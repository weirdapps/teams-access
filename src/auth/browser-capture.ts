// src/auth/browser-capture.ts
import type { BrowserContext, Request as PWRequest } from 'playwright';
import {
  readSession,
  writeSession,
  type Session,
  type SessionCookie,
  type SessionRegion,
  type AudienceToken,
} from '../session/store';
import { decodeJwt } from '../session/jwt';
import { harvestMsalTokens } from './msal-harvest';
import { ExitCode, ExitWithCode } from '../util/exit-codes';

// Teams web moved to the cloud.microsoft domain, the same consolidation that
// produced m365.cloud.microsoft (already in the nav list below). teams.microsoft.com
// still answers 200 and redirects, so nothing here failed loudly: the redirect just
// lands on a shell whose app rail does not match the selectors in
// clickAppRailForGraph, the click that "actually provokes Graph" reported
// "found no match", and renewal captured zero audiences while the session itself was
// perfectly healthy. Measured 2026-09-03, after headless renewal had been returning
// "(none)" for days against a tenant where Teams web worked fine in a normal browser.
const TEAMS_ROOT = 'https://teams.cloud.microsoft/';

// Both hosts, for the "is this a Teams page?" test. A redirect can leave either one
// in the URL, and the nav list below still visits the legacy host as a fallback in
// case a tenant has not been migrated yet.
const TEAMS_HOSTS = [
  'https://teams.cloud.microsoft/',
  'https://teams.microsoft.com/',
  // MCAS (Defender for Cloud Apps) Conditional Access App Control serves the
  // whole tenant through "<original-fqdn>.mcas.ms". These are SEPARATE ORIGINS
  // with their own cookie jar and their own localStorage, which is why a session
  // can be alive on one and dead on the other. Confirmed on this tenant: the
  // browser's real Teams runs at teams.cloud.microsoft.mcas.ms.
  'https://teams.cloud.microsoft.mcas.ms/',
  'https://teams.microsoft.com.mcas.ms/',
];

// Microsoft Graph well-known appid + audience strings.
const GRAPH_APPID = '00000003-0000-0000-c000-000000000000';
const GRAPH_AUD_RE =
  /(^|\/)(graph\.microsoft\.com|graph\.microsoft\.us|microsoftgraph\.chinacloudapi\.cn)(\/|$)/i;

export interface CaptureOptions {
  loginTimeoutMs: number;
  chromeChannel: string;
  profileDir?: string;
  /**
   * Minimum number of distinct audience tokens to capture before resolving.
   * Default is 1 (resolve on first Bearer). Set higher on headless VPS
   * environments where MCAS proxies trickle tokens across multiple navigations.
   * The browser stays open (collecting tokens) until this count is reached or
   * loginTimeoutMs expires — whichever comes first.
   */
  minAudiences?: number;
  /**
   * When set (>0), after a Graph token is accepted, the browser remains open
   * for this many additional milliseconds. This window is used purely for
   * diagnostic capture — we keep logging every Bearer the page requests, so
   * the trace file grows to include URLs we'd miss if we closed immediately.
   * Defaults to 0 (close immediately on accept). Set higher to discover more
   * (URL, audience) tuples for Path B (multi-service) architecture work.
   */
  diagnosticExtraMs?: number;
  /**
   * When true, launch Chromium headless. Requires a persistent profileDir
   * with a valid ESTSAUTHPERSISTENT cookie so Entra silently re-issues
   * Bearer tokens without user interaction. Used by `teams-cli auth-renew`.
   */
  headless?: boolean;
}

export interface CapturedTokenInfo {
  bearerToken: string;
  upn?: string;
  oid?: string;
  tid?: string;
  aud?: string;
  appid?: string;
  scp?: string;
}

/**
 * Pure function: given a Playwright Request, decide if it carries an acceptable Bearer.
 *
 * Path B: we accept ANY audience-bound Bearer because the multi-token Session
 * stores all of them keyed by aud. The "primary" bearerToken on Session is
 * just whichever one we happened to capture first — it's a backward-compat
 * shim, no longer the only credential.
 *
 * Returns null only if the request has no Bearer at all OR the Bearer doesn't
 * decode as a 3-segment JWT.
 */
export function extractBearerFromRequest(req: PWRequest): CapturedTokenInfo | null {
  const headers = req.headers();
  const auth = headers['authorization'] ?? headers['Authorization'];
  if (!auth || !auth.startsWith('Bearer ')) return null;
  const token = auth.slice(7).trim();
  if (token.split('.').length !== 3) return null;
  let claims;
  try {
    claims = decodeJwt(token);
  } catch {
    return null;
  }
  const aud = typeof claims.aud === 'string' ? claims.aud : undefined;
  return {
    bearerToken: token,
    upn: claims.upn ?? claims.unique_name ?? undefined,
    oid: claims.oid as string | undefined,
    tid: claims.tid as string | undefined,
    aud,
    appid: claims.appid as string | undefined,
    scp: claims.scp as string | undefined,
  };
}

/** True if the audience claim resolves to Microsoft Graph. */
export function isGraphAudience(aud: string | undefined): boolean {
  return aud != null && (aud === GRAPH_APPID || GRAPH_AUD_RE.test(aud));
}

interface InspectedBearer {
  shortLine: string; // brief line for stderr
  jsonLine: string; // structured JSON line for trace file
  audience: string; // decoded aud claim
  hostPath: string; // host + path for dedup
}

/**
 * Inspect ALL Bearer-bearing requests (regardless of audience) for diagnostic logging.
 * Returns the brief stderr line + structured JSON line for the trace file, or null
 * if the request has no Bearer.
 */
function inspectBearerForLog(req: PWRequest): InspectedBearer | null {
  const headers = req.headers();
  const auth = headers['authorization'] ?? headers['Authorization'];
  if (!auth || !auth.startsWith('Bearer ')) return null;
  const token = auth.slice(7).trim();
  if (token.split('.').length !== 3) return null;
  let aud = '?';
  let appid = '?';
  try {
    const claims = decodeJwt(token);
    aud = typeof claims.aud === 'string' ? claims.aud : '?';
    appid = (claims.appid as string | undefined) ?? '?';
  } catch {
    /* ignore */
  }
  let host = '?';
  let path = '?';
  let url = req.url();
  try {
    const u = new URL(url);
    host = u.host;
    path = u.pathname + (u.search ? '?' + u.search.slice(1).split('&').slice(0, 3).join('&') : '');
  } catch {
    /* ignore */
  }
  const method = req.method();
  const hostPath = `${method} ${host}${path}`;
  return {
    shortLine: `[bearer-seen] ${method} ${host}${path.length > 80 ? path.slice(0, 80) + '…' : path} aud=${aud}`,
    jsonLine: JSON.stringify({ ts: new Date().toISOString(), method, host, path, aud, appid }),
    audience: aud,
    hostPath,
  };
}

/**
 * Click the Teams app rail so the SPA asks Entra for a Graph token.
 *
 * The old comment here said auto-clicking was tried and that "Teams web's DOM
 * doesn't expose stable selectors for those buttons". That is not true of the
 * current UI: `[aria-label*="Calendar" i]` matched and clicked first attempt on
 * 2026-09-02, and the Graph Bearer landed within seconds. Files is kept as a
 * fallback because it is SharePoint-backed and also calls Graph.
 *
 * Deliberately best-effort: every failure is swallowed. A missing selector must
 * never break a capture that is otherwise succeeding on the other audiences.
 */
async function clickAppRailForGraph(page: {
  locator: (sel: string) => {
    first: () => {
      isVisible: (o?: object) => Promise<boolean>;
      click: (o?: object) => Promise<void>;
    };
  };
  waitForTimeout: (ms: number) => Promise<void>;
}): Promise<void> {
  const candidates = [
    '[aria-label*="Calendar" i]',
    '[data-tid*="calendar" i]',
    'button:has-text("Calendar")',
    '[aria-label*="Files" i]',
    '[data-tid*="files" i]',
  ];
  for (const sel of candidates) {
    try {
      const el = page.locator(sel).first();
      if (await el.isVisible({ timeout: 2500 })) {
        await el.click({ timeout: 5000 });
        process.stderr.write(`[teams-cli login] clicked app rail (${sel}) to provoke Graph\n`);
        await page.waitForTimeout(8000);
        return;
      }
    } catch {
      /* try the next selector */
    }
  }
  process.stderr.write(
    '[teams-cli login] app-rail click found no match; Graph may not be captured\n',
  );
}

export async function captureSession(opts: CaptureOptions): Promise<Session> {
  const { chromium } = await import('playwright');
  let context: BrowserContext | undefined;
  // Track distinct audiences seen so we can give a useful diagnostic on timeout.
  const audiencesSeen = new Map<string, number>();
  // Dedup stderr lines by (method+host+path+audience) so output stays readable
  // even when Teams web bursts hundreds of identical-shape requests.
  const seenHostPathAud = new Set<string>();
  // Path B foundation: collect the FIRST Bearer seen for each distinct audience.
  // Written to ~/.teams-cli/multi-tokens.json (mode 0600) so we can probe the
  // discovered Skype/IC3/chatsvcagg/etc endpoints with the right token per
  // audience. This is the seed of Path B's multi-token session model.
  interface CapturedToken {
    token: string;
    aud: string;
    appid?: string;
    scp?: string;
    exp?: number;
    capturedAt: string;
  }
  const tokensByAud = new Map<string, CapturedToken>();
  let primaryResolved = false;

  // Merge a batch of MSAL-cache tokens into tokensByAud, persist, and settle the
  // login if the network listener has not already done so. Declared here (rather
  // than inline) because it is called once per navigated origin AND once at the
  // end: localStorage is per-origin, so a single harvest only ever sees one
  // surface's cache.
  let mergeHarvest: (batch: import('./msal-harvest').HarvestedToken[]) => void = () => {};

  // Open the trace file. Each Bearer-bearing request appends one JSON line.
  // Path is fixed under ~/.teams-cli/ so subsequent debug runs can grep it.
  const { writeFileSync, appendFileSync, mkdirSync, chmodSync, renameSync } =
    await import('node:fs');
  const { join } = await import('node:path');
  const { homedir } = await import('node:os');
  const traceDir = join(process.env.HOME ?? homedir(), '.teams-cli');
  try {
    mkdirSync(traceDir, { recursive: true, mode: 0o700 });
  } catch {
    /* ignore */
  }
  const tracePath = join(traceDir, 'login-trace.jsonl');
  const multiTokensPath = join(traceDir, 'multi-tokens.json');

  // Atomic write of the per-audience token map. Extracted so the network path
  // and the MSAL-cache harvest share one implementation rather than drifting.
  const writeMultiTokens = (map: Map<string, CapturedToken>): void => {
    try {
      const payload: Record<
        string,
        { token: string; appid?: string; scp?: string; exp?: number; capturedAt: string }
      > = {};
      for (const [aud, t] of map.entries()) {
        payload[aud] = {
          token: t.token,
          appid: t.appid,
          scp: t.scp,
          exp: t.exp,
          capturedAt: t.capturedAt,
        };
      }
      const tmpPath = `${multiTokensPath}.tmp`;
      writeFileSync(tmpPath, JSON.stringify(payload, null, 2), { mode: 0o600 });
      chmodSync(tmpPath, 0o600);
      renameSync(tmpPath, multiTokensPath);
    } catch (e) {
      process.stderr.write(`[multi-tokens] WARN write failed: ${(e as Error).message}\n`);
    }
  };
  // Truncate any prior trace.
  try {
    writeFileSync(tracePath, '');
  } catch {
    /* ignore */
  }

  try {
    context = await chromium.launchPersistentContext(opts.profileDir ?? '', {
      channel: opts.chromeChannel,
      headless: opts.headless ?? false,
      viewport: null,
    });

    // Capture strategy (Path B): grab the FIRST Bearer of any audience and resolve.
    // The diagnostic window collects more audiences after that; the multi-token
    // session is built from everything seen. We no longer require a Graph token
    // to be the "primary" — it's just one of N tokens in session.tokens.
    // Hoisted so the MSAL-cache harvest below can settle this promise too. The
    // network listener is no longer the only way a login can succeed.
    let resolveCaptured: ((v: CapturedTokenInfo) => void) | undefined;
    let captureTimer: ReturnType<typeof setTimeout> | undefined;

    const captured = new Promise<CapturedTokenInfo>((resolve, reject) => {
      resolveCaptured = resolve;
      const timer = setTimeout(() => {
        const summary = Array.from(audiencesSeen.entries())
          .map(([aud, n]) => `  ${aud} (×${n})`)
          .join('\n');
        reject(
          new ExitWithCode(ExitCode.AuthRequired, {
            code: 'auth_required',
            message:
              `Login window timed out after ${opts.loginTimeoutMs}ms with no Bearer captured. ` +
              `Distinct audiences seen during the wait:\n${summary || '  (none)'}\n` +
              `If you see no audiences, the page hasn't loaded yet — increase --login-timeout.`,
            audiencesSeen: Object.fromEntries(audiencesSeen.entries()),
          }),
        );
      }, opts.loginTimeoutMs);
      captureTimer = timer;

      context!.on('request', (req) => {
        // Diagnostic log for every Bearer (regardless of audience).
        const log = inspectBearerForLog(req);
        if (log) {
          // Always append the structured line to the trace file (no dedup there).
          try {
            appendFileSync(tracePath, log.jsonLine + '\n');
          } catch {
            /* ignore */
          }
          // Dedup the stderr line: only print first occurrence of each unique
          // (method+host+path+audience) tuple so the user can read it.
          const dedupKey = `${log.hostPath} aud=${log.audience}`;
          if (!seenHostPathAud.has(dedupKey)) {
            seenHostPathAud.add(dedupKey);
            process.stderr.write(log.shortLine + '\n');
          }
          audiencesSeen.set(log.audience, (audiencesSeen.get(log.audience) ?? 0) + 1);

          // Multi-audience token capture: keep the FIRST Bearer for each new
          // audience. Later runs of the same Teams session may issue refreshed
          // tokens with the same audience; keeping the first is fine because
          // they all share the same scope and lifetime profile.
          if (log.audience && log.audience !== '?' && !tokensByAud.has(log.audience)) {
            const headers = req.headers();
            const auth = headers['authorization'] ?? headers['Authorization'];
            if (auth && auth.startsWith('Bearer ')) {
              const tok = auth.slice(7).trim();
              let exp: number | undefined;
              let scp: string | undefined;
              let appid: string | undefined;
              try {
                const c = decodeJwt(tok);
                exp = typeof c.exp === 'number' ? c.exp : undefined;
                scp = typeof c.scp === 'string' ? c.scp : undefined;
                appid = typeof c.appid === 'string' ? c.appid : undefined;
              } catch {
                /* ignore */
              }
              tokensByAud.set(log.audience, {
                token: tok,
                aud: log.audience,
                appid,
                scp,
                exp,
                capturedAt: new Date().toISOString(),
              });
              // Persist the multi-tokens file IMMEDIATELY so even a timed-out
              // login leaves us with whatever tokens were collected. Atomic
              // write via temp+rename to avoid torn reads.
              try {
                const payload: Record<
                  string,
                  { token: string; appid?: string; scp?: string; exp?: number; capturedAt: string }
                > = {};
                for (const [aud, t] of tokensByAud.entries()) {
                  payload[aud] = {
                    token: t.token,
                    appid: t.appid,
                    scp: t.scp,
                    exp: t.exp,
                    capturedAt: t.capturedAt,
                  };
                }
                const tmpPath = `${multiTokensPath}.tmp`;
                writeFileSync(tmpPath, JSON.stringify(payload, null, 2), { mode: 0o600 });
                chmodSync(tmpPath, 0o600);
                renameSync(tmpPath, multiTokensPath);
                process.stderr.write(
                  `[multi-tokens] +${log.audience} (now have ${tokensByAud.size} audiences)\n`,
                );
              } catch (e) {
                process.stderr.write(`[multi-tokens] WARN write failed: ${(e as Error).message}\n`);
              }
            }
          }
        }
        // Path B accept: take the FIRST Bearer of any audience, prefer Graph
        // if it shows up early. Once resolved, additional Bearers continue to
        // be collected into tokensByAud (above) for the diagnostic window.
        //
        // minAudiences gate: when set > 1, delay resolution until we have
        // captured at least that many distinct audience tokens. This is needed
        // on VPS/headless environments where MCAS proxies deliver tokens one
        // at a time across multiple navigation steps.
        const minAud = opts.minAudiences ?? 1;
        const info = extractBearerFromRequest(req);
        if (info && !primaryResolved) {
          if (tokensByAud.size < minAud) {
            process.stderr.write(
              `[waiting] aud=${info.aud} (${tokensByAud.size}/${minAud} audiences — need more)\n`,
            );
            return;
          }
          // If the first-seen token isn't Graph but Graph is already in
          // tokensByAud (rare but possible due to async), prefer Graph.
          const graphAud = Array.from(tokensByAud.keys()).find(isGraphAudience);
          if (graphAud) {
            const t = tokensByAud.get(graphAud)!;
            process.stderr.write(
              `[accepted] aud=${graphAud} (preferred Graph over ${info.aud}) scp=${t.scp ?? '(none)'}\n`,
            );
            primaryResolved = true;
            clearTimeout(timer);
            resolve({
              bearerToken: t.token,
              upn: info.upn,
              oid: info.oid,
              tid: info.tid,
              aud: graphAud,
              appid: t.appid,
              scp: t.scp,
            });
          } else {
            process.stderr.write(
              `[accepted] aud=${info.aud} appid=${info.appid} scp=${info.scp ?? '(none)'}\n`,
            );
            primaryResolved = true;
            clearTimeout(timer);
            resolve(info);
          }
        }
      });
    });

    mergeHarvest = (batch) => {
      const fresh: string[] = [];
      for (const t of batch) {
        const prev = tokensByAud.get(t.aud);
        // Keep whichever expires later: a later origin can hold a fresher copy.
        if (prev && (prev.exp ?? 0) >= (t.exp ?? 0)) continue;
        tokensByAud.set(t.aud, {
          token: t.token,
          aud: t.aud,
          exp: t.exp,
          scp: t.scp,
          appid: t.appid,
          capturedAt: new Date().toISOString(),
        });
        if (!prev) fresh.push(t.aud);
      }
      if (!fresh.length) return;
      process.stderr.write(
        `[msal-cache] +${fresh.length} audience(s): ${fresh.join(', ')} (${tokensByAud.size} total)\n`,
      );
      writeMultiTokens(tokensByAud);

      // Settle the login on the cache when the wire produced nothing, which is
      // the normal case on an MCAS tenant. Prefer Graph as the primary token so
      // the resulting session matches what the network path used to produce.
      if (primaryResolved || !resolveCaptured) return;
      const graph =
        Array.from(tokensByAud.values()).find((t) => isGraphAudience(t.aud)) ??
        Array.from(tokensByAud.values())[0];
      if (!graph) return;
      primaryResolved = true;
      if (captureTimer) clearTimeout(captureTimer);
      let claims: Record<string, unknown> = {};
      try {
        claims = decodeJwt(graph.token) as Record<string, unknown>;
      } catch {
        /* a token we cannot decode is still usable as a Bearer */
      }
      resolveCaptured({
        bearerToken: graph.token,
        aud: graph.aud,
        appid: graph.appid,
        scp: graph.scp,
        upn:
          (typeof claims.upn === 'string' ? claims.upn : undefined) ??
          (typeof claims.preferred_username === 'string' ? claims.preferred_username : undefined),
        oid: typeof claims.oid === 'string' ? claims.oid : undefined,
        tid: typeof claims.tid === 'string' ? claims.tid : undefined,
      });
    };

    const page = await context.newPage();
    await page.goto(TEAMS_ROOT);

    // For an interactive login, go to the MCAS origin instead of the canonical
    // one, because only the MCAS origin tells the truth about the session.
    //
    // Teams registers a Service Worker that serves a complete offline shell from
    // IndexedDB. On a DEAD session that shell still renders: measured 2026-09-03,
    // a profile whose every MSAL token had expired 28 hours earlier loaded
    // "(1) Chat | ... | Microsoft Teams" with 292 [data-tid] nodes, real chat
    // history and a clickable app rail. The app never reached the network, so it
    // never discovered it was logged out, so NO SIGN-IN PAGE EVER APPEARED. A
    // human sitting in front of that window has nothing to sign in to. That is
    // why repeated interactive logins produced one audience or none.
    //
    // Unregistering the Service Worker is NOT sufficient: tried, it reports
    // "unregistered 1" and the very next load still returns the cached shell
    // with zero input fields, because the app is also in the HTTP cache.
    //
    // Going to <host>.mcas.ms does work, and is measured, not assumed:
    //   teams.cloud.microsoft/v2/          -> title "(1) Chat | ...", 0 inputs
    //   teams.cloud.microsoft.mcas.ms/v2/  -> login.microsoftonline.com,
    //                                         title "Sign in to your account",
    //                                         1 email input, 1 password input
    // Same profile, same second. The MCAS origin is a separate origin with its
    // own cookie jar and no Service Worker cache of its own, so it cannot serve
    // a stale shell.
    //
    // A tenant without Conditional Access App Control has no such host; the goto
    // fails fast and we keep the canonical page already loaded above.
    if (!opts.headless) {
      const mcasEntry = 'https://teams.cloud.microsoft.mcas.ms/v2/?view=Chat';
      try {
        await page.goto(mcasEntry, { waitUntil: 'domcontentloaded', timeout: 30000 });
        process.stderr.write(
          `[teams-cli login] entered via the MCAS origin so a dead session shows a real sign-in page rather than the cached offline shell\n`,
        );
      } catch {
        process.stderr.write(
          `[teams-cli login] no MCAS origin on this tenant; staying on ${TEAMS_ROOT}\n`,
        );
      }
    }
    process.stderr.write(`[teams-cli login] navigated to ${page.url()}\n`);
    // Poll the MSAL cache of whatever origin the page is currently on, until the
    // login settles or the window closes.
    //
    // A one-shot harvest is not enough for two independent reasons. On a live
    // profile the Teams SPA populates its cache within seconds and we want to
    // finish immediately, without navigating anywhere. On a dead profile the
    // human has to sign in first, and that takes as long as it takes: fixed
    // harvest points all fire before the credentials are even typed, then report
    // "session is dead" about a session that came up thirty seconds later.
    // Polling covers both, and costs one cheap localStorage read every 5s.
    const harvestPoll = setInterval(() => {
      void (async () => {
        if (primaryResolved) return;
        try {
          mergeHarvest(await harvestMsalTokens(page));
        } catch {
          /* page navigating or closed; the next tick retries */
        }
      })();
    }, 5000);
    if (typeof harvestPoll.unref === 'function') harvestPoll.unref();
    const stopHarvestPoll = (): void => clearInterval(harvestPoll);
    page.once('close', stopHarvestPoll);
    context.once('close', stopHarvestPoll);
    process.stderr.write(
      `[teams-cli login] waiting for a Graph-audience Bearer (timeout ${opts.loginTimeoutMs}ms)\n`,
    );
    process.stderr.write(`[teams-cli login] full request trace: ${tracePath}\n`);
    if (!opts.headless) {
      process.stderr.write(
        `[teams-cli login] HINT: open a chat AND scroll its history. That triggers the chat-list and chat-message endpoints we want to discover.\n`,
      );
    }

    // Drive the browser through M365 surfaces that reliably provoke Graph +
    // chatsvcagg + ic3 audience requests. Validated empirically: this sequence
    // captures all critical Phase 1 audiences (graph.microsoft.com, chatsvcagg,
    // ic3, presence) in ~45s.
    //
    // This used to be gated on `opts.headless`, which split the two halves of a
    // working capture apart and meant neither mode could finish the job:
    //   headed   authenticates fine, but only printed the hint above and waited
    //            for a human to click. Left alone it captured one audience, or
    //            none at all if nobody happened to be watching the window.
    //   headless clicks everything correctly, but on a Conditional-Access tenant
    //            never authenticates: measured 2026-09-03, an entire headless run
    //            wrote a ZERO-BYTE request trace, i.e. not one Bearer of any
    //            audience, while the same profile worked headed.
    // Running the sequence in both modes gives the combination that actually
    // works: a real, CA-satisfying Chrome session that also clicks the app rail
    // itself instead of hoping someone is at the keyboard.
    {
      // Run navigations in a non-blocking IIFE so the `await captured` below
      // still drives the timeout/resolve loop. We only need the navigations to
      // FIRE the requests; the request listener captures Bearers as they go.
      //
      // Order matters:
      //   1. Teams /v2/?view=Chat   — provokes chatsvcagg.teams.microsoft.com
      //                                (chat list/updates endpoint). Without
      //                                this, only base Graph token is captured
      //                                and downstream `list-messages` etc fail
      //                                401 with "no chatsvcagg token in
      //                                session". 7s settle because the chat
      //                                SPA panel takes longer to init than a
      //                                static page.
      //   2. outlook.office.com     — provokes Graph + outlook audiences.
      //   3. www.office.com         — catches presence + remaining tenant
      //                                surfaces.
      void (async () => {
        for (const { url, settleMs } of [
          // MCAS origin first: on a proxied tenant this is where the live session
          // and its storage actually are, and it is the only one that reveals a
          // dead session by redirecting to login instead of serving a cached shell.
          { url: 'https://teams.cloud.microsoft.mcas.ms/v2/?view=Chat', settleMs: 9000 },
          { url: 'https://teams.cloud.microsoft/v2/?view=Chat', settleMs: 7000 },
          // Legacy host, kept as a fallback for a tenant that has not been moved to
          // cloud.microsoft yet. On a migrated tenant this just redirects and costs
          // a few seconds; it must NOT come first, or the redirect chain is what the
          // app-rail click lands on.
          { url: 'https://teams.microsoft.com/v2/?view=Chat', settleMs: 5000 },
          //   2. outlook.office.com/mail/ — the Graph provoker. It must stay
          //      SECOND so both required audiences (chatsvcagg + graph) land
          //      early, inside the token-sync's auth-renew timeout budget.
          //
          //      The /mail/ path is load-bearing. Bare outlook.office.com and
          //      m365.cloud.microsoft used to yield a Graph Bearer and stopped:
          //      on 2026-08-11 both were measured over 30s each and produced
          //      NONE, while /mail/ produced three Graph calls (/v1.0/<tid>/,
          //      /v1.0/organization, /beta/me/settings) carrying exactly the
          //      scopes this CLI needs (Chat.Read, Chat.ReadWrite,
          //      Channel.ReadBasic.All). Bare office.com just redirects to
          //      m365.cloud.microsoft, so it is not a second chance at Graph.
          //
          //      Symptom when this breaks: renewal captures 4-6 audiences,
          //      every one of them non-Graph, so auth-renew fails its
          //      REQUIRED_AUDIENCES check and parks the teams sentinel. The
          //      session itself is fine, which is why it reads as "interactive
          //      login required" when no human input would have helped.
          //      2026-09-03: on an MCAS tenant this host is not where Outlook web
          //      actually lives. Defender for Cloud Apps proxies through
          //      "<original-fqdn>.mcas.ms", and the measured live URL was
          //      outlook.office365.com.mcas.ms/mail/. Navigating to the canonical
          //      host on such a tenant brings up no SPA, so the Graph provoker
          //      provoked nothing and renewal saw zero audiences. Both forms are
          //      tried; .mcas.ms is the generic proxy hostname (the tenant is in
          //      the McasTsid query param), so nothing tenant-specific is pinned.
          { url: 'https://outlook.office.com/mail/', settleMs: 15000 },
          { url: 'https://outlook.office365.com.mcas.ms/mail/', settleMs: 15000 },
          //   3+. Kept as belt and braces for presence / tenant surfaces, and
          //      in case Microsoft restores Graph on the M365 home.
          { url: 'https://m365.cloud.microsoft/', settleMs: 10000 },
          { url: 'https://www.office.com/', settleMs: 5000 },
        ]) {
          try {
            process.stderr.write(`[teams-cli login] headless: navigating to ${url}\n`);
            await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
            await page.waitForTimeout(settleMs);
            // Clicking the Teams app rail is what actually provokes Graph.
            // Measured 2026-09-02: NAVIGATION ALONE IS NOT ENOUGH. Loading
            // teams.microsoft.com/v2/?view=Calendar and ?view=Files headlessly,
            // 20s settle each, yielded 5 audiences and no Graph. Clicking the
            // Calendar rail item in the same session immediately produced
            // GET graph.microsoft.com/v1.0/<tid>/subscribedskus. The SPA only
            // issues the Graph call in response to a real UI interaction.
            if (TEAMS_HOSTS.some((h) => url.startsWith(h))) {
              await clickAppRailForGraph(page);
            }

            // Harvest HERE, before navigating away. localStorage is per-origin,
            // so the MSAL cache that Teams writes on teams.cloud.microsoft is
            // simply not readable once the page is on outlook.office.com or
            // www.office.com. Harvesting only at the end of the loop read the
            // LAST origin and reported "session is dead" while a full set of
            // valid Teams tokens sat one origin behind it. Each surface also
            // keeps its own cache, so sweeping every origin is what collects
            // Teams, Outlook and the M365 shell audiences in one pass.
            const here = await harvestMsalTokens(page);
            if (here.length) mergeHarvest(here);
          } catch (e) {
            const msg = e instanceof Error ? e.message : String(e);
            process.stderr.write(`[teams-cli login] headless nav warning (${url}): ${msg}\n`);
          }
        }

        // Final sweep, for the last origin visited.
        try {
          mergeHarvest(await harvestMsalTokens(page));
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          process.stderr.write(`[msal-cache] harvest failed: ${msg}\n`);
        }

        if (tokensByAud.size === 0) {
          process.stderr.write(
            `[msal-cache] no unexpired AccessToken in any visited origin's localStorage. The ` +
              `profile's session is dead (MSAL keeps expired entries, so this means expired, not ` +
              `absent). Sign in interactively.\n`,
          );
        }
      })();
    }

    const info = await captured;

    const hasGraph = Array.from(tokensByAud.keys()).some(isGraphAudience);
    if (!hasGraph) {
      process.stderr.write(
        `[teams-cli login] WARNING: no Graph-audience token was captured. ` +
          `Commands that use Microsoft Graph (list-teams, list-channels, send-message, auth-check) will fail with 401. ` +
          `To get a Graph token, click the Calendar tab or Files tab in Teams web while the diagnostic window is open.\n`,
      );
    }
    process.stderr.write(
      `[teams-cli login] tokens captured so far: ${Array.from(tokensByAud.keys()).join(', ')}\n`,
    );

    // Optional diagnostic window: keep listening (and logging) for additional
    // ms after Graph token accept, so the user can interact with Teams (open
    // a chat, scroll history) and we capture URL+audience tuples for the
    // ic3 / chatsvcagg endpoints that didn't fire in the initial burst.
    if (opts.diagnosticExtraMs && opts.diagnosticExtraMs > 0) {
      process.stderr.write(
        `[teams-cli login] Graph token captured. Diagnostic window: ${opts.diagnosticExtraMs}ms (keep clicking around in Teams — open a chat, scroll history)\n`,
      );
      await new Promise((r) => setTimeout(r, opts.diagnosticExtraMs));
      process.stderr.write(`[teams-cli login] diagnostic window over\n`);
    }

    process.stderr.write(
      `[teams-cli login] trace summary (unique audience × method+host+path tuples):\n`,
    );
    for (const k of Array.from(seenHostPathAud).sort()) {
      process.stderr.write(`  ${k}\n`);
    }
    process.stderr.write(
      `[teams-cli login] full trace at ${tracePath} — ${Array.from(seenHostPathAud).length} unique tuples / ${Array.from(audiencesSeen.values()).reduce((a, b) => a + b, 0)} total Bearer requests\n`,
    );

    // Write the multi-audience token map for Path B probing / future use.
    if (tokensByAud.size > 0) {
      const payload: Record<
        string,
        { token: string; appid?: string; scp?: string; exp?: number; capturedAt: string }
      > = {};
      for (const [aud, t] of tokensByAud.entries()) {
        payload[aud] = {
          token: t.token,
          appid: t.appid,
          scp: t.scp,
          exp: t.exp,
          capturedAt: t.capturedAt,
        };
      }
      try {
        writeFileSync(multiTokensPath, JSON.stringify(payload, null, 2), { mode: 0o600 });
        chmodSync(multiTokensPath, 0o600);
        process.stderr.write(
          `[teams-cli login] multi-audience token map written to ${multiTokensPath} (${tokensByAud.size} audiences, mode 0600)\n`,
        );
      } catch (e) {
        process.stderr.write(
          `[teams-cli login] WARN: could not write multi-tokens file: ${(e as Error).message}\n`,
        );
      }
    }

    const pwCookies = await context.cookies();
    const cookies: SessionCookie[] = pwCookies.map((c) => ({
      name: c.name,
      value: c.value,
      domain: c.domain,
      path: c.path,
      expires: c.expires,
      httpOnly: c.httpOnly,
      secure: c.secure,
      sameSite: c.sameSite as 'Strict' | 'Lax' | 'None' | undefined,
    }));

    // Build the multi-audience token map for the session. Seed from the PRIOR
    // session so an audience not re-observed this run keeps its last-known token
    // instead of being dropped. Graph is captured only flakily headless (Teams
    // mints it on M365-home load — see the nav list above); without this seed a
    // single miss would clobber a still-valid Graph token and 401 every
    // Graph-backed command until the next lucky capture. This run's fresh
    // captures (loop below) overwrite the seeded entries, so newer always wins.
    const priorTokens = readSession()?.tokens ?? {};
    const tokens: Record<string, AudienceToken> = { ...priorTokens };
    for (const [aud, t] of tokensByAud.entries()) {
      tokens[aud] = {
        bearerToken: t.token,
        appid: t.appid,
        scp: t.scp,
        exp: t.exp,
        capturedAt: t.capturedAt,
      };
    }
    // Make sure the Graph token is always present in tokens map under the
    // canonical audience key, even if the capture loop missed it (e.g.
    // when the Graph host appears with a different aud-claim form).
    if (info.aud && !tokens[info.aud]) {
      tokens[info.aud] = {
        bearerToken: info.bearerToken,
        appid: info.appid,
        scp: info.scp,
        capturedAt: new Date().toISOString(),
      };
    }

    // Detect tenant region from observed URL paths.
    // Each regex picks out the region segment that appears in the URL path.
    const region: SessionRegion = {};
    for (const tuple of seenHostPathAud) {
      // tuple: "<METHOD> <host><path> aud=<aud>"
      let m: RegExpMatchArray | null;
      if (!region.chatsvc && (m = tuple.match(/\/api\/chatsvc\/([^/]+)\//))) region.chatsvc = m[1];
      if (!region.csa && (m = tuple.match(/\/api\/csa\/([^/]+)\//))) region.csa = m[1];
      if (!region.mt && (m = tuple.match(/\/api\/mt\/part\/([^/]+)\//))) region.mt = m[1];
      if (!region.asyncgw && (m = tuple.match(/^\S+\s+([\w-]+)\.asyncgw\.teams\.microsoft\.com/)))
        region.asyncgw = m[1];
    }
    process.stderr.write(`[teams-cli login] detected region: ${JSON.stringify(region)}\n`);

    // Primary bearer = the Graph token (by design). Prefer whatever Graph token
    // is now in the map — freshly captured OR preserved from the prior session —
    // over info.bearerToken, which is merely the FIRST audience seen this run
    // (often skype/chatsvc) and useless for Graph calls.
    const graphToken = tokens['https://graph.microsoft.com'];
    const session: Session = {
      bearerToken: graphToken?.bearerToken ?? info.bearerToken,
      cookies,
      capturedAt: new Date().toISOString(),
      account: { upn: info.upn, oid: info.oid, tid: info.tid },
      tokens,
      region,
    };
    writeSession(session);
    return session;
  } finally {
    if (context) await context.close();
  }
}
