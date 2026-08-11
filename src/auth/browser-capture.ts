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
import { ExitCode, ExitWithCode } from '../util/exit-codes';

const TEAMS_ROOT = 'https://teams.microsoft.com/';

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
    const captured = new Promise<CapturedTokenInfo>((resolve, reject) => {
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

    const page = await context.newPage();
    await page.goto(TEAMS_ROOT);
    process.stderr.write(`[teams-cli login] navigated to ${TEAMS_ROOT}\n`);
    process.stderr.write(
      `[teams-cli login] waiting for a Graph-audience Bearer (timeout ${opts.loginTimeoutMs}ms)\n`,
    );
    process.stderr.write(`[teams-cli login] full request trace: ${tracePath}\n`);
    if (!opts.headless) {
      process.stderr.write(
        `[teams-cli login] HINT: open a chat AND scroll its history. That triggers the chat-list and chat-message endpoints we want to discover.\n`,
      );
    }

    // In headless mode, the Teams SPA doesn't auto-navigate, so the user-clicks
    // hint above doesn't apply. Drive the browser through M365 surfaces that
    // reliably provoke Graph + chatsvcagg + ic3 audience requests. Validated
    // empirically: this sequence captures all critical Phase 1 audiences
    // (graph.microsoft.com, chatsvcagg, ic3, presence) in ~45s headless.
    if (opts.headless) {
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
          { url: 'https://teams.microsoft.com/v2/?view=Chat', settleMs: 7000 },
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
          { url: 'https://outlook.office.com/mail/', settleMs: 15000 },
          //   3+. Kept as belt and braces for presence / tenant surfaces, and
          //      in case Microsoft restores Graph on the M365 home.
          { url: 'https://m365.cloud.microsoft/', settleMs: 10000 },
          { url: 'https://www.office.com/', settleMs: 5000 },
        ]) {
          try {
            process.stderr.write(`[teams-cli login] headless: navigating to ${url}\n`);
            await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
            await page.waitForTimeout(settleMs);
          } catch (e) {
            const msg = e instanceof Error ? e.message : String(e);
            process.stderr.write(`[teams-cli login] headless nav warning (${url}): ${msg}\n`);
          }
        }
      })();
    }

    const info = await captured;

    // Note: we tried auto-clicking Calendar / Files tabs to provoke Graph
    // token acquisition, but Teams web's DOM doesn't expose stable selectors
    // for those buttons. Graph tokens get acquired only when Teams' own
    // background sync decides to fetch SharePoint sites or calendar items.
    // For now, surface a hint to the user if Graph token wasn't captured.
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
