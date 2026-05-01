// src/auth/browser-capture.ts
import { chromium, type BrowserContext, type Request as PWRequest } from 'playwright';
import { writeSession, type Session, type SessionCookie } from '../session/store';
import { decodeJwt } from '../session/jwt';
import { ExitCode, ExitWithCode } from '../util/exit-codes';

const TEAMS_ROOT = 'https://teams.microsoft.com/';

// Microsoft Graph well-known appid + audience strings.
const GRAPH_APPID = '00000003-0000-0000-c000-000000000000';
const GRAPH_AUD_RE = /(^|\/)(graph\.microsoft\.com|graph\.microsoft\.us|microsoftgraph\.chinacloudapi\.cn)(\/|$)/i;

export interface CaptureOptions {
  loginTimeoutMs: number;
  chromeChannel: string;
  profileDir?: string;
  /**
   * When set (>0), after a Graph token is accepted, the browser remains open
   * for this many additional milliseconds. This window is used purely for
   * diagnostic capture — we keep logging every Bearer the page requests, so
   * the trace file grows to include URLs we'd miss if we closed immediately.
   * Defaults to 0 (close immediately on accept). Set higher to discover more
   * (URL, audience) tuples for Path B (multi-service) architecture work.
   */
  diagnosticExtraMs?: number;
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
 * Pure function: given a Playwright Request, decide if it carries a Graph-bound Bearer.
 *
 * We accept ONLY tokens whose `aud` claim resolves to Microsoft Graph (commercial,
 * GovCloud, or China cloud), or whose `aud` is the Graph well-known appid GUID.
 * Tokens for other audiences (api.spaces.skype.com, chatsvc.teams.microsoft.com,
 * outlook.office.com etc.) are returned as null even though Teams web requests
 * them — only Graph tokens work against the public Graph REST surface this CLI
 * targets.
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
  const isGraphAud = aud != null && (aud === GRAPH_APPID || GRAPH_AUD_RE.test(aud));
  if (!isGraphAud) return null;
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

interface InspectedBearer {
  shortLine: string;       // brief line for stderr
  jsonLine: string;        // structured JSON line for trace file
  audience: string;        // decoded aud claim
  hostPath: string;        // host + path for dedup
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
  } catch { /* ignore */ }
  let host = '?';
  let path = '?';
  let url = req.url();
  try {
    const u = new URL(url);
    host = u.host;
    path = u.pathname + (u.search ? '?' + u.search.slice(1).split('&').slice(0, 3).join('&') : '');
  } catch { /* ignore */ }
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
  interface CapturedToken { token: string; aud: string; appid?: string; scp?: string; exp?: number; capturedAt: string }
  const tokensByAud = new Map<string, CapturedToken>();

  // Open the trace file. Each Bearer-bearing request appends one JSON line.
  // Path is fixed under ~/.teams-cli/ so subsequent debug runs can grep it.
  const { writeFileSync, appendFileSync, mkdirSync, chmodSync } = await import('node:fs');
  const { join } = await import('node:path');
  const { homedir } = await import('node:os');
  const traceDir = join(process.env.HOME ?? homedir(), '.teams-cli');
  try { mkdirSync(traceDir, { recursive: true, mode: 0o700 }); } catch { /* ignore */ }
  const tracePath = join(traceDir, 'login-trace.jsonl');
  const multiTokensPath = join(traceDir, 'multi-tokens.json');
  // Truncate any prior trace.
  try { writeFileSync(tracePath, ''); } catch { /* ignore */ }

  try {
    context = await chromium.launchPersistentContext(opts.profileDir ?? '', {
      channel: opts.chromeChannel,
      headless: false,
      viewport: null,
    });

    const captured = new Promise<CapturedTokenInfo>((resolve, reject) => {
      const timer = setTimeout(() => {
        const summary = Array.from(audiencesSeen.entries())
          .map(([aud, n]) => `  ${aud} (×${n})`)
          .join('\n');
        reject(new ExitWithCode(ExitCode.AuthRequired, {
          code: 'auth_required',
          message:
            `Login window timed out after ${opts.loginTimeoutMs}ms with no Graph-audience Bearer captured. ` +
            `Distinct audiences seen during the wait:\n${summary || '  (none)'}\n` +
            `If 'https://graph.microsoft.com' is not in this list, Teams web on your tenant did not request a Graph token. ` +
            `Try: open a chat, view someone's profile (presence triggers Graph), or open the calendar tab.`,
          audiencesSeen: Object.fromEntries(audiencesSeen.entries()),
        }));
      }, opts.loginTimeoutMs);

      context!.on('request', (req) => {
        // Diagnostic log for every Bearer (regardless of audience).
        const log = inspectBearerForLog(req);
        if (log) {
          // Always append the structured line to the trace file (no dedup there).
          try { appendFileSync(tracePath, log.jsonLine + '\n'); } catch { /* ignore */ }
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
              } catch { /* ignore */ }
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
                const payload: Record<string, { token: string; appid?: string; scp?: string; exp?: number; capturedAt: string }> = {};
                for (const [aud, t] of tokensByAud.entries()) {
                  payload[aud] = { token: t.token, appid: t.appid, scp: t.scp, exp: t.exp, capturedAt: t.capturedAt };
                }
                const tmpPath = `${multiTokensPath}.tmp`;
                writeFileSync(tmpPath, JSON.stringify(payload, null, 2), { mode: 0o600 });
                chmodSync(tmpPath, 0o600);
                const { renameSync } = require('node:fs');
                renameSync(tmpPath, multiTokensPath);
                process.stderr.write(`[multi-tokens] +${log.audience} (now have ${tokensByAud.size} audiences)\n`);
              } catch (e) {
                process.stderr.write(`[multi-tokens] WARN write failed: ${(e as Error).message}\n`);
              }
            }
          }
        }
        // Selective accept: only Graph-audience tokens.
        const info = extractBearerFromRequest(req);
        if (info) {
          process.stderr.write(`[accepted] aud=${info.aud} appid=${info.appid} scp=${info.scp ?? '(none)'}\n`);
          clearTimeout(timer);
          resolve(info);
        }
      });
    });

    const page = await context.newPage();
    await page.goto(TEAMS_ROOT);
    process.stderr.write(`[teams-cli login] navigated to ${TEAMS_ROOT}\n`);
    process.stderr.write(`[teams-cli login] waiting for a Graph-audience Bearer (timeout ${opts.loginTimeoutMs}ms)\n`);
    process.stderr.write(`[teams-cli login] full request trace: ${tracePath}\n`);
    process.stderr.write(`[teams-cli login] HINT: open a chat AND scroll its history. That triggers the chat-list and chat-message endpoints we want to discover.\n`);

    const info = await captured;

    // Optional diagnostic window: keep listening (and logging) for additional
    // ms after Graph token accept, so the user can interact with Teams (open
    // a chat, scroll history) and we capture URL+audience tuples for the
    // ic3 / chatsvcagg endpoints that didn't fire in the initial burst.
    if (opts.diagnosticExtraMs && opts.diagnosticExtraMs > 0) {
      process.stderr.write(`[teams-cli login] Graph token captured. Diagnostic window: ${opts.diagnosticExtraMs}ms (keep clicking around in Teams — open a chat, scroll history)\n`);
      await new Promise(r => setTimeout(r, opts.diagnosticExtraMs));
      process.stderr.write(`[teams-cli login] diagnostic window over\n`);
    }

    process.stderr.write(`[teams-cli login] trace summary (unique audience × method+host+path tuples):\n`);
    for (const k of Array.from(seenHostPathAud).sort()) {
      process.stderr.write(`  ${k}\n`);
    }
    process.stderr.write(`[teams-cli login] full trace at ${tracePath} — ${Array.from(seenHostPathAud).length} unique tuples / ${Array.from(audiencesSeen.values()).reduce((a, b) => a + b, 0)} total Bearer requests\n`);

    // Write the multi-audience token map for Path B probing / future use.
    if (tokensByAud.size > 0) {
      const payload: Record<string, { token: string; appid?: string; scp?: string; exp?: number; capturedAt: string }> = {};
      for (const [aud, t] of tokensByAud.entries()) {
        payload[aud] = { token: t.token, appid: t.appid, scp: t.scp, exp: t.exp, capturedAt: t.capturedAt };
      }
      try {
        writeFileSync(multiTokensPath, JSON.stringify(payload, null, 2), { mode: 0o600 });
        chmodSync(multiTokensPath, 0o600);
        process.stderr.write(`[teams-cli login] multi-audience token map written to ${multiTokensPath} (${tokensByAud.size} audiences, mode 0600)\n`);
      } catch (e) {
        process.stderr.write(`[teams-cli login] WARN: could not write multi-tokens file: ${(e as Error).message}\n`);
      }
    }

    const pwCookies = await context.cookies();
    const cookies: SessionCookie[] = pwCookies.map(c => ({
      name: c.name,
      value: c.value,
      domain: c.domain,
      path: c.path,
      expires: c.expires,
      httpOnly: c.httpOnly,
      secure: c.secure,
      sameSite: c.sameSite as 'Strict' | 'Lax' | 'None' | undefined,
    }));

    const session: Session = {
      bearerToken: info.bearerToken,
      cookies,
      capturedAt: new Date().toISOString(),
      account: { upn: info.upn, oid: info.oid, tid: info.tid },
    };
    writeSession(session);
    return session;
  } finally {
    if (context) await context.close();
  }
}
