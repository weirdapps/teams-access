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

/**
 * Inspect ALL Bearer-bearing requests (regardless of audience) for diagnostic logging.
 * Returns a brief summary string the caller can write to stderr.
 */
function inspectBearerForLog(req: PWRequest): string | null {
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
  try { host = new URL(req.url()).host; } catch { /* ignore */ }
  return `[bearer-seen] host=${host} aud=${aud} appid=${appid}`;
}

export async function captureSession(opts: CaptureOptions): Promise<Session> {
  let context: BrowserContext | undefined;
  // Track distinct audiences seen so we can give a useful diagnostic on timeout.
  const audiencesSeen = new Map<string, number>();

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
          process.stderr.write(log + '\n');
          // Track audience histogram for the timeout summary.
          const audMatch = log.match(/aud=([^\s]+)/);
          if (audMatch) {
            const aud = audMatch[1];
            audiencesSeen.set(aud, (audiencesSeen.get(aud) ?? 0) + 1);
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
    process.stderr.write(`[teams-cli login] HINT: click around in Teams (open a chat, view a profile, open the calendar tab) to provoke Graph calls.\n`);

    const info = await captured;

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
