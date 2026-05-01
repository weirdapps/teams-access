// src/auth/browser-capture.ts
import { chromium, type BrowserContext, type Request as PWRequest } from 'playwright';
import { writeSession, type Session, type SessionCookie } from '../session/store';
import { decodeJwt } from '../session/jwt';
import { ExitCode, ExitWithCode } from '../util/exit-codes';

const TEAMS_ROOT = 'https://teams.microsoft.com/';
const TARGET_HOST_RE = /(graph\.microsoft\.com|teams\.microsoft\.com|outlook\.office\.com)/;

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
}

/** Pure function: given a Playwright Request, decide if it carries a Bearer we want to keep. */
export function extractBearerFromRequest(req: PWRequest): CapturedTokenInfo | null {
  const url = req.url();
  if (!TARGET_HOST_RE.test(url)) return null;
  const headers = req.headers();
  const auth = headers['authorization'] ?? headers['Authorization'];
  if (!auth || !auth.startsWith('Bearer ')) return null;
  const token = auth.slice(7).trim();
  if (token.split('.').length !== 3) return null;
  try {
    const claims = decodeJwt(token);
    return {
      bearerToken: token,
      upn: claims.upn ?? claims.unique_name ?? undefined,
      oid: claims.oid as string | undefined,
      tid: claims.tid as string | undefined,
    };
  } catch {
    return null;
  }
}

export async function captureSession(opts: CaptureOptions): Promise<Session> {
  let context: BrowserContext | undefined;
  try {
    context = await chromium.launchPersistentContext(opts.profileDir ?? '', {
      channel: opts.chromeChannel,
      headless: false,
      viewport: null,
    });

    const captured = new Promise<CapturedTokenInfo>((resolve, reject) => {
      const timer = setTimeout(
        () => reject(new ExitWithCode(ExitCode.AuthRequired, {
          code: 'auth_required',
          message: `Login window timed out after ${opts.loginTimeoutMs}ms with no Bearer captured`,
        })),
        opts.loginTimeoutMs,
      );
      context!.on('request', (req) => {
        const info = extractBearerFromRequest(req);
        if (info) { clearTimeout(timer); resolve(info); }
      });
    });

    const page = await context.newPage();
    await page.goto(TEAMS_ROOT);
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
