// src/commands/auth-renew.ts
//
// Silent (headless) bearer renewal. Uses the persisted Playwright profile
// (~/.teams-cli/playwright-profile/) to re-issue Teams Bearer tokens without
// opening a visible browser window.
//
// Works while the device-trust cookie (ESTSAUTHPERSISTENT, ~90 days) is
// alive. When that cookie expires or NBG forces re-MFA, this command fails
// and the caller must run `teams-cli login` interactively.
//
// Mirrors outlook-cli's auth-renew (commands/auth-renew.ts in outlook-access).

import { join } from 'node:path';
import { homedir } from 'node:os';

import { captureSession } from '../auth/browser-capture';
import { readSession, writeSession, type Session } from '../session/store';
import { ExitCode, ExitWithCode } from '../util/exit-codes';

/** Default headless renewal timeout. Headless mode drives 3 navigations
 *  (teams → outlook → office.com) to provoke Graph + chatsvcagg + ic3
 *  audience captures; needs ~45s headroom. */
const DEFAULT_RENEW_TIMEOUT_MS = 90_000;
/** After Graph token captured, keep listening this long for additional
 *  audiences to land (chatsvcagg, ic3, etc). */
const DEFAULT_DIAGNOSTIC_EXTRA_MS = 15_000;

export interface AuthRenewOptions {
  /** Override the renew-specific timeout (default 30000ms). */
  timeoutMs?: number;
  /** Override the Chromium channel (defaults to "chrome"). */
  chromeChannel?: string;
  /** Override the persisted profile directory. */
  profileDir?: string;
}

export interface AuthRenewResult {
  status: 'ok';
  sessionFile: string;
  account: { upn?: string; oid?: string; tid?: string };
  /** Wall-clock duration of the renewal in milliseconds. */
  durationMs: number;
  /** Number of audience tokens captured during the silent renewal. */
  audiencesCaptured: number;
}

function defaultProfileDir(): string {
  return join(process.env.HOME ?? homedir(), '.teams-cli', 'playwright-profile');
}

function defaultSessionPath(): string {
  return join(process.env.HOME ?? homedir(), '.teams-cli', 'session.json');
}

export async function runAuthRenew(opts: AuthRenewOptions = {}): Promise<AuthRenewResult> {
  // A renewal only makes sense if a prior interactive login left a profile
  // and a session file behind. Fail fast otherwise — the caller must run `login`.
  const sessionPath = defaultSessionPath();
  const existing = readSession();
  if (existing === null) {
    throw new ExitWithCode(ExitCode.AuthRequired, {
      code: 'auth_no_reauth',
      message: 'No cached session to renew. Run `teams-cli login` first.',
    });
  }

  const t0 = Date.now();

  let captured: Session;
  try {
    captured = await captureSession({
      loginTimeoutMs: opts.timeoutMs ?? DEFAULT_RENEW_TIMEOUT_MS,
      chromeChannel: opts.chromeChannel ?? 'chrome',
      profileDir: opts.profileDir ?? defaultProfileDir(),
      headless: true,
      diagnosticExtraMs: DEFAULT_DIAGNOSTIC_EXTRA_MS,
    });
  } catch (err) {
    // Headless renewal failed. Most likely cause: ESTSAUTHPERSISTENT cookie
    // expired or NBG forced re-MFA. Caller must run interactive login.
    const msg = err instanceof Error ? err.message : String(err);
    throw new ExitWithCode(ExitCode.AuthRequired, {
      code: 'auth_renew_failed',
      message: `Headless renewal failed: ${msg}. Run \`teams-cli login\`.`,
    });
  }

  // Persist the freshly-captured session.
  writeSession(captured);

  return {
    status: 'ok',
    sessionFile: sessionPath,
    account: captured.account ?? {},
    durationMs: Date.now() - t0,
    audiencesCaptured: Object.keys(captured.tokens ?? {}).length,
  };
}
