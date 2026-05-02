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

/** Default headless renewal timeout. Headless mode drives 4 navigations
 *  (teams root → teams /v2/?view=Chat → outlook → office.com) to provoke
 *  Graph + chatsvcagg + outlook + presence audience captures; needs ~60s
 *  headroom (chat panel init alone takes ~7s). */
const DEFAULT_RENEW_TIMEOUT_MS = 90_000;
/** After Graph token captured, keep listening this long for additional
 *  audiences to land (chatsvcagg, ic3, etc). */
const DEFAULT_DIAGNOSTIC_EXTRA_MS = 15_000;

/**
 * Audiences that downstream commands actually need. Renewal is considered
 * incomplete (and exits AuthRequired) if any of these are missing — this
 * prevents the silent-success failure mode where renew "worked" but
 * list-messages still 401s on chatsvcagg.
 *
 * Keep in sync with what `health-check` probes: anything probed there should
 * be required here, otherwise health-check will still detect breakage that
 * renewal claimed to fix.
 */
const REQUIRED_AUDIENCES = [
  'https://graph.microsoft.com',
  'https://chatsvcagg.teams.microsoft.com',
] as const;

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

  // Persist the freshly-captured session BEFORE the audience check — even an
  // incomplete capture is more useful on disk than nothing (e.g. Graph token
  // refreshed for `list-teams` while chatsvcagg is still missing).
  writeSession(captured);

  // Strict validation: renew is only "ok" if every audience downstream
  // commands need was captured. Without this, headless flow drift (e.g. Teams
  // SPA redesign that stops loading chatsvcagg on the chat URL) silently
  // produces a session that looks fine but fails on first real use.
  const capturedAudiences = Object.keys(captured.tokens ?? {});
  const missing = REQUIRED_AUDIENCES.filter(aud => !capturedAudiences.includes(aud));
  if (missing.length > 0) {
    throw new ExitWithCode(ExitCode.AuthRequired, {
      code: 'auth_renew_incomplete',
      message:
        `Headless renewal captured ${capturedAudiences.length} audiences but ` +
        `${missing.length} required audience(s) are missing: ${missing.join(', ')}. ` +
        `Run \`teams-cli login\` interactively (open the Chat tab in the diagnostic window).`,
      capturedAudiences,
      missingAudiences: missing,
    });
  }

  return {
    status: 'ok',
    sessionFile: sessionPath,
    account: captured.account ?? {},
    durationMs: Date.now() - t0,
    audiencesCaptured: capturedAudiences.length,
  };
}
