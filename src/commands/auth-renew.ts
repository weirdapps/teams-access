// src/commands/auth-renew.ts
//
// Silent (headless) bearer renewal. Uses the persisted Playwright profile
// (~/.teams-cli/playwright-profile/) to re-issue Teams Bearer tokens without
// opening a visible browser window.
//
// Works while the device-trust cookie (ESTSAUTHPERSISTENT, ~90 days) is
// alive. When that cookie expires or tenant policy forces re-MFA, this command fails
// and the caller must run `teams-cli login` interactively.
//
// Mirrors outlook-cli's auth-renew (commands/auth-renew.ts in outlook-access).

import { join } from 'node:path';
import { homedir } from 'node:os';

import { captureSession } from '../auth/browser-capture';
import { acquireLock } from '../auth/lock';
import { decodeJwt } from '../session/jwt';
import { readSession, writeSession, type AudienceToken, type Session } from '../session/store';
import { ExitCode, ExitWithCode, type ExitCodeValue } from '../util/exit-codes';

/** Default headless renewal timeout. Headless mode drives 4 navigations
 *  (teams root → teams /v2/?view=Chat → outlook → office.com) to provoke
 *  Graph + chatsvcagg + outlook + presence audience captures; needs ~60s
 *  headroom (chat panel init alone takes ~7s). */
const DEFAULT_RENEW_TIMEOUT_MS = 90_000;
/** Keep the headless browser open this long after the FIRST bearer so the
 *  slower surfaces finish and their audiences land. The M365 home
 *  (m365.cloud.microsoft) fires its graph.microsoft.com call ~20s after that
 *  nav — with the old 15s window the context closed first and Graph was
 *  perpetually missed (the multi-day "graph audience missing" outage). 40s
 *  clears M365's Graph fetch with margin; it's headless, so the extra wall-clock
 *  is invisible. */
const DEFAULT_DIAGNOSTIC_EXTRA_MS = 40_000;

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

/**
 * The shortest life a required audience may leave this command with, in
 * seconds. It is the producer's own gate: `sync-tokens-to-vps.sh` withholds its
 * dead-man ping below DEADMAN_MIN_TTL=900, because a bearer with less than one
 * 15-minute push interval left expires on the VPS before the next push lands.
 * Nothing under it is captured as fresh (wire or cache), and a required
 * audience still under it fails the renewal instead of passing as ok.
 */
export const RENEW_FLOOR_TTL_S = 900;

/**
 * Cached access tokens with less life than this are evicted before each page
 * boots, so the SPA mints fresh ones (see `evictNearExpiryInPage`). Above the
 * floor on purpose: a token kept at 1499s still outlives one push interval,
 * which is the margin for a push that fails on a bad network. Measured
 * 2026-09-23, the capture's required-audience TTL fell 5049, 4047, 3046, 2047,
 * 1046 across five runs: the same token, copied down. With this the fifth run
 * re-mints instead, so a ~5000s token serves four pushes and the shortest one
 * delivered stays above 1500s rather than reaching the 250-640s that failed.
 */
export const RENEW_EVICT_BELOW_TTL_S = 1500;

/** Seconds of life left on a session token, or undefined when it cannot be read. */
function secondsLeft(token: AudienceToken | undefined, nowS: number): number | undefined {
  if (!token) return undefined;
  let exp = typeof token.exp === 'number' ? token.exp : undefined;
  if (exp === undefined && typeof token.bearerToken === 'string') {
    try {
      const claimed = decodeJwt(token.bearerToken).exp;
      exp = typeof claimed === 'number' ? claimed : undefined;
    } catch {
      exp = undefined;
    }
  }
  return exp === undefined ? undefined : exp - nowS;
}

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

/**
 * Which exit code a captureSession() failure deserves.
 *
 * Renewal fails for two unrelated reasons and used to report both as
 * AuthRequired, which every caller reads as "the device-trust cookie died, go
 * and log in at the keyboard". Measured over 132 runs on 2026-09-22, 14 failed
 * and the four that left a trace were ERR_INTERNET_DISCONNECTED,
 * ERR_CERT_AUTHORITY_INVALID twice (a captive portal intercepting TLS) and a
 * launchPersistentContext timeout. No login would have fixed any of them, and
 * they cost six alert emails asking for one.
 *
 * Only what is unambiguously transport or launch moves to Upstream. Anything
 * else, including a navigation timeout that may well be a login page, keeps the
 * old AuthRequired: misreporting a real auth failure as transient would hide the
 * one case a human has to act on.
 */
export function classifyCaptureFailure(message: string): ExitCodeValue {
  if (/net::ERR_/.test(message)) return ExitCode.Upstream;
  if (/browserType\.launch/.test(message)) return ExitCode.Upstream;
  return ExitCode.AuthRequired;
}

function defaultProfileDir(): string {
  return join(process.env.HOME ?? homedir(), '.teams-cli', 'playwright-profile');
}

function defaultSessionPath(): string {
  return join(process.env.HOME ?? homedir(), '.teams-cli', 'session.json');
}

export function defaultLockPath(): string {
  return join(process.env.HOME ?? homedir(), '.teams-cli', '.browser.lock');
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

  // One process at a time in the persistent profile. Without this, an
  // overlapping invocation waits on Chromium's own profile lock and burns the
  // full three-minute Playwright budget before reporting a launch timeout.
  // Upstream rather than AuthRequired: contention is transient and the next
  // scheduled run gets it.
  let release: () => Promise<void>;
  try {
    release = await acquireLock(defaultLockPath());
  } catch (err) {
    throw new ExitWithCode(ExitCode.Upstream, {
      code: 'auth_renew_locked',
      message: `${err instanceof Error ? err.message : String(err)}. Transient; the next run retries.`,
    });
  }

  try {
    return await captureAndValidate(opts, sessionPath);
  } finally {
    // A lock this run cannot release would wedge every later run behind a dead
    // owner until the PID probe reclaims it. Never let it mask the real error.
    await release().catch(() => undefined);
  }
}

async function captureAndValidate(
  opts: AuthRenewOptions,
  sessionPath: string,
): Promise<AuthRenewResult> {
  const t0 = Date.now();

  let captured: Session;
  try {
    captured = await captureSession({
      loginTimeoutMs: opts.timeoutMs ?? DEFAULT_RENEW_TIMEOUT_MS,
      chromeChannel: opts.chromeChannel ?? 'chrome',
      profileDir: opts.profileDir ?? defaultProfileDir(),
      headless: true,
      diagnosticExtraMs: DEFAULT_DIAGNOSTIC_EXTRA_MS,
      evictBelowTtlS: RENEW_EVICT_BELOW_TTL_S,
      captureFloorTtlS: RENEW_FLOOR_TTL_S,
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    const code = classifyCaptureFailure(msg);
    if (code === ExitCode.Upstream) {
      // Never reached Microsoft, or Chrome never started. Transient: the next
      // scheduled run retries, and no human action would help.
      throw new ExitWithCode(ExitCode.Upstream, {
        code: 'auth_renew_unreachable',
        message: `Headless renewal could not reach Microsoft: ${msg}. Transient; no login needed.`,
      });
    }
    // Reached Microsoft and came back empty. Most likely cause:
    // ESTSAUTHPERSISTENT cookie expired or tenant policy forced re-MFA. Caller
    // must run interactive login.
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
  const missing = REQUIRED_AUDIENCES.filter((aud) => !capturedAudiences.includes(aud));
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

  // Present is not the same as alive. A required audience can only be under
  // the floor here if no fresh copy was captured and the prior session's own
  // token was carried forward (the capture refuses to take a new one that
  // short), so this is a renewal that did not renew. Exiting ok used to let the
  // producer push it anyway, and the VPS lost Teams 12-16 minutes later.
  // AuthRequired to match the incomplete gate above: both mean the caller's
  // credentials are about to stop working. The token itself is never reported.
  const nowS = Math.floor(Date.now() / 1000);
  const stale = REQUIRED_AUDIENCES.flatMap((aud) => {
    const ttl = secondsLeft(captured.tokens?.[aud], nowS);
    return ttl !== undefined && ttl < RENEW_FLOOR_TTL_S ? [{ aud, ttlSeconds: ttl }] : [];
  });
  if (stale.length > 0) {
    throw new ExitWithCode(ExitCode.AuthRequired, {
      code: 'auth_renew_stale',
      message:
        `Headless renewal could not mint a fresh token for ` +
        stale.map((s) => `${s.aud} (${s.ttlSeconds}s left)`).join(', ') +
        `: under the ${RENEW_FLOOR_TTL_S}s floor, a bearer expires on the VPS before the ` +
        `next 15-minute push. Near-expiry tokens are evicted so the page re-mints them, and ` +
        `this run captured no fresh copy. If this persists, run \`teams-cli login\`.`,
      staleAudiences: stale,
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
