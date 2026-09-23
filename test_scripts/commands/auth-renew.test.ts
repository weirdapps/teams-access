// test_scripts/commands/auth-renew.test.ts
//
// Headless renewal fails for two very different reasons and used to report both
// the same way. Every error out of captureSession() became AuthRequired (4),
// which reads as "the device-trust cookie died, go and log in at the keyboard".
//
// Measured 2026-09-22 on this Mac: 14 of 132 runs failed, and the four with a
// trace were `net::ERR_INTERNET_DISCONNECTED`, `net::ERR_CERT_AUTHORITY_INVALID`
// twice (a captive portal intercepting TLS after a flight), and
// `browserType.launchPersistentContext: Timeout 180000ms exceeded`. None was an
// auth failure, no login would have fixed any of them, and they cost six alert
// emails telling the user to go and run `teams-cli login`.
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { classifyCaptureFailure } from '../../src/commands/auth-renew';
import { ExitCode, type ExitWithCode } from '../../src/util/exit-codes';

describe('classifyCaptureFailure', () => {
  it('maps Chromium transport errors to Upstream, not AuthRequired', () => {
    for (const message of [
      'page.goto: net::ERR_INTERNET_DISCONNECTED at https://teams.cloud.microsoft/',
      'page.goto: net::ERR_CERT_AUTHORITY_INVALID at https://teams.cloud.microsoft/',
      'page.goto: net::ERR_NAME_NOT_RESOLVED at https://teams.cloud.microsoft/',
      'page.goto: net::ERR_PROXY_CONNECTION_FAILED at https://teams.cloud.microsoft/',
    ]) {
      expect(classifyCaptureFailure(message)).toBe(ExitCode.Upstream);
    }
  });

  it('maps a browser that never started to Upstream', () => {
    expect(
      classifyCaptureFailure('browserType.launchPersistentContext: Timeout 180000ms exceeded.'),
    ).toBe(ExitCode.Upstream);
    expect(classifyCaptureFailure('browserType.launch: Executable doesn’t exist')).toBe(
      ExitCode.Upstream,
    );
  });

  it('keeps everything else on AuthRequired', () => {
    // The real signal: Chromium reached Microsoft and no token came back.
    for (const message of [
      'Timed out waiting for a Teams bearer token',
      'page.goto: Timeout 30000ms exceeded',
      '',
    ]) {
      expect(classifyCaptureFailure(message)).toBe(ExitCode.AuthRequired);
    }
  });
});

vi.mock('../../src/auth/browser-capture', () => ({
  captureSession: vi.fn(),
}));
vi.mock('../../src/session/store', () => ({
  readSession: vi.fn(),
  writeSession: vi.fn(),
}));

describe('runAuthRenew exit codes', () => {
  // HOME is redirected because runAuthRenew now takes a real lock under
  // $HOME/.teams-cli. Without this these tests reach for the live lock and fail
  // with auth_renew_locked whenever the 15-minute token sync happens to be
  // mid-renew, which is both flaky and a test interfering with production.
  let home: string;
  let originalHome: string | undefined;

  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    originalHome = process.env.HOME;
    home = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-renew-rc-'));
    process.env.HOME = home;
  });

  afterEach(() => {
    if (originalHome === undefined) delete process.env.HOME;
    else process.env.HOME = originalHome;
    fs.rmSync(home, { recursive: true, force: true });
  });

  async function renewRejectingWith(err: Error) {
    const { captureSession } = await import('../../src/auth/browser-capture');
    const { readSession } = await import('../../src/session/store');
    vi.mocked(readSession).mockReturnValue({ tokens: {} } as never);
    vi.mocked(captureSession).mockRejectedValue(err);
    const { runAuthRenew } = await import('../../src/commands/auth-renew');
    return runAuthRenew().then(
      () => null,
      (e: unknown) => e as ExitWithCode,
    );
  }

  it('exits Upstream when the network is down, so the caller does not demand a login', async () => {
    const thrown = await renewRejectingWith(
      new Error('page.goto: net::ERR_INTERNET_DISCONNECTED at https://teams.cloud.microsoft/'),
    );
    // Not toBeInstanceOf: vi.resetModules() gives the dynamic import its own
    // copy of the ExitWithCode class, so identity never matches the static one.
    expect(thrown?.name).toBe('ExitWithCode');
    expect(thrown?.code).toBe(ExitCode.Upstream);
    expect(thrown?.payload.code).toBe('auth_renew_unreachable');
  });

  it('exits Upstream when Chrome never launched', async () => {
    const thrown = await renewRejectingWith(
      new Error('browserType.launchPersistentContext: Timeout 180000ms exceeded.'),
    );
    expect(thrown?.code).toBe(ExitCode.Upstream);
  });

  it('still exits AuthRequired when the capture reached Microsoft and got nothing', async () => {
    const thrown = await renewRejectingWith(
      new Error('Timed out waiting for a Teams bearer token'),
    );
    expect(thrown?.code).toBe(ExitCode.AuthRequired);
    expect(thrown?.payload.code).toBe('auth_renew_failed');
  });
});

// A 698 MB persistent profile driven by two processes at once is how
// `launchPersistentContext: Timeout 180000ms exceeded` happens. HOME is
// redirected so these never touch the real ~/.teams-cli/.browser.lock, which the
// live 15-minute token sync uses.
describe('runAuthRenew holds the browser lock', () => {
  let home: string;
  let originalHome: string | undefined;

  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    originalHome = process.env.HOME;
    home = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-renew-'));
    process.env.HOME = home;
  });

  afterEach(() => {
    if (originalHome === undefined) delete process.env.HOME;
    else process.env.HOME = originalHome;
    fs.rmSync(home, { recursive: true, force: true });
  });

  const lockPath = () => path.join(home, '.teams-cli', '.browser.lock');

  async function renew(captureImpl: () => Promise<unknown>) {
    const { captureSession } = await import('../../src/auth/browser-capture');
    const { readSession } = await import('../../src/session/store');
    vi.mocked(readSession).mockReturnValue({ tokens: {} } as never);
    vi.mocked(captureSession).mockImplementation(captureImpl as never);
    const { runAuthRenew } = await import('../../src/commands/auth-renew');
    return runAuthRenew().then(
      (r) => r,
      (e: unknown) => e as ExitWithCode,
    );
  }

  it('takes the lock while capturing and frees it afterwards', async () => {
    let heldDuringCapture = false;
    await renew(async () => {
      heldDuringCapture = fs.existsSync(lockPath());
      throw new Error('Timed out waiting for a Teams bearer token');
    });
    expect(heldDuringCapture).toBe(true);
    expect(fs.existsSync(lockPath())).toBe(false);
  });

  it('frees the lock even when the capture throws, so one blip cannot wedge every later run', async () => {
    await renew(async () => {
      throw new Error(
        'page.goto: net::ERR_INTERNET_DISCONNECTED at https://teams.cloud.microsoft/',
      );
    });
    expect(fs.existsSync(lockPath())).toBe(false);
  });

  it('returns ok and reports every audience it captured', async () => {
    const { writeSession } = await import('../../src/session/store');
    const tokens = {
      'https://graph.microsoft.com': {},
      'https://chatsvcagg.teams.microsoft.com': {},
      'https://outlook.office.com/': {},
    };
    const result = (await renew(async () => ({
      tokens,
      account: { upn: 'user@example.com' },
    }))) as { status: string; audiencesCaptured: number; account: { upn?: string } };
    expect(result.status).toBe('ok');
    expect(result.audiencesCaptured).toBe(3);
    expect(result.account.upn).toBe('user@example.com');
    expect(vi.mocked(writeSession)).toHaveBeenCalledOnce();
    expect(fs.existsSync(lockPath())).toBe(false);
  });

  it('persists the session before rejecting an incomplete capture', async () => {
    // A Graph token refreshed for list-teams is worth keeping even when
    // chatsvcagg is missing, so the write must happen before the audience gate.
    const { writeSession } = await import('../../src/session/store');
    const thrown = (await renew(async () => ({
      tokens: { 'https://graph.microsoft.com': {} },
    }))) as ExitWithCode;
    expect(thrown?.code).toBe(ExitCode.AuthRequired);
    expect(thrown?.payload.code).toBe('auth_renew_incomplete');
    expect(thrown?.payload.missingAudiences).toEqual(['https://chatsvcagg.teams.microsoft.com']);
    expect(vi.mocked(writeSession)).toHaveBeenCalledOnce();
    expect(fs.existsSync(lockPath())).toBe(false);
  });

  it('never takes the lock when there is no session to renew', async () => {
    const { captureSession } = await import('../../src/auth/browser-capture');
    const { readSession } = await import('../../src/session/store');
    vi.mocked(readSession).mockReturnValue(null);
    const { runAuthRenew } = await import('../../src/commands/auth-renew');
    const thrown = await runAuthRenew().then(
      () => null,
      (e: unknown) => e as ExitWithCode,
    );
    expect(thrown?.payload.code).toBe('auth_no_reauth');
    expect(fs.existsSync(lockPath())).toBe(false);
    expect(vi.mocked(captureSession)).not.toHaveBeenCalled();
  });

  it('exits Upstream, not AuthRequired, when another instance holds it', async () => {
    fs.mkdirSync(path.dirname(lockPath()), { recursive: true, mode: 0o700 });
    fs.writeFileSync(lockPath(), `${process.pid}\n`);
    const thrown = (await renew(async () => ({ tokens: {} }))) as ExitWithCode;
    expect(thrown?.code).toBe(ExitCode.Upstream);
    expect(thrown?.payload.code).toBe('auth_renew_locked');
    // The live holder's lock must survive: releasing it here would hand the
    // profile to a third process while the real owner is still inside it.
    expect(fs.existsSync(lockPath())).toBe(true);
  });
});

// A renewal that hands back a bearer about to die is worse than one that fails:
// the producer pushes it, the VPS reads it as fresh, and it expires before the
// next push lands. Measured 2026-09-22/23, 4 of 4 teams-session outages on the
// VPS followed a push whose shortest required audience had 250-640s left. The
// capture now evicts near-expiry tokens so the SPA mints fresh ones, and the
// command refuses to call a renewal ok while a required audience is still
// under the floor the producer's dead-man switch uses (900s).
describe('runAuthRenew refuses a stale capture', () => {
  let home: string;
  let originalHome: string | undefined;

  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    originalHome = process.env.HOME;
    home = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-renew-stale-'));
    process.env.HOME = home;
  });

  afterEach(() => {
    if (originalHome === undefined) delete process.env.HOME;
    else process.env.HOME = originalHome;
    fs.rmSync(home, { recursive: true, force: true });
  });

  const nowS = () => Math.floor(Date.now() / 1000);

  function jwt(aud: string, exp: number): string {
    const b64 = (o: object) => Buffer.from(JSON.stringify(o)).toString('base64url');
    return `${b64({ alg: 'none' })}.${b64({ aud, exp })}.sig`;
  }

  async function renewCapturing(tokens: Record<string, unknown>) {
    const { captureSession } = await import('../../src/auth/browser-capture');
    const { readSession } = await import('../../src/session/store');
    vi.mocked(readSession).mockReturnValue({ tokens: {} } as never);
    vi.mocked(captureSession).mockResolvedValue({ tokens, account: {} } as never);
    const { runAuthRenew } = await import('../../src/commands/auth-renew');
    const outcome = await runAuthRenew().then(
      (r) => r,
      (e: unknown) => e as ExitWithCode,
    );
    return { outcome, captureSession: vi.mocked(captureSession) };
  }

  it('asks the capture to evict near-expiry tokens and to capture nothing under the floor', async () => {
    const { captureSession } = await renewCapturing({
      'https://graph.microsoft.com': { exp: nowS() + 4000 },
      'https://chatsvcagg.teams.microsoft.com': { exp: nowS() + 4000 },
    });
    const opts = captureSession.mock.calls[0][0];
    expect(opts.headless).toBe(true);
    expect(opts.evictBelowTtlS).toBe(1500);
    expect(opts.captureFloorTtlS).toBe(900);
    // Evicting above the floor is what leaves room for a missed push: a token
    // kept at 1499s still outlives one 900s push interval.
    expect(opts.evictBelowTtlS!).toBeGreaterThan(opts.captureFloorTtlS!);
  });

  it('exits AuthRequired when a required audience comes back under the floor', async () => {
    const graph = jwt('https://graph.microsoft.com', nowS() + 300);
    const { outcome } = await renewCapturing({
      'https://graph.microsoft.com': { bearerToken: graph, exp: nowS() + 300 },
      'https://chatsvcagg.teams.microsoft.com': { exp: nowS() + 4000 },
    });
    const thrown = outcome as ExitWithCode;
    expect(thrown?.code).toBe(ExitCode.AuthRequired);
    expect(thrown?.payload.code).toBe('auth_renew_stale');
    const stale = thrown?.payload.staleAudiences as Array<{ aud: string; ttlSeconds: number }>;
    expect(stale.map((s) => s.aud)).toEqual(['https://graph.microsoft.com']);
    expect(stale[0].ttlSeconds).toBeGreaterThan(290);
    expect(stale[0].ttlSeconds).toBeLessThanOrEqual(300);
    // Named and dated, never shown.
    expect(JSON.stringify(thrown?.payload)).not.toContain(graph);
  });

  it('reads the expiry off the JWT when the entry does not carry one', async () => {
    const { outcome } = await renewCapturing({
      'https://graph.microsoft.com': { exp: nowS() + 4000 },
      'https://chatsvcagg.teams.microsoft.com': {
        bearerToken: jwt('https://chatsvcagg.teams.microsoft.com', nowS() + 120),
      },
    });
    expect((outcome as ExitWithCode)?.payload.code).toBe('auth_renew_stale');
  });

  it('still persists what it captured before refusing, as the incomplete gate does', async () => {
    const { writeSession } = await import('../../src/session/store');
    await renewCapturing({
      'https://graph.microsoft.com': { exp: nowS() + 100 },
      'https://chatsvcagg.teams.microsoft.com': { exp: nowS() + 4000 },
    });
    expect(vi.mocked(writeSession)).toHaveBeenCalledOnce();
  });

  it('passes a fresh capture, and one whose expiry it cannot read', async () => {
    const { outcome } = await renewCapturing({
      'https://graph.microsoft.com': { exp: nowS() + 3000 },
      'https://chatsvcagg.teams.microsoft.com': {},
    });
    expect((outcome as { status: string }).status).toBe('ok');
  });

  it('does not call an unreadable bearer stale on a guess', async () => {
    const { outcome } = await renewCapturing({
      'https://graph.microsoft.com': { bearerToken: 'not-a-jwt' },
      'https://chatsvcagg.teams.microsoft.com': { exp: nowS() + 3000 },
    });
    expect((outcome as { status: string }).status).toBe('ok');
  });
});
