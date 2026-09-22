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
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
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
