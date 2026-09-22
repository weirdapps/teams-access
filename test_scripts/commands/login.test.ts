// test_scripts/commands/login.test.ts
//
// login drives the same ~700 MB persistent profile as auth-renew, so it takes
// the same lock. A lock only auth-renew honoured would still let an interactive
// login collide with the 15-minute scheduled renew, and that is the worse case:
// the human is at the keyboard waiting for a window that never opens.
//
// HOME is redirected so none of this touches the live ~/.teams-cli/.browser.lock.
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

vi.mock('../../src/auth/browser-capture', () => ({
  captureSession: vi.fn(),
}));

const CONFIG = { loginTimeoutMs: 1000, chromeChannel: 'chrome' } as never;

describe('runLogin browser lock', () => {
  let home: string;
  let originalHome: string | undefined;

  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    originalHome = process.env.HOME;
    home = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-login-'));
    process.env.HOME = home;
  });

  afterEach(() => {
    if (originalHome === undefined) delete process.env.HOME;
    else process.env.HOME = originalHome;
    fs.rmSync(home, { recursive: true, force: true });
  });

  const lockPath = () => path.join(home, '.teams-cli', '.browser.lock');

  async function login(captureImpl: () => Promise<unknown>) {
    const { captureSession } = await import('../../src/auth/browser-capture');
    vi.mocked(captureSession).mockImplementation(captureImpl as never);
    const { runLogin } = await import('../../src/commands/login');
    return runLogin({ config: CONFIG });
  }

  it('holds the lock while capturing and frees it on success', async () => {
    let heldDuringCapture = false;
    const result = await login(async () => {
      heldDuringCapture = fs.existsSync(lockPath());
      return { account: { upn: 'user@example.com' } };
    });
    expect(heldDuringCapture).toBe(true);
    expect(result.status).toBe('ok');
    expect(result.account.upn).toBe('user@example.com');
    expect(fs.existsSync(lockPath())).toBe(false);
  });

  it('frees the lock when the capture throws', async () => {
    await expect(
      login(async () => {
        throw new Error('no Bearer token captured');
      }),
    ).rejects.toThrow(/no Bearer token captured/);
    expect(fs.existsSync(lockPath())).toBe(false);
  });

  it('refuses to start while a renew holds the profile', async () => {
    fs.mkdirSync(path.dirname(lockPath()), { recursive: true, mode: 0o700 });
    fs.writeFileSync(lockPath(), `${process.pid}\n`);
    const { captureSession } = await import('../../src/auth/browser-capture');
    await expect(login(async () => ({ account: {} }))).rejects.toThrow(
      /another teams-cli instance holds the lock/,
    );
    // The point of the lock: the second Chromium is never launched at all.
    expect(vi.mocked(captureSession)).not.toHaveBeenCalled();
  });

  it('reports no account when the capture returns one without it', async () => {
    const result = await login(async () => ({}));
    expect(result.account).toEqual({});
  });
});
