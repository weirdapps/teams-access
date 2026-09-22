// test_scripts/auth/lock.test.ts
//
// teams-cli drives a 698 MB persistent Chromium profile and, unlike
// outlook-cli, took no lock while doing it. Two overlapping invocations (the
// 15-minute token sync and an interactive command, or the MCP bridge) both open
// the same profile, and Chromium's own profile lock makes the loser hang until
// Playwright gives up. Observed 2026-09-22 as
// `browserType.launchPersistentContext: Timeout 180000ms exceeded`.
//
// Ported from outlook-access/src/auth/lock.ts, which has had this since the
// start. Advisory PID lock: O_EXCL create, PID as the content, stale entries
// reclaimed when the owner is gone.
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { acquireLock } from '../../src/auth/lock';

let dir: string;
let lockPath: string;

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-lock-'));
  lockPath = path.join(dir, 'sub', '.browser.lock');
});

afterEach(() => {
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('acquireLock', () => {
  it('creates the lock and its parent, holding this process PID', async () => {
    const release = await acquireLock(lockPath);
    expect(fs.readFileSync(lockPath, 'utf8').trim()).toBe(String(process.pid));
    expect(fs.statSync(path.dirname(lockPath)).mode & 0o777).toBe(0o700);
    await release();
  });

  it('refuses a second holder while the first is alive', async () => {
    const release = await acquireLock(lockPath);
    await expect(acquireLock(lockPath)).rejects.toThrow(
      /another teams-cli instance holds the lock/,
    );
    await release();
  });

  it('reclaims a lock whose owner is gone', async () => {
    fs.mkdirSync(path.dirname(lockPath), { recursive: true });
    // PID 2^22 is above the macOS and Linux defaults, so nothing owns it.
    fs.writeFileSync(lockPath, '4194304\n');
    const release = await acquireLock(lockPath);
    expect(fs.readFileSync(lockPath, 'utf8').trim()).toBe(String(process.pid));
    await release();
  });

  it('reclaims a lock whose content is not a PID', async () => {
    fs.mkdirSync(path.dirname(lockPath), { recursive: true });
    fs.writeFileSync(lockPath, '');
    const release = await acquireLock(lockPath);
    await release();
    expect(fs.existsSync(lockPath)).toBe(false);
  });

  it('treats a PID it may not signal as alive rather than stealing the lock', async () => {
    // PID 1 is launchd/init, owned by root, so process.kill(1, 0) raises EPERM
    // for an ordinary user. Guessing "not mine, therefore dead" would hand the
    // profile to a second Chromium while the first is still inside it.
    fs.mkdirSync(path.dirname(lockPath), { recursive: true });
    fs.writeFileSync(lockPath, '1\n');
    await expect(acquireLock(lockPath)).rejects.toThrow(
      /another teams-cli instance holds the lock/,
    );
    expect(fs.readFileSync(lockPath, 'utf8').trim()).toBe('1');
  });

  it('reclaims a lock holding a nonsense PID', async () => {
    fs.mkdirSync(path.dirname(lockPath), { recursive: true });
    fs.writeFileSync(lockPath, 'not-a-pid\n');
    const release = await acquireLock(lockPath);
    expect(fs.readFileSync(lockPath, 'utf8').trim()).toBe(String(process.pid));
    await release();
  });

  it('survives the lock vanishing under it, since release must not throw', async () => {
    const release = await acquireLock(lockPath);
    fs.unlinkSync(lockPath);
    await expect(release()).resolves.toBeUndefined();
  });

  it('releases idempotently, so a finally block cannot throw', async () => {
    const release = await acquireLock(lockPath);
    await release();
    await expect(release()).resolves.toBeUndefined();
    expect(fs.existsSync(lockPath)).toBe(false);
  });

  it('frees the lock for the next caller once released', async () => {
    await (
      await acquireLock(lockPath)
    )();
    const second = await acquireLock(lockPath);
    await second();
  });
});
