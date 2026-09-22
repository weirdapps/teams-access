// src/auth/lock.ts
//
// Advisory PID lock for the browser-capture flow. Ported from
// outlook-access/src/auth/lock.ts, which has had one since the start.
//
// Why teams-cli needs it: the persisted profile under ~/.teams-cli is ~700 MB
// and Chromium takes its own exclusive lock on it. Two overlapping invocations
// (the 15-minute token sync, an interactive command, the MCP bridge) therefore
// leave the loser waiting on a profile it will never get, and Playwright
// reports that as `browserType.launchPersistentContext: Timeout 180000ms
// exceeded` three minutes later. Failing fast and saying why beats burning the
// caller's whole timeout budget.
//
// Creates the file with O_CREAT|O_EXCL|O_WRONLY ('wx') at mode 0o600, holding
// the owner's PID. On EEXIST the stored PID is probed: if that process is gone
// the lock is stale, and it is removed and retried exactly once.

import * as fs from 'node:fs';
import * as nodePath from 'node:path';

/**
 * Acquire an advisory lock at `path`. Returns a release function that is safe
 * to call more than once, so a `finally` block cannot throw.
 *
 * @throws Error('another teams-cli instance holds the lock: ' + path)
 *         when a live process already owns it.
 */
export async function acquireLock(path: string): Promise<() => Promise<void>> {
  // First run: ~/.teams-cli may not exist yet. 0o700 because a sibling of this
  // file is the session.
  fs.mkdirSync(nodePath.dirname(path), { recursive: true, mode: 0o700 });

  const tryOpen = (): number => fs.openSync(path, 'wx', 0o600);

  let fd: number;
  try {
    fd = tryOpen();
  } catch (err) {
    if (!isEexist(err)) throw err;

    const existingPid = readLockPid(path);
    if (existingPid !== null && isProcessAlive(existingPid)) {
      throw new Error('another teams-cli instance holds the lock: ' + path, { cause: err });
    }

    // Stale or unreadable. Remove and retry exactly once: a loop here would
    // race two reclaimers against each other forever.
    try {
      fs.unlinkSync(path);
    } catch (unlinkErr) {
      if (!isEnoent(unlinkErr)) throw unlinkErr;
    }

    try {
      fd = tryOpen();
    } catch (err2) {
      if (isEexist(err2)) {
        throw new Error('another teams-cli instance holds the lock: ' + path, { cause: err2 });
      }
      throw err2;
    }
  }

  try {
    fs.writeSync(fd, `${process.pid}\n`);
  } finally {
    try {
      fs.closeSync(fd);
    } catch {
      // Closing after a failed write must not mask the write's own error.
    }
  }

  let released = false;
  return async (): Promise<void> => {
    if (released) return;
    released = true;
    try {
      fs.unlinkSync(path);
    } catch (err) {
      if (!isEnoent(err)) throw err;
      // Already gone: idempotent success.
    }
  };
}

// ── Internals ────────────────────────────────────────────────────────────────

function readLockPid(path: string): number | null {
  let content: string;
  try {
    content = fs.readFileSync(path, 'utf8');
  } catch {
    return null;
  }

  const firstLine = content.split('\n')[0]?.trim() ?? '';
  if (firstLine.length === 0) return null;

  const n = Number.parseInt(firstLine, 10);
  if (!Number.isFinite(n) || n <= 0) return null;
  return n;
}

function isProcessAlive(pid: number): boolean {
  try {
    // Signal 0 is an existence and permission probe; it is never delivered.
    process.kill(pid, 0);
    return true;
  } catch (err) {
    const code = (err as NodeJS.ErrnoException)?.code;
    if (code === 'ESRCH') return false;
    // EPERM means it exists but belongs to another user, so it is alive.
    if (code === 'EPERM') return true;
    // Anything else: assume alive rather than clobber a real lock.
    return true;
  }
}

function isEexist(err: unknown): boolean {
  return (err as NodeJS.ErrnoException | undefined)?.code === 'EEXIST';
}

function isEnoent(err: unknown): boolean {
  return (err as NodeJS.ErrnoException | undefined)?.code === 'ENOENT';
}
