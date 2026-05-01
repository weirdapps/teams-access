// test_scripts/session/store.test.ts
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtempSync, rmSync, statSync, readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir, platform } from 'node:os';
import { writeSession, readSession, sessionPath, type Session } from '../../src/session/store';

const SAMPLE: Session = {
  bearerToken: 'eyJ.X.Y',
  cookies: [{ name: 'AAA', value: 'bbb', domain: '.teams.microsoft.com' }],
  capturedAt: '2026-05-01T10:00:00Z',
};

describe('session store', () => {
  let tmp: string;
  let oldHome: string | undefined;

  beforeEach(() => {
    tmp = mkdtempSync(join(tmpdir(), 'teams-cli-store-'));
    oldHome = process.env.HOME;
    process.env.HOME = tmp;
  });

  afterEach(() => {
    process.env.HOME = oldHome;
    rmSync(tmp, { recursive: true, force: true });
  });

  it('writes session.json under ~/.teams-cli/ with file mode 0600 (POSIX)', () => {
    writeSession(SAMPLE);
    const path = sessionPath();
    expect(existsSync(path)).toBe(true);
    if (platform() !== 'win32') {
      const mode = statSync(path).mode & 0o777;
      expect(mode).toBe(0o600);
    }
  });

  it('parent dir is mode 0700 (POSIX)', () => {
    writeSession(SAMPLE);
    const dir = sessionPath().replace(/\/session\.json$/, '');
    if (platform() !== 'win32') {
      const mode = statSync(dir).mode & 0o777;
      expect(mode).toBe(0o700);
    }
  });

  it('round-trips a session value', () => {
    writeSession(SAMPLE);
    const read = readSession();
    expect(read).toEqual(SAMPLE);
  });

  it('readSession returns null when file is absent', () => {
    expect(readSession()).toBeNull();
  });

  it('writes JSON content matching the input', () => {
    writeSession(SAMPLE);
    const raw = readFileSync(sessionPath(), 'utf8');
    expect(JSON.parse(raw)).toEqual(SAMPLE);
  });
});
