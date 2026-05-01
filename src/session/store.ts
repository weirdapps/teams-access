// src/session/store.ts
import {
  mkdirSync,
  writeFileSync,
  readFileSync,
  existsSync,
  chmodSync,
  renameSync,
  statSync,
} from 'node:fs';
import { join } from 'node:path';
import { homedir, platform } from 'node:os';
import { randomBytes } from 'node:crypto';

export interface SessionCookie {
  name: string;
  value: string;
  domain: string;
  path?: string;
  expires?: number;
  httpOnly?: boolean;
  secure?: boolean;
  sameSite?: 'Strict' | 'Lax' | 'None';
}

export interface Session {
  bearerToken: string;
  cookies: SessionCookie[];
  capturedAt: string;
  account?: { upn?: string; oid?: string; tid?: string };
}

const HOME = () => process.env.HOME ?? homedir();
const DIR = () => join(HOME(), '.teams-cli');
export const sessionPath = (): string => join(DIR(), 'session.json');

function ensureDir(): void {
  const dir = DIR();
  if (!existsSync(dir)) {
    mkdirSync(dir, { recursive: true, mode: 0o700 });
  } else if (platform() !== 'win32') {
    const current = statSync(dir).mode & 0o777;
    if (current !== 0o700) chmodSync(dir, 0o700);
  }
}

export function writeSession(session: Session): void {
  ensureDir();
  const finalPath = sessionPath();
  const tmpPath = `${finalPath}.${randomBytes(6).toString('hex')}.tmp`;
  writeFileSync(tmpPath, JSON.stringify(session, null, 2), { mode: 0o600 });
  renameSync(tmpPath, finalPath);
  if (platform() !== 'win32') {
    chmodSync(finalPath, 0o600);
  }
}

export function readSession(): Session | null {
  const path = sessionPath();
  if (!existsSync(path)) return null;
  const raw = readFileSync(path, 'utf8');
  return JSON.parse(raw) as Session;
}
