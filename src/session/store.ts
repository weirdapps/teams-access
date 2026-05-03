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

export interface AudienceToken {
  bearerToken: string;
  appid?: string;
  scp?: string;
  exp?: number;
  capturedAt: string;
}

export interface SessionRegion {
  chatsvc?: string; // e.g. "emea"
  csa?: string; // e.g. "emea"
  mt?: string; // e.g. "emea-03"
  asyncgw?: string; // e.g. "eu-prod"
}

export interface Session {
  // Primary token = the Graph one. Kept top-level for backward compat with
  // the original GraphClient + auth-check + tests written before Path B.
  bearerToken: string;
  cookies: SessionCookie[];
  capturedAt: string;
  account?: { upn?: string; oid?: string; tid?: string };

  // Path B (amendment 1): per-audience Bearer tokens captured at login.
  // Keys are the JWT 'aud' claim string. Includes the Graph audience
  // (which duplicates bearerToken above, by design — keeps lookup uniform).
  tokens?: Record<string, AudienceToken>;

  // Tenant region detected from URL paths in the capture trace.
  region?: SessionRegion;

  // Per-team General-channel-id cache. Key: team UUID. Value: 19:...@thread.tacv2.
  // Populated lazily by list-channels / list-messages.
  generalChannelByTeamId?: Record<string, string>;
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
