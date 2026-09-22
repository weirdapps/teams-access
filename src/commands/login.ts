// src/commands/login.ts
import { captureSession } from '../auth/browser-capture';
import { acquireLock } from '../auth/lock';
import { defaultLockPath } from './auth-renew';
import type { Config } from '../config/load';

export interface LoginOptions {
  config: Config;
  profileDir?: string;
  diagnosticExtraMs?: number;
  minAudiences?: number;
}

export interface LoginResult {
  status: 'ok';
  account: { upn?: string; oid?: string; tid?: string };
}

export async function runLogin(opts: LoginOptions): Promise<LoginResult> {
  // Same profile, same lock as auth-renew. A login started while the 15-minute
  // token sync is mid-renew would otherwise sit on Chromium's own profile lock
  // with no explanation, which is the worst place to be confused: the human is
  // at the keyboard waiting for a window that will never open.
  const release = await acquireLock(defaultLockPath());
  try {
    const session = await captureSession({
      loginTimeoutMs: opts.config.loginTimeoutMs,
      chromeChannel: opts.config.chromeChannel,
      profileDir: opts.profileDir,
      diagnosticExtraMs: opts.diagnosticExtraMs,
      minAudiences: opts.minAudiences,
    });
    return {
      status: 'ok',
      account: session.account ?? {},
    };
  } finally {
    await release().catch(() => undefined);
  }
}
