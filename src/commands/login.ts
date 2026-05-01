// src/commands/login.ts
import { captureSession } from '../auth/browser-capture';
import type { Config } from '../config/load';

export interface LoginOptions {
  config: Config;
  profileDir?: string;
}

export interface LoginResult {
  status: 'ok';
  account: { upn?: string; oid?: string; tid?: string };
}

export async function runLogin(opts: LoginOptions): Promise<LoginResult> {
  const session = await captureSession({
    loginTimeoutMs: opts.config.loginTimeoutMs,
    chromeChannel: opts.config.chromeChannel,
    profileDir: opts.profileDir,
  });
  return {
    status: 'ok',
    account: session.account ?? {},
  };
}
