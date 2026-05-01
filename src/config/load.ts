// src/config/load.ts
import { ExitCode, ExitWithCode } from '../util/exit-codes';

export interface CliOverrides {
  timeoutMs?: number;
  loginTimeoutMs?: number;
  chromeChannel?: string;
}

export interface Config {
  httpTimeoutMs: number;
  loginTimeoutMs: number;
  chromeChannel: string;
}

const DEFAULTS: Config = {
  httpTimeoutMs: 30_000,
  loginTimeoutMs: 300_000,
  chromeChannel: 'chrome',
};

function parseIntStrict(name: string, raw: string): number {
  const n = Number(raw);
  if (!Number.isFinite(n) || n <= 0 || !Number.isInteger(n)) {
    throw new ExitWithCode(ExitCode.Config, {
      code: 'config_error',
      message: `Invalid ${name} (must be positive integer): "${raw}"`,
      setting: name,
    });
  }
  return n;
}

export function loadConfig(overrides: CliOverrides): Config {
  const httpEnv = process.env.TEAMS_CLI_HTTP_TIMEOUT_MS;
  const loginEnv = process.env.TEAMS_CLI_LOGIN_TIMEOUT_MS;
  const channelEnv = process.env.TEAMS_CLI_CHROME_CHANNEL;

  const httpTimeoutMs =
    overrides.timeoutMs ??
    (httpEnv ? parseIntStrict('TEAMS_CLI_HTTP_TIMEOUT_MS', httpEnv) : DEFAULTS.httpTimeoutMs);
  const loginTimeoutMs =
    overrides.loginTimeoutMs ??
    (loginEnv ? parseIntStrict('TEAMS_CLI_LOGIN_TIMEOUT_MS', loginEnv) : DEFAULTS.loginTimeoutMs);
  const chromeChannel = overrides.chromeChannel ?? channelEnv ?? DEFAULTS.chromeChannel;

  return { httpTimeoutMs, loginTimeoutMs, chromeChannel };
}
