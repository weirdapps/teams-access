// test_scripts/config/load.test.ts
import { describe, it, expect, beforeEach } from 'vitest';
import { loadConfig } from '../../src/config/load';

describe('loadConfig', () => {
  beforeEach(() => {
    delete process.env.TEAMS_CLI_HTTP_TIMEOUT_MS;
    delete process.env.TEAMS_CLI_LOGIN_TIMEOUT_MS;
    delete process.env.TEAMS_CLI_CHROME_CHANNEL;
  });

  it('returns defaults when no flags or env vars set', () => {
    const cfg = loadConfig({});
    expect(cfg.httpTimeoutMs).toBe(30_000);
    expect(cfg.loginTimeoutMs).toBe(300_000);
    expect(cfg.chromeChannel).toBe('chrome');
  });

  it('env vars override defaults', () => {
    process.env.TEAMS_CLI_HTTP_TIMEOUT_MS = '5000';
    process.env.TEAMS_CLI_CHROME_CHANNEL = 'msedge';
    const cfg = loadConfig({});
    expect(cfg.httpTimeoutMs).toBe(5000);
    expect(cfg.chromeChannel).toBe('msedge');
  });

  it('CLI flags override env vars', () => {
    process.env.TEAMS_CLI_HTTP_TIMEOUT_MS = '5000';
    const cfg = loadConfig({ timeoutMs: 9999 });
    expect(cfg.httpTimeoutMs).toBe(9999);
  });

  it('throws ExitWithCode(Config) on malformed env timeout', () => {
    process.env.TEAMS_CLI_HTTP_TIMEOUT_MS = 'not-a-number';
    expect(() => loadConfig({})).toThrow(/timeout/i);
  });
});
