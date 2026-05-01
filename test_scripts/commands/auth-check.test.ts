// test_scripts/commands/auth-check.test.ts
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { runAuthCheck } from '../../src/commands/auth-check';
import { ExitWithCode } from '../../src/util/exit-codes';

const SESSION = {
  bearerToken: 'eyJ.X.Y',
  cookies: [],
  capturedAt: '2026-05-01T10:00:00Z',
};

describe('auth-check command', () => {
  let originalFetch: typeof globalThis.fetch;
  beforeEach(() => { originalFetch = globalThis.fetch; });
  afterEach(() => { globalThis.fetch = originalFetch; });

  it('returns ok status + account info when GET /me succeeds', async () => {
    globalThis.fetch = (async () =>
      new Response(JSON.stringify({
        id: 'oid-1',
        userPrincipalName: 'user@example.com',
        displayName: 'Test User',
      }), { status: 200 })) as typeof globalThis.fetch;
    const result = await runAuthCheck({ session: SESSION, httpTimeoutMs: 1000 });
    expect(result.status).toBe('ok');
    expect(result.account.upn).toBe('user@example.com');
  });

  it('throws ExitWithCode(AuthRequired) when GET /me returns 401', async () => {
    globalThis.fetch = (async () =>
      new Response(JSON.stringify({ error: { code: 'X', message: 'expired' } }), { status: 401 })) as typeof globalThis.fetch;
    await expect(runAuthCheck({ session: SESSION, httpTimeoutMs: 1000 })).rejects.toBeInstanceOf(ExitWithCode);
  });
});
