import { describe, it, expect, vi, beforeEach } from 'vitest';
import { runResolveMri } from '../src/commands/resolve-mri';
import type { Session } from '../src/session/store';

const fakeSession: Session = {
  bearerToken: 'fake',
  cookies: [],
  region: 'eu',
  expiresAt: Date.now() + 3600_000,
  tokens: { graph: 'fake-graph-token' },
} as unknown as Session;

describe('runResolveMri', () => {
  beforeEach(() => vi.restoreAllMocks());

  it('extracts aad-oid from MRI and returns user info', async () => {
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      new Response(
        JSON.stringify({
          id: '1234-aaaa',
          mail: 'user@example.com',
          displayName: 'User Name',
          userPrincipalName: 'user@example.com',
        }),
        { status: 200 }
      )
    );

    const result = await runResolveMri({
      session: fakeSession,
      httpTimeoutMs: 30000,
      mri: '8:orgid:1234-aaaa',
    });

    expect(result.email).toBe('user@example.com');
    expect(result.displayName).toBe('User Name');
    expect(result.id).toBe('1234-aaaa');
    expect(fetchSpy).toHaveBeenCalledWith(
      expect.stringContaining('/users/1234-aaaa'),
      expect.anything()
    );
  });

  it('rejects malformed MRI (no orgid prefix)', async () => {
    await expect(
      runResolveMri({ session: fakeSession, httpTimeoutMs: 30000, mri: 'not-a-mri' })
    ).rejects.toThrow(/invalid MRI/i);
  });

  it('returns null email when Graph response has no mail field', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      new Response(
        JSON.stringify({ id: 'x', displayName: 'Guest', userPrincipalName: 'guest#EXT#@ext.onmicrosoft.com' }),
        { status: 200 }
      )
    );

    const result = await runResolveMri({
      session: fakeSession,
      httpTimeoutMs: 30000,
      mri: '8:orgid:x',
    });

    expect(result.email).toBeNull();
    expect(result.displayName).toBe('Guest');
  });
});
