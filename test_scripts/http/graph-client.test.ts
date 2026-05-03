// test_scripts/http/graph-client.test.ts
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { GraphClient } from '../../src/http/graph-client';
import { AuthRequiredError, UpstreamError } from '../../src/http/errors';

const SESSION = {
  bearerToken: 'eyJ.X.Y',
  cookies: [],
  capturedAt: '2026-05-01T10:00:00Z',
};

function mockFetch(impl: typeof globalThis.fetch): typeof globalThis.fetch {
  globalThis.fetch = impl as typeof globalThis.fetch;
  return impl;
}

describe('GraphClient', () => {
  let originalFetch: typeof globalThis.fetch;
  beforeEach(() => {
    originalFetch = globalThis.fetch;
  });
  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('attaches Bearer header from the session', async () => {
    let captured: Headers | undefined;
    mockFetch(async (_url, init) => {
      captured = new Headers(init?.headers);
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    });
    const c = new GraphClient(SESSION, { httpTimeoutMs: 1000 });
    await c.get('/me');
    expect(captured?.get('Authorization')).toBe('Bearer eyJ.X.Y');
  });

  it('throws AuthRequiredError on 401', async () => {
    mockFetch(
      async () =>
        new Response(
          JSON.stringify({ error: { code: 'InvalidAuthenticationToken', message: 'expired' } }),
          { status: 401 },
        ),
    );
    const c = new GraphClient(SESSION, { httpTimeoutMs: 1000 });
    await expect(c.get('/me')).rejects.toBeInstanceOf(AuthRequiredError);
  });

  it('throws UpstreamError on 5xx (after retries exhausted)', async () => {
    let calls = 0;
    mockFetch(async () => {
      calls++;
      return new Response(
        JSON.stringify({ error: { code: 'ServiceUnavailable', message: 'busy' } }),
        { status: 503 },
      );
    });
    const c = new GraphClient(SESSION, { httpTimeoutMs: 1000, retries: 2, retryBaseMs: 1 });
    await expect(c.get('/me')).rejects.toBeInstanceOf(UpstreamError);
    expect(calls).toBe(3);
  });

  it('returns parsed JSON on 200', async () => {
    mockFetch(async () => new Response(JSON.stringify({ value: [{ id: 'a' }] }), { status: 200 }));
    const c = new GraphClient(SESSION, { httpTimeoutMs: 1000 });
    const r = await c.get<{ value: { id: string }[] }>('/teams');
    expect(r.value[0].id).toBe('a');
  });

  it('paginate() walks @odata.nextLink until exhausted or cap reached', async () => {
    let call = 0;
    mockFetch(async () => {
      call++;
      if (call === 1) {
        return new Response(
          JSON.stringify({
            value: [{ id: '1' }, { id: '2' }],
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=abc',
          }),
          { status: 200 },
        );
      }
      return new Response(JSON.stringify({ value: [{ id: '3' }] }), { status: 200 });
    });
    const c = new GraphClient(SESSION, { httpTimeoutMs: 1000 });
    const all = await c.paginate<{ id: string }>('/me/chats', { maxResults: 100 });
    expect(all.items.map((i) => i.id)).toEqual(['1', '2', '3']);
    expect(all.cappedEarly).toBe(false);
  });

  it('paginate() respects maxResults cap', async () => {
    mockFetch(
      async () =>
        new Response(
          JSON.stringify({
            value: Array.from({ length: 50 }, (_, i) => ({ id: String(i) })),
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/x?next',
          }),
          { status: 200 },
        ),
    );
    const c = new GraphClient(SESSION, { httpTimeoutMs: 1000 });
    const all = await c.paginate<{ id: string }>('/x', { maxResults: 30 });
    expect(all.items.length).toBe(30);
    expect(all.cappedEarly).toBe(true);
  });
});
