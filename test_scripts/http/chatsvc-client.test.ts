// test_scripts/http/chatsvc-client.test.ts
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { ChatsvcClient } from '../../src/http/chatsvc-client';
import { AuthRequiredError, GraphHttpError } from '../../src/http/errors';

const SESSION_NO_TOKEN = {
  bearerToken: 'eyJ.X.Y',
  cookies: [],
  capturedAt: '2026-05-01T10:00:00Z',
};
const SESSION_WITH_TOKEN = {
  bearerToken: 'eyJ.X.Y',
  cookies: [],
  capturedAt: '2026-05-01T10:00:00Z',
  region: { chatsvc: 'emea' },
  tokens: {
    'https://ic3.teams.office.com': {
      bearerToken: 'eyJ.IC3.TOKEN',
      capturedAt: '2026-05-01T10:00:00Z',
    },
  },
};

describe('ChatsvcClient', () => {
  let originalFetch: typeof globalThis.fetch;
  beforeEach(() => {
    originalFetch = globalThis.fetch;
  });
  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('throws AuthRequired when no IC3 token in session', async () => {
    const c = new ChatsvcClient(SESSION_NO_TOKEN, { httpTimeoutMs: 1000 });
    await expect(c.getChatMessages('19:abc@thread.v2')).rejects.toBeInstanceOf(AuthRequiredError);
  });

  it('uses IC3 token for Authorization', async () => {
    let captured: string | null = null;
    globalThis.fetch = (async (_url, init) => {
      captured = new Headers(init?.headers).get('Authorization');
      return new Response(JSON.stringify({ messages: [] }), { status: 200 });
    }) as typeof globalThis.fetch;
    const c = new ChatsvcClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    await c.getChatMessages('19:abc@thread.v2');
    expect(captured).toBe('Bearer eyJ.IC3.TOKEN');
  });

  it('builds the chatsvc URL with region + URL-encoded threadId', async () => {
    let url = '';
    globalThis.fetch = (async (u) => {
      url = String(u);
      return new Response(JSON.stringify({ messages: [] }), { status: 200 });
    }) as typeof globalThis.fetch;
    const c = new ChatsvcClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    await c.getChatMessages('19:abc@thread.v2', { pageSize: 5 });
    expect(url).toContain('/api/chatsvc/emea/v1/users/ME/conversations/');
    expect(url).toContain('19%3Aabc%40thread.v2/messages');
    expect(url).toContain('pageSize=5');
    expect(url).toContain('startTime=1');
    expect(url).toContain('view=msnp24Equivalent');
  });

  it('uses syncState URL verbatim when provided', async () => {
    let url = '';
    globalThis.fetch = (async (u) => {
      url = String(u);
      return new Response(JSON.stringify({ messages: [] }), { status: 200 });
    }) as typeof globalThis.fetch;
    const c = new ChatsvcClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    await c.getChatMessages('ignored', {
      syncState: 'https://teams.microsoft.com/api/chatsvc/emea/v1/foo?syncState=abc',
    });
    expect(url).toBe('https://teams.microsoft.com/api/chatsvc/emea/v1/foo?syncState=abc');
  });

  it('parses messages from response', async () => {
    globalThis.fetch = (async () =>
      new Response(
        JSON.stringify({
          messages: [
            { id: 'm1', content: 'hi', composetime: '2026-05-01T00:00:00Z' },
            { id: 'm2', content: 'hey', composetime: '2026-05-01T00:01:00Z' },
          ],
          tenantId: 'tid',
          _metadata: { syncState: 'https://...' },
        }),
        { status: 200 },
      )) as typeof globalThis.fetch;
    const c = new ChatsvcClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    const r = await c.getChatMessages('19:abc@thread.v2');
    expect(r.messages.length).toBe(2);
    expect(r.messages[0].id).toBe('m1');
    expect(r._metadata?.syncState).toBe('https://...');
  });

  it('throws GraphHttpError on 404', async () => {
    globalThis.fetch = (async () =>
      new Response('Not found', { status: 404 })) as typeof globalThis.fetch;
    const c = new ChatsvcClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    await expect(c.getChatMessages('19:nope@thread.v2')).rejects.toBeInstanceOf(GraphHttpError);
  });
});
