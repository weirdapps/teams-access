// test_scripts/http/chatsvcagg-client.test.ts
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { ChatsvcaggClient } from '../../src/http/chatsvcagg-client';
import { AuthRequiredError, GraphHttpError } from '../../src/http/errors';

const SESSION_WITHOUT_TOKEN = {
  bearerToken: 'eyJ.X.Y',
  cookies: [],
  capturedAt: '2026-05-01T10:00:00Z',
};

const SESSION_WITH_TOKEN = {
  bearerToken: 'eyJ.X.Y',
  cookies: [],
  capturedAt: '2026-05-01T10:00:00Z',
  region: { csa: 'emea' },
  tokens: {
    'https://chatsvcagg.teams.microsoft.com': {
      bearerToken: 'eyJ.CHATSVCAGG.TOKEN',
      capturedAt: '2026-05-01T10:00:00Z',
    },
  },
};

describe('ChatsvcaggClient', () => {
  let originalFetch: typeof globalThis.fetch;
  beforeEach(() => { originalFetch = globalThis.fetch; });
  afterEach(() => { globalThis.fetch = originalFetch; });

  it('throws AuthRequired when no chatsvcagg token in session', async () => {
    const c = new ChatsvcaggClient(SESSION_WITHOUT_TOKEN, { httpTimeoutMs: 1000 });
    await expect(c.listChats()).rejects.toBeInstanceOf(AuthRequiredError);
  });

  it('uses chatsvcagg token (not session.bearerToken) for Authorization', async () => {
    let captured: string | null = null;
    globalThis.fetch = (async (_url, init) => {
      captured = new Headers(init?.headers).get('Authorization');
      return new Response(JSON.stringify({ chats: [] }), { status: 200 });
    }) as typeof globalThis.fetch;
    const c = new ChatsvcaggClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    await c.listChats();
    expect(captured).toBe('Bearer eyJ.CHATSVCAGG.TOKEN');
  });

  it('listChats hits v1/updates with region from session', async () => {
    let capturedUrl = '';
    globalThis.fetch = (async (url) => {
      capturedUrl = String(url);
      return new Response(JSON.stringify({ chats: [{ id: 'c1' }, { id: 'c2' }], teams: [], channels: [] }), { status: 200 });
    }) as typeof globalThis.fetch;
    const c = new ChatsvcaggClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    const r = await c.listChats();
    expect(capturedUrl).toContain('/api/csa/emea/api/v1/teams/users/me/updates');
    expect(capturedUrl).toContain('isPrefetch=false');
    expect(r.chats.length).toBe(2);
    expect(r.chats[0].id).toBe('c1');
  });

  it('listChatsViaDiscover (legacy alias) still works', async () => {
    globalThis.fetch = (async () =>
      new Response(JSON.stringify({ chats: [{ id: 'c9' }] }), { status: 200 })) as typeof globalThis.fetch;
    const c = new ChatsvcaggClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    const r = await c.listChatsViaDiscover();
    expect(r.chats[0].id).toBe('c9');
  });

  it('listChannelPosts URL includes channelId, modality, pageSize, teamId', async () => {
    let capturedUrl = '';
    globalThis.fetch = (async (url) => {
      capturedUrl = String(url);
      return new Response(JSON.stringify({ posts: [], hasMore: false }), { status: 200 });
    }) as typeof globalThis.fetch;
    const c = new ChatsvcaggClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    await c.listChannelPosts(
      '19:abc@thread.tacv2',
      '19:general@thread.tacv2',
      { pageSize: 5 },
    );
    expect(capturedUrl).toContain('/containers/19%3Aabc%40thread.tacv2/posts');
    expect(capturedUrl).toContain('modality=post');
    expect(capturedUrl).toContain('pageSize=5');
    // teamId is the General channel id, encoded
    expect(decodeURIComponent(capturedUrl)).toContain('teamId=19:general@thread.tacv2');
  });

  it('throws GraphHttpError on 4xx', async () => {
    globalThis.fetch = (async () => new Response('Bad request.', { status: 400 })) as typeof globalThis.fetch;
    const c = new ChatsvcaggClient(SESSION_WITH_TOKEN, { httpTimeoutMs: 1000 });
    await expect(c.listChats()).rejects.toBeInstanceOf(GraphHttpError);
  });
});
