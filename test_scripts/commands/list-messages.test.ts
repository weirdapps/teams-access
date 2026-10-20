// test_scripts/commands/list-messages.test.ts
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { runListMessages } from '../../src/commands/list-messages';
import { ExitCode, ExitWithCode } from '../../src/util/exit-codes';
import type { Session } from '../../src/session/store';

const SESSION = {
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
} as unknown as Session;

const CHAT = '19:abc@thread.v2';
const HOUR = 3600 * 1000;
const NOW = Date.parse('2026-10-20T12:00:00Z');
const LINK = 'https://teams.microsoft.com/api/chatsvc/emea/v1/users/ME/conversations/x/messages';

/** A chat of `count` messages `step` ms apart, newest first, in pages of `size`. */
function chatPages(count: number, size: number, step = HOUR) {
  const all = Array.from({ length: count }, (_, i) => ({
    id: String(NOW - i * step),
    composetime: new Date(NOW - i * step).toISOString(),
  }));
  const pages = [];
  for (let start = 0; start < all.length; start += size) {
    const last = start + size >= all.length;
    pages.push({
      messages: all.slice(start, start + size),
      _metadata: last ? {} : { backwardLink: `${LINK}?page=${pages.length + 1}` },
    });
  }
  return { all, pages };
}

describe('runListMessages chat paging', () => {
  let originalFetch: typeof globalThis.fetch;
  let requested: string[];

  function serve(pages: object[]) {
    requested = [];
    globalThis.fetch = (async (u) => {
      const url = String(u);
      requested.push(url);
      const m = /[?&]page=(\d+)/.exec(url);
      const page = m ? pages[Number(m[1])] : pages[0];
      return new Response(JSON.stringify(page ?? { messages: [] }), { status: 200 });
    }) as typeof globalThis.fetch;
  }

  beforeEach(() => {
    originalFetch = globalThis.fetch;
  });
  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('reads the newest page only by default, as before', async () => {
    const { pages } = chatPages(500, 200);
    serve(pages);
    const r = await runListMessages({
      session: SESSION,
      httpTimeoutMs: 1000,
      chat: CHAT,
      pageSize: 200,
    });
    expect(requested).toHaveLength(1);
    expect(r.messages).toHaveLength(200);
    expect(r.pagesRead).toBe(1);
  });

  it('recovers a three-day gap by paging back to the newest stored message', async () => {
    // A chat busy enough to fill a page of 200 in a day and a half: one message
    // every 11 minutes, so three days hold about 390 messages, two pages.
    const step = 11 * 60 * 1000;
    const { all, pages } = chatPages(600, 200, step);
    serve(pages);
    const stored = all[Math.round((72 * HOUR) / step)].composetime; // three days back
    const r = await runListMessages({
      session: SESSION,
      httpTimeoutMs: 1000,
      chat: CHAT,
      pageSize: 200,
      until: stored,
      maxPages: 10,
    });
    const ids = new Set(r.messages.map((m) => m.id));
    for (const m of all.filter((x) => Date.parse(x.composetime) > Date.parse(stored))) {
      expect(ids.has(m.id)).toBe(true);
    }
    expect(r.complete).toBe(true);
    expect(r.pagesRead).toBe(2);
  });

  it('stops after one page when that page already reaches the stored message', async () => {
    const { all, pages } = chatPages(600, 200);
    serve(pages);
    const r = await runListMessages({
      session: SESSION,
      httpTimeoutMs: 1000,
      chat: CHAT,
      pageSize: 200,
      until: all[30].composetime,
      maxPages: 10,
    });
    expect(requested).toHaveLength(1);
    expect(r.complete).toBe(true);
  });

  it('says complete:false when it stops at the page cap', async () => {
    const { all, pages } = chatPages(1000, 200);
    serve(pages);
    const r = await runListMessages({
      session: SESSION,
      httpTimeoutMs: 1000,
      chat: CHAT,
      pageSize: 200,
      until: all[999].composetime,
      maxPages: 3,
    });
    expect(requested).toHaveLength(3);
    expect(r.messages).toHaveLength(600);
    expect(r.complete).toBe(false);
  });

  it('is complete when the chat runs out before the time is reached', async () => {
    const { pages } = chatPages(300, 200);
    serve(pages);
    const r = await runListMessages({
      session: SESSION,
      httpTimeoutMs: 1000,
      chat: CHAT,
      pageSize: 200,
      until: '2020-01-01T00:00:00Z',
      maxPages: 10,
    });
    expect(r.messages).toHaveLength(300);
    expect(r.complete).toBe(true);
  });

  it('accepts unix milliseconds for --until', async () => {
    const { all, pages } = chatPages(600, 200);
    serve(pages);
    const r = await runListMessages({
      session: SESSION,
      httpTimeoutMs: 1000,
      chat: CHAT,
      pageSize: 200,
      until: String(Date.parse(all[450].composetime)),
      maxPages: 10,
    });
    expect(r.pagesRead).toBe(3);
  });

  it('never sends the token to a paging link on another host', async () => {
    const pages = [
      {
        messages: [{ id: '1', composetime: new Date(NOW).toISOString() }],
        _metadata: { backwardLink: 'https://elsewhere.example/steal?page=1' },
      },
    ];
    serve(pages);
    await expect(
      runListMessages({
        session: SESSION,
        httpTimeoutMs: 1000,
        chat: CHAT,
        until: '2020-01-01T00:00:00Z',
        maxPages: 5,
      }),
    ).rejects.toBeInstanceOf(ExitWithCode);
    expect(requested.every((u) => u.startsWith('https://teams.microsoft.com/'))).toBe(true);
  });

  it('refuses paging flags on a channel and a time that does not parse', async () => {
    serve([{ messages: [] }]);
    const invalid = { code: ExitCode.InvalidInput };
    await expect(
      runListMessages({
        session: SESSION,
        httpTimeoutMs: 1000,
        team: 't',
        channel: 'c',
        maxPages: 2,
      }),
    ).rejects.toMatchObject(invalid);
    await expect(
      runListMessages({ session: SESSION, httpTimeoutMs: 1000, chat: CHAT, until: 'yesterday' }),
    ).rejects.toMatchObject(invalid);
    await expect(
      runListMessages({ session: SESSION, httpTimeoutMs: 1000, chat: CHAT, maxPages: 0 }),
    ).rejects.toMatchObject(invalid);
    expect(requested).toHaveLength(0);
  });
});
