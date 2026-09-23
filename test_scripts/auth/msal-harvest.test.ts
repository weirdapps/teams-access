// test_scripts/auth/msal-harvest.test.ts
//
// auth-renew copied whatever access token MSAL happened to have cached. The SPA
// only refreshes a token when it is inside MSAL's own renewal offset (about five
// minutes), so every 15-minute push from the producer carried the SAME bearer
// with ~1000s less life, and one run in four or five landed with 250-640s left.
// That token then died on the VPS before the next push arrived: measured
// 2026-09-22/23, 4 of 4 teams-session failures on the VPS followed such a push
// by 12-16 minutes, and each one cost sb-teams-sync a failed hour.
//
// The fix is to take near-expiry tokens OUT of the cache before the page boots,
// so the SPA has to mint fresh ones from its refresh token, and to refuse to
// capture anything that is still near expiry. These tests pin both halves
// without a browser: the eviction runs against a fake Storage, and the harvest
// against a fake page.
import { describe, it, expect, afterEach } from 'vitest';

import {
  evictNearExpiryInPage,
  harvestMsalTokens,
  readEvictionReport,
  type EvaluatablePage,
} from '../../src/auth/msal-harvest';

const NOW = 1_790_000_000;

/** A JWT whose payload carries `aud` and `exp`. The signature is never checked. */
function jwt(aud: string, exp: number): string {
  const b64 = (o: object) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'none', typ: 'JWT' })}.${b64({ aud, exp })}.sig`;
}

function accessToken(aud: string, exp: number): string {
  return JSON.stringify({
    credentialType: 'AccessToken',
    secret: jwt(aud, exp),
    target: `${aud}/.default`,
    expiresOn: String(exp),
  });
}

/** Just enough of the DOM Storage interface for the in-page function. */
class FakeStorage {
  private readonly map = new Map<string, string>();
  constructor(entries: Record<string, string>) {
    for (const [k, v] of Object.entries(entries)) this.map.set(k, v);
  }
  get length(): number {
    return this.map.size;
  }
  key(i: number): string | null {
    return Array.from(this.map.keys())[i] ?? null;
  }
  getItem(k: string): string | null {
    return this.map.get(k) ?? null;
  }
  removeItem(k: string): void {
    this.map.delete(k);
  }
  keys(): string[] {
    return Array.from(this.map.keys());
  }
}

const AT = (aud: string) => `msal.2|home.tid|login.windows.net|accesstoken|client|tid|${aud}`;

afterEach(() => {
  delete (globalThis as Record<string, unknown>).__teamsCliEvicted;
});

describe('evictNearExpiryInPage', () => {
  it('removes only the access tokens that expire inside the window', () => {
    const graphSoon = jwt('https://graph.microsoft.com', NOW + 400);
    const store = new FakeStorage({
      [AT('graph')]: accessToken('https://graph.microsoft.com', NOW + 400),
      [AT('chatsvcagg')]: accessToken('https://chatsvcagg.teams.microsoft.com', NOW + 4000),
      'msal.2|home.tid|login.windows.net|refreshtoken|client||': JSON.stringify({
        credentialType: 'RefreshToken',
        secret: 'refresh-secret',
      }),
      'msal.2|home.tid|login.windows.net|idtoken|client|tid|': JSON.stringify({
        credentialType: 'IdToken',
        secret: jwt('client', NOW + 10),
      }),
      'some.app.setting': '{"theme":"dark"}',
      'msal.2|broken|accesstoken|x': 'not json at all',
    });

    const evicted = evictNearExpiryInPage({ minTtlSeconds: 1500, nowSeconds: NOW, storage: store });

    expect(evicted).toEqual([{ aud: 'https://graph.microsoft.com', ttlSeconds: 400 }]);
    expect(store.getItem(AT('graph'))).toBeNull();
    // The long-lived access token, the refresh token the SPA needs to mint a new
    // one, the id token and anything that is not an MSAL credential all stay.
    expect(store.keys()).toHaveLength(5);
    expect(store.getItem(AT('chatsvcagg'))).not.toBeNull();
    expect(store.getItem('msal.2|home.tid|login.windows.net|refreshtoken|client||')).not.toBeNull();
    // The report names the audience and the life it had left, never the token.
    expect(JSON.stringify(evicted)).not.toContain(graphSoon);
  });

  it('removes an access token that is already dead', () => {
    const store = new FakeStorage({
      [AT('ic3')]: accessToken('https://ic3.teams.office.com', NOW - 60),
    });
    const evicted = evictNearExpiryInPage({ minTtlSeconds: 1500, nowSeconds: NOW, storage: store });
    expect(evicted).toEqual([{ aud: 'https://ic3.teams.office.com', ttlSeconds: -60 }]);
    expect(store.keys()).toEqual([]);
  });

  it('keeps a token whose expiry it cannot prove', () => {
    // No JWT exp and no MSAL expiresOn: evicting on a guess could strip a
    // perfectly good token out of a working session.
    const store = new FakeStorage({
      [AT('graph')]: JSON.stringify({ credentialType: 'AccessToken', secret: 'opaque' }),
    });
    expect(evictNearExpiryInPage({ minTtlSeconds: 1500, nowSeconds: NOW, storage: store })).toEqual(
      [],
    );
    expect(store.keys()).toHaveLength(1);
  });

  it('leaves its report where the capture loop reads it, once', async () => {
    const store = new FakeStorage({
      [AT('graph')]: accessToken('https://graph.microsoft.com', NOW + 100),
    });
    evictNearExpiryInPage({ minTtlSeconds: 1500, nowSeconds: NOW, storage: store });

    // The same object the init script wrote to, read back through the same
    // evaluate path the capture loop uses. Proves the two sides agree on where
    // the report lives, since neither can import the other's constant.
    const page: EvaluatablePage = {
      evaluate: async (fn: () => unknown) => fn(),
    } as EvaluatablePage;
    expect(await readEvictionReport(page)).toEqual([
      { aud: 'https://graph.microsoft.com', ttlSeconds: 100 },
    ]);
    expect(await readEvictionReport(page)).toEqual([]);
  });

  it('is self-contained, because Playwright ships it to the page as source text', () => {
    // addInitScript serialises the function with toString() and runs it in a
    // page that has none of this module's imports. A reference to anything
    // outside the function body would throw there and silently evict nothing.
    const src = evictNearExpiryInPage.toString();
    expect(src).not.toMatch(/require\(|import\(|__awaiter|_msal_harvest|exports\./);
  });
});

describe('harvestMsalTokens skew', () => {
  function pageWith(entries: Array<{ aud: string; exp: number }>): EvaluatablePage {
    return {
      evaluate: async () =>
        entries.map((e) => ({ secret: jwt(e.aud, e.exp), target: e.aud, expiresOn: e.exp })),
    } as unknown as EvaluatablePage;
  }

  it('drops a token under the floor it is given, and keeps it under the default', async () => {
    const realNow = Date.now;
    Date.now = () => NOW * 1000;
    try {
      const page = pageWith([{ aud: 'https://graph.microsoft.com', exp: NOW + 800 }]);
      expect(await harvestMsalTokens(page, 900)).toEqual([]);
      expect((await harvestMsalTokens(page)).map((t) => t.aud)).toEqual([
        'https://graph.microsoft.com',
      ]);
    } finally {
      Date.now = realNow;
    }
  });
});
