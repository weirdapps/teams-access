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
import { describe, it, expect, afterEach, vi } from 'vitest';

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

describe('the renew floor on the wire', () => {
  it('reports seconds left only under the floor, and never guesses', async () => {
    const { secondsLeftUnderFloor } = await import('../../src/auth/msal-harvest');
    expect(secondsLeftUnderFloor(jwt('https://graph.microsoft.com', NOW + 300), 900, NOW)).toBe(
      300,
    );
    expect(secondsLeftUnderFloor(jwt('https://graph.microsoft.com', NOW + 3000), 900, NOW)).toBe(
      undefined,
    );
    // No floor is an interactive login: nothing is ever refused.
    expect(secondsLeftUnderFloor(jwt('https://graph.microsoft.com', NOW + 1), 0, NOW)).toBe(
      undefined,
    );
    // Unreadable or exp-less tokens are not refused on a guess.
    expect(secondsLeftUnderFloor('not-a-jwt', 900, NOW)).toBe(undefined);
    const noExp = `${Buffer.from('{}').toString('base64url')}.${Buffer.from('{"aud":"x"}').toString('base64url')}.s`;
    expect(secondsLeftUnderFloor(noExp, 900, NOW)).toBe(undefined);
  });

  it('refuses a near-expiry wire token and says so once per audience', async () => {
    const { wireFloorGuard } = await import('../../src/auth/msal-harvest');
    const lines: string[] = [];
    const refuse = wireFloorGuard(900, (l) => lines.push(l));
    const soon = jwt('https://graph.microsoft.com', Math.floor(Date.now() / 1000) + 300);
    const later = jwt('https://graph.microsoft.com', Math.floor(Date.now() / 1000) + 4000);
    expect(refuse(soon, 'https://graph.microsoft.com')).toBe(true);
    expect(refuse(soon, 'https://graph.microsoft.com')).toBe(true);
    expect(refuse(later, 'https://graph.microsoft.com')).toBe(false);
    expect(lines).toHaveLength(1);
    expect(lines[0]).toContain('aud=https://graph.microsoft.com');
    expect(lines[0]).toContain('900s renew floor');
    expect(lines[0]).not.toContain(soon);
  });

  it('never refuses anything without a floor', async () => {
    const { wireFloorGuard } = await import('../../src/auth/msal-harvest');
    const refuse = wireFloorGuard(0, () => {
      throw new Error('must not log');
    });
    expect(refuse(jwt('https://graph.microsoft.com', 1), 'https://graph.microsoft.com')).toBe(
      false,
    );
  });

  it('describes an eviction by audience and seconds left, and says nothing for none', async () => {
    const { describeEvictions } = await import('../../src/auth/msal-harvest');
    expect(describeEvictions('https://teams.cloud.microsoft/', [])).toBe('');
    const line = describeEvictions('https://teams.cloud.microsoft/', [
      { aud: 'https://graph.microsoft.com', ttlSeconds: 400 },
      { aud: 'https://ic3.teams.office.com', ttlSeconds: -60 },
    ]);
    expect(line).toContain('evicted 2 near-expiry token(s) on https://teams.cloud.microsoft/');
    expect(line).toContain('https://graph.microsoft.com (400s left)');
    expect(line).toContain('https://ic3.teams.office.com (-60s left)');
  });
});

describe('renew eviction wiring', () => {
  it('installs the in-page eviction with its threshold and says so', async () => {
    const { installRenewEviction, evictNearExpiryInPage: fn } =
      await import('../../src/auth/msal-harvest');
    const calls: Array<{ script: unknown; arg: unknown }> = [];
    const lines: string[] = [];
    const context = {
      addInitScript: async (script: unknown, arg: unknown) => {
        calls.push({ script, arg });
      },
    };
    expect(await installRenewEviction(context, 1500, 900, (l) => lines.push(l))).toBe(true);
    expect(calls).toEqual([{ script: fn, arg: { minTtlSeconds: 1500 } }]);
    expect(lines[0]).toContain('under 1500s');
    expect(lines[0]).toContain('nothing under 900s');
  });

  it('installs nothing without a threshold, so a login is untouched', async () => {
    const { installRenewEviction } = await import('../../src/auth/msal-harvest');
    const context = {
      addInitScript: async () => {
        throw new Error('must not install');
      },
    };
    expect(await installRenewEviction(context, 0, 0, () => undefined)).toBe(false);
  });

  it('reports a document eviction once, and stays quiet for none', async () => {
    const { reportEvictions } = await import('../../src/auth/msal-harvest');
    (globalThis as Record<string, unknown>).__teamsCliEvicted = [
      { aud: 'https://graph.microsoft.com', ttlSeconds: 400 },
    ];
    const page = { evaluate: async (fn: () => unknown) => fn() } as EvaluatablePage;
    const lines: string[] = [];
    await reportEvictions(page, 'https://teams.cloud.microsoft/', (l) => lines.push(l));
    await reportEvictions(page, 'https://teams.cloud.microsoft/', (l) => lines.push(l));
    expect(lines).toHaveLength(1);
    expect(lines[0]).toContain('https://graph.microsoft.com (400s left)');
  });

  it('reads an unreadable report as nothing evicted', async () => {
    const { readEvictionReport } = await import('../../src/auth/msal-harvest');
    const page = {
      evaluate: async () => {
        throw new Error('page navigating');
      },
    } as unknown as EvaluatablePage;
    expect(await readEvictionReport(page)).toEqual([]);
  });
});

describe('evictNearExpiryInPage fallbacks', () => {
  it('reads MSAL expiresOn and names the scopes when the secret is not a JWT', () => {
    const store = new FakeStorage({
      [AT('opaque')]: JSON.stringify({
        credentialType: 'AccessToken',
        secret: 'opaque.not-base64!.x',
        target: 'https://presence.teams.microsoft.com/.default',
        expiresOn: String(NOW + 200),
      }),
    });
    expect(evictNearExpiryInPage({ minTtlSeconds: 1500, nowSeconds: NOW, storage: store })).toEqual(
      [{ aud: 'https://presence.teams.microsoft.com/.default', ttlSeconds: 200 }],
    );
  });

  it('says unknown audience rather than inventing one', () => {
    const store = new FakeStorage({
      [AT('bare')]: JSON.stringify({
        credentialType: 'AccessToken',
        secret: 'opaque',
        expiresOn: NOW + 10,
      }),
    });
    expect(
      evictNearExpiryInPage({ minTtlSeconds: 1500, nowSeconds: NOW, storage: store })[0].aud,
    ).toBe('(unknown audience)');
  });

  it('evicts nothing, and does not throw, when storage is blocked', () => {
    const blocked = {
      get length(): number {
        throw new Error('SecurityError');
      },
      key: () => null,
      getItem: () => null,
      removeItem: () => undefined,
    };
    expect(
      evictNearExpiryInPage({ minTtlSeconds: 1500, nowSeconds: NOW, storage: blocked }),
    ).toEqual([]);
  });
});

describe('what reaches stderr by default', () => {
  it('writes the wire refusal, the install line and the eviction report to stderr', async () => {
    const { wireFloorGuard, installRenewEviction, reportEvictions } =
      await import('../../src/auth/msal-harvest');
    const written: string[] = [];
    const spy = vi.spyOn(process.stderr, 'write').mockImplementation((chunk: unknown) => {
      written.push(String(chunk));
      return true;
    });
    try {
      const soon = jwt('https://graph.microsoft.com', Math.floor(Date.now() / 1000) + 60);
      wireFloorGuard(900)(soon, 'https://graph.microsoft.com');
      await installRenewEviction({ addInitScript: async () => undefined }, 1500, 900);
      (globalThis as Record<string, unknown>).__teamsCliEvicted = [
        { aud: 'https://ic3.teams.office.com', ttlSeconds: 30 },
      ];
      await reportEvictions(
        { evaluate: async (fn: () => unknown) => fn() } as EvaluatablePage,
        'https://teams.cloud.microsoft/',
      );
      expect(written).toHaveLength(3);
      expect(written.join('')).not.toContain(soon);
    } finally {
      spy.mockRestore();
    }
  });
});

describe('evictNearExpiryInPage skips what is not an access token', () => {
  it('leaves a non-AccessToken credential stored under an access-token key alone', () => {
    const store = new FakeStorage({
      [AT('odd')]: JSON.stringify({ credentialType: 'RefreshToken', secret: 'r', expiresOn: 1 }),
      [AT('null')]: 'null',
    });
    expect(evictNearExpiryInPage({ minTtlSeconds: 1500, nowSeconds: NOW, storage: store })).toEqual(
      [],
    );
    expect(store.keys()).toHaveLength(2);
  });
});
