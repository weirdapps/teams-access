// test_scripts/commands/lookup-users.test.ts
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { parseAddressInput, runLookupUsers } from '../../src/commands/lookup-users';
import { ExitWithCode } from '../../src/util/exit-codes';
import type { Session } from '../../src/session/store';

const fakeSession: Session = {
  bearerToken: 'fake',
  cookies: [],
  capturedAt: '2026-10-18T10:00:00Z',
} as unknown as Session;

interface SubRequest {
  id: string;
  method: string;
  url: string;
  headers?: Record<string, string>;
}

interface SubResponse {
  status: number;
  headers?: Record<string, string>;
  body?: unknown;
}

const ALICE = {
  id: 'id-alice',
  displayName: 'Alice Example',
  mail: 'alice@example.com',
  userPrincipalName: 'alice@example.com',
  jobTitle: 'Analyst',
  department: 'Payments',
  officeLocation: 'Floor 3',
  companyName: 'Example Co',
  accountEnabled: true,
  manager: {
    '@odata.type': '#microsoft.graph.user',
    id: 'id-bob',
    displayName: 'Bob Example',
    mail: 'bob@example.com',
    jobTitle: 'Head of Payments',
  },
};

const NOT_FOUND: SubResponse = {
  status: 404,
  body: { error: { code: 'Request_ResourceNotFound', message: 'does not exist' } },
};

/** The address a full `/users/{addr}?$select=...` lookup asks for, or null for any other shape. */
function directAddress(url: string): string | null {
  const m = /^\/users\/([^?]+)\?\$select=id,displayName/.exec(url);
  return m ? decodeURIComponent(m[1]) : null;
}

/** The address inside a proxyAddresses filter sub-request, or null for any other shape. */
function proxyAddress(url: string): string | null {
  if (!url.startsWith('/users?')) return null;
  const filter = new URLSearchParams(url.slice('/users?'.length)).get('$filter') ?? '';
  const m = /^proxyAddresses\/any\(p:p eq 'smtp:(.*)'\)$/.exec(filter);
  return m ? m[1].replace(/''/g, "'") : null;
}

/**
 * Installs a fake Graph `$batch` endpoint. `handler` answers each sub-request;
 * `seen` is how many times that exact URL has been asked, so a test can
 * throttle the first attempt and succeed on the next. Responses come back in
 * reverse order to prove the command maps them by id, not by position.
 */
function installGraph(handler: (req: SubRequest, seen: number) => SubResponse): {
  batches: SubRequest[][];
  urls: string[];
} {
  const batches: SubRequest[][] = [];
  const urls: string[] = [];
  const seen = new Map<string, number>();
  vi.spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
    urls.push(String(input));
    const { requests } = JSON.parse(String(init?.body)) as { requests: SubRequest[] };
    batches.push(requests);
    const responses = requests.map((r) => {
      const n = (seen.get(r.url) ?? 0) + 1;
      seen.set(r.url, n);
      return { id: r.id, ...handler(r, n) };
    });
    return new Response(JSON.stringify({ responses: responses.reverse() }), { status: 200 });
  });
  return { batches, urls };
}

function userFor(addr: string): Record<string, unknown> {
  return {
    ...ALICE,
    id: `id-${addr}`,
    mail: addr,
    userPrincipalName: addr,
    displayName: addr.split('@')[0],
  };
}

const noSleep = (): ((ms: number) => Promise<void>) => vi.fn(async () => undefined);

describe('parseAddressInput', () => {
  it('trims, lower-cases, de-duplicates and skips blanks, keeping first-seen order', () => {
    expect(parseAddressInput([' B@Example.com ', '', 'a@example.com', 'b@example.com  '])).toEqual([
      'b@example.com',
      'a@example.com',
    ]);
  });

  it('splits stdin text on newlines (CRLF too) and appends it after the arguments', () => {
    const stdin = 'c@example.com\r\n\r\n  A@EXAMPLE.COM\nd@example.com\n';
    expect(parseAddressInput(['a@example.com'], stdin)).toEqual([
      'a@example.com',
      'c@example.com',
      'd@example.com',
    ]);
  });

  it('refuses an empty input with invalid_input', () => {
    try {
      parseAddressInput(['  ', ''], '\n\n');
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(ExitWithCode);
      expect((e as ExitWithCode).code).toBe(2);
      expect((e as ExitWithCode).payload.code).toBe('invalid_input');
    }
  });

  it('refuses more than 5,000 distinct addresses with invalid_input', () => {
    const many = Array.from({ length: 5001 }, (_, i) => `u${i}@example.com`);
    expect(() => parseAddressInput(many)).toThrow(/5,?000/);
    expect(parseAddressInput(many.slice(0, 5000))).toHaveLength(5000);
    // duplicates do not count against the cap
    expect(parseAddressInput([...many.slice(0, 5000), 'U0@example.com'])).toHaveLength(5000);
  });
});

describe('runLookupUsers', () => {
  beforeEach(() => vi.restoreAllMocks());

  it('returns the profile and manager of a known address', async () => {
    const { urls, batches } = installGraph((req) =>
      directAddress(req.url) === 'alice@example.com' ? { status: 200, body: ALICE } : NOT_FOUND,
    );

    const result = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['alice@example.com'],
      sleep: noSleep(),
    });

    expect(urls).toEqual(['https://graph.microsoft.com/v1.0/$batch']);
    expect(batches[0][0].method).toBe('GET');
    expect(batches[0][0].url).toBe(
      '/users/alice%40example.com?$select=id,displayName,mail,userPrincipalName,jobTitle,' +
        'department,officeLocation,companyName,accountEnabled' +
        '&$expand=manager($select=displayName,mail,jobTitle)',
    );
    expect(result).toEqual({
      users: [
        {
          query: 'alice@example.com',
          status: 'ok',
          id: 'id-alice',
          displayName: 'Alice Example',
          mail: 'alice@example.com',
          userPrincipalName: 'alice@example.com',
          jobTitle: 'Analyst',
          department: 'Payments',
          officeLocation: 'Floor 3',
          companyName: 'Example Co',
          accountEnabled: true,
          manager: {
            displayName: 'Bob Example',
            mail: 'bob@example.com',
            jobTitle: 'Head of Payments',
          },
        },
      ],
    });
  });

  it('gives null for missing fields and a null manager when Graph has none', async () => {
    installGraph(() => ({
      status: 200,
      body: { id: 'id-carol', displayName: 'Carol', userPrincipalName: 'carol@example.com' },
    }));

    const { users } = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['carol@example.com'],
      sleep: noSleep(),
    });

    expect(users[0]).toEqual({
      query: 'carol@example.com',
      status: 'ok',
      id: 'id-carol',
      displayName: 'Carol',
      mail: null,
      userPrincipalName: 'carol@example.com',
      jobTitle: null,
      department: null,
      officeLocation: null,
      companyName: null,
      accountEnabled: null,
      manager: null,
    });
  });

  it('sends 20 requests per batch and returns one entry per distinct query in input order', async () => {
    const { batches } = installGraph((req) => ({
      status: 200,
      body: userFor(directAddress(req.url) ?? ''),
    }));
    const input = Array.from({ length: 45 }, (_, i) => `user${44 - i}@example.com`);

    const { users } = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: [...input, 'USER0@example.com'],
      sleep: noSleep(),
    });

    expect(batches.map((b) => b.length)).toEqual([20, 20, 5]);
    expect(users.map((u) => u.query)).toEqual(input);
    expect(users.every((u) => u.status === 'ok')).toBe(true);
    expect(users.map((u) => (u as { mail: string }).mail)).toEqual(input);
  });

  it('looks a 404 up again by proxyAddresses and fetches its manager separately', async () => {
    const { batches } = installGraph((req) => {
      if (directAddress(req.url) !== null) return NOT_FOUND;
      if (proxyAddress(req.url) === "alias.o'neil@example.com") {
        const { manager: _manager, ...rest } = ALICE;
        return { status: 200, body: { '@odata.count': 1, value: [rest] } };
      }
      if (proxyAddress(req.url) !== null) return { status: 200, body: { value: [] } };
      if (req.url.startsWith('/users/id-alice?')) {
        return { status: 200, body: { id: 'id-alice', manager: ALICE.manager } };
      }
      return NOT_FOUND;
    });

    const { users } = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ["alias.o'neil@example.com", 'nobody@example.com'],
      sleep: noSleep(),
    });

    expect(batches).toHaveLength(3);
    const proxyReq = batches[1].find((r) => proxyAddress(r.url) === "alias.o'neil@example.com");
    expect(proxyReq?.headers).toEqual({ ConsistencyLevel: 'eventual' });
    expect(proxyReq?.url).toContain('$count=true');
    expect(proxyReq?.url).toContain("o''neil");
    expect(batches[2].map((r) => r.url)).toEqual([
      '/users/id-alice?$select=id&$expand=manager($select=displayName,mail,jobTitle)',
    ]);
    expect(users).toEqual([
      expect.objectContaining({
        query: "alias.o'neil@example.com",
        status: 'ok',
        id: 'id-alice',
        mail: 'alice@example.com',
        manager: {
          displayName: 'Bob Example',
          mail: 'bob@example.com',
          jobTitle: 'Head of Payments',
        },
      }),
      { query: 'nobody@example.com', status: 'not_found' },
    ]);
  });

  it('gives a null manager to an alias match whose follow-up shows none', async () => {
    installGraph((req) => {
      if (directAddress(req.url) !== null) return NOT_FOUND;
      if (proxyAddress(req.url) !== null) {
        return { status: 200, body: { value: [{ id: 'id-dan', displayName: 'Dan' }] } };
      }
      return { status: 200, body: { id: 'id-dan' } };
    });

    const { users } = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['dan.alias@example.com'],
      sleep: noSleep(),
    });

    expect(users[0]).toEqual(
      expect.objectContaining({ status: 'ok', id: 'id-dan', mail: null, manager: null }),
    );
  });

  it('retries only throttled items after their Retry-After, capped at 30 s', async () => {
    const { batches } = installGraph((req, seen) => {
      const addr = directAddress(req.url) ?? '';
      if (addr === 'slow@example.com' && seen === 1) {
        return { status: 429, headers: { 'Retry-After': '7' }, body: {} };
      }
      if (addr === 'busy@example.com' && seen === 1) {
        return { status: 503, headers: { 'retry-after': '120' }, body: {} };
      }
      return { status: 200, body: userFor(addr) };
    });
    const sleep = noSleep();

    const { users } = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['ok@example.com', 'slow@example.com', 'busy@example.com'],
      sleep,
    });

    expect(batches.map((b) => b.length)).toEqual([3, 2]);
    expect(sleep).toHaveBeenCalledTimes(1);
    expect(sleep).toHaveBeenCalledWith(30_000);
    expect(users.map((u) => u.status)).toEqual(['ok', 'ok', 'ok']);
  });

  it('marks an item still throttled after 3 retries as error and keeps the rest', async () => {
    const { batches } = installGraph((req) => {
      const addr = directAddress(req.url) ?? '';
      if (addr === 'stuck@example.com') {
        return { status: 429, headers: { 'Retry-After': '1' }, body: {} };
      }
      return { status: 200, body: userFor(addr) };
    });
    const sleep = noSleep();

    const { users } = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['stuck@example.com', 'fine@example.com'],
      sleep,
    });

    expect(batches.map((b) => b.length)).toEqual([2, 1, 1, 1]);
    expect(sleep).toHaveBeenCalledTimes(3);
    expect(sleep).toHaveBeenCalledWith(1000);
    expect(users[0]).toEqual({
      query: 'stuck@example.com',
      status: 'error',
      error: expect.stringMatching(/429/),
    });
    expect(users[1].status).toBe('ok');
  });

  it('marks any other per-item failure as error with the Graph message', async () => {
    installGraph(() => ({
      status: 403,
      body: { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } },
    }));

    const { users } = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['eve@example.com'],
      sleep: noSleep(),
    });

    expect(users).toEqual([
      {
        query: 'eve@example.com',
        status: 'error',
        error: expect.stringMatching(/403.*Authorization_RequestDenied.*Insufficient privileges/),
      },
    ]);
  });

  it('marks an alias match as error when its manager follow-up fails', async () => {
    installGraph((req) => {
      if (directAddress(req.url) !== null) return NOT_FOUND;
      if (proxyAddress(req.url) !== null) {
        return { status: 200, body: { value: [{ id: 'id-fay', displayName: 'Fay' }] } };
      }
      return { status: 500, body: { error: { code: 'InternalServerError', message: 'boom' } } };
    });

    const { users } = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['fay.alias@example.com'],
      sleep: noSleep(),
    });

    expect(users).toEqual([
      { query: 'fay.alias@example.com', status: 'error', error: expect.stringMatching(/manager/) },
    ]);
  });

  it('exits 4 auth_required when Graph refuses the whole call with 401', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      new Response(
        JSON.stringify({ error: { code: 'InvalidAuthenticationToken', message: 'x' } }),
        {
          status: 401,
        },
      ),
    );

    const err = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['alice@example.com'],
      sleep: noSleep(),
    }).catch((e: unknown) => e);

    expect(err).toBeInstanceOf(ExitWithCode);
    expect((err as ExitWithCode).code).toBe(4);
    expect((err as ExitWithCode).payload.code).toBe('auth_required');
  });

  it('exits 5 upstream when Graph refuses the whole call otherwise', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      new Response(JSON.stringify({ error: { code: 'BadRequest', message: 'bad batch' } }), {
        status: 400,
      }),
    );

    const err = await runLookupUsers({
      session: fakeSession,
      httpTimeoutMs: 1000,
      addresses: ['alice@example.com'],
      sleep: noSleep(),
    }).catch((e: unknown) => e);

    expect(err).toBeInstanceOf(ExitWithCode);
    expect((err as ExitWithCode).code).toBe(5);
    expect((err as ExitWithCode).payload).toEqual(
      expect.objectContaining({ code: 'upstream', status: 400 }),
    );
  });

  it('refuses bad input before calling Graph', async () => {
    const fetchSpy = vi.spyOn(globalThis, 'fetch');

    await expect(
      runLookupUsers({ session: fakeSession, httpTimeoutMs: 1000, addresses: [' '] }),
    ).rejects.toThrow(/no addresses/i);
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});
