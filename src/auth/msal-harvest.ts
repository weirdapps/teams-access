// Read access tokens out of the page's MSAL cache instead of intercepting them
// on the wire.
//
// WHY THIS EXISTS
// Network capture is how this CLI has always worked, and on 2026-09-03 it stopped
// producing anything on an MCAS (Defender for Cloud Apps) tenant. A measured run
// against a fully-rendered Teams page saw 583 requests and ZERO with an
// Authorization header: 412 went to teams.public.onecdn.static.microsoft and 104
// to the teams.cloud.microsoft shell, while not one API host (chatsvcagg, ic3,
// presence, graph) appeared at all. The API calls are dispatched from a Service
// Worker under an MCAS-proxied origin and never surfaced to a listener we can
// attach to; Playwright also refuses a CDP session on a service-worker target
// ("expected Page or Frame").
//
// Meanwhile every token we want is sitting in localStorage, because that is where
// MSAL puts it and where the SPA itself reads it from. The same measured run found
// AccessToken entries for graph.microsoft.com, chatsvcagg.teams.microsoft.com,
// ic3.teams.office.com, presence.teams.microsoft.com, api.spaces.skype.com and
// outlook.office.com/search, plus the RefreshToken and IdToken.
//
// Reading the cache is strictly more reliable than provoking traffic: it does not
// depend on which surfaces the SPA happens to call, on clicking the right app-rail
// item, on MCAS not rewriting hosts, or on a Service Worker being interceptable.
// It is also what a token expiry check should read, since a captured Bearer is
// only ever a copy of this.
//
// MSAL cache layout (schema "msal.2"), one localStorage entry per credential:
//   key   msal.2|<homeAccountId>|<environment>|accesstoken|<clientId>|<realm>|<target>
//   value {"credentialType":"AccessToken","secret":"<JWT>","target":"<scopes>",
//          "expiresOn":<epoch-seconds>,"clientId":...,"realm":...}
// We key results by the JWT `aud` claim, matching what the network path produced,
// so downstream code and existing session.json files stay compatible.

export interface HarvestedToken {
  /** Raw JWT. */
  token: string;
  /** JWT `aud` claim: the key the rest of the CLI indexes tokens by. */
  aud: string;
  exp?: number;
  scp?: string;
  appid?: string;
  /** MSAL's own expiry, used only when the JWT has no `exp`. */
  msalExpiresOn?: number;
}

interface RawEntry {
  secret: string;
  target?: string;
  expiresOn?: string | number;
}

/** Minimal surface we need from a Playwright Page, kept narrow for testability. */
export interface EvaluatablePage {
  evaluate<R>(fn: () => R): Promise<R>;
}

/** A token evicted from the cache, described without its secret. */
export interface EvictedToken {
  /** JWT `aud`, or MSAL's `target` scopes when the JWT will not decode. Never the token. */
  aud: string;
  /** Seconds of life it had left when it was removed; negative if already dead. */
  ttlSeconds: number;
}

/** Just enough of DOM Storage for {@link evictNearExpiryInPage}; a fake in tests. */
export interface StorageLike {
  readonly length: number;
  key(index: number): string | null;
  getItem(key: string): string | null;
  removeItem(key: string): void;
}

/**
 * Remove every MSAL AccessToken that expires within `minTtlSeconds`, so the SPA
 * has to mint a fresh one from its refresh token instead of handing back a copy.
 *
 * WHY. MSAL only refreshes a cached token inside its own renewal offset, about
 * five minutes, so a harvest copies whatever the cache holds. The producer
 * renews every 15 minutes, and each push carried the same bearer with ~1000s
 * less life until one landed with 250-640s left and died on the VPS before the
 * next push arrived: 4 of 4 teams-session outages there on 2026-09-22/23
 * followed such a push by 12-16 minutes. Taking the token OUT of the cache
 * before the page boots is the only lever that makes the SPA ask Entra again;
 * reading the cache harder cannot.
 *
 * WHAT IT TOUCHES. Access tokens only, and only those whose expiry it can read
 * (JWT `exp`, else MSAL's `expiresOn`): the refresh token the SPA mints from,
 * the id token, long-lived access tokens and anything that is not an MSAL
 * credential all stay. A token with no readable expiry is kept, because
 * evicting on a guess could strip a good token out of a working session.
 *
 * RUNS INSIDE THE PAGE, installed with `context.addInitScript` so it fires on
 * every document before the SPA's own scripts. Playwright ships it as source
 * text, so it must stay SELF-CONTAINED: no imports and nothing from this module,
 * or it throws in the page and silently evicts nothing. `storage` and
 * `nowSeconds` exist for tests; the page passes only `minTtlSeconds`.
 *
 * Returns the evicted tokens and leaves the same list on `globalThis`, where
 * {@link readEvictionReport} picks it up: an init script's return value goes
 * nowhere.
 */
export function evictNearExpiryInPage(arg: {
  minTtlSeconds: number;
  nowSeconds?: number;
  storage?: StorageLike;
}): EvictedToken[] {
  const out: EvictedToken[] = [];
  try {
    const store: StorageLike = arg.storage ?? localStorage;
    const now = arg.nowSeconds ?? Math.floor(Date.now() / 1000);
    const keys: string[] = [];
    for (let i = 0; i < store.length; i++) {
      const k = store.key(i);
      if (k && /accesstoken/i.test(k)) keys.push(k);
    }
    for (const k of keys) {
      let entry: Record<string, unknown>;
      try {
        entry = JSON.parse(store.getItem(k) ?? '') as Record<string, unknown>;
      } catch {
        continue; // not JSON: not an MSAL credential
      }
      if (!entry || entry.credentialType !== 'AccessToken') continue;
      let claims: Record<string, unknown> | null = null;
      const parts = typeof entry.secret === 'string' ? entry.secret.split('.') : [];
      if (parts.length === 3) {
        try {
          const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
          claims = JSON.parse(atob(b64 + '='.repeat((4 - (b64.length % 4)) % 4))) as Record<
            string,
            unknown
          >;
        } catch {
          claims = null;
        }
      }
      const jwtExp = claims && typeof claims.exp === 'number' ? claims.exp : undefined;
      const msalExp = entry.expiresOn !== undefined ? Number(entry.expiresOn) : NaN;
      const exp = jwtExp ?? (Number.isFinite(msalExp) && msalExp > 0 ? msalExp : undefined);
      if (exp === undefined || exp - now >= arg.minTtlSeconds) continue;
      store.removeItem(k);
      const aud =
        claims && typeof claims.aud === 'string'
          ? claims.aud
          : typeof entry.target === 'string'
            ? entry.target
            : '(unknown audience)';
      out.push({ aud, ttlSeconds: exp - now });
    }
  } catch {
    /* storage blocked: nothing evicted, and the capture proceeds as before */
  }
  (globalThis as unknown as Record<string, unknown>).__teamsCliEvicted = out;
  return out;
}

/**
 * Seconds left on a JWT when that is under `floorS`, else undefined: also when
 * there is no floor, or the token carries no readable `exp`, because refusing
 * a token on a guess could throw away the only copy of a working credential.
 */
export function secondsLeftUnderFloor(
  jwt: string,
  floorS: number,
  nowS: number = Math.floor(Date.now() / 1000),
): number | undefined {
  if (!(floorS > 0)) return undefined;
  const claims = decodeClaims(jwt);
  const exp = claims && typeof claims.exp === 'number' ? claims.exp : undefined;
  if (exp === undefined) return undefined;
  const left = exp - nowS;
  return left < floorS ? left : undefined;
}

/**
 * A predicate for the network listener: true when a bearer seen on the wire is
 * under the renew floor and must not be captured or become the session's
 * primary. Says so once per audience, since the SPA re-sends the same token on
 * every call. With no floor (interactive login) it never refuses.
 */
export function wireFloorGuard(
  floorS: number,
  say: (line: string) => void = (line) => process.stderr.write(line),
): (token: string, aud: string) => boolean {
  const told = new Set<string>();
  return (token, aud) => {
    const left = secondsLeftUnderFloor(token, floorS);
    if (left === undefined) return false;
    if (!told.has(aud)) {
      told.add(aud);
      say(`[wire] not capturing aud=${aud}: ${left}s left, under the ${floorS}s renew floor\n`);
    }
    return true;
  };
}

/** The slice of a Playwright BrowserContext the renew eviction needs. */
export interface InitScriptTarget {
  addInitScript(
    script: (arg: { minTtlSeconds: number }) => unknown,
    arg: { minTtlSeconds: number },
  ): Promise<unknown>;
}

/**
 * Install {@link evictNearExpiryInPage} to run before every document on the
 * context. An init script rather than evict-then-reload: it costs no extra page
 * load, and the renew budget (first bearer + 40s) already cuts the navigation
 * list short. Returns false, installing nothing, when there is no threshold.
 */
export async function installRenewEviction(
  context: InitScriptTarget,
  minTtlSeconds: number,
  floorS: number,
  say: (line: string) => void = (line) => process.stderr.write(line),
): Promise<boolean> {
  if (!(minTtlSeconds > 0)) return false;
  await context.addInitScript(evictNearExpiryInPage, { minTtlSeconds });
  say(
    `[msal-cache] renew: evicting cached access tokens with under ${minTtlSeconds}s left ` +
      `before each page boots, and capturing nothing under ${floorS}s\n`,
  );
  return true;
}

/** Say what the init script took out of the current document's cache, if anything. */
export async function reportEvictions(
  page: EvaluatablePage,
  where: string,
  say: (line: string) => void = (line) => process.stderr.write(line),
): Promise<void> {
  const line = describeEvictions(where, await readEvictionReport(page));
  if (line) say(line);
}

/** The stderr line for one document's evictions, or '' for none. Never a token. */
export function describeEvictions(where: string, evicted: EvictedToken[]): string {
  if (!evicted.length) return '';
  return (
    `[msal-cache] evicted ${evicted.length} near-expiry token(s) on ${where} so the SPA ` +
    `mints fresh ones: ${evicted.map((e) => `${e.aud} (${e.ttlSeconds}s left)`).join(', ')}\n`
  );
}

/**
 * What {@link evictNearExpiryInPage} removed in the page's current document,
 * then cleared so the same eviction is never reported twice. [] on any error.
 */
export async function readEvictionReport(page: EvaluatablePage): Promise<EvictedToken[]> {
  try {
    return await page.evaluate<EvictedToken[]>(() => {
      const g = globalThis as unknown as Record<string, unknown>;
      const report = Array.isArray(g.__teamsCliEvicted)
        ? (g.__teamsCliEvicted as EvictedToken[])
        : [];
      g.__teamsCliEvicted = [];
      return report;
    });
  } catch {
    return [];
  }
}

function decodeClaims(jwt: string): Record<string, unknown> | null {
  const parts = jwt.split('.');
  if (parts.length !== 3) return null;
  try {
    const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    return JSON.parse(Buffer.from(b64, 'base64').toString('utf8')) as Record<string, unknown>;
  } catch {
    return null;
  }
}

/**
 * Pull every non-expired AccessToken from the page's MSAL localStorage cache.
 *
 * Returns [] rather than throwing when the page has no cache, is on the wrong
 * origin, or storage is unreadable: the caller treats this as one source among
 * several, never as the sole authority.
 *
 * `skewSeconds` drops tokens about to expire, so we never persist a credential
 * that dies before the next command runs.
 */
export async function harvestMsalTokens(
  page: EvaluatablePage,
  skewSeconds = 120,
): Promise<HarvestedToken[]> {
  let raw: RawEntry[];
  try {
    raw = await page.evaluate<RawEntry[]>(() => {
      const out: RawEntry[] = [];
      try {
        for (let i = 0; i < localStorage.length; i++) {
          const k = localStorage.key(i);
          if (!k || !/accesstoken/i.test(k)) continue;
          const v = localStorage.getItem(k);
          if (!v) continue;
          try {
            const o = JSON.parse(v) as Record<string, unknown>;
            if (o?.credentialType !== 'AccessToken') continue;
            if (typeof o.secret !== 'string' || !o.secret) continue;
            out.push({
              secret: o.secret,
              target: typeof o.target === 'string' ? o.target : undefined,
              expiresOn: (o.expiresOn as string | number | undefined) ?? undefined,
            });
          } catch {
            /* not JSON, or not an MSAL credential: skip */
          }
        }
      } catch {
        /* storage blocked (rare, e.g. third-party-cookie policies): return what we have */
      }
      return out;
    });
  } catch {
    return [];
  }

  const now = Math.floor(Date.now() / 1000);
  const byAud = new Map<string, HarvestedToken>();

  for (const entry of raw) {
    const claims = decodeClaims(entry.secret);
    if (!claims) continue;

    const aud = typeof claims.aud === 'string' ? claims.aud : undefined;
    if (!aud) continue;

    const exp = typeof claims.exp === 'number' ? claims.exp : undefined;
    const msalExpiresOn =
      entry.expiresOn !== undefined ? Number(entry.expiresOn) || undefined : undefined;

    // Drop anything already dead or dying. The MSAL cache legitimately retains
    // expired entries until eviction, so without this we would happily persist a
    // token that has been useless for hours, which is exactly the failure this
    // module was written to end.
    const effectiveExp = exp ?? msalExpiresOn;
    if (effectiveExp !== undefined && effectiveExp <= now + skewSeconds) continue;

    // Keep the longest-lived token per audience: MSAL can hold several entries
    // for one audience with different scope sets.
    const prev = byAud.get(aud);
    if (prev && (prev.exp ?? 0) >= (exp ?? 0)) continue;

    byAud.set(aud, {
      token: entry.secret,
      aud,
      exp,
      scp: typeof claims.scp === 'string' ? claims.scp : entry.target,
      appid: typeof claims.appid === 'string' ? claims.appid : undefined,
      msalExpiresOn,
    });
  }

  return Array.from(byAud.values());
}
