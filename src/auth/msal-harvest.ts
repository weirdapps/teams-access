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
