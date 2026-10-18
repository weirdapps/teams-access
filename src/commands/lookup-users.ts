// src/commands/lookup-users.ts: look up users by address via Graph $batch, profile fields plus manager.

import { GraphClient } from '../http/graph-client';
import { AuthRequiredError, GraphHttpError } from '../http/errors';
import { ExitCode, ExitWithCode } from '../util/exit-codes';
import type { Session } from '../session/store';

export const MAX_ADDRESSES = 5000;
const BATCH_SIZE = 20; // Graph's per-$batch request limit
const MAX_RETRIES = 3;
const RETRY_AFTER_CAP_S = 30;

const USER_SELECT =
  'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,companyName,accountEnabled';
const MANAGER_EXPAND = 'manager($select=displayName,mail,jobTitle)';

// Addresses come from mail headers: anything outside this set, or a leading
// "-", is reported per item and never sent to Graph.
const ADDRESS_RE = /^[A-Za-z0-9._%+'@-]+$/;
const INVALID_ADDRESS =
  "invalid address: only letters, digits and . _ % + ' @ - are allowed, and no leading -";

export interface LookupUsersOptions {
  session: Session;
  httpTimeoutMs: number;
  addresses: string[];
  sleep?: (ms: number) => Promise<void>; // injectable so tests skip Retry-After waits
}

export interface UserManager {
  displayName: string | null;
  mail: string | null;
  jobTitle: string | null;
}

export interface FoundUser {
  query: string;
  status: 'ok';
  id: string;
  displayName: string | null;
  mail: string | null;
  userPrincipalName: string | null;
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
  companyName: string | null;
  accountEnabled: boolean | null;
  manager: UserManager | null;
}

export type LookupUserEntry =
  | FoundUser
  | { query: string; status: 'not_found' }
  | { query: string; status: 'error'; error: string };

export interface LookupUsersResult {
  users: LookupUserEntry[];
}

interface GraphUser {
  id: string;
  displayName?: string | null;
  mail?: string | null;
  userPrincipalName?: string | null;
  jobTitle?: string | null;
  department?: string | null;
  officeLocation?: string | null;
  companyName?: string | null;
  accountEnabled?: boolean | null;
  manager?: { displayName?: string | null; mail?: string | null; jobTitle?: string | null } | null;
}

interface BatchRequest {
  id: string;
  method: 'GET';
  url: string;
  headers?: Record<string, string>;
}

interface BatchResponse {
  id: string;
  status: number;
  headers?: Record<string, string>;
  body?: unknown;
}

/** Arguments plus optional stdin lines: trimmed, lower-cased, de-duplicated, blanks skipped. */
export function parseAddressInput(args: string[], stdinText?: string): string[] {
  const lines = stdinText === undefined ? [] : stdinText.split(/\r?\n/);
  const seen = new Set<string>();
  for (const raw of [...args, ...lines]) {
    const addr = raw.trim().toLowerCase();
    if (addr) seen.add(addr);
  }
  if (seen.size === 0) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: 'No addresses given: pass them as arguments, or one per line with --stdin',
    });
  }
  if (seen.size > MAX_ADDRESSES) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: `Too many addresses: ${seen.size} given, at most 5,000 per call`,
    });
  }
  return [...seen];
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

function isThrottled(r: BatchResponse | undefined): boolean {
  return r?.status === 429 || r?.status === 503;
}

function retryAfterMs(r: BatchResponse | undefined, attempt: number): number {
  const header = Object.entries(r?.headers ?? {}).find(([k]) => k.toLowerCase() === 'retry-after');
  const seconds = header ? Number(header[1]) : NaN;
  const wait = Number.isFinite(seconds) && seconds >= 0 ? seconds : 2 ** attempt;
  return Math.min(wait, RETRY_AFTER_CAP_S) * 1000;
}

/** Error text for a final sub-response; a throttled one has used up its retries. */
function describeFailure(r: BatchResponse | undefined): string {
  if (!r) return 'no response for this request in the Graph batch';
  const err = (r.body as { error?: { code?: string; message?: string } } | undefined)?.error;
  const text = `Graph ${r.status} ${err?.code ?? `HTTP_${r.status}`}: ${err?.message ?? `HTTP ${r.status}`}`;
  return isThrottled(r) ? `${text} (still throttled after ${MAX_RETRIES} retries)` : text;
}

/**
 * Send `requests` through `POST /$batch`, 20 at a time. Items answered 429 or
 * 503 are retried on their own after the longest Retry-After among them
 * (capped at 30 s), up to 3 times; whatever is still throttled keeps its last
 * response. A whole-call failure throws from GraphClient.
 */
async function sendBatch(
  graph: GraphClient,
  requests: BatchRequest[],
  sleep: (ms: number) => Promise<void>,
): Promise<Map<string, BatchResponse>> {
  const out = new Map<string, BatchResponse>();
  let pending = requests;
  for (let attempt = 0; ; attempt++) {
    for (let i = 0; i < pending.length; i += BATCH_SIZE) {
      const { responses } = await graph.post<{ responses: BatchResponse[] }>('/$batch', {
        requests: pending.slice(i, i + BATCH_SIZE),
      });
      for (const r of responses) out.set(r.id, r);
    }
    const throttled = pending.filter((req) => isThrottled(out.get(req.id)));
    if (throttled.length === 0 || attempt === MAX_RETRIES) return out;
    await sleep(Math.max(...throttled.map((req) => retryAfterMs(out.get(req.id), attempt))));
    pending = throttled;
  }
}

function found(query: string, u: GraphUser): FoundUser {
  const m = u.manager;
  return {
    query,
    status: 'ok',
    id: u.id,
    displayName: u.displayName ?? null,
    mail: u.mail ?? null,
    userPrincipalName: u.userPrincipalName ?? null,
    jobTitle: u.jobTitle ?? null,
    department: u.department ?? null,
    officeLocation: u.officeLocation ?? null,
    companyName: u.companyName ?? null,
    accountEnabled: u.accountEnabled ?? null,
    manager: m
      ? { displayName: m.displayName ?? null, mail: m.mail ?? null, jobTitle: m.jobTitle ?? null }
      : null,
  };
}

function proxyUrl(addr: string): string {
  const filter = `proxyAddresses/any(p:p eq 'smtp:${addr.replace(/'/g, "''")}')`;
  return `/users?$filter=${encodeURIComponent(filter)}&$count=true&$select=${USER_SELECT}`;
}

async function lookup(
  graph: GraphClient,
  queries: string[],
  sleep: (ms: number) => Promise<void>,
): Promise<LookupUserEntry[]> {
  const entries = new Map<string, LookupUserEntry>();
  const valid: string[] = [];
  for (const q of queries) {
    if (!q.startsWith('-') && ADDRESS_RE.test(q)) valid.push(q);
    else entries.set(q, { query: q, status: 'error', error: INVALID_ADDRESS });
  }

  // Round 1: the address as a user id or UPN, manager expanded inline.
  const direct = await sendBatch(
    graph,
    valid.map((q, i) => ({
      id: String(i),
      method: 'GET',
      url: `/users/${encodeURIComponent(q)}?$select=${USER_SELECT}&$expand=${MANAGER_EXPAND}`,
    })),
    sleep,
  );
  const missing: string[] = [];
  valid.forEach((q, i) => {
    const r = direct.get(String(i));
    if (r?.status === 404) missing.push(q);
    else if (r?.status === 200) entries.set(q, found(q, r.body as GraphUser));
    else entries.set(q, { query: q, status: 'error', error: describeFailure(r) });
  });

  // Round 2: an alias that is not the UPN, through proxyAddresses. This is an
  // advanced query (ConsistencyLevel + $count), which may refuse $expand, so
  // the manager of a match comes from round 3.
  const byProxy = await sendBatch(
    graph,
    missing.map((q, i) => ({
      id: String(i),
      method: 'GET',
      url: proxyUrl(q),
      headers: { ConsistencyLevel: 'eventual' },
    })),
    sleep,
  );
  const aliased: Array<{ query: string; user: GraphUser }> = [];
  missing.forEach((q, i) => {
    const r = byProxy.get(String(i));
    const user = r?.status === 200 ? (r.body as { value?: GraphUser[] }).value?.[0] : undefined;
    if (user) aliased.push({ query: q, user });
    else if (r?.status === 200) entries.set(q, { query: q, status: 'not_found' });
    else entries.set(q, { query: q, status: 'error', error: describeFailure(r) });
  });

  // Round 3: the manager of each alias match.
  const managers = await sendBatch(
    graph,
    aliased.map(({ user }, i) => ({
      id: String(i),
      method: 'GET',
      url: `/users/${encodeURIComponent(user.id)}?$select=id&$expand=${MANAGER_EXPAND}`,
    })),
    sleep,
  );
  aliased.forEach(({ query, user }, i) => {
    const r = managers.get(String(i));
    if (r?.status === 200) {
      entries.set(query, found(query, { ...user, manager: (r.body as GraphUser).manager }));
    } else {
      entries.set(query, {
        query,
        status: 'error',
        error: `manager lookup failed: ${describeFailure(r)}`,
      });
    }
  });

  return queries.map((q) => entries.get(q) as LookupUserEntry);
}

export async function runLookupUsers(opts: LookupUsersOptions): Promise<LookupUsersResult> {
  const queries = parseAddressInput(opts.addresses);
  const graph = new GraphClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
  try {
    return { users: await lookup(graph, queries, opts.sleep ?? defaultSleep) };
  } catch (e) {
    if (e instanceof AuthRequiredError) {
      throw new ExitWithCode(ExitCode.AuthRequired, { code: 'auth_required', message: e.message });
    }
    if (e instanceof GraphHttpError) {
      throw new ExitWithCode(ExitCode.Upstream, {
        code: 'upstream',
        message: e.message,
        status: e.status,
      });
    }
    throw e;
  }
}
