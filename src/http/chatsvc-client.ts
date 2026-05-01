// src/http/chatsvc-client.ts
//
// Path B: client for the private chatsvc API at teams.microsoft.com.
// Used for: chat messages content (text, HTML, mentions, sender, timestamps).
//
// Audience: https://ic3.teams.office.com
// Region prefix: typically "emea" for European tenants. Read from session.region.chatsvc.
//
// IMPORTANT: This is a private, undocumented Microsoft API. Endpoints may change
// without notice. When breakage is detected (via teams-cli health-check or in
// production), follow docs/private-api-cookbook.md to rediscover the URLs.

import type { Session } from '../session/store';
import { AuthRequiredError, GraphHttpError, UpstreamError } from './errors';
import { randomUUID } from 'node:crypto';

const AUDIENCE = 'https://ic3.teams.office.com';
const HOST = 'https://teams.microsoft.com';
const CLIENT_VERSION = '1415/25021922252';

export interface ChatsvcClientOptions {
  httpTimeoutMs: number;
  region?: string; // chatsvc region, defaults to "emea"
  retries?: number;
  retryBaseMs?: number;
}

export interface ChatsvcMessage {
  id: string;
  conversationid?: string;
  conversationLink?: string;
  contenttype?: string;
  type?: string;
  messagetype?: string;          // e.g. "Text", "RichText/Html", "ThreadActivity/AddMember", ...
  content?: string;              // raw body — text or HTML depending on messagetype
  from?: string;                 // contact link, ends with .../contacts/8:orgid:<oid>
  imdisplayname?: string;        // sender display name
  fromDisplayNameInToken?: string;
  composetime?: string;          // ISO timestamp
  originalarrivaltime?: string;
  sequenceId?: number;
  version?: string;
  clientmessageid?: string;
  properties?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface ChatsvcMessagesResponse {
  messages: ChatsvcMessage[];
  tenantId?: string;
  _metadata?: {
    lastCompleteSegmentStartTime?: number;
    lastCompleteSegmentEndTime?: number;
    backwardLink?: string;       // URL for paging backwards (older messages)
    syncState?: string;          // URL for incremental sync (newer messages)
  };
}

function isRetryable(status: number): boolean {
  return status === 408 || status === 429 || (status >= 500 && status < 600);
}

async function withTimeout<T>(p: Promise<T>, ms: number, label: string): Promise<T> {
  let to: NodeJS.Timeout | undefined;
  const timeout = new Promise<never>((_, reject) => {
    to = setTimeout(() => reject(new UpstreamError(599, 'Timeout', `${label} timed out after ${ms}ms`)), ms);
  });
  try {
    return await Promise.race([p, timeout]);
  } finally {
    if (to) clearTimeout(to);
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise(r => setTimeout(r, ms));
}

export class ChatsvcClient {
  private readonly region: string;
  private readonly retries: number;
  private readonly retryBaseMs: number;
  constructor(
    private readonly session: Session,
    private readonly opts: ChatsvcClientOptions,
  ) {
    this.region = opts.region ?? session.region?.chatsvc ?? 'emea';
    this.retries = opts.retries ?? 3;
    this.retryBaseMs = opts.retryBaseMs ?? 500;
  }

  private token(): string {
    const t = this.session.tokens?.[AUDIENCE];
    if (!t) {
      throw new AuthRequiredError(`No '${AUDIENCE}' token in session. Re-run \`teams-cli login\` and ensure the diagnostic window captures IC3 requests (open a chat in Teams).`);
    }
    return t.bearerToken;
  }

  private async request<T>(method: string, pathOrUrl: string, body?: unknown): Promise<T> {
    const url = pathOrUrl.startsWith('http') ? pathOrUrl : `${HOST}${pathOrUrl}`;
    const headers: Record<string, string> = {
      Authorization: `Bearer ${this.token()}`,
      Accept: 'application/json',
      'x-ms-client-version': CLIENT_VERSION,
      'x-ms-session-id': randomUUID(),
    };
    if (body !== undefined) headers['Content-Type'] = 'application/json';

    const init: RequestInit = {
      method,
      headers,
      ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
    };

    let lastErr: Error | undefined;
    for (let attempt = 0; attempt <= this.retries; attempt++) {
      try {
        const res = await withTimeout(fetch(url, init), this.opts.httpTimeoutMs, `${method} ${url}`);
        const requestId = res.headers.get('request-id') ?? res.headers.get('x-ms-correlation-id') ?? undefined;
        if (res.status === 401) {
          let msg = 'unauthenticated';
          try { msg = (await res.text()).slice(0, 200); } catch {}
          throw new AuthRequiredError(msg, requestId);
        }
        if (res.status >= 200 && res.status < 300) {
          if (res.status === 204) return undefined as unknown as T;
          return await res.json() as T;
        }
        let bodyText = '';
        try { bodyText = await res.text(); } catch {}
        const code = `HTTP_${res.status}`;
        const message = bodyText.slice(0, 200) || `HTTP ${res.status}`;
        if (isRetryable(res.status) && attempt < this.retries) {
          lastErr = new UpstreamError(res.status, code, message, requestId);
          await sleep(this.retryBaseMs * Math.pow(2, attempt));
          continue;
        }
        if (isRetryable(res.status)) throw new UpstreamError(res.status, code, message, requestId);
        throw new GraphHttpError(res.status, code, message, requestId);
      } catch (e) {
        if (e instanceof UpstreamError && attempt < this.retries) {
          lastErr = e;
          await sleep(this.retryBaseMs * Math.pow(2, attempt));
          continue;
        }
        throw e;
      }
    }
    throw lastErr ?? new UpstreamError(599, 'RetriesExhausted', 'No more retries available');
  }

  /**
   * Fetch messages from a chat thread.
   *
   * @param threadId  The chat's '19:...@unq.gbl.spaces' or '19:...@thread.v2' id.
   * @param opts.pageSize  Default 50.
   * @param opts.startTime  unix ms; default 1 (= "from beginning"). Use a recent
   *                        timestamp to fetch only new messages.
   * @param opts.syncState  URL from a prior response's _metadata.syncState; pass
   *                        this for incremental sync (will return only messages
   *                        newer than the last fetch).
   * @param opts.backwardLink  URL from prior response's _metadata.backwardLink
   *                           for paging through history.
   */
  async getChatMessages(threadId: string, opts: { pageSize?: number; startTime?: number; syncState?: string; backwardLink?: string } = {}): Promise<ChatsvcMessagesResponse> {
    // If a paging URL is provided, use it directly (Teams returns absolute URLs).
    if (opts.syncState) return await this.request<ChatsvcMessagesResponse>('GET', opts.syncState);
    if (opts.backwardLink) return await this.request<ChatsvcMessagesResponse>('GET', opts.backwardLink);

    const params = new URLSearchParams();
    params.set('view', 'msnp24Equivalent|supportsMessageProperties');
    params.set('pageSize', String(opts.pageSize ?? 50));
    params.set('startTime', String(opts.startTime ?? 1));
    const path = `/api/chatsvc/${this.region}/v1/users/ME/conversations/${encodeURIComponent(threadId)}/messages?${params.toString()}`;
    return await this.request<ChatsvcMessagesResponse>('GET', path);
  }
}
