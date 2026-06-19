// src/http/chatsvcagg-client.ts
//
// Path B: client for the private chatsvcagg API at teams.microsoft.com.
// Used for: chat list (via /discover or /updates) and channel posts.
//
// Audience: https://chatsvcagg.teams.microsoft.com
// Region prefix: typically "emea" for European tenants. Read from session.region.csa.
//
// IMPORTANT: This is a private, undocumented Microsoft API. Endpoints may change
// without notice. When breakage is detected (via teams-cli health-check or in
// production), follow docs/private-api-cookbook.md to rediscover the URLs.

import type { Session } from '../session/store';
import { AuthRequiredError, GraphHttpError, UpstreamError } from './errors';
import { randomUUID } from 'node:crypto';

const AUDIENCE = 'https://chatsvcagg.teams.microsoft.com';
const HOST = 'https://teams.microsoft.com';

// x-ms-client-version header value Teams web sends. Pinning a known-good value
// helps requests look more identical to real Teams traffic. May need refresh
// when Microsoft updates Teams web.
const CLIENT_VERSION = '1415/25021922252';

export interface ChatsvcaggClientOptions {
  httpTimeoutMs: number;
  region?: string; // csa region, defaults to "emea"
  retries?: number;
  retryBaseMs?: number;
}

export interface DiscoverChat {
  id: string;
  type?: string;
  threadType?: string;
  topic?: string | null;
  lastMessage?: { content?: string; from?: string; composeTime?: string };
  members?: Array<{ mri?: string; displayName?: string }>;
  // Other fields preserved as-is; consumer can introspect raw response.
  [key: string]: unknown;
}

export interface DiscoverChatsResponse {
  chats: DiscoverChat[];
  continuationToken?: string; // present only with discover-style pagination (deprecated)
  raw: unknown; // full response for debugging or extra fields
}

export interface ChannelPost {
  containerId: string;
  id: string;
  latestMessageTime?: string;
  message?: {
    messageType?: string;
    contentType?: string;
    content?: string;
    imDisplayName?: string;
    properties?: Record<string, unknown>;
    [key: string]: unknown;
  };
}

export interface ChannelPostsResponse {
  posts: ChannelPost[];
  hasMore?: boolean;
  hasMoreBackward?: boolean;
  continuationToken?: string;
}

function isRetryable(status: number): boolean {
  return status === 408 || status === 429 || (status >= 500 && status < 600);
}

async function withTimeout<T>(p: Promise<T>, ms: number, label: string): Promise<T> {
  let to: NodeJS.Timeout | undefined;
  const timeout = new Promise<never>((_, reject) => {
    to = setTimeout(
      () => reject(new UpstreamError(599, 'Timeout', `${label} timed out after ${ms}ms`)),
      ms,
    );
  });
  try {
    return await Promise.race([p, timeout]);
  } finally {
    if (to) clearTimeout(to);
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

export class ChatsvcaggClient {
  private readonly region: string;
  private readonly retries: number;
  private readonly retryBaseMs: number;
  constructor(
    private readonly session: Session,
    private readonly opts: ChatsvcaggClientOptions,
  ) {
    this.region = opts.region ?? session.region?.csa ?? 'emea';
    this.retries = opts.retries ?? 3;
    this.retryBaseMs = opts.retryBaseMs ?? 500;
  }

  private token(): string {
    const t = this.session.tokens?.[AUDIENCE];
    if (!t) {
      throw new AuthRequiredError(
        `No '${AUDIENCE}' token in session. Re-run \`teams-cli login\` and ensure the diagnostic window captures chatsvcagg requests (open the chat tab in Teams).`,
      );
    }
    return t.bearerToken;
  }

  private async request<T>(method: string, path: string, body?: unknown): Promise<T> {
    const url = `${HOST}${path}`;
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
        const res = await withTimeout(
          fetch(url, init),
          this.opts.httpTimeoutMs,
          `${method} ${url}`,
        );
        const requestId =
          res.headers.get('request-id') ?? res.headers.get('x-ms-correlation-id') ?? undefined;
        if (res.status === 401) {
          let msg = 'unauthenticated';
          try {
            msg = (await res.text()).slice(0, 200);
          } catch {
            /* ignore — keep default 'unauthenticated' message */
          }
          throw new AuthRequiredError(msg, requestId);
        }
        if (res.status >= 200 && res.status < 300) {
          if (res.status === 204) return undefined as unknown as T;
          return (await res.json()) as T;
        }
        // Non-2xx: parse body if possible.
        let bodyText = '';
        try {
          bodyText = await res.text();
        } catch {
          /* ignore — fall back to empty body */
        }
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
   * List chats via the v1/updates endpoint. THE chat list endpoint Teams web uses.
   *
   * IMPORTANT: v3/updates returns 500 (server-side bug or undocumented header
   * requirement); v1 and v2 work. We use v1.
   *
   * Returns the user's full chat list (no pagination — the response is one big
   * object with all chats + teams + channels). Can easily be 30+ MB for an
   * active user. Caller is responsible for filtering / sub-paging in memory.
   */
  async listChats(opts: { isPrefetch?: boolean } = {}): Promise<DiscoverChatsResponse> {
    const params = new URLSearchParams();
    params.set('isPrefetch', String(opts.isPrefetch ?? false));
    const path = `/api/csa/${this.region}/api/v1/teams/users/me/updates?${params.toString()}`;
    const raw = await this.request<{
      chats?: DiscoverChat[];
      teams?: unknown[];
      channels?: unknown[];
    }>('GET', path);
    return {
      chats: raw.chats ?? [],
      raw,
    };
  }

  /** Backwards-compat alias for the original method name. Prefer listChats(). */
  async listChatsViaDiscover(
    _opts: { pageSize?: number; continuationToken?: string } = {},
  ): Promise<DiscoverChatsResponse> {
    return this.listChats({ isPrefetch: false });
  }

  /**
   * List posts (top-level messages) in a channel.
   *
   * @param channelId          The channel's '19:...@thread.tacv2' id.
   * @param generalChannelId   The team's General channel id (also '19:...@thread.tacv2').
   *                           This is the chatsvcagg API's 'teamId' parameter — NOT the
   *                           team UUID. See spike-results.md for the quirk.
   */
  async listChannelPosts(
    channelId: string,
    generalChannelId: string,
    opts: { pageSize?: number } = {},
  ): Promise<ChannelPostsResponse> {
    const params = new URLSearchParams();
    params.set('modality', 'post');
    params.set('pageSize', String(opts.pageSize ?? 20));
    params.set('teamId', generalChannelId);
    const path = `/api/csa/${this.region}/api/v1/containers/${encodeURIComponent(channelId)}/posts?${params.toString()}`;
    return await this.request<ChannelPostsResponse>('GET', path);
  }

  /** Fetch the user's pinned channels. */
  async getPinnedChannels(): Promise<{ orderVersion?: number; pinChannelOrder: string[] }> {
    const path = `/api/csa/${this.region}/api/v1/teams/users/me/pinnedChannels`;
    return await this.request<{ orderVersion?: number; pinChannelOrder: string[] }>('GET', path);
  }
}
