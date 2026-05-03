// src/http/graph-client.ts
import type { Session } from '../session/store';
import { AuthRequiredError, GraphHttpError, UpstreamError } from './errors';

const GRAPH_AUDIENCE = 'https://graph.microsoft.com';

/**
 * Pick the right Bearer for Graph from a Session: prefer a multi-audience
 * token under tokens['https://graph.microsoft.com'] (Path B), fall back
 * to the legacy session.bearerToken for backward compat with single-token
 * sessions / older tests.
 */
function graphBearer(session: Session): string {
  const t = session.tokens?.[GRAPH_AUDIENCE];
  if (t?.bearerToken) return t.bearerToken;
  return session.bearerToken;
}

const DEFAULT_BASE = 'https://graph.microsoft.com/v1.0';

export interface GraphClientOptions {
  httpTimeoutMs: number;
  baseUrl?: string;
  retries?: number; // default 3
  retryBaseMs?: number; // default 500 (exponential)
}

export interface PaginateOptions {
  maxResults: number; // safety cap; cappedEarly:true if hit
  pageSize?: number; // optional $top to add to first call
}

export interface PaginateResult<T> {
  items: T[];
  cappedEarly: boolean;
}

interface GraphErrorBody {
  error?: { code?: string; message?: string };
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

export class GraphClient {
  private readonly base: string;
  private readonly retries: number;
  private readonly retryBaseMs: number;
  constructor(
    private readonly session: Session,
    private readonly opts: GraphClientOptions,
  ) {
    this.base = opts.baseUrl ?? DEFAULT_BASE;
    this.retries = opts.retries ?? 3;
    this.retryBaseMs = opts.retryBaseMs ?? 500;
  }

  private url(pathOrUrl: string): string {
    if (pathOrUrl.startsWith('http://') || pathOrUrl.startsWith('https://')) return pathOrUrl;
    return this.base + pathOrUrl;
  }

  private async request<T>(method: string, pathOrUrl: string, body?: unknown): Promise<T> {
    const url = this.url(pathOrUrl);
    const init: RequestInit = {
      method,
      headers: {
        Authorization: `Bearer ${graphBearer(this.session)}`,
        Accept: 'application/json',
        ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}),
      },
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
        const requestId = res.headers.get('request-id') ?? undefined;
        if (res.status === 401) {
          let msg = 'unauthenticated';
          try {
            msg = ((await res.json()) as GraphErrorBody).error?.message ?? msg;
          } catch {}
          throw new AuthRequiredError(msg, requestId);
        }
        if (res.status >= 200 && res.status < 300) {
          if (res.status === 204) return undefined as unknown as T;
          return (await res.json()) as T;
        }
        let body: GraphErrorBody = {};
        try {
          body = (await res.json()) as GraphErrorBody;
        } catch {}
        const code = body.error?.code ?? `HTTP_${res.status}`;
        const message = body.error?.message ?? `HTTP ${res.status}`;
        if (isRetryable(res.status) && attempt < this.retries) {
          lastErr = new UpstreamError(res.status, code, message, requestId);
          await sleep(this.retryBaseMs * Math.pow(2, attempt));
          continue;
        }
        if (isRetryable(res.status)) {
          throw new UpstreamError(res.status, code, message, requestId);
        }
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

  get<T>(path: string): Promise<T> {
    return this.request<T>('GET', path);
  }
  post<T>(path: string, body: unknown): Promise<T> {
    return this.request<T>('POST', path, body);
  }
  patch<T>(path: string, body: unknown): Promise<T> {
    return this.request<T>('PATCH', path, body);
  }
  delete<T>(path: string): Promise<T> {
    return this.request<T>('DELETE', path);
  }

  async paginate<T>(initialPath: string, opts: PaginateOptions): Promise<PaginateResult<T>> {
    const items: T[] = [];
    let url = initialPath;
    if (opts.pageSize && !url.includes('$top=')) {
      const sep = url.includes('?') ? '&' : '?';
      url = `${url}${sep}$top=${opts.pageSize}`;
    }
    let cappedEarly = false;
    while (url) {
      const page = await this.request<{ value: T[]; '@odata.nextLink'?: string }>('GET', url);
      for (const v of page.value) {
        if (items.length >= opts.maxResults) {
          cappedEarly = true;
          break;
        }
        items.push(v);
      }
      if (cappedEarly || !page['@odata.nextLink']) break;
      url = page['@odata.nextLink'];
    }
    return { items, cappedEarly };
  }
}
