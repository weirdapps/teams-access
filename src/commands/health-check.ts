// src/commands/health-check.ts
//
// Probes one read of each kind to detect breakage early. Reports per-probe
// status. Designed for cron / launchd scheduling so you find out the moment
// Microsoft changes a private API URL or scope expires.

import { GraphClient } from '../http/graph-client';
import { ChatsvcClient } from '../http/chatsvc-client';
import { ChatsvcaggClient } from '../http/chatsvcagg-client';
import type { Session } from '../session/store';

export interface HealthCheckOptions {
  session: Session;
  httpTimeoutMs: number;
  // Optional: a known thread id to probe for chat messages. If absent, we
  // try to discover one via chatsvcagg.
  probeChatThreadId?: string;
  // Optional: known team UUID + channel id for channel probe.
  probeTeamUuid?: string;
  probeChannelId?: string;
}

export interface HealthProbeResult {
  name: string;
  ok: boolean;
  status?: number;
  detail: string;
  durationMs: number;
}

export interface HealthCheckResult {
  overall: 'ok' | 'degraded' | 'broken';
  probes: HealthProbeResult[];
  account?: { upn?: string; oid?: string };
}

async function timed<T>(
  fn: () => Promise<T>,
): Promise<
  { ok: true; value: T; durationMs: number } | { ok: false; err: Error; durationMs: number }
> {
  const t0 = Date.now();
  try {
    const value = await fn();
    return { ok: true, value, durationMs: Date.now() - t0 };
  } catch (err) {
    return { ok: false, err: err as Error, durationMs: Date.now() - t0 };
  }
}

export async function runHealthCheck(opts: HealthCheckOptions): Promise<HealthCheckResult> {
  const probes: HealthProbeResult[] = [];
  let account: { upn?: string; oid?: string } | undefined;

  // Probe 1: Graph /me
  {
    const r = await timed(async () => {
      const g = new GraphClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
      return await g.get<{ id?: string; userPrincipalName?: string }>('/me');
    });
    if (r.ok) {
      account = { upn: r.value.userPrincipalName, oid: r.value.id };
      probes.push({
        name: 'graph_me',
        ok: true,
        detail: `upn=${r.value.userPrincipalName}`,
        durationMs: r.durationMs,
      });
    } else {
      probes.push({
        name: 'graph_me',
        ok: false,
        detail: r.err.message.slice(0, 200),
        durationMs: r.durationMs,
      });
    }
  }

  // Probe 2: Graph /me/joinedTeams
  {
    const r = await timed(async () => {
      const g = new GraphClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
      return await g.get<{ value: unknown[] }>('/me/joinedTeams');
    });
    if (r.ok) {
      probes.push({
        name: 'graph_joined_teams',
        ok: true,
        detail: `count=${r.value.value.length}`,
        durationMs: r.durationMs,
      });
    } else {
      probes.push({
        name: 'graph_joined_teams',
        ok: false,
        detail: r.err.message.slice(0, 200),
        durationMs: r.durationMs,
      });
    }
  }

  // Probe 3: Chatsvcagg v1/updates (chat list)
  let firstChatId: string | undefined = opts.probeChatThreadId;
  {
    const r = await timed(async () => {
      const c = new ChatsvcaggClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
      return await c.listChats();
    });
    if (r.ok) {
      probes.push({
        name: 'chatsvcagg_updates',
        ok: true,
        detail: `chats=${r.value.chats.length}`,
        durationMs: r.durationMs,
      });
      if (!firstChatId && r.value.chats.length > 0) {
        firstChatId = r.value.chats[0].id;
      }
    } else {
      probes.push({
        name: 'chatsvcagg_updates',
        ok: false,
        detail: r.err.message.slice(0, 200),
        durationMs: r.durationMs,
      });
    }
  }

  // Probe 4: Chatsvc messages (use first chat from discover, if available)
  if (firstChatId) {
    const r = await timed(async () => {
      const c = new ChatsvcClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
      return await c.getChatMessages(firstChatId!, { pageSize: 1 });
    });
    if (r.ok) {
      probes.push({
        name: 'chatsvc_messages',
        ok: true,
        detail: `messages=${r.value.messages.length} thread=${firstChatId}`,
        durationMs: r.durationMs,
      });
    } else {
      probes.push({
        name: 'chatsvc_messages',
        ok: false,
        detail: r.err.message.slice(0, 200),
        durationMs: r.durationMs,
      });
    }
  } else {
    probes.push({
      name: 'chatsvc_messages',
      ok: false,
      detail: 'skipped — no chat thread id available (provide --probe-chat-thread-id)',
      durationMs: 0,
    });
  }

  // Verdict
  const okCount = probes.filter((p) => p.ok).length;
  const overall: HealthCheckResult['overall'] =
    okCount === probes.length ? 'ok' : okCount === 0 ? 'broken' : 'degraded';

  return { overall, probes, account };
}
