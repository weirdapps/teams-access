// src/commands/list-messages.ts
//
// Path B: list messages in a chat OR channel. Mutually exclusive flags:
//   --chat <threadId>                                  (chat scope, uses ChatsvcClient)
//   --team <teamUuid> --channel <channelId>            (channel scope, uses ChatsvcaggClient)
//
// For channel scope we need the team's General-channel-id as the chatsvcagg
// 'teamId' parameter. We look it up via Graph and cache per-team in session.

import { GraphClient } from '../http/graph-client';
import { ChatsvcClient, type ChatsvcMessage } from '../http/chatsvc-client';
import { ChatsvcaggClient, type ChannelPost } from '../http/chatsvcagg-client';
import { AuthRequiredError, GraphHttpError } from '../http/errors';
import { ExitCode, ExitWithCode } from '../util/exit-codes';
import type { Session } from '../session/store';
import type { Channel } from '../http/types';
import { writeSession } from '../session/store';

export interface ListMessagesOptions {
  session: Session;
  httpTimeoutMs: number;
  pageSize?: number;
  chat?: string;
  team?: string;
  channel?: string;
  /** Chat scope: page back until a page reaches this time (ISO 8601 or unix ms). */
  until?: string;
  /** Chat scope: the most pages one call reads. Default 1, the newest page only. */
  maxPages?: number;
}

export type Scope =
  { kind: 'chat'; threadId: string } | { kind: 'channel'; teamUuid: string; channelId: string };

export interface ListMessagesResult {
  scope: Scope;
  messages: ChatsvcMessage[]; // for chat
  posts?: ChannelPost[]; // for channel
  /** Chat scope: the pages read, newest first. */
  pagesRead?: number;
  /**
   * Chat scope: false when the read stopped at maxPages before reaching `until`
   * (or the start of the chat), so older messages between are not in `messages`.
   */
  complete?: boolean;
}

function validateScope(opts: ListMessagesOptions): Scope {
  const hasChat = !!opts.chat;
  const hasTeam = !!opts.team;
  const hasChannel = !!opts.channel;
  if (hasChat && (hasTeam || hasChannel)) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: '--chat is mutually exclusive with --team / --channel',
    });
  }
  if (!hasChat && !(hasTeam && hasChannel)) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: 'Provide either --chat <id> OR both --team <uuid> and --channel <id>',
    });
  }
  if (!hasChat && (opts.until !== undefined || opts.maxPages !== undefined)) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: '--until and --max-pages page back through a --chat only',
    });
  }
  if (hasChat) return { kind: 'chat', threadId: opts.chat! };
  return { kind: 'channel', teamUuid: opts.team!, channelId: opts.channel! };
}

/** `--until` as unix ms: digits are ms already, anything else must parse as a date. */
function parseUntil(until: string | undefined): number | undefined {
  if (until === undefined) return undefined;
  const ms = /^\d+$/.test(until) ? Number(until) : Date.parse(until);
  if (!Number.isFinite(ms)) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: '--until must be an ISO 8601 time or unix milliseconds',
    });
  }
  return ms;
}

function parseMaxPages(maxPages: number | undefined): number {
  if (maxPages === undefined) return 1;
  if (!Number.isInteger(maxPages) || maxPages < 1) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: '--max-pages must be a whole number of 1 or more',
    });
  }
  return maxPages;
}

/** The oldest time a page holds, in unix ms, or undefined when no message carries one. */
function oldestOf(messages: ChatsvcMessage[]): number | undefined {
  let oldest: number | undefined;
  for (const m of messages) {
    const t = Date.parse(m.composetime ?? m.originalarrivaltime ?? '');
    if (Number.isFinite(t) && (oldest === undefined || t < oldest)) oldest = t;
  }
  return oldest;
}

/**
 * Read a chat newest page first, following each page's backwardLink until a page
 * reaches `until`, the chat runs out, or `maxPages` pages are read. One page held
 * the newest 200 messages and no more, so a chat busier than that between two
 * reads lost the rest for good: the busiest chat fills 200 in about 37 hours.
 */
async function readChat(
  c: ChatsvcClient,
  threadId: string,
  pageSize: number,
  until: number | undefined,
  maxPages: number,
): Promise<{ messages: ChatsvcMessage[]; pagesRead: number; complete: boolean }> {
  let page = await c.getChatMessages(threadId, { pageSize });
  const messages = [...page.messages];
  let pagesRead = 1;
  for (;;) {
    const oldest = oldestOf(page.messages);
    const reached = until === undefined || (oldest !== undefined && oldest <= until);
    const link = page._metadata?.backwardLink;
    if (reached || !link || page.messages.length === 0)
      return { messages, pagesRead, complete: true };
    if (pagesRead >= maxPages) return { messages, pagesRead, complete: false };
    page = await c.getChatMessages(threadId, { backwardLink: link });
    messages.push(...page.messages);
    pagesRead += 1;
  }
}

/**
 * Look up a team's General-channel-id. Required as the 'teamId' parameter for
 * chatsvcagg channel-message reads. Cached in session.generalChannelByTeamId
 * so we only call Graph once per team.
 */
async function generalChannelId(
  session: Session,
  httpTimeoutMs: number,
  teamUuid: string,
): Promise<string> {
  const cached = session.generalChannelByTeamId?.[teamUuid];
  if (cached) return cached;
  // Ask Graph for the team's channels and find the one named 'General'.
  const graph = new GraphClient(session, { httpTimeoutMs });
  const r = await graph.get<{ value: Channel[] }>(
    `/teams/${encodeURIComponent(teamUuid)}/channels`,
  );
  const general = r.value.find((c) => c.displayName === 'General');
  if (!general) {
    throw new ExitWithCode(ExitCode.Upstream, {
      code: 'upstream',
      message: `Team ${teamUuid} has no channel named "General" — cannot derive chatsvcagg teamId.`,
    });
  }
  // Cache + persist.
  if (!session.generalChannelByTeamId) session.generalChannelByTeamId = {};
  session.generalChannelByTeamId[teamUuid] = general.id;
  writeSession(session);
  return general.id;
}

export async function runListMessages(opts: ListMessagesOptions): Promise<ListMessagesResult> {
  const scope = validateScope(opts);
  const pageSize = opts.pageSize ?? 50;
  const until = parseUntil(opts.until);
  const maxPages = parseMaxPages(opts.maxPages);

  try {
    if (scope.kind === 'chat') {
      const c = new ChatsvcClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
      const r = await readChat(c, scope.threadId, pageSize, until, maxPages);
      return { scope, ...r };
    }
    // channel
    const generalId = await generalChannelId(opts.session, opts.httpTimeoutMs, scope.teamUuid);
    const cagg = new ChatsvcaggClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
    const r = await cagg.listChannelPosts(scope.channelId, generalId, { pageSize });
    // Channel returns posts (top-level messages); map to messages-like for uniformity.
    return { scope, messages: [], posts: r.posts };
  } catch (e) {
    if (e instanceof AuthRequiredError) {
      throw new ExitWithCode(ExitCode.AuthRequired, {
        code: 'auth_required',
        message: e.message,
      });
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
