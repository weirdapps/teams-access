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
}

export type Scope =
  | { kind: 'chat'; threadId: string }
  | { kind: 'channel'; teamUuid: string; channelId: string };

export interface ListMessagesResult {
  scope: Scope;
  messages: ChatsvcMessage[]; // for chat
  posts?: ChannelPost[]; // for channel
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
  if (hasChat) return { kind: 'chat', threadId: opts.chat! };
  return { kind: 'channel', teamUuid: opts.team!, channelId: opts.channel! };
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

  try {
    if (scope.kind === 'chat') {
      const c = new ChatsvcClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
      const r = await c.getChatMessages(scope.threadId, { pageSize });
      return { scope, messages: r.messages };
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
