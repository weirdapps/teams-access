// src/commands/list-chats.ts
//
// Path B: list chats via the chatsvcagg /discover endpoint. The /updates
// endpoint returns 500 without precise Teams-web headers; /discover works
// today and returns recently-active chats with continuation paging.

import { ChatsvcaggClient, type DiscoverChat } from '../http/chatsvcagg-client';
import { AuthRequiredError, GraphHttpError } from '../http/errors';
import { ExitCode, ExitWithCode } from '../util/exit-codes';
import type { Session } from '../session/store';

export interface ListChatsOptions {
  session: Session;
  httpTimeoutMs: number;
  limit?: number;       // client-side cap on returned chats (no API pagination)
}

export interface ListChatsResult {
  chats: DiscoverChat[];
  totalAvailable: number;
  source: 'updates_v1';
}

export async function runListChats(opts: ListChatsOptions): Promise<ListChatsResult> {
  const client = new ChatsvcaggClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
  try {
    const r = await client.listChats({ isPrefetch: false });
    const all = r.chats;
    const limit = opts.limit ?? all.length;
    return {
      chats: all.slice(0, limit),
      totalAvailable: all.length,
      source: 'updates_v1',
    };
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
