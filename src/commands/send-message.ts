// src/commands/send-message.ts
//
// POST a message to a chat or channel via Microsoft Graph.
//
// Scope status (verified in spike-results.md):
//   - chat send: ChatMessage.Send IS in token → works
//   - channel send: ChannelMessage.Send is NOT in token → 403 expected
//                   The CLI accepts the channel send call but the user
//                   should expect failure until admin consent or alternative
//                   approach is added.

import { GraphClient } from '../http/graph-client';
import { AuthRequiredError, GraphHttpError } from '../http/errors';
import { ExitCode, ExitWithCode } from '../util/exit-codes';
import type { Session } from '../session/store';

export interface SendMessageOptions {
  session: Session;
  httpTimeoutMs: number;
  chat?: string;             // chat thread id
  team?: string;             // team uuid
  channel?: string;          // channel id
  text?: string;             // plain text body
  html?: string;             // HTML body (mutually exclusive with text)
  replyTo?: string;          // for channel replies, parent message id
}

export interface SendMessageResult {
  status: 'ok';
  messageId: string;
  webUrl?: string;
}

export async function runSendMessage(opts: SendMessageOptions): Promise<SendMessageResult> {
  // Scope validation
  const hasChat = !!opts.chat;
  const hasChannel = !!opts.team && !!opts.channel;
  if (hasChat && hasChannel) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: '--chat is mutually exclusive with --team/--channel',
    });
  }
  if (!hasChat && !hasChannel) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: 'Provide either --chat <id> OR both --team <id> and --channel <id>',
    });
  }
  // Body validation
  if (opts.text && opts.html) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: '--text and --html are mutually exclusive',
    });
  }
  if (!opts.text && !opts.html) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: 'Provide either --text "..." or --html "..."',
    });
  }

  const body = {
    body: {
      contentType: opts.html ? 'html' as const : 'text' as const,
      content: opts.html ?? opts.text!,
    },
  };

  let path: string;
  if (hasChat) {
    path = `/chats/${encodeURIComponent(opts.chat!)}/messages`;
  } else {
    path = opts.replyTo
      ? `/teams/${encodeURIComponent(opts.team!)}/channels/${encodeURIComponent(opts.channel!)}/messages/${encodeURIComponent(opts.replyTo)}/replies`
      : `/teams/${encodeURIComponent(opts.team!)}/channels/${encodeURIComponent(opts.channel!)}/messages`;
  }

  const c = new GraphClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
  try {
    const res = await c.post<{ id: string; webUrl?: string }>(path, body);
    return { status: 'ok', messageId: res.id, webUrl: res.webUrl };
  } catch (e) {
    if (e instanceof AuthRequiredError) {
      throw new ExitWithCode(ExitCode.AuthRequired, { code: 'auth_required', message: e.message });
    }
    if (e instanceof GraphHttpError) {
      throw new ExitWithCode(ExitCode.Upstream, {
        code: 'upstream',
        message: e.message + (hasChannel ? ' (channel sends require ChannelMessage.Send scope which is NOT in the Teams-web Graph token; this command currently only works for --chat)' : ''),
        status: e.status,
      });
    }
    throw e;
  }
}
