#!/usr/bin/env node
// src/cli.ts
import { Command } from 'commander';
import { loadConfig } from './config/load';
import { readSession, type Session } from './session/store';
import { writeJson } from './output/json';
import { ExitCode, ExitWithCode } from './util/exit-codes';
import { runLogin } from './commands/login';
import { runAuthCheck } from './commands/auth-check';
import { runListTeams } from './commands/list-teams';
import { runListChannels } from './commands/list-channels';
import { runListChats } from './commands/list-chats';
import { runListMessages } from './commands/list-messages';
import { runSendMessage } from './commands/send-message';
import { runHealthCheck } from './commands/health-check';
import { runResolveMri } from './commands/resolve-mri';
import { runAuthRenew } from './commands/auth-renew';
import { join } from 'node:path';
import { homedir } from 'node:os';

const program = new Command();
program
  .name('teams-cli')
  .description(
    'Microsoft Teams CLI via captured Bearer token. Path B: Graph + private chatsvc/chatsvcagg.',
  )
  .version('0.1.0')
  .option('--timeout <ms>', 'Per-HTTP-call timeout', (v) => Number(v))
  .option('--login-timeout <ms>', 'Max wait for interactive login', (v) => Number(v))
  .option('--chrome-channel <name>', 'Playwright Chrome channel')
  .option('--no-auto-reauth', 'Do NOT silently reopen the login browser on expired session');

function loadSessionOrThrow(): Session {
  const s = readSession();
  if (!s) {
    throw new ExitWithCode(ExitCode.AuthRequired, {
      code: 'auth_required',
      message: 'No cached session. Run `teams-cli login`.',
    });
  }
  return s;
}

function commonConfig(): { httpTimeoutMs: number } {
  const config = loadConfig({
    timeoutMs: program.opts().timeout,
    loginTimeoutMs: program.opts().loginTimeout,
    chromeChannel: program.opts().chromeChannel,
  });
  return { httpTimeoutMs: config.httpTimeoutMs };
}

program
  .command('login')
  .description('Capture a Teams web session by signing in via Playwright Chrome window.')
  .option(
    '--diagnostic-extra-ms <ms>',
    'After capture, keep browser open this long to collect more audience tokens',
    (v) => Number(v),
  )
  .action(async (cmd) => {
    const config = loadConfig({
      timeoutMs: program.opts().timeout,
      loginTimeoutMs: program.opts().loginTimeout,
      chromeChannel: program.opts().chromeChannel,
    });
    const result = await runLogin({
      config,
      profileDir: join(process.env.HOME ?? homedir(), '.teams-cli', 'playwright-profile'),
      diagnosticExtraMs: cmd.diagnosticExtraMs,
    });
    writeJson(result);
  });

program
  .command('auth-check')
  .description('Verify the cached session is still accepted by Microsoft Graph.')
  .action(async () => {
    const session = loadSessionOrThrow();
    const result = await runAuthCheck({ session, ...commonConfig() });
    writeJson(result);
  });

program
  .command('list-teams')
  .description('List teams I belong to.')
  .action(async () => {
    const session = loadSessionOrThrow();
    const result = await runListTeams({ session, ...commonConfig() });
    writeJson(result);
  });

program
  .command('list-channels')
  .description('List channels in a team or across all my teams.')
  .option('--team-id <uuid>', 'List channels of one team')
  .option('--all-teams', 'Flatten channels across every team I belong to')
  .action(async (cmd) => {
    const session = loadSessionOrThrow();
    const result = await runListChannels({
      session,
      ...commonConfig(),
      teamId: cmd.teamId,
      allTeams: !!cmd.allTeams,
    });
    writeJson(result);
  });

program
  .command('list-chats')
  .description('List my chats (1:1, group, meeting). Path B: chatsvcagg v1/updates.')
  .option('--limit <n>', 'Max chats to return (default: all)', (v) => Number(v))
  .action(async (cmd) => {
    const session = loadSessionOrThrow();
    const result = await runListChats({
      session,
      ...commonConfig(),
      limit: cmd.limit,
    });
    writeJson(result);
  });

program
  .command('list-messages')
  .description(
    'List messages in a chat OR channel. Path B: chatsvc (chat) or chatsvcagg (channel).',
  )
  .option('--chat <id>', 'Chat thread id (mutually exclusive with --team/--channel)')
  .option('--team <uuid>', 'Team UUID (paired with --channel)')
  .option('--channel <id>', 'Channel id (paired with --team)')
  .option('--page-size <n>', 'Page size', (v) => Number(v), 50)
  .action(async (cmd) => {
    const session = loadSessionOrThrow();
    const result = await runListMessages({
      session,
      ...commonConfig(),
      pageSize: cmd.pageSize,
      chat: cmd.chat,
      team: cmd.team,
      channel: cmd.channel,
    });
    writeJson(result);
  });

program
  .command('resolve-mri <mri>')
  .description(
    'Resolve a Teams MRI (8:orgid:<aad-oid>) to {id, email, displayName} via Graph /users/{id}',
  )
  .action(async (mri: string) => {
    const session = loadSessionOrThrow();
    const result = await runResolveMri({ session, ...commonConfig(), mri });
    writeJson(result);
  });

program
  .command('auth-renew')
  .description('Silently renew the Teams Bearer using the persisted browser profile (headless)')
  .option('--timeout <ms>', 'Headless capture timeout (default 30000)', (v) => Number(v))
  .action(async (cmd) => {
    const result = await runAuthRenew({
      timeoutMs: cmd.timeout,
      chromeChannel: program.opts().chromeChannel,
    });
    writeJson(result);
  });

program
  .command('send-message')
  .description('Send a message to a chat (Graph). Channel sends NOT supported (scope missing).')
  .option('--chat <id>', 'Chat id')
  .option('--team <uuid>', 'Team UUID (with --channel)')
  .option('--channel <id>', 'Channel id (with --team)')
  .option('--text <text>', 'Plain text body')
  .option('--html <html>', 'HTML body')
  .option('--reply-to <messageId>', 'For channel replies, parent message id')
  .action(async (cmd) => {
    const session = loadSessionOrThrow();
    const result = await runSendMessage({
      session,
      ...commonConfig(),
      chat: cmd.chat,
      team: cmd.team,
      channel: cmd.channel,
      text: cmd.text,
      html: cmd.html,
      replyTo: cmd.replyTo,
    });
    writeJson(result);
  });

program
  .command('health-check')
  .description('Probe one read of each kind (Graph + chatsvc + chatsvcagg) and report.')
  .option(
    '--probe-chat-thread-id <id>',
    'Use this chat thread id for the chatsvc probe instead of auto-discovering',
  )
  .action(async (cmd) => {
    const session = loadSessionOrThrow();
    const result = await runHealthCheck({
      session,
      ...commonConfig(),
      probeChatThreadId: cmd.probeChatThreadId,
    });
    writeJson(result);
    if (result.overall === 'broken') process.exit(ExitCode.Upstream);
    if (result.overall === 'degraded') process.exit(ExitCode.Internal); // soft alert
  });

async function main(): Promise<void> {
  try {
    await program.parseAsync(process.argv);
  } catch (e) {
    if (e instanceof ExitWithCode) {
      process.stderr.write(JSON.stringify(e.payload) + '\n');
      process.exit(e.code);
    }
    process.stderr.write(
      JSON.stringify({
        code: 'internal',
        message: (e as Error).message ?? String(e),
      }) + '\n',
    );
    process.exit(ExitCode.Internal);
  }
}

main();
