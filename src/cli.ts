#!/usr/bin/env node
// src/cli.ts
import { Command } from 'commander';
import { loadConfig } from './config/load';
import { readSession } from './session/store';
import { writeJson } from './output/json';
import { ExitCode, ExitWithCode } from './util/exit-codes';
import { runLogin } from './commands/login';
import { runAuthCheck } from './commands/auth-check';
import { join } from 'node:path';
import { homedir } from 'node:os';

const program = new Command();
program
  .name('teams-cli')
  .description('Microsoft Teams CLI via captured Bearer token (mirrors outlook-cli).')
  .version('0.1.0')
  .option('--timeout <ms>', 'Per-HTTP-call timeout', v => Number(v))
  .option('--login-timeout <ms>', 'Max wait for interactive login', v => Number(v))
  .option('--chrome-channel <name>', 'Playwright Chrome channel')
  .option('--no-auto-reauth', 'Do NOT silently reopen the login browser on expired session');

program
  .command('login')
  .description('Capture a Teams web session by signing in via Playwright Chrome window.')
  .action(async () => {
    const config = loadConfig({
      timeoutMs: program.opts().timeout,
      loginTimeoutMs: program.opts().loginTimeout,
      chromeChannel: program.opts().chromeChannel,
    });
    const result = await runLogin({
      config,
      profileDir: join(process.env.HOME ?? homedir(), '.teams-cli', 'playwright-profile'),
    });
    writeJson(result);
  });

program
  .command('auth-check')
  .description('Verify the cached session is still accepted by Microsoft Graph.')
  .action(async () => {
    const config = loadConfig({
      timeoutMs: program.opts().timeout,
      loginTimeoutMs: program.opts().loginTimeout,
      chromeChannel: program.opts().chromeChannel,
    });
    const session = readSession();
    if (!session) {
      throw new ExitWithCode(ExitCode.AuthRequired, {
        code: 'auth_required',
        message: 'No cached session. Run `teams-cli login`.',
      });
    }
    const result = await runAuthCheck({ session, httpTimeoutMs: config.httpTimeoutMs });
    writeJson(result);
  });

async function main(): Promise<void> {
  try {
    await program.parseAsync(process.argv);
  } catch (e) {
    if (e instanceof ExitWithCode) {
      process.stderr.write(JSON.stringify(e.payload) + '\n');
      process.exit(e.code);
    }
    process.stderr.write(JSON.stringify({
      code: 'internal',
      message: (e as Error).message ?? String(e),
    }) + '\n');
    process.exit(ExitCode.Internal);
  }
}

main();
