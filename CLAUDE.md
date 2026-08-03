# teams-access (teams-cli)

TypeScript CLI for Microsoft Teams. Reads chats/channels via private Teams APIs (chatsvc + chatsvcagg, no admin consent needed); sends chat messages via Microsoft Graph. Same auth pattern as outlook-cli — capturing Bearer tokens from a live Teams browser session.

## Tech Stack

- TypeScript 6, compiled to CommonJS in `dist/`
- Node.js >= 20
- `commander` for CLI parsing, `playwright` for browser-based login
- `vitest` for tests, ESLint + Prettier, SonarCloud

## Build / Run

```bash
npm run build              # tsc + chmod dist/cli.js
npx tsc --noEmit           # type-check only
node dist/cli.js <command> # run compiled CLI
npm run cli                # dev run via ts-node (no compile)
```

**Always rebuild after code changes before running via `node dist/cli.js`.**

The binary is NOT on PATH — always invoke as `node ~/SourceCode/teams-access/dist/cli.js`.

## Test / Lint

```bash
npm test                   # vitest run (52 tests across 13 files)
npm run test:watch
npm run test:coverage      # v8 coverage
npm run lint
npm run format
./scripts/pii-gauntlet.sh  # tracked-file PII scan; pre-commit hook AND required CI job
```

## CLI Commands

All output JSON to stdout; errors JSON to stderr. Exit codes 0–6 (4=auth required, 5=upstream error).

| Command                       | Backend                                                    |
| ----------------------------- | ---------------------------------------------------------- |
| `login`                       | Playwright — captures Bearer tokens from Teams web session |
| `auth-check` / `auth-renew`   | Graph `/me` / headless Playwright                          |
| `list-teams`, `list-channels` | Graph                                                      |
| `list-chats`, `list-messages` | Private chatsvc/chatsvcagg APIs                            |
| `send-message`                | Graph POST (chat only; channel sends not supported)        |
| `resolve-mri`, `health-check` | Graph + all three backends                                 |

Session stored at `~/.teams-cli/session.json` (mode 0600, ~24h TTL).

## Code Organization

```text
src/
  cli.ts                    # Entrypoint + commander wiring
  auth/browser-capture.ts   # Playwright login
  commands/                 # One file per command
  config/load.ts            # Flag → env var → default precedence
  http/                     # graph-client, chatsvc-client, chatsvcagg-client, errors, types
  output/                   # json.ts, table.ts
  session/                  # jwt.ts, store.ts (atomic writes, 0600)
  util/                     # exit-codes.ts, redact.ts, time.ts
test_scripts/               # Mirrors src/ layout, vitest tests
scripts/pii-gauntlet.sh     # Tracked-file PII scan (pre-commit + CI gate)
dist/                       # Compiled output
docs/                       # private-api-cookbook.md, spike-results.md
```

## Key Conventions

- All output is JSON (stdout). All errors are JSON on stderr.
- `ExitWithCode` is the single throw-to-exit mechanism, caught in `main()`.
- Config precedence: CLI flag > env var > default. Env vars: `TEAMS_CLI_HTTP_TIMEOUT_MS`, `TEAMS_CLI_LOGIN_TIMEOUT_MS`, `TEAMS_CLI_CHROME_CHANNEL`.
- No tokens ever logged — `redact.ts` guards all error paths.
- `browser-capture.ts` and `cli.ts` excluded from coverage (Playwright-dependent / wiring only).
- Tests in `test_scripts/` (not `src/`), mirroring `src/` subdirectory layout.

## CI

- **ci.yml**: push/PR to master, Node 22, three jobs: lint (`tsc --noEmit`), test (`npm test`), pii-gauntlet (`bash scripts/pii-gauntlet.sh`)
- **sonarcloud.yml**: push/PR to master, coverage upload (skips if no `SONAR_TOKEN`)
- **dependabot-auto-merge.yml**: thin caller of `weirdapps/shared-workflows/.github/workflows/dependabot-auto-merge.yml@main`, passing `allow_major_in_group: false`. The merge logic is NOT in this repo; edit the shared workflow to change behaviour. Effect: patch/minor auto-squash, any major stays open for manual review
- **deps-refresh.yml**: thin caller of `weirdapps/shared-workflows/.github/workflows/deps-refresh.yml@main`; monthly (14th), gate is `tsc --noEmit && npm run build && npm test`

## Notes

- **Never trust `auth-check` as a session health signal.** It probes Graph `/me` only, so it returns `{"status":"ok"}` with exit 0 while the `chatsvc` / `chatsvcagg` / IC3 tokens are already expired and every `list-messages` call is 401-ing. Each audience has its own lifetime. Use `health-check` (four probes: `graph_me`, `graph_joined_teams`, `chatsvcagg_updates`, `chatsvc_messages`; exit 5 broken, 1 degraded) before concluding Teams reads work.
- `--no-auto-reauth` is declared in `src/cli.ts` but never read anywhere. It is a no-op and there is no auto-reauth: an expired session throws `auth_required` (exit 4). Do not document or rely on browser-reopening behaviour.
- The private chatsvc/chatsvcagg endpoints are undocumented. If `health-check` reports failures, see `docs/private-api-cookbook.md`.
- `[Claude]` prefix is mandatory on all Teams messages sent by automation (enforced at call site).
- Session tokens expire ~24h; `auth-renew` refreshes headlessly; `login` does full interactive re-auth.
- Writing a real home-directory path (`/Users/<name>/...`) or a personal email into any tracked file will fail `pii-gauntlet.sh` in pre-commit and in CI. Use `~/` or `/Users/user/` in examples.
