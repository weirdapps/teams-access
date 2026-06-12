# teams-cli

[![CI](https://github.com/weirdapps/teams-access/actions/workflows/ci.yml/badge.svg)](https://github.com/weirdapps/teams-access/actions/workflows/ci.yml)
[![SonarCloud](https://sonarcloud.io/api/project_badges/measure?project=weirdapps_teams-access&metric=alert_status)](https://sonarcloud.io/project/overview?id=weirdapps_teams-access)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js >=20](https://img.shields.io/badge/node-%3E%3D20-brightgreen.svg)](https://nodejs.org/)

A TypeScript command-line tool for Microsoft Teams. Reads chats and channels via private Teams APIs (no admin consent needed); sends chat messages via Microsoft Graph (using the delegated scope Teams web itself acquires).

Sister project to [`outlook-cli`](https://github.com/weirdapps/outlook-access). Same auth pattern (snoop the Bearer your browser session uses), same exit codes, same "no app registration" promise.

## Status

**v0.1.0 (Plan 1 + amendment 1 — Path B hybrid).** Eight commands shipped:

| Command                   | Backend                                                              | Verified                              |
| ------------------------- | -------------------------------------------------------------------- | ------------------------------------- |
| `login`                   | Playwright capture of Teams web Bearers                              | ✅ multi-audience                     |
| `auth-check`              | Microsoft Graph `/me`                                                | ✅                                    |
| `list-teams`              | Microsoft Graph `/me/joinedTeams`                                    | ✅                                    |
| `list-channels`           | Microsoft Graph `/teams/{id}/channels`                               | ✅                                    |
| `list-chats`              | private chatsvcagg `/api/csa/{region}/api/v1/teams/users/me/updates` | ✅ (5,901 chats on test account)      |
| `list-messages` (chat)    | private chatsvc `/api/chatsvc/{region}/v1/.../messages`              | ✅                                    |
| `list-messages` (channel) | private chatsvcagg `/api/csa/{region}/api/v1/containers/{id}/posts`  | ✅                                    |
| `send-message` (chat)     | Microsoft Graph `POST /chats/{id}/messages`                          | ⚠ scope verified, end-to-end deferred |
| `health-check`            | exercises one of each                                                | ✅                                    |

## Why this approach (Path B)

The original plan was to drive everything via Microsoft Graph. The spike (see `docs/spike-results.md`) found that the Teams web client's Graph token, on an enterprise tenant with restrictive scope policies, does **not** carry the scopes Microsoft Graph requires for `Chat.Read*` or `ChannelMessage.Read.All`. Graph rejects those calls with 403.

But Teams web itself reads chats and channels constantly — it just uses _different_ services with _different_ audience-bound tokens:

- `chatsvc` (audience `https://ic3.teams.office.com`) for chat message content
- `chatsvcagg` (audience `https://chatsvcagg.teams.microsoft.com`) for chat list + channel posts

Both APIs are private and undocumented. We capture all the audience tokens Teams web acquires at sign-in and use the right one per operation. Graph still gets used where it works (account info, team/channel listing, sending chat messages, calendar metadata).

## Prerequisites

- Node.js 20+
- A real Google Chrome or Microsoft Edge installation (Playwright launches your installed browser via `channel`)
- A Microsoft 365 work/school mailbox you can sign in to at `teams.microsoft.com`

## Build

```bash
git clone <this-repo> teams-access
cd teams-access
npm install
npm run build
```

Optional global install:

```bash
npm link
```

## First use

```bash
teams-cli login --diagnostic-extra-ms 30000
```

A Chrome window opens at `teams.microsoft.com`. Sign in normally. The script captures every Bearer token Teams web acquires for the next 30 seconds (the diagnostic window) and persists them all to `~/.teams-cli/session.json`.

**Important during login**: click the **Calendar tab** or **Files tab** in Teams web while the browser is open. Those trigger Microsoft Graph requests, so the Graph-audience token gets captured. Without that, commands using Graph (`auth-check`, `list-teams`, `list-channels`, `send-message`) will fail with 401.

After login:

```bash
teams-cli auth-check         # confirms the cached session works against Graph
teams-cli health-check       # probes one read of each kind
teams-cli list-chats --limit 10
teams-cli list-teams
```

## Configuration

| Setting        | Flag                      | Env var                      | Default |
| -------------- | ------------------------- | ---------------------------- | ------- |
| HTTP timeout   | `--timeout <ms>`          | `TEAMS_CLI_HTTP_TIMEOUT_MS`  | 30000   |
| Login timeout  | `--login-timeout <ms>`    | `TEAMS_CLI_LOGIN_TIMEOUT_MS` | 300000  |
| Chrome channel | `--chrome-channel <name>` | `TEAMS_CLI_CHROME_CHANNEL`   | chrome  |

Precedence: CLI flag > env var > default.

## Exit codes

| Code | Meaning                              |
| ---- | ------------------------------------ |
| 0    | Success                              |
| 1    | Unexpected internal error            |
| 2    | Invalid usage                        |
| 3    | Configuration error                  |
| 4    | Auth failure (run `teams-cli login`) |
| 5    | Upstream API error                   |
| 6    | IO error                             |

## Session file format

`~/.teams-cli/session.json` (mode 0600 in 0700 dir, atomic writes, never logged):

```json
{
  "bearerToken": "<the FIRST captured token, used as fallback for legacy code>",
  "cookies": [...],
  "capturedAt": "2026-05-01T...",
  "account": { "upn": "...", "oid": "...", "tid": "..." },
  "tokens": {
    "https://graph.microsoft.com": {"bearerToken": "...", "scp": "...", ...},
    "https://ic3.teams.office.com": {"bearerToken": "...", ...},
    "https://chatsvcagg.teams.microsoft.com": {"bearerToken": "...", ...},
    "https://api.spaces.skype.com": {"bearerToken": "...", ...},
    "https://presence.teams.microsoft.com/": {"bearerToken": "...", ...}
  },
  "region": { "chatsvc": "emea", "csa": "emea", "mt": "emea-03", "asyncgw": "eu-prod" },
  "generalChannelByTeamId": { "<team-uuid>": "19:...@thread.tacv2" }
}
```

## Brittleness

The chatsvc and chatsvcagg endpoints are private. Microsoft can change their URLs, response shapes, or accepted parameters without notice. If `health-check` starts reporting failures on `chatsvc_messages` or `chatsvcagg_updates`, see `docs/private-api-cookbook.md` for the rediscovery workflow.

## Testing

```bash
npm test                     # vitest run — 49 tests
npm run test:watch
```

## Security posture

- Cached session at `~/.teams-cli/session.json` contains live Bearer tokens for ≥6 audiences, expires ~24h after capture.
- Atomic writes (temp + rename), file mode 0600, parent dir 0700.
- No tokens ever logged or printed. Body-snippet redaction is in place for error paths.
- `~/.teams-cli/login-trace.jsonl` (audit trail of capture) records (URL, audience) tuples but **no token bytes**.
- `~/.teams-cli/multi-tokens.json` (legacy diagnostic side-file) DOES contain token bytes — same 0600 protection.

## License

MIT — see [LICENSE](./LICENSE).
