# Private API cookbook — when chatsvc / chatsvcagg breaks

This file is the runbook for when one of the private Microsoft Teams APIs we depend on stops working. The brittleness is by design — we accept it in exchange for not needing tenant admin consent. When the inevitable happens (Microsoft updates Teams web, changes URL paths, requires new headers), the steps below recover the system.

## Symptom: `teams-cli health-check` shows a probe failing

Example:

```json
{
  "overall": "degraded",
  "probes": [
    { "name": "graph_me", "ok": true, ... },
    { "name": "chatsvcagg_updates", "ok": false, "detail": "Graph 500 ...", "durationMs": 1234 },
    ...
  ]
}
```

Or you call `list-chats` / `list-messages` and get a 4xx/5xx that wasn't there yesterday.

## Step 1: confirm the URL/audience hasn't drifted

Re-run login with full diagnostic capture:

```bash
rm -f ~/.teams-cli/session.json
teams-cli login --diagnostic-extra-ms 60000
```

While the browser is open: open the Chat tab, click into a chat, scroll its history, click into a Channel, scroll. Whatever operation is failing — exercise it manually in the Teams UI so the request fires and gets traced.

After the diagnostic window closes, inspect `~/.teams-cli/login-trace.jsonl`:

```bash
# Show all distinct (audience, method, host, path) tuples
jq -r '"\(.aud) | \(.method) \(.host)\(.path[0:120])"' \
    ~/.teams-cli/login-trace.jsonl | sort -u
```

Compare to what's coded in `src/http/chatsvc-client.ts` and `src/http/chatsvcagg-client.ts`. Look for:

- New URL paths with familiar audiences (Microsoft moved the endpoint)
- New audience claims you've never seen (Microsoft introduced a new service)
- Different region segments (`emea` → something else)

## Step 2: probe the new URL

Get the audience-bound token from `~/.teams-cli/multi-tokens.json` and curl the new URL:

```bash
TOK=$(jq -r '."https://chatsvcagg.teams.microsoft.com".token' ~/.teams-cli/multi-tokens.json)
curl -sS -w "\nHTTP %{http_code}\n" \
  -H "Authorization: Bearer $TOK" \
  -H "Accept: application/json" \
  -H "x-ms-client-version: 1415/25021922252" \
  -H "x-ms-session-id: $(uuidgen)" \
  "https://teams.microsoft.com/<NEW PATH HERE>"
```

If 200, you've found the replacement URL. Update the client code (one constant at the top of the relevant client file). Update `docs/spike-results.md` with the new mapping. Bump CLIENT_VERSION constant if Teams web has updated theirs.

## Step 3: when headers are the issue

Sometimes the same URL returns 200 from Teams web but 500 from our curl, even with identical Bearer + audience. That means Teams web is sending an additional header we're not. Common culprits:

- `x-ms-correlation-id`: random UUID per request
- `x-ms-scenario-id`: scenario tracking
- `behavioroverride: redirectAs404`: server behavior toggle
- `referer: https://teams.microsoft.com/`: origin check
- `cookie`: session cookies (we have them in session.json)

To inspect what Teams web actually sends, modify `src/auth/browser-capture.ts` to log full request headers temporarily (be careful — Cookie and Authorization headers contain secrets). Or use Chrome DevTools' Network tab during a real Teams session.

## Step 4: when audience claims have changed

If a previously-working audience claim is no longer issued by Teams web (e.g., they renamed `https://api.spaces.skype.com` to something else), update the AUDIENCE constant in the relevant client file and the documentation.

If a NEW audience appears in the trace and is needed for the failing operation, add a new client class. Pattern: copy `chatsvcagg-client.ts`, change the AUDIENCE constant, change the URL paths, write the response interfaces.

## Step 5: when Microsoft Graph scope set changes

The Teams-web Graph token's `scp` claim is set by Microsoft on the Teams web client itself, not by your tenant. Microsoft can add scopes (great — more works) or remove scopes (bad — some Graph commands stop working).

If `auth-check` succeeds but a specific Graph command starts failing 403 with "missing scope X", the scope was probably removed from the Teams-web Graph token. Check the actual scope set:

```bash
node -e "
const t = require('/Users/<you>/.teams-cli/multi-tokens.json')['https://graph.microsoft.com'].token;
const p = t.split('.')[1].replace(/-/g,'+').replace(/_/g,'/');
const padded = p + '='.repeat((4 - p.length % 4) % 4);
console.log(JSON.parse(Buffer.from(padded,'base64').toString('utf8')).scp);
"
```

If the scope was removed, the affected Graph command is now blocked at the audience level. Either (a) wait for Microsoft to add it back, (b) add a Path-B equivalent via the private APIs, or (c) accept admin consent as the right answer for that one operation.

## Step 6: when tokens expire mid-operation

Tokens last ~24h after capture. Multi-token sessions may have audiences with different expiry times. If a command starts failing with 401 mid-run, run `teams-cli login` to refresh.

Future: an `auth-check --verbose` mode that reports the expiry of every audience in the session. (Not implemented yet.)

## Update sequence (when fixing breakage)

1. **Reproduce** the failure with `health-check` and a single failing command.
2. **Re-capture** the trace (`teams-cli login --diagnostic-extra-ms 60000` + active interaction).
3. **Identify** the new URL / header / audience by diffing the trace against current code.
4. **Probe** the new shape with curl + known-audience token from `multi-tokens.json`.
5. **Patch** the relevant client file (one URL constant, one audience constant, or one new client class).
6. **Update** `docs/spike-results.md` — append a dated entry under "URL drift" listing what changed.
7. **Run** `npm test` (unit tests should pass — they mock fetch). Then run `health-check` against the live tenant.
8. **Commit** with a `fix(http/<client>): URL/audience update for ...` message.

## Telemetry to add later

Eventually we should:

- Add a launchd plist that runs `teams-cli health-check` daily and writes results to a log
- Set up a macOS notification when a probe transitions from green → red
- Optionally email a digest weekly so the user notices drift before they need the broken command
