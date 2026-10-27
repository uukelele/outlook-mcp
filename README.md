# outlook-mcp

An [MCP](https://modelcontextprotocol.io) server for Microsoft 365 Outlook — mail, calendar and contacts — that talks to the same internal JSON API the Outlook web app (OWA) uses: `https://outlook.cloud.microsoft/owa/service.svc`.

It does not use Microsoft Graph, EWS SOAP or a browser. Every tool call is a plain HTTPS request with a bearer token, so it works anywhere Node runs and needs no app registration or admin consent beyond what Outlook on the web already has.

- 40 tools: read, search, send, reply, forward, drafts, attachments, folders, flags, categories, calendar events and invitations, contacts, directory lookup
- Authentication through the Azure CLI by default (tokens are refreshed automatically), or any token source you provide
- Read-only mode for safe exploration
- Compact, normalized JSON output (UTC times, SMTP addresses, text bodies by default)

> **Caveat:** the OWA service API is undocumented and owned by Microsoft. It has been stable for years and is what several open-source clients use, but it can change without notice. See [docs/OWA-API.md](docs/OWA-API.md) for exactly what was verified and how.

## Requirements

- Node.js 20 or newer — Windows, macOS and Linux are all supported
- A Microsoft 365 (work or school) mailbox
- One of:
  - the [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli), signed in with `az login` (recommended), or
  - some other way to obtain an access token for `https://outlook.office.com`

## Install

### From GitHub (with the GitHub CLI)

```sh
gh repo clone uukelele/outlook-mcp
cd outlook-mcp
npm install
npm run build
az login --allow-no-subscriptions   # work/school accounts often have no Azure subscription
npm run check                       # verifies auth and prints your mailbox identity
```

`gh` itself: `winget install GitHub.cli` on Windows, `brew install gh` on macOS, or your Linux package manager — then `gh auth login` once.

### From a local copy

```sh
npm install
npm run build
az login            # if not already signed in
npm run check
```

`npm run check` prints something like:

```json
{ "ok": true, "auth": "azure-cli", "mailbox": "you@contoso.com", "displayName": "You", "timeZone": "GMT Standard Time", "inbox": { "total": 2535, "unread": 119 }, "readOnly": false }
```

### Platform notes

The server has no native dependencies and runs anywhere Node 20+ runs, including Windows ARM64. A few platform details are handled for you:

- **Windows:** the Azure CLI is `az.cmd`; the server invokes it through the shell automatically. `download_attachment` accepts `C:\...` paths, and `OUTLOOK_TOKEN_COMMAND` runs through `cmd.exe`.
- **Time zones:** user-facing zones are IANA names (e.g. `Europe/London`), resolved via Node's bundled ICU, so `Intl` works the same on every platform. The mailbox's own zone stays a Windows zone ID, as Exchange expects.
- **Line endings:** the repo pins LF via `.gitattributes`, and the test scripts use `cross-env`, so `npm test` / `npm run test:live` work in cmd and PowerShell as well as POSIX shells.
- **macOS:** install the Azure CLI with `brew install azure-cli` if you don't have it.

### Add it to an MCP client

**OpenCode** (`~/.config/opencode/opencode.jsonc` on macOS/Linux, `%USERPROFILE%\.config\opencode\opencode.jsonc` on Windows; v2 config):

```jsonc
{
  "mcp": {
    "servers": {
      "outlook": {
        "type": "local",
        "command": ["node", "/path/to/outlook-mcp/dist/index.js"],
        "disabled": false
      }
    }
  }
}
```

Windows path form: `"command": ["node", "C:\\path\\to\\outlook-mcp\\dist\\index.js"]`.

**Claude Desktop / other clients** (`mcpServers` format; the config file is `claude_desktop_config.json` — see your client's docs):

```json
{
  "mcpServers": {
    "outlook": { "command": "node", "args": ["/path/to/outlook-mcp/dist/index.js"] }
  }
}
```

Add `"environment"` / `"env"` entries for any settings below, e.g. `"OUTLOOK_READ_ONLY": "1"`.

## Authentication

The server needs a bearer token whose audience is `https://outlook.office.com`. **Microsoft Graph tokens do not work** (the server detects them and says so). Sources, in the order they are tried when `OUTLOOK_AUTH` is unset:

| Source | Setting | Refresh |
| --- | --- | --- |
| Static token | `OUTLOOK_ACCESS_TOKEN` | No — restart with a new token when it expires (~1 h) |
| Token file | `OUTLOOK_TOKEN_FILE` (raw token or JSON with `accessToken`/`access_token`) | Re-read when expired or rejected |
| Command | `OUTLOOK_TOKEN_COMMAND` (prints a token or JSON) | Re-run when expired or rejected |
| OWA cookie *(experimental, untested)* | `OUTLOOK_COOKIE` (+ `OUTLOOK_CANARY`) | No |
| **Azure CLI** (default) | `OUTLOOK_AZ_TENANT` optional | Automatic |

The Azure CLI source runs `az account get-access-token --resource https://outlook.office.com`. Tokens are cached in memory until five minutes before expiry and refreshed once automatically if the service rejects one. Tokens are never logged or written to disk.

## Configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| `OUTLOOK_AUTH` | `auto` | Force a source: `azure-cli`, `token`, `token-file`, `command`, `cookie` |
| `OUTLOOK_READ_ONLY` | off | `1` registers only the 15 read-only tools |
| `OUTLOOK_TIMEZONE` | mailbox setting | Windows time zone ID (e.g. `GMT Standard Time`) for times given without an offset |
| `OUTLOOK_MAX_DOWNLOAD_MB` | `25` | Maximum attachment size for downloads and uploads |
| `OUTLOOK_REQUEST_TIMEOUT_MS` | `30000` | Per-request timeout |
| `OUTLOOK_ORIGIN` | `https://outlook.cloud.microsoft` | OWA origin |
| `OUTLOOK_ATTACHMENTS_ORIGIN` | `https://attachments.office.net` | Attachment download host |
| `OUTLOOK_ANCHOR_MAILBOX` | from token (`PUID:…@tenant`) | Override the `X-AnchorMailbox` routing header |
| `OUTLOOK_DEBUG` | off | `1` logs each request (action and status only) to stderr |

## Tools

IDs returned by one tool can be passed to any other. Folders can be given as a well-known name (`inbox`, `sent`, `drafts`, `trash`, `junk`, `archive`, …), a display name, a path like `Inbox/Projects`, or an ID.

**Mailbox** — `get_mailbox_info`, `list_categories`

**Folders** — `list_folders`, `create_folder`, `rename_folder`, `move_folder`, `delete_folder`

**Reading mail**

| Tool | Notes |
| --- | --- |
| `list_messages` | Folder listing, newest first; `filter` = `unread` / `flagged` / `toOrCcMe`; paging via `offset`/`nextOffset` |
| `search_messages` | Text, sender, recipient, subject, date range, category, unread/flagged/attachments. See *Search* below |
| `get_message` | Full message as text (default) or sanitized HTML, recipients, attachment list. Does not mark read unless asked |
| `get_thread` | Whole conversation, oldest first, each message showing only its new text |

**Writing mail** — `send_message`, `create_draft`, `update_draft`, `send_draft`, `reply_message`, `forward_message`

`reply_message` and `forward_message` send immediately unless `send: false`, which saves a draft and returns its ID. The original message is quoted automatically.

**Triage** (batched, with per-item results) — `mark_messages_read`, `flag_messages`, `categorize_messages`, `move_messages`, `copy_messages`, `delete_messages`

**Attachments** — `list_attachments`, `download_attachment`, `add_attachment`, `remove_attachment`

`download_attachment` returns text files as text, images as images and other files as an embedded resource; pass `savePath` to write the file to disk instead. Attachments can be added from a local `path`, base64 or plain text.

**Calendar** — `list_calendars`, `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `respond_to_event`

**Contacts & people** — `list_contacts`, `get_contact`, `create_contact`, `update_contact`, `delete_contact`, `search_people` (organization directory and personal contacts)

### Behavior worth knowing

- **Times.** Results are UTC. Inputs with an offset (`2026-11-02T14:00:00Z`, `+01:00`) are exact. Inputs without one are local time — in the IANA `timeZone` argument if given, otherwise in the mailbox's time zone, resolved by Exchange itself (daylight saving included).
- **Search** scans the newest messages of each folder (`scanLimit`, default 1000 per folder) and matches subject, sender, recipient display names and the ~250-character preview. It does not search full bodies or attachment contents. Server-side search was investigated and is not reachable with these tokens — see [docs/OWA-API.md](docs/OWA-API.md#search). Results say when a folder was only partly scanned; add `after` to search further back cheaply.
- **Meeting invitations.** Creating an event with attendees always emails invitations — Exchange ignores "don't send" for new meetings. `update_event` and `delete_event` notify attendees unless `notifyAttendees: false`.
- **Sending** uses Exchange's normal submission path and keeps a copy in Sent Items, exactly like Outlook on the web.
- **Safety.** Tools carry MCP annotations (`readOnlyHint`, `destructiveHint`, `openWorldHint`) so clients can ask for confirmation. The server instructions tell the model to confirm before contacting other people and to prefer drafts. `OUTLOOK_READ_ONLY=1` removes every tool that changes anything.

## Development

```sh
npm run build        # compile to dist/
npm test             # 73 offline tests (fake OWA endpoint, no network)
npm run test:live    # 9 live tests against your mailbox — read-only by construction
```

The offline tests check payloads against the request shapes that were verified on the live service, cover the client's retry/auth/error handling, and run the full MCP protocol over an in-memory transport. The live suite starts the server in read-only mode, so write tools do not exist during the run.

```
src/
  index.ts            CLI entry (stdio server, --check, --help)
  server.ts           wires config → auth → client → services → tools
  config.ts           environment settings
  auth/               token sources (Azure CLI, command, file, static, cookie) and JWT helpers
  owa/client.ts       HTTP client: headers, envelope, refresh-on-401, throttling, error mapping
  owa/builders.ts     request payload builders (__type ordering, folder refs, recipients)
  owa/normalize.ts    raw OWA objects → compact JSON
  owa/time.ts         date/time parsing and time zone handling
  services/           mail, folders, attachments, calendar, contacts, mailbox
  tools/              MCP tool definitions
docs/OWA-API.md       protocol notes and verification record
```

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| `Azure CLI could not issue an Outlook token` | Run `az login` — with a work/school account that has no Azure subscription, use `az login --allow-no-subscriptions` |
| `az` not found on Windows | Install with `winget install Microsoft.AzureCLI` (or the [MSI](https://aka.ms/installazurecliwindows)), then reopen the terminal so `az.cmd` is on PATH |
| `token is for Microsoft Graph` | Request the token with `--resource https://outlook.office.com` |
| `Outlook rejected the credentials … (HTTP 401)` | Token expired or lacks mailbox access; with a static token, supply a fresh one |
| `returned an HTML page instead of JSON` | Not authenticated (often cookie mode); use a bearer token |
| `… (the service rejected the request shape)` | A payload the service does not accept; please report it with `OUTLOOK_DEBUG=1` output |

## License

MIT
