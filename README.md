# Claude Outlook MCP Tool — Windows Edition

A Model Context Protocol (MCP) server that lets Claude interact with the
**Microsoft Outlook desktop application on Windows** through the native
**Outlook COM Object Model** (driven via PowerShell).

> This branch is the **Windows-only** build. The original macOS
> (AppleScript/JXA) implementation lives on `main`. This edition is a clean,
> modular, fully type-checked and unit-tested rewrite that uses the
> Windows-native automation surface instead of AppleScript.

## Why this design

| Concern | macOS original | Windows edition |
| --- | --- | --- |
| Automation surface | AppleScript via `run-applescript` | **PowerShell + Outlook COM** (`Outlook.Application`) |
| Data into the engine | string-interpolated into script source (`replace(/"/g, …)`) | **out-of-band JSON payload** in an env var / temp file — never interpolated |
| Data out of the engine | brittle `/\{([^}]+)\}/` regex | **structured JSON** (`ConvertTo-Json`) parsed safely |
| Structure | one 1,800-line file | modular `src/` with an injectable runner |
| Tests | none | **86 unit/integration tests** + typecheck + CI |

The result is injection-safe (a subject line like `"; Remove-Item C:\ #` is
carried as data, never executed), robust to odd output, and maintainable.

## Features

All operations target the locally-installed Outlook desktop app for the
signed-in profile.

- **`outlook_mail`** — `unread`, `read`, `search`, `get` (full message by
  EntryID), `send`, `draft`, `reply`, `forward`, `move`, `delete`,
  `mark` (read/unread), `flag`, `folders`. Supports HTML bodies, CC/BCC, and
  file attachments.
- **`outlook_calendar`** — `today`, `upcoming`, `search`, `create`,
  `update`, `delete`. Adding attendees turns a `create` into a meeting invite.
- **`outlook_contacts`** — `list`, `search`, `create`.
- **`outlook_attachments`** — `list` and `download` (save attachments from a
  message to disk).
- **`outlook_files`** — `read`, `write`, `info`, `list`, `delete` on the
  Windows filesystem. Text is UTF‑8; binary is base64. Use these to stage an
  attachment before sending, read a downloaded attachment, or save content.

Every list/read operation returns the message **EntryID**, which you pass to
`get`, `reply`, `move`, `outlook_attachments`, etc.

## Prerequisites

- **Windows 10/11** with the **classic Microsoft Outlook desktop app**
  (Microsoft 365 / Outlook 2016–2021) installed, configured, and signed in.
  See the note on "New Outlook" under [Limitations](#limitations).
- **Windows PowerShell 5.1** (built in) or **PowerShell 7+** (`pwsh`).
- [Bun](https://bun.sh/) (`powershell -c "irm bun.sh/install.ps1 | iex"`).
- [Claude Desktop](https://claude.ai/desktop).

## Installation

```powershell
git clone https://github.com/syedazharmbnr1/claude-outlook-mcp.git
cd claude-outlook-mcp
# Develop/checkout this Windows branch, then:
powershell -ExecutionPolicy Bypass -File install.ps1
```

`install.ps1` verifies Bun, checks Outlook COM, installs dependencies, runs
the test suite, and prints a ready-to-paste config snippet.

### Manual setup

```powershell
bun install
bun run check          # tsc --noEmit && bun test
```

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "C:\\Users\\YOU\\.bun\\bin\\bun.exe",
      "args": ["run", "C:\\path\\to\\claude-outlook-mcp\\index.ts"]
    }
  }
}
```

Restart Claude Desktop.

## Validate Outlook access first

Before wiring into Claude, confirm Outlook is reachable on the machine:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\smoke-test.ps1
# optionally also send yourself a test message:
powershell -ExecutionPolicy Bypass -File scripts\smoke-test.ps1 -SendTest -To you@example.com
```

It exercises the same COM calls the server uses (mail, folders, calendar,
contacts) and prints a pass/fail summary.

## Usage examples

```
Check my unread Outlook emails
Search Outlook for "quarterly report" in the Inbox
Show the full message with EntryID 00000000ABCD…
Send an HTML email to john@example.com, subject "Update", attaching C:\reports\q2.pdf
Reply-all to that message saying I'll join at 3pm
Download all attachments from that email to C:\Users\me\Downloads\q2
What's on my calendar today? / for the next 14 days?
Create a meeting "Design review" tomorrow 2–3pm with a@x.com, b@x.com
Add a contact: Jane Doe, jane@contoso.com, Acme, Director
Read C:\Users\me\Downloads\q2\summary.csv
```

## Configuration

Environment variables (all optional):

| Variable | Purpose | Default |
| --- | --- | --- |
| `OUTLOOK_MCP_PWSH` | Path/name of the PowerShell executable | `powershell.exe` |
| `OUTLOOK_MCP_TIMEOUT_MS` | Per-operation timeout | `60000` |
| `OUTLOOK_MCP_DOWNLOAD_DIR` | Default attachment download directory | `%USERPROFILE%\Downloads` |

## How it works

```
Claude ⇄ MCP (stdio) ⇄ index.ts ⇄ src/server.ts (dispatch)
                                     ├─ src/outlook/*  → build static PowerShell body
                                     │                   + JSON $payload (out-of-band)
                                     ├─ src/powershell.ts → spawn powershell.exe -File op.ps1
                                     │                   → Outlook COM → ConvertTo-Json
                                     └─ src/files.ts   → local fs (read/write/info/list/delete)
```

Each operation script is **static** (it references `$payload.subject`, never
the literal value). The payload is delivered through the `OUTLOOK_MCP_PAYLOAD`
environment variable (or a temp file when large), parsed with
`ConvertFrom-Json`, and the result is emitted as a single JSON object.

## Project layout

```
index.ts                 entry point (wires runner → stdio)
src/
  server.ts              tool registration + dispatch (exported, testable)
  tools.ts               MCP tool schemas
  powershell.ts          production runner (spawns PowerShell)
  files.ts               local filesystem operations
  ps/preamble.ts         Outlook COM helpers + JSON I/O contract
  outlook/               mail / calendar / contacts / attachments
  util/                  ps-encode, parse, dates, paths, validate, format
test/                    86 tests (bun test)
scripts/smoke-test.ps1   live Windows COM validation
```

## Testing & development

```bash
bun run typecheck   # tsc --noEmit
bun test            # 86 tests
bun run check       # both
```

The PowerShell/Outlook layer is hidden behind an injectable `OutlookRunner`
interface, so the entire dispatch flow is tested with a fake runner — no
Windows or live Outlook required for the suite. CI runs typecheck + tests on
Ubuntu and Windows, plus PSScriptAnalyzer on the `.ps1` scripts. The live COM
path is validated on a real machine with `scripts\smoke-test.ps1`.

## Limitations

- **Classic Outlook only.** The COM object model is exposed by the classic
  Outlook desktop app. The new "**Outlook for Windows**" (the web-based
  Monarch app) does **not** support COM automation; switch off *"Try the new
  Outlook"* or use classic Outlook.
- Outlook must be installed and signed in for the current Windows user.
- Calendar `Restrict` date filtering uses the US `MM/dd/yyyy hh:mm tt`
  format, which Outlook accepts broadly; exotic locale configurations may
  need adjustment.
- This server controls the local desktop client — it is not a Microsoft Graph
  / cloud integration.

## Troubleshooting

- **"requires Windows … Detected platform: …"** — you're running on a
  non-Windows host. This build only works on Windows.
- **Cannot start Outlook via COM** — open Outlook once and complete profile
  setup; ensure classic (not "new") Outlook; try `OUTLOOK_MCP_PWSH=pwsh.exe`.
- **Attachment not attached** — pass an absolute Windows path that the user
  running the server can read (e.g. `C:\Users\me\file.pdf`).
- **Timeouts on large mailboxes** — raise `OUTLOOK_MCP_TIMEOUT_MS`.

## License

MIT
