<p align="center"><img src="packaging/icon.png" width="96" alt=""></p>

# post.

**Your mailbox for your AI. No add-in, no Microsoft 365 subscription, no service in between.**

post. is an MCP server that connects Claude Desktop, Claude Code, Codex or a local model to your own mailbox. The AI can read mail, search it, read attachments, summarise threads and draft replies or forwards. Nothing is sent unless you explicitly allow it.

```
Claude Desktop, Claude Code, Codex, local models
                │  MCP (stdio, on this computer only)
              post.
                │  Outlook object model (COM)  or  IMAP/SMTP
           your mailbox
```

## Why this way round?

Plenty of people want to use their AI subscription for their own tools. Anthropic does not allow the login of a Claude subscription to be used inside other programs. post. turns it around: **post. does not call the AI, your AI app calls post.** Your subscription stays in the official app, post. only supplies the tools. That is within the rules and works with any app that speaks MCP.

## What happens to my mail?

post. itself runs on your computer and sends nothing to a server of its own. But whatever your AI reads from your mail, the AI app sends to its provider (Anthropic, OpenAI …), exactly as if you had pasted the text into the chat. With a local model everything stays in house.

## Installation

### Windows (recommended)

1. Download `post.exe` from this repository's releases page and double-click it.
   For new programs Windows often shows "Windows protected your PC": click **More info** and then **Run anyway**.
2. Setup opens in your browser. post. finds your Outlook, Thunderbird and your AI apps by itself.
3. Pick your mailboxes, set the permissions, pick your AI app. post. copies itself to `%LOCALAPPDATA%\Programs\post\`; you can delete the file you downloaded afterwards.
4. Quit the AI app completely and open it again. For Claude Desktop that means: right-click the Claude icon in the notification area at the bottom right, choose **Quit**.
5. Try it: "Which mailboxes can you see?"

Alternatively, for Claude Desktop: double-click `post-<version>-windows.mcpb` (a desktop extension). With classic Outlook you do not have to fill in anything there. With "Allow sorting and deleting" the AI may move, delete and mark mail.

### macOS and Linux

```bash
pipx install gg-post        # or: uv tool install gg-post
post setup
```

There is an extension for Claude Desktop here too: double-click `post-<version>-uv.mcpb`. It carries the source with it and sets itself up on first start; for that, [uv](https://docs.astral.sh/uv/) has to be installed (`brew install uv` or `curl -LsSf https://astral.sh/uv/install.sh | sh`). The first start then takes a few seconds longer, because uv fetches the dependencies.

## What does post. detect by itself?

| Program | How post. uses it | Password needed? |
|---|---|---|
| Outlook (classic), Office 2016 to 2024 or Microsoft 365 Apps | Drives Outlook directly. Every account, shared mailbox and PST archive of the default profile | No |
| Outlook (new) | Cannot be driven from outside. post. connects to the same mailbox over IMAP | Yes |
| Thunderbird | Takes host, port and user name from your profile | Yes |
| Apple Mail and everything else | Enter your address, post. looks up the server settings itself | Yes |

post. picks the route by itself, and it picks the one that carries without you lifting a finger: first an address you entered explicitly, then classic Outlook (which needs no password), then Thunderbird with the server settings from its profile. If Thunderbird is your default mail program, it comes first. The other mailboxes it found stay reachable as further accounts. If post. finds nothing that carries by itself, it says in plain words what is missing instead of sitting there empty (`post doctor` shows the chosen route).

post. looks up the server settings the way Thunderbird does: the provider's autoconfig, the Mozilla ISPDB, the MX record (your own domain at IONOS, Strato, Google …), DNS SRV records, a built-in list, and guessing last. For Gmail, GMX, WEB.DE, Telekom, iCloud and Yahoo, setup tells you directly what you have to enable there (app password, IMAP access).

## Tools for the AI

| Tool | What it does |
|---|---|
| `mail_accounts` | Connected mailboxes and the allowed actions |
| `mail_folders` | Folders with total and unread counts |
| `mail_search` | Lists or searches mail, across all mailboxes and, on request, all folders (`folder="all"`). Filters: words, sender, time range (`today`, `yesterday`, `7d`, a date), unread |
| `mail_read` | A whole mail with recipients and attachment names, on request with the full thread (quotes removed) |
| `mail_attachment` | Text from attachments: PDF, Word, Excel, text, CSV, HTML, .eml |
| `mail_draft` | A new draft, reply or forward with attachments. In Outlook with your signature, on request straight as an open window. A revision replaces the earlier draft |
| `mail_discard_draft` | Deletes a draft (drafts only) |
| `mail_move`, `mail_mark` | Sort and mark, only when enabled |
| `mail_delete` | Deletes a mail into the trash, or for good when you explicitly ask. Only when sorting is enabled |
| `mail_send` | Only visible once you have enabled sending |
| `post_setup` | Opens setup in the browser |

Examples:

- "Summarise my unread mail from today."
- "What did the last customer want? Draft a reply and open it in Outlook."
- "Read the order in the PDF attachment and check whether quantity and date match our quote."
- "Forward the invoice from supplier X to accounting@company.com."

## Permissions

| Permission | Default | Meaning |
|---|---|---|
| Read and search | always on | Nothing is marked as read |
| Create drafts | always on | Lands in "Drafts" or opens in Outlook; you send it yourself |
| Sort, delete and mark | off | Sorting into folders, deleting to the trash, marking as spam |
| Send on its own | off | The AI is instructed to ask first. post. cannot enforce that |

Deleting always goes to the trash first, so you can get the mail back. Erasing for good is a separate, explicit request. post. cannot create mail rules or server-side filters: IMAP has no concept of them, and Gmail filters can only be set in Gmail itself.

## Security

- **Mail is foreign content.** Every mail body and every attachment is fenced for the AI between markers carrying a random per-start id; fence look-alikes and invisible control characters are stripped, in subjects and previews too. The AI is instructed never to act on instructions found in mail without asking you first.
- **Passwords** live only in the Windows Credential Manager, the macOS Keychain or the Linux Secret Service. None of them is in the configuration file.
- **Audit log:** `audit.log` in the settings folder records which tool ran when. Metadata only, never contents.
- **Setup in the browser** runs on `127.0.0.1` only, with a random port and a one-time token, guarded against DNS rebinding and other web pages.
- When registering with AI apps, post. makes a backup of each settings file first and leaves other entries untouched.
- **Outlook security prompt:** on some company PCs Outlook asks "A program is trying to access e-mail address information". Allow it. To settle it permanently: File > Options > Trust Center > Trust Center Settings > Programmatic Access (needs administrator rights).

## Commands

```
post               Double-click: setup. Started by an AI app: the server.
post setup         Setup in the browser
post doctor        Check the connections
post detect        Detected mail programs and AI apps as JSON
post register      Register with AI apps (claude_desktop, claude_code, codex)
post uninstall     Remove from every AI app (--all: settings and passwords too)
post serve         MCP server over stdio
```

Settings live in `%APPDATA%\post\config.json` (Windows), `~/Library/Application Support/post/` (macOS) or `~/.config/post/` (Linux). For special cases: `POST_CONFIG_DIR`, `POST_EMAIL`, `POST_PASSWORD`, `POST_ALLOW_SEND`, `POST_ALLOW_MODIFY` (the last two can only switch permissions on once a setup has been saved).

### For local models and other MCP clients

```json
{ "mcpServers": { "post": { "command": "C:\\Users\\YOU\\AppData\\Local\\Programs\\post\\post.exe", "args": ["serve"] } } }
```

## Limits

- Outlook.com, Hotmail and Microsoft 365 no longer allow IMAP with a password. Those accounts go through classic Outlook or through Thunderbird.
- With classic Outlook, post. uses the default profile.
- Scanned PDFs without text and the old Office formats (.doc, .xls) cannot be read.
- Drafts cannot attach files of their own yet (forwarding with the original attachments does work).
- Mail rules and server-side filters are not possible over IMAP.
- Calendar and contacts are not in yet.

## Development

```bash
pip install -e ".[dev]"
pytest                                   # a real IMAP test server (pymap) and a rebuilt Outlook
python scripts/demo_setup.py             # setup with sample data in the browser
python scripts/demo_setup.py --shots out # screenshots
pyinstaller packaging/pyinstaller/post.spec --noconfirm   # post.exe
python scripts/build_mcpb.py windows     # the Claude Desktop extension
```

The IMAP tests need pymap, which needs POSIX signal handlers, so they run on Linux and macOS and are skipped on Windows. `tests/test_stdio_smoke.py` needs no mailbox and therefore also checks the built `post.exe` (set `POST_BINARY`).

The Windows build and the release files are produced automatically by GitHub Actions on every `v*` tag.

```
src/ggpost/
  backends/outlook_com.py   Outlook over COM, in its own STA thread
  backends/imap.py          IMAP/SMTP
  detect/clients.py         Outlook (registry), Thunderbird (prefs.js), default mail program
  detect/autoconfig.py      finding the server settings for an address
  detect/aiapps.py          finding and registering with Claude Desktop, Claude Code, Codex
  extract.py                text from PDF, Word, Excel
  setup/                    setup in the browser
  server.py                 the MCP tools
```

Some German stays on purpose: the folder names post. matches (`Posteingang`, `Gelöschte Elemente`), the openers of a quoted history (`Von:`, `Am … schrieb`), the subject prefixes (`AW:`, `WG:`) and the date words `heute` and `gestern`. They are how German mailboxes and German mail programs actually write things, so they are matching rules, not interface text.

## Licence

AGPL-3.0. A tool by gedacht.gemacht.
Fonts: DM Sans and JetBrains Mono, both under the SIL Open Font License.
