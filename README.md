# OutlookAI — Free Open-Source Outlook AI Add-in

[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![Build](https://img.shields.io/badge/build-passing-brightgreen)]()
[![Tests](https://img.shields.io/badge/tests-546%2F546-brightgreen)]()
[![Branch](https://img.shields.io/badge/branch-feature%2Fcodex--oauth--migration-orange)]()

> ⚠️ **Still in active development.** Chat, Inbox Copilot, Inbox Reports, and
> Excel/PDF exports are live on the `feature/codex-oauth-migration` branch and
> are being merged to `master`. Track the current state in the
> [open pull requests](https://github.com/kirklandsig/OutlookAI/pulls).

The free open-source alternative to **GPT for Outlook**, **Mailbutler**,
**Lavender**, **Compose AI**, **OtterMail**, **Boomerang Respondable**,
**Mailmaestro**, **EmailTree**, **SaneBox AI**, and **Spike Magic AI** —
running inside Microsoft Outlook desktop and billed against your **own
ChatGPT subscription**. No per-seat fee, no proxy server, no data sent through
a third party.

## TL;DR

OutlookAI is a free, open-source Outlook AI add-in that replaces the paid
"AI for Outlook" market. Sign in once with your own ChatGPT Plus or Pro
account; OutlookAI runs entirely inside your Outlook process — no proxy, no
SaaS middleware, no telemetry. You get:

- **Chat** in the compose pane: WebView2 surface with streaming, multi-round
  tool dispatch, model + reasoning-effort picker.
- **Inbox Copilot** taskpane: selection-aware actions, full multi-round chat
  over the messages you highlighted.
- **Inbox Reports** taskpane: templated chips that turn your inbox into
  markdown reports, action items, vendor breakdowns.
- **Excel and PDF exports**: model-callable export tools plus a per-message
  Save as PDF button. Files land in `Documents\OutlookAI\Reports\` with
  `Open` / `Show in folder`.
- **15 Outlook tools** the model can call (11 always on + 4 admin-gated
  safe writes; no send / delete / move tools by design).
- **Voice transcription** from the same ChatGPT credential.
- Open source under MIT, fork-friendly, auditable end-to-end.

## What you get

### 💬 Chat tab (in compose pane)

WebView2-based chat that lives next to the draft you are writing. Streaming
text, multi-round tool dispatch with visible tool cards, cancellation that
preserves partial text, and light / dark / high-contrast themes. Per-turn
reasoning-effort override. Clear, copy-to-clipboard, and Save-as-PDF buttons
on every assistant message.

### 📥 Inbox Copilot (taskpane)

Open the AI Assistant taskpane on any Outlook explorer window and the same
chat surface attaches to your inbox. Highlighting one or more messages feeds
them into the context. Quick-action chips for common workflows ("summarize
this thread", "draft a reply") and full freeform chat with the entire
15-tool mailbox surface.

### 📊 Inbox Reports (taskpane)

A second taskpane focused on generating reports from your mailbox. Six
default templated chips (action items, top senders, out-of-office digest,
project status, conversations with a specific person, stats by sender). Each
chip prompts the model with a structured intent and renders the markdown
response inline. Save the report as PDF, or export underlying data as Excel.

### 📑 Excel and PDF exports

Two model-callable export tools:

- `outlook_export_excel` produces a styled `.xlsx` with bold/frozen header
  row, autofilter, and per-column formatting (text / number / currency /
  date / datetime / boolean) via ClosedXML.
- `outlook_export_pdf` renders polished markdown through an isolated
  off-screen WebView2 instance into A4 PDF with header bar and no chat-UI
  chrome.

Both save to `~\Documents\OutlookAI\Reports\` with auto-generated,
timestamped, collision-safe filenames. Tool results surface as inline file
cards with `Open` and `Show in folder` buttons. Every file action goes
through a path-policy gate that rejects any path outside the Reports
directory.

### 📎 Per-message Save as PDF

Every assistant message gets a small button that exports just that message —
the markdown the chat is showing, not the rendered HTML — to PDF. One click,
no model round-trip.

### 🛠 15 model-callable Outlook tools

**Always on (11):**

- `outlook_get_current_compose_state`
- `outlook_get_current_selection`
- `outlook_list_folders`
- `outlook_search_messages`
- `outlook_read_message`
- `outlook_read_messages` (bulk)
- `outlook_count_messages`
- `outlook_aggregate_messages` (group + top-N)
- `outlook_list_recent_threads_with`
- `outlook_export_excel`
- `outlook_export_pdf`

**Admin-gated safe writes (4):**

- `outlook_create_draft`
- `outlook_mark_as_read`
- `outlook_flag_message`
- `outlook_set_category`

By design there is no `send`, `delete`, `move-to-deleted`, or
permanent-mutation tool. The admin password gates which write tools, if any,
the model can call.

### 🎙 Voice

The mic button opens a Realtime WebSocket at
`wss://api.openai.com/v1/realtime?model=gpt-realtime-1.5` using the same OAuth
token. Transcription lands back in the prompt textbox. No separate Whisper
API key.

### ⚙️ Settings

The gear icon opens a password-gated Settings dialog with:

- ChatGPT account: Sign In / Sign Out / Refresh.
- Model picker: 7 options (`gpt-5.5`, `gpt-5.5-pro`, `gpt-5.4`,
  `gpt-5.4-mini`, `gpt-4.1-mini`, `gpt-4.1-nano`, `gpt-5.3-codex`).
- Reasoning effort dropdown, filtered per model.
- 4 checkboxes for the safe-write tools (each can be individually enabled).
- Admin password rotation.

Per-user settings persist to `%APPDATA%\OutlookAI\config.xml`.

## How it compares

| Feature | **OutlookAI** | GPT for Outlook | Mailbutler | Lavender | OtterMail | Compose AI | Boomerang Respondable | Mailmaestro | EmailTree | SaneBox AI |
|---|---|---|---|---|---|---|---|---|---|---|
| Price | **$0** (BYO ChatGPT sub) | $7-$15/user/mo | $9.95-$32.95/mo | $29-$89/mo | $10-$20/mo | $9.99-$29/mo | $4.99-$22.99/mo | $19-$39/mo | enterprise (POA) | $7-$36/mo |
| Source code | **MIT, public** | closed | closed | closed | closed | closed | closed | closed | closed | closed |
| OAuth via your own ChatGPT sub | **yes** | no | no | no | no | no | no | no | no | no |
| Tool calling on real mailbox data | **15 tools** | partial | partial | no | partial | no | no | no | yes (proprietary) | partial |
| Runs entirely in-process (no proxy) | **yes** | no | no | no | no | no | no | no | no | no |
| Inbox Copilot taskpane | **yes** | no | no | no | no | no | no | partial | partial | no |
| Inbox Reports + Excel/PDF export | **yes** | no | no | no | no | no | no | no | partial | no |
| Save any chat output as PDF | **yes** | no | no | no | no | no | no | no | no | no |
| Voice (speech-to-text) | **yes** | yes | partial | no | yes | no | no | no | no | no |
| Per-tool write permissions | **yes** | no | no | no | no | no | no | no | no | no |
| Telemetry / data leaves your machine | **none** | yes | yes | yes | yes | yes | yes | yes | yes | yes |
| Outlook desktop (VSTO) | **yes** | yes | yes | partial | yes | no | partial | partial | yes | partial |
| Auditable code | **yes** | no | no | no | no | no | no | no | no | no |

*Public list prices as of writing (May 2026), rounded for readability. The
paid alternatives charge per user per month. OutlookAI is free and bills
inference against your existing ChatGPT subscription.*

## Why OutlookAI

- **Bring your own ChatGPT subscription — no extra monthly fee.** OutlookAI
  uses the OAuth flow Codex CLI uses
  (`client_id` `app_EMoamEEZ73f0CkXaXp7hrann`), so signing in with your
  ChatGPT account is enough. There is no second OpenAI API key to manage and
  no vendor middleman charging a per-seat license.
- **Runs entirely in your Outlook process.** No proxy server sees your mail
  or your OAuth token. Every paid Outlook AI add-in we know of routes either
  your email content, your API key, or both through their own servers.
  OutlookAI talks directly to
  `chatgpt.com/backend-api/codex/responses` (text) and
  `wss://api.openai.com/v1/realtime` (voice) from inside `Outlook.exe`.
- **Real tools, real reports, real exports.** Not just "rewrite this email."
  The model can search your mailbox, summarize threads, aggregate by sender
  or day, draft replies, and export the results to Excel or PDF — all from
  inside the chat window, with explicit tool cards so you can see every
  read.
- **Open source under MIT.** Read the code, fork it, audit the security
  model, contribute. The competitors are black boxes.

## Install

For a single workstation:

```powershell
# 1. Clone the repo or download the latest Release zip
git clone https://github.com/kirklandsig/OutlookAI.git
cd OutlookAI

# 2. (Optional) Refresh the vendored WebView2 bootstrapper
.\Deploy\Fetch-WebView2Bootstrapper.ps1

# 3. Publish Release
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" `
  /p:PublishDir="C:\OutlookAI\"

# 4. Install (elevated)
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
.\Deploy\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"

# 5. Open Outlook → AI Assistant → sign in with your ChatGPT account.
```

For multi-user RDS / Terminal Server installs and IT-managed images, see
[`Deploy/README.txt`](Deploy/README.txt) (and the short
[`docs/Install.md`](docs/Install.md) summary).

## Architecture

```
Outlook.exe
  └── AITaskPane (WinForms tab control)
        ├── Chat                  → ChatController → CodexChatService → OpenAI Codex Responses API
        ├── Inbox Copilot         → InboxCopilotController → CodexChatService → tool catalog
        └── Inbox Reports         → InboxReportsController → CodexChatService → tool catalog

CodexChatService
  └── OutlookToolHost / ToolDispatcher
        └── 15 IOutlookTool implementations
              ├── LiveOutlookSurface (Outlook COM)
              └── Services/Export
                    ├── ExcelWorkbookBuilder       (ClosedXML)
                    ├── PdfRenderer                (off-screen WebView2 + PrintToPdfAsync)
                    └── PrintTemplateRenderer

ExportBridge (WebMessageReceived)
  ├── export_pdf
  ├── open_file              ← path policy
  └── reveal_in_explorer     ← path policy
```

Key components:

- `Services/CodexChatService.cs` — Codex Responses request/streaming,
  multi-round tool dispatch, parallel tool calls, cancellation.
- `Services/Tools/OutlookToolHost.cs` — tool catalog construction and
  per-tool admin-write gating.
- `Services/Tools/LiveOutlookSurface.cs` — Outlook COM surface for
  read/write/search; AdvancedSearch + iterative-folder fallback; budgeted
  / early-stop scan policy.
- `Services/Export/ExcelWorkbookBuilder.cs`,
  `Services/Export/ExcelCellCoercer.cs` — ClosedXML Excel construction.
- `Services/Export/PdfRenderer.cs`,
  `Services/Export/PrintTemplateRenderer.cs` — off-screen WebView2 PDF
  rendering with an isolated user-data folder.
- `TaskPane/Chat/ExportBridge.cs` — WebView2 host-message bridge for
  `export_pdf` / `open_file` / `reveal_in_explorer`, guarded by
  `IExportPathPolicy`.
- `WebUI/*` — `chat.js`, `markdown.js`, `styles.css`, `print-template.html`,
  `print-styles.css`. Extracted from `chat.js` so the same markdown renderer
  drives chat and PDF.

## Security model

- **OAuth via ChatGPT.** Tokens stored at:
  - `C:\ProgramData\OutlookAI\auth.json` (machine, RDS-shared by design).
  - `%APPDATA%\OutlookAI\config.xml` (per-user settings).
- **No telemetry.** Outbound traffic is exactly two endpoints from inside
  `Outlook.exe`:
  - `https://chatgpt.com/backend-api/codex/responses` (text inference).
  - `wss://api.openai.com/v1/realtime` (voice transcription).
- **Path policy on file actions.**
  `IExportPathPolicy.RequireInsideReportsDir(...)` rejects any open/reveal
  path that escapes `Documents\OutlookAI\Reports\`. Path traversal attempts
  are logged and never launched.
- **No destructive tools.** No `outlook_send_message`,
  `outlook_delete_message`, or `outlook_move_to_deleted`. The admin can
  additionally disable any safe-write tool from Settings.
- **RDS shared-credential.** On multi-user servers, `auth.json` grants
  `Authenticated Users: Modify` — explicit accepted risk. Rotation procedure
  is documented in [`Deploy/README.txt`](Deploy/README.txt).

## FAQ

### Is there a free alternative to GPT for Outlook?

Yes — OutlookAI. It is open-source under MIT, free to use, and bills inference
against your existing ChatGPT Plus or Pro subscription instead of charging a
separate per-seat fee.

### Can I use my ChatGPT Plus subscription inside Outlook?

Yes. OutlookAI uses the same OAuth client ID Codex CLI uses, so signing in
with your ChatGPT account is enough. There is no second OpenAI API key to
manage.

### How does OutlookAI compare to Mailbutler / Lavender / Compose AI?

Mailbutler bundles AI drafting with snooze, send-later, tracking, and other
email-productivity features that route data through their servers. Lavender
focuses on tone scoring and predictive suggestions during typing. Compose AI
is a browser autocomplete. OutlookAI is a native VSTO add-in with multi-round
chat, taskpane Copilot, taskpane Reports, and Excel/PDF exports — different
design point, broader capability surface, $0 vs $5-$89/mo.

### Does OutlookAI work with Outlook Web (OWA) or Outlook for Mac?

Not yet. OutlookAI is a VSTO add-in for Outlook desktop on Windows (Outlook
2016, 2019, 2021, 2024, and Microsoft 365). OWA and Outlook for Mac would
require a separate Office.js add-in; that is on the long-term roadmap.

### What model does it use?

Default `gpt-5.5` for text and `gpt-realtime-1.5` for voice. Admin can change
the model from Settings.

### Where does my data go?

Two endpoints, both directly from `Outlook.exe`. No third-party proxy, no
telemetry, no analytics. See the **Security model** section above.

### Does it work on Windows Server / RDS?

Yes — that is the primary deployment target. See
[`Deploy/README.txt`](Deploy/README.txt).

### How do I uninstall?

```powershell
.\Deploy\Uninstall-OutlookAI.ps1
```

Removes the add-in registration, the install folder, and the local OAuth
artifacts. The `Backups/` subfolder is preserved.

## Status and roadmap

**Shipped on `feature/codex-oauth-migration` (merging to `master` now):**

- Chat tab in compose pane.
- Inbox Copilot taskpane.
- Inbox Reports taskpane with six templated chips.
- Excel export tool with typed columns and per-column formatting.
- PDF export tool with off-screen WebView2 renderer.
- Per-message Save as PDF in chat and reports.
- File cards with Open / Show in folder, guarded by path policy.
- OAuth via ChatGPT (no API keys to manage).
- Voice transcription via OpenAI Realtime.
- RDS / Terminal Server install path.

**Known gaps / explicit follow-ups:**

- Multi-sheet Excel workbooks.
- PDF page numbers, footer, table of contents, embedded images.
- Settings UI for picking a custom Reports folder.
- CSV, `.docx`, and other export formats.
- Mac / OWA support (would require an Office.js add-in).

## Contributing

```powershell
# Clone and restore
git clone https://github.com/kirklandsig/OutlookAI.git
cd OutlookAI

# Build (Debug)
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"

# Test (current branch: 546/546 passing)
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Branch model:

- `master` — released code.
- `feature/<feature-name>` — feature branches with their own spec under
  `docs/superpowers/specs/` and plan under `docs/superpowers/plans/`.

Spec / plan workflow: see `docs/superpowers/` for the spec → plan →
implementation cycle that drove every phase of OutlookAI development.

## Requirements

- Windows 10 / 11, or Windows Server 2019 / 2022 / 2025.
- Microsoft Outlook 2016 / 2019 / 2021 / 2024 / Microsoft 365 (desktop).
- .NET Framework 4.7.2.
- [Visual Studio Tools for Office Runtime](https://aka.ms/VSTORuntime).
- [Microsoft Edge WebView2 Evergreen Runtime](https://developer.microsoft.com/microsoft-edge/webview2/)
  (installed by `Install-OutlookAI.ps1` if missing).

## License

MIT — see [`LICENSE`](LICENSE).
