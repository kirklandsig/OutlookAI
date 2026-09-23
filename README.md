# OutlookAI — Free Open-Source Outlook AI Add-in

[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![Release](https://img.shields.io/github/v/release/kirklandsig/OutlookAI)](https://github.com/kirklandsig/OutlookAI/releases/latest)
[![Tests](https://img.shields.io/badge/tests-898%2F898-brightgreen)](#contributing)

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
- **16 Outlook tools** the model can call (12 always on + 4 admin-gated
  safe writes; no send / delete / move tools by design).
- **Voice transcription** from the same ChatGPT credential.
- **In-app updates and a live model list**: admins install new releases and
  refresh the models ChatGPT offers from Settings, no reinstall needed.
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
16-tool mailbox surface.

### 📊 Inbox Reports (taskpane)

A second taskpane focused on generating reports from your mailbox. Six
default templated chips (action items, top senders, out-of-office digest,
project status, conversations with a specific person, stats by sender). Each
chip prompts the model with a structured intent and renders the markdown
response inline. Save the report as PDF, or export underlying data as Excel.

### 📑 Excel and PDF exports

Three model-callable export tools:

- `outlook_export_excel` produces a styled `.xlsx` with bold/frozen header
  row, autofilter, and per-column formatting (text / number / currency /
  date / datetime / boolean) via ClosedXML.
- `outlook_export_pdf` renders polished markdown through an isolated
  off-screen WebView2 instance into A4 PDF with header bar and no chat-UI
  chrome.
- `outlook_export_search_results` exports the complete list of messages
  matching a search to Excel, not just the first page of results: it counts
  the true total, collects up to a server-set ceiling (2,000 rows by default,
  10,000 at most), and reports how many it exported out of how many matched.

All three save to `~\Documents\OutlookAI\Reports\` (or
`%LOCALAPPDATA%\OutlookAI\Reports\` when Documents is redirected to a network
share, as on many RDS servers) with auto-generated, timestamped,
collision-safe filenames. Tool results surface as inline file
cards with `Open` and `Show in folder` buttons. Every file action goes
through a path-policy gate that rejects any path outside the Reports
directory.

### 📎 Per-message Save as PDF

Every assistant message gets a small button that exports just that message —
the markdown the chat is showing, not the rendered HTML — to PDF. One click,
no model round-trip.

### 🛠 16 model-callable Outlook tools

**Always on (12):**

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
- `outlook_export_search_results` (complete list to Excel)

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

The gear icon opens a password-gated Settings dialog. Its choices apply to
every user on the machine (saved in `C:\ProgramData\OutlookAI\config.xml`):

- ChatGPT account: Sign In / Sign Out / Refresh.
- Model picker driven by the model catalog ChatGPT publishes for your
  account. **Update Models** fetches the current list and each model's
  supported reasoning efforts, so new OpenAI models (and new effort levels
  such as `Max`) show up without an OutlookAI release. The list is cached in
  `C:\ProgramData\OutlookAI\models.json` and shared with every user on the
  machine; retirement dates from the catalog are shown next to the model, and
  a retired model automatically falls back to its announced replacement.
- Reasoning effort dropdown, filtered per model. `Auto` (called `None` before
  v2.2.1) sends no effort, so the model's own default applies (medium today).
  `config.xml` still stores it as `None`, which older versions read the same way.
- 4 checkboxes for the safe-write tools (each can be individually enabled).
- Admin password rotation.
- Updates: **Check Now** / **Install Update** for new OutlookAI releases (see
  [Updating](#updating)).

## How it compares

| Feature | **OutlookAI** | GPT for Outlook | Mailbutler | Lavender | OtterMail | Compose AI | Boomerang Respondable | Mailmaestro | EmailTree | SaneBox AI |
|---|---|---|---|---|---|---|---|---|---|---|
| Price | **$0** (BYO ChatGPT sub) | $7-$15/user/mo | $9.95-$32.95/mo | $29-$89/mo | $10-$20/mo | $9.99-$29/mo | $4.99-$22.99/mo | $19-$39/mo | enterprise (POA) | $7-$36/mo |
| Source code | **MIT, public** | closed | closed | closed | closed | closed | closed | closed | closed | closed |
| OAuth via your own ChatGPT sub | **yes** | no | no | no | no | no | no | no | no | no |
| Tool calling on real mailbox data | **16 tools** | partial | partial | no | partial | no | no | no | yes (proprietary) | partial |
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

Each [release](https://github.com/kirklandsig/OutlookAI/releases/latest)
has an install bundle, `OutlookAI-vX.Y.Z-RDS-Deploy.zip`, and its `.sha256`.
Download both, then in an elevated PowerShell in the download folder:

```powershell
# 1. Check the download (must print True) and extract it
$zip = ".\OutlookAI-vX.Y.Z-RDS-Deploy.zip"
(Get-FileHash $zip -Algorithm SHA256).Hash -eq (Get-Content "$zip.sha256").Trim()
Unblock-File $zip
Expand-Archive $zip -DestinationPath C:\OutlookAI

# 2. Install
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
C:\OutlookAI\Install-OutlookAI.ps1 -SourcePath C:\OutlookAI

# 3. Open Outlook → AI Assistant → sign in with your ChatGPT account.
```

The same bundle covers single workstations, multi-user RDS / Terminal Server
and IT-managed images; see [`Deploy/README.txt`](Deploy/README.txt), which
is also inside the zip. To build the bundle from source instead, see
[Contributing](#contributing).

### Updating

Admins update from inside Outlook: Settings → **Updates** → **Check Now**,
then **Install Update** when a newer release is out. It downloads the
release, checks its SHA256 and runs the installer with administrator rights.
The installer closes Outlook for every user on the machine and leaves it
closed, so warn RDS users first; everyone reopens Outlook afterwards.
Updates keep the settings.

## Architecture

```
Outlook.exe
  └── AITaskPane (WinForms tab control)
        ├── Chat                  → ChatController → CodexChatService → OpenAI Codex Responses API
        ├── Inbox Copilot         → InboxCopilotController → CodexChatService → tool catalog
        └── Inbox Reports         → InboxReportsController → CodexChatService → tool catalog

CodexChatService
  └── OutlookToolHost / ToolDispatcher
        └── 16 IOutlookTool implementations
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
- `Services/OutlookToolHost.cs` — tool catalog construction and
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
- `Services/Models/` — the model catalog behind **Update Models**: fetch,
  cache (`models.json`), per-model reasoning efforts, retirement routing.
- `Services/Updates/` — the in-app updater: GitHub Releases lookup, download
  and SHA256 check, elevated install, update history.
- `Config.cs`, `SettingsForm.cs` — the config layers and the Settings dialog;
  saves are merged into the ProgramData `config.xml` under a cross-process
  lock (`Services/FileLock.cs`).

## Security model

- **OAuth via ChatGPT.** Tokens stored at
  `C:\ProgramData\OutlookAI\auth.json` (machine, RDS-shared by design).
- **Settings** (model, effort, write tools and the admin password, in plain
  text) live in `C:\ProgramData\OutlookAI\config.xml`, which every signed-in
  user on the machine can read and change, like `auth.json`.
- **No telemetry.** In normal use, outbound traffic is two endpoints from
  inside `Outlook.exe`:
  - `https://chatgpt.com/backend-api/codex/responses` (text inference).
  - `wss://api.openai.com/v1/realtime` (voice transcription).

  Admin actions in Settings add, only when clicked:
  - **Check Now / Install Update:** `api.github.com` and the release download
    for `kirklandsig/OutlookAI`.
  - **Update Models:** `https://chatgpt.com/backend-api/codex/models` (the
    model catalog), one rejected `/responses` call that reports the server's
    reasoning-effort values, and `api.github.com` for the latest `openai/codex`
    release number. Only the ChatGPT calls carry the sign-in token.
- **Path policy on file actions.**
  `IExportPathPolicy.RequireInsideReportsDir(...)` rejects any open/reveal
  path that escapes the Reports folder (`Documents\OutlookAI\Reports\`, or
  `%LOCALAPPDATA%\OutlookAI\Reports\` when Documents is on a network share).
  Path traversal attempts are logged and never launched.
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

For text, whatever the admin picks in Settings from the models ChatGPT offers
your account (the default is the catalog's top model; `gpt-6-astra` as of
September 2026). Voice uses `gpt-realtime-1.5`. Settings → **Update Models**
refreshes the list.

### Where does my data go?

Two endpoints in normal use, both directly from `Outlook.exe`, plus the
update and model-list checks an admin runs from Settings. No third-party
proxy, no telemetry, no analytics. See the **Security model** section above.

### Does it work on Windows Server / RDS?

Yes — that is the primary deployment target. See
[`Deploy/README.txt`](Deploy/README.txt).

### How do I uninstall?

In an elevated PowerShell, run the uninstaller from the install bundle (or
from `Deploy\` in a clone):

```powershell
C:\OutlookAI\Uninstall-OutlookAI.ps1
```

It removes the add-in registration, the install folder and the shared
ChatGPT sign-in (`C:\ProgramData\OutlookAI\auth.json`). It keeps the rest of
`C:\ProgramData\OutlookAI` — the Settings in `config.xml` (including the
admin password), `models.json` and `Backups\` — so a reinstall picks them
up; delete that folder too to remove everything.

## Status and roadmap

**Shipped** (details on the
[Releases](https://github.com/kirklandsig/OutlookAI/releases) page):

- **v2 (ChatGPT OAuth):** Chat tab in the compose pane, Inbox Copilot and
  Inbox Reports taskpanes, Excel and PDF export tools, per-message Save as
  PDF, file cards guarded by the path policy, voice transcription, and the
  RDS / Terminal Server install path. No API keys to manage.
- **v2.1.0:** in-app updater (Settings → Updates).
- **v2.1.1:** exports work when Documents is redirected to a network share;
  admin Settings are shared through `C:\ProgramData\OutlookAI\config.xml`.
- **v2.1.2 / v2.1.3:** `outlook_export_search_results` for complete-list
  Excel exports (with a `folder` column); searches report the true total
  when they truncate.
- **v2.2.0:** Update Models — the model list and reasoning efforts come from
  ChatGPT's catalog, and retired models route to their replacement.
- **v2.2.1:** updates keep `config.xml` and per-user settings; the `None`
  effort is now called `Auto`.
- **v2.2.2:** Settings saves for every user, and the ProgramData file wins
  over old per-user copies.

**Known gaps / explicit follow-ups:**

- Multi-sheet Excel workbooks.
- PDF page numbers, footer, table of contents, embedded images.
- Settings UI for picking a custom Reports folder.
- CSV, `.docx`, and other export formats.
- Automatic model-list refresh (Update Models is a manual button).
- Mac / OWA support (would require an Office.js add-in).

## Contributing

```powershell
# Clone
git clone https://github.com/kirklandsig/OutlookAI.git
cd OutlookAI

# One-time: a certificate to sign the add-in's manifests. The maintainer's
# isn't in the repo, and installs trust the add-in by location, so any will do.
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=OutlookAI dev" `
  -CertStoreLocation Cert:\CurrentUser\My

# Restore and build (Debug)
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /restore /p:RestorePackagesConfig=true /p:Configuration=Debug `
  /p:Platform="Any CPU" /p:ManifestCertificateThumbprint=$($cert.Thumbprint)

# Test (898 tests, all passing)
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"

# Build an install bundle (the same zip a release has) into out\
.\Deploy\Make-ReleaseZip.ps1 -Tag v0.0.0-dev -OutDir out -CertThumbprint $cert.Thumbprint
```

Building needs Visual Studio with the Office/SharePoint development
workload. To install your build, run
`out\staging-v0.0.0-dev\Install-OutlookAI.ps1 -SourcePath out\staging-v0.0.0-dev`
elevated. In a later session, find the certificate again with
`Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert`.

Branch model:

- `master` — released code; each release is tagged `vX.Y.Z`.
- `feature/<name>` and `fix/<name>` — work branches, merged with `--no-ff`.
  Larger features get a spec under `docs/superpowers/specs/` and a plan under
  `docs/superpowers/plans/`.

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
