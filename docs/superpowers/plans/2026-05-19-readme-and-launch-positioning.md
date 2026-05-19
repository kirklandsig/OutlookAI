# README and Launch Positioning Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the Phase 2-era `README.md` and `Deploy/README.txt` with a refreshed product landing page and three-shape deployment guide that match what is actually shipped on `feature/codex-oauth-migration`, then merge that branch to `master`.

**Architecture:** Docs-only pass. Three files are rewritten or created: `README.md`, `Deploy/README.txt`, and `docs/Install.md`. No production code is touched. Verification gates rely on the existing build/test suite and the open PR (`#1`).

**Tech Stack:** Markdown (README + docs), plain text (`Deploy/README.txt`). Verification uses the same VS MSBuild + VSTest + `node --check` toolchain documented in `docs/superpowers/plans/2026-05-18-phase-5-exports.md`.

**Spec:** `docs/superpowers/specs/2026-05-19-readme-and-launch-positioning-design.md`

---

## File Structure

**Modified files:**

- `README.md` — full rewrite. Landing-page structure with hero, TL;DR, "What you get", competitor comparison, "Why OutlookAI", install snippet, architecture, security, FAQ, status/roadmap, contributing, requirements, license.
- `Deploy/README.txt` — full rewrite around three install shapes (single workstation, RDS/Terminal Server, IT-managed image) with shared prerequisites, troubleshooting, uninstall, and rollback.

**New files:**

- `docs/Install.md` — short pointer doc that summarizes the three install shapes and links into `Deploy/README.txt` as the canonical source.

**Verification commands** (use throughout this plan, run from
`C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-codex-oauth-migration`):

```powershell
node --check VSTO2\OutlookAI\WebUI\chat.js
node --check VSTO2\OutlookAI\WebUI\markdown.js
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Baseline before starting: **546 tests passing**. Target at end: **546 tests passing** (docs-only; no test count change).

---

## Task 1: Rewrite `README.md`

Replace the existing `README.md` in full. The new content uses the structure agreed in the spec.

**Files:**

- Modify: `README.md`

- [ ] **Step 1: Read current `README.md`**

```powershell
Get-Content -LiteralPath "README.md" | Select-Object -First 30
```

This is a sanity check so you know what you are replacing.

- [ ] **Step 2: Replace `README.md` with the new content**

Use `Write`/`Set-Content -LiteralPath "README.md"` with the following exact content (no surrounding fence). Use UTF-8.

````markdown
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
````

- [ ] **Step 3: Sanity-check the file**

```powershell
Get-Content -LiteralPath "README.md" -TotalCount 6
(Get-Content -LiteralPath "README.md").Length
Select-String -Path "README.md" -Pattern "546/546|15 Outlook tools|Inbox Reports|outlook_export" | Select-Object -First 8
```

Expected: header line `# OutlookAI — Free Open-Source Outlook AI Add-in`,
roughly 280–340 lines, and the grep finds matches for `546/546`,
`15 Outlook tools`, `Inbox Reports`, and `outlook_export`.

- [ ] **Step 4: Verify no broken references**

```powershell
Select-String -Path "README.md" -Pattern "Phase 2|94/94|10 first-class|gpt-realtime-1.5\?" | Where-Object { $_.Line -notmatch "still wired|gpt-realtime-1.5" }
```

Expected: no matches (we removed the Phase 2 era language). If anything
returns, fix the offending line and re-run.

- [ ] **Step 5: Commit**

```powershell
git add README.md
git commit -m "docs(readme): rewrite as product landing page for v2 launch"
```

---

## Task 2: Rewrite `Deploy/README.txt`

Replace `Deploy/README.txt` with the three-shape deployment guide.

**Files:**

- Modify: `Deploy/README.txt`

- [ ] **Step 1: Replace `Deploy/README.txt` with the new content**

Write the file (UTF-8, no BOM) with the following exact content:

```text
OutlookAI - Deployment Guide
============================

Three install shapes are supported:
  A. Single workstation (developer or power user)
  B. Multi-user RDS / Terminal Server (primary deployment target)
  C. IT-managed image / silent install

All three share the same installer:
  Deploy/Install-OutlookAI.ps1

The installer requires Administrator. It is idempotent: running it again
upgrades in place.


PREREQUISITES (all shapes)
---------------------------
- Windows 10 / 11 (Pro or Enterprise) or Windows Server 2019 / 2022 / 2025.
- Microsoft Outlook desktop (2016, 2019, 2021, 2024, or Microsoft 365).
- .NET Framework 4.7.2 or later.
- Visual Studio Tools for Office Runtime:
  https://aka.ms/VSTORuntime
- Microsoft Edge WebView2 Evergreen Runtime. The installer will install it
  silently if missing. To pre-stage offline, vendor the bootstrapper:
    .\Deploy\Fetch-WebView2Bootstrapper.ps1


WHAT THE INSTALLER DOES
-----------------------
1. Backs up any v1 config to:
     C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>
2. Cleans up stale Outlook add-in registrations for every user profile on
   the machine (ClickOnce subscription, VSTA, VSTO SolutionMetadata,
   Inclusion list, Add/Remove Programs, Outlook AddInLoadTimes, ClickOnce
   app cache). This prevents AddInAlreadyInstalledException on upgrade.
3. Closes any running Outlook.exe.
4. Copies the published build to:
     C:\Program Files\OutlookAI
5. Writes a fresh v2 config.xml at C:\Program Files\OutlookAI\config.xml
   if none exists. Server-authoritative defaults: Model, MaxTokens,
   CodexAuthPath.
6. Creates the shared OAuth credential directory at:
     C:\ProgramData\OutlookAI
   with Authenticated Users: Modify (RDS shared-credential model).
7. Renames any per-user %APPDATA%\OutlookAI\config.xml that lacks the v2
   CodexAuthPath element to <name>.v1.backup.<timestamp>.
8. Configures VSTO trust + Inclusion list (HKLM, 64-bit + WOW6432Node).
9. Registers OutlookAI for all users.
10. Configures the Default User profile so new RDS users auto-load it.


SHAPE A - SINGLE WORKSTATION
-----------------------------
Use case: developer machine, power user, single-user laptop or desktop.

Steps:

1. Clone the repo or download the latest Release zip.

2. Publish a Release build:

   & "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
     "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release `
     /p:Platform="Any CPU" /p:PublishDir="C:\OutlookAI\"

3. Run the installer elevated:

   Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
   .\Deploy\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"

4. Open Outlook -> click AI Assistant on the ribbon.

5. Open a compose window or use the taskpane button -> click any action ->
   the default browser opens the ChatGPT OAuth consent page (the consent
   screen says "Codex CLI" because OutlookAI reuses the public Codex
   client_id; this is expected).

6. Sign in. The browser confirms and OutlookAI returns to its taskpane.

7. Verification:
     SHA256 of C:\Program Files\OutlookAI\OutlookAI.dll matches the
     SHA256 of the staged C:\OutlookAI\OutlookAI.dll.


SHAPE B - MULTI-USER RDS / TERMINAL SERVER
-------------------------------------------
Use case: shared server where many interactive users open Outlook with
their own profile and you want the same ChatGPT credential to back all of
them.

Steps:

1. From an admin workstation, build a Release publish bundle and copy it
   to the RDS server, e.g. to C:\OutlookAI on the server.

2. On the RDS server, in an elevated PowerShell:

   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
   cd C:\OutlookAI
   .\Install-OutlookAI.ps1 -SourcePath C:\OutlookAI

3. Designate one admin user as the "first run" user. Have them log in,
   open Outlook, open AI Assistant, and click any action. The OAuth
   browser flow runs once; the resulting auth.json is shared with every
   user on this server.


ACCEPTED RISK - SHARED OAuth CREDENTIAL
----------------------------------------
auth.json sits in C:\ProgramData\OutlookAI with Authenticated Users:
Modify. Any signed-in interactive user on this server can:
  - Read auth.json and copy tokens off the box.
  - Use those tokens to call OpenAI directly until revoked.
  - Delete or corrupt auth.json (signs everyone out).
  - Replace auth.json with their own ChatGPT tokens (other users'
    traffic then bills to the attacker's ChatGPT account and the
    attacker can observe every call).

Only deploy this build to RDS servers where every interactive user is
trusted with the ChatGPT account that signs in.


ROTATING CREDENTIALS
--------------------
Phase 1 has no remote revocation; rotation is a manual two-step:

1. On the RDS server, as any user who knows the OutlookAI admin password:
   - Open Outlook -> AI Assistant -> gear icon (Settings).
   - Enter the OutlookAI admin password.
   - Click "Sign Out" in the ChatGPT Account section.
   - Click "Sign In" and authenticate with the new ChatGPT account.

2. From the OpenAI side (recommended after any suspected leak):
   - Sign the previous account out at https://chatgpt.com/#settings
     (rotates session tokens server-side).


SHAPE C - IT-MANAGED IMAGE / SILENT INSTALL
--------------------------------------------
Use case: corporate Windows image, MDT/SCCM rollout, or any deployment
where the install must complete without operator interaction.

Pre-stage the WebView2 runtime so the installer has no internet dependency
during image build:

   .\Deploy\Fetch-WebView2Bootstrapper.ps1

Then bake the installer into the image:

1. Copy the published OutlookAI bundle (the contents of the PublishDir)
   to a known location on the gold image, e.g. C:\OutlookAI.

2. Add a run-once install task (Group Policy startup script, SCCM task
   sequence, MDT package, or Task Scheduler "At startup") that calls:

   powershell.exe -NoProfile -ExecutionPolicy Bypass `
     -File C:\OutlookAI\Install-OutlookAI.ps1 -SourcePath C:\OutlookAI

   The script enforces #Requires -RunAsAdministrator. Schedule the task
   to run as SYSTEM or a local admin account.

3. Post-install verification (run on a sample VM provisioned from the
   image):

   $staged    = (Get-FileHash 'C:\OutlookAI\OutlookAI.dll' -Algorithm SHA256).Hash
   $installed = (Get-FileHash 'C:\Program Files\OutlookAI\OutlookAI.dll' -Algorithm SHA256).Hash
   if ($staged -ne $installed) { throw "OutlookAI DLL hash mismatch" }

4. First-run experience for end users: open Outlook -> AI Assistant ->
   first action triggers the OAuth flow per user (or, on RDS-style images,
   relies on the shared auth.json from Shape B).


VERIFICATION (all shapes)
--------------------------
1. Outlook -> File -> Options -> Add-ins lists "OutlookAI".
2. Outlook ribbon shows the AI Assistant group.
3. Click AI Assistant -> taskpane opens.
4. Click a Quick Action / type a chat message -> response streams in.
5. C:\ProgramData\OutlookAI\auth.json exists (after first sign-in).
6. (Optional) SHA256 of installed OutlookAI.dll matches staged build.


TROUBLESHOOTING
---------------

Add-in shows in list but won't load / keeps unchecking:
  1. Confirm VSTO Runtime is installed (Programs and Features:
     "Microsoft Visual Studio 2010 Tools for Office Runtime").
  2. File -> Options -> Add-ins -> Manage: Disabled Items -> Go ->
     enable OutlookAI if listed there.
  3. File -> Options -> Add-ins -> Manage: COM Add-ins -> Go ->
     tick OutlookAI; note any error.
  4. Event Viewer -> Windows Logs -> Application -> look for "Outlook"
     or ".NET Runtime" errors.

OAuth sign-in doesn't open a browser:
  - Confirm the user has a default browser configured.
  - Confirm http://localhost:1455 is not blocked locally; the installer
    does not modify firewall rules because the listener is loopback only.

Sign-in returns immediately with "OAuth state mismatch":
  - Click Sign In again; this is usually a stale browser tab racing the
    fresh authorize URL.

ChatGPT Codex backend returns 4xx:
  - 401: token rotated remotely; click Sign Out then Sign In.
  - 429: ChatGPT subscription rate-limited; wait or upgrade plan.
  - 403 with HTML body: Cloudflare challenge; retry once.

Realtime voice fails with "beta_api_shape_disabled":
  - Should not happen on this build (no OpenAI-Beta header is sent).
    If you see it, file an issue with the exact error_id.

PDF export fails with HRESULT 0x8007139F:
  - Indicates a WebView2 user-data folder conflict. The current build
    isolates the PDF renderer at
    C:\Users\<user>\AppData\Local\OutlookAI\WebView2PdfData. If conflicts
    persist, delete that folder and retry.

Search takes very long on large mailboxes:
  - The current build caps interactive broad all-mail scans at 200
    folders and early-stops once enough candidates have been collected.
    See VSTO2\OutlookAI\Services\Tools\SearchFallbackBudget.cs.


UNINSTALL
---------
1. Open PowerShell as Administrator.
2. Run: .\Uninstall-OutlookAI.ps1

This removes:
  - HKLM Outlook add-in registration (64-bit + WOW6432Node).
  - C:\Program Files\OutlookAI install directory.
  - C:\ProgramData\OutlookAI\auth.json + sidecar refresh lock.

Backups under C:\ProgramData\OutlookAI\Backups are intentionally preserved
for rollback.


ROLLBACK TO v1
--------------
1. Run Uninstall-OutlookAI.ps1.
2. Reinstall the v1.x publish artifacts using the v1 installer.
3. Restore the latest backup over the new install:

   Copy-Item `
     "C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>" `
     "C:\Program Files\OutlookAI\config.xml" -Force


SUPPORT
-------
For issues, file a GitHub issue at:
  https://github.com/kirklandsig/OutlookAI/issues
```

- [ ] **Step 2: Sanity-check the file**

```powershell
(Get-Content -LiteralPath "Deploy\README.txt").Length
Select-String -Path "Deploy\README.txt" -Pattern "SHAPE A|SHAPE B|SHAPE C|ACCEPTED RISK|ROTATING CREDENTIALS|HRESULT 0x8007139F|SearchFallbackBudget" | Select-Object -First 10
```

Expected: roughly 230–290 lines and grep matches every section header listed
above.

- [ ] **Step 3: Commit**

```powershell
git add Deploy/README.txt
git commit -m "docs(deploy): rewrite deployment guide for three install shapes"
```

---

## Task 3: Add `docs/Install.md`

Add a short pointer doc so the README can link to a single
markdown-friendly install summary without bloating the README body.

**Files:**

- Create: `docs/Install.md`

- [ ] **Step 1: Confirm `docs/` exists**

```powershell
if (-not (Test-Path -LiteralPath "docs")) { throw "Missing docs directory" }
```

Expected: no output.

- [ ] **Step 2: Write `docs/Install.md`**

Use UTF-8 and the following exact content:

````markdown
# Installing OutlookAI

OutlookAI supports three install shapes. All three share the same installer
(`Deploy/Install-OutlookAI.ps1`) and the same OAuth flow (sign in once with
your ChatGPT account, then OutlookAI uses your existing subscription for
inference).

| Shape | Use case | Detail |
|---|---|---|
| **Single workstation** | One developer or power user. | [Deploy/README.txt — Shape A](../Deploy/README.txt) |
| **Multi-user RDS / Terminal Server** | Shared server, many interactive users, one shared ChatGPT credential. | [Deploy/README.txt — Shape B](../Deploy/README.txt) |
| **IT-managed image / silent install** | MDT / SCCM / corporate gold image. | [Deploy/README.txt — Shape C](../Deploy/README.txt) |

## Quick start (single workstation)

```powershell
git clone https://github.com/kirklandsig/OutlookAI.git
cd OutlookAI

# Publish Release into a staging folder
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" `
  /p:PublishDir="C:\OutlookAI\"

# Install elevated
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
.\Deploy\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"

# Open Outlook → AI Assistant → sign in with your ChatGPT account.
```

For the full deployment story (cleanup, shared credentials, rotation,
troubleshooting, rollback, uninstall), see
[`Deploy/README.txt`](../Deploy/README.txt). That file is the canonical
install guide; this page is a pointer.
````

- [ ] **Step 3: Sanity-check the file**

```powershell
Test-Path -LiteralPath "docs/Install.md"
Select-String -Path "docs/Install.md" -Pattern "Shape A|Shape B|Shape C" | Select-Object -First 6
```

Expected: `True` and three matches.

- [ ] **Step 4: Commit**

```powershell
git add docs/Install.md
git commit -m "docs(install): add three-shape install pointer doc"
```

---

## Task 4: Verification gate

Confirm the docs-only changes did not break build or tests, and that no
README anchor links are broken.

**Files:** none (verification only).

- [ ] **Step 1: Run WebUI syntax checks**

```powershell
node --check VSTO2\OutlookAI\WebUI\chat.js
node --check VSTO2\OutlookAI\WebUI\markdown.js
```

Expected: both commands exit 0 with no output.

- [ ] **Step 2: Build Debug**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds. The existing `MSB3277` warning about
`System.Runtime.CompilerServices.Unsafe` is preserved (it was present
before this plan).

- [ ] **Step 3: Run the full VSTest suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Total tests: 546   Passed: 546`.

- [ ] **Step 4: Check README links**

```powershell
Select-String -Path "README.md" -Pattern "\]\(([^)]+)\)" -AllMatches | ForEach-Object {
    foreach ($m in $_.Matches) { $m.Groups[1].Value }
} | Where-Object { $_ -notmatch "^(https?://|#)" } | Sort-Object -Unique
```

Expected output (every line must resolve to a real file on disk):

```
Deploy/README.txt
LICENSE
docs/Install.md
```

For each printed path, verify it exists:

```powershell
foreach ($p in @("Deploy/README.txt","LICENSE","docs/Install.md")) {
    if (-not (Test-Path -LiteralPath $p)) { throw "README links to missing file: $p" }
}
```

Expected: no output. If `LICENSE` is missing, list the actual files in the
repo root with `Get-ChildItem -LiteralPath "." -Filter "LICENSE*"` and
adjust the README to point to the real filename.

- [ ] **Step 5: Confirm working tree is clean**

```powershell
git status --short
```

Expected: empty output (all three docs commits already landed in Tasks 1–3).

---

## Task 5: Refresh PR #1 description and push branch

Update the open pull request so its body reflects the launch positioning,
then push the branch.

**Files:** none (PR description + remote update only).

- [ ] **Step 1: Push the branch**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
git push -u origin feature/codex-oauth-migration
```

Expected: push succeeds; the remote `feature/codex-oauth-migration` is now
ahead of the previous PR head.

- [ ] **Step 2: Verify the PR is still open and pointed at this branch**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr list --head feature/codex-oauth-migration --json url,state,title,number
```

Expected: one PR with `"state": "OPEN"` and `"number": 1`. If `state` is
not `OPEN`, stop and surface the result before proceeding.

- [ ] **Step 3: Update PR #1 body**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
$body = @'
## Summary
- OutlookAI v2 launch: replace the Phase 2-era README and deployment guide with a product landing page and three-shape install guide that match what is actually shipped on `feature/codex-oauth-migration`.
- Highlight Chat, Inbox Copilot, Inbox Reports, Excel and PDF exports, per-message Save as PDF, voice, and the 15-tool mailbox surface (11 read/utility + 4 admin-gated safe writes).
- Rewrite `Deploy/README.txt` around three install shapes (single workstation, RDS/Terminal Server, IT-managed image) with explicit RDS shared-credential risk and rotation steps.
- Add `docs/Install.md` as a short pointer doc.

## Test Plan
- `node --check VSTO2\OutlookAI\WebUI\chat.js`
- `node --check VSTO2\OutlookAI\WebUI\markdown.js`
- VS MSBuild Debug Any CPU succeeded (known `MSB3277` warning preserved)
- VSTest passed `546/546`
- Installed Release DLL hash matched staged build during the most recent install/smoke (`9FD98ED352F7A14022D412E30154BC931B4BBBA3778DF27EB4664BE9301D5325`)
- Smoke validated Action Items runtime, Excel export, per-message PDF export, PDF rendering, and file-card open/reveal
- README link sanity check (`Deploy/README.txt`, `LICENSE`, `docs/Install.md` all resolve)
'@
gh pr edit 1 --body $body
```

Expected: command prints the PR URL.

- [ ] **Step 4: Surface the PR URL**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr view 1 --json url --jq '.url'
```

Expected: `https://github.com/kirklandsig/OutlookAI/pull/1` (or the
equivalent for the repo).

---

## Task 6: Merge PR #1 to `master` and push

Merge the launch PR and update local `master`.

**Files:** none (merge + push only).

- [ ] **Step 1: Confirm CI / required checks (if any) are clean**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr view 1 --json mergeStateStatus,mergeable,statusCheckRollup
```

Expected: `"mergeable": "MERGEABLE"` and either an empty `statusCheckRollup`
(no required checks configured) or a rollup with all checks in
`"conclusion": "SUCCESS"`. If anything else, stop and surface the result.

- [ ] **Step 2: Merge the PR (merge commit, no squash)**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr merge 1 --merge --subject "Merge OutlookAI v2 launch (Chat, Inbox Copilot, Inbox Reports, exports)"
```

Expected: command exits 0 and prints `Merged pull request #1`.

- [ ] **Step 3: Update local `master`**

```powershell
git fetch origin master
git checkout master
git pull --ff-only origin master
git log -1 --oneline
```

Expected: the latest commit on `master` is the merge commit just created.

- [ ] **Step 4: Verify the build still works on `master`**

```powershell
node --check VSTO2\OutlookAI\WebUI\chat.js
node --check VSTO2\OutlookAI\WebUI\markdown.js
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: both `node --check` exits 0, MSBuild succeeds with the known
`MSB3277` warning, and VSTest reports `Total tests: 546   Passed: 546`.

- [ ] **Step 5: Push `master` if it is behind**

```powershell
git status -sb
```

If `git status` shows `master` is `ahead` of `origin/master`, push:

```powershell
git push origin master
```

Expected: push succeeds or no-op (`Everything up-to-date`).

- [ ] **Step 6: Switch back to the feature worktree branch (so subsequent work continues there)**

```powershell
git checkout feature/codex-oauth-migration
git status -sb
```

Expected: `feature/codex-oauth-migration...origin/feature/codex-oauth-migration`
with no extra ahead/behind tracking diff.

---

## Self-Review

**1. Spec coverage:**

- Spec "Goal 1: Rewrite `README.md` as a product landing page" → Task 1.
- Spec "Goal 2: Tell the full story (Chat, Inbox Copilot, Inbox Reports,
  Excel/PDF exports, voice, write-tool permissions, RDS/Server deployment)"
  → Task 1 (all sections covered in the README body).
- Spec "Goal 3: Refresh `Deploy/README.txt` and add `docs/Install.md`" →
  Tasks 2 and 3.
- Spec "Goal 4: Keep all claims factual and dated to what is actually
  shipped" → Tasks 1–3 (texts pinned to 546/546, 15 tools, six default
  report chips, etc.).
- Spec "Goal 5: Set explicit still-in-development expectations and link to
  the open PR" → Task 1 (banner at top of README), Task 5 (PR description
  refresh).
- Spec "Acceptance criterion: tests, build, installed DLL hash remain
  unchanged" → Task 4 (Debug build + VSTest), Task 6 (post-merge re-verify).
- Spec "Acceptance criterion: README link sanity" → Task 4 Step 4.
- Spec "Acceptance criterion: merge to master and push" → Task 6.

**2. Placeholder scan:** No `TBD`, `TODO`, `implement later`, or `similar to
Task N` phrases in this plan. All file contents are inlined verbatim in the
relevant task step.

**3. Type / name consistency:**

- Tool count is `15` (11 + 4) in spec, README body, and PR description.
- Test count is `546/546` in spec, README body, contributing block, PR
  description, and verification gate.
- File paths are consistent: `Deploy/README.txt`, `docs/Install.md`,
  `README.md`, `LICENSE`, `VSTO2/OutlookAI/Services/Tools/SearchFallbackBudget.cs`.
- Branch name `feature/codex-oauth-migration` and PR `#1` are used
  consistently in Tasks 5 and 6.

---

## Final State After This Plan

- `README.md` rewritten as a product landing page.
- `Deploy/README.txt` rewritten around three install shapes.
- `docs/Install.md` created as a short pointer doc.
- Test count, build, and installed DLL hash unchanged from branch tip.
- PR #1 body refreshed and merged into `master`.
- Local `master` updated to the merge commit.
- Working tree returned to `feature/codex-oauth-migration` for any
  follow-up work.
