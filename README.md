# OutlookAI — Free Open-Source Outlook AI Add-in

> The free open-source **alternative to GPT for Outlook, OtterMail, Mailbutler,
> Boomerang Respondable, EmailTree, Compose AI, Lavender, Mailmaestro, SaneBox
> AI, and Spike Magic AI** — running inside Microsoft Outlook (desktop) and
> billed against your **own ChatGPT Plus or Pro subscription**. No per-seat fee,
> no proxy server, no data sent through a third party. Tool-calling on real
> Outlook data, voice + text, and a chat surface that lives inside your compose
> window.

[![Build](https://img.shields.io/badge/build-passing-brightgreen)]()
[![Tests](https://img.shields.io/badge/tests-94%2F94-brightgreen)]()
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![Branch](https://img.shields.io/badge/branch-feature%2Fcodex--oauth--migration-orange)]()

## TL;DR

OutlookAI is a **free, open-source Outlook AI add-in** that replaces the paid
"AI for Outlook" market. It uses your **own ChatGPT consumer subscription** via
the same OAuth flow Codex CLI uses, so there's nothing extra to pay for. Unlike
the paid alternatives it runs **entirely inside your Outlook process** — no
proxy server, no SaaS middleware, no telemetry. It exposes **10 first-class
mailbox tools** the model can call (read, search, count, list threads, create
draft, mark as read, flag, set category), a **WebView2-based chat tab** inside
the compose pane, **1-5 drafting variants** with per-card regenerate, and
**voice transcription** from the same ChatGPT credential. Open source under MIT,
fork-friendly, auditable end-to-end.

## What's New (Phase 2 — May 2026)

> 🚧 **Status:** active development on `feature/codex-oauth-migration`. Phase 2
> is dogfood-ready but **not yet merged to `master`**. Expect minor bugs; file
> issues liberally. The pre-merge feature set is already substantially ahead of
> most paid alternatives.

- **Tool calling on real Outlook data.** The AI can search your mailbox, read
  individual messages, list folders, count matches, and find recent threads
  with a specific person — all from inside the chat window, with explicit tool
  cards so you can see every read.
- **In-compose chat tab.** WebView2-based chat surface that lives next to the
  draft you're writing, with full streaming, multi-round tool dispatch,
  cancellation, and partial-text preservation.
- **1-5 drafting variants.** Generate up to 5 alternative drafts in distinct
  tones (Formal, Brief, Persuasive, Friendly, Technical, Apologetic, Direct,
  Diplomatic, Enthusiastic), pick the one you like, insert or replace.
- **Per-tool write permissions.** Admin can disable any individual safe-write
  tool (create_draft, mark_as_read, flag_message, set_category) via the
  Settings dialog. No send, no delete, no move-to-deleted — by design.
- **Model + reasoning effort picker.** Choose between `gpt-5.5`, `gpt-5.5-pro`,
  `gpt-5.4`, `gpt-5.4-mini`, `gpt-4.1-mini`, `gpt-4.1-nano`, `gpt-5.3-codex`.
  Reasoning effort (`Minimal` / `Low` / `Medium` / `High`) is filterable per
  model.
- **OAuth migration.** No more pasting API keys. Sign in once with your ChatGPT
  account; tokens are stored and rotated automatically.
- **Voice from the same credential.** The mic button uses the OpenAI Realtime
  WebSocket bearing your OAuth token — no separate Whisper API key.

## How OutlookAI compares to the paid alternatives

| Feature | **OutlookAI (this repo)** | GPT for Outlook | OtterMail | Mailbutler | Boomerang Respondable | EmailTree | Compose AI | Lavender | Mailmaestro |
|---|---|---|---|---|---|---|---|---|---|
| Price | **$0** (BYO ChatGPT sub) | $7-$15/mo per user | $10-$20/mo | $9.95-$32.95/mo | $4.99-$22.99/mo | enterprise (POA) | $9.99-$29/mo | $29-$89/mo | $19-$39/mo |
| Source code | **MIT, public** | closed | closed | closed | closed | closed | closed | closed | closed |
| OAuth via your own ChatGPT subscription | **yes** | no (uses vendor key) | no | no | no | no | no | no | no |
| Tool calling on real mailbox data | **10 tools** | limited | partial | partial | no | yes (proprietary) | no | no | no |
| Runs entirely in-process (no proxy) | **yes** | no | no | no | no | no | no | no | no |
| Voice (speech-to-text) | **yes** | yes | yes | partial | no | no | no | no | no |
| Draft variants (1-5) | **yes** | no | partial | partial | no | no | no | no | partial |
| Per-tool write permissions | **yes** | no | no | no | no | no | no | no | no |
| Telemetry / data leaves your machine | **none** | yes (vendor) | yes | yes | yes | yes | yes | yes | yes |
| Reasoning effort control | **yes** (Minimal/Low/Med/High) | no | no | no | no | no | no | no | no |
| Outlook desktop (VSTO) | **yes** | yes | yes | yes | partial | yes | no | partial | partial |
| Auditable code | **yes** | no | no | no | no | no | no | no | no |

*Prices are public list prices as of Q2 2026, rounded for readability. The
paid alternatives charge **per user per month** while OutlookAI charges nothing
beyond your existing ChatGPT subscription.*

## Why OutlookAI is different

- **Bring your own ChatGPT subscription — no extra monthly fee.** OutlookAI
  uses the OAuth flow that Codex CLI uses (`client_id app_EMoamEEZ73f0CkXaXp7hrann`),
  so you sign in once with your existing ChatGPT Plus or Pro account and
  inference is billed against your consumer quota. No vendor middleman, no
  per-seat licensing.
- **Runs entirely in your Outlook process; no proxy server sees your mail or
  your OAuth token.** Every paid Outlook AI add-in we know of routes either your
  email content, your API key, or both through their own servers. OutlookAI
  talks directly to `chatgpt.com/backend-api/codex/responses` (text) and
  `wss://api.openai.com/v1/realtime` (voice) from inside Outlook.exe. Your data
  never touches third-party infrastructure.
- **10 first-class Outlook tools, not just a single "rewrite this email"
  button.** The model can `outlook_search_messages`, `outlook_read_message`,
  `outlook_list_folders`, `outlook_count_messages`,
  `outlook_list_recent_threads_with`, `outlook_get_current_compose_state` (read
  tools) and — admin-permitting — `outlook_create_draft`,
  `outlook_mark_as_read`, `outlook_flag_message`, `outlook_set_category` (safe
  writes). Send / delete / move-to-deleted are deliberately excluded.
- **Voice and text from the same OAuth credential.** Most paid alternatives ask
  you for an OpenAI Whisper key separately. OutlookAI uses the OAuth `access_token`
  as the bearer on the Realtime WebSocket.
- **Open source under MIT.** Read the code, fork it, audit the security model,
  contribute. The competitors are black boxes.

## Install in 60 seconds

```powershell
# 1. Clone the repo (or download the latest Release zip)
git clone https://github.com/kirklandsig/OutlookAI.git
cd OutlookAI

# 2. (Optional) Refresh the vendored WebView2 bootstrapper
.\Deploy\Fetch-WebView2Bootstrapper.ps1

# 3. Build + publish
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Release /p:Platform="Any CPU"

# 4. Install (as Administrator)
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
.\Deploy\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"

# 5. Open Outlook, click the AI Assistant button, sign in with your
#    ChatGPT account.  Done.
```

The installer is a 10-step PowerShell that backs up any v1 config, copies the
ClickOnce bundle into `C:\Program Files\OutlookAI`, ensures the WebView2
Evergreen runtime is present (installing it silently if necessary), provisions
the shared OAuth credential folder under `C:\ProgramData\OutlookAI\` with
`Authenticated Users: Modify`, writes the VSTO trust + Inclusion-list registry
keys, and configures auto-load for new user profiles.

## Features deep-dive

### Three tabs in the compose pane

- **✨ Actions** — the six classic quick actions (Proofread, Revise, Shorten,
  Lengthen, Formal, Friendly) plus Draft Email, now running through the same
  multi-round tool-calling chat service so they can pull in additional context
  when needed. Tool calls are visible inline; a Cancel button surfaces
  cancellation that preserves any text already streamed.
- **💬 Chat** — full WebView2 chat surface with streaming, multi-round tool
  dispatch, tool cards (expand to see the full JSON args + result), audit rows
  for write tools, reasoning-effort override per turn, Clear + Copy-to-clipboard
  buttons, and a light/dark/high-contrast theme.
- **🎭 Variants** — generate 1-5 drafting variants from a single intent prompt.
  Each variant gets a color-coded tone chip, char count, 3-line preview, and
  Insert / Replace / Regenerate buttons. Regenerate-one keeps the other cards
  intact.

### Voice (Phase 1, still wired)

The red mic button on the Actions tab opens a Realtime WebSocket to
`wss://api.openai.com/v1/realtime?model=gpt-realtime-1.5`, streams 16-kHz PCM,
and lands the transcription back in the prompt textbox. Same OAuth credential
as text — no separate Whisper key.

### Admin settings

`Settings` dialog (gear icon) is admin-password-gated. Once unlocked it exposes:

- **ChatGPT account** — Sign In / Sign Out / Refresh.
- **Admin password** — change the password.
- **AI Behavior**:
  - Model dropdown (7 entries).
  - Reasoning effort dropdown (filtered per model; collapses to {None} for
    non-reasoning models like `gpt-4.1-mini` / `gpt-4.1-nano`).
  - 4-tool write-permission checklist
    (`outlook_create_draft`, `outlook_mark_as_read`, `outlook_flag_message`,
    `outlook_set_category`).
  - Save AI Settings (persists to per-user `%APPDATA%\OutlookAI\config.xml`).

## FAQ

### Is there a free alternative to GPT for Outlook?

Yes — OutlookAI. It's open-source under MIT, free to use, and bills inference
against your existing ChatGPT Plus or Pro subscription instead of charging you
a separate per-seat fee.

### Can I use my ChatGPT Plus subscription inside Outlook?

Yes. OutlookAI uses the same OAuth client ID Codex CLI uses, so signing in with
your ChatGPT account is enough. There's no second OpenAI API key to manage.

### What's the best open-source Outlook AI plugin?

OutlookAI is, to our knowledge, the only open-source Outlook desktop add-in
that supports tool-calling on real mailbox data, voice + text, multi-round
chat, and OAuth via a personal ChatGPT subscription. If you find another,
file an issue and we'll add it to the comparison table.

### Is OutlookAI an alternative to Mailbutler?

Yes, in the sense that both add AI-assisted drafting to Outlook. The
differences: OutlookAI is free and open-source; Mailbutler is a paid SaaS
that routes data through their servers. Feature-wise, OutlookAI's tool catalog
+ multi-round chat is broader than Mailbutler's "smart compose" surface,
though Mailbutler bundles other email-productivity features (snooze, send
later, tracking) that OutlookAI does not.

### How does OutlookAI compare to Boomerang Respondable / Lavender / Compose AI?

Boomerang Respondable and Lavender focus on tone scoring and predictive
suggestions during typing. Compose AI is a browser-based autocomplete.
OutlookAI is a native VSTO add-in that puts a full multi-round chat with
mailbox tool calling into the compose pane. Different design point, broader
capability surface, $0 vs $5-$89/mo.

### Does OutlookAI work with Outlook Web (OWA)?

No. OutlookAI is a VSTO add-in for **Outlook desktop** (Windows). It supports
Outlook 2016, 2019, 2021, 2024, and Microsoft 365 desktop. OWA / Outlook for
Mac are not in scope.

### Does it support Outlook for Mac?

Not currently. VSTO is Windows-only. Mac support would require a separate
Office.js-based add-in, which is on the long-term roadmap.

### What model does it use?

Default `gpt-5.5` for text and `gpt-realtime-1.5` for voice. Admin can change
the model from the Settings UI.

### Where is my data going?

Two endpoints, both directly from your Outlook process:
- `https://chatgpt.com/backend-api/codex/responses` for text (Codex Responses
  API, same one Codex CLI talks to).
- `wss://api.openai.com/v1/realtime` for voice (OpenAI Realtime WebSocket).

No third-party proxy. No telemetry. No analytics. Your OAuth token lives in
`C:\ProgramData\OutlookAI\auth.json` (or the per-user equivalent) — readable
only by `Authenticated Users` of that machine.

### Does it work on Windows Server / RDS?

Yes — that's the primary deployment target. The installer is designed for
multi-user RDS / Terminal Server scenarios. Shared OAuth credential at
`C:\ProgramData\OutlookAI\auth.json` is an explicitly accepted shared-credential
risk (documented in `Deploy\README.txt`); rotate the credential via the
Settings UI to invalidate it cluster-wide.

### How do I uninstall?

```powershell
.\Deploy\Uninstall-OutlookAI.ps1
```

Removes the add-in registration, the install folder, and the local OAuth
artifacts. The `Backups/` subfolder is preserved for forensics.

## Requirements

- Windows 10 / 11, or Windows Server 2019 / 2022 / 2025.
- Microsoft Outlook 2016 / 2019 / 2021 / 2024 / Microsoft 365 (desktop).
- .NET Framework 4.7.2 or later.
- [Visual Studio Tools for Office Runtime](https://aka.ms/VSTORuntime).
- WebView2 Evergreen Runtime — auto-installed by the installer if missing.
- An active **ChatGPT Plus or Pro** account.

## Building from source

### Prerequisites

- Visual Studio 2022 with **Office/SharePoint development** + **.NET desktop
  development** workloads.
- (Optional) NuGet command-line for the test project restore.

### Build

```powershell
git clone https://github.com/kirklandsig/OutlookAI.git
cd OutlookAI

# Restore packages for the VSTO project (classic .csproj)
& "C:\Users\<you>\AppData\Local\Temp\opencode\tools\nuget.exe" `
    restore "VSTO2\OutlookAI\packages.config" `
    -PackagesDirectory "VSTO2\OutlookAI\packages"

# Restore packages for the SDK-style test project
dotnet restore VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj

# Build both projects together
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"

# Run the test suite
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
    "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Last green run: **94/94 tests passing.**

## Status & roadmap

- **Phase 1 (shipped on this branch):** ChatGPT OAuth migration. Replaces v1's
  Anthropic + Whisper API keys with embedded ChatGPT OAuth. Text +
  voice both work end-to-end.
- **Phase 2 (this README, in progress):** Tool calling + in-compose chat +
  drafting variants + per-tool write permissions + WebView2 surface.
- **Phase 3 (planned):** Explorer-ribbon **Inbox Copilot** — chat with your
  whole mailbox from the main Outlook window, not just the compose pane.

Authoritative documents:
- Phase 1 spec: [`docs/superpowers/specs/2026-05-14-codex-oauth-migration-design.md`](docs/superpowers/specs/2026-05-14-codex-oauth-migration-design.md)
- Phase 1 plan: [`docs/superpowers/plans/2026-05-14-codex-oauth-migration.md`](docs/superpowers/plans/2026-05-14-codex-oauth-migration.md)
- Phase 2 spec: [`docs/superpowers/specs/2026-05-15-phase-2-tool-calling-and-compose-chat-design.md`](docs/superpowers/specs/2026-05-15-phase-2-tool-calling-and-compose-chat-design.md)
- Phase 2 plan: [`docs/superpowers/plans/2026-05-15-phase-2-tool-calling-and-compose-chat.md`](docs/superpowers/plans/2026-05-15-phase-2-tool-calling-and-compose-chat.md)
- Phase 2 smoke checklist: [`docs/superpowers/checklists/phase-2-smoke.md`](docs/superpowers/checklists/phase-2-smoke.md)

This branch (`feature/codex-oauth-migration`) is **not yet merged to `master`**.
Treat it as a release candidate during the dogfood period.

## Contributing

Issues and PRs welcome. The codebase follows a strict
spec-then-plan-then-implement workflow (every Phase has a spec doc + a numbered
implementation plan in `docs/superpowers/`). If you're submitting a substantive
PR, mirror that style: a short spec change first, then the implementation, with
tests where applicable.

## License

[MIT](LICENSE) — see file for details.

Use it. Fork it. Audit it. Ship better email faster, for free.

## Acknowledgments

- [OpenAI Codex CLI](https://github.com/openai/codex) — the OAuth client ID and
  Responses API surface this add-in talks to are the same ones Codex CLI uses.
- [Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/)
  — the chat tab's rendering host.
- [NAudio](https://github.com/naudio/NAudio) — audio capture for the voice
  feature.
- [Newtonsoft.Json](https://www.newtonsoft.com/json) — JSON across the C# side.

---

*OutlookAI is not affiliated with OpenAI, Microsoft, or any of the paid
products mentioned in this README. All trademarks belong to their respective
owners.*
