# OutlookAI

> ## ⚠️ Active Development Branch — `feature/codex-oauth-migration`
>
> **This branch is work-in-progress and may contain bugs. Do not deploy to production.**
>
> - Phase 1 (this branch): replaces the Anthropic/OpenAI API-key auth with embedded ChatGPT OAuth.
>   Text → `chatgpt.com/backend-api/codex/responses`. Voice → `wss://api.openai.com/v1/realtime`.
>   Both billed against the user's ChatGPT consumer subscription.
> - Phase 2 (planned): tool calling so the chat service can search/read mailbox via Outlook OOM.
> - Phase 3 (planned): Inbox-Copilot UI on the Explorer ribbon (chat with your mailbox).
>
> Sections below still describe the v1 (Claude + Whisper API-key) shape and have not been
> rewritten yet. Authoritative spec for this branch lives at
> [`docs/superpowers/specs/2026-05-14-codex-oauth-migration-design.md`](docs/superpowers/specs/2026-05-14-codex-oauth-migration-design.md).
> Implementation plan: [`docs/superpowers/plans/2026-05-14-codex-oauth-migration.md`](docs/superpowers/plans/2026-05-14-codex-oauth-migration.md).
>
> This branch is intentionally not being merged to `master` until validated.

An AI-powered email writing assistant for Microsoft Outlook, built as a VSTO add-in.

<img width="283" height="317" alt="image" src="https://github.com/user-attachments/assets/7513e75c-c226-4791-853a-d1aacd897883" />

## What's New (April 2026)

- **Updated AI models** - Now uses Claude Opus 4.6 (writing) and GPT-4o Transcribe (voice)
- **Global config file** - Admins can set API keys and models via `C:\Program Files\OutlookAI\config.xml` without rebuilding
- **Server 2025 fix** - Fixed invisible button text on Windows Server 2025
- **Resiliency protection** - Outlook can no longer auto-disable the add-in
- **Better error messages** - API errors now show in a full dialog instead of a truncated label

## Features

- **Quick Actions** - One-click buttons to improve your email drafts:
  - Proofread (grammar, spelling, punctuation)
  - Revise (clarity and flow)
  - Shorten / Lengthen
  - Formal / Friendly tone

- **Draft New Emails** - Describe what you want to write and let AI generate the email
  - Voice input support using OpenAI Whisper
  - Context-aware replies (AI sees the email chain)

- **Insert or Replace** - Choose to add content at the top (preserving email chain) or replace everything

## Requirements

- Windows 10/11 or Windows Server 2019/2022/2025
- Microsoft Outlook 2016, 2019, 2021, or 2024 (desktop version)
- .NET Framework 4.8
- [Visual Studio Tools for Office Runtime](https://aka.ms/VSTORuntime)

## API Keys Required

This add-in requires:
- **Anthropic API Key** (Claude) - Required for all AI features. Get one at [console.anthropic.com](https://console.anthropic.com)
- **OpenAI API Key** (Whisper) - Optional, for voice input. Get one at [platform.openai.com](https://platform.openai.com)

## Installation

### Option 1: Pre-configured Build (Enterprise/RDS)

1. Build the solution in Release mode
3. Publish from Visual Studio (Right-click project > Publish)
4. Copy the publish folder to your deployment location
5. Run `Deploy\Install-OutlookAI.ps1` as Administrator:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
Unblock-File -Path "C:\OutlookAI\Install-OutlookAI.ps1"
cd C:\OutlookAI
.\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"
```
6. Edit `C:\Program Files\OutlookAI\config.xml` to set your API keys and model preferences

### Option 2: Per-User Install

1. Build and publish the solution
2. Run `setup.exe` from the publish folder
3. Open Outlook and configure API keys in the Settings panel

## Usage

1. Open Outlook and compose a new email (New, Reply, or Forward)
2. Click the **AI Assistant** button in the ribbon
3. The task pane opens on the right side

### Quick Actions
- Write your email draft first
- Click any Quick Action button (Proofread, Revise, etc.)
- Review the result and click **Insert** or **Replace**

### Draft New Email
- Type or speak your instructions (e.g., "Write a thank you email to John for the meeting")
- Click **Draft Email**
- Review and insert the result

### Voice Input
- Click the red circle button to start recording
- Speak your instructions
- Click again to stop and transcribe
- Requires OpenAI API key

## Building from Source

### Prerequisites
- Visual Studio 2022
- Office/SharePoint development workload
- .NET desktop development workload

### Build Steps
1. Clone this repository
2. Open `VSTO2\OutlookAI\OutlookAI.sln`
3. Restore NuGet packages
4. Build > Rebuild Solution

## Configuration

Settings are loaded in this order (each overrides the previous):
1. **Hardcoded defaults** in the compiled DLL
2. **Global config** - `C:\Program Files\OutlookAI\config.xml` (admin-managed, applies to all users)
3. **Per-user config** - `%APPDATA%\OutlookAI\config.xml` (created when a user saves settings)

The install script creates a default global config. Edit it to set API keys and models for all users without rebuilding.

Access the Settings panel by clicking the gear icon in the add-in. The default admin password is `admin`.

## Deployment Scripts

Located in the `Deploy` folder:

- `Install-OutlookAI.ps1` - Per-machine install for all users (RDS/Terminal Server)
- `Uninstall-OutlookAI.ps1` - Remove the add-in
- `Enable-OutlookAI-User.ps1` - Re-enable if Outlook disabled the add-in

## Troubleshooting

### Add-in doesn't appear
- Restart Outlook
- Check File > Options > Add-ins
- Run `Enable-OutlookAI-User.ps1`

### Add-in keeps getting disabled
- Outlook's "Resiliency" feature may disable slow-loading add-ins
- Run `Enable-OutlookAI-User.ps1` - this now sets a Group Policy key to permanently prevent disabling
- Add to logon scripts for automatic protection

### "Untrusted" or security errors
- Ensure all files are unblocked (Right-click > Properties > Unblock)
- Or run: `Get-ChildItem -Path "C:\Program Files\OutlookAI" -Recurse | Unblock-File`

### API errors
- Verify your API keys are correct
- Check your API account has credits/quota
- Ensure TLS 1.2 is enabled (default on modern Windows)

## License

MIT License - See [LICENSE](LICENSE) file

## Acknowledgments

- [Anthropic Claude API](https://www.anthropic.com) - AI text generation
- [OpenAI Whisper API](https://openai.com) - Speech-to-text
- [NAudio](https://github.com/naudio/NAudio) - Audio recording
