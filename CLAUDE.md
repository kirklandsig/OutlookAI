# OutlookAI — Developer Guide

## Project Overview
OutlookAI is a VSTO add-in for Outlook desktop on Windows: chat in the compose
pane, Inbox Copilot and Inbox Reports taskpanes, Excel/PDF exports and voice
transcription, all running in-process on the signed-in user's ChatGPT
subscription (Codex OAuth; no API keys, no proxy). The primary deployment
target is multi-user RDS / Terminal Servers, where one shared ChatGPT sign-in
and one set of admin Settings serve every user. Releases are published on
GitHub and installed by the in-app updater.

Stack: C# 7.3 on .NET Framework 4.7.2 (VSTO, WinForms), WebView2 with HTML/JS
for the chat UI, xUnit tests, Windows PowerShell 5.1 install scripts.

Maintainers keep a local, gitignored `handoff.md` with release history and
current state; read it first when it exists.

## Quick Start
```powershell
$msbuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$vstest  = "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe"

# Build (Debug). Without the maintainer's signing certificate, add
# /p:ManifestCertificateThumbprint=<your cert> (README.md -> Contributing).
& $msbuild VSTO2\OutlookAI.sln /restore /p:RestorePackagesConfig=true /p:Configuration=Debug /p:Platform="Any CPU"

# Test: everything, or a filtered subset
& $vstest VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll
& $vstest VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll /TestCaseFilter:"FullyQualifiedName~ConfigTests"

# Install bundle (zip + .sha256, as published in a release) into out\
.\Deploy\Make-ReleaseZip.ps1 -Tag vX.Y.Z -OutDir out
```

## Architecture
- `VSTO2/OutlookAI/` — the add-in.
  - `ThisAddIn.cs`, `Ribbon.cs`; `TaskPane/` hosts the Chat, Inbox Copilot,
    Inbox Reports and Variants controllers and the WebView2 bridge.
  - `Services/CodexChatService.cs` (Responses API streaming, multi-round tool
    dispatch), `Services/CodexAuthService.cs` (OAuth, token refresh),
    `Services/OutlookToolHost.cs` + `Services/Tools/` (the 16 model-callable
    tools over Outlook COM), `Services/Export/` (ClosedXML Excel, WebView2
    PDF), `Services/Models/` (model catalog, Update Models),
    `Services/Updates/` (in-app updater).
  - `Config.cs` (config layers, Settings save) and `SettingsForm.cs` (the
    admin Settings dialog).
  - `WebUI/` — `chat.js`, `markdown.js`, styles and the print template.
- `VSTO2/OutlookAI.Tests/` — xUnit (SDK-style project): unit tests, STA
  WinForms tests, and installer tests that run steps in PowerShell 5.1.
- `Deploy/` — install/uninstall scripts, `Make-ReleaseZip.ps1`, and
  `README.txt` (the deployment guide, shipped inside the zip).
- `.github/workflows/release.yml` — verify-only: checks a published
  release's assets and SHA256. Builds happen locally.
- `docs/superpowers/` — specs and plans per feature. They are point-in-time;
  addenda note later changes.

## Documentation Lookup Table
| Feature | Documentation | Key files |
|---|---|---|
| Install, RDS, updates, uninstall | `Deploy/README.txt` | `Deploy/Install-OutlookAI.ps1`, `Services/Updates/` |
| Config layers, Settings for every user | `Deploy/README.txt` (SETTINGS) | `Config.cs`, `SettingsForm.cs`, `Services/FileLock.cs` |
| Model catalog (Update Models) | `docs/superpowers/specs/2026-09-22-model-catalog-refresh-design.md` | `Services/Models/` |
| In-app updater | `docs/superpowers/specs/2026-05-20-in-app-updater-design.md` | `Services/Updates/` |
| OAuth and the Codex backend | `docs/superpowers/specs/2026-05-14-codex-oauth-migration-design.md` | `Services/CodexAuthService.cs`, `Services/CodexChatService.cs` |
| Tools | `README.md` (tools) | `Services/OutlookToolHost.cs`, `Services/Tools/` |
| Exports | `docs/superpowers/specs/2026-05-18-phase-5-exports-design.md` | `Services/Export/`, `TaskPane/Chat/ExportBridge.cs` |

## Key Patterns

### Add-in (C#)
- C# 7.3 and .NET Framework 4.7.2. The add-in is an old-style csproj: add
  each new `.cs` file as a `<Compile Include>` entry.
- Config layers, setting by setting, later wins: defaults →
  `C:\Program Files\OutlookAI\config.xml` (also the only source of
  CodexAuthPath, VoiceModel, MaxBulkExportRows, ModelCatalogClientVersion) →
  per-user `%APPDATA%\OutlookAI\config.xml` →
  `C:\ProgramData\OutlookAI\config.xml`. Settings writes only the
  ProgramData file, and only what changed in the dialog (write tools per
  tool), merged under `FileLock` and swapped in atomically.
- The Settings dialog measures edits against baselines (what it last loaded,
  reloaded or saved). Values it shows on its own, such as a retired model's
  replacement or Auto for a model without the chosen effort, never count as
  edits.
- Outlook COM access goes through `OutlookThreadMarshaller` onto the UI
  thread.
- WinForms buttons need explicit `ForeColor`/`BackColor` and
  `UseVisualStyleBackColor = false`; Windows Server 2025 renders them blank
  otherwise.
- Test seams are `internal` (e.g. `SettingsForm.SaveSettings`,
  `Config.SaveSettingsTo`, `Config.Clock`), visible to the tests through
  `InternalsVisibleTo`.

### Installer (PowerShell)
- Windows PowerShell 5.1 reads the BOM-less scripts as ANSI: keep
  `Install-OutlookAI.ps1` ASCII-only.
- Steps are delimited by `# --- N.` banners, which `InstallScriptConfigTests`
  uses to slice them; keep the banners.
- Updates keep both `config.xml` files and the per-user files.

### Testing
- Tests that touch `Config` statics use `[Collection("Config")]` and
  `ConfigStateScope`; WinForms tests run on an STA thread (`Sta.Run`).
- Settings dialog tests drive the form with helpers (`Saves`, `Save`,
  `Reload`, `Pick`) while the save seam records instead of writing.
- For a behavior change, check that the new test fails without the fix.

## Release Process
1. Merge to `master` with `--no-ff` ("Merge: … (vX.Y.Z)") and push.
2. Push the annotated tag `vX.Y.Z` first; it starts the verify-only Release
   workflow, which waits for the release to appear.
3. Run `.\Deploy\Make-ReleaseZip.ps1 -Tag vX.Y.Z -OutDir out`, then
   `gh release create vX.Y.Z --title vX.Y.Z --notes-file <notes> out\*.zip out\*.zip.sha256`.
4. Check the workflow run and `releases/latest`, then delete `out\`.

## Code Quality Standards

### Code Review Checkpoints
After completing any feature or significant change, perform a self-review:
- [ ] All tests pass
- [ ] No hardcoded secrets, credentials, or PII
- [ ] Admin-only actions stay behind the Settings admin password
- [ ] Error handling covers edge cases (unreadable or busy files, another
      session saving at the same time)
- [ ] No unused usings or dead code introduced
- [ ] Documentation updated if behavior changed (`README.md`,
      `Deploy/README.txt`)

### Security
- Never commit secrets, tokens, `auth.json` or signing keys (`*.pfx` is
  gitignored).
- Validate input at the boundaries: config.xml values against the model
  catalog and the tool list; file actions through `IExportPathPolicy`.
- No send, delete or move tools: model writes are limited to the four
  admin-gated safe writes.

## Data Sanitization
- Tests and docs use fake data only (generic names, example.com addresses,
  555-xxxx phone numbers).
- Never commit real mailbox content, trace logs or local credentials; keep
  local test tokens in gitignored `*.local.json` files.

## Important Gotchas
- **Subscription OAuth:** inference goes to
  `chatgpt.com/backend-api/codex/responses` (not `api.openai.com/v1/*`) with
  `stream: true` and dotted model slugs (`gpt-5.5`). The model list comes
  from `.../codex/models?client_version=X`, and the client version gates
  which models appear.
- **Antivirus** (e.g. Bitdefender) can quarantine `OutlookAI.dll` under
  Program Files or the updater's downloads, and deletes scratch `.ps1` files
  that contain delete loops; drive such experiments from Python.
- **WebView2:** the PDF renderer needs its own user-data folder
  (`%LOCALAPPDATA%\OutlookAI\WebView2PdfData`); sharing the chat's fails with
  HRESULT 0x8007139F.
- **RDS:** `C:\ProgramData\OutlookAI` (shared sign-in, Settings, model list)
  is writable by every signed-in user, an accepted risk documented in
  `Deploy/README.txt`.
