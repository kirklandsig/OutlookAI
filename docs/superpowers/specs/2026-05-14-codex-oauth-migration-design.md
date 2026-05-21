# Phase 1 Design: ChatGPT/Codex OAuth Migration

## Executive Summary

Phase 1 replaces OutlookAI's Anthropic API-key path with embedded ChatGPT/Codex OAuth that uses the user's consumer ChatGPT subscription. The add-in keeps the existing compose-window UX: Proofread, Revise, Shorten, Lengthen, Formal, Friendly, Draft New Email, voice input, and Insert/Replace/Discard. Text generation moves from Anthropic Messages to Codex's ChatGPT-managed Responses backend (`chatgpt.com/backend-api/codex/responses`). Voice input moves from a separate OpenAI Whisper API key to the **OpenAI Realtime WebSocket** (`wss://api.openai.com/v1/realtime`). Both paths authenticate with the same OAuth `access_token` and bill against the user's ChatGPT consumer subscription.

The target deployment remains Windows RDS with Office 2024 and on-prem Exchange. Each RDS server owns one shared OAuth session at `C:\ProgramData\OutlookAI\auth.json`; all interactive users on that server share that OpenAI identity. This is an accepted Phase 1 tradeoff.

## Confirmed Decisions

| Item | Decision |
| --- | --- |
| Phase scope | OAuth/auth/provider migration only |
| Target framework | Keep `.NET Framework 4.7.2`; do not upgrade to 4.8 in Phase 1 |
| Chat provider | ChatGPT Codex backend Responses endpoint |
| Chat model | `gpt-5.5` |
| Voice provider | OpenAI Realtime WebSocket (`wss://api.openai.com/v1/realtime`) using same OAuth `access_token` |
| Voice model | `gpt-realtime-1.5` (per AIReceptionist precedent) |
| MaxTokens default | `65536` |
| Auth mechanism | Embedded browser OAuth + localhost callback + PKCE |
| OAuth client ID | Codex CLI public client: `app_EMoamEEZ73f0CkXaXp7hrann` |
| Consent screen | Accept that OpenAI may show `Codex CLI` |
| OAuth dependency | No Codex CLI, Node.js, or manual `auth.json` copy |
| Auth file | `C:\ProgramData\OutlookAI\auth.json` |
| Auth sharing | Per server, shared by all users on that server |
| Auth ACL | `Authenticated Users: Modify`; accepted risk for trusted RDS user base |
| API-key fallback | None in Phase 1 |
| JSON library | `Newtonsoft.Json 13.0.3` via NuGet |
| HTTP library | `System.Net.Http` BCL reference |
| Settings UI | Account status + Sign In / Sign Out / Refresh; no API key/model/max-token fields |
| Install path | Hardcode `C:\Program Files\OutlookAI`; remove `-InstallPath` parameter |
| Per-user v1 config | Rename to `config.xml.v1.backup.<timestamp>` |
| First implementation task | 30-60 minute OAuth spike gates the rest of implementation; completed against `https://chatgpt.com/backend-api/codex/responses` |

## Architecture

```text
ThisAddIn
  owns singleton CodexAuthService
        |
        | GetAccessTokenAsync() returns ChatGPT OAuth access_token
        |
        +---------------+-------------------+
        v               v                   v
CodexChatService  RealtimeVoiceService  SettingsForm
  POST /backend-     wss://api.openai.    sign in / sign out /
  api/codex/         com/v1/realtime      refresh / status
  responses          ?model=gpt-realtime
        ^               ^
        |               |
        +-------+-------+
                |
            AITaskPane
              compose UX (text actions + mic)
```

`CodexAuthService` is process-wide. It handles sign-in, refresh, atomic token persistence, and status reporting. It exposes `GetAccessTokenAsync()` which returns the ChatGPT OAuth `access_token`, matching Codex's `CodexAuth::Chatgpt` behavior. It does not mint or store an `OPENAI_API_KEY`. Both text and voice services consume the same bearer.

`CodexChatService` replaces `ClaudeService`. It keeps the existing action enum and prompt semantics. It posts SSE-streamed Responses-API requests to `https://chatgpt.com/backend-api/codex/responses` with the OAuth `access_token` and the `ChatGPT-Account-ID` header.

`RealtimeVoiceService` replaces the inline Whisper REST call in `AITaskPane`. It opens a `System.Net.WebSockets.ClientWebSocket` to `wss://api.openai.com/v1/realtime?model=gpt-realtime-1.5` with the OAuth `access_token` as `Authorization: Bearer ...` and the `OpenAI-Beta: realtime=v1` header. It streams microphone PCM frames in and surfaces the final user transcript as a single string for the task pane to drop into the Draft prompt.

`SettingsForm` is extracted from `TaskPane\AITaskPane.cs` into a separate file. It removes legacy API key, model, and max-token controls.

## OAuth Details

| Field | Value |
| --- | --- |
| Issuer | `https://auth.openai.com` |
| Authorize endpoint | `https://auth.openai.com/oauth/authorize` |
| Token endpoint | `https://auth.openai.com/oauth/token` |
| Revoke endpoint | `https://auth.openai.com/oauth/revoke` |
| Redirect primary | `http://localhost:1455/auth/callback` |
| Redirect fallback | `http://localhost:1457/auth/callback` |
| Scopes | `openid profile email offline_access api.connectors.read api.connectors.invoke` |
| PKCE | S256, 64 random bytes |

The OAuth flow starts only when the user clicks Sign In or triggers first-use sign-in. It must not run from `ThisAddIn_Startup`.

The authorization-code flow returns `access_token`, `id_token`, and `refresh_token`. Requests use `Authorization: Bearer <access_token>` and include `ChatGPT-Account-ID` when the ID token contains `https://api.openai.com/auth.chatgpt_account_id`.

Text endpoint (chat actions):

```text
POST https://chatgpt.com/backend-api/codex/responses
Accept: text/event-stream
Authorization: Bearer <access_token>
ChatGPT-Account-ID: <chatgpt_account_id, when present>
```

Voice endpoint (microphone input → transcript):

```text
GET wss://api.openai.com/v1/realtime?model=gpt-realtime-1.5
Authorization: Bearer <access_token>
```

Verified via spike: the GA Realtime API accepts the ChatGPT OAuth `access_token` directly. Do NOT send `OpenAI-Beta: realtime=v1` — the server now returns `beta_api_shape_disabled` for that header.

This is the critical distinction from platform API usage: Phase 1 must not call `https://api.openai.com/v1/chat/completions`, `https://api.openai.com/v1/audio/transcriptions`, or use a token-exchanged `OPENAI_API_KEY`. Text uses the Codex backend; voice uses the Realtime WebSocket. Both bill against the user's ChatGPT consumer subscription, exactly as Codex CLI does for chat and as AIReceptionist does for voice.

## Concurrency And Persistence

`CodexAuthService` must guard both initial sign-in and refresh with a process-wide `SemaphoreSlim(1, 1)`. Multiple compose windows can host multiple task panes; they must not start competing localhost listeners or token refreshes.

Cross-process refresh is protected by `C:\ProgramData\OutlookAI\auth.json.refresh.lock`. Token writes use a temp file and atomic replacement so refresh-token rotation cannot corrupt `auth.json`.

## Config Schema

V2 config at `C:\Program Files\OutlookAI\config.xml`:

```xml
<Config>
  <AdminPassword>admin</AdminPassword>
  <CodexAuthPath>C:\ProgramData\OutlookAI\auth.json</CodexAuthPath>
  <Model>gpt-5.5</Model>
  <MaxTokens>65536</MaxTokens>
</Config>
```

Legacy elements are ignored if encountered: `ApiKey`, `OpenAIApiKey`, `WhisperModel`, `TranscribeModel`, and Claude model names.

Config precedence changes in Phase 1:

- Defaults load first.
- Global config loads next.
- Per-user AppData config may override only `AdminPassword`.
- Server-authoritative fields (`CodexAuthPath`, `Model`, `MaxTokens`) are not overridden by per-user v1 files.

## File Changes

### `VSTO2\OutlookAI\OutlookAI.csproj`

- Keep `<TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>`.
- Add `System.Net.Http` reference.
- Add `Newtonsoft.Json` reference.
- Add compile entries for `Services\CodexAuthService.cs`, `Services\CodexChatService.cs`, `Services\RealtimeVoiceService.cs`, and `SettingsForm.cs`.
- Remove compile entry for `Services\ClaudeServices.cs`.

### `VSTO2\OutlookAI\packages.config`

- Add `Newtonsoft.Json` version `13.0.3` for `net472`.

### `VSTO2\OutlookAI\Config.cs`

- Remove `ApiKey`, `OpenAIApiKey`, `WhisperModel`, `TranscribeModel`, and Claude `AvailableModels`.
- Keep `AdminPassword`.
- Add `CodexAuthPath`, `Model`, and `MaxTokens` defaults.
- Implement v2 load/save semantics and ignore v1 API-key fields.

### `VSTO2\OutlookAI\Services\ClaudeServices.cs`

- Delete. This also removes the current global certificate-validation bypass.

### `VSTO2\OutlookAI\Services\CodexAuthService.cs`

- New auth service.
- Owns OAuth, refresh, token persistence, and auth status.
- Exposes `GetAccessTokenAsync`, `SignInAsync`, `SignOutAsync`, `GetStatus`, and `StatusChanged`.
- Accepts injectable `HttpMessageHandler` for tests.

### `VSTO2\OutlookAI\Services\CodexChatService.cs`

- New ChatGPT Codex service.
- Owns Codex Responses HTTP calls.
- Preserves existing action enum and prompt behavior.
- Uses `CodexAuthService.GetAccessTokenAsync` and posts to `https://chatgpt.com/backend-api/codex/responses`.

### `VSTO2\OutlookAI\Services\RealtimeVoiceService.cs`

- New Realtime voice service. Replaces inline `TranscribeWithWhisper` in `AITaskPane`.
- Wraps a `System.Net.WebSockets.ClientWebSocket` connection to `wss://api.openai.com/v1/realtime?model=gpt-realtime-1.5`.
- Uses `CodexAuthService.GetAccessTokenAsync()` for the bearer.
- Sets header: `Authorization: Bearer <token>`. (Do not send `OpenAI-Beta`; the GA endpoint rejects it with `beta_api_shape_disabled`.)
- API: `Task<string> TranscribeAsync(Stream pcmAudio, CancellationToken ct)` returns the final user transcript.
- Streams `input_audio_buffer.append` frames built from the existing 16-kHz PCM `WaveInEvent` capture.
- Sends `input_audio_buffer.commit` then `response.create` with `modalities: ["text"]` and `input_audio_transcription.model: gpt-4o-mini-transcribe` so the server returns the user transcript without speaking back.
- Listens for `conversation.item.input_audio_transcription.completed` events and resolves with the `transcript` field.
- Closes the socket cleanly after the transcript arrives or on cancellation.

### `VSTO2\OutlookAI\SettingsForm.cs`

- New extracted form.
- Keeps admin password gate.
- Removes API-key/model/max-token fields.
- Adds account status, Sign In, Sign Out, and Refresh.

### `VSTO2\OutlookAI\TaskPane\AITaskPane.cs`

- Replace `_claudeService` with `_codexChatService` (`CodexChatService`).
- Replace all `ClaudeService.ActionType` references with `CodexChatService.ActionType`.
- Delete inline `TranscribeWithWhisper`. Replace its call site with `await _voiceService.TranscribeAsync(pcmStream, ct)` where `_voiceService` is the singleton `RealtimeVoiceService` from `Globals.ThisAddIn`.
- Keep the existing `WaveInEvent` mic capture, but stream PCM frames into the new `RealtimeVoiceService` instead of writing a `.wav` and POSTing it.
- Remove the global `ServicePointManager.ServerCertificateValidationCallback` bypass.
- Await service methods directly; remove `Task.Run(... .Result)`.
- Delete nested `SettingsForm` class.
- Add first-use auth guard before chat and voice operations.

### `VSTO2\OutlookAI\ThisAddIn.cs`

- Add singletons:
  - `public CodexAuthService AuthService { get; private set; }`
  - `public CodexChatService ChatService { get; private set; }`
  - `public RealtimeVoiceService VoiceService { get; private set; }`
- Initialize all three in `ThisAddIn_Startup`. Do not auto-launch OAuth; sign-in is user-triggered.
- Dispose all three in `ThisAddIn_Shutdown`.

### `Ribbon.cs` / `Ribbon.xml`

- No Phase 1 change.

### `Deploy\Install-OutlookAI.ps1`

- Remove `-InstallPath` parameter.
- Back up existing `config.xml` to `C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>` before deleting install files.
- If existing config lacks `<CodexAuthPath>`, write fresh v2 config and carry forward only `AdminPassword`.
- Create `C:\ProgramData\OutlookAI`.
- Grant `Authenticated Users: Modify` to `C:\ProgramData\OutlookAI`.
- Rename per-user `%APPDATA%\OutlookAI\config.xml` files missing `<CodexAuthPath>` to `config.xml.v1.backup.<timestamp>`.
- Print accepted-risk warning about shared OAuth credentials.

### `Deploy\Uninstall-OutlookAI.ps1`

- Remove local auth artifacts: `C:\ProgramData\OutlookAI\auth.json` and `auth.json.refresh.lock`.
- Preserve `C:\ProgramData\OutlookAI\Backups`.
- Do not attempt server-side token revocation in Phase 1.

### `Deploy\README.txt`

- Document shared OAuth credential risk.
- Add runbook for rotating shared auth.

### `VSTO2\OutlookAI.sln`

- New solution file containing the VSTO project and test project.

### `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj`

- New classic `.NET Framework 4.7.2` xUnit test project.
- Tests avoid Outlook COM and WinForms UI dependencies.

## Testing Strategy

Unit tests target pure services and config logic only:

- `CodexAuthServiceTests`: PKCE, authorization-code exchange shape, refresh-token exchange shape, expiry, atomic write, port fallback, single-flight.
- `CodexChatServiceTests`: Codex backend request shape, SSE response parsing, error mapping.
- `RealtimeVoiceServiceTests`: WebSocket request URL, Authorization/OpenAI-Beta headers, session-config payload shape, transcript event parsing. Driven by an injectable WebSocket transport seam (interface around `ClientWebSocket`).
- `ConfigMigrationTests`: v1 config detection, v2 config creation, per-user backup behavior using explicit temp paths.

Verification commands:

```powershell
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\OutlookAI.Tests.dll"
```

Manual smoke checklist:

- Open Outlook compose window.
- AI Assistant ribbon opens task pane.
- First-use sign-in opens browser and captures callback.
- Quick Actions generate text.
- Draft New Email works with and without reply context.
- Voice input records and transcribes.
- Insert, Replace, and Discard behave as before.
- Settings Refresh updates account status.
- Sign Out clears local auth and returns to unsigned-in state.
- Restart Outlook; persistent sign-in works.

## OAuth Spike Gate

Before full implementation, run a 30-60 minute spike to prove:

- Browser launch from an Outlook/VSTO-style host works.
- `HttpListener` can bind to `http://localhost:1455/auth/callback` under the Outlook user context.
- Callback capture works.
- Code exchange produces a valid ChatGPT OAuth `access_token`.
- A minimal `POST /backend-api/codex/responses` with `gpt-5.5` succeeds (verified — text spike returned 200 OK on `chatgpt.com/backend-api/codex/responses`).
- A minimal `wss://api.openai.com/v1/realtime?model=gpt-realtime-1.5` connection accepts the OAuth bearer and emits `session.created`. Precedent: AIReceptionist `receptionist/agent.py:1138-1142` already uses this exact bearer flow against `openai.realtime.RealtimeModel`, so we expect this to succeed.

If `HttpListener` fails, test raw `TcpListener`. If both fail inside Outlook, switch to a small out-of-process helper executable before implementing the rest.

## Rollout

- Deploy to one canary RDS server.
- Have one authorized admin sign in once on that server.
- Run a one-week internal soak.
- Roll out to remaining RDS servers.

## Rollback

- Uninstall v2.
- Reinstall v1.x publish artifacts.
- Restore `C:\Program Files\OutlookAI\config.xml` from `C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>`.

## Accepted Risks

- `Authenticated Users: Modify` means any signed-in RDS user can read, copy, corrupt, delete, or replace the shared OAuth tokens. This is accepted for the trusted RDS user base. If trust changes, rotate the credential and reconsider a dedicated group.
- Consumer ChatGPT quota/rate limits are account-plan limits. OutlookAI must surface ChatGPT backend rate-limit or credit errors as account/subscription issues, not platform API-key quota errors.
- Phase 1 does not perform server-side token revocation on uninstall.
- Consent page may say `Codex CLI`.
- Code signing is deferred.

## Out Of Scope

- Tool/function calling.
- Multi-turn conversation state.
- Outlook Object Model search/filter tools.
- Explorer-window task pane or new Explorer ribbon button.
- Inbox Copilot chat UI.
- Report rendering or export.
- Custom Instructions, Summarize, Translate, Export Images to PDF.
- Dedicated OutlookAI OAuth client registration.
- Code signing.

## Deferred Follow-ups (Phase 1.x)

These items extend the Settings UX once Phase 1 is verified in production.
They are intentionally not in the initial Phase 1 cut because they expand
the admin surface area, but they will land before (or alongside) Phase 2.

### Admin-selectable model + reasoning effort

The Settings dialog admin panel (already gated by `AdminPassword`) gains two
dropdowns:

| Control | Source of options | Persisted to |
| --- | --- | --- |
| **Model** | Static list of consumer-subscription-accessible models (`gpt-5.5`, `gpt-5.5-pro`, `gpt-5.4`, `gpt-5.4-mini`, `gpt-4.1-mini`, `gpt-4.1-nano`, `gpt-5.3-codex`). Source of truth: `codex-rs/skills/src/assets/samples/openai-docs/references/latest-model.md` and the live OpenAI models docs. | `Config.Model` in the **global** `config.xml` (server-authoritative). |
| **Reasoning Effort** | `None` / `Minimal` / `Low` / `Medium` / `High`. Options are filtered per selected model — non-reasoning models (e.g. `gpt-4.1-nano`, `gpt-4.1-mini`) only expose `None`. | `Config.ReasoningEffort` in the **global** `config.xml`. |

Behavior:

- The admin saves both fields once; every user on the server uses those
  values for every chat request. There is no per-user override (consistent
  with the existing server-authoritative precedence for `Model`).
- `CodexChatService.BuildResponsesRequest` switches from
  `"reasoning": null` to `"reasoning": { "effort": "<value>" }` whenever
  `Config.ReasoningEffort` is anything but `None`. When `None`, the
  property is omitted entirely so non-reasoning models accept the request.
- Selecting a model in the dropdown re-filters the reasoning effort
  dropdown and snaps the current selection to a valid value.
- The XML schema grows two elements; legacy configs without them keep the
  current `gpt-5.5` / `None` defaults.

This is the only Phase 1.x follow-up currently planned. Anything beyond it
belongs in Phase 2 (tool calling) or Phase 3 (Inbox Copilot UI).

## Addendum — v2.1.1 precedence change (commit 2dec846)

The original design above stated: *"`Config.ReasoningEffort` in the **global**
`config.xml`. … There is no per-user override (consistent with the existing
server-authoritative precedence for `Model`)."*

That model broke on RDS: the admin-gated `SettingsForm` writes to
`%APPDATA%\OutlookAI\config.xml` (the admin's own profile). Other users on the
same RDS host don't see those changes because their `%APPDATA%` is separate.
As shipped in v2.1.0, an admin setting "Medium" only affected their own future
sessions — not the rest of the server.

v2.1.1 (commit `2dec846`) introduces a third storage layer between global and
per-user:

| Layer | Location | Writable by | Purpose |
| --- | --- | --- | --- |
| 1. Defaults | hardcoded in `Config.cs` | n/a | Last-resort fallback. |
| 2. Global (server-authoritative) | `C:\Program Files\OutlookAI\config.xml` | installer only | Fields the server enforces: `CodexAuthPath`, `VoiceModel`, etc. |
| 3. **Shared admin defaults (new)** | `C:\ProgramData\OutlookAI\config.xml` | Admins (ProgramData ACL) | What the admin chose as the server-wide default for user-tunable fields. |
| 4. Per-user override | `%APPDATA%\OutlookAI\config.xml` | The user | Optional per-user tweak. |

`SaveConfig()` writes user-tunable fields to BOTH layer 3 and layer 4 — the
layer-3 write silently no-ops for non-admin users (ProgramData ACL denies
write), so only admin saves change the shared defaults. Layer 4 always
succeeds, so the saving user's own preference still takes effect immediately.

Load order is 1 → 2 → 3 → 4 with later layers winning. Net effect:

- Admin sets `ReasoningEffort = Medium` on the RDS host → ProgramData gets
  `Medium` → every user's next launch loads `Medium` as their default → users
  can still override to their own value if they want (Layer 4).
- A non-admin user setting their own value silently no-ops the ProgramData
  write (no permission) — only their AppData is updated, only their session
  is affected.

Tests covering this precedence are in `VSTO2/OutlookAI.Tests/ConfigTests.cs`:
`LoadConfigFromPaths_SharedDefaultsAppliedWhenUserAbsent`,
`LoadConfigFromPaths_UserOverridesSharedDefaults`,
`LoadConfigFromPaths_SharedConfigPathNull_TreatedAsAbsent`,
`LoadConfigFromPaths_GlobalAndSharedAndUser_UserBeatsSharedBeatsGlobal`.
