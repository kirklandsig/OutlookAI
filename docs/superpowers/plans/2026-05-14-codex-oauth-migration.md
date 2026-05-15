# ChatGPT/Codex OAuth Migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace OutlookAI's Anthropic/API-key integration with embedded ChatGPT/Codex OAuth and ChatGPT consumer-plan `gpt-5.5` access while preserving the existing Outlook compose-window text UX.

**Architecture:** Add a process-wide `CodexAuthService` that owns OAuth, refresh, and `auth.json` persistence. Replace `ClaudeService` with `CodexChatService`, which calls `https://chatgpt.com/backend-api/codex/responses` with the raw ChatGPT OAuth `access_token`. Keep Outlook/WinForms integration thin and test plain service/config code outside COM.

**Tech Stack:** .NET Framework 4.7.2 VSTO, Windows Forms, `HttpClient`, `Newtonsoft.Json 13.0.3`, xUnit v2, `packages.config`, MSBuild/VSTest on Windows.

---

## Guardrails

- Worktree: `C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-codex-oauth-migration`.
- Branch: `feature/codex-oauth-migration`.
- Do not commit unless the user explicitly asks.
- Production C# behavior changes use red-green-refactor.
- OAuth spike is a user-approved throwaway prototype gate and must not be committed as production code.
- Deployment scripts can affect real machines; test their pure migration behavior with temp paths before manual admin testing.

## File Structure

- Create: `VSTO2\OutlookAI.sln` — solution containing product and tests.
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj` — references, package hint paths, compile includes.
- Modify: `VSTO2\OutlookAI\packages.config` — add `Newtonsoft.Json`.
- Modify: `VSTO2\OutlookAI\Config.cs` — v2 config schema and v1 compatibility behavior.
- Delete: `VSTO2\OutlookAI\Services\ClaudeServices.cs` — old Anthropic service.
- Create: `VSTO2\OutlookAI\Services\CodexAuthService.cs` — OAuth, refresh, token exchange, persistence.
- Create: `VSTO2\OutlookAI\Services\CodexChatService.cs` — ChatGPT Codex backend Responses calls.
- Create: `VSTO2\OutlookAI\Services\RealtimeVoiceService.cs` — OpenAI Realtime WebSocket transcription.
- Create: `VSTO2\OutlookAI\SettingsForm.cs` — account/settings UI extracted from task pane.
- Modify: `VSTO2\OutlookAI\TaskPane\AITaskPane.cs` — service rewiring; replace inline Whisper REST with `RealtimeVoiceService.TranscribeAsync`.
- Modify: `VSTO2\OutlookAI\ThisAddIn.cs` — singleton auth service lifecycle.
- Modify: `Deploy\Install-OutlookAI.ps1` — v1-to-v2 config migration, ProgramData ACL, AppData cleanup.
- Modify: `Deploy\Uninstall-OutlookAI.ps1` — remove local auth artifacts, preserve backups.
- Modify: `Deploy\README.txt` — shared-token risk and rotation runbook.
- Create: `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj` — xUnit test project.
- Create: `VSTO2\OutlookAI.Tests\packages.config` — test dependencies.
- Create: `VSTO2\OutlookAI.Tests\Helpers\FakeHttpMessageHandler.cs` — controllable HTTP seam.
- Create: `VSTO2\OutlookAI.Tests\Services\CodexAuthServiceTests.cs` — auth behavior tests.
- Create: `VSTO2\OutlookAI.Tests\Services\CodexChatServiceTests.cs` — Codex backend request/response tests.
- Create: `VSTO2\OutlookAI.Tests\Services\RealtimeVoiceServiceTests.cs` — Realtime WebSocket request/response tests via injectable transport.
- Create: `VSTO2\OutlookAI.Tests\Helpers\FakeWebSocketTransport.cs` — controllable WebSocket seam.
- Create: `VSTO2\OutlookAI.Tests\ConfigTests.cs` — v2 config behavior tests.
- Create: `VSTO2\OutlookAI.Tests\DeployScriptTests.cs` — temp-path deploy migration tests.

## Baseline Commands

- [x] Restore packages:

```powershell
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
```

- [x] Build baseline:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI\OutlookAI.csproj" /p:Configuration=Debug /p:Platform="AnyCPU"
```

Expected baseline result: `Build succeeded. 0 Warning(s) 0 Error(s)`.

---

### Task 1: OAuth Spike Gate

**Files:**
- Create outside repo: `C:\Users\MDASR\AppData\Local\Temp\opencode\outlookai-oauth-spike\CodexOAuthSpike.csproj`
- Create outside repo: `C:\Users\MDASR\AppData\Local\Temp\opencode\outlookai-oauth-spike\Program.cs`

- [ ] **Step 1: Create throwaway console project**

Run:

```powershell
Test-Path -LiteralPath "C:\Users\MDASR\AppData\Local\Temp\opencode"
if (!(Test-Path -LiteralPath "C:\Users\MDASR\AppData\Local\Temp\opencode\outlookai-oauth-spike")) { New-Item -ItemType Directory -Path "C:\Users\MDASR\AppData\Local\Temp\opencode\outlookai-oauth-spike" | Out-Null }
dotnet new console --framework net472 --output "C:\Users\MDASR\AppData\Local\Temp\opencode\outlookai-oauth-spike"
```

Expected: project is created outside the repo.

- [ ] **Step 2: Implement spike with browser launch and localhost callback**

Use `HttpListener` on `http://localhost:1455/auth/callback/`, generate PKCE S256, open the authorize URL, capture `code` and `state`, exchange code at `https://auth.openai.com/oauth/token`, then call `POST https://chatgpt.com/backend-api/codex/responses` with the returned OAuth `access_token`, model `gpt-5.5`, and prompt `Say OAuth spike success in five words.`.

The spike must print:

```text
Callback received: true
Token exchange status: OK
ChatGPT Codex backend status: OK
```

- [ ] **Step 3: Run spike**

Run:

```powershell
dotnet run --project "C:\Users\MDASR\AppData\Local\Temp\opencode\outlookai-oauth-spike\CodexOAuthSpike.csproj"
```

Expected: browser sign-in succeeds and all three status lines print.

- [ ] **Step 4: Decide on listener implementation**

If `HttpListener` fails to bind under the Outlook user context, repeat Step 2 using raw `TcpListener` and keep that approach for `CodexAuthService`. If both listeners fail, stop and design a helper executable before touching product code.

---

### Task 2: Test Project And Package Scaffolding

**Files:**
- Create: `VSTO2\OutlookAI.sln`
- Create: `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj`
- Create: `VSTO2\OutlookAI.Tests\packages.config`
- Create: `VSTO2\OutlookAI.Tests\Helpers\FakeHttpMessageHandler.cs`
- Modify: `VSTO2\OutlookAI\packages.config`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

- [ ] **Step 1: Add failing test helper compile target**

Create `VSTO2\OutlookAI.Tests\Helpers\FakeHttpMessageHandler.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Tests.Helpers
{
    public sealed class FakeHttpMessageHandler : HttpMessageHandler
    {
        private readonly Queue<Func<HttpRequestMessage, HttpResponseMessage>> _responses = new Queue<Func<HttpRequestMessage, HttpResponseMessage>>();

        public List<HttpRequestMessage> Requests { get; } = new List<HttpRequestMessage>();

        public void QueueJson(HttpStatusCode statusCode, string json)
        {
            _responses.Enqueue(_ => new HttpResponseMessage(statusCode)
            {
                Content = new StringContent(json)
            });
        }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests.Add(request);
            if (_responses.Count == 0)
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.InternalServerError)
                {
                    Content = new StringContent("{\"error\":\"no fake response queued\"}")
                });
            }

            return Task.FromResult(_responses.Dequeue()(request));
        }
    }
}
```

- [ ] **Step 2: Create test project files**

Create `VSTO2\OutlookAI.Tests\packages.config`:

```xml
<?xml version="1.0" encoding="utf-8"?>
<packages>
  <package id="xunit" version="2.9.3" targetFramework="net472" />
  <package id="xunit.abstractions" version="2.0.3" targetFramework="net472" />
  <package id="xunit.assert" version="2.9.3" targetFramework="net472" />
  <package id="xunit.core" version="2.9.3" targetFramework="net472" />
  <package id="xunit.extensibility.core" version="2.9.3" targetFramework="net472" />
  <package id="xunit.extensibility.execution" version="2.9.3" targetFramework="net472" />
  <package id="xunit.runner.visualstudio" version="3.1.5" targetFramework="net472" />
  <package id="Microsoft.NET.Test.Sdk" version="17.14.1" targetFramework="net472" />
  <package id="Newtonsoft.Json" version="13.0.3" targetFramework="net472" />
</packages>
```

Create `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj` as a classic C# library targeting `v4.7.2`, referencing `System`, `System.Core`, `System.Net.Http`, `Newtonsoft.Json`, xUnit assemblies under `packages`, and the product project `..\OutlookAI\OutlookAI.csproj`.

- [ ] **Step 3: Add solution file**

Run:

```powershell
dotnet new sln --name OutlookAI --output VSTO2
dotnet sln VSTO2\OutlookAI.sln add VSTO2\OutlookAI\OutlookAI.csproj
dotnet sln VSTO2\OutlookAI.sln add VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
```

Expected: `VSTO2\OutlookAI.sln` lists both projects.

- [ ] **Step 4: Restore and verify test scaffold**

Run:

```powershell
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI.Tests\packages.config" -PackagesDirectory "VSTO2\packages"
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"
```

Expected: test project compiles with no product behavior changes.

---

### Task 3: Config V2 Behavior

**Files:**
- Test: `VSTO2\OutlookAI.Tests\ConfigTests.cs`
- Modify: `VSTO2\OutlookAI\Config.cs`

- [ ] **Step 1: Write failing tests for v2 defaults and v1 ignore behavior**

Create `VSTO2\OutlookAI.Tests\ConfigTests.cs`:

```csharp
using System.IO;
using OutlookAI;
using Xunit;

namespace OutlookAI.Tests
{
    public class ConfigTests
    {
        [Fact]
        public void LoadConfigFromPaths_UsesV2DefaultsWhenFilesAreMissing()
        {
            var dir = Path.Combine(Path.GetTempPath(), "outlookai-config-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);

            Config.LoadConfigForTest(Path.Combine(dir, "global.xml"), Path.Combine(dir, "user.xml"));

            Assert.Equal("admin", Config.AdminPassword);
            Assert.Equal(@"C:\ProgramData\OutlookAI\auth.json", Config.CodexAuthPath);
            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal(65536, Config.MaxTokens);
        }

        [Fact]
        public void LoadConfigFromPaths_DoesNotLetPerUserV1OverrideServerFields()
        {
            var dir = Path.Combine(Path.GetTempPath(), "outlookai-config-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);
            var global = Path.Combine(dir, "global.xml");
            var user = Path.Combine(dir, "user.xml");

            File.WriteAllText(global, "<Config><AdminPassword>server</AdminPassword><CodexAuthPath>C:\\ProgramData\\OutlookAI\\auth.json</CodexAuthPath><Model>gpt-5.5</Model><MaxTokens>65536</MaxTokens></Config>");
            File.WriteAllText(user, "<Config><ApiKey>anthropic</ApiKey><OpenAIApiKey>openai</OpenAIApiKey><AdminPassword>userpass</AdminPassword><Model>claude-opus-4-6</Model><WhisperModel>whisper-1</WhisperModel><MaxTokens>2048</MaxTokens></Config>");

            Config.LoadConfigForTest(global, user);

            Assert.Equal("userpass", Config.AdminPassword);
            Assert.Equal(@"C:\ProgramData\OutlookAI\auth.json", Config.CodexAuthPath);
            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal(65536, Config.MaxTokens);
        }
    }
}
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"
```

Expected: compile fails because `Config.LoadConfigForTest` and `CodexAuthPath` do not exist.

- [ ] **Step 3: Implement minimal Config v2 API**

Replace legacy fields in `Config.cs` with:

```csharp
public static string AdminPassword { get; set; } = "admin";
public static string CodexAuthPath { get; set; } = @"C:\ProgramData\OutlookAI\auth.json";
public static string Model { get; set; } = "gpt-5.5";
public static int MaxTokens { get; set; } = 65536;

public static void LoadConfigForTest(string globalPath, string userPath)
{
    ResetDefaults();
    LoadFromFile(globalPath, allowServerFields: true);
    LoadFromFile(userPath, allowServerFields: false);
}
```

Use `LoadFromFile(string filePath, bool allowServerFields)` so per-user files can change only `AdminPassword`.

- [ ] **Step 4: Run tests and verify GREEN**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\OutlookAI.Tests.dll"
```

Expected: both config tests pass.

---

### Task 4: CodexAuthService Core

**Files:**
- Test: `VSTO2\OutlookAI.Tests\Services\CodexAuthServiceTests.cs`
- Create: `VSTO2\OutlookAI\Services\CodexAuthService.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

- [ ] **Step 1: Write failing tests for authorization-code exchange request**

Create a test that constructs `CodexAuthService` with `FakeHttpMessageHandler`, queues an authorization-code JSON response containing `access_token`, `id_token`, and `refresh_token`, calls `ExchangeCodeForTokensForTestAsync("code", "redirect", "verifier")`, and asserts the request body contains:

```text
grant_type=authorization_code
client_id=app_EMoamEEZ73f0CkXaXp7hrann
code=code
redirect_uri=redirect
code_verifier=verifier
```

- [ ] **Step 2: Run and verify RED**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"
```

Expected: compile fails because `CodexAuthService` does not exist.

- [ ] **Step 3: Implement minimal token exchange**

Create `CodexAuthService.cs` with constants for OAuth endpoints, constructor `CodexAuthService(string authPath, HttpMessageHandler handler = null)`, and method:

```csharp
public async Task<AuthTokens> ExchangeCodeForTokensForTestAsync(string code, string redirectUri, string codeVerifier, CancellationToken cancellationToken = default(CancellationToken))
{
    return await ExchangeCodeForTokensAsync(code, redirectUri, codeVerifier, cancellationToken).ConfigureAwait(false);
}
```

`ExchangeCodeForTokensAsync` posts URL-encoded form data to `https://auth.openai.com/oauth/token` and parses `access_token`, `id_token`, and `refresh_token` using Newtonsoft.Json.

- [ ] **Step 4: Run and verify GREEN**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\OutlookAI.Tests.dll"
```

Expected: authorization-code exchange test passes.

- [ ] **Step 5: Add tests for persistence and status**

Add tests proving `SignOutAsync` deletes auth file, `GetStatus` reports unauthenticated with no file, and expired bearer triggers refresh using queued fake responses.

- [ ] **Step 6: Implement persistence and status**

Add `AuthStatus`, `StatusChanged`, `SignOutAsync`, `GetStatus`, atomic write, and refresh lock path behavior. Use a sidecar lock file named `auth.json.refresh.lock` next to auth.json.

- [ ] **Step 7: Run full test project**

Run build and `vstest.console.exe`. Expected: all service/config tests pass.

---

### Task 5: CodexChatService Core

**Files:**
- Test: `VSTO2\OutlookAI.Tests\Services\CodexChatServiceTests.cs`
- Create: `VSTO2\OutlookAI\Services\CodexChatService.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

- [ ] **Step 1: Write failing Codex backend request test**

Test behavior:

```csharp
[Fact]
public async Task ProcessEmailAsync_SendsCodexResponsesRequestWithConfiguredModelAndBearer()
{
    var handler = new FakeHttpMessageHandler();
    handler.QueueText(HttpStatusCode.OK, "data: {\"type\":\"response.output_text.delta\",\"delta\":\"fixed email\"}\n\ndata: {\"type\":\"response.completed\"}\n\n");
    var auth = new FakeAuthService("chatgpt-access-token", "account-123");
    var service = new CodexChatService(auth, handler);

    var result = await service.ProcessEmailAsync(CodexChatService.ActionType.Proofread, "helo world");

    Assert.Equal("fixed email", result);
    Assert.Equal("Bearer chatgpt-access-token", handler.Requests[0].Headers.Authorization.ToString());
    Assert.Equal("account-123", handler.Requests[0].Headers.GetValues("ChatGPT-Account-ID").Single());
    Assert.Equal("https://chatgpt.com/backend-api/codex/responses", handler.Requests[0].RequestUri.ToString());
}
```

Use a small test-only fake auth provider that returns `chatgpt-access-token` and `account-123`.

- [ ] **Step 2: Verify RED**

Run MSBuild. Expected: compile fails because `CodexChatService` does not exist.

- [ ] **Step 3: Implement minimal Codex responses call**

Create `CodexChatService.cs` with `ActionType`, `ProcessEmailAsync`, copied prompts from `ClaudeService`, JSON body matching Codex `ResponsesApiRequest` (`model`, `instructions`, `input`, `tools`, `tool_choice`, `parallel_tool_calls`, `store`, `stream`, `include`), and SSE parsing for output text deltas.

- [ ] **Step 4: Verify GREEN**

Run build and tests. Expected: chat request test passes.

- [ ] **Step 5: Add error mapping tests**

Queue `429` with `{"detail":"usage limit reached"}` and assert the thrown exception message contains `usage limit reached` and identifies it as a ChatGPT/Codex backend error.

- [ ] **Step 6: Implement error mapping**

For non-success responses, read response body and throw `Exception("ChatGPT Codex Error: " + body)`.

---

### Task 5b: RealtimeVoiceService Core (voice via OAuth)

**Files:**
- Test: `VSTO2\OutlookAI.Tests\Services\RealtimeVoiceServiceTests.cs`
- Test: `VSTO2\OutlookAI.Tests\Helpers\FakeWebSocketTransport.cs`
- Create: `VSTO2\OutlookAI\Services\RealtimeVoiceService.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

Precedent: `C:\Users\MDASR\Desktop\Projects\AIReceptionist\receptionist\agent.py:1138-1142` passes the Codex OAuth `access_token` straight into `openai.realtime.RealtimeModel` and connects to `wss://api.openai.com/v1/realtime`. We replicate the same auth pattern in .NET 4.7.2 using `System.Net.WebSockets.ClientWebSocket`.

- [ ] **Step 1: Define injectable WebSocket transport seam**

Create `IRealtimeWebSocket` with `ConnectAsync(Uri url, IDictionary<string,string> headers, CancellationToken ct)`, `SendAsync(byte[] payload, WebSocketMessageType type, bool endOfMessage, CancellationToken ct)`, `ReceiveAsync(ArraySegment<byte> buffer, CancellationToken ct)`, `CloseAsync(WebSocketCloseStatus status, string description, CancellationToken ct)`, and `Dispose()`. Production implementation wraps `ClientWebSocket`.

- [ ] **Step 2: Write failing connect-headers test**

```csharp
[Fact]
public async Task TranscribeAsync_OpensRealtimeWebSocketWithBearerAndBetaHeader()
{
    var fake = new FakeWebSocketTransport();
    fake.QueueIncomingJson("{\"type\":\"session.created\",\"session\":{\"id\":\"sess_1\"}}");
    fake.QueueIncomingJson("{\"type\":\"conversation.item.input_audio_transcription.completed\",\"transcript\":\"hello world\"}");
    var auth = new FakeAuthService("chatgpt-access-token", "account-123");
    var service = new RealtimeVoiceService(auth, () => fake);

    using (var pcm = new MemoryStream(new byte[3200]))
    {
        var transcript = await service.TranscribeAsync(pcm, CancellationToken.None);
        Assert.Equal("hello world", transcript);
    }

    Assert.Equal(new Uri("wss://api.openai.com/v1/realtime?model=gpt-realtime-1.5"), fake.LastConnectUri);
    Assert.Equal("Bearer chatgpt-access-token", fake.LastConnectHeaders["Authorization"]);
    Assert.False(fake.LastConnectHeaders.ContainsKey("OpenAI-Beta"), "GA endpoint rejects OpenAI-Beta header");
}
```

- [ ] **Step 3: Verify RED**

Run MSBuild. Expected: compile fails because `RealtimeVoiceService` does not exist.

- [ ] **Step 4: Implement minimal RealtimeVoiceService**

Create `RealtimeVoiceService.cs`. Constructor takes `CodexAuthService auth` and `Func<IRealtimeWebSocket> transportFactory`. `TranscribeAsync(Stream pcm, CancellationToken ct)`:

1. Get bearer via `auth.GetAccessTokenAsync(ct)`.
2. Open transport at `wss://api.openai.com/v1/realtime?model=gpt-realtime-1.5` with header `Authorization: Bearer <token>`. Do not send `OpenAI-Beta: realtime=v1` — the GA endpoint rejects it with `beta_api_shape_disabled`.
3. Wait for `session.created`.
4. Send `session.update` with `{"type":"session.update","session":{"input_audio_format":"pcm16","input_audio_transcription":{"model":"gpt-4o-mini-transcribe"},"turn_detection":null,"modalities":["text"]}}`.
5. Read PCM from the stream in 16KB chunks; for each chunk send `{"type":"input_audio_buffer.append","audio":"<base64>"}`.
6. Send `{"type":"input_audio_buffer.commit"}`.
7. Send `{"type":"response.create","response":{"modalities":["text"]}}`.
8. Read messages until a `conversation.item.input_audio_transcription.completed` event arrives; return its `transcript` field.
9. Close socket cleanly.

- [ ] **Step 5: Verify GREEN**

Run build and tests. Expected: connect-headers test passes.

- [ ] **Step 6: Add transcript-event parsing test**

Queue an event without `transcript` and assert the service throws a clear "no transcript returned" error so the task pane can surface a friendly message.

- [ ] **Step 7: Add error-mapping test**

Queue `{"type":"error","error":{"message":"insufficient subscription"}}` and assert the thrown exception message contains the server's error message.

- [ ] **Step 8: Implement `ClientWebSocket`-backed production transport**

`ClientWebSocketRealtimeTransport` wraps `System.Net.WebSockets.ClientWebSocket`. It sets `ClientWebSocket.Options.SetRequestHeader("Authorization", header)` and `SetRequestHeader("OpenAI-Beta", "realtime=v1")` before `ConnectAsync`. Add a smoke test (skipped by default; tagged `[Trait("Category","Integration")]`) that opens a real WebSocket and asserts `session.created` arrives. Run only when an `OUTLOOKAI_REALTIME_INTEGRATION` env var is set so CI doesn't burn quota.

---

### Task 6: Product Project Wiring

**Files:**
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`
- Modify: `VSTO2\OutlookAI\packages.config`
- Delete: `VSTO2\OutlookAI\Services\ClaudeServices.cs`

- [ ] **Step 1: Add Newtonsoft.Json and System.Net.Http**

Add to product `packages.config`:

```xml
<package id="Newtonsoft.Json" version="13.0.3" targetFramework="net472" />
```

Add to `OutlookAI.csproj` references:

```xml
<Reference Include="Newtonsoft.Json, Version=13.0.0.0, Culture=neutral, PublicKeyToken=30ad4fe6b2a6aeed, processorArchitecture=MSIL">
  <HintPath>packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll</HintPath>
</Reference>
<Reference Include="System.Net.Http" />
```

- [ ] **Step 2: Replace compile entries**

Add compile entries for `Services\CodexAuthService.cs`, `Services\CodexChatService.cs`, `Services\RealtimeVoiceService.cs`, and `SettingsForm.cs`. Remove `Services\ClaudeServices.cs`.

- [ ] **Step 3: Delete Claude service**

Delete `VSTO2\OutlookAI\Services\ClaudeServices.cs` after all prompt text has been copied into `CodexChatService`.

- [ ] **Step 4: Build product**

Run package restore and MSBuild. Expected failure at this point may mention task pane references to `ClaudeService`; Task 7 resolves those references.

---

### Task 7: Task Pane And Settings UI

**Files:**
- Create: `VSTO2\OutlookAI\SettingsForm.cs`
- Modify: `VSTO2\OutlookAI\TaskPane\AITaskPane.cs`

- [ ] **Step 1: Extract SettingsForm**

Move the nested `SettingsForm` class from `AITaskPane.cs` into `SettingsForm.cs` under namespace `OutlookAI`. Keep admin password login, remove API key/model/max token controls, and add `Sign In`, `Sign Out`, `Refresh`, and account status label.

- [ ] **Step 2: Rewire task pane actions**

Change field and constructor:

```csharp
private readonly CodexChatService _codexChatService;

public AITaskPane()
{
    InitializeComponent();
    _codexChatService = new CodexChatService(Globals.ThisAddIn.AuthService);
}
```

Replace all `ClaudeService.ActionType` references with `CodexChatService.ActionType`.

- [ ] **Step 3: Replace blocking async call**

Change processing call to:

```csharp
string result = await _codexChatService.ProcessEmailAsync(action, emailContent, prompt);
```

- [ ] **Step 4: Replace Whisper REST with RealtimeVoiceService**

Delete `TranscribeWithWhisper` and the inline multipart POST to `https://api.openai.com/v1/audio/transcriptions`. Replace with:

```csharp
using (var pcm = OpenPcmStreamFromWaveCapture(audioFile))
{
    string transcript = await Globals.ThisAddIn.VoiceService.TranscribeAsync(pcm, ct);
    txtDraftPrompt.Text = transcript;
}
```

The mic button still uses the existing `WaveInEvent` (16kHz, 16-bit, mono) capture but the captured PCM is streamed to `RealtimeVoiceService` instead of being POSTed as a `.wav`. Surface server errors via the existing `lblStatus` label.

- [ ] **Step 5: Build product**

Run MSBuild. Expected: no `ClaudeService`, `Config.OpenAIApiKey`, `Config.ApiKey`, or `Config.WhisperModel` references remain.

---

### Task 8: ThisAddIn Service Lifecycle

**Files:**
- Modify: `VSTO2\OutlookAI\ThisAddIn.cs`

- [ ] **Step 1: Add service singletons**

Add properties:

```csharp
public CodexAuthService AuthService { get; private set; }
public CodexChatService ChatService { get; private set; }
public RealtimeVoiceService VoiceService { get; private set; }
```

Initialize in `ThisAddIn_Startup`:

```csharp
AuthService = new CodexAuthService(Config.CodexAuthPath);
ChatService = new CodexChatService(AuthService);
VoiceService = new RealtimeVoiceService(AuthService, () => new ClientWebSocketRealtimeTransport());
```

Dispose in shutdown:

```csharp
VoiceService?.Dispose();
ChatService?.Dispose();
AuthService?.Dispose();
```

- [ ] **Step 2: Build product**

Run MSBuild. Expected: startup wiring compiles and does not launch OAuth automatically.

---

### Task 9: Deploy Script Migration

**Files:**
- Test: `VSTO2\OutlookAI.Tests\DeployScriptTests.cs`
- Modify: `Deploy\Install-OutlookAI.ps1`
- Modify: `Deploy\Uninstall-OutlookAI.ps1`
- Modify: `Deploy\README.txt`

- [ ] **Step 1: Write failing text-level tests for installer invariants**

Create tests that read `Deploy\Install-OutlookAI.ps1` and assert it contains:

```text
C:\ProgramData\OutlookAI\Backups
config.xml.v1.backup.
CodexAuthPath
gpt-5.5
Authenticated Users
```

Also assert it does not contain:

```text
param(
[string]$InstallPath
<ApiKey>
<OpenAIApiKey>
```

- [ ] **Step 2: Verify RED**

Run VSTest. Expected: deploy script invariant tests fail against current installer.

- [ ] **Step 3: Update installer**

Remove `-InstallPath`, hardcode `$InstallPath = "C:\Program Files\OutlookAI"`, add external config backup, write v2 config when `CodexAuthPath` is absent, create ProgramData folder, grant `Authenticated Users: Modify`, and rename per-user v1 AppData configs to timestamped backups.

- [ ] **Step 4: Update uninstaller and deploy README**

Uninstaller removes local auth artifacts but preserves `C:\ProgramData\OutlookAI\Backups`. README documents shared token risk and credential rotation steps.

- [ ] **Step 5: Verify GREEN**

Run VSTest. Expected: deploy script invariant tests pass.

---

### Task 10: End-to-End Verification

**Files:**
- Modify: `handoff.md`

- [ ] **Step 1: Restore packages**

Run:

```powershell
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI.Tests\packages.config" -PackagesDirectory "VSTO2\packages"
```

- [ ] **Step 2: Build solution**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU"
```

Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`.

- [ ] **Step 3: Run tests**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\OutlookAI.Tests.dll"
```

Expected: all tests pass.

- [ ] **Step 4: Search for removed legacy references**

Run:

```powershell
rg "ClaudeService|ApiKey|OpenAIApiKey|WhisperModel|claude-" VSTO2\OutlookAI Deploy
```

Expected: no product or deploy references remain except historical text in docs if intentionally retained.

- [ ] **Step 5: Update handoff**

Append a concise entry pointing to:

```text
docs/superpowers/specs/2026-05-14-codex-oauth-migration-design.md
docs/superpowers/plans/2026-05-14-codex-oauth-migration.md
```

Mention the worktree path and that commits have not been created.

---

### Task 11: Admin-selectable model + reasoning effort (Phase 1.x)

**Scope:** Restore a single Settings dropdown for `Model` and add a new
dropdown for `ReasoningEffort`. Both persist to the global config and apply
to every user on the server. Filters reasoning options per model (e.g.
`gpt-4.1-nano` only allows `None`).

**Files:**
- Modify: `VSTO2\OutlookAI\Config.cs` — add `ReasoningEffort` property + `AvailableModels`/`AvailableReasoningEfforts` constants and per-model filter.
- Modify: `VSTO2\OutlookAI\SettingsForm.cs` — add `ComboBox` for model + `ComboBox` for reasoning effort under the Account group, gated by admin auth.
- Modify: `VSTO2\OutlookAI\Services\CodexChatService.cs` — emit `"reasoning": { "effort": "<value>" }` in the Responses body when effort is not `None`; omit otherwise.
- Modify: `Deploy\Install-OutlookAI.ps1` — extend the v2 `config.xml` template to include `<ReasoningEffort>None</ReasoningEffort>` (preserved if already set).

- [ ] **Step 1: Add `ReasoningEffort` enum + Config field**

In `Config.cs`:

```csharp
public enum ReasoningEffort { None, Minimal, Low, Medium, High }

public const string DefaultReasoningEffort = "None";

public static string ReasoningEffort { get; set; } = DefaultReasoningEffort;

public static readonly string[] AvailableModels =
{
    "gpt-5.5",
    "gpt-5.5-pro",
    "gpt-5.4",
    "gpt-5.4-mini",
    "gpt-4.1-mini",
    "gpt-4.1-nano",
    "gpt-5.3-codex"
};

public static string[] GetReasoningEffortsForModel(string model)
{
    // Non-reasoning models accept only "None".
    if (model == "gpt-4.1-mini" || model == "gpt-4.1-nano")
    {
        return new[] { "None" };
    }
    return new[] { "None", "Minimal", "Low", "Medium", "High" };
}
```

Update `LoadFromFile` to read `<ReasoningEffort>` from the global section
only (server-authoritative). Update `ResetDefaults` to set
`ReasoningEffort = DefaultReasoningEffort`.

- [ ] **Step 2: Wire dropdowns into SettingsForm**

In the Account group, under the existing Sign In / Sign Out / Refresh row,
add:

```csharp
var lblModel = new Label { Text = "Model:", AutoSize = true };
var cboModel = new ComboBox {
    DropDownStyle = ComboBoxStyle.DropDownList,
    Width = 200
};
cboModel.Items.AddRange(Config.AvailableModels);
cboModel.SelectedItem = Config.Model;

var lblEffort = new Label { Text = "Reasoning effort:", AutoSize = true };
var cboEffort = new ComboBox {
    DropDownStyle = ComboBoxStyle.DropDownList,
    Width = 200
};
cboModel.SelectedIndexChanged += (s, e) => {
    var opts = Config.GetReasoningEffortsForModel((string)cboModel.SelectedItem);
    cboEffort.Items.Clear();
    cboEffort.Items.AddRange(opts);
    cboEffort.SelectedItem = opts.Contains(Config.ReasoningEffort) ? Config.ReasoningEffort : opts[0];
};
```

Persist on **Save**:

```csharp
Config.Model = (string)cboModel.SelectedItem;
Config.ReasoningEffort = (string)cboEffort.SelectedItem;
Config.SaveConfig();
```

Note: Phase 1's `SaveConfig` only writes the per-user file (AdminPassword
only). Persisting `Model` and `ReasoningEffort` requires writing the global
config in `C:\Program Files\OutlookAI\config.xml`. This requires elevated
write access. Two options:

- (a) Have `SaveConfig` attempt to write the global file when the admin
  has changed `Model`/`ReasoningEffort`; show a clear error if not running
  elevated (and ask the user to relaunch Outlook elevated for that save).
- (b) Persist these as per-user overrides in AppData and special-case them
  in `LoadFromFile` (`allowServerFields: true` for these specific keys).

Pick (b) — it avoids forcing elevated Outlook, matches normal Office UX
expectations, and the admin password gate already bounds who can change it.

- [ ] **Step 3: Emit `reasoning.effort` in chat requests**

In `CodexChatService.BuildResponsesRequest`:

```csharp
JToken reasoning = JValue.CreateNull();
if (!string.Equals(Config.ReasoningEffort, "None", StringComparison.OrdinalIgnoreCase))
{
    reasoning = new JObject(
        new JProperty("effort", Config.ReasoningEffort.ToLowerInvariant()));
}

return new JObject(
    new JProperty("model", Config.Model),
    new JProperty("instructions", instructions ?? ""),
    new JProperty("input", new JArray(/* ... */)),
    new JProperty("tools", new JArray()),
    new JProperty("tool_choice", "auto"),
    new JProperty("parallel_tool_calls", false),
    new JProperty("reasoning", reasoning),
    new JProperty("store", false),
    new JProperty("stream", true),
    new JProperty("include", new JArray()));
```

- [ ] **Step 4: Update install script's v2 config template**

In `Deploy\Install-OutlookAI.ps1`, extend the `$v2Config` here-string:

```xml
<Config>
  <AdminPassword>$preservedAdminPassword</AdminPassword>
  <CodexAuthPath>C:\ProgramData\OutlookAI\auth.json</CodexAuthPath>
  <Model>gpt-5.5</Model>
  <VoiceModel>gpt-realtime-1.5</VoiceModel>
  <ReasoningEffort>None</ReasoningEffort>
  <MaxTokens>65536</MaxTokens>
</Config>
```

Preserve any previously-set `<ReasoningEffort>` from the most recent
backup file the same way `AdminPassword` is preserved.

- [ ] **Step 5: Build and verify**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    "VSTO2\OutlookAI\OutlookAI.csproj" /p:Configuration=Debug /p:Platform="AnyCPU"
```

Expected: clean build. Manual smoke: open Settings, change the dropdowns,
save, restart Outlook, confirm the next chat request shows the new model
and reasoning effort in fiddler / packet capture.
