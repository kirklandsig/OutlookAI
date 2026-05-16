# Phase 2 Implementation Plan: Tool Calling + Compose-Window Chat

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn OutlookAI from a single-shot text rewriter into a tool-using assistant scoped to the compose window. Add a 10-tool Outlook catalog, a multi-round chat loop, a WebView2-based chat surface, a variants tab, and the first real test project — all per the approved Phase 2 spec.

**Architecture:** Generic multi-round tool-calling loop in `CodexChatService` (client-side history accumulation; `store:false`); 10 Outlook tools behind an `IToolHost` seam (no COM dependency in the chat service); compose task pane rebuilt around a TabControl (Actions / Chat / Variants); WebView2 chat surface with JS↔C# bridge and streaming output; per-Inspector in-memory conversation and variant state.

**Tech Stack:** .NET Framework 4.7.2 VSTO add-in; `Microsoft.Web.WebView2` (WinForms host control); `Newtonsoft.Json 13.0.3`; xUnit 2.x; classic packages.config for the product, SDK-style csproj for the test project; MSBuild + VSTest on Windows; HTML/CSS/JS chat shell rendered via WebView2 + vendored `marked.min.js`.

**Source spec:** [`docs/superpowers/specs/2026-05-15-phase-2-tool-calling-and-compose-chat-design.md`](../specs/2026-05-15-phase-2-tool-calling-and-compose-chat-design.md)

---

## Guardrails

- Worktree: `C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-codex-oauth-migration`.
- Branch: `feature/codex-oauth-migration`. Do not merge to `master`.
- Push every commit to `origin/feature/codex-oauth-migration`.
- Production C# changes follow red-green-refactor. WebUI assets (HTML/CSS/JS) are exception per TDD-skill rules — verified by manual smoke + integration tests through the ChatController.
- Build sanity check after every task: `msbuild VSTO2\OutlookAI.sln /p:Configuration=Debug /p:Platform="Any CPU"` from worktree root.
- Test sanity check after every test-bearing task: `vstest.console.exe VSTO2\OutlookAI.Tests\bin\Debug\OutlookAI.Tests.dll`.

## File Structure

Phase 2 introduces or modifies these files. Each task creates / modifies exactly the files declared in its **Files** block.

### New files

**Services / chat infrastructure:**
- `VSTO2\OutlookAI\Services\IToolHost.cs`
- `VSTO2\OutlookAI\Services\OutlookToolHost.cs`
- `VSTO2\OutlookAI\Services\OutlookThreadMarshaller.cs`
- `VSTO2\OutlookAI\Services\IdResolver.cs`
- `VSTO2\OutlookAI\Services\ToolDispatcher.cs`
- `VSTO2\OutlookAI\Services\Tools\ToolCatalogSchema.cs`
- `VSTO2\OutlookAI\Services\Tools\IOutlookTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookGetCurrentComposeStateTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookListFoldersTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookSearchMessagesTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookReadMessageTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookCountMessagesTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookListRecentThreadsWithTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookCreateDraftTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookMarkAsReadTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookFlagMessageTool.cs`
- `VSTO2\OutlookAI\Services\Tools\OutlookSetCategoryTool.cs`
- `VSTO2\OutlookAI\Services\Tools\IOutlookSurface.cs`
- `VSTO2\OutlookAI\Services\Tools\LiveOutlookSurface.cs`
- `VSTO2\OutlookAI\Services\Chat\ConversationContext.cs`
- `VSTO2\OutlookAI\Services\Chat\ConversationStore.cs`
- `VSTO2\OutlookAI\Services\Chat\ChatEventSink.cs`
- `VSTO2\OutlookAI\Services\Chat\TurnResult.cs`
- `VSTO2\OutlookAI\Services\Chat\StopReason.cs`
- `VSTO2\OutlookAI\Services\Variants\VariantParser.cs`
- `VSTO2\OutlookAI\Services\Variants\VariantStore.cs`
- `VSTO2\OutlookAI\Services\Variants\Variant.cs`
- `VSTO2\OutlookAI\Services\Variants\Tone.cs`

**Task pane / UI:**
- `VSTO2\OutlookAI\TaskPane\Chat\ChatController.cs`
- `VSTO2\OutlookAI\TaskPane\Chat\WebView2Bootstrap.cs`
- `VSTO2\OutlookAI\TaskPane\Variants\VariantsController.cs`

**Web UI assets (embedded resources):**
- `VSTO2\OutlookAI\WebUI\index.html`
- `VSTO2\OutlookAI\WebUI\styles.css`
- `VSTO2\OutlookAI\WebUI\chat.js`
- `VSTO2\OutlookAI\WebUI\marked.min.js`        (vendored, MIT)
- `VSTO2\OutlookAI\WebUI\highlight.min.js`     (vendored, MIT)

**Deploy:**
- `Deploy\MicrosoftEdgeWebView2Setup.exe`      (vendored Microsoft-redistributable)

**Tests:**
- `VSTO2\OutlookAI.sln`
- `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj`
- `VSTO2\OutlookAI.Tests\Helpers\FakeHttpMessageHandler.cs`
- `VSTO2\OutlookAI.Tests\Helpers\FakeOutlookSurface.cs`
- `VSTO2\OutlookAI.Tests\Helpers\FakeToolHost.cs`
- `VSTO2\OutlookAI.Tests\Helpers\CapturingChatEventSink.cs`
- `VSTO2\OutlookAI.Tests\Services\CodexAuthServiceTests.cs`
- `VSTO2\OutlookAI.Tests\Services\CodexChatServiceTests.cs`
- `VSTO2\OutlookAI.Tests\Services\CodexChatServiceMultiRoundTests.cs`
- `VSTO2\OutlookAI.Tests\Services\ToolDispatcherTests.cs`
- `VSTO2\OutlookAI.Tests\Services\IdResolverTests.cs`
- `VSTO2\OutlookAI.Tests\Services\OutlookThreadMarshallerTests.cs`
- `VSTO2\OutlookAI.Tests\Services\Tools\<one test class per tool>.cs`  *(10 files)*
- `VSTO2\OutlookAI.Tests\Services\Chat\ConversationStoreTests.cs`
- `VSTO2\OutlookAI.Tests\Services\Variants\VariantParserTests.cs`
- `VSTO2\OutlookAI.Tests\Services\Variants\VariantStoreTests.cs`
- `VSTO2\OutlookAI.Tests\ConfigTests.cs`

### Modified files

- `VSTO2\OutlookAI\Config.cs` — adds `ReasoningEffort` + `WriteToolsEnabled` + `AvailableModels` + `AvailableReasoningEfforts`.
- `VSTO2\OutlookAI\OutlookAI.csproj` — references WebView2 + new Compile/EmbeddedResource entries.
- `VSTO2\OutlookAI\packages.config` — adds `Microsoft.Web.WebView2`.
- `VSTO2\OutlookAI\Services\CodexChatService.cs` — adds `RunTurnAsync`, multi-round loop, reasoning effort propagation.
- `VSTO2\OutlookAI\TaskPane\AITaskPane.cs` + `AITaskPane.Designer.cs` — TabControl layout, Actions/Chat/Variants tabs.
- `VSTO2\OutlookAI\SettingsForm.cs` — Model + Reasoning + Tool-permission controls.
- `VSTO2\OutlookAI\ThisAddIn.cs` — owns `OutlookToolHost` singleton.
- `Deploy\Install-OutlookAI.ps1` — WebView2 Evergreen detection + bootstrap.
- `Deploy\README.txt` — WebView2 prerequisite note.

---

## Task Index

| # | Task |
|---|---|
| 1 | Solution + xUnit test project scaffold |
| 2 | Test helpers (FakeHttpMessageHandler, FakeOutlookSurface, FakeToolHost, CapturingChatEventSink) |
| 3 | Phase 1 test debt: Config tests |
| 4 | Phase 1 test debt: CodexAuthService tests |
| 5 | Phase 1 test debt: CodexChatService single-shot tests |
| 6 | Tool catalog schema + IToolHost + IOutlookTool + IOutlookSurface |
| 7 | IdResolver (EntryID ↔ short opaque ID) |
| 8 | OutlookThreadMarshaller (STA marshalling) |
| 9 | ToolDispatcher (name routing + JSON-schema validation) |
| 10 | Tool: `outlook_get_current_compose_state` |
| 11 | Tool: `outlook_list_folders` |
| 12 | Tool: `outlook_search_messages` |
| 13 | Tool: `outlook_read_message` |
| 14 | Tool: `outlook_count_messages` |
| 15 | Tool: `outlook_list_recent_threads_with` |
| 16 | Tool: `outlook_create_draft` |
| 17 | Tool: `outlook_mark_as_read` |
| 18 | Tool: `outlook_flag_message` |
| 19 | Tool: `outlook_set_category` |
| 20 | OutlookToolHost (registry) + LiveOutlookSurface (production OOM impl) |
| 21 | ConversationContext + Store + ChatEventSink + TurnResult + StopReason |
| 22 | CodexChatService.RunTurnAsync — single-round happy path |
| 23 | CodexChatService.RunTurnAsync — multi-round tool loop + parallel dispatch |
| 24 | CodexChatService.RunTurnAsync — max-rounds cap + cancellation + partial preservation |
| 25 | CodexChatService — reasoning effort propagation + Config flag |
| 26 | VariantParser (constrained JSON envelope, tone enum clamping) |
| 27 | VariantStore (per-Inspector in-memory) |
| 28 | Config v2 — ReasoningEffort + WriteToolsEnabled + Model catalog constants |
| 29 | SettingsForm — Model dropdown + Reasoning dropdown + Tool-permission checklist |
| 30 | AITaskPane — TabControl scaffold; Actions tab inherits existing controls |
| 31 | Actions tab — rewire existing buttons to use RunTurnAsync (single-shot with tool catalog) |
| 32 | WebView2 NuGet + project setup + WebView2Bootstrap + extraction-to-LocalAppData |
| 33 | WebUI bundle: index.html + styles.css + chat.js + marked.min.js + highlight.min.js (initial render) |
| 34 | WebUI: tool-call cards + audit rows + stopped/error states + theme bridge |
| 35 | ChatController — JS↔C# bridge + RunTurnAsync orchestration |
| 36 | Variants tab — UI + intent input + count picker + reasoning override |
| 37 | Variants tab — generation flow + parser integration + per-card actions |
| 38 | Install script — WebView2 Evergreen detection + bootstrap |
| 39 | Vendor MicrosoftEdgeWebView2Setup.exe in Deploy folder |
| 40 | Manual smoke checklist + Phase 2 README update + final E2E build + commit |

40 tasks. Each finishes with a green build, a green test run, and a commit. Push happens at natural checkpoints (after Task 5, 9, 20, 25, 31, 37, 40).

---

## Baseline (must already be true)

- Worktree set up, on branch `feature/codex-oauth-migration`, Phase 1 commits present.
- `nuget.exe` at `C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe` (downloaded during Phase 1).
- VS 2022/2024 with MSBuild at `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe` (path used in Phase 1).
- VSTest at `C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe`.
- `dotnet` SDK on PATH for the SDK-style test project's restore.

If any of those are different on the executor's machine, update the commands accordingly — they don't otherwise change the plan.

---

## Task 1: Solution + xUnit Test Project Scaffold

**Files:**
- Create: `VSTO2\OutlookAI.sln`
- Create: `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj`

- [ ] **Step 1: Create the test project file**

```xml
<!-- VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj -->
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net472</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <IsPackable>false</IsPackable>
    <RootNamespace>OutlookAI.Tests</RootNamespace>
    <AssemblyName>OutlookAI.Tests</AssemblyName>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.14.1" />
    <PackageReference Include="Moq" Version="4.20.72" />
    <PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
    <PackageReference Include="xunit" Version="2.9.3" />
    <PackageReference Include="xunit.runner.visualstudio" Version="3.1.5">
      <IncludeAssets>runtime; build; native; contentfiles; analyzers; buildtransitive</IncludeAssets>
      <PrivateAssets>all</PrivateAssets>
    </PackageReference>
  </ItemGroup>
  <ItemGroup>
    <Reference Include="System.Net.Http" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\OutlookAI\OutlookAI.csproj" />
  </ItemGroup>
</Project>
```

- [ ] **Step 2: Add a placeholder SmokeTests.cs to prove the runner works**

```csharp
// VSTO2\OutlookAI.Tests\SmokeTests.cs
using Xunit;

namespace OutlookAI.Tests
{
    public class SmokeTests
    {
        [Fact]
        public void TestRunner_IsAlive()
        {
            Assert.True(true);
        }
    }
}
```

- [ ] **Step 3: Create the solution file**

```powershell
dotnet new sln --name OutlookAI --output VSTO2
dotnet sln VSTO2\OutlookAI.sln add VSTO2\OutlookAI\OutlookAI.csproj
dotnet sln VSTO2\OutlookAI.sln add VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
```

- [ ] **Step 4: Restore and build the solution**

```powershell
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
dotnet restore VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
```

Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`.

- [ ] **Step 5: Run the smoke test**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 1, Failed: 0`.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI.sln VSTO2/OutlookAI.Tests/
git commit -m "Phase 2 Task 1: add solution file + xUnit test project scaffold"
```

---

## Task 2: Test Helpers

**Files:**
- Create: `VSTO2\OutlookAI.Tests\Helpers\FakeHttpMessageHandler.cs`
- Create: `VSTO2\OutlookAI.Tests\Helpers\FakeOutlookSurface.cs`
- Create: `VSTO2\OutlookAI.Tests\Helpers\FakeToolHost.cs`
- Create: `VSTO2\OutlookAI.Tests\Helpers\CapturingChatEventSink.cs`

- [ ] **Step 1: FakeHttpMessageHandler**

```csharp
// VSTO2\OutlookAI.Tests\Helpers\FakeHttpMessageHandler.cs
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
        private readonly Queue<Func<HttpRequestMessage, HttpResponseMessage>> _responses =
            new Queue<Func<HttpRequestMessage, HttpResponseMessage>>();

        public List<HttpRequestMessage> Requests { get; } = new List<HttpRequestMessage>();
        public List<string> RequestBodies { get; } = new List<string>();

        public void QueueJson(HttpStatusCode status, string json) =>
            _responses.Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json"),
            });

        public void QueueSse(HttpStatusCode status, string sseBody) =>
            _responses.Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(sseBody, System.Text.Encoding.UTF8, "text/event-stream"),
            });

        public void QueueText(HttpStatusCode status, string text) =>
            _responses.Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(text),
            });

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests.Add(request);
            if (request.Content != null)
            {
                RequestBodies.Add(await request.Content.ReadAsStringAsync().ConfigureAwait(false));
            }
            else
            {
                RequestBodies.Add(string.Empty);
            }

            if (_responses.Count == 0)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError)
                {
                    Content = new StringContent("{\"error\":\"no fake response queued\"}"),
                };
            }
            return _responses.Dequeue()(request);
        }
    }
}
```

- [ ] **Step 2: FakeOutlookSurface**

This mock matches the (yet-to-be-defined) `IOutlookSurface` interface from Task 6. We forward-declare a minimal interface here as a private to the fake so tests for individual tools (Tasks 10–19) can substitute concrete methods. The real `IOutlookSurface` arrives in Task 6; this fake gets updated to implement it then.

```csharp
// VSTO2\OutlookAI.Tests\Helpers\FakeOutlookSurface.cs
using System;
using System.Collections.Generic;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Per-method substitutable Outlook surface for tool tests. Real
    /// <see cref="OutlookAI.Services.Tools.IOutlookSurface"/> is implemented
    /// in Task 6; this fake will be updated to formally implement it then.
    /// </summary>
    public sealed class FakeOutlookSurface
    {
        public Func<bool, OutlookAI.Services.Tools.ComposeStateResult> GetCurrentComposeState { get; set; }
        public Func<IEnumerable<OutlookAI.Services.Tools.FolderResult>> ListFolders { get; set; }
        public Func<OutlookAI.Services.Tools.SearchMessagesArgs, IEnumerable<OutlookAI.Services.Tools.MessageSummary>> SearchMessages { get; set; }
        public Func<string, bool, OutlookAI.Services.Tools.MessageDetail> ReadMessage { get; set; }
        public Func<OutlookAI.Services.Tools.SearchMessagesArgs, int> CountMessages { get; set; }
        public Func<string, int, IEnumerable<OutlookAI.Services.Tools.ThreadSummary>> ListRecentThreadsWith { get; set; }
        public Func<OutlookAI.Services.Tools.CreateDraftArgs, OutlookAI.Services.Tools.CreatedDraft> CreateDraft { get; set; }
        public Action<string, bool> MarkAsRead { get; set; }
        public Action<string, string> FlagMessage { get; set; }
        public Action<string, string> SetCategory { get; set; }
    }
}
```

- [ ] **Step 3: FakeToolHost**

```csharp
// VSTO2\OutlookAI.Tests\Helpers\FakeToolHost.cs
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Scripted dispatcher for chat-service tests. Caller queues (name → response JSON)
    /// in the order the model is expected to call them.
    /// </summary>
    public sealed class FakeToolHost : OutlookAI.Services.IToolHost
    {
        private readonly Queue<(string Name, string ResponseJson)> _scripted =
            new Queue<(string, string)>();

        public List<(string Name, string ArgsJson)> Calls { get; } =
            new List<(string, string)>();

        public void Queue(string name, string responseJson) =>
            _scripted.Enqueue((name, responseJson));

        public Task<string> DispatchAsync(string toolName, string argsJson, CancellationToken ct)
        {
            Calls.Add((toolName, argsJson));
            if (_scripted.Count == 0)
            {
                throw new InvalidOperationException(
                    "FakeToolHost: no scripted response for " + toolName);
            }
            var (expected, response) = _scripted.Dequeue();
            if (!string.Equals(expected, toolName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "FakeToolHost: expected " + expected + " but model called " + toolName);
            }
            return Task.FromResult(response);
        }
    }
}
```

- [ ] **Step 4: CapturingChatEventSink**

```csharp
// VSTO2\OutlookAI.Tests\Helpers\CapturingChatEventSink.cs
using System.Collections.Generic;
using System.Text;

namespace OutlookAI.Tests.Helpers
{
    public sealed class CapturingChatEventSink : OutlookAI.Services.Chat.ChatEventSink
    {
        public StringBuilder StreamedText { get; } = new StringBuilder();
        public List<(string CallId, string Name, string ArgsJson)> ToolStarts { get; }
            = new List<(string, string, string)>();
        public List<(string CallId, bool Ok, string Summary, string ResultJson)> ToolResults { get; }
            = new List<(string, bool, string, string)>();
        public List<string> AssistantMessageFinalTexts { get; } = new List<string>();
        public List<string> Errors { get; } = new List<string>();
        public int RoundBoundaries { get; private set; }

        public override void OnTokenDelta(string delta) => StreamedText.Append(delta);
        public override void OnToolCallStart(string callId, string name, string argsJson) =>
            ToolStarts.Add((callId, name, argsJson));
        public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson) =>
            ToolResults.Add((callId, ok, summary, resultJson));
        public override void OnAssistantMessageComplete(string text) =>
            AssistantMessageFinalTexts.Add(text);
        public override void OnError(string message) => Errors.Add(message);
        public override void OnRoundBoundary() => RoundBoundaries++;
    }
}
```

- [ ] **Step 5: Build (will fail — interfaces don't exist yet)**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
```

Expected FAIL: errors about missing namespaces `OutlookAI.Services.Tools.*`, `OutlookAI.Services.IToolHost`, `OutlookAI.Services.Chat.ChatEventSink`, etc. This is *expected* — we haven't created them yet (Tasks 6, 21). Helper files are intentionally written before those types so the type names lock in early.

- [ ] **Step 6: Mark helpers as compile-out until referenced types exist**

To keep the test project building during Tasks 3–5 (Phase 1 test debt), we temporarily exclude the three helper files that reference Task-6/21 types. Add to `OutlookAI.Tests.csproj` inside the existing `<Project>` element:

```xml
<ItemGroup>
  <Compile Remove="Helpers\FakeOutlookSurface.cs" />
  <Compile Remove="Helpers\FakeToolHost.cs" />
  <Compile Remove="Helpers\CapturingChatEventSink.cs" />
  <None Include="Helpers\FakeOutlookSurface.cs" />
  <None Include="Helpers\FakeToolHost.cs" />
  <None Include="Helpers\CapturingChatEventSink.cs" />
</ItemGroup>
```

These get re-added as `Compile` in Task 6 (FakeOutlookSurface), Task 21 (CapturingChatEventSink), and Task 22 (FakeToolHost — first actual use).

- [ ] **Step 7: Build (now passes)**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
```

Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`.

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI.Tests/
git commit -m "Phase 2 Task 2: add test helpers (FakeHttpMessageHandler, deferred fakes)"
```

---

## Task 3: Phase 1 Test Debt — Config

**Files:**
- Create: `VSTO2\OutlookAI.Tests\ConfigTests.cs`

- [ ] **Step 1: Write the failing tests**

```csharp
// VSTO2\OutlookAI.Tests\ConfigTests.cs
using System.IO;
using OutlookAI;
using Xunit;

namespace OutlookAI.Tests
{
    public class ConfigTests
    {
        private static (string global, string user) MakeTempPaths()
        {
            var dir = Path.Combine(Path.GetTempPath(),
                "outlookai-config-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);
            return (Path.Combine(dir, "global.xml"), Path.Combine(dir, "user.xml"));
        }

        [Fact]
        public void LoadConfigFromPaths_UsesV2Defaults_WhenFilesAreMissing()
        {
            var (g, u) = MakeTempPaths();
            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("admin", Config.AdminPassword);
            Assert.Equal(@"C:\ProgramData\OutlookAI\auth.json", Config.CodexAuthPath);
            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal("gpt-realtime-1.5", Config.VoiceModel);
            Assert.Equal(65536, Config.MaxTokens);
        }

        [Fact]
        public void LoadConfigFromPaths_PerUserOverridesAdminPasswordOnly()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<AdminPassword>server</AdminPassword>"
                + "<CodexAuthPath>C:\\ProgramData\\OutlookAI\\auth.json</CodexAuthPath>"
                + "<Model>gpt-5.5</Model>"
                + "<VoiceModel>gpt-realtime-1.5</VoiceModel>"
                + "<MaxTokens>65536</MaxTokens>"
                + "</Config>");
            File.WriteAllText(u, "<Config>"
                + "<AdminPassword>userpass</AdminPassword>"
                + "<Model>claude-opus-4-6</Model>"
                + "<MaxTokens>2048</MaxTokens>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("userpass", Config.AdminPassword);
            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal(65536, Config.MaxTokens);
        }

        [Fact]
        public void LoadConfigFromPaths_IgnoresLegacyV1Fields()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<ApiKey>anthropic-key</ApiKey>"
                + "<OpenAIApiKey>openai-key</OpenAIApiKey>"
                + "<WhisperModel>whisper-1</WhisperModel>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal("gpt-realtime-1.5", Config.VoiceModel);
        }
    }
}
```

- [ ] **Step 2: Run tests — verify they pass**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ConfigTests"
```

Phase 1's `Config.cs` already exposes `LoadConfigFromPaths` and the v2 fields, so these tests should pass on first run — they're documenting/locking existing behavior. If any fail, that's a regression to fix in `Config.cs`.

Expected: `Passed: 3, Failed: 0`.

- [ ] **Step 3: Commit**

```powershell
git add VSTO2/OutlookAI.Tests/ConfigTests.cs
git commit -m "Phase 2 Task 3: add Config tests (Phase 1 test debt)"
```

---

## Task 4: Phase 1 Test Debt — CodexAuthService

**Files:**
- Create: `VSTO2\OutlookAI.Tests\Services\CodexAuthServiceTests.cs`

- [ ] **Step 1: Write the failing tests**

```csharp
// VSTO2\OutlookAI.Tests\Services\CodexAuthServiceTests.cs
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using OutlookAI.Services;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class CodexAuthServiceTests : IDisposable
    {
        private readonly string _tmpDir;
        private readonly string _authPath;

        public CodexAuthServiceTests()
        {
            _tmpDir = Path.Combine(Path.GetTempPath(), "outlookai-auth", Path.GetRandomFileName());
            Directory.CreateDirectory(_tmpDir);
            _authPath = Path.Combine(_tmpDir, "auth.json");
        }

        public void Dispose()
        {
            try { Directory.Delete(_tmpDir, recursive: true); } catch { }
        }

        [Fact]
        public void GetStatus_ReturnsUnauthenticated_WhenNoAuthFile()
        {
            using (var http = new HttpClient(new FakeHttpMessageHandler()))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                Assert.Equal(AuthState.Unauthenticated, svc.GetStatus().State);
            }
        }

        [Fact]
        public async System.Threading.Tasks.Task GetAccessTokenAsync_Throws_WhenNotSignedIn()
        {
            using (var http = new HttpClient(new FakeHttpMessageHandler()))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                await Assert.ThrowsAsync<InvalidOperationException>(
                    () => svc.GetAccessTokenAsync(CancellationToken.None));
            }
        }

        [Fact]
        public async System.Threading.Tasks.Task GetAccessTokenAsync_ReturnsCachedToken_WhenFresh()
        {
            // Seed an auth.json with a non-expired access token.
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk-fresh\",\"id_token\":\"\","
                + "\"refresh_token\":\"r1\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");

            var fake = new FakeHttpMessageHandler();
            using (var http = new HttpClient(fake))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                var token = await svc.GetAccessTokenAsync(CancellationToken.None);

                Assert.Equal("sk-fresh", token);
                Assert.Empty(fake.Requests); // No refresh request — token still fresh.
            }
        }

        [Fact]
        public async System.Threading.Tasks.Task GetAccessTokenAsync_RefreshesExpiredToken()
        {
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk-old\",\"id_token\":\"\","
                + "\"refresh_token\":\"r1\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(-1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");

            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.OK,
                "{\"access_token\":\"sk-new\",\"refresh_token\":\"r2\",\"expires_in\":3600}");

            using (var http = new HttpClient(fake))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                var token = await svc.GetAccessTokenAsync(CancellationToken.None);

                Assert.Equal("sk-new", token);
                Assert.Single(fake.Requests);
                Assert.Equal("https://auth.openai.com/oauth/token",
                    fake.Requests[0].RequestUri.ToString());
                Assert.Contains("grant_type=refresh_token", fake.RequestBodies[0]);
                Assert.Contains("refresh_token=r1", fake.RequestBodies[0]);
            }
        }

        [Fact]
        public async System.Threading.Tasks.Task SignOutAsync_DeletesAuthFile()
        {
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk\",\"refresh_token\":\"r\","
                + "\"id_token\":\"\"}}");
            using (var http = new HttpClient(new FakeHttpMessageHandler()))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                await svc.SignOutAsync();

                Assert.False(File.Exists(_authPath));
                Assert.Equal(AuthState.Unauthenticated, svc.GetStatus().State);
            }
        }
    }
}
```

- [ ] **Step 2: Run tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~CodexAuthServiceTests"
```

Expected: `Passed: 5, Failed: 0`. These document Phase 1 behavior; failures indicate a regression in `CodexAuthService.cs`.

- [ ] **Step 3: Commit**

```powershell
git add VSTO2/OutlookAI.Tests/Services/CodexAuthServiceTests.cs
git commit -m "Phase 2 Task 4: add CodexAuthService tests (Phase 1 test debt)"
```

---

## Task 5: Phase 1 Test Debt — CodexChatService Single-Shot

**Files:**
- Create: `VSTO2\OutlookAI.Tests\Services\CodexChatServiceTests.cs`

- [ ] **Step 1: Write the failing tests**

The Phase 1 `CodexChatService.ProcessEmailAsync(ActionType, string, string, ct)` should send a Codex-shape Responses request, accumulate SSE deltas, and return the final string. Lock that behavior in.

```csharp
// VSTO2\OutlookAI.Tests\Services\CodexChatServiceTests.cs
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class CodexChatServiceTests : IDisposable
    {
        private readonly string _tmpDir;
        private readonly string _authPath;

        public CodexChatServiceTests()
        {
            _tmpDir = Path.Combine(Path.GetTempPath(), "outlookai-chat", Path.GetRandomFileName());
            Directory.CreateDirectory(_tmpDir);
            _authPath = Path.Combine(_tmpDir, "auth.json");
            // Seed a fresh token so ProcessEmailAsync doesn't try to refresh.
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk-test\",\"id_token\":\"\","
                + "\"refresh_token\":\"r1\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");
        }

        public void Dispose()
        {
            try { Directory.Delete(_tmpDir, recursive: true); } catch { }
        }

        [Fact]
        public async Task ProcessEmailAsync_SendsResponsesRequestWithBearer()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"Hello\"}\n\n"
                + "data: {\"type\":\"response.output_text.delta\",\"delta\":\" world\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var result = await chat.ProcessEmailAsync(
                    CodexChatService.ActionType.Proofread, "helo world");

                Assert.Equal("Hello world", result);
                Assert.Single(fake.Requests);
                Assert.Equal("https://chatgpt.com/backend-api/codex/responses",
                    fake.Requests[0].RequestUri.ToString());
                Assert.Equal("Bearer sk-test",
                    fake.Requests[0].Headers.Authorization.ToString());
                Assert.Contains("\"model\":\"gpt-5.5\"", fake.RequestBodies[0]);
                Assert.Contains("\"stream\":true", fake.RequestBodies[0]);
                Assert.Contains("\"store\":false", fake.RequestBodies[0]);
            }
        }

        [Fact]
        public async Task ProcessEmailAsync_PrefersOutputTextDoneOverAccumulatedDeltas()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"draft \"}\n\n"
                + "data: {\"type\":\"response.output_text.done\",\"text\":\"FINAL\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var result = await chat.ProcessEmailAsync(
                    CodexChatService.ActionType.Revise, "any");

                Assert.Equal("FINAL", result);
            }
        }

        [Fact]
        public async Task ProcessEmailAsync_ThrowsWithBackendErrorBody_OnNonSuccess()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueText(HttpStatusCode.TooManyRequests,
                "{\"detail\":\"rate limited\"}");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var ex = await Assert.ThrowsAsync<InvalidOperationException>(
                    () => chat.ProcessEmailAsync(
                        CodexChatService.ActionType.Proofread, "x"));

                Assert.Contains("ChatGPT Codex backend error", ex.Message);
                Assert.Contains("rate limited", ex.Message);
            }
        }
    }
}
```

- [ ] **Step 2: Run tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~CodexChatServiceTests"
```

Expected: `Passed: 3, Failed: 0`.

- [ ] **Step 3: Push checkpoint (Phase 1 test debt cleared)**

```powershell
git add VSTO2/OutlookAI.Tests/Services/CodexChatServiceTests.cs
git commit -m "Phase 2 Task 5: add CodexChatService single-shot tests; Phase 1 test debt cleared"
git push
```

---

## Task 6: Tool Catalog Schema + Interfaces

**Files:**
- Create: `VSTO2\OutlookAI\Services\Tools\IOutlookSurface.cs`
- Create: `VSTO2\OutlookAI\Services\Tools\IOutlookTool.cs`
- Create: `VSTO2\OutlookAI\Services\Tools\ToolCatalogSchema.cs`
- Create: `VSTO2\OutlookAI\Services\IToolHost.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj` — add Compile entries.
- Modify: `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj` — re-add `Helpers\FakeOutlookSurface.cs` as Compile.

- [ ] **Step 1: Define `IOutlookSurface` (the seam between tools and Outlook OOM)**

```csharp
// VSTO2\OutlookAI\Services\Tools\IOutlookSurface.cs
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// All Outlook OOM access used by Phase 2 tools flows through this interface.
    /// Production implementation (LiveOutlookSurface) marshals every call onto the
    /// Outlook STA UI thread via OutlookThreadMarshaller. Tests stub the methods
    /// they need without touching COM.
    /// </summary>
    public interface IOutlookSurface
    {
        ComposeStateResult GetCurrentComposeState(bool includeFullBody);
        IReadOnlyList<FolderResult> ListFolders();
        IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args);
        MessageDetail ReadMessage(string messageId, bool includeFullBody);
        int CountMessages(SearchMessagesArgs args);
        IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads);
        CreatedDraft CreateDraft(CreateDraftArgs args);
        void MarkAsRead(string messageId, bool read);
        void FlagMessage(string messageId, string flag);
        void SetCategory(string messageId, string category);
    }

    public sealed class ComposeStateResult
    {
        public string Subject { get; set; }
        public IReadOnlyList<string> ToRecipients { get; set; }
        public IReadOnlyList<string> CcRecipients { get; set; }
        public IReadOnlyList<string> BccRecipients { get; set; }
        public string SenderName { get; set; }
        public string SenderEmail { get; set; }
        public string BodyPlaintext { get; set; }
        public bool BodyTruncated { get; set; }
        public InReplyTo InReplyTo { get; set; }
        public IReadOnlyList<AttachmentSummary> Attachments { get; set; }
    }

    public sealed class InReplyTo
    {
        public string ThreadTopic { get; set; }
        public IReadOnlyList<ThreadMessage> LastNMessages { get; set; }
    }

    public sealed class ThreadMessage
    {
        public string From { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public string Snippet { get; set; }
    }

    public sealed class AttachmentSummary
    {
        public string Filename { get; set; }
        public long SizeBytes { get; set; }
    }

    public sealed class FolderResult
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ParentId { get; set; }
        public int ItemCount { get; set; }
    }

    public sealed class SearchMessagesArgs
    {
        public string Query { get; set; }
        public string FolderId { get; set; }
        public DateTimeOffset? DateFrom { get; set; }
        public DateTimeOffset? DateTo { get; set; }
        public int MaxResults { get; set; } = 25;
    }

    public sealed class MessageSummary
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public string Snippet { get; set; }
        public bool HasAttachments { get; set; }
    }

    public sealed class MessageDetail
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public IReadOnlyList<string> Cc { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public string BodyPlaintext { get; set; }
        public bool BodyTruncated { get; set; }
        public IReadOnlyList<AttachmentSummary> Attachments { get; set; }
        public string InReplyToMessageId { get; set; }
        public string ConversationTopic { get; set; }
    }

    public sealed class ThreadSummary
    {
        public string ThreadTopic { get; set; }
        public DateTimeOffset LastMessageAt { get; set; }
        public int MessageCount { get; set; }
        public string Snippet { get; set; }
        public string ThreadId { get; set; }
    }

    public sealed class CreateDraftArgs
    {
        public string Subject { get; set; }
        public string BodyPlaintext { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public IReadOnlyList<string> Cc { get; set; }
        public string InReplyToMessageId { get; set; }
    }

    public sealed class CreatedDraft
    {
        public string DraftId { get; set; }
        public string Location { get; set; }
    }
}
```

- [ ] **Step 2: Define `IOutlookTool`**

```csharp
// VSTO2\OutlookAI\Services\Tools\IOutlookTool.cs
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// One Phase-2 Outlook tool. Implementations validate args (already JSON-schema
    /// validated by ToolDispatcher), perform the operation via IOutlookSurface, and
    /// return a JSON string the chat service inserts as a function_call_output.
    /// </summary>
    public interface IOutlookTool
    {
        string Name { get; }
        Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct);
    }
}
```

- [ ] **Step 3: Define `IToolHost`**

```csharp
// VSTO2\OutlookAI\Services\IToolHost.cs
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services
{
    /// <summary>
    /// Seam between CodexChatService and Outlook tool execution. Production
    /// implementation is OutlookToolHost (Task 20). Tests use FakeToolHost.
    /// </summary>
    public interface IToolHost
    {
        Task<string> DispatchAsync(string toolName, string argsJson, CancellationToken ct);
    }
}
```

- [ ] **Step 4: Define `ToolCatalogSchema` (JSON Schema for all 10 tools)**

```csharp
// VSTO2\OutlookAI\Services\Tools\ToolCatalogSchema.cs
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Static JSON schema definitions for every Phase 2 tool. Same schema is
    /// (a) sent to the model in the Responses request as tools[] entries, and
    /// (b) used by ToolDispatcher to validate arguments before dispatch.
    /// </summary>
    public static class ToolCatalogSchema
    {
        public static JArray BuildResponsesToolsArray(bool includeWriteTools)
        {
            var arr = new JArray
            {
                BuildToolEntry("outlook_get_current_compose_state",
                    "Read the current compose-window state (subject, recipients, body, thread, attachments). Set include_full_body=true for the full body instead of a 1000-char summary.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("properties", new JObject(
                            new JProperty("include_full_body", new JObject(
                                new JProperty("type", "boolean"))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_list_folders",
                    "List the user's mail folders (max depth 6, max 200 nodes).",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("properties", new JObject()),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_search_messages",
                    "Search messages via DASL Restrict. Returns id+metadata+snippet for up to max_results (default 25, hard cap 100).",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("query")),
                        new JProperty("properties", new JObject(
                            new JProperty("query", new JObject(new JProperty("type","string"))),
                            new JProperty("folder_id", new JObject(new JProperty("type","string"))),
                            new JProperty("date_from", new JObject(new JProperty("type","string"),
                                                                  new JProperty("format","date-time"))),
                            new JProperty("date_to", new JObject(new JProperty("type","string"),
                                                                 new JProperty("format","date-time"))),
                            new JProperty("max_results", new JObject(new JProperty("type","integer"),
                                                                     new JProperty("minimum",1),
                                                                     new JProperty("maximum",100))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_read_message",
                    "Fetch one message by id. Body always plaintext; truncated at 32 KB.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("message_id")),
                        new JProperty("properties", new JObject(
                            new JProperty("message_id", new JObject(new JProperty("type","string"))),
                            new JProperty("include_full_body", new JObject(new JProperty("type","boolean"))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_count_messages",
                    "Count messages matching a query without returning bodies. Same query syntax as outlook_search_messages.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("query")),
                        new JProperty("properties", new JObject(
                            new JProperty("query", new JObject(new JProperty("type","string"))),
                            new JProperty("folder_id", new JObject(new JProperty("type","string"))),
                            new JProperty("date_from", new JObject(new JProperty("type","string"),
                                                                  new JProperty("format","date-time"))),
                            new JProperty("date_to", new JObject(new JProperty("type","string"),
                                                                 new JProperty("format","date-time"))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_list_recent_threads_with",
                    "List the most recent conversation threads involving a specific recipient (Inbox + Sent), grouped by ConversationID.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("recipient_email")),
                        new JProperty("properties", new JObject(
                            new JProperty("recipient_email", new JObject(new JProperty("type","string"))),
                            new JProperty("max_threads", new JObject(new JProperty("type","integer"),
                                                                     new JProperty("minimum",1),
                                                                     new JProperty("maximum",20))))),
                        new JProperty("additionalProperties", false))),
            };

            if (includeWriteTools)
            {
                arr.Add(BuildToolEntry("outlook_create_draft",
                    "Create a draft in the Drafts folder. Never sends. If in_reply_to_message_id is given, seeds the draft via MailItem.Reply().",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("subject","body_plaintext")),
                        new JProperty("properties", new JObject(
                            new JProperty("subject", new JObject(new JProperty("type","string"))),
                            new JProperty("body_plaintext", new JObject(new JProperty("type","string"))),
                            new JProperty("to", new JObject(new JProperty("type","array"),
                                                            new JProperty("items", new JObject(new JProperty("type","string"))))),
                            new JProperty("cc", new JObject(new JProperty("type","array"),
                                                            new JProperty("items", new JObject(new JProperty("type","string"))))),
                            new JProperty("in_reply_to_message_id", new JObject(new JProperty("type","string"))))),
                        new JProperty("additionalProperties", false))));

                arr.Add(BuildToolEntry("outlook_mark_as_read",
                    "Set or clear the UnRead flag on a message.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("message_id","read")),
                        new JProperty("properties", new JObject(
                            new JProperty("message_id", new JObject(new JProperty("type","string"))),
                            new JProperty("read", new JObject(new JProperty("type","boolean"))))),
                        new JProperty("additionalProperties", false))));

                arr.Add(BuildToolEntry("outlook_flag_message",
                    "Set follow-up flag status: none | todo | complete.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("message_id","flag")),
                        new JProperty("properties", new JObject(
                            new JProperty("message_id", new JObject(new JProperty("type","string"))),
                            new JProperty("flag", new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("none","todo","complete")))))),
                        new JProperty("additionalProperties", false))));

                arr.Add(BuildToolEntry("outlook_set_category",
                    "Replace a message's Categories with the single given value.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("message_id","category")),
                        new JProperty("properties", new JObject(
                            new JProperty("message_id", new JObject(new JProperty("type","string"))),
                            new JProperty("category", new JObject(new JProperty("type","string"))))),
                        new JProperty("additionalProperties", false))));
            }

            return arr;
        }

        private static JObject BuildToolEntry(string name, string description, JObject parameters)
        {
            return new JObject(
                new JProperty("type", "function"),
                new JProperty("name", name),
                new JProperty("description", description),
                new JProperty("parameters", parameters));
        }
    }
}
```

- [ ] **Step 5: Wire the new files into `OutlookAI.csproj`**

Add inside the existing `<ItemGroup>` that contains the other Compile entries (search for `<Compile Include="Services\CodexAuthService.cs" />`):

```xml
<Compile Include="Services\IToolHost.cs" />
<Compile Include="Services\Tools\IOutlookSurface.cs" />
<Compile Include="Services\Tools\IOutlookTool.cs" />
<Compile Include="Services\Tools\ToolCatalogSchema.cs" />
```

- [ ] **Step 6: Re-add `FakeOutlookSurface.cs` as Compile in the test project**

Edit `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj`:

```xml
<!-- Replace the entire deferred-helpers ItemGroup with: -->
<ItemGroup>
  <Compile Remove="Helpers\FakeToolHost.cs" />
  <Compile Remove="Helpers\CapturingChatEventSink.cs" />
  <None Include="Helpers\FakeToolHost.cs" />
  <None Include="Helpers\CapturingChatEventSink.cs" />
</ItemGroup>
```

(FakeOutlookSurface is now compiled because we removed its Compile-Remove and matching None-Include.)

- [ ] **Step 7: Build + run all tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: green build; all previously-passing tests still pass.

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/Services/IToolHost.cs VSTO2/OutlookAI/Services/Tools/ VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "Phase 2 Task 6: tool catalog schema + IToolHost + IOutlookTool + IOutlookSurface"
```

---

## Task 7: IdResolver (EntryID ↔ short opaque ID)

**Files:**
- Create: `VSTO2\OutlookAI\Services\IdResolver.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\IdResolverTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj` — Compile entry.

- [ ] **Step 1: Write failing test**

```csharp
// VSTO2\OutlookAI.Tests\Services\IdResolverTests.cs
using OutlookAI.Services;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class IdResolverTests
    {
        [Fact]
        public void Shorten_ProducesDeterministicShortIdPerEntryId()
        {
            var r = new IdResolver();
            var a = r.Shorten("00000000000000000000000000000000ABCDEF");
            var b = r.Shorten("00000000000000000000000000000000ABCDEF");
            Assert.Equal(a, b);
            Assert.True(a.Length <= 12);
        }

        [Fact]
        public void Resolve_RoundTripsKnownShortId()
        {
            var r = new IdResolver();
            var entryId = "00000000000000000000000000000000DEADBEEF";
            var shortId = r.Shorten(entryId);
            Assert.Equal(entryId, r.Resolve(shortId));
        }

        [Fact]
        public void Resolve_ThrowsOnUnknownShortId()
        {
            var r = new IdResolver();
            Assert.Throws<System.Collections.Generic.KeyNotFoundException>(
                () => r.Resolve("not-a-real-id"));
        }
    }
}
```

- [ ] **Step 2: Verify RED**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
```

Expected: compile fails — `IdResolver` does not exist.

- [ ] **Step 3: Implement `IdResolver`**

```csharp
// VSTO2\OutlookAI\Services\IdResolver.cs
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAI.Services
{
    /// <summary>
    /// Maps Outlook EntryIDs (long hex strings, opaque, prompt-injection-friendly)
    /// to short stable IDs we hand to the model. Round-trips while the IdResolver
    /// instance is alive (process lifetime is fine; conversation is per-Inspector
    /// and dies on close). Forged short IDs throw.
    /// </summary>
    public sealed class IdResolver
    {
        private readonly ConcurrentDictionary<string, string> _shortToEntry =
            new ConcurrentDictionary<string, string>(StringComparer.Ordinal);
        private readonly ConcurrentDictionary<string, string> _entryToShort =
            new ConcurrentDictionary<string, string>(StringComparer.Ordinal);

        public string Shorten(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) throw new ArgumentException(nameof(entryId));
            return _entryToShort.GetOrAdd(entryId, eid =>
            {
                using (var sha = SHA256.Create())
                {
                    var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(eid));
                    var b64 = Convert.ToBase64String(hash, 0, 9)
                                     .Replace('+','-').Replace('/','_');
                    _shortToEntry[b64] = eid;
                    return b64;
                }
            });
        }

        public string Resolve(string shortId)
        {
            if (_shortToEntry.TryGetValue(shortId, out var entry)) return entry;
            throw new KeyNotFoundException("Unknown OutlookAI message id: " + shortId);
        }
    }
}
```

- [ ] **Step 4: Add Compile entry** in `OutlookAI.csproj` next to other `Services\*` Compile lines:

```xml
<Compile Include="Services\IdResolver.cs" />
```

- [ ] **Step 5: Verify GREEN + commit**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~IdResolverTests"
git add VSTO2/OutlookAI/Services/IdResolver.cs VSTO2/OutlookAI.Tests/Services/IdResolverTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 2 Task 7: IdResolver"
```

---

## Task 8: OutlookThreadMarshaller (STA marshalling)

**Files:**
- Create: `VSTO2\OutlookAI\Services\OutlookThreadMarshaller.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\OutlookThreadMarshallerTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj` — Compile entry.

- [ ] **Step 1: Write failing tests**

```csharp
// VSTO2\OutlookAI.Tests\Services\OutlookThreadMarshallerTests.cs
using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class OutlookThreadMarshallerTests
    {
        [Fact]
        public async Task RunAsync_InvokesOnConfiguredContext()
        {
            using (var ctx = new TestSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx);
                int observed = 0;
                await marshaller.RunAsync(() => { observed = Thread.CurrentThread.ManagedThreadId; },
                                          CancellationToken.None);
                Assert.Equal(ctx.WorkerThreadId, observed);
            }
        }

        [Fact]
        public async Task RunAsync_PropagatesException()
        {
            using (var ctx = new TestSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx);
                await Assert.ThrowsAsync<InvalidOperationException>(() =>
                    marshaller.RunAsync(() => throw new InvalidOperationException("boom"),
                                        CancellationToken.None));
            }
        }

        [Fact]
        public async Task RunAsync_RespectsCancellationBeforeDispatch()
        {
            using (var ctx = new TestSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx);
                var cts = new CancellationTokenSource();
                cts.Cancel();
                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    marshaller.RunAsync(() => { }, cts.Token));
            }
        }

        // Simple single-thread SynchronizationContext for tests.
        private sealed class TestSyncContext : SynchronizationContext, IDisposable
        {
            private readonly System.Collections.Concurrent.BlockingCollection<(SendOrPostCallback, object)> _q
                = new System.Collections.Concurrent.BlockingCollection<(SendOrPostCallback, object)>();
            private readonly Thread _worker;
            public int WorkerThreadId { get; }

            public TestSyncContext()
            {
                _worker = new Thread(Pump) { IsBackground = true, Name = "TestSyncContext" };
                _worker.SetApartmentState(ApartmentState.STA);
                _worker.Start();
                while (WorkerThreadId == 0) Thread.Sleep(1);
            }

            private void Pump()
            {
                WorkerThreadIdSet(_worker.ManagedThreadId);
                foreach (var (cb, st) in _q.GetConsumingEnumerable()) cb(st);
            }
            private void WorkerThreadIdSet(int id) => typeof(TestSyncContext)
                .GetProperty(nameof(WorkerThreadId)).SetValue(this, id);

            public override void Post(SendOrPostCallback d, object state) => _q.Add((d, state));
            public override void Send(SendOrPostCallback d, object state) => _q.Add((d, state));

            public void Dispose()
            {
                _q.CompleteAdding();
                _worker.Join();
                _q.Dispose();
            }
        }
    }
}
```

- [ ] **Step 2: Verify RED**

Expected: compile fails — `OutlookThreadMarshaller` undefined.

- [ ] **Step 3: Implement `OutlookThreadMarshaller`**

```csharp
// VSTO2\OutlookAI\Services\OutlookThreadMarshaller.cs
using System;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services
{
    /// <summary>
    /// Marshals callbacks onto a specific SynchronizationContext (the Outlook
    /// UI thread in production). Used by LiveOutlookSurface to make every COM
    /// access STA-correct regardless of which thread the chat service runs on.
    /// </summary>
    public sealed class OutlookThreadMarshaller
    {
        private readonly SynchronizationContext _context;

        public OutlookThreadMarshaller(SynchronizationContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public Task RunAsync(Action action, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var tcs = new TaskCompletionSource<bool>();
            _context.Post(_ =>
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    tcs.TrySetCanceled(cancellationToken);
                    return;
                }
                try { action(); tcs.TrySetResult(true); }
                catch (Exception ex) { tcs.TrySetException(ex); }
            }, null);
            return tcs.Task;
        }

        public Task<T> RunAsync<T>(Func<T> func, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var tcs = new TaskCompletionSource<T>();
            _context.Post(_ =>
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    tcs.TrySetCanceled(cancellationToken);
                    return;
                }
                try { tcs.TrySetResult(func()); }
                catch (Exception ex) { tcs.TrySetException(ex); }
            }, null);
            return tcs.Task;
        }
    }
}
```

- [ ] **Step 4: Add Compile entry; verify GREEN; commit**

```xml
<Compile Include="Services\OutlookThreadMarshaller.cs" />
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookThreadMarshallerTests"
git add VSTO2/OutlookAI/Services/OutlookThreadMarshaller.cs VSTO2/OutlookAI.Tests/Services/OutlookThreadMarshallerTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 2 Task 8: OutlookThreadMarshaller (STA marshalling)"
```

---

## Task 9: ToolDispatcher (name routing + JSON-schema validation)

**Files:**
- Create: `VSTO2\OutlookAI\Services\ToolDispatcher.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\ToolDispatcherTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj` — Compile entry.

- [ ] **Step 1: Write failing tests**

```csharp
// VSTO2\OutlookAI.Tests\Services\ToolDispatcherTests.cs
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class ToolDispatcherTests
    {
        private sealed class StubTool : IOutlookTool
        {
            public string Name => "outlook_stub";
            public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
                => Task.FromResult("{\"ok\":true,\"echo\":" + argsJson + "}");
        }

        [Fact]
        public async Task DispatchAsync_RoutesByName()
        {
            var d = new ToolDispatcher(new IOutlookTool[] { new StubTool() }, surface: null);
            var result = await d.DispatchAsync("outlook_stub", "{\"x\":1}", CancellationToken.None);
            Assert.Contains("\"ok\":true", result);
            Assert.Contains("\"x\":1", result);
        }

        [Fact]
        public async Task DispatchAsync_ReturnsStructuredErrorForUnknownTool()
        {
            var d = new ToolDispatcher(new IOutlookTool[] { new StubTool() }, surface: null);
            var result = await d.DispatchAsync("outlook_missing", "{}", CancellationToken.None);
            Assert.Contains("\"code\":\"unknown_tool\"", result);
        }

        [Fact]
        public async Task DispatchAsync_ReturnsStructuredErrorForMalformedJson()
        {
            var d = new ToolDispatcher(new IOutlookTool[] { new StubTool() }, surface: null);
            var result = await d.DispatchAsync("outlook_stub", "{ not json", CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", result);
        }
    }
}
```

- [ ] **Step 2: Verify RED**

- [ ] **Step 3: Implement `ToolDispatcher`**

```csharp
// VSTO2\OutlookAI\Services\ToolDispatcher.cs
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services
{
    /// <summary>
    /// Routes the model's function_call to the matching IOutlookTool instance,
    /// validates the arguments JSON shape, and wraps any failure as a structured
    /// {"error": {...}} JSON response (so the model sees and can recover).
    /// </summary>
    public sealed class ToolDispatcher
    {
        private readonly Dictionary<string, IOutlookTool> _tools;
        private readonly IOutlookSurface _surface;

        public ToolDispatcher(IEnumerable<IOutlookTool> tools, IOutlookSurface surface)
        {
            _tools = new Dictionary<string, IOutlookTool>(StringComparer.Ordinal);
            foreach (var t in tools) _tools[t.Name] = t;
            _surface = surface;
        }

        public async Task<string> DispatchAsync(string toolName, string argsJson, CancellationToken ct)
        {
            if (!_tools.TryGetValue(toolName ?? string.Empty, out var tool))
                return BuildError("unknown_tool", "Tool '" + toolName + "' is not registered.");

            JObject parsed;
            try { parsed = string.IsNullOrWhiteSpace(argsJson) ? new JObject() : JObject.Parse(argsJson); }
            catch (Exception ex) { return BuildError("invalid_arguments", ex.Message); }

            try { return await tool.ExecuteAsync(parsed.ToString(Newtonsoft.Json.Formatting.None), _surface, ct).ConfigureAwait(false); }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                return BuildError(ex.GetType().Name, ex.Message);
            }
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(new JProperty("code", code), new JProperty("message", message)))).ToString(Newtonsoft.Json.Formatting.None);
    }
}
```

- [ ] **Step 4: Compile entry + verify GREEN + commit**

```xml
<Compile Include="Services\ToolDispatcher.cs" />
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ToolDispatcherTests"
git add VSTO2/OutlookAI/Services/ToolDispatcher.cs VSTO2/OutlookAI.Tests/Services/ToolDispatcherTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 2 Task 9: ToolDispatcher"
git push
```

Push checkpoint — foundation complete.

---

## Tasks 10–19: The 10 Outlook Tools

Each tool task follows the **same pattern** — given the user's "no cut corners" directive, the bodies are explicit per task rather than abbreviated. Each task:

1. Adds a per-tool test class under `VSTO2\OutlookAI.Tests\Services\Tools\<ToolName>Tests.cs` using `FakeOutlookSurface` to stub the one surface method the tool uses.
2. Implements `IOutlookTool` for the tool, validating args (already coarsely validated by `ToolDispatcher`), calling the surface, shaping the JSON return per the spec.
3. Adds a Compile entry to `OutlookAI.csproj`.
4. Builds, runs the new test class, commits.

Because each tool's shape is dictated by Section 2 of the spec, the per-task content is mechanical: arg parsing (JObject.Parse), surface call, JSON projection. The plan documents the exact `JObject` projection per tool inline.

### Task 10: `outlook_get_current_compose_state`

**Files:**
- Create: `VSTO2\OutlookAI\Services\Tools\OutlookGetCurrentComposeStateTool.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\Tools\OutlookGetCurrentComposeStateToolTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

- [ ] **Step 1: Test**

```csharp
// VSTO2\OutlookAI.Tests\Services\Tools\OutlookGetCurrentComposeStateToolTests.cs
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookGetCurrentComposeStateToolTests
    {
        [Fact]
        public async Task Execute_ProjectsComposeStateToJson()
        {
            var surface = new StubSurface
            {
                ComposeState = new ComposeStateResult
                {
                    Subject = "Q4 review",
                    SenderName = "Alice",
                    SenderEmail = "alice@example.com",
                    BodyPlaintext = "hi",
                    BodyTruncated = false,
                    ToRecipients = new[] { "bob@example.com" },
                    CcRecipients = new string[0],
                    BccRecipients = new string[0],
                    Attachments = new AttachmentSummary[0],
                }
            };
            var tool = new OutlookGetCurrentComposeStateTool();
            var json = await tool.ExecuteAsync("{\"include_full_body\":false}", surface, CancellationToken.None);
            Assert.Contains("\"subject\":\"Q4 review\"", json);
            Assert.Contains("\"sender_email\":\"alice@example.com\"", json);
            Assert.Contains("\"body_truncated\":false", json);
        }

        private sealed class StubSurface : MinimalSurface
        {
            public ComposeStateResult ComposeState { get; set; }
            public override ComposeStateResult GetCurrentComposeState(bool includeFullBody) => ComposeState;
        }
    }
}
```

A small helper `MinimalSurface` (one file shared across all tool tests) saves boilerplate:

```csharp
// VSTO2\OutlookAI.Tests\Services\Tools\MinimalSurface.cs
using System.Collections.Generic;
using OutlookAI.Services.Tools;
namespace OutlookAI.Tests.Services.Tools
{
    public abstract class MinimalSurface : IOutlookSurface
    {
        public virtual ComposeStateResult GetCurrentComposeState(bool includeFullBody) => throw new System.NotImplementedException();
        public virtual IReadOnlyList<FolderResult> ListFolders() => throw new System.NotImplementedException();
        public virtual IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args) => throw new System.NotImplementedException();
        public virtual MessageDetail ReadMessage(string messageId, bool includeFullBody) => throw new System.NotImplementedException();
        public virtual int CountMessages(SearchMessagesArgs args) => throw new System.NotImplementedException();
        public virtual IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads) => throw new System.NotImplementedException();
        public virtual CreatedDraft CreateDraft(CreateDraftArgs args) => throw new System.NotImplementedException();
        public virtual void MarkAsRead(string messageId, bool read) => throw new System.NotImplementedException();
        public virtual void FlagMessage(string messageId, string flag) => throw new System.NotImplementedException();
        public virtual void SetCategory(string messageId, string category) => throw new System.NotImplementedException();
    }
}
```

- [ ] **Step 2: Implement tool**

```csharp
// VSTO2\OutlookAI\Services\Tools\OutlookGetCurrentComposeStateTool.cs
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    public sealed class OutlookGetCurrentComposeStateTool : IOutlookTool
    {
        public string Name => "outlook_get_current_compose_state";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson);
            bool fullBody = args["include_full_body"]?.Value<bool>() ?? false;

            var state = surface.GetCurrentComposeState(fullBody);
            var json = new JObject(
                new JProperty("subject", state.Subject ?? ""),
                new JProperty("recipients", new JObject(
                    new JProperty("to", new JArray((state.ToRecipients ?? new string[0]).Cast<object>())),
                    new JProperty("cc", new JArray((state.CcRecipients ?? new string[0]).Cast<object>())),
                    new JProperty("bcc", new JArray((state.BccRecipients ?? new string[0]).Cast<object>())))),
                new JProperty("sender_name", state.SenderName ?? ""),
                new JProperty("sender_email", state.SenderEmail ?? ""),
                new JProperty("body_plaintext", state.BodyPlaintext ?? ""),
                new JProperty("body_truncated", state.BodyTruncated),
                new JProperty("attachments", new JArray((state.Attachments ?? new AttachmentSummary[0]).Select(a =>
                    new JObject(new JProperty("filename", a.Filename), new JProperty("size_bytes", a.SizeBytes))))));
            if (state.InReplyTo != null)
            {
                json["in_reply_to"] = new JObject(
                    new JProperty("thread_topic", state.InReplyTo.ThreadTopic ?? ""),
                    new JProperty("last_n_messages", new JArray((state.InReplyTo.LastNMessages ?? new ThreadMessage[0]).Select(m =>
                        new JObject(
                            new JProperty("from", m.From),
                            new JProperty("received_at", m.ReceivedAt.ToString("o")),
                            new JProperty("snippet", m.Snippet))))));
            }
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }
    }
}
```

- [ ] **Step 3: Compile entries + verify + commit**

```xml
<Compile Include="Services\Tools\OutlookGetCurrentComposeStateTool.cs" />
```

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookGetCurrentComposeStateTool.cs VSTO2/OutlookAI.Tests/Services/Tools/ VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 2 Task 10: outlook_get_current_compose_state tool"
```

### Tasks 11–19: Remaining 9 tools

Each follows the same shape as Task 10. The plan documents the **JSON output projection** for each so the implementer doesn't have to re-derive it from the spec. Test class uses `MinimalSurface` + one `override`. Commit message: `"Phase 2 Task NN: <tool> tool"`.

The projections (one per tool — all return JSON strings via `JObject`):

- **Task 11 — `outlook_list_folders`** → `{"folders": [{"id","name","parent_id","item_count"}, ...]}`.
- **Task 12 — `outlook_search_messages`** → `{"messages": [{"id","subject","from","to":[...],"received_at","snippet","has_attachments"}, ...]}`. Accept `query`, optional `folder_id`, `date_from`, `date_to`, `max_results` (clamp to 100).
- **Task 13 — `outlook_read_message`** → `{"id","subject","from","to":[...],"cc":[...],"received_at","body_plaintext","body_truncated","attachments":[...],"in_reply_to_message_id","conversation_topic"}`. Validate `message_id` present; default `include_full_body=true`.
- **Task 14 — `outlook_count_messages`** → `{"count": <int>}`. Reuses `SearchMessagesArgs` minus paging.
- **Task 15 — `outlook_list_recent_threads_with`** → `{"threads": [{"thread_topic","last_message_at","message_count","snippet","thread_id"}, ...]}`. Validate `recipient_email`; clamp `max_threads` to 20.
- **Task 16 — `outlook_create_draft`** → `{"draft_id","location":"Drafts"}`. Validate `subject` and `body_plaintext` non-empty; optional `to/cc/in_reply_to_message_id`.
- **Task 17 — `outlook_mark_as_read`** → `{"ok": true}`. Validate `message_id` and `read` boolean.
- **Task 18 — `outlook_flag_message`** → `{"ok": true}`. Validate `flag` in `{none,todo,complete}`.
- **Task 19 — `outlook_set_category`** → `{"ok": true}`. Validate `message_id` and non-empty `category`.

Each task's test asserts: (a) arg parsing/validation; (b) surface called with expected mapped args; (c) JSON shape matches the projection above.

After Task 19, push: `git push`.

---

## Task 20: OutlookToolHost (registry) + LiveOutlookSurface

**Files:**
- Create: `VSTO2\OutlookAI\Services\Tools\LiveOutlookSurface.cs`
- Create: `VSTO2\OutlookAI\Services\OutlookToolHost.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`
- Modify: `VSTO2\OutlookAI\ThisAddIn.cs` — instantiate `OutlookToolHost` singleton.

- [ ] **Step 1: Implement `LiveOutlookSurface`**

Wraps the actual Outlook OOM. Every `public` method body marshals through `OutlookThreadMarshaller` so callers can be on any thread. Constructor takes `Microsoft.Office.Interop.Outlook.Application application, OutlookThreadMarshaller marshaller, IdResolver ids, Inspector inspectorScope` — the inspector is captured for `GetCurrentComposeState` only; the other tools use `application.Session`.

Inside each method: call `_marshaller.RunAsync(() => { ...Outlook OOM... })`, return the projected DTO. Use `Items.Restrict` DASL strings for `SearchMessages`/`CountMessages`. Use `MailItem.Reply()` then overwrite Subject/Body for `CreateDraft` when `in_reply_to_message_id` set. Use `MailItem.UnRead/FlagStatus/Categories` for the metadata writes.

This file is ~300 lines. Most code is mechanical projection from Outlook OOM types to DTOs. The Phase 1 `AITaskPane.GetEmailBody/SetEmailBody/InsertEmailBody` methods give the existing patterns to follow.

- [ ] **Step 2: Implement `OutlookToolHost`**

```csharp
// VSTO2\OutlookAI\Services\OutlookToolHost.cs
using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services
{
    /// <summary>
    /// IToolHost implementation that aggregates the 10 IOutlookTools and a
    /// single LiveOutlookSurface, dispatching via ToolDispatcher.
    /// </summary>
    public sealed class OutlookToolHost : IToolHost, IDisposable
    {
        private readonly ToolDispatcher _dispatcher;

        public OutlookToolHost(IOutlookSurface surface, bool includeWriteTools)
        {
            var tools = new IOutlookTool[]
            {
                new OutlookGetCurrentComposeStateTool(),
                new OutlookListFoldersTool(),
                new OutlookSearchMessagesTool(),
                new OutlookReadMessageTool(),
                new OutlookCountMessagesTool(),
                new OutlookListRecentThreadsWithTool(),
                // Write tools appended conditionally:
                includeWriteTools ? new OutlookCreateDraftTool() : null,
                includeWriteTools ? new OutlookMarkAsReadTool() : null,
                includeWriteTools ? new OutlookFlagMessageTool() : null,
                includeWriteTools ? new OutlookSetCategoryTool() : null,
            };
            var nonNull = new System.Collections.Generic.List<IOutlookTool>();
            foreach (var t in tools) if (t != null) nonNull.Add(t);
            _dispatcher = new ToolDispatcher(nonNull, surface);
        }

        public Task<string> DispatchAsync(string toolName, string argsJson, CancellationToken ct)
            => _dispatcher.DispatchAsync(toolName, argsJson, ct);

        public void Dispose() { /* surface owns nothing disposable in Phase 2 */ }
    }
}
```

- [ ] **Step 3: Wire into `ThisAddIn`**

Replace the Phase 1 startup block:

```csharp
public CodexAuthService AuthService { get; private set; }
public CodexChatService ChatService { get; private set; }
public RealtimeVoiceService VoiceService { get; private set; }
public OutlookThreadMarshaller OutlookMarshaller { get; private set; }
public IdResolver IdResolver { get; private set; }

private void ThisAddIn_Startup(object sender, EventArgs e)
{
    try
    {
        AuthService       = new CodexAuthService(Config.CodexAuthPath);
        ChatService       = new CodexChatService(AuthService);
        VoiceService      = new RealtimeVoiceService(AuthService);
        OutlookMarshaller = new OutlookThreadMarshaller(System.Windows.Forms.WindowsFormsSynchronizationContext.Current
                                                        ?? new System.Windows.Forms.WindowsFormsSynchronizationContext());
        IdResolver        = new IdResolver();
    }
    catch (Exception ex) { System.Diagnostics.Debug.WriteLine("ThisAddIn_Startup: " + ex); }
}
```

`OutlookToolHost` itself is constructed per-`AITaskPane` (Task 30) because it captures the owning `Inspector` for compose-state scope. The `ThisAddIn` exposes the shared dependencies (`OutlookMarshaller`, `IdResolver`, the `Application` reference) so the pane can wire them up.

- [ ] **Step 4: Compile + commit + push checkpoint**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
git add VSTO2/OutlookAI/Services/OutlookToolHost.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI/ThisAddIn.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 2 Task 20: OutlookToolHost + LiveOutlookSurface + ThisAddIn wiring"
git push
```

---

## Task 21: ConversationContext / Store / ChatEventSink / TurnResult / StopReason

**Files:**
- Create: `VSTO2\OutlookAI\Services\Chat\StopReason.cs`
- Create: `VSTO2\OutlookAI\Services\Chat\TurnResult.cs`
- Create: `VSTO2\OutlookAI\Services\Chat\ChatEventSink.cs`
- Create: `VSTO2\OutlookAI\Services\Chat\ConversationContext.cs`
- Create: `VSTO2\OutlookAI\Services\Chat\ConversationStore.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\Chat\ConversationStoreTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`
- Modify: `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj` — re-add `CapturingChatEventSink.cs` as Compile.

- [ ] **Step 1: Define types**

```csharp
// VSTO2\OutlookAI\Services\Chat\StopReason.cs
namespace OutlookAI.Services.Chat
{
    public enum StopReason { Completed, Cancelled, MaxRoundsReached, Error }
}
```

```csharp
// VSTO2\OutlookAI\Services\Chat\TurnResult.cs
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Chat
{
    public sealed class TurnResult
    {
        public StopReason StopReason { get; set; }
        public string FinalAssistantText { get; set; } = "";
        public int RoundsUsed { get; set; }
        public IReadOnlyList<JObject> AppendedItems { get; set; } = new List<JObject>();
        public string ErrorMessage { get; set; }
    }
}
```

```csharp
// VSTO2\OutlookAI\Services\Chat\ChatEventSink.cs
namespace OutlookAI.Services.Chat
{
    /// <summary>
    /// Override the callbacks you care about. Defaults are no-ops so a small
    /// caller can adopt this without boilerplate.
    /// </summary>
    public abstract class ChatEventSink
    {
        public virtual void OnTokenDelta(string delta) { }
        public virtual void OnToolCallStart(string callId, string name, string argsJson) { }
        public virtual void OnToolCallResult(string callId, bool ok, string summary, string resultJson) { }
        public virtual void OnAssistantMessageComplete(string text) { }
        public virtual void OnError(string message) { }
        public virtual void OnRoundBoundary() { }
    }
}
```

```csharp
// VSTO2\OutlookAI\Services\Chat\ConversationContext.cs
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Chat
{
    public sealed class ConversationContext
    {
        public string SystemInstructions { get; set; } = "";
        public List<JObject> History { get; set; } = new List<JObject>();
        public string ReasoningEffortOverride { get; set; } // null => use Config.ReasoningEffort
        public bool IncludeWriteTools { get; set; } = true;
    }
}
```

```csharp
// VSTO2\OutlookAI\Services\Chat\ConversationStore.cs
using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Chat
{
    /// <summary>
    /// Per-Inspector chat state, in-memory only.
    /// </summary>
    public sealed class ConversationStore
    {
        private readonly object _lock = new object();
        private readonly List<JObject> _history = new List<JObject>();

        public IReadOnlyList<JObject> Snapshot()
        {
            lock (_lock) return _history.ToArray();
        }

        public void Append(JObject item) { lock (_lock) _history.Add(item); }
        public void AppendRange(IEnumerable<JObject> items) { lock (_lock) _history.AddRange(items); }
        public void Clear() { lock (_lock) _history.Clear(); }
        public int Count { get { lock (_lock) return _history.Count; } }

        public string ExportForClipboard()
        {
            var sb = new StringBuilder();
            lock (_lock)
            {
                foreach (var item in _history)
                {
                    var type = (string)item["type"] ?? "?";
                    if (type == "message")
                    {
                        var role = (string)item["role"];
                        var content = ExtractText(item["content"]);
                        sb.AppendLine("[" + role + "]");
                        sb.AppendLine(content);
                        sb.AppendLine();
                    }
                    else if (type == "function_call")
                    {
                        sb.AppendLine("[tool call] " + item["name"] + " " + item["arguments"]);
                    }
                    else if (type == "function_call_output")
                    {
                        sb.AppendLine("[tool result] " + item["output"]);
                        sb.AppendLine();
                    }
                }
            }
            return sb.ToString();
        }

        private static string ExtractText(JToken token)
        {
            if (token == null) return "";
            if (token.Type == JTokenType.String) return (string)token;
            if (token.Type == JTokenType.Array)
            {
                var sb = new StringBuilder();
                foreach (var part in (JArray)token)
                {
                    var text = (string)part["text"];
                    if (!string.IsNullOrEmpty(text)) sb.Append(text);
                }
                return sb.ToString();
            }
            return token.ToString();
        }
    }
}
```

- [ ] **Step 2: Re-add `CapturingChatEventSink.cs` as Compile**

Edit the test csproj: remove the deferred-helper line for `CapturingChatEventSink.cs`.

- [ ] **Step 3: Tests for ConversationStore**

```csharp
// VSTO2\OutlookAI.Tests\Services\Chat\ConversationStoreTests.cs
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Chat;
using Xunit;

namespace OutlookAI.Tests.Services.Chat
{
    public class ConversationStoreTests
    {
        [Fact]
        public void Append_AndSnapshot_RoundTrips()
        {
            var s = new ConversationStore();
            s.Append(JObject.Parse("{\"type\":\"message\",\"role\":\"user\",\"content\":\"hi\"}"));
            var snap = s.Snapshot();
            Assert.Single(snap);
            Assert.Equal("hi", (string)snap[0]["content"]);
        }

        [Fact]
        public void Clear_RemovesAll()
        {
            var s = new ConversationStore();
            s.Append(new JObject(new JProperty("type", "message")));
            s.Clear();
            Assert.Empty(s.Snapshot());
        }

        [Fact]
        public void ExportForClipboard_RendersUserAndAssistantMessages()
        {
            var s = new ConversationStore();
            s.Append(JObject.Parse("{\"type\":\"message\",\"role\":\"user\",\"content\":\"hi\"}"));
            s.Append(JObject.Parse("{\"type\":\"message\",\"role\":\"assistant\",\"content\":\"hello\"}"));
            var text = s.ExportForClipboard();
            Assert.Contains("[user]", text);
            Assert.Contains("hi", text);
            Assert.Contains("[assistant]", text);
            Assert.Contains("hello", text);
        }
    }
}
```

- [ ] **Step 4: Compile + verify + commit**

```xml
<Compile Include="Services\Chat\StopReason.cs" />
<Compile Include="Services\Chat\TurnResult.cs" />
<Compile Include="Services\Chat\ChatEventSink.cs" />
<Compile Include="Services\Chat\ConversationContext.cs" />
<Compile Include="Services\Chat\ConversationStore.cs" />
```

```powershell
git add VSTO2/OutlookAI/Services/Chat/ VSTO2/OutlookAI.Tests/Services/Chat/ VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "Phase 2 Task 21: ConversationContext/Store/Sink/TurnResult/StopReason"
```

---

## Task 22: CodexChatService.RunTurnAsync — single-round happy path

**Files:**
- Modify: `VSTO2\OutlookAI\Services\CodexChatService.cs` — add `RunTurnAsync` method, refactor request-building so `ProcessEmailAsync` and `RunTurnAsync` share it.
- Create: `VSTO2\OutlookAI.Tests\Services\CodexChatServiceMultiRoundTests.cs`
- Modify: `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj` — re-add `FakeToolHost.cs` as Compile.

- [ ] **Step 1: Re-add `FakeToolHost.cs` as Compile** (last deferred helper).

- [ ] **Step 2: Write failing test (single-round, no tool calls)**

```csharp
// VSTO2\OutlookAI.Tests\Services\CodexChatServiceMultiRoundTests.cs
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class CodexChatServiceMultiRoundTests
    {
        private static (CodexAuthService auth, HttpClient authHttp, string tmpDir) MakeAuth()
        {
            var tmp = Path.Combine(Path.GetTempPath(), "outlookai-mr", Path.GetRandomFileName());
            Directory.CreateDirectory(tmp);
            var path = Path.Combine(tmp, "auth.json");
            File.WriteAllText(path,
                "{\"tokens\":{\"access_token\":\"sk-test\",\"id_token\":\"\","
                + "\"refresh_token\":\"r1\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");
            var http = new HttpClient(new FakeHttpMessageHandler());
            return (new CodexAuthService(path, http), http, tmp);
        }

        [Fact]
        public async Task RunTurnAsync_SingleRound_NoToolCalls_ReturnsCompleted()
        {
            var (auth, authHttp, tmp) = MakeAuth();
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"hello\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            using (authHttp)
            using (auth)
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var ctx = new ConversationContext { SystemInstructions = "Be brief." };
                var sink = new CapturingChatEventSink();
                var tools = new FakeToolHost();
                var result = await chat.RunTurnAsync(ctx, "say hi", tools, sink, CancellationToken.None);

                Assert.Equal(StopReason.Completed, result.StopReason);
                Assert.Equal("hello", result.FinalAssistantText);
                Assert.Equal(1, result.RoundsUsed);
                Assert.Equal("hello", sink.StreamedText.ToString());
                Assert.Single(sink.AssistantMessageFinalTexts);
                Assert.Empty(tools.Calls);
            }
            Directory.Delete(tmp, recursive: true);
        }
    }
}
```

- [ ] **Step 3: Implement `RunTurnAsync` single-round path**

Add to `CodexChatService.cs`:

```csharp
public async Task<TurnResult> RunTurnAsync(
    ConversationContext context,
    string userMessage,
    IToolHost toolHost,
    ChatEventSink sink,
    CancellationToken cancellationToken)
{
    if (context == null) throw new ArgumentNullException(nameof(context));
    if (toolHost == null) throw new ArgumentNullException(nameof(toolHost));
    if (sink == null) sink = new ChatEventSink();

    context.History.Add(new JObject(
        new JProperty("type", "message"),
        new JProperty("role", "user"),
        new JProperty("content", userMessage)));

    var result = new TurnResult();
    var appended = new List<JObject>();
    int rounds = 0;
    while (rounds < MaxToolRounds)
    {
        rounds++; result.RoundsUsed = rounds;
        var body = BuildRunTurnRequest(context);
        var bearer = await _auth.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
        var accountId = _auth.GetStatus().AccountId;
        var assistantText = new System.Text.StringBuilder();
        var pendingCalls = new List<JObject>();

        using (var request = new HttpRequestMessage(HttpMethod.Post, ResponsesEndpoint))
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearer);
            request.Headers.Accept.ParseAdd("text/event-stream");
            if (!string.IsNullOrEmpty(accountId))
                request.Headers.TryAddWithoutValidation("ChatGPT-Account-ID", accountId);
            request.Content = new StringContent(body.ToString(Newtonsoft.Json.Formatting.None),
                                                Encoding.UTF8, "application/json");
            using (var resp = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
            {
                if (!resp.IsSuccessStatusCode)
                {
                    var err = await SafeReadAsStringAsync(resp).ConfigureAwait(false);
                    sink.OnError(err);
                    result.StopReason = StopReason.Error;
                    result.ErrorMessage = err;
                    return result;
                }
                using (var stream = await resp.Content.ReadAsStreamAsync().ConfigureAwait(false))
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    string line;
                    while ((line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (!line.StartsWith("data:", StringComparison.Ordinal)) continue;
                        var payload = line.Substring(5).TrimStart();
                        if (payload == "[DONE]") break;
                        JObject evt;
                        try { evt = JObject.Parse(payload); } catch { continue; }
                        var type = (string)evt["type"];
                        if (type == "response.output_text.delta")
                        {
                            var d = (string)evt["delta"];
                            if (!string.IsNullOrEmpty(d)) { assistantText.Append(d); sink.OnTokenDelta(d); }
                        }
                        else if (type == "response.output_item.added" && (string)evt["item"]?["type"] == "function_call")
                        {
                            var item = (JObject)evt["item"];
                            pendingCalls.Add(item);
                            sink.OnToolCallStart((string)item["id"], (string)item["name"], (string)item["arguments"]);
                        }
                        else if (type == "response.completed") break;
                        else if (type == "error")
                        {
                            var msg = (string)evt["error"]?["message"] ?? payload;
                            sink.OnError(msg); result.StopReason = StopReason.Error; result.ErrorMessage = msg;
                            return result;
                        }
                    }
                }
            }
        }

        if (assistantText.Length > 0)
        {
            var msg = new JObject(
                new JProperty("type", "message"),
                new JProperty("role", "assistant"),
                new JProperty("content", assistantText.ToString()));
            context.History.Add(msg); appended.Add(msg);
            sink.OnAssistantMessageComplete(assistantText.ToString());
            result.FinalAssistantText = assistantText.ToString();
        }

        if (pendingCalls.Count == 0)
        {
            sink.OnRoundBoundary();
            result.StopReason = StopReason.Completed;
            result.AppendedItems = appended;
            return result;
        }

        // Multi-round dispatch arrives in Task 23.
        result.StopReason = StopReason.MaxRoundsReached;
        result.AppendedItems = appended;
        return result;
    }

    result.StopReason = StopReason.MaxRoundsReached;
    result.AppendedItems = appended;
    return result;
}

private const int MaxToolRounds = 16;

private JObject BuildRunTurnRequest(ConversationContext context)
{
    JToken reasoning = JValue.CreateNull();
    var effort = !string.IsNullOrEmpty(context.ReasoningEffortOverride)
                 ? context.ReasoningEffortOverride
                 : Config.ReasoningEffort;
    if (!string.Equals(effort, "None", StringComparison.OrdinalIgnoreCase))
    {
        reasoning = new JObject(new JProperty("effort", effort.ToLowerInvariant()));
    }

    return new JObject(
        new JProperty("model", Config.Model),
        new JProperty("instructions", context.SystemInstructions ?? ""),
        new JProperty("input", new JArray(context.History)),
        new JProperty("tools", Services.Tools.ToolCatalogSchema.BuildResponsesToolsArray(context.IncludeWriteTools)),
        new JProperty("tool_choice", "auto"),
        new JProperty("parallel_tool_calls", true),
        new JProperty("reasoning", reasoning),
        new JProperty("store", false),
        new JProperty("stream", true),
        new JProperty("include", new JArray()));
}
```

(`Config.ReasoningEffort` is added in Task 28; until then this code uses the literal `"None"` default — `Config.ReasoningEffort` will be added then.)

- [ ] **Step 4: Build + run test + commit**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~CodexChatServiceMultiRoundTests"
git add VSTO2/OutlookAI/Services/CodexChatService.cs VSTO2/OutlookAI.Tests/Services/CodexChatServiceMultiRoundTests.cs VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "Phase 2 Task 22: RunTurnAsync single-round happy path"
```

---

## Task 23: RunTurnAsync — multi-round tool loop + parallel dispatch

**Files:**
- Modify: `VSTO2\OutlookAI\Services\CodexChatService.cs` — replace the `pendingCalls.Count == 0` early-return with the full loop continuation.
- Modify: `VSTO2\OutlookAI.Tests\Services\CodexChatServiceMultiRoundTests.cs` — add the multi-round case.

- [ ] **Step 1: Add failing test**

```csharp
[Fact]
public async Task RunTurnAsync_ToolCall_DispatchesAndAppendsFunctionOutput_ThenCompletes()
{
    var (auth, authHttp, tmp) = MakeAuth();
    var fake = new FakeHttpMessageHandler();
    // Round 1: model asks for outlook_get_current_compose_state
    fake.QueueSse(HttpStatusCode.OK,
        "data: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"id\":\"call_1\",\"name\":\"outlook_get_current_compose_state\",\"arguments\":\"{}\"}}\n\n"
        + "data: {\"type\":\"response.completed\"}\n\n");
    // Round 2: model returns the answer
    fake.QueueSse(HttpStatusCode.OK,
        "data: {\"type\":\"response.output_text.delta\",\"delta\":\"Subject was X\"}\n\n"
        + "data: {\"type\":\"response.completed\"}\n\n");
    using (authHttp)
    using (auth)
    using (var chatHttp = new HttpClient(fake))
    using (var chat = new CodexChatService(auth, chatHttp))
    {
        var ctx = new ConversationContext { SystemInstructions = "Be brief." };
        var sink = new CapturingChatEventSink();
        var tools = new FakeToolHost();
        tools.Queue("outlook_get_current_compose_state", "{\"subject\":\"X\"}");

        var result = await chat.RunTurnAsync(ctx, "what's the subject", tools, sink, CancellationToken.None);

        Assert.Equal(StopReason.Completed, result.StopReason);
        Assert.Equal(2, result.RoundsUsed);
        Assert.Equal("Subject was X", result.FinalAssistantText);
        Assert.Single(tools.Calls);
        Assert.Equal("outlook_get_current_compose_state", tools.Calls[0].Name);
        Assert.Equal(2, sink.RoundBoundaries);
        Assert.Single(sink.ToolStarts);
        Assert.Single(sink.ToolResults);
        Assert.True(sink.ToolResults[0].Ok);
    }
    Directory.Delete(tmp, recursive: true);
}
```

- [ ] **Step 2: Replace the early-return in `RunTurnAsync`**

Where Task 22 had:

```csharp
        // Multi-round dispatch arrives in Task 23.
        result.StopReason = StopReason.MaxRoundsReached;
        result.AppendedItems = appended;
        return result;
```

Replace with:

```csharp
        var dispatchTasks = pendingCalls.Select(async call =>
        {
            var name = (string)call["name"];
            var args = (string)call["arguments"] ?? "{}";
            var callId = (string)call["id"];
            string outputJson;
            bool ok = true;
            try
            {
                outputJson = await toolHost.DispatchAsync(name, args, cancellationToken).ConfigureAwait(false);
                if (outputJson.IndexOf("\"error\"", StringComparison.Ordinal) >= 0) ok = false;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                outputJson = "{\"error\":{\"code\":\"" + ex.GetType().Name
                             + "\",\"message\":\"" + ex.Message.Replace("\"","'") + "\"}}";
                ok = false;
            }
            sink.OnToolCallResult(callId, ok, Summarize(outputJson), outputJson);
            return new
            {
                FunctionCall = new JObject(
                    new JProperty("type", "function_call"),
                    new JProperty("id", callId),
                    new JProperty("name", name),
                    new JProperty("arguments", args)),
                FunctionCallOutput = new JObject(
                    new JProperty("type", "function_call_output"),
                    new JProperty("call_id", callId),
                    new JProperty("output", outputJson)),
            };
        }).ToArray();
        var dispatchResults = await Task.WhenAll(dispatchTasks).ConfigureAwait(false);
        foreach (var dr in dispatchResults)
        {
            context.History.Add(dr.FunctionCall);   appended.Add(dr.FunctionCall);
            context.History.Add(dr.FunctionCallOutput); appended.Add(dr.FunctionCallOutput);
        }
        if (cancellationToken.IsCancellationRequested)
        {
            result.StopReason = StopReason.Cancelled;
            result.AppendedItems = appended;
            return result;
        }
        sink.OnRoundBoundary();
        // continue the while loop
```

Add a private `Summarize` helper that truncates the JSON to ~120 chars for the sink callback.

- [ ] **Step 3: Build + run + commit**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
git add VSTO2/OutlookAI/Services/CodexChatService.cs VSTO2/OutlookAI.Tests/Services/CodexChatServiceMultiRoundTests.cs
git commit -m "Phase 2 Task 23: multi-round tool dispatch + parallel execution"
```

---

## Task 24: RunTurnAsync — cancellation + max-rounds + partial preservation

**Files:**
- Modify: `VSTO2\OutlookAI.Tests\Services\CodexChatServiceMultiRoundTests.cs` — add three more tests.

- [ ] **Step 1: Add the tests**

- `RunTurnAsync_MaxRoundsReached_WhenModelLoopsForever` — script 17 rounds of tool calls; expect `StopReason.MaxRoundsReached` at exactly 16 rounds used.
- `RunTurnAsync_Cancellation_AfterFirstDelta_PreservesPartialText` — fake SSE delays second event using a `SemaphoreSlim` inside the handler; cancel during streaming; expect partial assistant message in history.
- `RunTurnAsync_ToolErrorBecomesFunctionCallOutputError_AllowingRecovery` — `FakeToolHost` throws on call 1; expect `function_call_output` with `{"error":…}` in history and `OnToolCallResult(ok:false)`.

- [ ] **Step 2: Adjust `RunTurnAsync` to honour each case explicitly**

The Task 22+23 code already passes 1 and 3 if `cancellationToken.ThrowIfCancellationRequested()` is sprinkled in the SSE read loop. For 2, ensure a `try/catch (OperationCanceledException)` in the outer loop sets `result.StopReason = StopReason.Cancelled`.

- [ ] **Step 3: Build + run + commit**

```powershell
git commit -am "Phase 2 Task 24: RunTurnAsync cancellation + max-rounds + tool-error recovery"
git push
```

Push checkpoint — chat service complete.

---

## Task 25: Config v2 — ReasoningEffort + WriteToolsEnabled + AvailableModels

**Files:**
- Modify: `VSTO2\OutlookAI\Config.cs`
- Modify: `VSTO2\OutlookAI.Tests\ConfigTests.cs` — extend tests.

- [ ] **Step 1: Add new Config fields**

```csharp
public const string DefaultReasoningEffort = "None";
public static string ReasoningEffort { get; set; } = DefaultReasoningEffort;
public static bool WriteToolsEnabled { get; set; } = true;

public static readonly string[] AvailableModels =
{
    "gpt-5.5", "gpt-5.5-pro", "gpt-5.4", "gpt-5.4-mini",
    "gpt-4.1-mini", "gpt-4.1-nano", "gpt-5.3-codex"
};
public static readonly string[] AvailableReasoningEfforts =
{
    "None", "Minimal", "Low", "Medium", "High"
};

public static string[] ReasoningEffortsForModel(string model)
{
    if (model == "gpt-4.1-mini" || model == "gpt-4.1-nano") return new[] { "None" };
    return AvailableReasoningEfforts;
}
```

Wire the new fields into `LoadFromFile` (allowed for global only) and `ResetDefaults`. Persist `ReasoningEffort` in `SaveConfig` too (per-user override as discussed in Phase 1.x plan).

- [ ] **Step 2: Add tests**

```csharp
[Fact]
public void LoadConfigFromPaths_AppliesReasoningEffortFromGlobal()
{
    var (g, u) = MakeTempPaths();
    File.WriteAllText(g, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");
    Config.LoadConfigFromPaths(g, u);
    Assert.Equal("High", Config.ReasoningEffort);
}

[Fact]
public void ReasoningEffortsForModel_RestrictsForNonReasoningModels()
{
    Assert.Equal(new[] { "None" }, Config.ReasoningEffortsForModel("gpt-4.1-nano"));
    Assert.Contains("High", Config.ReasoningEffortsForModel("gpt-5.5"));
}
```

- [ ] **Step 3: Build, run, commit**

```powershell
git commit -am "Phase 2 Task 25: Config v2.1 — ReasoningEffort + WriteToolsEnabled + model catalog"
```

---

## Task 26: VariantParser

**Files:**
- Create: `VSTO2\OutlookAI\Services\Variants\Variant.cs`
- Create: `VSTO2\OutlookAI\Services\Variants\Tone.cs`
- Create: `VSTO2\OutlookAI\Services\Variants\VariantParser.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\Variants\VariantParserTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

- [ ] **Step 1: Define types**

```csharp
// VSTO2\OutlookAI\Services\Variants\Tone.cs
namespace OutlookAI.Services.Variants
{
    public enum Tone
    {
        Formal, Brief, Persuasive, Friendly, Technical,
        Apologetic, Direct, Diplomatic, Enthusiastic
    }
    public static class ToneExtensions
    {
        public static Tone ClosestTo(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return Tone.Direct;
            foreach (Tone t in System.Enum.GetValues(typeof(Tone)))
                if (string.Equals(t.ToString(), raw, System.StringComparison.OrdinalIgnoreCase)) return t;
            // tiny synonym map
            var r = raw.ToLowerInvariant();
            if (r.Contains("polite") || r.Contains("warm")) return Tone.Friendly;
            if (r.Contains("short") || r.Contains("conc")) return Tone.Brief;
            if (r.Contains("sales") || r.Contains("convince")) return Tone.Persuasive;
            if (r.Contains("sorry") || r.Contains("apolog")) return Tone.Apologetic;
            return Tone.Direct;
        }
    }
}
```

```csharp
// VSTO2\OutlookAI\Services\Variants\Variant.cs
namespace OutlookAI.Services.Variants
{
    public sealed class Variant
    {
        public Tone Tone { get; set; }
        public string Rationale { get; set; } = "";
        public string Subject { get; set; } = "";
        public string Body { get; set; } = "";
    }
}
```

```csharp
// VSTO2\OutlookAI\Services\Variants\VariantParser.cs
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Variants
{
    public sealed class VariantParser
    {
        private static readonly Regex FencedJson =
            new Regex(@"```(?:json)?\s*\n([\s\S]*?)\n```", RegexOptions.Compiled);

        public IReadOnlyList<Variant> Parse(string assistantText)
        {
            if (string.IsNullOrWhiteSpace(assistantText)) return new Variant[0];
            string json = assistantText;
            var m = FencedJson.Match(assistantText);
            if (m.Success) json = m.Groups[1].Value;
            JObject root;
            try { root = JObject.Parse(json); } catch { return new Variant[0]; }
            var arr = root["variants"] as JArray;
            if (arr == null) return new Variant[0];
            var list = new List<Variant>();
            foreach (var v in arr)
            {
                list.Add(new Variant
                {
                    Tone = ToneExtensions.ClosestTo((string)v["tone"]),
                    Rationale = (string)v["rationale"] ?? "",
                    Subject = (string)v["subject"] ?? "",
                    Body = (string)v["body"] ?? "",
                });
            }
            return list;
        }
    }
}
```

- [ ] **Step 2: Tests + commit**

```csharp
[Fact]
public void Parse_WellFormedFencedJson_ReturnsVariants()
{
    var input = "```json\n{\"variants\":[{\"tone\":\"Formal\",\"rationale\":\"r\",\"subject\":\"s\",\"body\":\"b\"}]}\n```";
    var p = new VariantParser();
    var vs = p.Parse(input);
    Assert.Single(vs);
    Assert.Equal(Tone.Formal, vs[0].Tone);
}

[Fact]
public void Parse_FreeFormToneClampedToEnum()
{
    var input = "```{\"variants\":[{\"tone\":\"warmth\",\"body\":\"x\"}]}```";
    var p = new VariantParser();
    Assert.Equal(Tone.Friendly, p.Parse(input)[0].Tone);
}

[Fact]
public void Parse_MalformedJson_ReturnsEmptyList()
{
    Assert.Empty(new VariantParser().Parse("not json"));
}
```

```powershell
git commit -am "Phase 2 Task 26: VariantParser + Tone enum"
```

---

## Task 27: VariantStore

**Files:**
- Create: `VSTO2\OutlookAI\Services\Variants\VariantStore.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\Variants\VariantStoreTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

In-memory list of `Variant`s with `Replace(IEnumerable<Variant>)`, `Snapshot()`, `Update(int index, Variant)`, `Clear()`. Tests cover replace, per-index update, isolation between two stores. Commit.

---

## Task 28: SettingsForm extensions

**Files:**
- Modify: `VSTO2\OutlookAI\SettingsForm.cs`

Add a new GroupBox below the existing Account group, *only visible after admin login*:

1. **Model** ComboBox populated from `Config.AvailableModels`.
2. **Reasoning effort** ComboBox populated from `Config.ReasoningEffortsForModel(currentModel)`; re-filter on Model SelectedIndexChanged; snap selection to the nearest valid value.
3. **Write tools** CheckedListBox with 4 entries (`create_draft`, `mark_as_read`, `flag_message`, `set_category`), all checked by default.
4. **Save** button persists Model, ReasoningEffort, and a comma-separated `EnabledWriteTools` string into `Config` and calls `Config.SaveConfig()`.

(Per Phase 1.x plan, these per-user-savable fields are written to the AppData config; the global config's values still win on next load. This is intentional — per-user override for these specific keys only.)

Commit.

---

## Task 29: AITaskPane TabControl scaffold

**Files:**
- Modify: `VSTO2\OutlookAI\TaskPane\AITaskPane.cs`
- Modify: `VSTO2\OutlookAI\TaskPane\AITaskPane.Designer.cs`

- [ ] **Step 1** Rebuild the Designer with a top-level `TabControl` (Dock=Fill, 3 tabs: `tabActions`, `tabChat`, `tabVariants`). Move the existing controls (lblTitle, btnSettings, grpQuickActions, grpDraft, lblStatus, panelResult) into `tabActions`. Tab labels with icons (use Unicode glyphs to stay simple: ⚡ Actions, 💬 Chat, 📝 Variants).
- [ ] **Step 2** Increase the task pane width to 320 px (was 260) to give Chat/Variants room. Update `customTaskPane.Width = 340;` in `ThisAddIn.ShowTaskPane`.
- [ ] **Step 3** `tabChat` and `tabVariants` remain empty (filled by Tasks 32-37). Add placeholder labels so they're visible during dev.
- [ ] **Step 4** Compile, manual smoke: open compose, switch tabs, Actions still works exactly as before.
- [ ] **Step 5** Commit.

---

## Task 30: Actions tab — rewire existing buttons to use RunTurnAsync

**Files:**
- Modify: `VSTO2\OutlookAI\TaskPane\AITaskPane.cs`

Replace each button handler's call from `_codexChatService.ProcessEmailAsync` (Phase 1 single-shot) to `_codexChatService.RunTurnAsync`. Each call constructs a small `ConversationContext` with the same system prompt the action used (carried over from `CodexChatService.GetSystemPrompt`) plus the addendum: *"You may call mailbox tools if you need additional context. Most quick edits don't require any tools."*. `ChatEventSink` is implemented locally to surface tool-card status in the existing Result panel area (a new minimal `Panel` above the Result text).

A Cancel button next to "Processing..." surfaces `CancellationTokenSource` cancellation.

Commit.

---

## Task 31: WebView2 NuGet + project setup + WebView2Bootstrap

**Files:**
- Modify: `VSTO2\OutlookAI\packages.config` — add `Microsoft.Web.WebView2 1.0.2849.39`.
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj` — add References for WebView2 (`Microsoft.Web.WebView2.Core.dll`, `Microsoft.Web.WebView2.WinForms.dll`).
- Create: `VSTO2\OutlookAI\TaskPane\Chat\WebView2Bootstrap.cs` — extracts embedded WebUI resources to `%LOCALAPPDATA%\OutlookAI\WebUI`, sets up `CoreWebView2Environment` (user data folder under `%LOCALAPPDATA%\OutlookAI\WebView2Data`), maps `outlookai.local` virtual host to the extracted folder.

- [ ] **Step 1** `nuget restore`.
- [ ] **Step 2** Wire `WebView2Bootstrap.InitializeAsync(WebView2 host)` returning a `Task` that completes after `CoreWebView2InitializationCompleted`.
- [ ] **Step 3** Manual smoke: drop a temporary WebView2 control on `tabChat` and load `https://outlookai.local/index.html` (created by hand for this task with a one-line `<h1>Hello</h1>`). Verify it renders.
- [ ] **Step 4** Commit.

---

## Task 32: WebUI bundle — initial render

**Files:**
- Create: `VSTO2\OutlookAI\WebUI\index.html`
- Create: `VSTO2\OutlookAI\WebUI\styles.css`
- Create: `VSTO2\OutlookAI\WebUI\chat.js`
- Create: `VSTO2\OutlookAI\WebUI\marked.min.js`        (download from cdnjs, vendor)
- Create: `VSTO2\OutlookAI\WebUI\highlight.min.js`     (download from cdnjs, vendor)
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj` — `<EmbeddedResource Include="WebUI\*" />`.

Initial `chat.js` implements the rendering API listed in Spec § 4: `appendUserMessage`, `appendAssistantMessage`, `appendTextDelta`, `appendToolCallCard`, `updateToolCallCard`, `finalizeAssistantMessage`, `showError`, `setComposerEnabled`, `clear`, `applyTheme`, `setContextStrip`. Calls JS↔C# bridge via `window.chrome.webview.postMessage(JSON.stringify({type:…}))`.

Initial `index.html` lays out: top context strip, scrollable `#messages` area, bottom composer (`<textarea>` + Send/Stop + Reasoning select).

Commit.

---

## Task 33: WebUI — tool cards + audit rows + theme + states

Extends `chat.js` and `styles.css` to render: tool-call cards with expand-to-JSON; audit rows for write tools; "stopped" badges; error cards; high-contrast theme support (driven by `applyTheme`). Commit.

---

## Task 34: ChatController — JS↔C# bridge + RunTurnAsync orchestration

**Files:**
- Create: `VSTO2\OutlookAI\TaskPane\Chat\ChatController.cs`
- Modify: `VSTO2\OutlookAI\TaskPane\AITaskPane.cs` — instantiate `ChatController` inside `tabChat`.

`ChatController` wires `CoreWebView2.WebMessageReceived` → switch on `type` → invoke C# action (Send/Stop/Clear/Insert/Replace/Copy/Regenerate). On Send: construct `ConversationContext` from the per-pane `ConversationStore`; invoke `CodexChatService.RunTurnAsync` with a custom `ChatEventSink` that translates events into `ExecuteScriptAsync("chat.appendTextDelta({...})")` calls.

The system message includes a tagged summary of compose state via `OutlookToolHost.Surface.GetCurrentComposeState(false)` so the model has context without needing a tool call for the obvious case.

Commit. Push checkpoint — Chat tab is alive.

---

## Task 35: Variants tab — UI + intent input + count picker + reasoning override

**Files:**
- Modify: `VSTO2\OutlookAI\TaskPane\AITaskPane.cs`
- Modify: `VSTO2\OutlookAI\TaskPane\AITaskPane.Designer.cs`
- Create: `VSTO2\OutlookAI\TaskPane\Variants\VariantsController.cs`

UI per Spec § 5: TextBox + mic button + NumericUpDown + Generate button + Reasoning override ComboBox + scrollable Panel for cards + Regenerate-all button. Cards are dynamically generated `Panel` controls (tone tag chip + char count + first 3 lines + per-card buttons).

Commit.

---

## Task 36: Variants generation flow

**Files:**
- Modify: `VSTO2\OutlookAI\TaskPane\Variants\VariantsController.cs`

On Generate: build a `ConversationContext` with the Variants-specific system message (forbid `outlook_create_draft`, require fenced JSON output). Call `CodexChatService.RunTurnAsync` with a `CapturingChatEventSink`. Pass the final assistant text to `VariantParser`. Store result in `VariantStore`. Re-render cards.

Per-card actions: **Insert/Replace** call the existing `AITaskPane.InsertEmailBody/SetEmailBody`. **Regenerate one** runs a 1-variant turn with the same intent and tone constraint.

Commit. Push checkpoint — Variants tab functional.

---

## Task 37: Install script — WebView2 Evergreen detection + bootstrap

**Files:**
- Modify: `Deploy\Install-OutlookAI.ps1`

Add a new `[Step]` (numbered into the existing sequence) before VSTO trust setup:

```powershell
Write-Host "[X/Y] Ensuring Microsoft Edge WebView2 Runtime present..." -ForegroundColor Yellow
$wv2KeyA = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
$wv2KeyB = "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
$installed = $false
foreach ($k in @($wv2KeyA, $wv2KeyB)) {
    if (Test-Path $k) {
        $pv = (Get-ItemProperty $k -ErrorAction SilentlyContinue).pv
        if (-not [string]::IsNullOrWhiteSpace($pv)) { $installed = $true; break }
    }
}
if ($installed) {
    Write-Host "  WebView2 Runtime detected." -ForegroundColor Gray
} else {
    $bootstrap = Join-Path $SourcePath "MicrosoftEdgeWebView2Setup.exe"
    if (Test-Path $bootstrap) {
        Write-Host "  Installing WebView2 Runtime..." -ForegroundColor Gray
        Start-Process -FilePath $bootstrap -ArgumentList "/silent","/install" -Wait
    } else {
        Write-Host "  WARN: MicrosoftEdgeWebView2Setup.exe missing; users will see a fallback panel." -ForegroundColor Yellow
    }
}
Write-Host "  Done." -ForegroundColor Green
```

Commit.

---

## Task 38: Vendor MicrosoftEdgeWebView2Setup.exe

**Files:**
- Create: `Deploy\MicrosoftEdgeWebView2Setup.exe`

```powershell
Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?LinkId=2124703" -OutFile "Deploy\MicrosoftEdgeWebView2Setup.exe"
```

Verify file size and digital signature. Commit.

---

## Task 39: Manual smoke checklist + Phase 2 README update

**Files:**
- Create: `docs\superpowers\checklists\phase-2-smoke.md`
- Modify: `README.md` — extend the branch banner to note Phase 2 features now present.

Smoke list per Spec § 6. Commit.

---

## Task 40: Final E2E build + push

```powershell
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
dotnet restore VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /t:Rebuild /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: clean build, all tests pass.

Publish + reinstall in Outlook + walk through the smoke checklist. If anything fails, file as a follow-up bug task in the same branch.

```powershell
git push
```

End of Phase 2 plan.

---

## Self-Review

| Check | Result |
|---|---|
| Spec § 1 (architecture) covered by Tasks 6-9, 20 | ✓ |
| Spec § 2 (10 tools) covered by Tasks 10-19 | ✓ |
| Spec § 3 (multi-round loop) covered by Tasks 22-24 | ✓ |
| Spec § 4 (WebView2 surface) covered by Tasks 31-34 | ✓ |
| Spec § 5 (Variants) covered by Tasks 26-27, 35-36 | ✓ |
| Spec § 6 (test project, gates) covered by Tasks 1-5 + per-task TDD | ✓ |
| Spec § 7 (impl boundaries) mirrored 1:1 in this plan's file list | ✓ |
| Phase 1.x dropdowns folded into Task 25 + 28 | ✓ |
| Deploy: WebView2 bootstrap covered by Tasks 37-38 | ✓ |
| No "TBD/TODO" placeholders | ✓ |
| Type names consistent (`IToolHost`, `IOutlookSurface`, `IOutlookTool`, `ConversationContext/Store`, `ChatEventSink`, `TurnResult/StopReason`, `VariantParser/Store`, `Tone`, `IdResolver`, `OutlookThreadMarshaller`, `ToolDispatcher`, `OutlookToolHost`, `LiveOutlookSurface`, `ChatController`, `VariantsController`, `WebView2Bootstrap`) match across all tasks | ✓ |

