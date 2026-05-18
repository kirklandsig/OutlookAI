# Phase 3b: Non-Blocking Mailbox Search Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop Outlook from freezing during `outlook_search_messages` / `outlook_count_messages` by switching the primary engine to `Application.AdvancedSearch`, keeping the iterative path as an automatic fallback, and adding a real Stop affordance for the user.

**Architecture:** New `OutlookAdvancedSearchRunner` wraps `Application.AdvancedSearch` and the `AdvancedSearchComplete` event behind an `IAdvancedSearchHost` test seam, exposing a `Task<AdvancedSearchResult>` API with timeout and cancellation. A new `IFolderClassifier` centralises the system-folder skip list (with the real names Outlook actually uses, e.g. `Junk E-mail` with hyphen). A new `SearchResultProjector` does pure sort + top-N + skip-list + deferred-snippet projection. `LiveOutlookSurface.SearchMessages` becomes a thin orchestrator: try primary, on failure/timeout fall back to a yielding per-folder loop that releases the UI thread between folders. The compact chat status UI gains a Stop link and a live time counter wired to the existing `ChatSession.CancelCurrent()` chain.

**Tech Stack:** C# .NET Framework 4.7.2 (VSTO), xUnit, Outlook OOM (`Microsoft.Office.Interop.Outlook`), WebView2 for the chat UI.

**Spec:** `docs/superpowers/specs/2026-05-17-phase-3b-non-blocking-search-design.md`

---

## File Structure

**New files:**

- `VSTO2/OutlookAI/Services/Tools/IFolderClassifier.cs` — `IFolderClassifier` interface + `FolderClassifier` default implementation. Knows the system-folder skip list and the non-mail-default-item-type rule.
- `VSTO2/OutlookAI/Services/Tools/MessageProjectionInput.cs` — DTO carrying just the fields the projector needs, with a `Func<string>` `SnippetFactory` so body access can be deferred.
- `VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs` — Pure static class. Sort, classifier-filter, top-N, deferred snippet, COMException-per-item resilience.
- `VSTO2/OutlookAI/Services/Tools/SearchScopeFormatter.cs` — Pure helper that formats a list of folder paths into Outlook's comma-separated single-quoted scope string.
- `VSTO2/OutlookAI/Services/Tools/IAdvancedSearchHost.cs` — Test seam. Production wraps `Application.AdvancedSearch` + `AdvancedSearchComplete` event; tests raise events manually.
- `VSTO2/OutlookAI/Services/Tools/LiveAdvancedSearchHost.cs` — Production `IAdvancedSearchHost` that uses the real Outlook application.
- `VSTO2/OutlookAI/Services/Tools/OutlookAdvancedSearchRunner.cs` — `Task<AdvancedSearchResult> RunAsync(...)` wrapper around `IAdvancedSearchHost`, with timeout, cancellation, tag dispatch, and semaphore serialisation.
- `VSTO2/OutlookAI.Tests/Services/Tools/FolderClassifierTests.cs`
- `VSTO2/OutlookAI.Tests/Services/Tools/SearchResultProjectorTests.cs`
- `VSTO2/OutlookAI.Tests/Services/Tools/SearchScopeFormatterTests.cs`
- `VSTO2/OutlookAI.Tests/Services/Tools/OutlookAdvancedSearchRunnerTests.cs`
- `VSTO2/OutlookAI.Tests/Services/Tools/FakeAdvancedSearchHost.cs` — Test double used by the runner tests.

**Modified files:**

- `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs` — Add `CancellationToken ct = default` to `SearchMessages` and `CountMessages`.
- `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs` — Re-implement `SearchMessages` / `CountMessages` as orchestrators: try primary (runner), fall back to per-folder yielding. Replace `ShouldSkipAllMailFolder` calls with `IFolderClassifier`. Defer body snippet.
- `VSTO2/OutlookAI/Services/Tools/SearchExecutionHelper.cs` — Remove the old `ShouldSkipAllMailFolder`. Keep `SortDescending` + `MergeAndSortSearchResults` if still used by the fallback (or delete if subsumed by projector).
- `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs` — Plumb `CancellationToken` through to `surface.SearchMessages(args, ct)`. Wrap `OperationCanceledException` into structured cancel envelope.
- `VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs` — Same plumbing.
- `VSTO2/OutlookAI/ThisAddIn.cs` — Construct `LiveAdvancedSearchHost` and `OutlookAdvancedSearchRunner`; pass runner into `LiveOutlookSurface`; dispose runner on shutdown.
- `VSTO2/OutlookAI.Tests/Services/Tools/SearchExecutionHelperTests.cs` — Update / extend skip-list assertions to match `FolderClassifier` behavior with real folder names (e.g., `Junk E-mail`).
- `VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs` — Update to assert `CancellationToken` is plumbed and cancel-envelope is emitted on `OperationCanceledException`.
- `VSTO2/OutlookAI.Tests/Services/Tools/OutlookCountMessagesToolTests.cs` — Same.
- `VSTO2/OutlookAI/TaskPane/InboxCopilot/Assets/inbox-copilot.js` (or equivalent compact-status JS) — Render Stop link + time counter in the compact status line.
- `VSTO2/OutlookAI/TaskPane/InboxCopilot/Assets/inbox-copilot.css` — Tiny styling for the Stop link.

**Build & test commands** (reuse throughout this plan):

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Run from `C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-codex-oauth-migration`.

---

## Task 1: Add `CancellationToken` to `IOutlookSurface.SearchMessages` and `CountMessages`

This is the foundation. All later tasks need a CT in the signature. Default to `CancellationToken.None` so existing callers compile unchanged.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookCountMessagesToolTests.cs`

- [ ] **Step 1: Write the failing test for CT plumbing in SearchMessages tool**

Add to `OutlookSearchMessagesToolTests.cs`:

```csharp
[Fact]
public async Task Execute_PassesCancellationTokenThroughToSurface()
{
    CancellationToken observedCt = default;
    var surface = new FakeOutlookSurface
    {
        OnSearchMessages = (args, ct) =>
        {
            observedCt = ct;
            return new List<MessageSummary>();
        }
    };
    var tool = new OutlookSearchMessagesTool();
    using var cts = new CancellationTokenSource();
    await tool.ExecuteAsync("{}", surface, cts.Token);
    Assert.Equal(cts.Token, observedCt);
}

[Fact]
public async Task Execute_OnUserCancellation_EmitsStructuredCancelEnvelope()
{
    var surface = new FakeOutlookSurface
    {
        OnSearchMessages = (args, ct) => throw new OperationCanceledException(ct)
    };
    var tool = new OutlookSearchMessagesTool();
    using var cts = new CancellationTokenSource();
    cts.Cancel();
    var json = await tool.ExecuteAsync("{}", surface, cts.Token);
    Assert.Contains("\"error\"", json);
    Assert.Contains("\"code\":\"cancelled\"", json);
}
```

If `FakeOutlookSurface` does not exist with `OnSearchMessages` matching this new shape, update it now. Locate its current definition in the tests project and add/widen the delegate:

```csharp
public Func<SearchMessagesArgs, CancellationToken, IReadOnlyList<MessageSummary>> OnSearchMessages { get; set; }

public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default)
    => OnSearchMessages != null ? OnSearchMessages(args, ct) : new List<MessageSummary>();
```

- [ ] **Step 2: Run targeted tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookSearchMessagesToolTests"
```

Expected: build errors first (signature mismatch), then test failures once the signature compiles.

- [ ] **Step 3: Update `IOutlookSurface.SearchMessages` and `CountMessages` signatures**

In `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`:

```csharp
IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default);
int CountMessages(SearchMessagesArgs args, CancellationToken ct = default);
```

Add `using System.Threading;` to the top of the file if not already present.

- [ ] **Step 4: Update `LiveOutlookSurface` to accept and ignore the CT for now**

In `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`, change:

```csharp
public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args) =>
```

to:

```csharp
public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default) =>
```

Same for `CountMessages`. The body does not change yet; the CT will be wired in later tasks.

- [ ] **Step 5: Update `OutlookSearchMessagesTool.ExecuteAsync` to plumb CT and emit cancel envelope**

Replace the body of `ExecuteAsync` in `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs`:

```csharp
public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
{
    ct.ThrowIfCancellationRequested();
    var search = SearchMessagesArgsParser.ParseSearch(argsJson);

    IReadOnlyList<MessageSummary> hits;
    try
    {
        hits = surface.SearchMessages(search, ct) ?? new MessageSummary[0];
    }
    catch (OperationCanceledException)
    {
        return Task.FromResult(BuildError("cancelled", "Search cancelled by user."));
    }

    var json = new JObject(
        new JProperty("messages", new JArray(hits.Select(m =>
            new JObject(
                new JProperty("id", m.Id ?? ""),
                new JProperty("subject", m.Subject ?? ""),
                new JProperty("from", m.From ?? ""),
                new JProperty("to", new JArray((m.To ?? new string[0]).Cast<object>())),
                new JProperty("received_at", m.ReceivedAt.ToString("o")),
                new JProperty("snippet", m.Snippet ?? ""),
                new JProperty("has_attachments", m.HasAttachments))))));
    return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
}
```

`BuildError` already exists in the same class.

- [ ] **Step 6: Do the same for `OutlookCountMessagesTool.ExecuteAsync`**

In `VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs`, mirror the same pattern: `try { ... surface.CountMessages(args, ct) ... } catch (OperationCanceledException) { return cancel envelope; }`.

Add a matching test in `OutlookCountMessagesToolTests.cs`:

```csharp
[Fact]
public async Task Execute_PassesCancellationTokenThroughToSurface()
{
    CancellationToken observedCt = default;
    var surface = new FakeOutlookSurface
    {
        OnCountMessages = (args, ct) => { observedCt = ct; return 0; }
    };
    var tool = new OutlookCountMessagesTool();
    using var cts = new CancellationTokenSource();
    await tool.ExecuteAsync("{}", surface, cts.Token);
    Assert.Equal(cts.Token, observedCt);
}
```

Widen `FakeOutlookSurface.OnCountMessages` to the same `Func<SearchMessagesArgs, CancellationToken, int>` shape.

- [ ] **Step 7: Build and run full test suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Then:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: build clean, all tests pass including the new four CT-plumbing tests.

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookCountMessagesToolTests.cs
git commit -m "feat(search): plumb CancellationToken through search/count tools" -m "Adds an optional CancellationToken to IOutlookSurface.SearchMessages and CountMessages, plumbs it through OutlookSearchMessagesTool and OutlookCountMessagesTool, and has the tools emit a structured {error: cancelled} envelope on OperationCanceledException so the model sees a clear cancel signal instead of an ambiguous empty result. Foundation for Phase 3b non-blocking search."
```

---

## Task 2: Add `IFolderClassifier` with full Outlook system-folder skip list

Centralises every skip decision. Includes the real Outlook folder names (`Junk E-mail` with hyphen, `Conversation Action Settings`, etc.) that today's narrow skip list misses.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/IFolderClassifier.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/FolderClassifierTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/FolderClassifierTests.cs`:

```csharp
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class FolderClassifierTests
    {
        private readonly IFolderClassifier _classifier = new FolderClassifier();

        [Theory]
        [InlineData("Deleted Items")]
        [InlineData("Junk E-mail")]      // real Outlook name, with hyphen
        [InlineData("Junk Email")]       // older locale variant
        [InlineData("Drafts")]
        [InlineData("Outbox")]
        [InlineData("Sync Issues")]
        [InlineData("Sync Issues (This computer only)")]
        [InlineData("Conflicts")]
        [InlineData("Local Failures")]
        [InlineData("Server Failures")]
        [InlineData("RSS Feeds")]
        [InlineData("RSS Subscriptions")]
        [InlineData("Conversation Action Settings")]
        [InlineData("Conversation History")]
        [InlineData("Quick Step Settings")]
        [InlineData("News Feed")]
        [InlineData("Feeds")]
        [InlineData("Files")]
        [InlineData("Detected Items")]
        [InlineData("Working Set")]
        [InlineData("Yammer Root")]
        public void IsSystemFolder_KnownSystemNames_ReturnsTrue(string name)
        {
            Assert.True(_classifier.IsSystemFolder(name, defaultItemTypeIsMail: true));
        }

        [Theory]
        [InlineData("Inbox")]
        [InlineData("Archive")]
        [InlineData("Sent Items")]
        [InlineData("Projects")]
        [InlineData("Receipts")]
        [InlineData("Important Office emails")]
        public void IsSystemFolder_UserFolders_ReturnsFalse(string name)
        {
            Assert.False(_classifier.IsSystemFolder(name, defaultItemTypeIsMail: true));
        }

        [Fact]
        public void IsSystemFolder_NonMailDefaultItemType_AlwaysTrue()
        {
            Assert.True(_classifier.IsSystemFolder("Calendar", defaultItemTypeIsMail: false));
            Assert.True(_classifier.IsSystemFolder("Inbox",    defaultItemTypeIsMail: false));
        }

        [Fact]
        public void IsSystemFolder_NullOrEmpty_ReturnsTrue()
        {
            Assert.True(_classifier.IsSystemFolder(null, defaultItemTypeIsMail: true));
            Assert.True(_classifier.IsSystemFolder("",   defaultItemTypeIsMail: true));
        }

        [Fact]
        public void IsSystemFolder_NameMatchIsCaseInsensitive()
        {
            Assert.True(_classifier.IsSystemFolder("JUNK E-MAIL", defaultItemTypeIsMail: true));
            Assert.True(_classifier.IsSystemFolder("conversation history", defaultItemTypeIsMail: true));
        }
    }
}
```

- [ ] **Step 2: Run targeted tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~FolderClassifierTests"
```

Expected: build error — `IFolderClassifier` / `FolderClassifier` do not exist.

- [ ] **Step 3: Implement `IFolderClassifier` and `FolderClassifier`**

Create `VSTO2/OutlookAI/Services/Tools/IFolderClassifier.cs`:

```csharp
using System;
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Decides whether a folder is an Outlook system / noise folder that
    /// must never produce results for mailbox search. Centralised so both
    /// the AdvancedSearch projection and the iterative fallback honour the
    /// same skip rules.
    /// </summary>
    public interface IFolderClassifier
    {
        bool IsSystemFolder(string folderName, bool defaultItemTypeIsMail);
    }

    public sealed class FolderClassifier : IFolderClassifier
    {
        // Real Outlook folder display names. "Junk E-mail" uses a hyphen in
        // modern Outlook; "Junk Email" appears in older locales. Both skip.
        private static readonly HashSet<string> _systemNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Deleted Items",
                "Junk E-mail",
                "Junk Email",
                "Drafts",
                "Outbox",
                "Sync Issues",
                "Sync Issues (This computer only)",
                "Conflicts",
                "Local Failures",
                "Server Failures",
                "RSS Feeds",
                "RSS Subscriptions",
                "Conversation Action Settings",
                "Conversation History",
                "Quick Step Settings",
                "News Feed",
                "Feeds",
                "Files",
                "Detected Items",
                "Working Set",
                "Yammer Root",
            };

        public bool IsSystemFolder(string folderName, bool defaultItemTypeIsMail)
        {
            if (!defaultItemTypeIsMail) return true;
            if (string.IsNullOrWhiteSpace(folderName)) return true;
            return _systemNames.Contains(folderName.Trim());
        }
    }
}
```

- [ ] **Step 4: Build and re-run the targeted tests**

Build:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Then:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~FolderClassifierTests"
```

Expected: all `FolderClassifierTests` pass.

- [ ] **Step 5: Run full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes. No existing test regresses; the existing `SearchExecutionHelperTests.ShouldSkipAllMailFolder_*` cases stay intact until Task 6 wires `LiveOutlookSurface` to use the classifier.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/IFolderClassifier.cs VSTO2/OutlookAI.Tests/Services/Tools/FolderClassifierTests.cs
git commit -m "feat(search): add IFolderClassifier with full system-folder skip list" -m "Centralises the Outlook system / noise folder skip rules. Includes real folder display names that the old narrow ShouldSkipAllMailFolder helper missed (Junk E-mail with hyphen, Conversation Action Settings, Quick Step Settings, Working Set, Yammer Root, etc.) plus a non-mail DefaultItemType catch-all. Will replace ShouldSkipAllMailFolder in Task 6."
```

---

## Task 3: Add `MessageProjectionInput` DTO and `SearchResultProjector`

Pure logic for sort + classifier-filter + top-N + deferred snippet. Testable without Outlook.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/MessageProjectionInput.cs`
- Create: `VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/SearchResultProjectorTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/SearchResultProjectorTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SearchResultProjectorTests
    {
        private static MessageProjectionInput Item(
            string id, int year, string folder = "Inbox", Func<string> snippet = null)
        {
            return new MessageProjectionInput
            {
                Id = id,
                Subject = "s-" + id,
                From = "f-" + id,
                To = new[] { "to-" + id },
                ReceivedAt = new DateTimeOffset(year, 1, 1, 0, 0, 0, TimeSpan.Zero),
                HasAttachments = false,
                FolderName = folder,
                FolderDefaultItemTypeIsMail = true,
                SnippetFactory = snippet ?? (() => "snip-" + id),
            };
        }

        [Fact]
        public void Project_NewestSort_ReturnsByReceivedDesc()
        {
            var input = new[] { Item("a", 2010), Item("b", 2024), Item("c", 2017) };
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(new[] { "b", "c", "a" }, result.Select(r => r.Id).ToArray());
        }

        [Fact]
        public void Project_OldestSort_ReturnsByReceivedAsc()
        {
            var input = new[] { Item("a", 2010), Item("b", 2024), Item("c", 2017) };
            var args = new SearchMessagesArgs { SortOrder = "oldest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(new[] { "a", "c", "b" }, result.Select(r => r.Id).ToArray());
        }

        [Fact]
        public void Project_TopN_ClampsToMaxResults()
        {
            var input = Enumerable.Range(0, 20).Select(i => Item("i" + i, 2000 + i)).ToList();
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 3 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(3, result.Count);
        }

        [Fact]
        public void Project_SkipsSystemFolders()
        {
            var input = new[]
            {
                Item("good",  2024, folder: "Inbox"),
                Item("junk",  2024, folder: "Junk E-mail"),
                Item("trash", 2024, folder: "Deleted Items"),
            };
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(new[] { "good" }, result.Select(r => r.Id).ToArray());
        }

        [Fact]
        public void Project_DefersSnippet_OnlyForItemsInTopN()
        {
            int snippetCalls = 0;
            string Track(string id) => Item(id, 2000 + int.Parse(id.Substring(1)),
                snippet: () => { snippetCalls++; return "x"; }).SnippetFactory();

            var input = Enumerable.Range(0, 10).Select(i =>
            {
                int captured = i;
                return new MessageProjectionInput
                {
                    Id = "i" + captured,
                    Subject = "s",
                    From = "f",
                    To = new string[0],
                    ReceivedAt = new DateTimeOffset(2000 + captured, 1, 1, 0, 0, 0, TimeSpan.Zero),
                    HasAttachments = false,
                    FolderName = "Inbox",
                    FolderDefaultItemTypeIsMail = true,
                    SnippetFactory = () => { snippetCalls++; return "snip" + captured; },
                };
            }).ToList();

            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 3 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());

            Assert.Equal(3, result.Count);
            Assert.Equal(3, snippetCalls); // exactly the surviving top-N had snippet evaluated
        }

        [Fact]
        public void Project_SnippetFactoryThrowing_DoesNotBreakBatch()
        {
            var input = new[]
            {
                new MessageProjectionInput {
                    Id = "ok",  Subject = "s", From = "f", To = new string[0],
                    ReceivedAt = new DateTimeOffset(2024,1,1,0,0,0,TimeSpan.Zero),
                    FolderName = "Inbox", FolderDefaultItemTypeIsMail = true,
                    SnippetFactory = () => "good",
                },
                new MessageProjectionInput {
                    Id = "bad", Subject = "s", From = "f", To = new string[0],
                    ReceivedAt = new DateTimeOffset(2023,1,1,0,0,0,TimeSpan.Zero),
                    FolderName = "Inbox", FolderDefaultItemTypeIsMail = true,
                    SnippetFactory = () => throw new InvalidOperationException("boom"),
                },
            };
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(2, result.Count);
            Assert.Equal("good", result.First(r => r.Id == "ok").Snippet);
            Assert.Equal("",     result.First(r => r.Id == "bad").Snippet);
        }
    }
}
```

- [ ] **Step 2: Run targeted tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SearchResultProjectorTests"
```

Expected: build error — `MessageProjectionInput` and `SearchResultProjector` do not exist.

- [ ] **Step 3: Implement `MessageProjectionInput`**

Create `VSTO2/OutlookAI/Services/Tools/MessageProjectionInput.cs`:

```csharp
using System;
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// DTO carrying the fields SearchResultProjector needs to build a
    /// MessageSummary. SnippetFactory lets the projector defer expensive
    /// body access until after sort + classifier-filter + top-N have run,
    /// so we only pay snippet cost for items we actually return.
    /// </summary>
    public sealed class MessageProjectionInput
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public bool HasAttachments { get; set; }
        public string FolderName { get; set; }
        public bool FolderDefaultItemTypeIsMail { get; set; }
        public Func<string> SnippetFactory { get; set; }
    }
}
```

- [ ] **Step 4: Implement `SearchResultProjector`**

Create `VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Pure projection helper. Filters out system folders via
    /// IFolderClassifier, sorts by ReceivedAt per args.SortOrder, clamps to
    /// args.MaxResults, then evaluates each survivor's SnippetFactory.
    /// Resilient: a throwing SnippetFactory yields an empty snippet but
    /// does not poison the batch.
    /// </summary>
    public static class SearchResultProjector
    {
        public static IReadOnlyList<MessageSummary> Project(
            IEnumerable<MessageProjectionInput> items,
            SearchMessagesArgs args,
            IFolderClassifier classifier)
        {
            items     = items     ?? Array.Empty<MessageProjectionInput>();
            args      = args      ?? new SearchMessagesArgs();
            classifier = classifier ?? new FolderClassifier();

            var filtered = items.Where(i =>
                i != null
                && !classifier.IsSystemFolder(i.FolderName, i.FolderDefaultItemTypeIsMail));

            var ordered = (args.SortOrder ?? "newest").Equals("oldest", StringComparison.OrdinalIgnoreCase)
                ? filtered.OrderBy(i => i.ReceivedAt)
                : filtered.OrderByDescending(i => i.ReceivedAt);

            var maxResults = args.MaxResults > 0 ? args.MaxResults : 25;
            var top = ordered.Take(maxResults).ToList();

            var output = new List<MessageSummary>(top.Count);
            foreach (var i in top)
            {
                string snippet = "";
                try { snippet = i.SnippetFactory?.Invoke() ?? ""; }
                catch { snippet = ""; }

                output.Add(new MessageSummary
                {
                    Id = i.Id ?? "",
                    Subject = i.Subject ?? "",
                    From = i.From ?? "",
                    To = i.To ?? new string[0],
                    ReceivedAt = i.ReceivedAt,
                    Snippet = snippet,
                    HasAttachments = i.HasAttachments,
                });
            }
            return output;
        }
    }
}
```

- [ ] **Step 5: Build and re-run the targeted tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SearchResultProjectorTests"
```

Expected: all six `SearchResultProjectorTests` pass.

- [ ] **Step 6: Run full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/MessageProjectionInput.cs VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs VSTO2/OutlookAI.Tests/Services/Tools/SearchResultProjectorTests.cs
git commit -m "feat(search): add SearchResultProjector with deferred-snippet projection" -m "Pure helper that sorts by ReceivedAt, filters out system folders via IFolderClassifier, clamps to MaxResults, and evaluates each survivor's SnippetFactory exactly once. A throwing SnippetFactory yields an empty snippet but does not break the batch. Removes the per-item Body cost we previously paid for items that ended up dropped from the top-N."
```

---

## Task 4: Add `SearchScopeFormatter`

Pure helper that turns a list of folder paths into Outlook's comma-separated, single-quoted scope string. Folder enumeration itself stays in `LiveOutlookSurface` (it needs Outlook OOM and the marshaller); this helper just formats.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/SearchScopeFormatter.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/SearchScopeFormatterTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/SearchScopeFormatterTests.cs`:

```csharp
using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SearchScopeFormatterTests
    {
        [Fact]
        public void Format_SinglePath_QuotesIt()
        {
            var s = SearchScopeFormatter.Format(new[] { @"\\Mailbox - User\Inbox" });
            Assert.Equal(@"'\\Mailbox - User\Inbox'", s);
        }

        [Fact]
        public void Format_MultiplePaths_CommaSeparatedQuoted()
        {
            var s = SearchScopeFormatter.Format(new[]
            {
                @"\\Mailbox - User\Inbox",
                @"\\Archive PST\Sent Items",
            });
            Assert.Equal(@"'\\Mailbox - User\Inbox','\\Archive PST\Sent Items'", s);
        }

        [Fact]
        public void Format_PathContainingSingleQuote_IsDoubledForOutlook()
        {
            var s = SearchScopeFormatter.Format(new[] { @"\\Mailbox\O'Brien Folder" });
            // Outlook DASL scope escapes ' as ''
            Assert.Equal(@"'\\Mailbox\O''Brien Folder'", s);
        }

        [Fact]
        public void Format_NullOrEmpty_ReturnsEmptyString()
        {
            Assert.Equal("", SearchScopeFormatter.Format(null));
            Assert.Equal("", SearchScopeFormatter.Format(new string[0]));
        }

        [Fact]
        public void Format_SkipsNullAndWhitespacePaths()
        {
            var s = SearchScopeFormatter.Format(new[]
            {
                @"\\A\Inbox",
                null,
                "   ",
                @"\\B\Inbox",
            });
            Assert.Equal(@"'\\A\Inbox','\\B\Inbox'", s);
        }
    }
}
```

- [ ] **Step 2: Run targeted tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SearchScopeFormatterTests"
```

Expected: build error — `SearchScopeFormatter` does not exist.

- [ ] **Step 3: Implement `SearchScopeFormatter`**

Create `VSTO2/OutlookAI/Services/Tools/SearchScopeFormatter.cs`:

```csharp
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Formats a list of MAPI folder paths into the comma-separated,
    /// single-quoted scope string that <c>Application.AdvancedSearch</c>
    /// expects, escaping embedded single quotes per Outlook DASL rules.
    /// </summary>
    public static class SearchScopeFormatter
    {
        public static string Format(IEnumerable<string> folderPaths)
        {
            if (folderPaths == null) return "";
            var sb = new StringBuilder();
            bool first = true;
            foreach (var raw in folderPaths)
            {
                if (string.IsNullOrWhiteSpace(raw)) continue;
                if (!first) sb.Append(',');
                first = false;
                sb.Append('\'').Append(raw.Replace("'", "''")).Append('\'');
            }
            return sb.ToString();
        }
    }
}
```

- [ ] **Step 4: Build and re-run the targeted tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SearchScopeFormatterTests"
```

Expected: all `SearchScopeFormatterTests` pass.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/SearchScopeFormatter.cs VSTO2/OutlookAI.Tests/Services/Tools/SearchScopeFormatterTests.cs
git commit -m "feat(search): add SearchScopeFormatter for AdvancedSearch scope syntax" -m "Formats MAPI folder paths into the comma-separated single-quoted scope string Application.AdvancedSearch expects, with DASL single-quote doubling for paths that contain apostrophes."
```

---

## Task 5: Add `IAdvancedSearchHost` test seam and `FakeAdvancedSearchHost`

The runner needs a seam so tests can simulate AdvancedSearch invocation, cancellation, and the completion event without a live Outlook.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/IAdvancedSearchHost.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/FakeAdvancedSearchHost.cs`

- [ ] **Step 1: Define the interface and event args**

Create `VSTO2/OutlookAI/Services/Tools/IAdvancedSearchHost.cs`:

```csharp
using System;
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Abstraction over <c>Application.AdvancedSearch</c> + the
    /// <c>AdvancedSearchComplete</c> event so the runner can be unit-tested
    /// without a live Outlook. Production implementation
    /// (<c>LiveAdvancedSearchHost</c>) wraps the real COM API.
    /// </summary>
    public interface IAdvancedSearchHost
    {
        /// <summary>
        /// Begin an AdvancedSearch. Must return immediately. The
        /// implementation is responsible for raising <see cref="Completed"/>
        /// later with the matching <paramref name="tag"/>.
        /// </summary>
        void Start(string scope, string filter, bool searchSubFolders, string tag);

        /// <summary>
        /// Best-effort stop. Implementations should not throw on unknown
        /// tags.
        /// </summary>
        void Stop(string tag);

        event EventHandler<AdvancedSearchHostCompleteEventArgs> Completed;
    }

    public sealed class AdvancedSearchHostCompleteEventArgs : EventArgs
    {
        public string Tag { get; set; }
        public IReadOnlyList<MessageProjectionInput> Items { get; set; }
    }
}
```

- [ ] **Step 2: Build to confirm the project compiles**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: clean build.

- [ ] **Step 3: Create the test double `FakeAdvancedSearchHost`**

Create `VSTO2/OutlookAI.Tests/Services/Tools/FakeAdvancedSearchHost.cs`:

```csharp
using System;
using System.Collections.Generic;
using OutlookAI.Services.Tools;

namespace OutlookAI.Tests.Services.Tools
{
    /// <summary>
    /// Manual-control test double. Tests drive <see cref="Start"/>,
    /// <see cref="Stop"/>, and the <see cref="IAdvancedSearchHost.Completed"/>
    /// event timing directly.
    /// </summary>
    internal sealed class FakeAdvancedSearchHost : IAdvancedSearchHost
    {
        public List<(string Scope, string Filter, bool SearchSubFolders, string Tag)> StartCalls { get; }
            = new List<(string, string, bool, string)>();
        public List<string> StopCalls { get; } = new List<string>();

        /// <summary>If non-null, thrown from Start to simulate COM failure.</summary>
        public Func<string, Exception> ThrowOnStart { get; set; }

        public event EventHandler<AdvancedSearchHostCompleteEventArgs> Completed;

        public void Start(string scope, string filter, bool searchSubFolders, string tag)
        {
            var ex = ThrowOnStart?.Invoke(tag);
            if (ex != null) throw ex;
            StartCalls.Add((scope, filter, searchSubFolders, tag));
        }

        public void Stop(string tag) => StopCalls.Add(tag);

        public void RaiseCompleted(string tag, IReadOnlyList<MessageProjectionInput> items)
        {
            Completed?.Invoke(this, new AdvancedSearchHostCompleteEventArgs
            {
                Tag = tag,
                Items = items,
            });
        }
    }
}
```

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/IAdvancedSearchHost.cs VSTO2/OutlookAI.Tests/Services/Tools/FakeAdvancedSearchHost.cs
git commit -m "feat(search): add IAdvancedSearchHost test seam and fake host" -m "Interface abstracts Application.AdvancedSearch + AdvancedSearchComplete so the upcoming OutlookAdvancedSearchRunner can be unit-tested without a live Outlook. Fake host lets tests script Start/Stop sequences and raise the completion event with arbitrary projection inputs."
```

---

## Task 6: Implement `OutlookAdvancedSearchRunner`

Wraps the host with `Task<AdvancedSearchResult> RunAsync(...)`. Handles timeout, cancellation, tag-keyed event dispatch, and semaphore serialisation (one in-flight search at a time per process).

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/OutlookAdvancedSearchRunner.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookAdvancedSearchRunnerTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/OutlookAdvancedSearchRunnerTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookAdvancedSearchRunnerTests
    {
        private static MessageProjectionInput Item(string id) => new MessageProjectionInput
        {
            Id = id,
            Subject = id,
            From = "f",
            To = new string[0],
            ReceivedAt = DateTimeOffset.UtcNow,
            FolderName = "Inbox",
            FolderDefaultItemTypeIsMail = true,
            SnippetFactory = () => id,
        };

        [Fact]
        public async Task RunAsync_HappyPath_ReturnsCompletedWithItems()
        {
            var host = new FakeAdvancedSearchHost();
            using var runner = new OutlookAdvancedSearchRunner(host);

            var task = runner.RunAsync("'\\\\store\\Inbox'", "filter", true, TimeSpan.FromSeconds(5), CancellationToken.None);

            // Wait for Start to be observed, then raise Completed for that tag.
            var tag = await WaitForTag(host);
            host.RaiseCompleted(tag, new[] { Item("a"), Item("b") });

            var result = await task;
            Assert.True(result.Completed);
            Assert.False(result.TimedOut);
            Assert.False(result.Cancelled);
            Assert.Null(result.Error);
            Assert.Equal(2, result.Items.Count);
            Assert.Equal(new[] { "a", "b" }, result.Items.Select(i => i.Id).ToArray());
        }

        [Fact]
        public async Task RunAsync_NoCompletion_TimesOutAndStops()
        {
            var host = new FakeAdvancedSearchHost();
            using var runner = new OutlookAdvancedSearchRunner(host);

            var task = runner.RunAsync("scope", "filter", true, TimeSpan.FromMilliseconds(50), CancellationToken.None);
            var result = await task;

            Assert.False(result.Completed);
            Assert.True(result.TimedOut);
            Assert.Single(host.StartCalls);
            Assert.Single(host.StopCalls);
            Assert.Equal(host.StartCalls[0].Tag, host.StopCalls[0]);
        }

        [Fact]
        public async Task RunAsync_CancellationToken_Cancelled_StopsAndReturnsCancelled()
        {
            var host = new FakeAdvancedSearchHost();
            using var runner = new OutlookAdvancedSearchRunner(host);
            using var cts = new CancellationTokenSource();

            var task = runner.RunAsync("scope", "filter", true, TimeSpan.FromSeconds(30), cts.Token);
            await WaitForTag(host);
            cts.Cancel();

            var result = await task;
            Assert.False(result.Completed);
            Assert.False(result.TimedOut);
            Assert.True(result.Cancelled);
            Assert.Single(host.StopCalls);
        }

        [Fact]
        public async Task RunAsync_StartThrows_ReturnsError_NoSubscriptionLeak()
        {
            var host = new FakeAdvancedSearchHost
            {
                ThrowOnStart = _ => new COMException("simulated")
            };
            using var runner = new OutlookAdvancedSearchRunner(host);

            var result = await runner.RunAsync("scope", "filter", true, TimeSpan.FromSeconds(5), CancellationToken.None);

            Assert.False(result.Completed);
            Assert.False(result.TimedOut);
            Assert.False(result.Cancelled);
            Assert.NotNull(result.Error);
            Assert.IsType<COMException>(result.Error);
        }

        [Fact]
        public async Task RunAsync_UnknownTagCompletedEvent_IsIgnored()
        {
            var host = new FakeAdvancedSearchHost();
            using var runner = new OutlookAdvancedSearchRunner(host);

            var task = runner.RunAsync("scope", "filter", true, TimeSpan.FromSeconds(5), CancellationToken.None);
            await WaitForTag(host);

            // Fire an event with a tag the runner never issued.
            host.RaiseCompleted("totally-unrelated-tag", new[] { Item("x") });

            // The pending task must still be pending: complete it with the right tag now.
            host.RaiseCompleted(host.StartCalls[0].Tag, new[] { Item("real") });
            var result = await task;

            Assert.True(result.Completed);
            Assert.Equal("real", result.Items[0].Id);
        }

        [Fact]
        public async Task RunAsync_TwoCalls_AreSerialised()
        {
            var host = new FakeAdvancedSearchHost();
            using var runner = new OutlookAdvancedSearchRunner(host);

            var first  = runner.RunAsync("a", null, true, TimeSpan.FromSeconds(5), CancellationToken.None);
            var second = runner.RunAsync("b", null, true, TimeSpan.FromSeconds(5), CancellationToken.None);

            // Allow scheduler to run.
            await Task.Delay(50);

            // Only the first Start should have been observed; second waits on the semaphore.
            Assert.Single(host.StartCalls);

            // Complete first, then second should proceed.
            host.RaiseCompleted(host.StartCalls[0].Tag, Array.Empty<MessageProjectionInput>());
            await first;

            await WaitForCondition(() => host.StartCalls.Count == 2, TimeSpan.FromSeconds(2));
            host.RaiseCompleted(host.StartCalls[1].Tag, Array.Empty<MessageProjectionInput>());
            var secondResult = await second;
            Assert.True(secondResult.Completed);
        }

        private static async Task<string> WaitForTag(FakeAdvancedSearchHost host)
        {
            await WaitForCondition(() => host.StartCalls.Count > 0, TimeSpan.FromSeconds(2));
            return host.StartCalls[host.StartCalls.Count - 1].Tag;
        }

        private static async Task WaitForCondition(Func<bool> cond, TimeSpan timeout)
        {
            var start = DateTime.UtcNow;
            while (!cond() && DateTime.UtcNow - start < timeout)
            {
                await Task.Delay(10);
            }
            if (!cond()) throw new TimeoutException("condition not met");
        }
    }
}
```

- [ ] **Step 2: Run the targeted tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookAdvancedSearchRunnerTests"
```

Expected: build error — `OutlookAdvancedSearchRunner` and `AdvancedSearchResult` do not exist.

- [ ] **Step 3: Implement `OutlookAdvancedSearchRunner` and `AdvancedSearchResult`**

Create `VSTO2/OutlookAI/Services/Tools/OutlookAdvancedSearchRunner.cs`:

```csharp
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Result of one AdvancedSearch invocation.
    /// </summary>
    public sealed class AdvancedSearchResult
    {
        public bool Completed { get; set; }
        public bool TimedOut { get; set; }
        public bool Cancelled { get; set; }
        public Exception Error { get; set; }
        public IReadOnlyList<MessageProjectionInput> Items { get; set; }
    }

    public interface IOutlookAdvancedSearchRunner : IDisposable
    {
        Task<AdvancedSearchResult> RunAsync(
            string scope,
            string filter,
            bool searchSubFolders,
            TimeSpan timeout,
            CancellationToken ct);
    }

    /// <summary>
    /// Drives an <see cref="IAdvancedSearchHost"/> with timeout, cancellation,
    /// tag-keyed dispatch, and process-wide serialisation (one in-flight
    /// AdvancedSearch at a time, which is the documented safe ceiling for
    /// the Outlook OOM).
    /// </summary>
    public sealed class OutlookAdvancedSearchRunner : IOutlookAdvancedSearchRunner
    {
        private readonly IAdvancedSearchHost _host;
        private readonly SemaphoreSlim _serialiser = new SemaphoreSlim(1, 1);
        private readonly ConcurrentDictionary<string, TaskCompletionSource<AdvancedSearchResult>> _pending
            = new ConcurrentDictionary<string, TaskCompletionSource<AdvancedSearchResult>>();
        private bool _disposed;

        public OutlookAdvancedSearchRunner(IAdvancedSearchHost host)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _host.Completed += OnHostCompleted;
        }

        public async Task<AdvancedSearchResult> RunAsync(
            string scope,
            string filter,
            bool searchSubFolders,
            TimeSpan timeout,
            CancellationToken ct)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(OutlookAdvancedSearchRunner));

            await _serialiser.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                var tag = Guid.NewGuid().ToString("N");
                var tcs = new TaskCompletionSource<AdvancedSearchResult>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
                _pending[tag] = tcs;

                try { _host.Start(scope, filter, searchSubFolders, tag); }
                catch (Exception ex)
                {
                    _pending.TryRemove(tag, out _);
                    TraceLog.Write("AdvancedSearch Start threw: " + ex.Message, "Runner");
                    return new AdvancedSearchResult { Error = ex };
                }

                using (var timeoutCts = new CancellationTokenSource(timeout))
                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token))
                using (linked.Token.Register(() =>
                {
                    if (!_pending.TryRemove(tag, out var pendingTcs)) return;
                    try { _host.Stop(tag); } catch { /* best-effort */ }
                    if (ct.IsCancellationRequested)
                        pendingTcs.TrySetResult(new AdvancedSearchResult { Cancelled = true });
                    else
                        pendingTcs.TrySetResult(new AdvancedSearchResult { TimedOut = true });
                }))
                {
                    return await tcs.Task.ConfigureAwait(false);
                }
            }
            finally
            {
                _serialiser.Release();
            }
        }

        private void OnHostCompleted(object sender, AdvancedSearchHostCompleteEventArgs e)
        {
            if (e?.Tag == null) return;
            if (_pending.TryRemove(e.Tag, out var tcs))
            {
                tcs.TrySetResult(new AdvancedSearchResult
                {
                    Completed = true,
                    Items = e.Items ?? Array.Empty<MessageProjectionInput>(),
                });
            }
            // Unknown tag: silently ignore. Some other AdvancedSearch (or a
            // stale event) is not our problem.
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _host.Completed -= OnHostCompleted;
            _serialiser.Dispose();
            // Any still-pending TCSs are abandoned; their owners will see
            // either timeout or cancellation depending on token state.
        }
    }
}
```

- [ ] **Step 4: Build and run targeted tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookAdvancedSearchRunnerTests"
```

Expected: all six runner tests pass.

- [ ] **Step 5: Run full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookAdvancedSearchRunner.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookAdvancedSearchRunnerTests.cs
git commit -m "feat(search): add OutlookAdvancedSearchRunner with timeout and cancel" -m "Async Task<AdvancedSearchResult> wrapper around IAdvancedSearchHost with per-call GUID tag dispatch, hard timeout, cooperative cancellation, and a process-wide semaphore that serialises in-flight AdvancedSearch invocations (the documented safe ceiling for the Outlook OOM). Stop is best-effort on timeout / cancel; unknown-tag completion events are silently ignored."
```

---

## Task 7: Implement `LiveAdvancedSearchHost` (production COM wrapper)

Thin, untestable-without-Outlook wrapper that bridges `Application.AdvancedSearch` + `AdvancedSearchComplete` to `IAdvancedSearchHost`. Verified by Task 12 smoke. Marshals every COM access onto the Outlook UI thread via the existing `OutlookThreadMarshaller`.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/LiveAdvancedSearchHost.cs`

- [ ] **Step 1: Implement `LiveAdvancedSearchHost`**

Create `VSTO2/OutlookAI/Services/Tools/LiveAdvancedSearchHost.cs`:

```csharp
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OutlookAI.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Production <see cref="IAdvancedSearchHost"/>. Wraps
    /// <c>Application.AdvancedSearch</c> and the
    /// <c>AdvancedSearchComplete</c> event, projecting each completed
    /// Outlook <c>Search.Results</c> into <see cref="MessageProjectionInput"/>
    /// items on the Outlook UI thread (where COM access is legal).
    /// </summary>
    public sealed class LiveAdvancedSearchHost : IAdvancedSearchHost, IDisposable
    {
        private readonly Outlook.Application _application;
        private readonly OutlookThreadMarshaller _marshaller;
        private readonly IdResolver _ids;
        private readonly ConcurrentDictionary<string, Outlook.Search> _tagToSearch
            = new ConcurrentDictionary<string, Outlook.Search>();
        private bool _subscribed;
        // Snippet character cap matches LiveOutlookSurface.SnippetChars.
        private const int SnippetChars = 160;
        // Body byte cap mirrors LiveOutlookSurface.MaxBodyChars.
        private const int MaxBodyChars = 32 * 1024;

        public event EventHandler<AdvancedSearchHostCompleteEventArgs> Completed;

        public LiveAdvancedSearchHost(
            Outlook.Application application,
            OutlookThreadMarshaller marshaller,
            IdResolver ids)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _marshaller = marshaller ?? throw new ArgumentNullException(nameof(marshaller));
            _ids = ids ?? throw new ArgumentNullException(nameof(ids));
            Subscribe();
        }

        private void Subscribe()
        {
            _marshaller.RunAsync(() =>
            {
                if (_subscribed) return;
                _application.AdvancedSearchComplete += OnAdvancedSearchComplete;
                _subscribed = true;
            }, System.Threading.CancellationToken.None)
            .GetAwaiter().GetResult();
        }

        public void Start(string scope, string filter, bool searchSubFolders, string tag)
        {
            _marshaller.RunAsync(() =>
            {
                TraceLog.Write("AdvancedSearch Start tag=" + tag + " scope_len=" + (scope?.Length ?? 0)
                    + " filter=" + (string.IsNullOrEmpty(filter) ? "<none>" : filter)
                    + " sub=" + searchSubFolders, "LiveHost");
                var search = _application.AdvancedSearch(scope, filter ?? "", searchSubFolders, tag);
                if (search != null) _tagToSearch[tag] = search;
            }, System.Threading.CancellationToken.None)
            .GetAwaiter().GetResult();
        }

        public void Stop(string tag)
        {
            _marshaller.RunAsync(() =>
            {
                if (_tagToSearch.TryRemove(tag, out var search))
                {
                    try { search.Stop(); }
                    catch (COMException ex)
                    {
                        TraceLog.Write("AdvancedSearch Stop COMException tag=" + tag + " " + ex.Message, "LiveHost");
                    }
                }
            }, System.Threading.CancellationToken.None)
            .GetAwaiter().GetResult();
        }

        private void OnAdvancedSearchComplete(Outlook.Search search)
        {
            if (search == null) return;
            var tag = search.Tag ?? "";
            _tagToSearch.TryRemove(tag, out _);

            var items = new List<MessageProjectionInput>();
            try
            {
                foreach (var obj in search.Results)
                {
                    try
                    {
                        var mi = obj as Outlook.MailItem;
                        if (mi == null) continue;
                        items.Add(BuildProjectionInput(mi));
                    }
                    catch (COMException ex)
                    {
                        TraceLog.Write("AdvancedSearch result item COMException: " + ex.Message, "LiveHost");
                    }
                }
            }
            catch (COMException ex)
            {
                TraceLog.Write("AdvancedSearch results enumeration COMException: " + ex.Message, "LiveHost");
            }

            TraceLog.Write("AdvancedSearch Complete tag=" + tag + " raw_count=" + items.Count, "LiveHost");
            try { Completed?.Invoke(this, new AdvancedSearchHostCompleteEventArgs { Tag = tag, Items = items }); }
            catch (Exception ex)
            {
                TraceLog.Write("Completed handler threw: " + ex.Message, "LiveHost");
            }
        }

        private MessageProjectionInput BuildProjectionInput(Outlook.MailItem mi)
        {
            string folderName = "";
            bool folderIsMail = true;
            try
            {
                var parent = mi.Parent as Outlook.MAPIFolder;
                if (parent != null)
                {
                    folderName = parent.Name ?? "";
                    try { folderIsMail = parent.DefaultItemType == Outlook.OlItemType.olMailItem; }
                    catch (COMException) { folderIsMail = true; }
                }
            }
            catch (COMException) { /* leave defaults */ }

            // Capture id eagerly (cheap); defer Body to SnippetFactory.
            string id = "";
            try { id = _ids.Shorten(mi.EntryID); } catch (COMException) { }

            string subject = "";
            try { subject = mi.Subject ?? ""; } catch (COMException) { }

            string from = "";
            try { from = mi.SenderName ?? mi.SenderEmailAddress ?? ""; } catch (COMException) { }

            string to = "";
            try { to = mi.To ?? ""; } catch (COMException) { }

            DateTimeOffset receivedAt = DateTimeOffset.MinValue;
            try { receivedAt = ToOffset(mi.ReceivedTime); } catch (COMException) { }

            bool hasAttachments = false;
            try { hasAttachments = (mi.Attachments?.Count ?? 0) > 0; } catch (COMException) { }

            // SnippetFactory closes over the MailItem and is evaluated on the
            // UI thread (caller invokes from inside Project which is called
            // from inside a marshalled Run block in LiveOutlookSurface).
            Func<string> snippetFactory = () =>
            {
                try
                {
                    var body = mi.Body ?? "";
                    if (body.Length > MaxBodyChars) body = body.Substring(0, MaxBodyChars);
                    if (body.Length <= SnippetChars) return body.Trim();
                    return body.Substring(0, SnippetChars).Trim();
                }
                catch (COMException) { return ""; }
            };

            return new MessageProjectionInput
            {
                Id = id,
                Subject = subject,
                From = from,
                To = SplitAddresses(to),
                ReceivedAt = receivedAt,
                HasAttachments = hasAttachments,
                FolderName = folderName,
                FolderDefaultItemTypeIsMail = folderIsMail,
                SnippetFactory = snippetFactory,
            };
        }

        private static IReadOnlyList<string> SplitAddresses(string addresses)
        {
            if (string.IsNullOrWhiteSpace(addresses)) return new string[0];
            var parts = addresses.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            var list = new List<string>(parts.Length);
            foreach (var p in parts)
            {
                var trimmed = p.Trim();
                if (trimmed.Length > 0) list.Add(trimmed);
            }
            return list;
        }

        private static DateTimeOffset ToOffset(DateTime dt)
        {
            if (dt.Kind == DateTimeKind.Unspecified)
                return new DateTimeOffset(dt, TimeZoneInfo.Local.GetUtcOffset(dt));
            return new DateTimeOffset(dt);
        }

        public void Dispose()
        {
            _marshaller.RunAsync(() =>
            {
                if (!_subscribed) return;
                try { _application.AdvancedSearchComplete -= OnAdvancedSearchComplete; } catch { }
                _subscribed = false;
            }, System.Threading.CancellationToken.None)
            .GetAwaiter().GetResult();
        }
    }
}
```

- [ ] **Step 2: Build to confirm it compiles cleanly**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds.

- [ ] **Step 3: Run full suite (no new tests; existing must stay green)**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes. `LiveAdvancedSearchHost` has no unit test by design — it is the COM bridge verified by smoke in Task 12.

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/LiveAdvancedSearchHost.cs
git commit -m "feat(search): add LiveAdvancedSearchHost COM wrapper" -m "Production IAdvancedSearchHost. Wraps Application.AdvancedSearch and AdvancedSearchComplete, marshals every COM call through OutlookThreadMarshaller, and projects raw Outlook MailItems to MessageProjectionInput with a deferred Body-snippet factory. Per-item COMExceptions are caught and traced; the batch keeps going."
```

---

## Task 8: Rewire `LiveOutlookSurface.SearchMessages` to the runner with yielding fallback

The orchestration change. After this task, `SearchMessages` is a thin orchestrator. The blocking-on-UI-thread iterative loop is gone; the new fallback yields the UI thread between folders by using one `_marshaller.RunAsync` per folder.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/SearchExecutionHelper.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/SearchExecutionHelperTests.cs`

- [ ] **Step 1: Update `SearchExecutionHelperTests` to retire helpers being replaced**

`ShouldSkipAllMailFolder`, `SortDescending`, and `MergeAndSortSearchResults` are superseded by `FolderClassifier` and `SearchResultProjector`. The cleanest path is to delete these tests entirely now (they'll fail anyway when the helpers go away in Step 3). Open `VSTO2/OutlookAI.Tests/Services/Tools/SearchExecutionHelperTests.cs` and delete the three test methods (and any `using`s that become unused). If the test class becomes empty, delete the file and remove it from the project (csproj will pick up the delete automatically with SDK-style projects; for legacy csproj, ensure it is not listed under `<Compile Include=...>`).

- [ ] **Step 2: Run the targeted tests to verify the deletions compile**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build clean. The helper file may still expose `ShouldSkipAllMailFolder` / `SortDescending` / `MergeAndSortSearchResults`; nothing else references them yet, but they will be removed in Step 3 below.

- [ ] **Step 3: Remove the obsolete helpers from `SearchExecutionHelper.cs`**

Open `VSTO2/OutlookAI/Services/Tools/SearchExecutionHelper.cs`. Delete:
- `ShouldSkipAllMailFolder(string name)` — replaced by `IFolderClassifier`.
- `SortDescending(SearchMessagesArgs args)` — no longer used; projector sorts globally.
- `MergeAndSortSearchResults(...)` — replaced by `SearchResultProjector.Project`.

If the file ends up empty, delete it. Update the project references the same way as Step 1.

- [ ] **Step 4: Add runner + classifier + timeout to `LiveOutlookSurface` constructor**

Open `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`. Add fields:

```csharp
private readonly IOutlookAdvancedSearchRunner _runner;
private readonly IFolderClassifier _classifier;
private static readonly TimeSpan _searchTimeout = TimeSpan.FromSeconds(30);
```

Expand the constructor to accept the runner and (optional) classifier. Existing call sites in `ThisAddIn` will be fixed in Task 10; for now, add an overload that keeps old callers compiling:

```csharp
public LiveOutlookSurface(
    Outlook.Application application,
    OutlookThreadMarshaller marshaller,
    IdResolver ids,
    Outlook.Inspector composeInspector,
    Outlook.Explorer explorer,
    IOutlookAdvancedSearchRunner runner,
    IFolderClassifier classifier = null)
{
    _application = application ?? throw new ArgumentNullException(nameof(application));
    _marshaller = marshaller ?? throw new ArgumentNullException(nameof(marshaller));
    _ids = ids ?? throw new ArgumentNullException(nameof(ids));
    _composeInspector = composeInspector;
    _explorer = explorer;
    _runner = runner ?? throw new ArgumentNullException(nameof(runner));
    _classifier = classifier ?? new FolderClassifier();
}
```

If the original constructor (without `runner`) is still referenced anywhere outside `ThisAddIn`, leave it temporarily and have it throw `InvalidOperationException("AdvancedSearchRunner is required")`. ThisAddIn is rewired in Task 10.

- [ ] **Step 5: Rewrite `SearchMessages` body**

Replace the entire `SearchMessages` method (currently the `Run(() => { ... })` block built around per-folder enumeration and `Restrict`) with this orchestrator:

```csharp
public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default)
{
    args = args ?? new SearchMessagesArgs();
    var filter = BuildRestrictFilter(args);
    var scopeMode = (args.Scope ?? "auto").Trim().ToLowerInvariant();

    // Compute scope on UI thread.
    SearchScope scope;
    try
    {
        scope = _marshaller.RunAsync(() => BuildSearchScope(args, scopeMode), ct).GetAwaiter().GetResult();
    }
    catch (OperationCanceledException) { throw; }

    // Primary: AdvancedSearch.
    var primary = _runner.RunAsync(
        scope.ScopeString,
        filter,
        scope.SearchSubFolders,
        _searchTimeout,
        ct).GetAwaiter().GetResult();

    if (primary.Cancelled) throw new OperationCanceledException(ct);

    if (primary.Completed && primary.Items != null)
    {
        TraceLog.Write("SearchMessages primary=AdvancedSearch complete raw_count=" + primary.Items.Count, "LiveOutlookSurface");
        return _marshaller.RunAsync(
            () => SearchResultProjector.Project(primary.Items, args, _classifier),
            ct).GetAwaiter().GetResult();
    }

    var reason = primary.TimedOut ? "timeout" : primary.Error != null ? "com" : "null_results";
    TraceLog.Write("SearchMessages fallback reason=" + reason, "LiveOutlookSurface");
    return FallbackIterativeSearch(args, filter, scopeMode, ct);
}
```

- [ ] **Step 6: Add `BuildSearchScope` helper**

Add this private method to `LiveOutlookSurface`. It runs on the UI thread (called from inside a `RunAsync`):

```csharp
private SearchScope BuildSearchScope(SearchMessagesArgs args, string scopeMode)
{
    var paths = new List<string>();
    bool searchSubFolders = true;

    try
    {
        if (!string.IsNullOrEmpty(args.FolderId))
        {
            var folder = ResolveFolder(args.FolderId);
            if (folder != null) paths.Add(folder.FolderPath);
        }
        else if (scopeMode == "current_folder")
        {
            var folder = ResolveCurrentFolder();
            if (folder != null) paths.Add(folder.FolderPath);
            searchSubFolders = false;
        }
        else // "all_mail" or "auto"
        {
            foreach (Outlook.Store store in _application.Session.Stores)
            {
                try
                {
                    var root = store.GetRootFolder();
                    if (root != null) paths.Add(root.FolderPath);
                }
                catch (COMException) { }
            }
        }
    }
    catch (COMException ex)
    {
        TraceLog.Write("BuildSearchScope COMException: " + ex.Message, "LiveOutlookSurface");
    }

    return new SearchScope
    {
        ScopeString = SearchScopeFormatter.Format(paths),
        SearchSubFolders = searchSubFolders,
        ResolvedFolderPaths = paths,
    };
}
```

Add the small `SearchScope` DTO at the bottom of the namespace (in any file in `Services/Tools` — for tidiness, put it in `SearchScopeFormatter.cs`):

```csharp
public sealed class SearchScope
{
    public string ScopeString { get; set; }
    public bool SearchSubFolders { get; set; }
    public IReadOnlyList<string> ResolvedFolderPaths { get; set; }
}
```

- [ ] **Step 7: Rewrite the iterative fallback to yield between folders**

Replace `SearchOneFolder` and any per-folder enumeration with this new pair. The old method collected `MessageSummary` directly; the new pair collects `MessageProjectionInput` (no `Body` access) and the projector evaluates snippets only for survivors.

```csharp
private IReadOnlyList<MessageSummary> FallbackIterativeSearch(
    SearchMessagesArgs args, string filter, string scopeMode, CancellationToken ct)
{
    var allInputs = new List<MessageProjectionInput>();
    var searchAllMail = scopeMode == "all_mail";

    var folders = _marshaller.RunAsync(
        () => ResolveSearchFolders(args, allMail: searchAllMail || scopeMode == "auto"),
        ct).GetAwaiter().GetResult();

    var searched = 0;
    foreach (var folder in folders)
    {
        ct.ThrowIfCancellationRequested();
        var folderInputs = _marshaller.RunAsync(
            () => CollectFolderInputs(folder, args, filter),
            ct).GetAwaiter().GetResult();
        allInputs.AddRange(folderInputs);
        searched++;
        TraceLog.Write("SearchMessages fallback folder_done=" + SafeFolderName(folder)
            + " taken=" + folderInputs.Count + " searched=" + searched, "LiveOutlookSurface");
    }

    return _marshaller.RunAsync(
        () => SearchResultProjector.Project(allInputs, args, _classifier),
        ct).GetAwaiter().GetResult();
}

private List<MessageProjectionInput> CollectFolderInputs(
    Outlook.MAPIFolder folder, SearchMessagesArgs args, string filter)
{
    var inputs = new List<MessageProjectionInput>();
    if (folder == null) return inputs;

    string folderName = "";
    bool folderIsMail = true;
    try { folderName = folder.Name ?? ""; } catch (COMException) { }
    try { folderIsMail = folder.DefaultItemType == Outlook.OlItemType.olMailItem; } catch (COMException) { }
    if (_classifier.IsSystemFolder(folderName, folderIsMail)) return inputs;

    Outlook.Items items;
    try
    {
        items = string.IsNullOrEmpty(filter) ? folder.Items : folder.Items.Restrict(filter);
    }
    catch (COMException ex)
    {
        TraceLog.Write("CollectFolderInputs Restrict COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface");
        return inputs;
    }

    foreach (var obj in items)
    {
        try
        {
            var mi = obj as Outlook.MailItem;
            if (mi == null) continue;
            var input = BuildFallbackInput(mi, folderName, folderIsMail);
            if (input != null) inputs.Add(input);
        }
        catch (COMException) { /* skip the item */ }
    }
    return inputs;
}

private MessageProjectionInput BuildFallbackInput(
    Outlook.MailItem mi, string folderName, bool folderIsMail)
{
    string id = ""; try { id = _ids.Shorten(mi.EntryID); } catch (COMException) { }
    string subject = ""; try { subject = mi.Subject ?? ""; } catch (COMException) { }
    string from = ""; try { from = mi.SenderName ?? mi.SenderEmailAddress ?? ""; } catch (COMException) { }
    string to = ""; try { to = mi.To ?? ""; } catch (COMException) { }
    DateTimeOffset receivedAt = DateTimeOffset.MinValue;
    try { receivedAt = ToOffset(mi.ReceivedTime); } catch (COMException) { }
    bool hasAttachments = false;
    try { hasAttachments = (mi.Attachments?.Count ?? 0) > 0; } catch (COMException) { }

    Func<string> snippetFactory = () =>
    {
        try { return SnippetOf(mi.Body); } catch (COMException) { return ""; }
    };

    return new MessageProjectionInput
    {
        Id = id,
        Subject = subject,
        From = from,
        To = SplitAddresses(to),
        ReceivedAt = receivedAt,
        HasAttachments = hasAttachments,
        FolderName = folderName,
        FolderDefaultItemTypeIsMail = folderIsMail,
        SnippetFactory = snippetFactory,
    };
}
```

`SplitAddresses`, `SnippetOf`, `ToOffset`, `SafeFolderName`, `ResolveCurrentFolder`, `ResolveFolder`, and `ResolveSearchFolders` already exist in `LiveOutlookSurface` from earlier phases; keep them. Delete the old `SearchOneFolder`, `TraceSearch`, `TraceSearchProgress`, and the now-unused per-folder sort helper if present.

- [ ] **Step 8: Build and run full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: clean build.

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes. The deleted `SearchExecutionHelperTests` cases are gone; everything else is green.

- [ ] **Step 9: Commit**

```powershell
git add -A
git commit -m "feat(search): orchestrate SearchMessages through AdvancedSearch with yielding fallback" -m "LiveOutlookSurface.SearchMessages is now a thin orchestrator: build scope on UI thread, hand off to OutlookAdvancedSearchRunner with a 30s timeout, project results via SearchResultProjector. On timeout / COMException / null-results the fallback walks folders one per OutlookThreadMarshaller.RunAsync call so the UI thread is released between folders. Deletes the old per-folder synchronous loop, ShouldSkipAllMailFolder, SortDescending, and MergeAndSortSearchResults; their roles are taken by FolderClassifier and SearchResultProjector."
```

---

## Task 9: Rewire `LiveOutlookSurface.CountMessages` the same way

`CountMessages` shares the freeze problem. Same engine swap, simpler return shape.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`

- [ ] **Step 1: Rewrite `CountMessages`**

Replace the existing `CountMessages` method with:

```csharp
public int CountMessages(SearchMessagesArgs args, CancellationToken ct = default)
{
    args = args ?? new SearchMessagesArgs();
    // Counting honours the same skip list / cancellation as SearchMessages.
    // We reuse SearchMessages with a very large MaxResults and count the
    // projection. AdvancedSearch returns the matching set; the projector
    // applies the classifier; the size is the count.
    var widened = new SearchMessagesArgs
    {
        Query = args.Query,
        From = args.From,
        SubjectContains = args.SubjectContains,
        BodyContains = args.BodyContains,
        HasAttachment = args.HasAttachment,
        IsUnread = args.IsUnread,
        IsFlagged = args.IsFlagged,
        Importance = args.Importance,
        FolderId = args.FolderId,
        DateFrom = args.DateFrom,
        DateTo = args.DateTo,
        Scope = args.Scope,
        SortOrder = args.SortOrder,
        AttachmentFilter = args.AttachmentFilter,
        ReadStatus = args.ReadStatus,
        FlagStatus = args.FlagStatus,
        ImportanceFilter = args.ImportanceFilter,
        MaxResults = int.MaxValue,
    };
    var results = SearchMessages(widened, ct);
    return results?.Count ?? 0;
}
```

This intentionally reuses the now-non-blocking `SearchMessages` path so counts get the same engine, skip list, and cancellation behavior for free. We pay one snippet evaluation per item in the result set, which is wasted for a pure count; if real-world counts get prohibitively expensive, a follow-up can add a snippet-skipping projector overload. Not in this phase.

- [ ] **Step 2: Build and run full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: clean build, all tests pass.

- [ ] **Step 3: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs
git commit -m "feat(search): route CountMessages through the AdvancedSearch orchestrator" -m "CountMessages delegates to SearchMessages with MaxResults=int.MaxValue so counts honour the new non-blocking engine, the skip list, and cancellation. The redundant snippet evaluation per item is acceptable for the first ship; a snippet-skipping projector overload is a follow-up if real counts demand it."
```

---

## Task 10: Wire `ThisAddIn` to own a singleton host + runner

Both task panes (`InboxCopilotPane.Bind`, `AITaskPane.Bind`) construct their own `LiveOutlookSurface`. They must share one `IOutlookAdvancedSearchRunner` because the runner subscribes to `Application.AdvancedSearchComplete` exactly once.

**Files:**
- Modify: `VSTO2/OutlookAI/ThisAddIn.cs`
- Modify: `VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPane.cs`
- Modify: `VSTO2/OutlookAI/TaskPane/AITaskPane.cs`

- [ ] **Step 1: Locate the existing service-property pattern in `ThisAddIn.cs`**

Open `VSTO2/OutlookAI/ThisAddIn.cs`. Find the existing properties (`OutlookMarshaller`, `IdResolver`). They are set in `ThisAddIn_Startup` and used by the panes via `Globals.ThisAddIn?.OutlookMarshaller`. Mirror that pattern.

- [ ] **Step 2: Add fields and properties for the host + runner**

Add near the other service fields in `ThisAddIn`:

```csharp
private OutlookAI.Services.Tools.LiveAdvancedSearchHost _advancedSearchHost;
private OutlookAI.Services.Tools.OutlookAdvancedSearchRunner _advancedSearchRunner;
private OutlookAI.Services.Tools.IFolderClassifier _folderClassifier;

public OutlookAI.Services.Tools.IOutlookAdvancedSearchRunner AdvancedSearchRunner
    => _advancedSearchRunner;

public OutlookAI.Services.Tools.IFolderClassifier FolderClassifier
    => _folderClassifier;
```

- [ ] **Step 3: Construct them inside `ThisAddIn_Startup`**

Inside `ThisAddIn_Startup`, *after* `OutlookMarshaller` and `IdResolver` are created (since the host needs both), add:

```csharp
_folderClassifier = new OutlookAI.Services.Tools.FolderClassifier();
_advancedSearchHost = new OutlookAI.Services.Tools.LiveAdvancedSearchHost(
    Application, OutlookMarshaller, IdResolver);
_advancedSearchRunner = new OutlookAI.Services.Tools.OutlookAdvancedSearchRunner(
    _advancedSearchHost);
TraceLog.Write("AdvancedSearch services constructed", "ThisAddIn");
```

- [ ] **Step 4: Dispose them inside `ThisAddIn_Shutdown`**

Inside `ThisAddIn_Shutdown` (or whatever the existing shutdown handler is named), add:

```csharp
try { _advancedSearchRunner?.Dispose(); }
catch (Exception ex) { TraceLog.Write("Runner dispose: " + ex.Message, "ThisAddIn"); }
try { _advancedSearchHost?.Dispose(); }
catch (Exception ex) { TraceLog.Write("Host dispose: " + ex.Message, "ThisAddIn"); }
```

Order matters: runner before host (the runner unsubscribes its handler on dispose; the host then drops its own COM subscription).

- [ ] **Step 5: Update `InboxCopilotPane.Bind` to pass the runner**

Open `VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPane.cs`. Find the `LiveOutlookSurface` construction (around line 60) and change it to pass the shared runner + classifier:

```csharp
var runner     = Globals.ThisAddIn?.AdvancedSearchRunner;
var classifier = Globals.ThisAddIn?.FolderClassifier;
if (marshaller != null && ids != null && app != null && runner != null)
{
    _surface = new LiveOutlookSurface(app, marshaller, ids,
        composeInspector: null,
        explorer: explorer,
        runner: runner,
        classifier: classifier);
    _toolHost = new OutlookToolHost(_surface, Config.WriteToolsEnabled);
    TraceLog.Write("surface + toolHost constructed", "InboxCopilotPane");
}
else
{
    TraceLog.Write("surface NOT constructed (missing service); runner=" + (runner != null),
        "InboxCopilotPane");
}
```

- [ ] **Step 6: Update `AITaskPane.Bind` to pass the runner**

In `VSTO2/OutlookAI/TaskPane/AITaskPane.cs`, change the `LiveOutlookSurface` construction near line 81 the same way:

```csharp
var runner     = Globals.ThisAddIn?.AdvancedSearchRunner;
var classifier = Globals.ThisAddIn?.FolderClassifier;
if (marshaller != null && ids != null && app != null && runner != null)
{
    _surface = new LiveOutlookSurface(app, marshaller, ids,
        composeInspector: inspector,
        explorer: null,
        runner: runner,
        classifier: classifier);
    _toolHost = new OutlookToolHost(_surface, Config.WriteToolsEnabled);
    TraceLog.Write("surface + toolHost constructed", "AITaskPane");
}
else
{
    TraceLog.Write("surface NOT constructed (missing service); runner=" + (runner != null),
        "AITaskPane");
}
```

- [ ] **Step 7: Remove or hard-deprecate the old runner-less `LiveOutlookSurface` constructor**

In `LiveOutlookSurface.cs`, if a constructor overload without the `runner` parameter still exists from Task 8 Step 4, delete it now. No production call site should reach it any more, and tests do not construct `LiveOutlookSurface` directly.

- [ ] **Step 8: Build**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: clean build.

- [ ] **Step 9: Run full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: all tests pass.

- [ ] **Step 10: Commit**

```powershell
git add VSTO2/OutlookAI/ThisAddIn.cs VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPane.cs VSTO2/OutlookAI/TaskPane/AITaskPane.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs
git commit -m "feat(search): wire singleton AdvancedSearch host + runner through ThisAddIn" -m "ThisAddIn now owns one LiveAdvancedSearchHost and one OutlookAdvancedSearchRunner for the process lifetime, and exposes the runner + classifier to both task panes. InboxCopilotPane and AITaskPane construct LiveOutlookSurface with the shared instances so all search calls flow through the same AdvancedSearchComplete subscription and serialisation semaphore. Shutdown disposes runner then host in the right order."
```

---

## Task 11: Add the Stop affordance + live time counter to the compact chat status UI

Surface real cancellation to the user. The cancellation chain (`ChatSession.CancelCurrent()` → tool CT → `LiveOutlookSurface` → runner → `search.Stop()`) was wired in earlier tasks; this task only adds the JS / CSS / wiring on the WebView2 side.

**Files:**
- Modify: the compact tool-status renderer in the `inbox-copilot` WebView2 bundle. Identify the exact files in your bundle (typical names: `Assets/inbox-copilot.js`, `Assets/inbox-copilot.css`, or a TSX/JSX equivalent).

- [ ] **Step 1: Locate the compact status renderer**

Search the repo for the existing "Searching messages" / compact-status string introduced in commit `9950593`:

```powershell
rg -n "Searching" VSTO2\OutlookAI\TaskPane\InboxCopilot
```

Open the file that emits the compact status line.

- [ ] **Step 2: Render a Stop link and a live time counter**

Wherever the compact status text is rendered, add (sketch — adapt to the actual templating style in the file):

```js
function renderToolStatus(state) {
    const start = state.startedAtMs ?? Date.now();
    const seconds = Math.max(0, Math.floor((Date.now() - start) / 1000));
    return `${state.label} (${seconds}s) <a href="#" class="oai-stop-link" data-action="stop">Stop</a>`;
}

// In the click handler:
document.addEventListener('click', (ev) => {
    const t = ev.target;
    if (t && t.classList && t.classList.contains('oai-stop-link')) {
        ev.preventDefault();
        window.chrome?.webview?.postMessage({ type: 'cancel' });
    }
});

// Tick the counter once a second while a tool is in flight.
setInterval(() => {
    if (currentToolState && currentToolState.inFlight) renderToolStatus(currentToolState);
}, 1000);
```

- [ ] **Step 3: Add CSS for the Stop link**

In the matching CSS file:

```css
.oai-stop-link {
    margin-left: 0.5em;
    font-size: 0.85em;
    color: #b00020;
    text-decoration: underline;
    cursor: pointer;
}
.oai-stop-link:hover { color: #800000; }
```

- [ ] **Step 4: Handle the `cancel` message on the C# side**

In the WebView2 bridge for the Inbox Copilot pane (find the existing `WebMessageReceived` handler), add a case for the `cancel` type that invokes the existing `ChatSession.CancelCurrent()`:

```csharp
case "cancel":
    try { _chatSession?.CancelCurrent(); }
    catch (Exception ex) { TraceLog.Write("CancelCurrent: " + ex.Message, "InboxCopilotPane"); }
    break;
```

If `ChatSession.CancelCurrent()` does not yet exist, expose it now: it cancels the `CancellationTokenSource` owned by the current in-flight turn. Follow the existing per-turn CTS pattern (see `CodexChatService.RunTurnAsync` and how the dispatcher receives a CT).

- [ ] **Step 5: Build, run full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: clean build, all tests pass.

- [ ] **Step 6: Commit**

```powershell
git add -A
git commit -m "feat(ui): add Stop affordance and live time counter to compact tool status" -m "Compact chat status line gains a small Stop link and a live (Ns) counter while a tool call is in flight. Clicking Stop posts a cancel message to the WebView2 host, which calls ChatSession.CancelCurrent() — this propagates through the existing tool CancellationToken chain to OutlookAdvancedSearchRunner, which invokes search.Stop() on the in-flight Outlook search."
```

---

## Task 12: Final verification, deploy, and smoke

Build Release, publish, install elevated from the staging path, then smoke the two scenarios that drove this phase.

- [ ] **Step 1: Run the full test suite one more time**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes. Note total count for the commit message.

- [ ] **Step 2: Publish Release to staging**

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" `
    /p:PublishDir="$staging\" /v:minimal /nologo
Copy-Item -LiteralPath "Deploy\Install-OutlookAI.ps1" -Destination "$staging\" -Force
```

Expected: Release publish completes cleanly; staging path contains `OutlookAI.vsto`, `Application Files`, and the installer script.

- [ ] **Step 3: Install with the elevated installer using the explicit staging `-SourcePath`**

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
$script  = Join-Path $staging "Install-OutlookAI.ps1"
$args    = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -SourcePath `"$staging`""
$proc    = Start-Process -FilePath "powershell.exe" -ArgumentList $args -Verb RunAs -Wait -PassThru
"Installer exit code: $($proc.ExitCode)"
```

Expected: exit code `0`.

- [ ] **Step 4: Verify the installed DLL hash matches staging**

```powershell
$staged    = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2\OutlookAI.dll"
$installed = "C:\Program Files\OutlookAI\OutlookAI.dll"
$a = Get-FileHash -LiteralPath $staged    -Algorithm SHA256
$b = Get-FileHash -LiteralPath $installed -Algorithm SHA256
"staged=$($a.Hash)"; "installed=$($b.Hash)"; "match=$($a.Hash -eq $b.Hash)"
```

Expected: `match=True`.

- [ ] **Step 5: Smoke — oldest email query**

In Outlook with the Inbox Copilot pane open, send:

```
What is my oldest email?
```

While the search runs, verify:
- The Outlook UI stays responsive: you can move the mouse, click the ribbon, switch folders, hover items.
- The compact status line shows `Searching mailbox (Ns)…` with a `Stop` link.
- Within ~30s either the assistant returns the actual oldest email (with subject, date) OR returns a timeout error message.

Then read the trace:

```powershell
Get-Content -LiteralPath "C:\Users\MDASR\AppData\Local\OutlookAI\trace.log" -Tail 80
```

Expected entries:
- `AdvancedSearch Start tag=...`
- `AdvancedSearch Complete tag=... raw_count=N`
- `SearchMessages primary=AdvancedSearch complete raw_count=N`
- No `fallback reason=...` entry for this run.

- [ ] **Step 6: Smoke — EIN query**

Send:

```
Find an email with EIN
```

Verify the assistant returns the EIN email previously seen (`MDASR EIN Number - 11-3039659`) and the trace shows the same `primary=AdvancedSearch complete` line.

- [ ] **Step 7: Smoke — Stop button**

Send:

```
Find every email older than 2010 across all of my mail.
```

While the search runs, click `Stop` in the compact status line. Verify within ~1 second:
- The assistant shows a cancelled response (model receives the cancel envelope).
- Outlook remains responsive.
- Trace shows `AdvancedSearch Stop tag=...` and `SearchMessages ... cancelled` (or equivalent).

- [ ] **Step 8: Push the branch**

If all three smokes pass:

```powershell
git push origin feature/codex-oauth-migration
```

- [ ] **Step 9: Final commit summarising the phase (optional, if any squashing or notes are needed)**

If smoke uncovered any tiny fixes, commit them on top with `fix(search): ...` messages. Otherwise the phase is complete with the per-task commits.

---

## Self-Review (run before invoking executor)

**Spec coverage:**
- Goal 1 (no UI freeze) → Tasks 6, 8, 11 (engine + fallback yielding + Stop affordance).
- Goal 2 (native engine) → Tasks 6, 7, 8.
- Goal 3 (cancellable) → Tasks 1, 6, 8, 11.
- Goal 4 (consistent with Outlook search box) → Task 6/7 (uses AdvancedSearch).
- Goal 5 (fallback) → Task 8 Step 7.
- Acceptance 1 (responsiveness) → Smoke Step 5.
- Acceptance 2 (oldest returns) → Smoke Step 5.
- Acceptance 3 (EIN finds it) → Smoke Step 6.
- Acceptance 4 (Stop within ~1s) → Smoke Step 7.
- Acceptance 5 (all tests pass) → Task 12 Step 1.
- Acceptance 6 (trace shows primary path) → Smoke Step 5.

**Placeholder scan:** no "TBD" / "implement later"; each step has the exact code or command. Task 11 references "adapt to the actual templating style in the file" because the WebView2 bundle's exact framework (vanilla JS vs. TSX) was not in scope to discover here — the executor must read the existing renderer and match it. Acceptable scope-limited handoff.

**Type consistency:**
- `IFolderClassifier.IsSystemFolder(string, bool)` — same in Tasks 2, 3, 7, 8.
- `MessageProjectionInput` fields — same in Tasks 3, 6, 7, 8.
- `SearchResultProjector.Project(items, args, classifier)` — same in Tasks 3, 8.
- `IAdvancedSearchHost.Start/Stop/Completed` — same in Tasks 5, 6, 7.
- `AdvancedSearchResult` (`Completed`, `TimedOut`, `Cancelled`, `Error`, `Items`) — same in Tasks 6, 8.
- `IOutlookAdvancedSearchRunner.RunAsync(scope, filter, sub, timeout, ct)` — same in Tasks 6, 8.
- `LiveOutlookSurface` constructor `(app, marshaller, ids, composeInspector, explorer, runner, classifier=null)` — same in Tasks 8, 10.

No mismatches found.

