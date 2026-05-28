# Bulk Export + Search Truncation Feedback Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop silent truncation of large `outlook_search_messages` results, and add a deterministic server-side bulk-list Excel export so attorneys get complete results instead of a silently-capped subset.

**Architecture:** Two complementary mechanisms. (1) `outlook_search_messages` returns `total_matches` + `truncated` so the model knows when it got an incomplete set. (2) A new `outlook_export_search_results` tool counts the true total, collects up to a configurable ceiling (default 2,000) via the existing fast Table-API search path, projects mechanical columns, builds an `.xlsx` through the existing `ExcelWorkbookBuilder`, and reports "exported N of M". No server-side date-window walking (the 100-cap is only in the model-facing parser, not the surface).

**Tech Stack:** C# 7.3 / .NET Framework 4.7.2, VSTO, Newtonsoft.Json, ClosedXML, xUnit. VS MSBuild + VSTest (NOT `dotnet`).

**Spec:** `docs/superpowers/specs/2026-05-28-bulk-export-and-truncation-feedback-design.md`

**Baseline:** 607 tests passing on `master` (`d8ee686`). Branch `feature/bulk-export-truncation-feedback`.

**Tool reminders:**
- Build: `& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo`
- Test project only build: `& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo`
- Test: `& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"`
- Focused: append `/TestCaseFilter:"FullyQualifiedName~<ClassName>"`
- Pre-existing benign `MSB3277` warning is expected; nothing else should warn.
- C# 7.3: no switch expressions, no default interface methods. `is T t` pattern matching and classic `switch` are fine.

---

## File Structure

| File | New/Modify | Responsibility |
|---|---|---|
| `VSTO2/OutlookAI/Config.cs` | Modify | Add `MaxBulkExportRows` (default 2000), global-config load + clamp. |
| `VSTO2/OutlookAI/Services/Tools/SearchResult.cs` | Create | Value type: clamped Messages + TotalMatches + Truncated. |
| `VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs` | Modify | Return `SearchResult` (count + truncated + clamped list). |
| `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs` | Modify | `SearchMessages` returns `SearchResult`. |
| `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs` | Modify | Return `SearchResult`; force `Truncated` when early-stop fired. |
| `VSTO2/OutlookAI/TaskPane/AITaskPane.cs` | Modify | `NullSurface.SearchMessages` returns `SearchResult`. |
| `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs` | Modify | Emit `total_matches` + `truncated`. |
| `VSTO2/OutlookAI/Services/Tools/ExportSearchResultsArgs.cs` | Create | Filter + column selection + filename hint. |
| `VSTO2/OutlookAI/Services/Tools/ExportSearchResultsArgsParser.cs` | Create | Parse + validate (reuses search filter parse; column allow-list). |
| `VSTO2/OutlookAI/Services/Tools/OutlookExportSearchResultsTool.cs` | Create | The bulk tool: count → collect → project → ExportExcel → report N/M. |
| `VSTO2/OutlookAI/Services/OutlookToolHost.cs` | Modify | Register the new tool. |
| `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs` | Modify | New tool entry; rewrite truncation/chunking steering. |
| `VSTO2/OutlookAI/OutlookAI.csproj` | Modify | Register the 4 new `.cs` files. |
| `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs` | Modify | `SearchMessages` returns `SearchResult`. |
| `VSTO2/OutlookAI.Tests/Services/Tools/SearchResultProjectorTests.cs` | Modify | Assert on `SearchResult` shape. |
| `VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs` | Modify | Fake returns `SearchResult`; assert new fields. |
| `VSTO2/OutlookAI.Tests/Services/Tools/ExportSearchResultsArgsParserTests.cs` | Create | Parser validation tests. |
| `VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportSearchResultsToolTests.cs` | Create | Bulk tool behavior tests. |
| `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs` | Modify | New-tool-present + steering assertions. |
| `VSTO2/OutlookAI.Tests/ConfigTests.cs` | Modify | `MaxBulkExportRows` tests. |

The test project is SDK-style and globs `**/*.cs`, so new test files need no csproj edit. New **production** `.cs` files MUST be added to `VSTO2/OutlookAI/OutlookAI.csproj`.

---

## Task 1: Config.MaxBulkExportRows

**Files:**
- Modify: `VSTO2/OutlookAI/Config.cs`
- Test: `VSTO2/OutlookAI.Tests/ConfigTests.cs`

- [ ] **Step 1: Write the failing tests**

Append these tests to the `ConfigTests` class in `VSTO2/OutlookAI.Tests/ConfigTests.cs` (before the closing brace):

```csharp
[Fact]
public void MaxBulkExportRows_DefaultsTo2000()
{
    var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
    var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
    Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
    Assert.Equal(2000, Config.MaxBulkExportRows);
}

[Fact]
public void MaxBulkExportRows_LoadsFromGlobalConfig()
{
    var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
    var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
    File.WriteAllText(g, "<Config><MaxBulkExportRows>500</MaxBulkExportRows></Config>");
    try
    {
        Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
        Assert.Equal(500, Config.MaxBulkExportRows);
    }
    finally { if (File.Exists(g)) File.Delete(g); }
}

[Fact]
public void MaxBulkExportRows_ClampsToFloorAndCeiling()
{
    var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
    var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

    File.WriteAllText(g, "<Config><MaxBulkExportRows>0</MaxBulkExportRows></Config>");
    try
    {
        Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
        Assert.Equal(1, Config.MaxBulkExportRows);   // floor

        File.WriteAllText(g, "<Config><MaxBulkExportRows>999999</MaxBulkExportRows></Config>");
        Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
        Assert.Equal(50000, Config.MaxBulkExportRows);  // ceiling
    }
    finally { if (File.Exists(g)) File.Delete(g); }
}

[Fact]
public void MaxBulkExportRows_NotUserOverridable()
{
    var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
    var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
    File.WriteAllText(g, "<Config><MaxBulkExportRows>750</MaxBulkExportRows></Config>");
    File.WriteAllText(u, "<Config><MaxBulkExportRows>3000</MaxBulkExportRows></Config>");
    try
    {
        Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
        Assert.Equal(750, Config.MaxBulkExportRows);  // user value ignored
    }
    finally { if (File.Exists(g)) File.Delete(g); if (File.Exists(u)) File.Delete(u); }
}
```

- [ ] **Step 2: Run the tests to verify they fail (RED)**

Run: test-project build, expect `CS0117: 'Config' does not contain a definition for 'MaxBulkExportRows'`.

- [ ] **Step 3: Add the Config field + default constant**

In `VSTO2/OutlookAI/Config.cs`, after the `DefaultWriteToolsEnabled` constant (line 24) add:

```csharp
        public const int DefaultMaxBulkExportRows = 2000;
        private const int MinBulkExportRows = 1;
        private const int MaxBulkExportRowsCeiling = 50000;
```

After the `WriteToolsEnabled` property (around line 46) add:

```csharp
        /// <summary>
        /// Hard ceiling on rows collected by outlook_export_search_results.
        /// Server-authoritative (global config only); not user-overridable.
        /// Bounds runtime/memory on large mailboxes. Clamped to
        /// [MinBulkExportRows, MaxBulkExportRowsCeiling] on load.
        /// </summary>
        public static int MaxBulkExportRows { get; set; } = DefaultMaxBulkExportRows;
```

- [ ] **Step 4: Reset + load the field**

In `ResetDefaults()` (after the `WriteToolsEnabled = DefaultWriteToolsEnabled;` line) add:

```csharp
            MaxBulkExportRows = DefaultMaxBulkExportRows;
```

In `LoadFromFile`, inside the `if (!allowServerFields) return;` server-only block (after the `codexAuthPath` block, near the `voiceModel` block), add:

```csharp
                var maxBulkExportRows = root.Element("MaxBulkExportRows");
                if (maxBulkExportRows != null && int.TryParse(maxBulkExportRows.Value, out var mber))
                {
                    if (mber < MinBulkExportRows) mber = MinBulkExportRows;
                    if (mber > MaxBulkExportRowsCeiling) mber = MaxBulkExportRowsCeiling;
                    MaxBulkExportRows = mber;
                }
```

- [ ] **Step 5: Run tests (GREEN)**

Run: `/TestCaseFilter:"FullyQualifiedName~ConfigTests"`. Expect all ConfigTests pass (existing + 4 new).

- [ ] **Step 6: Commit**

```
git add VSTO2/OutlookAI/Config.cs VSTO2/OutlookAI.Tests/ConfigTests.cs
git commit -m "feat(config): add MaxBulkExportRows ceiling (global, default 2000)"
```

---

## Task 2: SearchResult type + projector returns it (interface stable)

This introduces `SearchResult` and changes `SearchResultProjector.Project` to return it. `LiveOutlookSurface` unwraps `.Messages` so the `IOutlookSurface.SearchMessages` signature stays unchanged this task (build stays green; interface change is Task 3).

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/SearchResult.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs` (lines 216-218, 1309-1311)
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/Tools/SearchResultProjectorTests.cs`

- [ ] **Step 1: Create `SearchResult.cs`**

```csharp
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Result of a message search: the clamped page of summaries the model
    /// sees, plus the pre-clamp total and whether the page was truncated.
    /// TotalMatches is exact where the collection path saw every match (the
    /// common from:/subject:/body: case) and a floor where an early-stop
    /// fired (to:/broad-no-filter scans); in the floor case Truncated is
    /// forced true so the model never believes a capped page is complete.
    /// </summary>
    public sealed class SearchResult
    {
        public IReadOnlyList<MessageSummary> Messages { get; set; } = new MessageSummary[0];
        public int TotalMatches { get; set; }
        public bool Truncated { get; set; }
    }
}
```

- [ ] **Step 2: Update the projector tests (RED)**

In `VSTO2/OutlookAI.Tests/Services/Tools/SearchResultProjectorTests.cs`, the existing tests call `SearchResultProjector.Project(...)` and treat the return as the list. Update each call site to read `.Messages`, and add two new tests for the count/truncated fields. Replace the test that checks clamping (the `MaxResults = 3` case around line 50) and add new ones. Concretely, add these two tests to the class:

```csharp
[Fact]
public void Project_ReportsTotalMatchesAndTruncated_WhenClamped()
{
    var inputs = MakeInputs(10);   // 10 candidates
    var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 3 };

    var result = SearchResultProjector.Project(inputs, args, new FolderClassifier());

    Assert.Equal(3, result.Messages.Count);
    Assert.Equal(10, result.TotalMatches);
    Assert.True(result.Truncated);
}

[Fact]
public void Project_NotTruncated_WhenUnderLimit()
{
    var inputs = MakeInputs(2);
    var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 25 };

    var result = SearchResultProjector.Project(inputs, args, new FolderClassifier());

    Assert.Equal(2, result.Messages.Count);
    Assert.Equal(2, result.TotalMatches);
    Assert.False(result.Truncated);
}
```

Add this helper to the test class if one does not already exist (adapt to the existing `MessageProjectionInput` construction already used in the file — match its property names exactly):

```csharp
private static List<MessageProjectionInput> MakeInputs(int n)
{
    var list = new List<MessageProjectionInput>();
    for (var i = 0; i < n; i++)
    {
        list.Add(new MessageProjectionInput
        {
            Id = "id-" + i,
            Subject = "Subject " + i,
            From = "sender" + i + "@example.com",
            To = new[] { "me@example.com" },
            ReceivedAt = new System.DateTimeOffset(2026, 1, 1, 0, 0, 0, System.TimeSpan.Zero).AddMinutes(i),
            FolderName = "Inbox",
            FolderDefaultItemTypeIsMail = true,
            HasAttachments = false,
            SnippetFactory = () => "snippet " + i,
        });
    }
    return list;
}
```

Also update every PRE-EXISTING test in this file that did `var result = SearchResultProjector.Project(...)` followed by list assertions (e.g. `result.Count`, `result[0]`) to use `result.Messages.Count` / `result.Messages[0]`.

Run the test-project build: expect compile errors because `Project` still returns the list type / `.Messages` does not exist yet → RED.

- [ ] **Step 3: Change the projector to return `SearchResult`**

Replace the body of `SearchResultProjector.Project` in `VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs`. Change the signature return type from `IReadOnlyList<MessageSummary>` to `SearchResult`, and replace the tail (from `var ordered = ...` onward) with:

```csharp
            var ordered = string.Equals(args.SortOrder, "oldest", StringComparison.OrdinalIgnoreCase)
                ? filtered.OrderBy(i => i.ReceivedAt)
                : filtered.OrderByDescending(i => i.ReceivedAt);

            var orderedList = ordered.ToList();
            var total = orderedList.Count;

            var maxResults = args.MaxResults > 0 ? args.MaxResults : 25;
            var top = orderedList.Take(maxResults).ToList();

            var output = new List<MessageSummary>(top.Count);
            foreach (var i in top)
            {
                string snippet = "";
                try { snippet = i.SnippetFactory != null ? (i.SnippetFactory() ?? "") : ""; }
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

            return new SearchResult
            {
                Messages = output,
                TotalMatches = total,
                Truncated = total > output.Count,
            };
```

Also update the XML doc summary line to mention it returns a `SearchResult`.

- [ ] **Step 4: Keep `LiveOutlookSurface.SearchMessages` returning the list (unwrap `.Messages`)**

In `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`, two call sites consume `Project(...)`:

Lines 216-218 (AdvancedSearch path):
```csharp
                    return _marshaller.RunAsync(
                        () => SearchResultProjector.Project(primary.Items, args, _classifier),
                        ct).GetAwaiter().GetResult();
```
becomes:
```csharp
                    return _marshaller.RunAsync(
                        () => SearchResultProjector.Project(primary.Items, args, _classifier).Messages,
                        ct).GetAwaiter().GetResult();
```

Lines 1309-1311 (`FallbackIterativeSearch` tail):
```csharp
            return _marshaller.RunAsync(
                () => SearchResultProjector.Project(allInputs, args, _classifier),
                ct).GetAwaiter().GetResult();
```
becomes:
```csharp
            return _marshaller.RunAsync(
                () => SearchResultProjector.Project(allInputs, args, _classifier).Messages,
                ct).GetAwaiter().GetResult();
```

(`SearchMessages` / `FallbackIterativeSearch` still declare `IReadOnlyList<MessageSummary>` this task.)

- [ ] **Step 5: Register `SearchResult.cs` in the csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, in the `Services\Tools\` `<Compile>` group, add (alphabetical, near the other Search* files):

```xml
    <Compile Include="Services\Tools\SearchResult.cs" />
```

- [ ] **Step 6: Build + run (GREEN)**

Run full solution build (only `MSB3277`), then `/TestCaseFilter:"FullyQualifiedName~SearchResultProjectorTests"`. Expect all pass.

- [ ] **Step 7: Commit**

```
git add VSTO2/OutlookAI/Services/Tools/SearchResult.cs VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/Tools/SearchResultProjectorTests.cs
git commit -m "feat(search): SearchResultProjector returns total + truncated"
```

---

## Task 3: `IOutlookSurface.SearchMessages` returns `SearchResult` + search tool feedback

Now flip the interface to return `SearchResult`, update all implementers/fakes, force `Truncated` on early-stop, and have the search tool emit `total_matches` + `truncated`.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs:18`
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs` (SearchMessages + FallbackIterativeSearch signatures/returns; early-stop flag)
- Modify: `VSTO2/OutlookAI/TaskPane/AITaskPane.cs:605` (NullSurface)
- Modify: `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs:19`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs`

- [ ] **Step 1: Update the search-tool tests (RED)**

In `VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs`:

The fake surface (around lines 186-189) currently has:
```csharp
            public Func<SearchMessagesArgs, IReadOnlyList<MessageSummary>> OnSearch { get; set; }
            public Func<SearchMessagesArgs, CancellationToken, IReadOnlyList<MessageSummary>> OnSearchWithCt { get; set; }

            public override IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
```
Change the override's return type to `SearchResult` and wrap list returns. Replace those lines with:
```csharp
            public Func<SearchMessagesArgs, IReadOnlyList<MessageSummary>> OnSearch { get; set; }
            public Func<SearchMessagesArgs, CancellationToken, IReadOnlyList<MessageSummary>> OnSearchWithCt { get; set; }
            public int? TotalMatchesOverride { get; set; }

            public override SearchResult SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
```
Inside that method, wherever it currently returns an `IReadOnlyList<MessageSummary>` (e.g. `return OnSearch(args);`), wrap it. Find the return(s) and convert to:
```csharp
                var list = /* existing expression that produced the list */;
                return new SearchResult
                {
                    Messages = list ?? new MessageSummary[0],
                    TotalMatches = TotalMatchesOverride ?? (list?.Count ?? 0),
                    Truncated = TotalMatchesOverride.HasValue && TotalMatchesOverride.Value > (list?.Count ?? 0),
                };
```

Add a new test asserting the tool surfaces the new fields:
```csharp
[Fact]
public void Execute_IncludesTotalMatchesAndTruncated_WhenResultCapped()
{
    var surface = new FakeSurface
    {
        OnSearch = _ => new[]
        {
            new MessageSummary { Id = "a", Subject = "s", From = "f", To = new[] { "t" }, ReceivedAt = System.DateTimeOffset.UtcNow, Snippet = "x" },
        },
        TotalMatchesOverride = 42,
    };
    var tool = new OutlookSearchMessagesTool();

    var json = tool.ExecuteAsync("{\"from\":\"alice\"}", surface, System.Threading.CancellationToken.None).GetAwaiter().GetResult();
    var obj = Newtonsoft.Json.Linq.JObject.Parse(json);

    Assert.Equal(42, (int)obj["total_matches"]);
    Assert.True((bool)obj["truncated"]);
    Assert.Single((Newtonsoft.Json.Linq.JArray)obj["messages"]);
}
```

Run the test-project build: expect compile errors (interface mismatch / `total_matches` missing) → RED.

- [ ] **Step 2: Flip the interface**

In `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs:18`, change:
```csharp
        IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken));
```
to:
```csharp
        SearchResult SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken));
```

- [ ] **Step 3: Update `LiveOutlookSurface`**

Change `SearchMessages` (line 163) return type to `SearchResult`. Its AdvancedSearch return (now `.Messages` from Task 2) becomes the full `SearchResult`:

Lines 216-218 become:
```csharp
                    return _marshaller.RunAsync(
                        () => SearchResultProjector.Project(primary.Items, args, _classifier),
                        ct).GetAwaiter().GetResult();
```

Change `FallbackIterativeSearch` (line 1241 signature) return type to `SearchResult`. Add an early-stop flag and force `Truncated`. Replace the `break;` inside the recipient early-stop block and the broad early-stop block so each sets a flag first, and replace the tail return. Concretely:

Add near the top of `FallbackIterativeSearch` (after `var allInputs = ...`):
```csharp
            var earlyStop = false;
```
In the recipient early-stop block, change `break;` (line 1290) to:
```csharp
                    earlyStop = true;
                    break;
```
In the broad early-stop block, change `break;` (line 1305) to:
```csharp
                    earlyStop = true;
                    break;
```
Replace the tail (lines 1309-1311) with:
```csharp
            var projected = _marshaller.RunAsync(
                () => SearchResultProjector.Project(allInputs, args, _classifier),
                ct).GetAwaiter().GetResult();
            if (earlyStop) projected.Truncated = true;
            return projected;
```

Also change the `SearchMessages` method's declared return type on line 163 from `IReadOnlyList<MessageSummary>` to `SearchResult`, and `FallbackIterativeSearch`'s declared return type to `SearchResult`.

- [ ] **Step 4: Update `NullSurface` in `AITaskPane.cs`**

Line 605 currently:
```csharp
            public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => new MessageSummary[0];
```
becomes:
```csharp
            public SearchResult SearchMessages(SearchMessagesArgs args, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => new SearchResult();
```

- [ ] **Step 5: Update `MinimalSurface` test base**

In `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs:19`:
```csharp
        public virtual IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
```
Change the return type to `SearchResult` and update its body to return `new SearchResult()` (or wrap whatever it currently returns). Match the file's existing default-return style.

- [ ] **Step 6: Emit the new fields in the search tool**

In `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs`, replace lines 23-36 (the `var search = ...` through the `return Task.FromResult(...)`) with:

```csharp
                var search = SearchMessagesArgsParser.ParseSearch(argsJson);
                var result = surface.SearchMessages(search, ct) ?? new SearchResult();
                var hits = result.Messages ?? new MessageSummary[0];

                var json = new JObject(
                    new JProperty("messages", new JArray(hits.Select(m =>
                        new JObject(
                            new JProperty("id", m.Id ?? ""),
                            new JProperty("subject", m.Subject ?? ""),
                            new JProperty("from", m.From ?? ""),
                            new JProperty("to", new JArray((m.To ?? new string[0]).Cast<object>())),
                            new JProperty("received_at", m.ReceivedAt.ToString("o")),
                            new JProperty("snippet", m.Snippet ?? ""),
                            new JProperty("has_attachments", m.HasAttachments))))),
                    new JProperty("total_matches", result.TotalMatches),
                    new JProperty("truncated", result.Truncated));
                return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
```

- [ ] **Step 7: Build + run (GREEN)**

Full solution build (only `MSB3277`). Then run the full suite. Expect previous count + the 1 new tool test (and the Task 2 projector tests). No failures.

- [ ] **Step 8: Commit**

```
git add VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI/TaskPane/AITaskPane.cs VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs
git commit -m "feat(search): surface total_matches + truncated to the model"
```

---

## Task 4: ExportSearchResultsArgs + parser

The bulk tool's args: a search filter (reuse the same field names as `outlook_search_messages`) plus a `columns` allow-list and optional `filename_hint` / `sheet_name`.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/ExportSearchResultsArgs.cs`
- Create: `VSTO2/OutlookAI/Services/Tools/ExportSearchResultsArgsParser.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/Tools/ExportSearchResultsArgsParserTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/ExportSearchResultsArgsParserTests.cs`:

```csharp
using System.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class ExportSearchResultsArgsParserTests
    {
        [Fact]
        public void Parse_ExtractsFilterAndColumns()
        {
            var json = "{\"from\":\"IT Creations\",\"scope\":\"all_mail\",\"columns\":[\"subject\",\"from\",\"received_at\",\"snippet\"],\"filename_hint\":\"vendor\"}";
            var args = ExportSearchResultsArgsParser.Parse(json);

            Assert.Equal("IT Creations", args.Filter.From);
            Assert.Equal("all_mail", args.Filter.Scope);
            Assert.Equal(new[] { "subject", "from", "received_at", "snippet" }, args.Columns.ToArray());
            Assert.Equal("vendor", args.FilenameHint);
        }

        [Fact]
        public void Parse_DropsUnknownColumns()
        {
            var json = "{\"columns\":[\"subject\",\"bogus\",\"snippet\"]}";
            var args = ExportSearchResultsArgsParser.Parse(json);
            Assert.Equal(new[] { "subject", "snippet" }, args.Columns.ToArray());
        }

        [Fact]
        public void Parse_DefaultsColumnsWhenNoneProvided()
        {
            var args = ExportSearchResultsArgsParser.Parse("{\"from\":\"x\"}");
            Assert.Equal(ExportSearchResultsArgsParser.DefaultColumns, args.Columns.ToArray());
        }

        [Fact]
        public void Parse_AllUnknownColumns_FallsBackToDefaults()
        {
            var args = ExportSearchResultsArgsParser.Parse("{\"columns\":[\"bogus\",\"nope\"]}");
            Assert.Equal(ExportSearchResultsArgsParser.DefaultColumns, args.Columns.ToArray());
        }

        [Fact]
        public void Parse_DeduplicatesColumns()
        {
            var args = ExportSearchResultsArgsParser.Parse("{\"columns\":[\"subject\",\"subject\",\"from\"]}");
            Assert.Equal(new[] { "subject", "from" }, args.Columns.ToArray());
        }
    }
}
```

- [ ] **Step 2: Build test project to verify RED**

Expect `CS0246`/`CS0103` for `ExportSearchResultsArgsParser` / `ExportSearchResultsArgs`.

- [ ] **Step 3: Create `ExportSearchResultsArgs.cs`**

```csharp
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Args for outlook_export_search_results: a search filter (same shape as
    /// outlook_search_messages) plus the mechanical columns to project and an
    /// optional filename hint. No body reads — columns come from the search
    /// projection only.
    /// </summary>
    public sealed class ExportSearchResultsArgs
    {
        public SearchMessagesArgs Filter { get; set; }
        public IReadOnlyList<string> Columns { get; set; }
        public string FilenameHint { get; set; }
        public string SheetName { get; set; }
    }
}
```

- [ ] **Step 4: Create `ExportSearchResultsArgsParser.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Parses outlook_export_search_results args. Reuses SearchMessagesArgsParser
    /// for the filter (count-mode shape — no max_results clamp), validates the
    /// requested columns against a fixed allow-list of mechanical fields, drops
    /// unknown / duplicate columns, and falls back to DefaultColumns when none
    /// survive.
    /// </summary>
    internal static class ExportSearchResultsArgsParser
    {
        public static readonly string[] AllowedColumns =
        {
            "subject", "from", "to", "received_at", "snippet", "has_attachments", "folder",
        };

        public static readonly string[] DefaultColumns =
        {
            "received_at", "from", "to", "subject", "snippet",
        };

        private const string DefaultFilenameHint = "OutlookAI-Search-Export";

        public static ExportSearchResultsArgs Parse(string argsJson)
        {
            var raw = string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson;

            // Reuse the existing filter parser (count shape: no max_results cap).
            var filter = SearchMessagesArgsParser.ParseCount(raw);

            var obj = JObject.Parse(raw);

            var columns = ParseColumns(obj["columns"]);
            var filenameHint = CleanString(obj["filename_hint"]) ?? DefaultFilenameHint;
            var sheetName = CleanString(obj["sheet_name"]) ?? filenameHint;

            return new ExportSearchResultsArgs
            {
                Filter = filter,
                Columns = columns,
                FilenameHint = filenameHint,
                SheetName = sheetName,
            };
        }

        private static IReadOnlyList<string> ParseColumns(JToken token)
        {
            var arr = token as JArray;
            if (arr == null) return DefaultColumns;

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ordered = new List<string>();
            foreach (var t in arr)
            {
                var name = (t?.Type == JTokenType.String) ? ((string)t)?.Trim().ToLowerInvariant() : null;
                if (string.IsNullOrEmpty(name)) continue;
                if (!AllowedColumns.Contains(name)) continue;
                if (!seen.Add(name)) continue;
                ordered.Add(name);
            }

            return ordered.Count > 0 ? (IReadOnlyList<string>)ordered : DefaultColumns;
        }

        private static string CleanString(JToken token)
        {
            if (token == null || token.Type != JTokenType.String) return null;
            var v = ((string)token)?.Trim();
            return string.IsNullOrEmpty(v) ? null : v;
        }
    }
}
```

Note: `ParseCount` sets `MaxResults = int.MaxValue`; the bulk tool overrides `MaxResults` to the ceiling before searching (Task 5). Reusing `ParseCount` avoids re-implementing all the filter-field parsing.

- [ ] **Step 5: Register both files in csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj` `Services\Tools\` `<Compile>` group add:
```xml
    <Compile Include="Services\Tools\ExportSearchResultsArgs.cs" />
    <Compile Include="Services\Tools\ExportSearchResultsArgsParser.cs" />
```

- [ ] **Step 6: Build + run (GREEN)**

Full build, then `/TestCaseFilter:"FullyQualifiedName~ExportSearchResultsArgsParserTests"`. Expect 5 pass.

- [ ] **Step 7: Commit**

```
git add VSTO2/OutlookAI/Services/Tools/ExportSearchResultsArgs.cs VSTO2/OutlookAI/Services/Tools/ExportSearchResultsArgsParser.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/Tools/ExportSearchResultsArgsParserTests.cs
git commit -m "feat(export): add ExportSearchResultsArgs + parser (column allow-list)"
```

---

## Task 5: OutlookExportSearchResultsTool

Orchestrates `CountMessages` (true total M) → `SearchMessages` capped at the ceiling (collect N) → project mechanical columns → `ExportExcel` → report `{ file, exported:N, total_matches:M, truncated:(N<M) }`. Zero matches → structured `no_matches`, no file.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/OutlookExportSearchResultsTool.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportSearchResultsToolTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportSearchResultsToolTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookExportSearchResultsToolTests
    {
        private sealed class FakeSurface : MinimalSurface
        {
            public int Count { get; set; }
            public List<MessageSummary> Hits { get; set; } = new List<MessageSummary>();
            public ExportExcelArgs CapturedExcel { get; set; }
            public int? CapturedMaxResults { get; set; }

            public override int CountMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
                => Count;

            public override SearchResult SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
            {
                CapturedMaxResults = args.MaxResults;
                var page = Hits.Take(args.MaxResults).ToList();
                return new SearchResult { Messages = page, TotalMatches = Count, Truncated = Count > page.Count };
            }

            public override FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken))
            {
                CapturedExcel = args;
                return new FileSavedResult
                {
                    Path = @"C:\Users\x\AppData\Local\OutlookAI\Reports\vendor.xlsx",
                    FileUrl = "file:///C:/Users/x/AppData/Local/OutlookAI/Reports/vendor.xlsx",
                    Format = "xlsx",
                    Bytes = 1234,
                    Filename = "vendor.xlsx",
                };
            }
        }

        private static List<MessageSummary> Make(int n)
        {
            var list = new List<MessageSummary>();
            for (var i = 0; i < n; i++)
                list.Add(new MessageSummary
                {
                    Id = "id" + i, Subject = "S" + i, From = "f" + i + "@x.com",
                    To = new[] { "me@x.com" }, ReceivedAt = DateTimeOffset.UtcNow.AddMinutes(i),
                    Snippet = "snip" + i, HasAttachments = false,
                });
            return list;
        }

        [Fact]
        public void Execute_UnderCeiling_ExportsAllAndReportsNotTruncated()
        {
            var surface = new FakeSurface { Count = 12, Hits = Make(12) };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync(
                "{\"from\":\"IT Creations\",\"columns\":[\"subject\",\"from\",\"snippet\"],\"filename_hint\":\"vendor\"}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal("file_saved", (string)obj["result_type"]);
            Assert.Equal(12, (int)obj["exported"]);
            Assert.Equal(12, (int)obj["total_matches"]);
            Assert.False((bool)obj["truncated"]);
            Assert.Equal("vendor.xlsx", (string)obj["filename"]);
            // 3 requested columns => 3 Excel columns, 12 rows
            Assert.Equal(3, surface.CapturedExcel.Columns.Count);
            Assert.Equal(12, surface.CapturedExcel.Rows.Count);
        }

        [Fact]
        public void Execute_OverCeiling_CapsAndReportsTruncated()
        {
            var surface = new FakeSurface { Count = 5000, Hits = Make(5000) };
            var tool = new OutlookExportSearchResultsTool(maxRowsOverride: 2000);

            var json = tool.ExecuteAsync(
                "{\"from\":\"x\",\"columns\":[\"subject\"]}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal(2000, (int)obj["exported"]);
            Assert.Equal(5000, (int)obj["total_matches"]);
            Assert.True((bool)obj["truncated"]);
            Assert.Equal(2000, surface.CapturedMaxResults);   // search capped at ceiling
        }

        [Fact]
        public void Execute_ZeroMatches_ReturnsNoMatchesWithoutWritingFile()
        {
            var surface = new FakeSurface { Count = 0, Hits = Make(0) };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync(
                "{\"from\":\"nobody\",\"columns\":[\"subject\"]}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal("no_matches", (string)obj["result_type"]);
            Assert.Null(surface.CapturedExcel);   // ExportExcel never called
        }
    }
}
```

- [ ] **Step 2: Build test project to verify RED**

Expect `CS0246` for `OutlookExportSearchResultsTool`.

- [ ] **Step 3: Create the tool**

```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// outlook_export_search_results. Deterministic, complete (up to a
    /// configurable ceiling) mechanical Excel export of a search. Counts the
    /// true total, collects up to the ceiling via the fast Table-API search
    /// path (no per-message body reads), projects the requested mechanical
    /// columns, and reports "exported N of M". For AI-synthesized exports
    /// ("summarize each email") the model must still read + accumulate; this
    /// tool only emits raw projected fields.
    /// </summary>
    public sealed class OutlookExportSearchResultsTool : IOutlookTool
    {
        private readonly int? _maxRowsOverride;

        public OutlookExportSearchResultsTool() : this(null) { }

        // Test seam so the ceiling can be set without touching global Config.
        internal OutlookExportSearchResultsTool(int? maxRowsOverride)
        {
            _maxRowsOverride = maxRowsOverride;
        }

        public string Name => "outlook_export_search_results";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                var args = ExportSearchResultsArgsParser.Parse(argsJson);
                var ceiling = _maxRowsOverride ?? Config.MaxBulkExportRows;
                if (ceiling < 1) ceiling = 1;

                // True total (count-mode, bounded per folder).
                var total = surface.CountMessages(args.Filter, ct);
                if (total <= 0)
                {
                    return Task.FromResult(new JObject(
                        new JProperty("result_type", "no_matches"),
                        new JProperty("total_matches", 0),
                        new JProperty("message", "No messages matched the filter; nothing was exported."))
                        .ToString(Newtonsoft.Json.Formatting.None));
                }

                // Collect up to the ceiling via the normal search path.
                args.Filter.MaxResults = Math.Min(ceiling, total);
                var result = surface.SearchMessages(args.Filter, ct) ?? new SearchResult();
                var hits = result.Messages ?? new MessageSummary[0];

                var excelArgs = BuildExcelArgs(args, hits);
                var saved = surface.ExportExcel(excelArgs, ct);

                var exported = hits.Count;
                var truncated = exported < total;

                return Task.FromResult(new JObject(
                    new JProperty("result_type", "file_saved"),
                    new JProperty("path", saved.Path ?? ""),
                    new JProperty("file_url", saved.FileUrl ?? ""),
                    new JProperty("format", saved.Format ?? ""),
                    new JProperty("bytes", saved.Bytes),
                    new JProperty("filename", saved.Filename ?? ""),
                    new JProperty("exported", exported),
                    new JProperty("total_matches", total),
                    new JProperty("truncated", truncated))
                    .ToString(Newtonsoft.Json.Formatting.None));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(BuildError("cancelled", "Export cancelled by user."));
            }
            catch (ToolArgValidationException ex)
            {
                return Task.FromResult(BuildError(ex.Code, ex.Message));
            }
            catch (ExportException ex)
            {
                return Task.FromResult(BuildError(ex.Code, ex.Message));
            }
            catch (Exception ex)
            {
                return Task.FromResult(BuildError("export_failed", ex.Message));
            }
        }

        private static ExportExcelArgs BuildExcelArgs(ExportSearchResultsArgs args, IReadOnlyList<MessageSummary> hits)
        {
            var columns = new List<ExcelColumnSpec>();
            foreach (var c in args.Columns)
            {
                columns.Add(new ExcelColumnSpec { Name = HeaderFor(c), Type = TypeFor(c) });
            }

            var rows = new List<JToken[]>(hits.Count);
            foreach (var m in hits)
            {
                var cells = new JToken[args.Columns.Count];
                for (var i = 0; i < args.Columns.Count; i++)
                {
                    cells[i] = CellFor(args.Columns[i], m);
                }
                rows.Add(cells);
            }

            return new ExportExcelArgs
            {
                FilenameHint = args.FilenameHint,
                SheetName = Truncate(args.SheetName, 31),
                Columns = columns,
                Rows = rows,
            };
        }

        private static string HeaderFor(string col)
        {
            switch (col)
            {
                case "subject": return "Subject";
                case "from": return "From";
                case "to": return "To";
                case "received_at": return "Received";
                case "snippet": return "Snippet";
                case "has_attachments": return "Has Attachments";
                case "folder": return "Folder";
                default: return col;
            }
        }

        private static ExcelColumnType TypeFor(string col)
        {
            switch (col)
            {
                case "received_at": return ExcelColumnType.DateTime;
                case "has_attachments": return ExcelColumnType.Boolean;
                default: return ExcelColumnType.Text;
            }
        }

        private static JToken CellFor(string col, MessageSummary m)
        {
            switch (col)
            {
                case "subject": return m.Subject ?? "";
                case "from": return m.From ?? "";
                case "to": return string.Join("; ", m.To ?? new string[0]);
                case "received_at": return m.ReceivedAt.ToString("o");
                case "snippet": return m.Snippet ?? "";
                case "has_attachments": return m.HasAttachments;
                case "folder": return "";   // folder name not on MessageSummary; reserved
                default: return "";
            }
        }

        private static string Truncate(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Length <= max ? s : s.Substring(0, max);
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
```

Note on the `folder` column: `MessageSummary` does not carry a folder name, so the `folder` cell is reserved/empty for now. Keep it in the allow-list for forward-compat but it emits empty. (Documented; not a placeholder — intentional.)

- [ ] **Step 4: Register in csproj**

```xml
    <Compile Include="Services\Tools\OutlookExportSearchResultsTool.cs" />
```

- [ ] **Step 5: Build + run (GREEN)**

Full build; then `/TestCaseFilter:"FullyQualifiedName~OutlookExportSearchResultsToolTests"`. Expect 3 pass. Note the over-ceiling test asserts `CapturedMaxResults == 2000`, proving the search is capped at the ceiling.

- [ ] **Step 6: Commit**

```
git add VSTO2/OutlookAI/Services/Tools/OutlookExportSearchResultsTool.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportSearchResultsToolTests.cs
git commit -m "feat(export): add outlook_export_search_results bulk tool"
```

---

## Task 6: Register tool in host + schema + rewrite steering

**Files:**
- Modify: `VSTO2/OutlookAI/Services/OutlookToolHost.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`
- Test: `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`

- [ ] **Step 1: Write the failing schema tests**

Append to `ToolCatalogSchemaTests` in `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`:

```csharp
[Fact]
public void ExportSearchResults_Tool_IsRegistered()
{
    var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var tool = FindTool(tools, "outlook_export_search_results");
    Assert.NotNull(tool);
    var desc = (string)tool["description"];
    Assert.NotNull(desc);
    Assert.Contains("complete", desc, System.StringComparison.OrdinalIgnoreCase);
    // advertises a columns array
    Assert.NotNull(tool["parameters"]["properties"]["columns"]);
}

[Fact]
public void SearchMessages_Description_TeachesTruncatedEscalation()
{
    var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var search = FindTool(tools, "outlook_search_messages");
    var desc = (string)search["description"];
    Assert.NotNull(desc);
    Assert.Contains("truncated", desc, System.StringComparison.OrdinalIgnoreCase);
    Assert.Contains("outlook_export_search_results", desc, System.StringComparison.Ordinal);
}
```

- [ ] **Step 2: Build test project to verify RED**

Expect the two new tests fail (tool not found / strings absent).

- [ ] **Step 3: Register the tool in the host**

In `VSTO2/OutlookAI/Services/OutlookToolHost.cs`, in the `tools` list initializer (after `new OutlookExportPdfTool(),` line 33) add:

```csharp
                new OutlookExportSearchResultsTool(),    // v2.1.2: complete bulk list export
```

Update the XML doc comment count if it states a specific number ("aggregates the 10 Phase 2 ...").

- [ ] **Step 4: Add the schema entry**

In `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`, after the `outlook_export_pdf` `BuildToolEntry(...)` block (find it; it's the one whose description mentions `content_markdown` / `250000`), add a new entry:

```csharp
                BuildToolEntry("outlook_export_search_results",
                    "Export a COMPLETE list of messages matching a search filter to a styled Excel .xlsx, up to a server-configured ceiling (default 2000 rows). "
                    + "Use this whenever the user wants 'all', 'every', 'a list of', or 'a spreadsheet of' messages matching criteria and completeness matters - e.g. 'Excel of every email from IT Creations', 'list all messages I sent to Susan in 2025'. "
                    + "Unlike calling outlook_search_messages then outlook_export_excel (which caps at 100 results and silently drops the rest), this tool counts the true total and collects up to the ceiling server-side in one call, then reports how many it exported vs the true total. "
                    + "Pass the SAME filter fields as outlook_search_messages (from, to, query, subject_contains, body_contains, scope, date_from, date_to, etc.) plus a 'columns' array choosing from: subject, from, to, received_at, snippet, has_attachments, folder. Columns default to received_at/from/to/subject/snippet when omitted. "
                    + "This tool projects only those mechanical fields (it does NOT read full bodies and cannot summarize or synthesize). For a report that needs per-message AI summaries, instead page outlook_search_messages by date window, read up to 25 bodies per window, accumulate, and call outlook_export_pdf once. "
                    + "If the result says truncated:true, tell the user you exported N of M and suggest narrowing the date range to capture the rest.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("properties", new JObject(
                            new JProperty("query",            new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Free-form keyword(s) matched against subject + body. Use dedicated fields for sender/recipient/dates."))),
                            new JProperty("from",             new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Sender substring; matches display name OR email (case-insensitive)."))),
                            new JProperty("to",               new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Recipient substring."))),
                            new JProperty("subject_contains", new JObject(new JProperty("type","string"))),
                            new JProperty("body_contains",    new JObject(new JProperty("type","string"))),
                            new JProperty("scope",            new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("current_folder","all_mail","auto")),
                                                              new JProperty("description","Default auto. Use all_mail for 'ever'/'any'/'everything'. folder_id overrides scope."))),
                            new JProperty("folder_id",        new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Use outlook_list_folders to discover ids."))),
                            new JProperty("date_from",        new JObject(new JProperty("type","string"),
                                                              new JProperty("format","date-time"),
                                                              new JProperty("description","Inclusive lower bound, ISO-8601 UTC."))),
                            new JProperty("date_to",          new JObject(new JProperty("type","string"),
                                                              new JProperty("format","date-time"),
                                                              new JProperty("description","Exclusive upper bound, ISO-8601 UTC."))),
                            new JProperty("columns",          new JObject(new JProperty("type","array"),
                                                              new JProperty("items", new JObject(new JProperty("type","string"),
                                                                  new JProperty("enum", new JArray("subject","from","to","received_at","snippet","has_attachments","folder")))),
                                                              new JProperty("description","Mechanical columns to include. Defaults to received_at, from, to, subject, snippet."))),
                            new JProperty("filename_hint",    new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Optional base filename, e.g. 'IT-Creations-quotes'."))))),
                        new JProperty("additionalProperties", false))),
```

- [ ] **Step 5: Rewrite the `outlook_search_messages` truncation steering**

In the existing `outlook_search_messages` description string (lines 34-57), make two edits:

(a) Remove the self-contradicting clause on line 53 (`"Do not use max_results:100 on broad all_mail targeted lookups; use 25 unless the user explicitly asks for a large row count or a specific folder is selected. "`). Replace it with:

```csharp
                    + "The result includes total_matches and truncated. If truncated is true you did NOT get all matches - the array is capped at max_results (hard cap 100). When the user wants a COMPLETE list or spreadsheet, do NOT try to raise max_results past 100; instead call outlook_export_search_results, which collects the full set server-side. "
```

(b) Keep the existing final sentence about date-window paging (lines 56-57) for the synthesis case — the test `SearchMessages_Description_TeachesDateWindowedPagination` (asserts "page by date window") and `SearchMessages_Description_SteersTowardSnippetForBulkExports` (asserts "snippet" + "metadata-only") must still pass, so do not remove the words "page by date window", "snippet", or "metadata-only".

- [ ] **Step 6: Build + run (GREEN)**

Full build. Run `/TestCaseFilter:"FullyQualifiedName~ToolCatalogSchemaTests"` — all pass including the two new + the pre-existing date-window/snippet ones. Then run the full suite.

- [ ] **Step 7: Commit**

```
git add VSTO2/OutlookAI/Services/OutlookToolHost.cs VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs
git commit -m "feat(tools): register outlook_export_search_results + truncation steering"
```

---

## Task 7: Verification gate

**Files:** none (verification only).

- [ ] **Step 1: WebUI syntax check (unchanged files, sanity only)**

```
node --check VSTO2\OutlookAI\WebUI\chat.js
node --check VSTO2\OutlookAI\WebUI\markdown.js
```
Expected: exit 0, no output.

- [ ] **Step 2: Full Debug build**

```
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```
Expected: success, only the pre-existing `MSB3277` warning.

- [ ] **Step 3: Full VSTest suite**

```
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```
Expected: `Failed: 0`. New total ≈ 607 + 4 (Config) + 2 (projector) + 1 (search tool) + 5 (export args parser) + 3 (export tool) + 2 (schema) = **624**. Exact count may vary if pre-existing projector tests were edited rather than added; the binding requirement is **0 failures** and every new test present.

- [ ] **Step 4: No stray TODO/FIXME in new/changed files**

```
Select-String -Path "VSTO2\OutlookAI\Services\Tools\SearchResult.cs",
                    "VSTO2\OutlookAI\Services\Tools\ExportSearchResultsArgs.cs",
                    "VSTO2\OutlookAI\Services\Tools\ExportSearchResultsArgsParser.cs",
                    "VSTO2\OutlookAI\Services\Tools\OutlookExportSearchResultsTool.cs" `
              -Pattern "TODO|FIXME"
```
Expected: empty.

- [ ] **Step 5: Working tree clean**

```
git status --short
```
Expected: empty (all committed).

---

## Notes for release (post-merge, separate from this plan)

- Ships as **v2.1.2**. CI release workflow is still blocked (issue #9), so build locally with `Deploy\Make-ReleaseZip.ps1 -Tag v2.1.2` and publish with `gh release create v2.1.2 ...` (same flow as v2.1.0/v2.1.1; see `handoff.md` §7).
- Smoke on the dev box or RDS: ask for "an Excel of every email from `<frequent sender>`" and confirm the row count equals `outlook_count_messages` for the same filter (up to the 2,000 ceiling), and that the assistant reports "exported N of M" when over the ceiling.
