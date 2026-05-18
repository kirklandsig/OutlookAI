# Phase 4: Inbox Reports Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship the new Inbox Reports task pane (a parallel chat-style surface with 6 templated report chips, a reports-focused system prompt, and two new tools — `outlook_read_messages` and `outlook_aggregate_messages`) on top of the Phase 3b non-blocking search engine.

**Architecture:** New `InboxReportsPane` mirrors `InboxCopilotPane`, sharing all infrastructure (WebView2, ChatService, marshaller, AdvancedSearchRunner) but owning its own `ConversationStore`, `OutlookToolHost`, and prompt. Two new tools join the catalog: bulk-read by ID array, and group-and-count with top_n. Same threading + cancellation patterns as Phase 3b.

**Tech Stack:** C# .NET Framework 4.7.2 (VSTO), xUnit, Outlook OOM (`Microsoft.Office.Interop.Outlook`), WebView2 for the chat UI.

**Spec:** `docs/superpowers/specs/2026-05-18-phase-4-inbox-reports-design.md`

---

## File Structure

**New files:**

- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.cs` — task pane wrapper, parallel to `InboxCopilotPane.cs`.
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.Designer.cs` — designer for the WinForms host.
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.resx` — designer resource file.
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsController.cs` — WebView message handler + chat session, parallel to `InboxCopilotController.cs`.
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPromptBuilder.cs` — reports-focused system prompt builder.
- `VSTO2/OutlookAI/TaskPane/InboxReports/ReportQuickActionChip.cs` — POCO + `Defaults()` static.
- `VSTO2/OutlookAI/Services/Tools/SenderKeyNormalizer.cs` — pure helper.
- `VSTO2/OutlookAI/Services/Tools/DateBucketFormatter.cs` — pure helper.
- `VSTO2/OutlookAI/Services/Tools/TopNBucketSelector.cs` — pure helper.
- `VSTO2/OutlookAI/Services/Tools/AggregationBucket.cs` — DTO returned by aggregate tool.
- `VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgs.cs` — DTO for aggregate tool args.
- `VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgsParser.cs` — pure JSON parser.
- `VSTO2/OutlookAI/Services/Tools/OutlookReadMessagesTool.cs` — `IOutlookTool` for bulk read.
- `VSTO2/OutlookAI/Services/Tools/OutlookAggregateMessagesTool.cs` — `IOutlookTool` for aggregate.
- Matching test files under `VSTO2/OutlookAI.Tests/`.

**Modified files:**

- `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs` — add `ReadMessages` and `AggregateMessages` signatures, plus `AggregateMessagesArgs` / `AggregationBucket` references.
- `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs` — implement both new surface methods.
- `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs` — register schemas for both new tools.
- `VSTO2/OutlookAI/Services/OutlookToolHost.cs` — register both new tools in the catalog.
- `VSTO2/OutlookAI/TaskPane/AITaskPane.cs` — update `NullSurface` for the new signatures.
- `VSTO2/OutlookAI/Ribbon.xml` — add the Reports ribbon button under TabMail.
- `VSTO2/OutlookAI/Ribbon.cs` — add `OnReportsClick` handler.
- `VSTO2/OutlookAI/ThisAddIn.cs` — add `ShowReportsTaskPane(Explorer)` plus toggle/show logic mirroring `ShowExplorerTaskPane`.
- `VSTO2/OutlookAI/OutlookAI.csproj` — register all new compile items.
- `VSTO2/OutlookAI/WebUI/chat.js` — render a chip row when the reports payload is sent during ready; click handler fills the input.
- `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs` — add virtual overrides for the new surface methods.

**Build & test commands** (reuse throughout this plan):

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Run from `C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-codex-oauth-migration`.

---

## Task 1: Pure helper — `SenderKeyNormalizer`

Buckets keyed by sender should not duplicate "Jane Doe" and "jane.doe@example.com" if they refer to the same person from the user's mailbox. We standardize: prefer the display name if present, fall back to the email address, trim, lowercase only the email part.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/SenderKeyNormalizer.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/SenderKeyNormalizerTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/SenderKeyNormalizerTests.cs`:

```csharp
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SenderKeyNormalizerTests
    {
        [Theory]
        [InlineData("Jane Doe", "jane@example.com", "Jane Doe")]
        [InlineData("Jane Doe", null,               "Jane Doe")]
        [InlineData("Jane Doe", "",                 "Jane Doe")]
        public void NameWins_WhenPresent(string name, string email, string expected)
        {
            Assert.Equal(expected, SenderKeyNormalizer.Normalize(name, email));
        }

        [Theory]
        [InlineData("", "Jane@Example.com", "jane@example.com")]
        [InlineData(null, "Bob@Example.com", "bob@example.com")]
        [InlineData("   ", "Carol@example.com", "carol@example.com")]
        public void EmailLowercased_WhenNameMissing(string name, string email, string expected)
        {
            Assert.Equal(expected, SenderKeyNormalizer.Normalize(name, email));
        }

        [Theory]
        [InlineData("  Jane Doe  ", "jane@example.com", "Jane Doe")]
        public void NameTrimmed(string name, string email, string expected)
        {
            Assert.Equal(expected, SenderKeyNormalizer.Normalize(name, email));
        }

        [Fact]
        public void BothMissing_ReturnsUnknownSender()
        {
            Assert.Equal("(unknown sender)", SenderKeyNormalizer.Normalize(null, null));
            Assert.Equal("(unknown sender)", SenderKeyNormalizer.Normalize("", ""));
            Assert.Equal("(unknown sender)", SenderKeyNormalizer.Normalize("   ", "   "));
        }
    }
}
```

- [ ] **Step 2: Run the tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SenderKeyNormalizerTests"
```

Expected: build error — `SenderKeyNormalizer` does not exist.

- [ ] **Step 3: Implement `SenderKeyNormalizer`**

Create `VSTO2/OutlookAI/Services/Tools/SenderKeyNormalizer.cs`:

```csharp
namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Produces a stable bucket key for messages grouped by sender.
    /// Prefers the display name; falls back to a lowercased email so
    /// case differences do not split the same correspondent into two
    /// buckets. Returns a sentinel string when neither is available so
    /// downstream code never has to handle null bucket keys.
    /// </summary>
    public static class SenderKeyNormalizer
    {
        public const string UnknownSender = "(unknown sender)";

        public static string Normalize(string senderName, string senderEmail)
        {
            var trimmedName = (senderName ?? "").Trim();
            if (trimmedName.Length > 0) return trimmedName;
            var trimmedEmail = (senderEmail ?? "").Trim();
            if (trimmedEmail.Length > 0) return trimmedEmail.ToLowerInvariant();
            return UnknownSender;
        }
    }
}
```

- [ ] **Step 4: Register the new file in the csproj**

Open `VSTO2/OutlookAI/OutlookAI.csproj`. Find the `<Compile Include="Services\Tools\SearchScopeFormatter.cs" />` line. Immediately after it, add:

```xml
    <Compile Include="Services\Tools\SenderKeyNormalizer.cs" />
```

(The csproj uses alphabetical ordering for Services\Tools entries; this keeps that.)

- [ ] **Step 5: Build and run the targeted tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SenderKeyNormalizerTests"
```

Expected: all `SenderKeyNormalizerTests` pass.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/SenderKeyNormalizer.cs VSTO2/OutlookAI.Tests/Services/Tools/SenderKeyNormalizerTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(reports): add SenderKeyNormalizer for aggregate bucket keys" -m "Pure helper that produces a stable bucket key for messages grouped by sender. Prefers display name, falls back to lowercased email (so 'Jane@Example.com' and 'jane@example.com' do not produce two buckets), returns '(unknown sender)' when neither is available."
```

---

## Task 2: Pure helper — `DateBucketFormatter`

Buckets keyed by day use ISO date format (`yyyy-MM-dd`) in UTC so the same calendar day across timezones produces one bucket per Outlook's local-to-UTC view of the message.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/DateBucketFormatter.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/DateBucketFormatterTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/DateBucketFormatterTests.cs`:

```csharp
using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class DateBucketFormatterTests
    {
        [Fact]
        public void Format_UtcDateTime_ReturnsIsoDateOnly()
        {
            var dt = new DateTimeOffset(2026, 5, 14, 18, 32, 0, TimeSpan.Zero);
            Assert.Equal("2026-05-14", DateBucketFormatter.Format(dt));
        }

        [Fact]
        public void Format_NonUtcOffset_NormalizedToUtcBeforeDateExtract()
        {
            // 2026-05-14 02:00 +04:00 == 2026-05-13 22:00 UTC. Bucket should be 2026-05-13.
            var dt = new DateTimeOffset(2026, 5, 14, 2, 0, 0, TimeSpan.FromHours(4));
            Assert.Equal("2026-05-13", DateBucketFormatter.Format(dt));
        }

        [Fact]
        public void Format_MinValue_ReturnsSentinel()
        {
            // Items we cannot date should not pollute buckets; sentinel
            // keeps them visible without colliding with real days.
            Assert.Equal("(unknown date)", DateBucketFormatter.Format(DateTimeOffset.MinValue));
        }
    }
}
```

- [ ] **Step 2: Run tests, verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~DateBucketFormatterTests"
```

Expected: build error — `DateBucketFormatter` does not exist.

- [ ] **Step 3: Implement `DateBucketFormatter`**

Create `VSTO2/OutlookAI/Services/Tools/DateBucketFormatter.cs`:

```csharp
using System;
using System.Globalization;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Produces a stable bucket key for messages grouped by calendar day
    /// (UTC). Outlook's ReceivedTime arrives as a local DateTime which
    /// our ToOffset wrapper turns into a DateTimeOffset; this helper
    /// normalizes to UTC, takes the date component, and formats as
    /// ISO-8601 (yyyy-MM-dd) so buckets sort lexically.
    /// </summary>
    public static class DateBucketFormatter
    {
        public const string UnknownDate = "(unknown date)";

        public static string Format(DateTimeOffset receivedAt)
        {
            if (receivedAt == DateTimeOffset.MinValue) return UnknownDate;
            return receivedAt.UtcDateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
    }
}
```

- [ ] **Step 4: Register in csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, find the `<Compile Include="Services\Tools\SenderKeyNormalizer.cs" />` line you added in Task 1. Add immediately before it:

```xml
    <Compile Include="Services\Tools\DateBucketFormatter.cs" />
```

- [ ] **Step 5: Build and run the targeted tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~DateBucketFormatterTests"
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/DateBucketFormatter.cs VSTO2/OutlookAI.Tests/Services/Tools/DateBucketFormatterTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(reports): add DateBucketFormatter for day-grouped aggregations" -m "Pure helper that turns a DateTimeOffset into a stable UTC ISO-8601 calendar-day key. Returns '(unknown date)' sentinel for MinValue so undateable items remain visible without colliding with real days."
```

---

## Task 3: Pure helper — `TopNBucketSelector`

After grouping, the aggregate tool returns the top-N buckets by count, sorted descending, ties broken alphabetically by label so output is deterministic across runs.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/TopNBucketSelector.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/TopNBucketSelectorTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/TopNBucketSelectorTests.cs`:

```csharp
using System.Collections.Generic;
using System.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class TopNBucketSelectorTests
    {
        private static AggregationBucket B(string label, int count) =>
            new AggregationBucket { Label = label, Count = count };

        [Fact]
        public void TakeTop_OrdersByCountDescending()
        {
            var input = new[] { B("a", 3), B("b", 10), B("c", 1) };
            var result = TopNBucketSelector.TakeTop(input, 5);
            Assert.Equal(new[] { "b", "a", "c" }, result.Select(r => r.Label).ToArray());
        }

        [Fact]
        public void TakeTop_ClampsToN()
        {
            var input = new[] { B("a", 5), B("b", 4), B("c", 3), B("d", 2) };
            var result = TopNBucketSelector.TakeTop(input, 2);
            Assert.Equal(2, result.Count);
            Assert.Equal(new[] { "a", "b" }, result.Select(r => r.Label).ToArray());
        }

        [Fact]
        public void TakeTop_TiesBreakAlphabetically()
        {
            var input = new[] { B("zeta", 5), B("alpha", 5), B("mu", 5) };
            var result = TopNBucketSelector.TakeTop(input, 3);
            Assert.Equal(new[] { "alpha", "mu", "zeta" }, result.Select(r => r.Label).ToArray());
        }

        [Fact]
        public void TakeTop_NullOrEmptyInput_ReturnsEmpty()
        {
            Assert.Empty(TopNBucketSelector.TakeTop(null, 5));
            Assert.Empty(TopNBucketSelector.TakeTop(new AggregationBucket[0], 5));
        }

        [Fact]
        public void TakeTop_NonPositiveN_ReturnsEmpty()
        {
            var input = new[] { B("a", 5) };
            Assert.Empty(TopNBucketSelector.TakeTop(input, 0));
            Assert.Empty(TopNBucketSelector.TakeTop(input, -3));
        }
    }
}
```

- [ ] **Step 2: Run tests, verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~TopNBucketSelectorTests"
```

Expected: build errors — `TopNBucketSelector` and `AggregationBucket` do not exist.

- [ ] **Step 3: Implement `AggregationBucket` first (the DTO used by Top-N)**

Create `VSTO2/OutlookAI/Services/Tools/AggregationBucket.cs`:

```csharp
namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// One bucket in an aggregate-messages result. Label is the grouped
    /// key (sender name, "yyyy-MM-dd" date, or folder name depending on
    /// args.GroupBy). Count is the number of matching messages.
    /// </summary>
    public sealed class AggregationBucket
    {
        public string Label { get; set; }
        public int Count { get; set; }
    }
}
```

- [ ] **Step 4: Implement `TopNBucketSelector`**

Create `VSTO2/OutlookAI/Services/Tools/TopNBucketSelector.cs`:

```csharp
using System.Collections.Generic;
using System.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Sorts AggregationBucket records by count descending (alphabetical
    /// label tiebreak for deterministic output across runs) and returns
    /// the first N. Defensive against null / empty / non-positive N.
    /// </summary>
    public static class TopNBucketSelector
    {
        public static IReadOnlyList<AggregationBucket> TakeTop(
            IEnumerable<AggregationBucket> buckets, int n)
        {
            if (buckets == null || n <= 0) return new AggregationBucket[0];
            return buckets
                .OrderByDescending(b => b.Count)
                .ThenBy(b => b.Label, System.StringComparer.OrdinalIgnoreCase)
                .Take(n)
                .ToList();
        }
    }
}
```

- [ ] **Step 5: Register both new files in the csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, find the existing `<Compile Include="Services\Tools\` block. Add both alphabetically:

```xml
    <Compile Include="Services\Tools\AggregationBucket.cs" />
```

immediately after `<Compile Include="Services\Tools\AdvancedSearch...` block or before `<Compile Include="Services\Tools\BuildRestrictFilter...` — wherever `A` belongs alphabetically. And:

```xml
    <Compile Include="Services\Tools\TopNBucketSelector.cs" />
```

between the existing `Services\Tools\T*` entries (it sits after `TableMessageClassFilter`).

- [ ] **Step 6: Build and run the targeted tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~TopNBucketSelectorTests"
```

Expected: all five `TopNBucketSelectorTests` pass.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/AggregationBucket.cs VSTO2/OutlookAI/Services/Tools/TopNBucketSelector.cs VSTO2/OutlookAI.Tests/Services/Tools/TopNBucketSelectorTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(reports): add AggregationBucket DTO and TopNBucketSelector" -m "Pure helper that sorts AggregationBucket records by count desc with alphabetical-label tiebreak for deterministic output, then clamps to top N. Defensive against null / empty / non-positive N. AggregationBucket is the DTO returned by the upcoming outlook_aggregate_messages tool."
```

---

## Task 4: `AggregateMessagesArgs` DTO and `AggregateMessagesArgsParser`

Defines the args for the aggregate tool and the JSON parser that builds it from `argsJson`. The parser mirrors `SearchMessagesArgsParser`'s pattern (defensive defaults, enum normalization, hard cap on top_n).

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgs.cs`
- Create: `VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgsParser.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/AggregateMessagesArgsParserTests.cs`

- [ ] **Step 1: Implement `AggregateMessagesArgs` DTO first (the parser tests will reference it)**

Create `VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgs.cs`:

```csharp
using System;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Arguments accepted by outlook_aggregate_messages. Mirrors
    /// SearchMessagesArgs for the filter portion, plus group_by + top_n.
    /// </summary>
    public sealed class AggregateMessagesArgs
    {
        public string Scope { get; set; } = "auto";    // current_folder | all_mail | auto
        public string FolderId { get; set; }            // optional explicit folder
        public DateTimeOffset? DateFrom { get; set; }
        public DateTimeOffset? DateTo { get; set; }
        public string From { get; set; }
        public string SubjectContains { get; set; }
        public string BodyContains { get; set; }
        public string GroupBy { get; set; } = "sender"; // sender | day | folder
        public int TopN { get; set; } = 10;
    }
}
```

- [ ] **Step 2: Register the DTO file in the csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, add (alphabetically with the other `Services\Tools\A*` and just-added `AggregationBucket.cs`):

```xml
    <Compile Include="Services\Tools\AggregateMessagesArgs.cs" />
```

Confirm the build is clean:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

- [ ] **Step 3: Write the failing parser tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/AggregateMessagesArgsParserTests.cs`:

```csharp
using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class AggregateMessagesArgsParserTests
    {
        [Fact]
        public void Parse_EmptyJson_AppliesDefaults()
        {
            var args = AggregateMessagesArgsParser.Parse("{}");
            Assert.Equal("auto", args.Scope);
            Assert.Equal("sender", args.GroupBy);
            Assert.Equal(10, args.TopN);
            Assert.Null(args.From);
            Assert.Null(args.SubjectContains);
            Assert.Null(args.BodyContains);
            Assert.Null(args.DateFrom);
            Assert.Null(args.DateTo);
        }

        [Fact]
        public void Parse_AllFields_RoundTrips()
        {
            var json = "{"
                + "\"scope\":\"all_mail\","
                + "\"folder_id\":\"f1\","
                + "\"date_from\":\"2026-05-01T00:00:00Z\","
                + "\"date_to\":\"2026-05-31T00:00:00Z\","
                + "\"from\":\"jane@example.com\","
                + "\"subject_contains\":\"Q4\","
                + "\"body_contains\":\"draft\","
                + "\"group_by\":\"day\","
                + "\"top_n\":25}";
            var args = AggregateMessagesArgsParser.Parse(json);
            Assert.Equal("all_mail", args.Scope);
            Assert.Equal("f1", args.FolderId);
            Assert.Equal(new DateTimeOffset(2026, 5, 1, 0, 0, 0, TimeSpan.Zero), args.DateFrom);
            Assert.Equal(new DateTimeOffset(2026, 5, 31, 0, 0, 0, TimeSpan.Zero), args.DateTo);
            Assert.Equal("jane@example.com", args.From);
            Assert.Equal("Q4", args.SubjectContains);
            Assert.Equal("draft", args.BodyContains);
            Assert.Equal("day", args.GroupBy);
            Assert.Equal(25, args.TopN);
        }

        [Theory]
        [InlineData("SENDER", "sender")]
        [InlineData("Day",    "day")]
        [InlineData("folder", "folder")]
        [InlineData("garbage","sender")]    // unknown → default "sender"
        [InlineData("",       "sender")]    // empty   → default
        public void Parse_GroupBy_NormalizedOrDefaulted(string raw, string expected)
        {
            var args = AggregateMessagesArgsParser.Parse("{\"group_by\":\"" + raw + "\"}");
            Assert.Equal(expected, args.GroupBy);
        }

        [Theory]
        [InlineData(0,     1)]               // floored to 1
        [InlineData(-5,    1)]               // floored to 1
        [InlineData(50,    50)]
        [InlineData(9999,  100)]             // capped at 100
        public void Parse_TopN_ClampedToValidRange(int input, int expected)
        {
            var args = AggregateMessagesArgsParser.Parse("{\"top_n\":" + input + "}");
            Assert.Equal(expected, args.TopN);
        }

        [Theory]
        [InlineData("ALL_MAIL", "all_mail")]
        [InlineData("Current_Folder", "current_folder")]
        [InlineData("Auto", "auto")]
        [InlineData("garbage", "auto")]
        public void Parse_Scope_NormalizedOrDefaulted(string raw, string expected)
        {
            var args = AggregateMessagesArgsParser.Parse("{\"scope\":\"" + raw + "\"}");
            Assert.Equal(expected, args.Scope);
        }

        [Fact]
        public void Parse_NullOrWhitespaceJson_AppliesDefaults()
        {
            Assert.Equal("auto", AggregateMessagesArgsParser.Parse(null).Scope);
            Assert.Equal("auto", AggregateMessagesArgsParser.Parse("").Scope);
            Assert.Equal("auto", AggregateMessagesArgsParser.Parse("   ").Scope);
        }
    }
}
```

- [ ] **Step 4: Run tests, verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~AggregateMessagesArgsParserTests"
```

Expected: build error — `AggregateMessagesArgsParser` does not exist.

- [ ] **Step 5: Implement `AggregateMessagesArgsParser`**

Create `VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgsParser.cs`:

```csharp
using System;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    internal static class AggregateMessagesArgsParser
    {
        private const int TopNFloor = 1;
        private const int TopNCap = 100;

        public static AggregateMessagesArgs Parse(string argsJson)
        {
            var args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);
            return new AggregateMessagesArgs
            {
                Scope = EnumOrDefault(args["scope"], "auto", "auto", "current_folder", "all_mail"),
                FolderId = Clean(args["folder_id"]),
                DateFrom = ParseDate(args["date_from"]),
                DateTo = ParseDate(args["date_to"]),
                From = Clean(args["from"]),
                SubjectContains = Clean(args["subject_contains"]),
                BodyContains = Clean(args["body_contains"]),
                GroupBy = EnumOrDefault(args["group_by"], "sender", "sender", "day", "folder"),
                TopN = ClampTopN(args["top_n"]?.Value<int>() ?? 10),
            };
        }

        private static int ClampTopN(int value)
        {
            if (value < TopNFloor) return TopNFloor;
            if (value > TopNCap) return TopNCap;
            return value;
        }

        private static string Clean(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            var value = ((string)token)?.Trim();
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static string EnumOrDefault(JToken token, string fallback, params string[] allowed)
        {
            var value = Clean(token);
            if (value == null) return fallback;
            value = value.ToLowerInvariant();
            foreach (var allowedValue in allowed)
            {
                if (value == allowedValue) return value;
            }
            return fallback;
        }

        private static DateTimeOffset? ParseDate(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            DateTimeOffset value;
            if (!DateTimeOffset.TryParse(
                (string)token,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out value))
            {
                return null;
            }
            return value;
        }
    }
}
```

- [ ] **Step 6: Register the parser in the csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, add (alphabetically with other `Services\Tools\A*`):

```xml
    <Compile Include="Services\Tools\AggregateMessagesArgsParser.cs" />
```

- [ ] **Step 7: Build and run the parser tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~AggregateMessagesArgsParserTests"
```

Expected: all parser tests pass.

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgs.cs VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgsParser.cs VSTO2/OutlookAI.Tests/Services/Tools/AggregateMessagesArgsParserTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(reports): add AggregateMessagesArgs + parser" -m "DTO + JSON parser for outlook_aggregate_messages. Defaults: scope=auto, group_by=sender, top_n=10 (floor 1, cap 100). Scope and group_by are normalized through enum-or-default like SearchMessagesArgsParser does. Empty/null/whitespace JSON applies defaults."
```

---

## Task 5: `InboxReportsPromptBuilder`

Builds the system prompt for the Reports pane. Tells the model to use markdown, prefer the bulk-read and aggregate tools, ask before unresolved `[placeholders]`, and keep reports concise.

**Files:**
- Create: `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPromptBuilder.cs`
- Create: `VSTO2/OutlookAI.Tests/TaskPane/InboxReports/InboxReportsPromptBuilderTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/TaskPane/InboxReports/InboxReportsPromptBuilderTests.cs`:

```csharp
using OutlookAI.TaskPane.InboxReports;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxReports
{
    public class InboxReportsPromptBuilderTests
    {
        private readonly string _prompt = new InboxReportsPromptBuilder().Build();

        [Fact]
        public void Prompt_AlwaysIncludesRolePreamble()
        {
            Assert.Contains("mailbox reports assistant", _prompt, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Prompt_MentionsBulkReadTool()
        {
            Assert.Contains("outlook_read_messages", _prompt);
        }

        [Fact]
        public void Prompt_MentionsAggregateTool()
        {
            Assert.Contains("outlook_aggregate_messages", _prompt);
        }

        [Fact]
        public void Prompt_TellsModelToAskBeforeUnresolvedPlaceholders()
        {
            // The chip templates include placeholders like "[name or email]";
            // the model must ask for clarification rather than calling tools
            // with literal brackets in the args.
            Assert.Contains("placeholder", _prompt, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("ask", _prompt, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Prompt_DescribesMarkdownAndConciseFormat()
        {
            Assert.Contains("markdown", _prompt, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("concise", _prompt, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Prompt_DescribesHeaderWithScopeAndCount()
        {
            // Each report should open with "Searched: <scope>, <date range>, N messages" style.
            Assert.Contains("header", _prompt, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("scope", _prompt, System.StringComparison.OrdinalIgnoreCase);
        }
    }
}
```

- [ ] **Step 2: Run tests, verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~InboxReportsPromptBuilderTests"
```

Expected: build errors — `InboxReportsPromptBuilder` does not exist.

- [ ] **Step 3: Implement `InboxReportsPromptBuilder`**

Create `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPromptBuilder.cs`:

```csharp
namespace OutlookAI.TaskPane.InboxReports
{
    /// <summary>
    /// Builds the system prompt for the Inbox Reports pane. Steers the
    /// model toward concise, structured markdown reports and tells it
    /// when to prefer the bulk-read and aggregate tools.
    /// </summary>
    public sealed class InboxReportsPromptBuilder
    {
        public string Build()
        {
            return
@"You are a mailbox reports assistant inside Microsoft Outlook. Your job
is to produce concise, well-structured reports of the user's email
content using the provided tools.

Always:
- Use markdown structure (headers, bullets, short paragraphs, tables
  when appropriate).
- Start every report with a one-line header indicating WHAT was
  searched (scope + date range) and HOW MANY messages were processed.
- For action items, topic status, and conversation summaries: prefer
  outlook_read_messages (bulk) over many outlook_read_message calls.
  Bulk read is 5-10x faster.
- For sender/day/folder counts (statistics): use
  outlook_aggregate_messages.
- For finding which messages to read first: outlook_search_messages
  with date_from / date_to and (optional) from / subject_contains.
- If a placeholder like ""[name or email]"" or ""[start date]"" is still
  in the user's prompt text, ASK for clarification before calling any
  tool.

Never:
- Dump raw JSON. The user wants a human-readable report.
- Include long quoted email bodies. Quote only when essential, max
  ~3 lines per quote.
- Apologize, preamble, or pad. Concise is better than complete.";
        }
    }
}
```

- [ ] **Step 4: Register the new file in the csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, find the existing `<Compile Include="TaskPane\InboxCopilot\` block. Add a new alphabetically-after block:

```xml
    <Compile Include="TaskPane\InboxReports\InboxReportsPromptBuilder.cs" />
```

- [ ] **Step 5: Build and run the tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~InboxReportsPromptBuilderTests"
```

Expected: all six `InboxReportsPromptBuilderTests` pass.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPromptBuilder.cs VSTO2/OutlookAI.Tests/TaskPane/InboxReports/InboxReportsPromptBuilderTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(reports): add InboxReportsPromptBuilder" -m "Builds the system prompt for the Inbox Reports pane. Steers the model toward concise markdown, opens-with-scope-and-count header convention, prefer outlook_read_messages over many read_message calls, prefer outlook_aggregate_messages for counts/groupings, and ask for clarification when [placeholders] are unresolved."
```

---

## Task 6: `ReportQuickActionChip` + `Defaults()`

Defines the 6 chips with labels, templates, and emoji prefixes.

**Files:**
- Create: `VSTO2/OutlookAI/TaskPane/InboxReports/ReportQuickActionChip.cs`
- Create: `VSTO2/OutlookAI.Tests/TaskPane/InboxReports/ReportQuickActionChipTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/TaskPane/InboxReports/ReportQuickActionChipTests.cs`:

```csharp
using System.Linq;
using OutlookAI.TaskPane.InboxReports;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxReports
{
    public class ReportQuickActionChipTests
    {
        private readonly System.Collections.Generic.IReadOnlyList<ReportQuickActionChip> _chips
            = ReportQuickActionChip.Defaults();

        [Fact]
        public void Defaults_ReturnsSixChips()
        {
            Assert.Equal(6, _chips.Count);
        }

        [Fact]
        public void Defaults_EachChipHasLabelAndTemplate()
        {
            foreach (var c in _chips)
            {
                Assert.False(string.IsNullOrWhiteSpace(c.Label));
                Assert.False(string.IsNullOrWhiteSpace(c.TemplateText));
                // Templates should be substantive prompts, not one-liners.
                Assert.True(c.TemplateText.Length > 40,
                    "Chip template too short: " + c.Label);
            }
        }

        [Fact]
        public void Defaults_OrderingMatchesSpec()
        {
            // Spec defines the order: Digest, Conversation, Action items,
            // Project status, Stats, Out-of-office.
            Assert.Contains("digest", _chips[0].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("conversation", _chips[1].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("action", _chips[2].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("project", _chips[3].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stats", _chips[4].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("out", _chips[5].Label, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ConversationChip_TemplateMentionsPersonPlaceholder()
        {
            Assert.Contains("[name or email]", _chips[1].TemplateText);
        }

        [Fact]
        public void ProjectChip_TemplateMentionsTopicPlaceholder()
        {
            Assert.Contains("[topic", _chips[3].TemplateText);
        }

        [Fact]
        public void OutOfOfficeChip_TemplateMentionsDatePlaceholders()
        {
            Assert.Contains("[start date]", _chips[5].TemplateText);
            Assert.Contains("[end date]", _chips[5].TemplateText);
        }

        [Fact]
        public void StatsChip_TemplateMentionsAggregateTool()
        {
            // The stats chip explicitly nudges the model toward the
            // aggregate tool. Keeps the tool routing honest even if the
            // system prompt is later trimmed.
            Assert.Contains("outlook_aggregate_messages", _chips[4].TemplateText);
        }

        [Fact]
        public void Defaults_LabelsAreUnique()
        {
            Assert.Equal(_chips.Count, _chips.Select(c => c.Label).Distinct().Count());
        }
    }
}
```

- [ ] **Step 2: Run tests, verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ReportQuickActionChipTests"
```

Expected: build error — `ReportQuickActionChip` does not exist.

- [ ] **Step 3: Implement `ReportQuickActionChip`**

Create `VSTO2/OutlookAI/TaskPane/InboxReports/ReportQuickActionChip.cs`:

```csharp
using System.Collections.Generic;

namespace OutlookAI.TaskPane.InboxReports
{
    /// <summary>
    /// One quick-action chip in the Inbox Reports pane. Clicking the chip
    /// prefills the chat input with TemplateText. The user can edit
    /// any [placeholders] before sending; the system prompt instructs
    /// the model to ask for clarification if a placeholder is still
    /// present at tool-call time.
    /// </summary>
    public sealed class ReportQuickActionChip
    {
        public string Label { get; set; }
        public string TemplateText { get; set; }

        public static IReadOnlyList<ReportQuickActionChip> Defaults()
        {
            return new[]
            {
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCC5 This week's digest",
                    TemplateText = "Summarize what came into my Inbox over the past 7 days. Group by sender or topic. Highlight urgent items and emails I'm directly addressed in.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCAC Conversation summary",
                    TemplateText = "Summarize my recent email conversations with [name or email]. Show the chronological flow and key decisions/topics.",
                },
                new ReportQuickActionChip {
                    Label = "\u2713 Action items",
                    TemplateText = "Find action items I need to do based on emails from the past 7 days. Read the relevant messages, extract TODOs/deadlines/asks. Group by who's waiting on what.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCC1 Project status",
                    TemplateText = "Summarize the status of [topic/project name]. Find relevant emails, read the most recent ones, and give me: latest update, open questions, action items, key participants.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCCA Email stats",
                    TemplateText = "Give me email statistics for the past 30 days: top 10 senders, busiest days, breakdown by folder. Use outlook_aggregate_messages.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83C\uDFD6\uFE0F While I was out",
                    TemplateText = "I was out from [start date] to [end date]. Show me what's important from that timeframe: urgent items, direct asks, replies needed. De-prioritize newsletters and automated mail.",
                },
            };
        }
    }
}
```

- [ ] **Step 4: Register the new file in the csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, alphabetically with the other `TaskPane\InboxReports\*`:

```xml
    <Compile Include="TaskPane\InboxReports\ReportQuickActionChip.cs" />
```

- [ ] **Step 5: Build and run the tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ReportQuickActionChipTests"
```

Expected: all chip tests pass.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/TaskPane/InboxReports/ReportQuickActionChip.cs VSTO2/OutlookAI.Tests/TaskPane/InboxReports/ReportQuickActionChipTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(reports): add ReportQuickActionChip with six default templates" -m "Six quick-action chips for the Inbox Reports pane: digest, conversation summary, action items, project status, stats, out-of-office catchup. Templates include [placeholders] for chips that need user-supplied parameters (person, topic, dates). Stats template explicitly nudges the model toward outlook_aggregate_messages even if the system prompt is later trimmed."
```

---

## Task 7: Extend `IOutlookSurface` + `MinimalSurface` + `NullSurface` with the two new methods

Add the surface contract for `ReadMessages` and `AggregateMessages`. Default implementations: `MinimalSurface` (test base) throws `NotImplementedException` to keep tests honest; `NullSurface` (AITaskPane fallback) returns empty results.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs`
- Modify: `VSTO2/OutlookAI/TaskPane/AITaskPane.cs` (NullSurface inner class)

- [ ] **Step 1: Add `ReadMessages` and `AggregateMessages` to `IOutlookSurface`**

In `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`, add these two methods to the interface (placement: alphabetically after `ReadMessage` and `CountMessages`, or wherever the existing pattern keeps them):

```csharp
IReadOnlyList<MessageDetail> ReadMessages(
    string[] ids,
    bool includeBody,
    int maxItems,
    CancellationToken ct = default(CancellationToken));

IReadOnlyList<AggregationBucket> AggregateMessages(
    AggregateMessagesArgs args,
    CancellationToken ct = default(CancellationToken));
```

The interface currently has `using System.Threading;` from Phase 3b; both signatures rely on that. `MessageDetail` is already defined in this file. `AggregationBucket` and `AggregateMessagesArgs` were added in Tasks 3 and 4.

- [ ] **Step 2: Update `MinimalSurface`**

Open `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs`. Add the two new virtual methods alongside the existing ones (keep the throwing-default pattern so tests that don't override fail loudly):

```csharp
public virtual IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken))
    => throw new System.NotImplementedException();
public virtual IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken))
    => throw new System.NotImplementedException();
```

- [ ] **Step 3: Update `AITaskPane.NullSurface`**

In `VSTO2/OutlookAI/TaskPane/AITaskPane.cs`, find the inner `NullSurface : IOutlookSurface` class (around line 595) and add two implementations that return empty collections:

```csharp
public IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => new MessageDetail[0];
public IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => new AggregationBucket[0];
```

- [ ] **Step 4: Verify the build catches all callers**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: ONE remaining build error — `LiveOutlookSurface` does not implement `ReadMessages` and `AggregateMessages`. That's intentional — we'll add stub implementations in Step 5 here that throw `NotImplementedException`, then replace them with real implementations in Tasks 10 and 11.

- [ ] **Step 5: Add throwing stubs to `LiveOutlookSurface`**

In `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`, add these two methods anywhere inside the class (near `CountMessages` is natural):

```csharp
public IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken))
{
    throw new System.NotImplementedException("LiveOutlookSurface.ReadMessages will be implemented in a follow-up task.");
}

public IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken))
{
    throw new System.NotImplementedException("LiveOutlookSurface.AggregateMessages will be implemented in a follow-up task.");
}
```

- [ ] **Step 6: Build and run the full test suite to confirm nothing regressed**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: clean build, all existing tests still pass. The new methods throw `NotImplementedException` if any test exercises them, but no existing test does.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs VSTO2/OutlookAI/TaskPane/AITaskPane.cs
git commit -m "feat(reports): add ReadMessages and AggregateMessages to IOutlookSurface" -m "Surface contract for the upcoming outlook_read_messages and outlook_aggregate_messages tools. MinimalSurface keeps the throw-on-call pattern so tests fail loud if they accidentally exercise surface methods they did not override. AITaskPane.NullSurface returns empty collections. LiveOutlookSurface adds NotImplementedException stubs to be replaced by real implementations in follow-up tasks."
```

---

## Task 8: `OutlookReadMessagesTool` + tests

The tool that the model calls when it needs bodies for many messages at once.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/OutlookReadMessagesTool.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookReadMessagesToolTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/OutlookReadMessagesToolTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookReadMessagesToolTests
    {
        [Fact]
        public async Task Execute_PassesIdsThrough()
        {
            string[] observed = null;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = ids;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\",\"b\",\"c\"]}", surface, CancellationToken.None);
            Assert.Equal(new[] { "a", "b", "c" }, observed);
        }

        [Fact]
        public async Task Execute_DefaultIncludeBody_IsTrue()
        {
            bool? observed = null;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = includeBody;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\"]}", surface, CancellationToken.None);
            Assert.True(observed);
        }

        [Fact]
        public async Task Execute_DefaultMaxItems_Is25()
        {
            int observed = -1;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = maxItems;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\"]}", surface, CancellationToken.None);
            Assert.Equal(25, observed);
        }

        [Fact]
        public async Task Execute_MaxItemsClampedTo100()
        {
            int observed = -1;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = maxItems;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\"],\"max_items\":9999}", surface, CancellationToken.None);
            Assert.Equal(100, observed);
        }

        [Fact]
        public async Task Execute_IncludeBodyFalse_Honored()
        {
            bool? observed = null;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = includeBody;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\"],\"include_body\":false}", surface, CancellationToken.None);
            Assert.False(observed);
        }

        [Fact]
        public async Task Execute_ProjectsMessageDetailsToJson()
        {
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => new[]
                {
                    new MessageDetail
                    {
                        Id = "m1",
                        Subject = "Q4 plan",
                        From = "jane@example.com",
                        To = new[] { "bob@example.com" },
                        Cc = new string[0],
                        ReceivedAt = DateTimeOffset.Parse("2026-05-14T18:32:00Z"),
                        BodyPlaintext = "Body text",
                        BodyTruncated = false,
                        Attachments = new AttachmentSummary[0],
                        InReplyToMessageId = null,
                        ConversationTopic = "Q4",
                    }
                }
            };
            var tool = new OutlookReadMessagesTool();
            var json = await tool.ExecuteAsync("{\"ids\":[\"m1\"]}", surface, CancellationToken.None);
            Assert.Contains("\"id\":\"m1\"", json);
            Assert.Contains("\"subject\":\"Q4 plan\"", json);
            Assert.Contains("\"body_plaintext\":\"Body text\"", json);
            Assert.Contains("\"body_truncated\":false", json);
            Assert.Contains("\"conversation_topic\":\"Q4\"", json);
        }

        [Fact]
        public async Task Execute_EmptyIds_ReturnsEmptyArrayWithoutCallingSurface()
        {
            bool called = false;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => { called = true; return new MessageDetail[0]; }
            };
            var tool = new OutlookReadMessagesTool();
            var json = await tool.ExecuteAsync("{\"ids\":[]}", surface, CancellationToken.None);
            Assert.False(called);
            Assert.Contains("\"messages\":[]", json);
        }

        [Fact]
        public async Task Execute_MissingIds_ReturnsEmptyArray()
        {
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => new MessageDetail[0]
            };
            var tool = new OutlookReadMessagesTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"messages\":[]", json);
        }

        [Fact]
        public async Task Execute_PassesCancellationTokenThroughToSurface()
        {
            CancellationToken observed = default(CancellationToken);
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => { observed = ct; return new MessageDetail[0]; }
            };
            var tool = new OutlookReadMessagesTool();
            using (var cts = new CancellationTokenSource())
            {
                await tool.ExecuteAsync("{\"ids\":[\"a\"]}", surface, cts.Token);
                Assert.Equal(cts.Token, observed);
            }
        }

        [Fact]
        public async Task Execute_OnSurfaceCancellation_EmitsStructuredCancelEnvelope()
        {
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => throw new OperationCanceledException(ct)
            };
            var tool = new OutlookReadMessagesTool();
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();
                var json = await tool.ExecuteAsync("{\"ids\":[\"a\"]}", surface, cts.Token);
                Assert.Contains("\"error\"", json);
                Assert.Contains("\"code\":\"cancelled\"", json);
            }
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<string[], bool, int, CancellationToken, IReadOnlyList<MessageDetail>> OnReadMessages { get; set; }
            public override IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken))
                => OnReadMessages(ids, includeBody, maxItems, ct);
        }
    }
}
```

- [ ] **Step 2: Run tests, verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookReadMessagesToolTests"
```

Expected: build error — `OutlookReadMessagesTool` does not exist.

- [ ] **Step 3: Implement `OutlookReadMessagesTool`**

Create `VSTO2/OutlookAI/Services/Tools/OutlookReadMessagesTool.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_read_messages. Bulk-read message details by short
    /// ID array. Used by reports that need bodies for many messages
    /// (action items, topic status, conversation summaries). Replaces
    /// many outlook_read_message round trips with one call.
    /// </summary>
    public sealed class OutlookReadMessagesTool : IOutlookTool
    {
        private const int DefaultMaxItems = 25;
        private const int MaxItemsCap = 100;

        public string Name => "outlook_read_messages";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                var args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);

                string[] ids = (args["ids"] as JArray)?
                    .Select(t => (string)t)
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToArray() ?? new string[0];

                bool includeBody = args["include_body"]?.Type == JTokenType.Boolean
                    ? args["include_body"].Value<bool>()
                    : true;

                int maxItems = args["max_items"]?.Value<int>() ?? DefaultMaxItems;
                if (maxItems < 1) maxItems = 1;
                if (maxItems > MaxItemsCap) maxItems = MaxItemsCap;

                IReadOnlyList<MessageDetail> messages;
                if (ids.Length == 0)
                {
                    messages = new MessageDetail[0];
                }
                else
                {
                    messages = surface.ReadMessages(ids, includeBody, maxItems, ct) ?? new MessageDetail[0];
                }

                var json = new JObject(
                    new JProperty("messages", new JArray(messages.Select(m =>
                        new JObject(
                            new JProperty("id", m.Id ?? ""),
                            new JProperty("subject", m.Subject ?? ""),
                            new JProperty("from", m.From ?? ""),
                            new JProperty("to", new JArray((m.To ?? new string[0]).Cast<object>())),
                            new JProperty("cc", new JArray((m.Cc ?? new string[0]).Cast<object>())),
                            new JProperty("received_at", m.ReceivedAt.ToString("o")),
                            new JProperty("body_plaintext", m.BodyPlaintext ?? ""),
                            new JProperty("body_truncated", m.BodyTruncated),
                            new JProperty("attachments", new JArray((m.Attachments ?? new AttachmentSummary[0]).Select(a =>
                                new JObject(
                                    new JProperty("filename", a.Filename ?? ""),
                                    new JProperty("size_bytes", a.SizeBytes))))),
                            new JProperty("in_reply_to_message_id", m.InReplyToMessageId ?? ""),
                            new JProperty("conversation_topic", m.ConversationTopic ?? ""))))));
                return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(BuildError("cancelled", "Read cancelled by user."));
            }
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

- [ ] **Step 4: Register the new file in the csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, alphabetically with the other `Services\Tools\Outlook*Tool.cs`:

```xml
    <Compile Include="Services\Tools\OutlookReadMessagesTool.cs" />
```

- [ ] **Step 5: Build and run the tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookReadMessagesToolTests"
```

Expected: all ten tests pass.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookReadMessagesTool.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookReadMessagesToolTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(reports): add OutlookReadMessagesTool" -m "Bulk-read variant of outlook_read_message. Accepts ids[], include_body (default true), max_items (default 25, hard cap 100). Surface skips entirely when ids is empty. Projects each MessageDetail to JSON including attachments. OperationCanceledException becomes a structured cancel envelope."
```

---

## Task 9: `OutlookAggregateMessagesTool` + tests

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/OutlookAggregateMessagesTool.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookAggregateMessagesToolTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/OutlookAggregateMessagesToolTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookAggregateMessagesToolTests
    {
        [Fact]
        public async Task Execute_ParsesAllFieldsAndPassesThrough()
        {
            AggregateMessagesArgs observed = null;
            var surface = new Surface
            {
                OnAggregate = (args, ct) => { observed = args; return new AggregationBucket[0]; }
            };
            var tool = new OutlookAggregateMessagesTool();
            var argsJson = "{"
                + "\"scope\":\"all_mail\","
                + "\"date_from\":\"2026-05-01T00:00:00Z\","
                + "\"date_to\":\"2026-05-31T00:00:00Z\","
                + "\"from\":\"jane@example.com\","
                + "\"group_by\":\"sender\","
                + "\"top_n\":25}";
            await tool.ExecuteAsync(argsJson, surface, CancellationToken.None);

            Assert.NotNull(observed);
            Assert.Equal("all_mail", observed.Scope);
            Assert.Equal(new DateTimeOffset(2026, 5, 1, 0, 0, 0, TimeSpan.Zero), observed.DateFrom);
            Assert.Equal(new DateTimeOffset(2026, 5, 31, 0, 0, 0, TimeSpan.Zero), observed.DateTo);
            Assert.Equal("jane@example.com", observed.From);
            Assert.Equal("sender", observed.GroupBy);
            Assert.Equal(25, observed.TopN);
        }

        [Fact]
        public async Task Execute_ProjectsBucketsAndTotal()
        {
            var surface = new Surface
            {
                OnAggregate = (args, ct) => new[]
                {
                    new AggregationBucket { Label = "Jane Doe", Count = 47 },
                    new AggregationBucket { Label = "Bob Smith", Count = 31 },
                }
            };
            var tool = new OutlookAggregateMessagesTool();
            var json = await tool.ExecuteAsync(
                "{\"scope\":\"all_mail\",\"group_by\":\"sender\"}",
                surface, CancellationToken.None);

            Assert.Contains("\"buckets\":[", json);
            Assert.Contains("\"label\":\"Jane Doe\"", json);
            Assert.Contains("\"count\":47", json);
            Assert.Contains("\"label\":\"Bob Smith\"", json);
            Assert.Contains("\"count\":31", json);
            // total = sum of bucket counts
            Assert.Contains("\"total\":78", json);
        }

        [Fact]
        public async Task Execute_EmptyResult_ReturnsZeroTotal()
        {
            var surface = new Surface { OnAggregate = (args, ct) => new AggregationBucket[0] };
            var tool = new OutlookAggregateMessagesTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"buckets\":[]", json);
            Assert.Contains("\"total\":0", json);
        }

        [Fact]
        public async Task Execute_PassesCancellationTokenThroughToSurface()
        {
            CancellationToken observed = default(CancellationToken);
            var surface = new Surface
            {
                OnAggregate = (args, ct) => { observed = ct; return new AggregationBucket[0]; }
            };
            var tool = new OutlookAggregateMessagesTool();
            using (var cts = new CancellationTokenSource())
            {
                await tool.ExecuteAsync("{}", surface, cts.Token);
                Assert.Equal(cts.Token, observed);
            }
        }

        [Fact]
        public async Task Execute_OnSurfaceCancellation_EmitsStructuredCancelEnvelope()
        {
            var surface = new Surface
            {
                OnAggregate = (args, ct) => throw new OperationCanceledException(ct)
            };
            var tool = new OutlookAggregateMessagesTool();
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();
                var json = await tool.ExecuteAsync("{}", surface, cts.Token);
                Assert.Contains("\"error\"", json);
                Assert.Contains("\"code\":\"cancelled\"", json);
            }
        }

        [Fact]
        public async Task Execute_TopNDefault10_PassedThrough()
        {
            int observed = -1;
            var surface = new Surface
            {
                OnAggregate = (args, ct) => { observed = args.TopN; return new AggregationBucket[0]; }
            };
            var tool = new OutlookAggregateMessagesTool();
            await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Equal(10, observed);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<AggregateMessagesArgs, CancellationToken, IReadOnlyList<AggregationBucket>> OnAggregate { get; set; }
            public override IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken))
                => OnAggregate(args, ct);
        }
    }
}
```

- [ ] **Step 2: Run tests, verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookAggregateMessagesToolTests"
```

Expected: build error — `OutlookAggregateMessagesTool` does not exist.

- [ ] **Step 3: Implement `OutlookAggregateMessagesTool`**

Create `VSTO2/OutlookAI/Services/Tools/OutlookAggregateMessagesTool.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_aggregate_messages. Group messages by sender, day,
    /// or folder and return the top-N buckets by count. Used by stats
    /// and out-of-office-catchup reports.
    /// </summary>
    public sealed class OutlookAggregateMessagesTool : IOutlookTool
    {
        public string Name => "outlook_aggregate_messages";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                var args = AggregateMessagesArgsParser.Parse(argsJson);
                var buckets = surface.AggregateMessages(args, ct) ?? new AggregationBucket[0];

                var total = buckets.Sum(b => b.Count);

                var json = new JObject(
                    new JProperty("buckets", new JArray(buckets.Select(b =>
                        new JObject(
                            new JProperty("label", b.Label ?? ""),
                            new JProperty("count", b.Count))))),
                    new JProperty("total", total));
                return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(BuildError("cancelled", "Aggregation cancelled by user."));
            }
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

- [ ] **Step 4: Register in csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, alphabetically with the other `Services\Tools\Outlook*Tool.cs`:

```xml
    <Compile Include="Services\Tools\OutlookAggregateMessagesTool.cs" />
```

- [ ] **Step 5: Build and run the tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookAggregateMessagesToolTests"
```

Expected: all six tests pass.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookAggregateMessagesTool.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookAggregateMessagesToolTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(reports): add OutlookAggregateMessagesTool" -m "Tool wrapping IOutlookSurface.AggregateMessages. Uses AggregateMessagesArgsParser to parse JSON, sums bucket counts into a top-level total, projects each bucket as {label, count}. OperationCanceledException becomes a structured cancel envelope."
```

---

## Task 10: Implement `LiveOutlookSurface.ReadMessages`

Replace the throwing stub with a real implementation that resolves short IDs, calls `Application.Session.GetItemFromID`, builds `MessageDetail` per item, yields between items, and clamps to `maxItems`.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`

- [ ] **Step 1: Replace the stub**

In `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`, find the stub:

```csharp
public IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken))
{
    throw new System.NotImplementedException("LiveOutlookSurface.ReadMessages will be implemented in a follow-up task.");
}
```

Replace with:

```csharp
public IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken))
{
    if (ids == null || ids.Length == 0) return new MessageDetail[0];
    if (maxItems < 1) maxItems = 1;
    if (maxItems > 100) maxItems = 100;

    var capped = ids.Take(maxItems).ToArray();
    var results = new List<MessageDetail>(capped.Length);

    _marshaller.RunAsync(() =>
    {
        foreach (var shortId in capped)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                var entryId = _ids.Resolve(shortId);
                var item = _application.Session.GetItemFromID(entryId) as Outlook.MailItem;
                if (item == null) continue;
                results.Add(BuildMessageDetail(shortId, item, includeBody));
            }
            catch (COMException ex)
            {
                try { OutlookAI.Diagnostics.TraceLog.Write("ReadMessages COMException id=" + shortId + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {
                try { OutlookAI.Diagnostics.TraceLog.Write("ReadMessages unknown short id=" + shortId, "LiveOutlookSurface"); } catch { }
            }
            // Pump Outlook UI between items so a long ids[] does not
            // freeze the UI thread for the full read sweep.
            YieldUi(ct);
        }
    }, ct).GetAwaiter().GetResult();

    return results;
}

// Builds a MessageDetail for a single MailItem. Mirrors ReadMessage's
// projection so the bulk and single tools return identical shapes.
private MessageDetail BuildMessageDetail(string shortId, Outlook.MailItem item, bool includeBody)
{
    var body = "";
    bool truncated = false;
    if (includeBody)
    {
        try { body = item.Body ?? ""; } catch (COMException) { }
        if (body.Length > MaxBodyChars)
        {
            body = body.Substring(0, MaxBodyChars);
            truncated = true;
        }
    }

    var attachments = new List<AttachmentSummary>();
    try
    {
        foreach (Outlook.Attachment att in item.Attachments)
        {
            attachments.Add(new AttachmentSummary
            {
                Filename = att.FileName,
                SizeBytes = att.Size,
            });
        }
    }
    catch (COMException) { }

    string subject = ""; try { subject = item.Subject ?? ""; } catch (COMException) { }
    string sender = ""; try { sender = item.SenderName ?? item.SenderEmailAddress ?? ""; } catch (COMException) { }
    string to = ""; try { to = item.To ?? ""; } catch (COMException) { }
    string cc = ""; try { cc = item.CC ?? ""; } catch (COMException) { }
    DateTimeOffset receivedAt = DateTimeOffset.MinValue;
    try { receivedAt = ToOffset(item.ReceivedTime); } catch (COMException) { }
    string conversationTopic = ""; try { conversationTopic = item.ConversationTopic ?? ""; } catch (COMException) { }

    return new MessageDetail
    {
        Id = shortId,
        Subject = subject,
        From = sender,
        To = SplitAddresses(to),
        Cc = SplitAddresses(cc),
        ReceivedAt = receivedAt,
        BodyPlaintext = body,
        BodyTruncated = truncated,
        Attachments = attachments,
        InReplyToMessageId = null,
        ConversationTopic = conversationTopic,
    };
}
```

`MaxBodyChars`, `_marshaller`, `_application`, `_ids`, `SplitAddresses`, `ToOffset`, `YieldUi` are already defined in `LiveOutlookSurface` from Phase 3a/3b. No new helpers needed.

- [ ] **Step 2: Build to verify the file still compiles**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: clean build.

- [ ] **Step 3: Run full test suite (we have no unit test for `LiveOutlookSurface.ReadMessages` — verification is smoke-only — but other tests must still pass)**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes.

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs
git commit -m "feat(reports): implement LiveOutlookSurface.ReadMessages" -m "Replaces the NotImplementedException stub with a real implementation. Resolves each short id to an EntryID via IdResolver, fetches via Session.GetItemFromID on the marshalled UI thread, builds MessageDetail (same shape as ReadMessage). Yields between items so a long ids[] does not freeze Outlook for the full sweep. Bad ids and COM failures are traced + skipped, the batch continues. Body access honours includeBody and the same MaxBodyChars cap as ReadMessage."
```

---

## Task 11: Implement `LiveOutlookSurface.AggregateMessages`

Walk the resolved folder list, for each folder build a Table with the columns we need, classify mail-only rows, accumulate per-key counts in a Dictionary, then top-N at the end. Uses the same per-folder marshaller pattern as the fallback search path so the UI stays responsive.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`

- [ ] **Step 1: Replace the stub**

In `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`, find:

```csharp
public IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken))
{
    throw new System.NotImplementedException("LiveOutlookSurface.AggregateMessages will be implemented in a follow-up task.");
}
```

Replace with:

```csharp
public IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken))
{
    args = args ?? new AggregateMessagesArgs();
    var scopeMode = (args.Scope ?? "auto").Trim().ToLowerInvariant();
    var filter = BuildAggregateFilter(args);
    var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

    // Step 1: resolve folder list on UI thread (reuse the existing
    // method used by the search fallback).
    IReadOnlyList<Outlook.MAPIFolder> folders;
    try
    {
        folders = _marshaller.RunAsync(
            () => ResolveSearchFolders(
                new SearchMessagesArgs { FolderId = args.FolderId },
                allMail: scopeMode != "current_folder",
                ct),
            ct).GetAwaiter().GetResult();
    }
    catch (OperationCanceledException) { throw; }

    // Step 2: per-folder Table API read, classifier-filter mail rows,
    // group into the dictionary. One marshalled call per folder so the
    // UI thread is released between folders.
    foreach (var folder in folders)
    {
        ct.ThrowIfCancellationRequested();
        try
        {
            _marshaller.RunAsync(
                () => AccumulateFolderBuckets(folder, args, filter, counts, ct),
                ct).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException) { throw; }
    }

    // Step 3: convert dictionary to AggregationBucket list and clamp
    // via TopNBucketSelector for deterministic ordering.
    var allBuckets = counts.Select(kv => new AggregationBucket { Label = kv.Key, Count = kv.Value }).ToList();
    return TopNBucketSelector.TakeTop(allBuckets, args.TopN);
}

// Builds a DASL @SQL filter for the aggregate query. Mirrors the
// existing BuildRestrictFilter clauses for date_from / date_to / from /
// subject_contains / body_contains. Returns null if no clauses.
private static string BuildAggregateFilter(AggregateMessagesArgs args)
{
    var clauses = new List<string>();
    if (!string.IsNullOrEmpty(args.From))
    {
        var f = (args.From ?? "").Replace("'", "''");
        clauses.Add("(urn:schemas:httpmail:fromname LIKE '%" + f + "%' OR urn:schemas:httpmail:fromemail LIKE '%" + f + "%')");
    }
    if (!string.IsNullOrEmpty(args.SubjectContains))
    {
        clauses.Add("urn:schemas:httpmail:subject LIKE '%" + args.SubjectContains.Replace("'", "''") + "%'");
    }
    if (!string.IsNullOrEmpty(args.BodyContains))
    {
        clauses.Add("urn:schemas:httpmail:textdescription LIKE '%" + args.BodyContains.Replace("'", "''") + "%'");
    }
    if (args.DateFrom.HasValue)
    {
        clauses.Add("urn:schemas:httpmail:datereceived >= '" + args.DateFrom.Value.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture) + "'");
    }
    if (args.DateTo.HasValue)
    {
        clauses.Add("urn:schemas:httpmail:datereceived <= '" + args.DateTo.Value.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture) + "'");
    }
    if (clauses.Count == 0) return null;
    return "@SQL=" + string.Join(" AND ", clauses);
}

// Read one folder's rows via the Table API, classify mail-only, and
// accumulate counts into the shared dictionary keyed by args.GroupBy.
private void AccumulateFolderBuckets(
    Outlook.MAPIFolder folder,
    AggregateMessagesArgs args,
    string filter,
    Dictionary<string, int> counts,
    CancellationToken ct)
{
    if (folder == null) return;

    string folderName = "";
    bool folderIsMail = true;
    try { folderName = folder.Name ?? ""; } catch (COMException) { }
    try { folderIsMail = folder.DefaultItemType == Outlook.OlItemType.olMailItem; } catch (COMException) { }
    if (_classifier.IsSystemFolder(folderName, folderIsMail)) return;

    Outlook.Table table;
    try
    {
        table = folder.GetTable(filter ?? "", Outlook.OlTableContents.olUserItems);
    }
    catch (COMException ex)
    {
        try { OutlookAI.Diagnostics.TraceLog.Write("AggregateMessages GetTable COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
        return;
    }

    try
    {
        table.Columns.RemoveAll();
        table.Columns.Add("SenderName");
        table.Columns.Add("SenderEmailAddress");
        table.Columns.Add("ReceivedTime");
        table.Columns.Add("MessageClass");
    }
    catch (COMException) { }

    int rowsScanned = 0;
    try
    {
        while (!table.EndOfTable)
        {
            ct.ThrowIfCancellationRequested();
            rowsScanned++;
            Outlook.Row row;
            try { row = table.GetNextRow(); }
            catch (COMException) { break; }
            if (row == null) break;

            string messageClass = "";
            try { messageClass = row["MessageClass"] as string ?? ""; } catch (COMException) { }
            if (!TableMessageClassFilter.IsMailMessage(messageClass)) continue;

            string key = ResolveBucketKey(row, folderName, args.GroupBy);
            if (key == null) continue;

            int existing;
            counts.TryGetValue(key, out existing);
            counts[key] = existing + 1;

            // Pump UI every ~50 rows so a big folder doesn't freeze
            // the UI for the whole folder scan.
            if ((rowsScanned % 50) == 0) YieldUi(ct);
        }
    }
    catch (COMException ex)
    {
        try { OutlookAI.Diagnostics.TraceLog.Write("AggregateMessages row scan COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
    }
}

// Picks the bucket key for one row given args.GroupBy.
private static string ResolveBucketKey(Outlook.Row row, string folderName, string groupBy)
{
    if (string.Equals(groupBy, "folder", StringComparison.OrdinalIgnoreCase))
    {
        return string.IsNullOrEmpty(folderName) ? null : folderName;
    }

    if (string.Equals(groupBy, "day", StringComparison.OrdinalIgnoreCase))
    {
        try
        {
            var rt = row["ReceivedTime"];
            if (rt is DateTime dt) return DateBucketFormatter.Format(new DateTimeOffset(dt));
        }
        catch (COMException) { }
        return DateBucketFormatter.UnknownDate;
    }

    // Default: sender
    string name = "";
    string email = "";
    try { name = row["SenderName"] as string ?? ""; } catch (COMException) { }
    try { email = row["SenderEmailAddress"] as string ?? ""; } catch (COMException) { }
    return SenderKeyNormalizer.Normalize(name, email);
}
```

The helpers `_marshaller`, `_classifier`, `ResolveSearchFolders`, `YieldUi`, `TableMessageClassFilter`, `SenderKeyNormalizer`, `DateBucketFormatter`, `TopNBucketSelector` are all already in scope from earlier Phase 3b code or earlier tasks in this plan.

- [ ] **Step 2: Build to verify**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: clean build.

- [ ] **Step 3: Run full test suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes.

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs
git commit -m "feat(reports): implement LiveOutlookSurface.AggregateMessages" -m "Walks resolved folders, opens an Outlook.Table per folder with [SenderName, SenderEmailAddress, ReceivedTime, MessageClass] columns plus the DASL filter from args (from / subject_contains / body_contains / date_from / date_to). Filters non-mail rows via TableMessageClassFilter, accumulates counts in a Dictionary keyed by SenderKeyNormalizer / DateBucketFormatter / folder name (per args.GroupBy). Per-folder runs in its own marshalled UI-thread call so Outlook stays responsive between folders; within a folder, YieldUi pumps the UI every 50 rows. Final clamp via TopNBucketSelector."
```

---

## Task 12: Register both tools in `ToolCatalogSchema`

The tool catalog tells the model what tools exist and how to use them. The schema description steers the model toward when to prefer each tool — that's where we put "use bulk read instead of N×read_message" and similar guidance.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`

- [ ] **Step 1: Write the failing schema tests**

In `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`, add to the existing class:

```csharp
[Fact]
public void OutlookReadMessages_Schema_HasIdsArrayAndBodyToggle()
{
    var catalog = ToolCatalogSchema.BuildToolsArray(includeWriteTools: false);
    var tool = catalog.FirstOrDefault(t => (string)t["name"] == "outlook_read_messages");
    Assert.NotNull(tool);
    var properties = tool["parameters"]?["properties"];
    Assert.NotNull(properties?["ids"]);
    Assert.Equal("array", (string)properties["ids"]["type"]);
    Assert.NotNull(properties["include_body"]);
    Assert.Equal("boolean", (string)properties["include_body"]["type"]);
    Assert.NotNull(properties["max_items"]);
}

[Fact]
public void OutlookReadMessages_Description_HintsAtUseInsteadOfManyReadCalls()
{
    var catalog = ToolCatalogSchema.BuildToolsArray(includeWriteTools: false);
    var tool = catalog.FirstOrDefault(t => (string)t["name"] == "outlook_read_messages");
    var desc = (string)tool["description"] ?? "";
    Assert.Contains("read_message", desc, System.StringComparison.OrdinalIgnoreCase);
}

[Fact]
public void OutlookAggregateMessages_Schema_HasGroupByEnumAndTopN()
{
    var catalog = ToolCatalogSchema.BuildToolsArray(includeWriteTools: false);
    var tool = catalog.FirstOrDefault(t => (string)t["name"] == "outlook_aggregate_messages");
    Assert.NotNull(tool);
    var properties = tool["parameters"]?["properties"];
    Assert.NotNull(properties?["group_by"]);
    var groupByEnum = properties["group_by"]?["enum"] as Newtonsoft.Json.Linq.JArray;
    Assert.NotNull(groupByEnum);
    var enumValues = groupByEnum.Select(t => (string)t).ToArray();
    Assert.Contains("sender", enumValues);
    Assert.Contains("day", enumValues);
    Assert.Contains("folder", enumValues);
    Assert.NotNull(properties["top_n"]);
}

[Fact]
public void OutlookAggregateMessages_Description_HintsAtUseInsteadOfManyCountCalls()
{
    var catalog = ToolCatalogSchema.BuildToolsArray(includeWriteTools: false);
    var tool = catalog.FirstOrDefault(t => (string)t["name"] == "outlook_aggregate_messages");
    var desc = (string)tool["description"] ?? "";
    Assert.Contains("count", desc, System.StringComparison.OrdinalIgnoreCase);
}
```

Make sure `using System.Linq;` is at the top of the test file (it likely already is). Make sure `using Newtonsoft.Json.Linq;` is present.

- [ ] **Step 2: Run the new tests, verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ToolCatalogSchemaTests"
```

Expected: the four new tests fail because the two tools are missing from the catalog.

- [ ] **Step 3: Add both schemas to `ToolCatalogSchema`**

Open `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`. Find the existing `BuildToolsArray` method that returns a `JArray` of tool objects. Add two new entries (near the existing `outlook_read_message` and `outlook_count_messages` entries, before the write tools section):

```csharp
// outlook_read_messages — bulk read by ID array. Use this instead of
// many outlook_read_message round trips when you have an array of IDs
// from a prior outlook_search_messages and you need the bodies.
new JObject(
    new JProperty("type", "function"),
    new JProperty("name", "outlook_read_messages"),
    new JProperty("description",
        "Bulk-read message details by short ID array. Returns subject/sender/" +
        "date/body/attachments per id. Use this instead of multiple outlook_read_message " +
        "calls; ~5-10x faster when you have many IDs from a search."),
    new JProperty("parameters",
        new JObject(
            new JProperty("type", "object"),
            new JProperty("properties",
                new JObject(
                    new JProperty("ids",
                        new JObject(
                            new JProperty("type", "array"),
                            new JProperty("items", new JObject(new JProperty("type", "string"))),
                            new JProperty("description", "Short message IDs from a prior outlook_search_messages."))),
                    new JProperty("include_body",
                        new JObject(
                            new JProperty("type", "boolean"),
                            new JProperty("description", "Default true. Set false to skip body bodies for a faster lightweight read."))),
                    new JProperty("max_items",
                        new JObject(
                            new JProperty("type", "integer"),
                            new JProperty("description", "Default 25, hard cap 100."))))),
            new JProperty("required", new JArray("ids"))))),

// outlook_aggregate_messages — group + count. Use this for statistics
// instead of looping outlook_count_messages with different filters.
new JObject(
    new JProperty("type", "function"),
    new JProperty("name", "outlook_aggregate_messages"),
    new JProperty("description",
        "Group matching messages by sender, day, or folder and return the top-N buckets by count. " +
        "Use this instead of calling outlook_count_messages many times when the user wants statistics " +
        "(top senders, busiest days, breakdown by folder). Examples: 'top 10 senders this month', " +
        "'busiest days last week', 'breakdown by folder'."),
    new JProperty("parameters",
        new JObject(
            new JProperty("type", "object"),
            new JProperty("properties",
                new JObject(
                    new JProperty("scope",
                        new JObject(
                            new JProperty("type", "string"),
                            new JProperty("enum", new JArray("auto", "current_folder", "all_mail")),
                            new JProperty("description", "Default auto."))),
                    new JProperty("folder_id",
                        new JObject(
                            new JProperty("type", "string"),
                            new JProperty("description", "Optional explicit folder id. Overrides scope when provided."))),
                    new JProperty("date_from",
                        new JObject(
                            new JProperty("type", "string"),
                            new JProperty("description", "ISO timestamp. Optional lower bound on ReceivedTime."))),
                    new JProperty("date_to",
                        new JObject(
                            new JProperty("type", "string"),
                            new JProperty("description", "ISO timestamp. Optional upper bound on ReceivedTime."))),
                    new JProperty("from",
                        new JObject(
                            new JProperty("type", "string"),
                            new JProperty("description", "Optional sender name or email substring."))),
                    new JProperty("subject_contains",
                        new JObject(
                            new JProperty("type", "string"),
                            new JProperty("description", "Optional subject substring."))),
                    new JProperty("body_contains",
                        new JObject(
                            new JProperty("type", "string"),
                            new JProperty("description", "Optional body substring."))),
                    new JProperty("group_by",
                        new JObject(
                            new JProperty("type", "string"),
                            new JProperty("enum", new JArray("sender", "day", "folder")),
                            new JProperty("description", "How to bucket matching messages."))),
                    new JProperty("top_n",
                        new JObject(
                            new JProperty("type", "integer"),
                            new JProperty("description", "Default 10, hard cap 100."))))),
            new JProperty("required", new JArray("group_by"))))),
```

If `BuildToolsArray` uses a different style of JObject construction in your project, follow that style instead (look at the existing entries — they're the model to mirror).

- [ ] **Step 4: Run the schema tests**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ToolCatalogSchemaTests"
```

Expected: all `ToolCatalogSchemaTests` (existing + 4 new) pass.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs
git commit -m "feat(reports): register outlook_read_messages and outlook_aggregate_messages schemas" -m "Adds both tools to ToolCatalogSchema with descriptions that steer the model toward using bulk-read instead of many read_message calls and using aggregate-messages instead of many count_messages calls. group_by is an enum (sender|day|folder); ids is a required array of strings."
```

---

## Task 13: Register both tools in `OutlookToolHost`

`ToolCatalogSchema` advertises tools to the model. `OutlookToolHost` is what actually routes a dispatched tool name to its `IOutlookTool` instance. Both registrations are required.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/OutlookToolHost.cs`

- [ ] **Step 1: Find the existing tool registration**

Open `VSTO2/OutlookAI/Services/OutlookToolHost.cs`. Locate where the existing tools are registered (typically a list initialization in the constructor or a `RegisterTools` method, listing `OutlookSearchMessagesTool`, `OutlookCountMessagesTool`, `OutlookReadMessageTool`, etc.).

- [ ] **Step 2: Add both new tools**

In the same registration list / sequence, add:

```csharp
new OutlookReadMessagesTool(),
new OutlookAggregateMessagesTool(),
```

Order doesn't functionally matter — the dispatcher routes by name — but keep them near the related tools (read_messages near read_message, aggregate_messages near count_messages) for readability.

- [ ] **Step 3: Build and run full test suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes. The tool host is now ready to dispatch both new tools when the model emits a function call with one of their names.

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/OutlookToolHost.cs
git commit -m "feat(reports): wire outlook_read_messages and outlook_aggregate_messages into OutlookToolHost" -m "Adds both new IOutlookTool implementations to the tool host's registration list so the dispatcher can route function calls to them. After this commit the model can actually invoke the tools (schema was registered in the previous commit; this is the runtime dispatch wiring)."
```

---

## Task 14: Wire up the Inbox Reports task pane (full integration)

Adds the InboxReportsPane + Controller, the ribbon button, the `ThisAddIn.ShowReportsTaskPane(Explorer)` method, and connects it all so clicking the new ribbon button opens a working Reports pane with chips and chat.

This is the largest task in the plan. Many small steps but each is mechanical (most parallels the InboxCopilot pattern that already works).

**Files:**
- Create: `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.cs`
- Create: `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.Designer.cs`
- Create: `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.resx`
- Create: `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsController.cs`
- Modify: `VSTO2/OutlookAI/Ribbon.xml`
- Modify: `VSTO2/OutlookAI/Ribbon.cs`
- Modify: `VSTO2/OutlookAI/ThisAddIn.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`

- [ ] **Step 1: Create the pane (a UserControl, mirror of `InboxCopilotPane`)**

Create `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.cs`:

```csharp
using System;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.Chat;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane.InboxReports
{
    public partial class InboxReportsPane : UserControl
    {
        private Outlook.Explorer _explorer;
        private LiveOutlookSurface _surface;
        private OutlookToolHost _toolHost;
        private WebView2 _webView;
        private InboxReportsController _controller;

        public InboxReportsPane()
        {
            using (TraceLog.Scope("ctor", "InboxReportsPane"))
            {
                InitializeComponent();
                _webView = new WebView2 { Dock = DockStyle.Fill };
                Controls.Add(_webView);
            }
        }

        public CodexChatService ChatService => Globals.ThisAddIn?.ChatService;

        public void Bind(Outlook.Explorer explorer)
        {
            using (TraceLog.Scope("Bind", "InboxReportsPane"))
            {
                _explorer = explorer;
                try
                {
                    var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                    var ids = Globals.ThisAddIn?.IdResolver;
                    var app = Globals.ThisAddIn?.Application;
                    var runner = Globals.ThisAddIn?.AdvancedSearchRunner;
                    var classifier = Globals.ThisAddIn?.FolderClassifier;
                    TraceLog.Write("Services: marshaller=" + (marshaller != null) +
                        " ids=" + (ids != null) + " app=" + (app != null) +
                        " runner=" + (runner != null), "InboxReportsPane");
                    if (marshaller != null && ids != null && app != null && runner != null)
                    {
                        _surface = new LiveOutlookSurface(app, marshaller, ids,
                            composeInspector: null, explorer: explorer,
                            runner: runner, classifier: classifier);
                        _toolHost = new OutlookToolHost(_surface, Config.WriteToolsEnabled);
                        TraceLog.Write("surface + toolHost constructed", "InboxReportsPane");
                    }
                    else
                    {
                        TraceLog.Write("surface NOT constructed (missing service); runner=" + (runner != null),
                            "InboxReportsPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("surface/toolHost error: " + ex, "InboxReportsPane");
                }

                try
                {
                    if (ChatService != null && _toolHost != null && _surface != null)
                    {
                        _controller = new InboxReportsController(
                            _webView, _surface, _toolHost, ChatService);
                        TraceLog.Write("Controller constructed; firing InitializeAsync", "InboxReportsPane");
                        _ = _controller.InitializeAsync();
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("controller error: " + ex, "InboxReportsPane");
                }
            }
        }
    }
}
```

- [ ] **Step 2: Create the designer file**

Create `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.Designer.cs`:

```csharp
namespace OutlookAI.TaskPane.InboxReports
{
    partial class InboxReportsPane
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = false;
            this.Size = new System.Drawing.Size(340, 600);
        }
    }
}
```

Create empty `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.resx` (copy `VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPane.resx` and replace the class name only).

- [ ] **Step 3: Create the controller**

Create `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsController.cs`. The controller mirrors `InboxCopilotController` closely but:
- Uses `InboxReportsPromptBuilder` for the system prompt.
- Pushes `ReportQuickActionChip.Defaults()` once on ready (no selection-driven re-pushes).
- Does **not** subscribe to selection-change events.

Use `VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotController.cs` as the model. Concrete required structure:

```csharp
using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.Chat;

namespace OutlookAI.TaskPane.InboxReports
{
    public sealed class InboxReportsController : IDisposable
    {
        private readonly WebView2 _webView;
        private readonly LiveOutlookSurface _surface;
        private readonly OutlookToolHost _toolHost;
        private readonly CodexChatService _chat;
        private readonly ConversationStore _store = new ConversationStore();
        private readonly InboxReportsPromptBuilder _promptBuilder = new InboxReportsPromptBuilder();
        private CancellationTokenSource _activeCts;
        private bool _isReady;
        private bool _turnInFlight;

        public InboxReportsController(
            WebView2 webView,
            LiveOutlookSurface surface,
            OutlookToolHost toolHost,
            CodexChatService chat)
        {
            _webView = webView;
            _surface = surface;
            _toolHost = toolHost;
            _chat = chat;
        }

        public async Task InitializeAsync()
        {
            using (TraceLog.Scope("InitializeAsync (sync prefix)", "InboxReports"))
            {
                await WebView2Bootstrap.InitializeAsync(_webView).ConfigureAwait(true);
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
            }
        }

        private async void OnWebMessageReceived(object sender,
            Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var raw = e.WebMessageAsJson;
                TraceLog.Write("WebMessageReceived: " + raw, "InboxReports");
                var obj = JObject.Parse(raw);
                var type = (string)obj["type"];
                switch (type)
                {
                    case "ready":
                        await OnWebViewReady().ConfigureAwait(false);
                        break;
                    case "send":
                        var text = (string)obj["payload"]?["text"];
                        var reasoning = (string)obj["payload"]?["reasoning"];
                        await StartTurnAsync(text, reasoning).ConfigureAwait(false);
                        break;
                    case "stop":
                        try { _activeCts?.Cancel(); } catch { }
                        break;
                    case "clear":
                        _store.Clear();
                        await RunScript("outlookai.clear();").ConfigureAwait(false);
                        break;
                    case "copy":
                        var clip = _store.ExportForClipboard();
                        await RunScript("outlookai.setClipboard(" + JsString(clip) + ");").ConfigureAwait(false);
                        break;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived error: " + ex, "InboxReports");
            }
        }

        private async Task OnWebViewReady()
        {
            using (TraceLog.Scope("OnWebViewReady", "InboxReports"))
            {
                _isReady = true;
                await PushChips().ConfigureAwait(false);
            }
        }

        private async Task PushChips()
        {
            try
            {
                var chips = ReportQuickActionChip.Defaults();
                var chipsArr = new JArray();
                foreach (var c in chips)
                {
                    chipsArr.Add(new JObject(
                        new JProperty("label", c.Label),
                        new JProperty("prompt", c.TemplateText)));
                }
                await RunScript("outlookai.setQuickActions(" +
                    chipsArr.ToString(Newtonsoft.Json.Formatting.None) + ");").ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushChips error: " + ex.Message, "InboxReports");
            }
        }

        private async Task StartTurnAsync(string userText, string reasoningOverride)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "InboxReports");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "InboxReports");
                return;
            }
            _turnInFlight = true;
            await RunScript("outlookai.setComposerEnabled(false, true);").ConfigureAwait(false);

            _activeCts?.Dispose();
            _activeCts = new CancellationTokenSource();
            var ct = _activeCts.Token;

            try
            {
                _store.AppendUserMessage(userText);
                var systemPrompt = _promptBuilder.Build();
                var assistantId = "rep-" + Guid.NewGuid().ToString("N").Substring(0, 8);
                await RunScript("outlookai.beginAssistantMessage(" + JsString(assistantId) + ");").ConfigureAwait(false);

                var sink = new WebViewSink(this, assistantId);
                var result = await _chat.RunTurnAsync(
                    systemPrompt,
                    _store,
                    _toolHost,
                    sink,
                    reasoningOverride,
                    ct).ConfigureAwait(false);

                await RunScript("outlookai.finalizeAssistantMessage("
                    + JsString(assistantId) + ", "
                    + new JObject(
                        new JProperty("stopped", result.StopReason == StopReason.Cancelled),
                        new JProperty("error", result.StopReason == StopReason.Error)).ToString(Newtonsoft.Json.Formatting.None)
                    + ");").ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                await RunScript("outlookai.finalizeAssistantMessage(\"rep\", {stopped:true});").ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                TraceLog.Write("StartTurnAsync error: " + ex, "InboxReports");
            }
            finally
            {
                _turnInFlight = false;
                await RunScript("outlookai.setComposerEnabled(true, false);").ConfigureAwait(false);
                TraceLog.Write("<< StartTurnAsync", "InboxReports");
            }
        }

        internal async Task RunScript(string script)
        {
            try
            {
                if (_webView?.CoreWebView2 == null) return;
                await _webView.CoreWebView2.ExecuteScriptAsync(script).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunScript error: " + ex.Message, "InboxReports");
            }
        }

        internal static string JsString(string s)
        {
            return Newtonsoft.Json.JsonConvert.ToString(s ?? "");
        }

        public void Dispose()
        {
            try { _activeCts?.Cancel(); } catch { }
        }

        // Streams assistant text + tool-call status back into the WebView.
        private sealed class WebViewSink : Services.Chat.ChatEventSink
        {
            private readonly InboxReportsController _ctrl;
            private readonly string _assistantId;
            public WebViewSink(InboxReportsController c, string id) { _ctrl = c; _assistantId = id; }
            public override void OnAssistantDelta(string text)
            {
                _ = _ctrl.RunScript("outlookai.appendAssistantDelta("
                    + JsString(_assistantId) + ", " + JsString(text) + ");");
            }
            public override void OnToolCallStart(string callId, string toolName, string argsJson)
            {
                TraceLog.Write("Sink.OnToolCallStart " + toolName + " args=" + argsJson, "WebViewSink");
                _ = _ctrl.RunScript("outlookai.appendToolCallCard("
                    + JsString(callId) + ", " + JsString(toolName) + ", " + JsString(argsJson) + ");");
            }
            public override void OnToolCallEnd(string callId, bool ok, string summary, string resultJson)
            {
                _ = _ctrl.RunScript("outlookai.updateToolCallCard("
                    + JsString(callId) + ", " + (ok ? "true" : "false") + ", "
                    + JsString(summary) + ", " + JsString(resultJson) + ");");
            }
        }
    }
}
```

The controller mirrors InboxCopilotController's shape. Differences:
- No selection-change subscription.
- One-time chip push on ready (chips never change).
- Uses `InboxReportsPromptBuilder.Build()` directly with no context.

If the existing `WebView2Bootstrap`, `ChatEventSink`, or `ConversationStore` APIs differ from what's shown, mirror the existing `InboxCopilotController`'s exact usage and propagate any adjustments here.

- [ ] **Step 4: Register the controller in csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj`, add (alphabetically with other `TaskPane\InboxReports\*`):

```xml
    <Compile Include="TaskPane\InboxReports\InboxReportsController.cs" />
    <Compile Include="TaskPane\InboxReports\InboxReportsPane.cs">
      <SubType>UserControl</SubType>
    </Compile>
    <Compile Include="TaskPane\InboxReports\InboxReportsPane.Designer.cs">
      <DependentUpon>InboxReportsPane.cs</DependentUpon>
    </Compile>
```

And in the `<EmbeddedResource>` section (after the existing `InboxCopilotPane.resx`):

```xml
    <EmbeddedResource Include="TaskPane\InboxReports\InboxReportsPane.resx">
      <DependentUpon>InboxReportsPane.cs</DependentUpon>
    </EmbeddedResource>
```

- [ ] **Step 5: Add the ribbon button**

In `VSTO2/OutlookAI/Ribbon.xml`, modify the `TabMail` tab group to add a second button next to the existing AI Assistant one:

```xml
<tab idMso="TabMail">
    <group id="AIAssistantExplorerGroup" label="AI Assistant" insertAfterMso="GroupMailMove">
        <button id="btnAIAssistantExplorer"
                label="AI Assistant"
                size="large"
                onAction="OnAIAssistantClick"
                supertip="Open the AI Assistant to chat with your mailbox - summarize, search, and act on messages."
                imageMso="SmartArtChangeColorsGallery" />
        <button id="btnReports"
                label="Reports"
                size="large"
                onAction="OnReportsClick"
                supertip="Open the Inbox Reports pane to generate digests, action items, conversation summaries, and stats."
                imageMso="ChartTypeMenu" />
    </group>
</tab>
```

In `VSTO2/OutlookAI/Ribbon.cs`, add a handler:

```csharp
public void OnReportsClick(Office.IRibbonControl control)
{
    Globals.ThisAddIn.ShowReportsTaskPane();
}
```

- [ ] **Step 6: Add `ShowReportsTaskPane` to `ThisAddIn`**

Open `VSTO2/OutlookAI/ThisAddIn.cs`. After `ShowExplorerTaskPane`, add:

```csharp
public void ShowReportsTaskPane()
{
    using (TraceLog.Scope("ShowReportsTaskPane", "ThisAddIn"))
    try
    {
        object activeWindow = null;
        try { activeWindow = this.Application.ActiveWindow(); } catch { }
        TraceLog.Write("ActiveWindow=" + (activeWindow?.GetType().FullName ?? "<null>"), "ThisAddIn");

        if (activeWindow is Outlook.Explorer expl)
        {
            ShowReportsExplorerTaskPane(expl);
            return;
        }
        // Reports only makes sense on an Explorer (Inbox view), not a
        // compose window.
        System.Windows.Forms.MessageBox.Show(
            "Open Outlook to your Inbox, then click Reports.",
            "Inbox Reports",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
    }
    catch (Exception ex)
    {
        TraceLog.Write("ShowReportsTaskPane error: " + ex, "ThisAddIn");
        System.Windows.Forms.MessageBox.Show(
            $"Error: {ex.Message}",
            "Inbox Reports",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}

private void ShowReportsExplorerTaskPane(Outlook.Explorer explorer)
{
    foreach (Microsoft.Office.Tools.CustomTaskPane pane in this.CustomTaskPanes)
    {
        if (pane.Window == explorer && pane.Control is InboxReports.InboxReportsPane)
        {
            TraceLog.Write("Reusing existing Reports CustomTaskPane (toggle visibility)", "ThisAddIn");
            pane.Visible = !pane.Visible;
            return;
        }
    }
    TraceLog.Write("Creating new InboxReportsPane for Explorer", "ThisAddIn");
    var paneControl = new InboxReports.InboxReportsPane();
    paneControl.Bind(explorer);
    var ctp = this.CustomTaskPanes.Add(paneControl, "Inbox Reports", explorer);
    ctp.Width = 340;
    ctp.Visible = true;
    TraceLog.Write("Reports CustomTaskPane.Visible = true", "ThisAddIn");
}
```

Make sure `using OutlookAI.TaskPane.InboxReports;` is at the top of `ThisAddIn.cs` (it may already be if other panes are imported similarly; otherwise add it).

Also update the loop above for `ShowExplorerTaskPane` to check for the InboxCopilotPane specifically — the existing loop matches by `pane.Window == explorer` which will incorrectly toggle Reports when the InboxCopilot button is clicked. Change the existing `ShowExplorerTaskPane` to:

```csharp
foreach (CustomTaskPane pane in this.CustomTaskPanes)
{
    if (pane.Window == explorer && pane.Control is InboxCopilotPane)
    {
        // ... existing body unchanged
        pane.Visible = !pane.Visible;
        return;
    }
}
```

so each ribbon button matches its own pane.

- [ ] **Step 7: Build and run the full test suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: clean build, all tests pass.

- [ ] **Step 8: Commit**

```powershell
git add -A
git commit -m "feat(reports): wire Inbox Reports pane, controller, ribbon, and ThisAddIn entry" -m "Adds the InboxReportsPane (parallel to InboxCopilotPane), InboxReportsController (mirrors InboxCopilotController but no selection awareness — chips never change), a new Reports ribbon button on TabMail with its own OnReportsClick handler, and ThisAddIn.ShowReportsTaskPane(Explorer) which toggles a per-Explorer Reports CustomTaskPane just like the InboxCopilot path. The existing ShowExplorerTaskPane is tightened to match only InboxCopilotPane instances so each ribbon button toggles its own pane." -m "Conversation state, tools, runner, and marshaller are shared via Globals.ThisAddIn singletons; only ConversationStore and OutlookToolHost are per-pane so the two panes never collide on conversation history or tool dispatch."
```

---

## Task 15: Final verification, deploy, and smoke

- [ ] **Step 1: Run the full test suite one more time**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: every test passes. Note total count for the eventual commit message.

- [ ] **Step 2: Publish Release to staging**

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" `
    /p:PublishDir="$staging\" /v:minimal /nologo
Copy-Item -LiteralPath "Deploy\Install-OutlookAI.ps1" -Destination "$staging\" -Force
```

Expected: Release publish clean.

- [ ] **Step 3: Verify Outlook is closed, install elevated**

```powershell
$procs = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) { "Outlook is RUNNING - close before install"; exit }
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
$script  = Join-Path $staging "Install-OutlookAI.ps1"
$args    = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -SourcePath `"$staging`""
$proc    = Start-Process -FilePath "powershell.exe" -ArgumentList $args -Verb RunAs -Wait -PassThru
"Installer exit code: $($proc.ExitCode)"
```

Expected: installer exit code 0.

- [ ] **Step 4: Verify installed hash**

```powershell
$staged    = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2\OutlookAI.dll"
$installed = "C:\Program Files\OutlookAI\OutlookAI.dll"
"staged    = $((Get-FileHash -LiteralPath $staged -Algorithm SHA256).Hash)"
"installed = $((Get-FileHash -LiteralPath $installed -Algorithm SHA256).Hash)"
```

Expected: hashes match.

- [ ] **Step 5: Smoke — open Reports pane**

Open Outlook. Click the new "Reports" ribbon button in the Mail tab. Verify:
- A second task pane opens titled "Inbox Reports".
- It shows a row of six chips with the spec's labels.
- The chat composer is at the bottom.
- Clicking the same ribbon button again toggles the pane closed.

- [ ] **Step 6: Smoke — chips prefill correctly**

For each of the six chips:
- Click the chip.
- Verify the input box is prefilled with the chip's template text.
- For chips 2, 4, 6: verify the `[placeholder]` text is visible.
- Do NOT submit yet.

- [ ] **Step 7: Smoke — Daily Digest report end-to-end**

- Click the `📅 This week's digest` chip.
- Click Send.
- Verify: a structured markdown report renders in the pane chat within ~30s on the 200-folder mailbox. Header mentions scope + count. Outlook UI stays responsive throughout.

Read trace:
```powershell
Get-Content -LiteralPath "C:\Users\MDASR\AppData\Local\OutlookAI\trace.log" -Tail 60
```

Verify the trace shows `outlook_search_messages` was called with a date range.

- [ ] **Step 8: Smoke — Action Items uses bulk read**

- Click `✓ Action items` chip.
- Send.
- Verify report renders.
- Inspect trace for `outlook_read_messages` call (NOT a long series of `outlook_read_message` calls).

- [ ] **Step 9: Smoke — Stats uses aggregate**

- Click `📊 Email stats` chip.
- Send.
- Verify report renders with sender/day/folder breakdown.
- Inspect trace for `outlook_aggregate_messages` call.

- [ ] **Step 10: Smoke — Conversation summary asks for clarification**

- Click `💬 Conversation summary` chip.
- Without editing the `[name or email]` placeholder, click Send.
- Verify: the model asks for clarification rather than calling tools with literal `[name or email]` as the argument.
- Then reply with an actual person's name + send. Verify report renders.

- [ ] **Step 11: Smoke — Stop button works**

- Click any chip that triggers a long search (e.g., the digest or stats).
- Before it completes, click Stop on the composer.
- Verify: cancellation propagates within ~1s, the model shows a cancelled response.

- [ ] **Step 12: Smoke — both panes coexist**

- Open Reports pane (as above).
- Open the original AI Assistant pane via its ribbon button.
- Both should be visible simultaneously and independent.
- Try chatting in each — verify each pane has its own conversation history.

- [ ] **Step 13: Push the branch**

If smokes pass:

```powershell
git push origin feature/codex-oauth-migration
```

- [ ] **Step 14: Final summary commit (optional, if any tiny polish fixes came up)**

If smoke uncovered small issues, commit them as `fix(reports): ...` follow-ups. Otherwise the phase is complete with the per-task commits above.

---

## Self-Review

**Spec coverage:**
- Architecture (pane + controller + prompt builder + chips, parallel to InboxCopilot) → Tasks 5, 6, 14.
- `outlook_read_messages` tool + surface → Tasks 7, 8, 10.
- `outlook_aggregate_messages` tool + surface → Tasks 4, 7, 9, 11.
- Pure helpers (SenderKeyNormalizer, DateBucketFormatter, TopNBucketSelector, AggregateMessagesArgsParser) → Tasks 1, 2, 3, 4.
- Tool catalog schema registration → Task 12.
- Tool host registration → Task 13.
- Ribbon + ThisAddIn entry → Task 14.
- Threading & cancellation → reused from Phase 3b; explicit in each surface implementation.
- Error handling → consistent across all tools (cancel envelope, structured errors).
- Testing strategy → each task TDD's its pure unit; surface implementations smoke-only.
- Acceptance criteria (9 items) → mapped to smoke steps 5-12.
- Out of scope (persistence, scheduling, email-to-self, raw CSV, rich UI forms) → not in plan, as intended.

**Placeholder scan:** No "TBD", "later", "TODO" in step instructions. The `[placeholder]` text inside chip templates is intentional user-visible content.

**Type consistency:**
- `IOutlookSurface.ReadMessages` signature `(string[] ids, bool includeBody, int maxItems, CancellationToken ct)` — same across Tasks 7, 8, 10.
- `IOutlookSurface.AggregateMessages(AggregateMessagesArgs, CancellationToken)` — same across Tasks 4, 7, 9, 11.
- `AggregationBucket { Label, Count }` — same across Tasks 3, 9, 11, 12.
- `AggregateMessagesArgs` fields — same across Tasks 4, 9, 11, 12.
- Tool names `outlook_read_messages` and `outlook_aggregate_messages` — consistent across tasks 8, 9, 12, 13.

No mismatches found.

