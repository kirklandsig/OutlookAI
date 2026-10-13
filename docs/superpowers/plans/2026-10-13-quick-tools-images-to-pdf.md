# Quick Tools: Clean up images → PDF Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A ribbon Quick Tools command that turns an email's embedded and attached document photos into one clean, upright PDF, with the AI deciding geometry and the add-in doing deterministic image processing, after a review window.

**Architecture:** `ImagesToPdfCommand` (UI thread) finds the email, saves its image attachments to a per-run temp folder (`EmailImageCollector`), then runs `ImagesToPdfPipeline` on the thread pool: decode upright (`ImageLoader`), ask ChatGPT for a validated `PagePlan` per image (`PagePlanner` → `CodexChatService.CompleteStructuredAsync`, strict JSON schema, image as data URL), flatten and clean each page (`PageCleaner` = `PerspectiveWarp` + `PageLook`) into JPEG files. `ImagesToPdfReviewForm` lets the user adjust; `ImagePdfWriter` writes the PDF from JPEG files. Nothing full-size stays in memory.

**Tech Stack:** C# 7.3, .NET Framework 4.7.2 (VSTO add-in, WinForms, System.Drawing/GDI+, WIC via PresentationCore), Newtonsoft.Json, xUnit 2.9 + VSTest.

**Spec:** `docs/superpowers/specs/2026-10-13-quick-tools-images-to-pdf-design.md` (read it first).

**Deviations from the spec, decided while planning (Task 5 updates the spec):**
- Output page size uses the page's true aspect ratio estimated from the perspective (Zhang & He, whiteboard scanning; verified numerically 2026-10-13: exact on four projected rectangles where the spec's "longer opposite edges" rule was 7–24% off), falling back to the longer-edges rule when the photo shows no usable perspective.
- The command and the two forms live under `VSTO2/OutlookAI/QuickTools/` (UI); everything testable lives under `VSTO2/OutlookAI/Services/QuickTools/`.
- The default file name is whatever `ExportFilenameSanitizer.Build` produces for `"<sender> - <subject>"` (e.g. `Jane-Example-Lease-docs-2026-10-13-1130.pdf`).

## Global Constraints

- C# 7.3 only (no switch expressions, `using var`, ranges, `is not`, nullable reference types, default interface methods). .NET Framework 4.7.2.
- `VSTO2/OutlookAI/OutlookAI.csproj` is old-style: every new `.cs` file needs a `<Compile Include="…" />` line in the existing `<ItemGroup>`, in alphabetical position. The test project is SDK-style and picks up new files on its own.
- No new NuGet packages. Framework references `PresentationCore`, `WindowsBase`, `System.Xaml` are added (Task 4).
- The AI never produces pixels; it only returns a `page_plan` (kind, reason, corners, color, strength).
- Numbers from the spec: images under **200 px** on the short side are skipped; at most **30** pages per run; AI copy at most **1600 px** long edge, JPEG **85**; upright originals saved at JPEG **95**; cleaned pages JPEG **85**; **3** AI requests at a time; **90 s** per image; **one** retry after HTTP 429, waiting `Retry-After` capped at **20 s**, else **5 s**; cleaning **2** pages in parallel; output page capped at **2550 × 3300** px (or 3300 × 2550); PDF pages Letter **612 × 792** pt (landscape 792 × 612), **18 pt** margins; thumbnails **160 px**, preview at most **1200 px**; run folders older than **1 day** are purged.
- Trace log (`TraceLog.Write(…, "QuickTools")`): counts, sizes, durations, model, exception type names only. Never image data, attachment file names, subjects, or AI answers.
- WinForms buttons: `ForeColor = Color.Black`, `BackColor = SystemColors.ButtonFace`, `UseVisualStyleBackColor = false` (Windows Server 2025 renders them blank otherwise).
- Tests touching `Config` statics: `[Collection("Config")]` + `ConfigStateScope`. WinForms tests: `Sta.Run`. Test data is fake (example.com, 555-xxxx, "John Example").
- Every new test must be seen failing before its implementation (TDD), and every behavior fix needs a test that fails without it.
- Ribbon: every callback needs a public handler in `Ribbon.cs`; control ids unique across `Ribbon.xml`.
- Commits: `type(scope): subject` (scope `quick-tools`), only on branch `feat/quick-tools-images-to-pdf`, never `--no-verify`. Every commit message ends with:
  ```
  Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu
  ```

### Commands (PowerShell; Git Bash mangles MSBuild's `/p:` switches)

```powershell
# Build
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" VSTO2\OutlookAI.sln /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo

# Test (a filter, or no /TestCaseFilter for everything)
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll /TestCaseFilter:"FullyQualifiedName~GeometryTests"
```

"Build" and "Test X" in the steps below mean these two commands (X = the `FullyQualifiedName~` filter). The baseline before Task 1 is **908/908**.

## Review Focus

1. **Corners in a mirrored or self-crossing order** (e.g. TL, BL, BR, TR): a person expects the page never to come out mirrored; such an outline must be rejected and the image used uncropped. → `GeometryTests.Quad_MirroredOrder_IsRejected`, `Quad_SelfCrossingOrder_IsRejected`, `PagePlanTests.FromJson_MirroredCorners_UseTheFullImage` (Tasks 1–2).
2. **A transparent PNG** (screenshot, scanned signature page with alpha): transparent areas must come out white, not black. → `ImageLoaderTests.TransparentPng_ComesOutOnWhite` (Task 4).
3. **An AI answer wrapped in a code fence or a sentence** despite the strict schema: it should still be read. → `PagePlanTests.FromJson_ToleratesAFenceAroundTheJson` (Task 2).
4. **Saving over a PDF that's open in a viewer** (locked file): a clear message, and the review window stays open so nothing is lost. → `ImagesToPdfOutputTests.Publish_LockedDestination_SaysCloseIt` (Task 8), `ImagesToPdfReviewFormTests.Save_StaysOpen_WhenSavingFails` (Task 12).
5. **Cancel in the middle of a run**: no further AI calls, no review window, an `OperationCanceledException` the command catches, and the temp folder deleted. → `PagePlannerTests.PlanAsync_Cancelled_MakesNoCall`, `ImagesToPdfPipelineTests.RunAsync_Cancelled_Throws` (Tasks 10–11), `ImagesToPdfCommand_SourceTests` (Task 13).

---

### Task 1: Geometry (points, quadrilateral checks, homography)

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/Geometry.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/Homography.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj` (Compile entries)
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/GeometryTests.cs`

**Interfaces:**
- Produces: `struct PointD(double x, double y) { double X; double Y; }`; `static class Quad { double Area(PointD[] q); bool IsConvexClockwise(PointD[] q); }` (q = TL, TR, BR, BL); `sealed class Homography { static Homography FromPoints(PointD[] from, PointD[] to); PointD Map(double x, double y); double[] Coefficients /* 9 values, row-major, [8] == 1 */ }`. All `internal`, namespace `OutlookAI.Services.QuickTools`.

- [ ] **Step 1: Branch check**

Run: `git branch --show-current`
Expected: `feat/quick-tools-images-to-pdf` (the spec commits are already on it).

- [ ] **Step 2: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/GeometryTests.cs`:

```csharp
using System;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class GeometryTests
    {
        private static PointD[] P(params double[] xy)
        {
            var points = new PointD[xy.Length / 2];
            for (var i = 0; i < points.Length; i++) points[i] = new PointD(xy[2 * i], xy[2 * i + 1]);
            return points;
        }

        [Fact]
        public void Homography_MapsTheFourPointsExactly()
        {
            var from = P(0, 0, 850, 0, 850, 1100, 0, 1100);
            var to = P(470, 140, 1150, 190, 1290, 1080, 330, 1010);

            var map = Homography.FromPoints(from, to);

            for (var i = 0; i < 4; i++)
            {
                var mapped = map.Map(from[i].X, from[i].Y);
                Assert.Equal(to[i].X, mapped.X, 6);
                Assert.Equal(to[i].Y, mapped.Y, 6);
            }
        }

        [Fact]
        public void Homography_ThereAndBack_IsTheIdentity()
        {
            var square = P(0, 0, 1, 0, 1, 1, 0, 1);
            var trapezoid = P(0.29, 0.12, 0.72, 0.16, 0.81, 0.9, 0.21, 0.84);
            var there = Homography.FromPoints(square, trapezoid);
            var back = Homography.FromPoints(trapezoid, square);

            var mid = there.Map(0.3, 0.6);
            var round = back.Map(mid.X, mid.Y);

            Assert.Equal(0.3, round.X, 9);
            Assert.Equal(0.6, round.Y, 9);
        }

        [Fact]
        public void Homography_RejectsCollinearPoints()
        {
            Assert.Throws<InvalidOperationException>(
                () => Homography.FromPoints(P(0, 0, 1, 1, 2, 2, 3, 3), P(0, 0, 1, 0, 1, 1, 0, 1)));
        }

        [Fact]
        public void Quad_ReadingOrderOutline_IsConvexClockwise()
        {
            Assert.True(Quad.IsConvexClockwise(P(0.29, 0.12, 0.72, 0.16, 0.81, 0.9, 0.21, 0.84)));
        }

        // TL, BL, BR, TR would produce a mirrored page.
        [Fact]
        public void Quad_MirroredOrder_IsRejected()
        {
            Assert.False(Quad.IsConvexClockwise(P(0, 0, 0, 1, 1, 1, 1, 0)));
        }

        [Fact]
        public void Quad_SelfCrossingOrder_IsRejected()
        {
            Assert.False(Quad.IsConvexClockwise(P(0, 0, 1, 1, 1, 0, 0, 1)));
        }

        [Fact]
        public void Quad_AreaOfTheUnitSquare_IsOne()
        {
            Assert.Equal(1.0, Quad.Area(P(0, 0, 1, 0, 1, 1, 0, 1)), 9);
        }
    }
}
```

- [ ] **Step 3: Run to verify they fail**

Run: Build
Expected: the test project fails to compile with `CS0246: The type or namespace name 'PointD' could not be found` (and `Homography`, `Quad`).

- [ ] **Step 4: Implement**

`VSTO2/OutlookAI/Services/QuickTools/Geometry.cs`:

```csharp
using System;

namespace OutlookAI.Services.QuickTools
{
    internal struct PointD
    {
        public PointD(double x, double y)
        {
            X = x;
            Y = y;
        }

        public double X { get; }
        public double Y { get; }
    }

    /// <summary>Checks on a page outline given as TL, TR, BR, BL.</summary>
    internal static class Quad
    {
        public static double Area(PointD[] q)
        {
            double sum = 0;
            for (var i = 0; i < 4; i++)
            {
                var a = q[i];
                var b = q[(i + 1) % 4];
                sum += a.X * b.Y - b.X * a.Y;
            }
            return Math.Abs(sum) / 2;
        }

        /// <summary>
        /// True when the outline turns the same way at every corner, clockwise on
        /// screen (y grows downward), as TL, TR, BR, BL in reading order does. A
        /// mirrored or self-crossing order fails, so it can't flip the page.
        /// </summary>
        public static bool IsConvexClockwise(PointD[] q)
        {
            for (var i = 0; i < 4; i++)
            {
                var a = q[i];
                var b = q[(i + 1) % 4];
                var c = q[(i + 2) % 4];
                var cross = (b.X - a.X) * (c.Y - b.Y) - (b.Y - a.Y) * (c.X - b.X);
                if (cross <= 0) return false;
            }
            return true;
        }
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/Homography.cs`:

```csharp
using System;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// The projective map that sends four points onto four others, e.g. an
    /// output rectangle onto a photographed page's corners.
    /// </summary>
    internal sealed class Homography
    {
        private readonly double[] _h;

        private Homography(double[] h)
        {
            _h = h;
        }

        /// <summary>
        /// a..h and 1, row-major: x' = (a x + b y + c) / (g x + h y + 1),
        /// y' = (d x + e y + f) / (g x + h y + 1).
        /// </summary>
        public double[] Coefficients => (double[])_h.Clone();

        public static Homography FromPoints(PointD[] from, PointD[] to)
        {
            if (from == null || to == null || from.Length != 4 || to.Length != 4)
                throw new ArgumentException("Four point pairs are needed.");

            var a = new double[8, 9];
            for (var i = 0; i < 4; i++)
            {
                double x = from[i].X, y = from[i].Y, u = to[i].X, v = to[i].Y;
                var r = 2 * i;
                a[r, 0] = x; a[r, 1] = y; a[r, 2] = 1; a[r, 6] = -u * x; a[r, 7] = -u * y; a[r, 8] = u;
                a[r + 1, 3] = x; a[r + 1, 4] = y; a[r + 1, 5] = 1; a[r + 1, 6] = -v * x; a[r + 1, 7] = -v * y; a[r + 1, 8] = v;
            }

            // Gauss-Jordan elimination with partial pivoting; column 8 is the right-hand side.
            for (var col = 0; col < 8; col++)
            {
                var pivot = col;
                for (var row = col + 1; row < 8; row++)
                {
                    if (Math.Abs(a[row, col]) > Math.Abs(a[pivot, col])) pivot = row;
                }
                if (Math.Abs(a[pivot, col]) < 1e-12)
                    throw new InvalidOperationException("The points don't define a projective map.");
                if (pivot != col)
                {
                    for (var k = 0; k < 9; k++)
                    {
                        var t = a[col, k];
                        a[col, k] = a[pivot, k];
                        a[pivot, k] = t;
                    }
                }
                var p = a[col, col];
                for (var k = col; k < 9; k++) a[col, k] /= p;
                for (var row = 0; row < 8; row++)
                {
                    if (row == col) continue;
                    var f = a[row, col];
                    if (f == 0) continue;
                    for (var k = col; k < 9; k++) a[row, k] -= f * a[col, k];
                }
            }

            var h = new double[9];
            for (var i = 0; i < 8; i++) h[i] = a[i, 8];
            h[8] = 1;
            return new Homography(h);
        }

        public PointD Map(double x, double y)
        {
            var w = _h[6] * x + _h[7] * y + _h[8];
            return new PointD((_h[0] * x + _h[1] * y + _h[2]) / w, (_h[3] * x + _h[4] * y + _h[5]) / w);
        }
    }
}
```

In `VSTO2/OutlookAI/OutlookAI.csproj`, after `<Compile Include="Services\OutlookToolHost.cs" />` add:

```xml
    <Compile Include="Services\QuickTools\Geometry.cs" />
    <Compile Include="Services\QuickTools\Homography.cs" />
```

- [ ] **Step 5: Run to verify they pass**

Run: Build; Test `GeometryTests`
Expected: 7 passed.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools
git commit -m "feat(quick-tools): homography and page outline checks" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 2: PagePlan (the AI's answer, validated)

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/PagePlan.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/PagePlanTests.cs`

**Interfaces:**
- Consumes: `PointD`, `Quad` (Task 1).
- Produces: `enum PageKind { DocumentPage, NotDocument }`, `enum PageColor { Grayscale, Color }`, `enum CleanupStrength { Light, Normal, Strong }`; `sealed class PagePlan { const string SchemaName = "page_plan"; PageKind Kind; string Reason; PointD[] Corners /* fractions, TL TR BR BL */; PageColor Color; CleanupStrength Strength; bool Analyzed; static PointD[] FullImage; static PagePlan Fallback(); static PagePlan FromJson(string text); static JObject Schema(); }`.

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/PagePlanTests.cs`:

```csharp
using System.Linq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PagePlanTests
    {
        private const string Keystone =
            "[[0.294,0.117],[0.719,0.158],[0.806,0.9],[0.206,0.842]]";

        private static string Json(string corners, string kind = "document_page", string color = "grayscale", string strength = "normal") =>
            "{\"kind\":\"" + kind + "\",\"reason\":\" lease page 1 \",\"corners\":" + corners
            + ",\"color\":\"" + color + "\",\"strength\":\"" + strength + "\"}";

        [Fact]
        public void FromJson_ValidPlan_IsAnalyzed()
        {
            var plan = PagePlan.FromJson(Json(Keystone, color: "color", strength: "strong"));

            Assert.True(plan.Analyzed);
            Assert.Equal(PageKind.DocumentPage, plan.Kind);
            Assert.Equal("lease page 1", plan.Reason);
            Assert.Equal(PageColor.Color, plan.Color);
            Assert.Equal(CleanupStrength.Strong, plan.Strength);
            Assert.Equal(0.294, plan.Corners[0].X, 9);
            Assert.Equal(0.842, plan.Corners[3].Y, 9);
        }

        [Fact]
        public void FromJson_SlightlyOutside_IsClamped()
        {
            var plan = PagePlan.FromJson(Json("[[-0.01,0],[1.01,0],[1,1],[0,1]]"));

            Assert.True(plan.Analyzed);
            Assert.Equal(0.0, plan.Corners[0].X, 9);
            Assert.Equal(1.0, plan.Corners[1].X, 9);
        }

        [Fact]
        public void FromJson_OutOfRange_UsesTheFullImage_AndKeepsTheRest()
        {
            var plan = PagePlan.FromJson(Json("[[0,0],[1.2,0],[1,1],[0,1]]", color: "color"));

            Assert.False(plan.Analyzed);
            Assert.Equal(PageColor.Color, plan.Color);
            Assert.Equal(PagePlan.FullImage.Select(p => p.X), plan.Corners.Select(p => p.X));
        }

        [Fact]
        public void FromJson_MirroredCorners_UseTheFullImage()
        {
            var plan = PagePlan.FromJson(Json("[[0,0],[0,1],[1,1],[1,0]]"));

            Assert.False(plan.Analyzed);
            Assert.Equal(1.0, plan.Corners[1].X, 9);
            Assert.Equal(0.0, plan.Corners[1].Y, 9);
        }

        [Fact]
        public void FromJson_TinyOutline_UsesTheFullImage()
        {
            var plan = PagePlan.FromJson(Json("[[0.4,0.4],[0.5,0.4],[0.5,0.5],[0.4,0.5]]"));

            Assert.False(plan.Analyzed);
        }

        [Fact]
        public void FromJson_NotADocument_KeepsItsReason()
        {
            var plan = PagePlan.FromJson(Json("[[0,0],[1,0],[1,1],[0,1]]", kind: "not_document"));

            Assert.True(plan.Analyzed);
            Assert.Equal(PageKind.NotDocument, plan.Kind);
            Assert.Equal("lease page 1", plan.Reason);
        }

        [Theory]
        [InlineData("not json")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("{\"kind\":\"maybe\",\"reason\":\"\",\"corners\":[[0,0],[1,0],[1,1],[0,1]],\"color\":\"grayscale\",\"strength\":\"normal\"}")]
        [InlineData("{\"kind\":\"document_page\",\"reason\":\"\",\"corners\":[[0,0],[1,0],[1,1]],\"color\":\"grayscale\",\"strength\":\"normal\"}")]
        public void FromJson_Unusable_IsTheFallback(string text)
        {
            var plan = PagePlan.FromJson(text);

            Assert.False(plan.Analyzed);
            Assert.Equal(PageKind.DocumentPage, plan.Kind);
            Assert.Equal(PageColor.Grayscale, plan.Color);
            Assert.Equal(CleanupStrength.Normal, plan.Strength);
        }

        [Fact]
        public void FromJson_ToleratesAFenceAroundTheJson()
        {
            var plan = PagePlan.FromJson("Here you go:\n```json\n" + Json(Keystone) + "\n```");

            Assert.True(plan.Analyzed);
        }

        // Strict structured outputs need every property listed and nothing extra.
        [Fact]
        public void Schema_IsStrictFriendly()
        {
            var schema = PagePlan.Schema();

            Assert.False((bool)schema["additionalProperties"]);
            var required = ((JArray)schema["required"]).Select(t => (string)t).OrderBy(s => s);
            var properties = ((JObject)schema["properties"]).Properties().Select(p => p.Name).OrderBy(s => s);
            Assert.Equal(properties, required);
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile error `CS0103: The name 'PagePlan' does not exist` / `CS0246 PageKind`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/QuickTools/PagePlan.cs`:

```csharp
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.QuickTools
{
    internal enum PageKind { DocumentPage, NotDocument }

    internal enum PageColor { Grayscale, Color }

    internal enum CleanupStrength { Light, Normal, Strong }

    /// <summary>
    /// What the AI decided about one image, after validation. Corners are
    /// fractions of the image width and height: TL, TR, BR, BL as the text reads.
    /// </summary>
    internal sealed class PagePlan
    {
        public const string SchemaName = "page_plan";
        private const double Slack = 0.02;
        private const double MinAreaFraction = 0.05;

        private const string SchemaJson =
            "{\"type\":\"object\",\"additionalProperties\":false,"
            + "\"required\":[\"kind\",\"reason\",\"corners\",\"color\",\"strength\"],"
            + "\"properties\":{"
            + "\"kind\":{\"type\":\"string\",\"enum\":[\"document_page\",\"not_document\"]},"
            + "\"reason\":{\"type\":\"string\"},"
            + "\"corners\":{\"type\":\"array\",\"minItems\":4,\"maxItems\":4,\"items\":"
            + "{\"type\":\"array\",\"minItems\":2,\"maxItems\":2,\"items\":{\"type\":\"number\"}}},"
            + "\"color\":{\"type\":\"string\",\"enum\":[\"grayscale\",\"color\"]},"
            + "\"strength\":{\"type\":\"string\",\"enum\":[\"light\",\"normal\",\"strong\"]}}}";

        private PagePlan()
        {
        }

        public PageKind Kind { get; private set; }
        public string Reason { get; private set; }
        public PointD[] Corners { get; private set; }
        public PageColor Color { get; private set; }
        public CleanupStrength Strength { get; private set; }

        /// <summary>False when the AI failed or its outline was implausible; the image is then used uncropped.</summary>
        public bool Analyzed { get; private set; }

        public static PointD[] FullImage =>
            new[] { new PointD(0, 0), new PointD(1, 0), new PointD(1, 1), new PointD(0, 1) };

        public static JObject Schema() => JObject.Parse(SchemaJson);

        public static PagePlan Fallback() => new PagePlan
        {
            Kind = PageKind.DocumentPage,
            Reason = "",
            Corners = FullImage,
            Color = PageColor.Grayscale,
            Strength = CleanupStrength.Normal,
            Analyzed = false,
        };

        public static PagePlan FromJson(string text)
        {
            var o = ParseObject(text);
            if (o == null) return Fallback();

            PageKind kind;
            PageColor color;
            CleanupStrength strength;
            if (!TryKind((string)o["kind"], out kind)
                || !TryColor((string)o["color"], out color)
                || !TryStrength((string)o["strength"], out strength))
            {
                return Fallback();
            }

            var plan = new PagePlan
            {
                Kind = kind,
                Reason = ((string)o["reason"] ?? "").Trim(),
                Color = color,
                Strength = strength,
            };
            var corners = ReadCorners(o["corners"] as JArray);
            if (corners != null && Quad.IsConvexClockwise(corners) && Quad.Area(corners) >= MinAreaFraction)
            {
                plan.Corners = corners;
                plan.Analyzed = true;
            }
            else
            {
                plan.Corners = FullImage;
                plan.Analyzed = false;
            }
            return plan;
        }

        // Strict output is plain JSON; a fence or a sentence around it is tolerated.
        private static JObject ParseObject(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var start = text.IndexOf('{');
            var end = text.LastIndexOf('}');
            if (start < 0 || end <= start) return null;
            try
            {
                return JObject.Parse(text.Substring(start, end - start + 1));
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static PointD[] ReadCorners(JArray array)
        {
            if (array == null || array.Count != 4) return null;
            var corners = new PointD[4];
            for (var i = 0; i < 4; i++)
            {
                var pair = array[i] as JArray;
                if (pair == null || pair.Count != 2) return null;
                double x, y;
                if (!TryNumber(pair[0], out x) || !TryNumber(pair[1], out y)) return null;
                if (x < -Slack || x > 1 + Slack || y < -Slack || y > 1 + Slack) return null;
                corners[i] = new PointD(Clamp(x), Clamp(y));
            }
            return corners;
        }

        private static bool TryNumber(JToken token, out double value)
        {
            value = 0;
            if (token == null || (token.Type != JTokenType.Float && token.Type != JTokenType.Integer)) return false;
            value = token.Value<double>();
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static double Clamp(double v) => v < 0 ? 0 : v > 1 ? 1 : v;

        private static bool TryKind(string s, out PageKind kind)
        {
            kind = PageKind.DocumentPage;
            if (s == "document_page") return true;
            if (s == "not_document") { kind = PageKind.NotDocument; return true; }
            return false;
        }

        private static bool TryColor(string s, out PageColor color)
        {
            color = PageColor.Grayscale;
            if (s == "grayscale") return true;
            if (s == "color") { color = PageColor.Color; return true; }
            return false;
        }

        private static bool TryStrength(string s, out CleanupStrength strength)
        {
            strength = CleanupStrength.Normal;
            if (s == "normal") return true;
            if (s == "light") { strength = CleanupStrength.Light; return true; }
            if (s == "strong") { strength = CleanupStrength.Strong; return true; }
            return false;
        }
    }
}
```

`OutlookAI.csproj`, after `<Compile Include="Services\QuickTools\Homography.cs" />`:

```xml
    <Compile Include="Services\QuickTools\PagePlan.cs" />
```

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `PagePlanTests`
Expected: 13 passed (8 facts + 5 theory cases).

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools/PagePlan.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools/PagePlanTests.cs
git commit -m "feat(quick-tools): page plan schema and validation" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 3: Structured (image + schema) call on the Codex backend

**Files:**
- Create: `VSTO2/OutlookAI/Services/CodexBackendException.cs`
- Create: `VSTO2/OutlookAI/Services/IStructuredCompletion.cs`
- Modify: `VSTO2/OutlookAI/Services/CodexChatService.cs` (extract `SendResponsesAsync` from `ProcessEmailAsync`; add `CompleteStructuredAsync`; implement `IStructuredCompletion`)
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/CodexChatServiceTests.cs`

**Interfaces:**
- Produces: `public sealed class CodexBackendException : InvalidOperationException { HttpStatusCode StatusCode; TimeSpan? RetryAfter; }` (namespace `OutlookAI.Services`); `internal interface IStructuredCompletion { Task<string> CompleteStructuredAsync(string instructions, IReadOnlyList<JObject> contentParts, string schemaName, JObject schema, string effort, CancellationToken cancellationToken); }`; `CodexChatService : IStructuredCompletion`. `ProcessEmailAsync` now throws `CodexBackendException` (still an `InvalidOperationException`) on HTTP errors.

- [ ] **Step 1: Write the failing tests**

Add to `CodexChatServiceTests` (it is already `[Collection("Config")]` with `ConfigStateScope` and a fresh token in `_authPath`). Add `using System.Collections.Generic;`, `using System.Net.Http.Headers;` and `using Newtonsoft.Json.Linq;` to the file.

```csharp
        private static IReadOnlyList<JObject> TextAndImage() => new List<JObject>
        {
            new JObject(new JProperty("type", "input_text"), new JProperty("text", "Plan the cleanup of this image.")),
            new JObject(new JProperty("type", "input_image"), new JProperty("image_url", "data:image/jpeg;base64,AAAA")),
        };

        private static JObject Schema() => JObject.Parse("{\"type\":\"object\",\"additionalProperties\":false,\"required\":[\"ok\"],\"properties\":{\"ok\":{\"type\":\"boolean\"}}}");

        [Fact]
        public async Task CompleteStructuredAsync_SendsTheImageAndAStrictSchema()
        {
            Config.Model = "gpt-5.6-sol";
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.done\",\"text\":\"{\\\"ok\\\":true}\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var text = await chat.CompleteStructuredAsync("instructions here", TextAndImage(), "page_plan", Schema(), "Low", CancellationToken.None);

                Assert.Equal("{\"ok\":true}", text);
                var body = JObject.Parse(fake.RequestBodies[0]);
                Assert.Equal("gpt-5.6-sol", (string)body["model"]);
                Assert.Equal("instructions here", (string)body["instructions"]);
                var content = (JArray)body["input"][0]["content"];
                Assert.Equal("input_image", (string)content[1]["type"]);
                Assert.StartsWith("data:image/jpeg;base64,", (string)content[1]["image_url"]);
                var format = body["text"]["format"];
                Assert.Equal("json_schema", (string)format["type"]);
                Assert.Equal("page_plan", (string)format["name"]);
                Assert.True((bool)format["strict"]);
                Assert.True(JToken.DeepEquals(Schema(), format["schema"]));
                Assert.Equal("low", (string)body["reasoning"]["effort"]);
                Assert.False((bool)body["store"]);
                Assert.True((bool)body["stream"]);
                Assert.Null(body["tools"]);
            }
        }

        [Fact]
        public async Task CompleteStructuredAsync_OmitsAnEffortTheModelDoesntTake()
        {
            Config.Model = "gpt-5.5";
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK, "data: {\"type\":\"response.output_text.done\",\"text\":\"{}\"}\n\n");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                await chat.CompleteStructuredAsync("i", TextAndImage(), "page_plan", Schema(), "Ultra", CancellationToken.None);

                Assert.Equal(JTokenType.Null, JObject.Parse(fake.RequestBodies[0])["reasoning"].Type);
            }
        }

        [Fact]
        public async Task CompleteStructuredAsync_RateLimited_ThrowsWithStatusAndRetryAfter()
        {
            var fake = new FakeHttpMessageHandler();
            fake.Queue(_ =>
            {
                var response = new HttpResponseMessage((HttpStatusCode)429) { Content = new StringContent("{\"detail\":\"slow down\"}") };
                response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(7));
                return response;
            });
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var ex = await Assert.ThrowsAsync<CodexBackendException>(
                    () => chat.CompleteStructuredAsync("i", TextAndImage(), "page_plan", Schema(), "Low", CancellationToken.None));

                Assert.Equal(429, (int)ex.StatusCode);
                Assert.Equal(TimeSpan.FromSeconds(7), ex.RetryAfter);
                Assert.Contains("slow down", ex.Message);
            }
        }
```

Change the existing `ProcessEmailAsync_ThrowsWithBackendErrorBody_OnNonSuccess` to expect the typed exception (it is still an `InvalidOperationException`, so callers catching that keep working):

```csharp
                var ex = await Assert.ThrowsAsync<CodexBackendException>(
                    () => chat.ProcessEmailAsync(
                        CodexChatService.ActionType.Proofread, "x"));

                Assert.Equal(429, (int)ex.StatusCode);
                Assert.Contains("ChatGPT Codex backend error", ex.Message);
                Assert.Contains("rate limited", ex.Message);
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile errors `CS1061 'CodexChatService' does not contain a definition for 'CompleteStructuredAsync'` and `CS0246 'CodexBackendException'`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/CodexBackendException.cs`:

```csharp
using System;
using System.Net;

namespace OutlookAI.Services
{
    /// <summary>The Codex backend answered with an HTTP error.</summary>
    public sealed class CodexBackendException : InvalidOperationException
    {
        public CodexBackendException(HttpStatusCode statusCode, TimeSpan? retryAfter, string message)
            : base(message)
        {
            StatusCode = statusCode;
            RetryAfter = retryAfter;
        }

        public HttpStatusCode StatusCode { get; }

        /// <summary>The server's Retry-After, when it sent one.</summary>
        public TimeSpan? RetryAfter { get; }
    }
}
```

`VSTO2/OutlookAI/Services/IStructuredCompletion.cs`:

```csharp
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services
{
    /// <summary>One answer held to a JSON schema, from text and image parts.</summary>
    internal interface IStructuredCompletion
    {
        Task<string> CompleteStructuredAsync(
            string instructions,
            IReadOnlyList<JObject> contentParts,
            string schemaName,
            JObject schema,
            string effort,
            CancellationToken cancellationToken);
    }
}
```

In `CodexChatService.cs`, change the class declaration to:

```csharp
    public sealed class CodexChatService : IDisposable, IStructuredCompletion
```

Replace the body of `ProcessEmailAsync` (everything after `var body = BuildResponsesRequest(instructions, userMessage);`) with:

```csharp
            return await SendResponsesAsync(body, cancellationToken).ConfigureAwait(false);
        }
```

and add, next to `BuildResponsesRequest`:

```csharp
        /// <summary>
        /// One answer held to a strict JSON schema, from text and image parts
        /// (Quick Tools). Uses the model in Settings; an effort the model
        /// doesn't take is omitted.
        /// </summary>
        public async Task<string> CompleteStructuredAsync(
            string instructions,
            IReadOnlyList<JObject> contentParts,
            string schemaName,
            JObject schema,
            string effort,
            CancellationToken cancellationToken)
        {
            var catalog = Config.ModelCatalog;
            var model = catalog.ResolveEffectiveModel(Config.Model, Config.Clock());
            var wireEffort = catalog.ResolveWireEffort(model, effort);
            var body = new JObject(
                new JProperty("model", model),
                new JProperty("instructions", instructions ?? ""),
                new JProperty("input", new JArray(
                    new JObject(
                        new JProperty("type", "message"),
                        new JProperty("role", "user"),
                        new JProperty("content", new JArray(contentParts))))),
                new JProperty("text", new JObject(
                    new JProperty("format", new JObject(
                        new JProperty("type", "json_schema"),
                        new JProperty("name", schemaName),
                        new JProperty("strict", true),
                        new JProperty("schema", schema))))),
                new JProperty("reasoning", wireEffort != null
                    ? (JToken)new JObject(new JProperty("effort", wireEffort))
                    : JValue.CreateNull()),
                new JProperty("store", false),
                new JProperty("stream", true),
                new JProperty("include", new JArray()));
            return await SendResponsesAsync(body, cancellationToken).ConfigureAwait(false);
        }

        private async Task<string> SendResponsesAsync(JObject body, CancellationToken cancellationToken)
        {
            using (var request = new HttpRequestMessage(HttpMethod.Post, ResponsesEndpoint))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue(
                    "Bearer",
                    await _auth.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false));
                request.Headers.Accept.ParseAdd("text/event-stream");

                var status = _auth.GetStatus();
                if (status.State == AuthState.Authenticated && !string.IsNullOrEmpty(status.AccountId))
                {
                    request.Headers.TryAddWithoutValidation("ChatGPT-Account-ID", status.AccountId);
                }

                request.Content = new StringContent(body.ToString(Newtonsoft.Json.Formatting.None), Encoding.UTF8, "application/json");

                using (var response = await _http.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        var errorBody = await SafeReadAsStringAsync(response).ConfigureAwait(false);
                        throw new CodexBackendException(
                            response.StatusCode,
                            RetryAfterOf(response),
                            "ChatGPT Codex backend error: " + (int)response.StatusCode + " " + errorBody);
                    }

                    using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    {
                        return await ReadResponsesSseAsync(stream, cancellationToken).ConfigureAwait(false);
                    }
                }
            }
        }

        private static TimeSpan? RetryAfterOf(HttpResponseMessage response)
        {
            var retry = response.Headers.RetryAfter;
            if (retry == null) return null;
            if (retry.Delta.HasValue) return retry.Delta;
            if (retry.Date.HasValue) return retry.Date.Value - DateTimeOffset.UtcNow;
            return null;
        }
```

`OutlookAI.csproj`: after `<Compile Include="Services\CodexAuthService.cs" />` add `<Compile Include="Services\CodexBackendException.cs" />`; after `<Compile Include="Services\IdResolver.cs" />` add `<Compile Include="Services\IStructuredCompletion.cs" />` (check the exact neighbor lines with `Select-String Services\\I OutlookAI.csproj` and keep alphabetical order).

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `CodexChatServiceTests`
Expected: all pass (the 3 new + the existing ones, including the changed one).

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/CodexChatServiceTests.cs
git commit -m "feat(quick-tools): structured image call on the Codex backend" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 4: Imaging basics (pixel buffer, JPEG files, image loading)

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/PixelBuffer.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/JpegFiles.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/ImageLoader.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj` (Compile entries; References `PresentationCore`, `WindowsBase`, `System.Xaml`)
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/ImageLoaderTests.cs`

**Interfaces:**
- Produces: `sealed class PixelBuffer { PixelBuffer(int width, int height); int Width; int Height; byte[] Data /* BGR, rows packed */; static PixelBuffer FromBitmap(Bitmap); Bitmap ToBitmap(); PixelBuffer Clone(); }`; `static class JpegFiles { const long PageQuality = 85, OriginalQuality = 95; void Save(Image, string path, long quality); byte[] Encode(Image, long quality); Bitmap Load(string path); Bitmap Scaled(Image, int maxEdge); Bitmap LoadScaled(string path, int maxEdge); }`; `sealed class UnreadableImageException : Exception`; `static class ImageLoader { List<Bitmap> LoadPages(string path); internal List<Bitmap> LoadWithWic(string path); internal RotateFlipType Upright(int exifOrientation); }`.

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/ImageLoaderTests.cs`:

```csharp
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class ImageLoaderTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-imageloader", Path.GetRandomFileName());

        public ImageLoaderTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        // 40 x 20, left half red, right half blue.
        private static Bitmap Sample()
        {
            var bitmap = new Bitmap(40, 20, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.FillRectangle(Brushes.Red, 0, 0, 20, 20);
                g.FillRectangle(Brushes.Blue, 20, 0, 20, 20);
            }
            return bitmap;
        }

        private string SaveWithOrientation(int orientation)
        {
            var path = Path.Combine(_dir, "o" + orientation + ".jpg");
            using (var bitmap = Sample())
            {
                var item = (PropertyItem)FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                item.Id = 0x0112;
                item.Type = 3;
                item.Len = 2;
                item.Value = BitConverter.GetBytes((ushort)orientation);
                bitmap.SetPropertyItem(item);
                JpegFiles.Save(bitmap, path, 95);
            }
            return path;
        }

        [Theory]
        [InlineData(1, RotateFlipType.RotateNoneFlipNone)]
        [InlineData(2, RotateFlipType.RotateNoneFlipX)]
        [InlineData(3, RotateFlipType.Rotate180FlipNone)]
        [InlineData(4, RotateFlipType.Rotate180FlipX)]
        [InlineData(5, RotateFlipType.Rotate90FlipX)]
        [InlineData(6, RotateFlipType.Rotate90FlipNone)]
        [InlineData(7, RotateFlipType.Rotate270FlipX)]
        [InlineData(8, RotateFlipType.Rotate270FlipNone)]
        [InlineData(0, RotateFlipType.RotateNoneFlipNone)]
        public void Upright_MapsEveryExifOrientation(int orientation, RotateFlipType expected)
        {
            Assert.Equal(expected, ImageLoader.Upright(orientation));
        }

        [Fact]
        public void LoadPages_TurnsAPhoneShotUpright()
        {
            var pages = ImageLoader.LoadPages(SaveWithOrientation(6));
            try
            {
                var page = Assert.Single(pages);
                Assert.Equal(20, page.Width);
                Assert.Equal(40, page.Height);
                // Rotated 90° clockwise: the red left half is now on top.
                Assert.True(page.GetPixel(10, 5).R > 200);
                Assert.True(page.GetPixel(10, 35).B > 200);
            }
            finally { foreach (var p in pages) p.Dispose(); }
        }

        [Fact]
        public void LoadPages_MultiPageTiff_GivesEveryFrame()
        {
            var path = Path.Combine(_dir, "fax.tif");
            using (var first = Sample())
            using (var second = Sample())
            using (var third = Sample())
            {
                var tiff = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Tiff.Guid);
                using (var parameters = new EncoderParameters(1))
                {
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                    first.Save(path, tiff, parameters);
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                    first.SaveAdd(second, parameters);
                    first.SaveAdd(third, parameters);
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
                    first.SaveAdd(parameters);
                }
            }

            var pages = ImageLoader.LoadPages(path);
            try
            {
                Assert.Equal(3, pages.Count);
            }
            finally { foreach (var p in pages) p.Dispose(); }
        }

        [Fact]
        public void TransparentPng_ComesOutOnWhite()
        {
            var path = Path.Combine(_dir, "screenshot.png");
            using (var png = new Bitmap(30, 30, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(png)) g.Clear(Color.Transparent);
                png.SetPixel(15, 15, Color.Black);
                png.Save(path, ImageFormat.Png);
            }

            var pages = ImageLoader.LoadPages(path);
            try
            {
                var pixel = pages[0].GetPixel(2, 2);
                Assert.Equal(255, pixel.R);
                Assert.Equal(255, pixel.G);
                Assert.Equal(255, pixel.B);
            }
            finally { foreach (var p in pages) p.Dispose(); }
        }

        [Fact]
        public void LoadWithWic_DecodesWhatWindowsCan()
        {
            var path = Path.Combine(_dir, "plain.png");
            using (var bitmap = Sample()) bitmap.Save(path, ImageFormat.Png);

            var pages = ImageLoader.LoadWithWic(path);
            try
            {
                Assert.Equal(40, pages[0].Width);
                Assert.True(pages[0].GetPixel(5, 5).R > 200);
            }
            finally { foreach (var p in pages) p.Dispose(); }
        }

        [Fact]
        public void LoadPages_NotAnImage_IsUnreadable()
        {
            var path = Path.Combine(_dir, "fake.jpg");
            File.WriteAllText(path, "this is not an image");

            Assert.Throws<UnreadableImageException>(() => ImageLoader.LoadPages(path));
        }

        [Fact]
        public void JpegFiles_Scaled_CapsTheLongEdge_AndNeverEnlarges()
        {
            using (var big = new Bitmap(3200, 1600))
            using (var small = new Bitmap(300, 200))
            using (var a = JpegFiles.Scaled(big, 1600))
            using (var b = JpegFiles.Scaled(small, 1600))
            {
                Assert.Equal(new Size(1600, 800), a.Size);
                Assert.Equal(new Size(300, 200), b.Size);
            }
        }

        [Fact]
        public void JpegFiles_Load_DoesntKeepTheFileLocked()
        {
            var path = Path.Combine(_dir, "unlocked.jpg");
            using (var bitmap = Sample()) JpegFiles.Save(bitmap, path, 85);

            using (var loaded = JpegFiles.Load(path))
            {
                File.Delete(path);
                Assert.Equal(40, loaded.Width);
            }
        }

        [Fact]
        public void PixelBuffer_RoundTripsThroughABitmap()
        {
            using (var bitmap = Sample())
            {
                var buffer = PixelBuffer.FromBitmap(bitmap);
                Assert.Equal(new byte[] { 0, 0, 255 }, buffer.Data.Take(3).ToArray()); // BGR red
                using (var back = buffer.ToBitmap())
                {
                    Assert.Equal(Color.FromArgb(255, 0, 0, 255).ToArgb(), back.GetPixel(30, 10).ToArgb());
                }
            }
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile errors `CS0103 'ImageLoader'`, `'JpegFiles'`, `'PixelBuffer'`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/QuickTools/PixelBuffer.cs`:

```csharp
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>24-bit pixels in B, G, R order, rows packed without padding.</summary>
    internal sealed class PixelBuffer
    {
        public PixelBuffer(int width, int height)
        {
            if (width <= 0 || height <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            Width = width;
            Height = height;
            Data = new byte[checked(width * height * 3)];
        }

        public int Width { get; }
        public int Height { get; }
        public byte[] Data { get; }

        public static PixelBuffer FromBitmap(Bitmap bitmap)
        {
            var buffer = new PixelBuffer(bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            try
            {
                var row = buffer.Width * 3;
                for (var y = 0; y < buffer.Height; y++)
                {
                    Marshal.Copy(data.Scan0 + y * data.Stride, buffer.Data, y * row, row);
                }
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
            return buffer;
        }

        public Bitmap ToBitmap()
        {
            var bitmap = new Bitmap(Width, Height, PixelFormat.Format24bppRgb);
            var data = bitmap.LockBits(new Rectangle(0, 0, Width, Height), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
            try
            {
                var row = Width * 3;
                for (var y = 0; y < Height; y++)
                {
                    Marshal.Copy(Data, y * row, data.Scan0 + y * data.Stride, row);
                }
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
            return bitmap;
        }

        public PixelBuffer Clone()
        {
            var copy = new PixelBuffer(Width, Height);
            Buffer.BlockCopy(Data, 0, copy.Data, 0, Data.Length);
            return copy;
        }
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/JpegFiles.cs`:

```csharp
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace OutlookAI.Services.QuickTools
{
    internal static class JpegFiles
    {
        public const long PageQuality = 85;
        public const long OriginalQuality = 95;

        private static readonly ImageCodecInfo Codec =
            ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);

        public static void Save(Image image, string path, long quality)
        {
            using (var parameters = Quality(quality))
            {
                image.Save(path, Codec, parameters);
            }
        }

        public static byte[] Encode(Image image, long quality)
        {
            using (var parameters = Quality(quality))
            using (var stream = new MemoryStream())
            {
                image.Save(stream, Codec, parameters);
                return stream.ToArray();
            }
        }

        /// <summary>Reads the file fully first, so it isn't kept locked.</summary>
        public static Bitmap Load(string path)
        {
            using (var stream = new MemoryStream(File.ReadAllBytes(path)))
            using (var image = Image.FromStream(stream))
            {
                return new Bitmap(image);
            }
        }

        /// <summary>A copy at most maxEdge pixels on its long side (never enlarged).</summary>
        public static Bitmap Scaled(Image image, int maxEdge)
        {
            var scale = Math.Min(1.0, (double)maxEdge / Math.Max(image.Width, image.Height));
            var width = Math.Max(1, (int)Math.Round(image.Width * scale));
            var height = Math.Max(1, (int)Math.Round(image.Height * scale));
            var copy = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(copy))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(image, new Rectangle(0, 0, width, height));
            }
            return copy;
        }

        public static Bitmap LoadScaled(string path, int maxEdge)
        {
            using (var full = Load(path))
            {
                return Scaled(full, maxEdge);
            }
        }

        private static EncoderParameters Quality(long quality)
        {
            var parameters = new EncoderParameters(1);
            parameters.Param[0] = new EncoderParameter(Encoder.Quality, quality);
            return parameters;
        }
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/ImageLoader.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class UnreadableImageException : Exception
    {
        public UnreadableImageException(string message, Exception inner) : base(message, inner)
        {
        }
    }

    /// <summary>
    /// Decodes an image file into upright 24-bit pages on white: EXIF orientation
    /// applied, one page per TIFF frame, a GIF's first frame. GDI+ first; WIC
    /// (WebP, HEIC, … when Windows has the codec) when GDI+ can't.
    /// </summary>
    internal static class ImageLoader
    {
        private const int ExifOrientation = 0x0112;

        /// <summary>Caller disposes the bitmaps.</summary>
        public static List<Bitmap> LoadPages(string path)
        {
            try
            {
                return LoadWithGdi(path);
            }
            catch (Exception gdiError) when (gdiError is ArgumentException || gdiError is ExternalException)
            {
                try
                {
                    return LoadWithWic(path);
                }
                catch (Exception wicError)
                {
                    throw new UnreadableImageException("This file couldn't be decoded.", wicError);
                }
            }
        }

        internal static RotateFlipType Upright(int orientation)
        {
            switch (orientation)
            {
                case 2: return RotateFlipType.RotateNoneFlipX;
                case 3: return RotateFlipType.Rotate180FlipNone;
                case 4: return RotateFlipType.Rotate180FlipX;
                case 5: return RotateFlipType.Rotate90FlipX;
                case 6: return RotateFlipType.Rotate90FlipNone;
                case 7: return RotateFlipType.Rotate270FlipX;
                case 8: return RotateFlipType.Rotate270FlipNone;
                default: return RotateFlipType.RotateNoneFlipNone;
            }
        }

        private static List<Bitmap> LoadWithGdi(string path)
        {
            var pages = new List<Bitmap>();
            using (var stream = new MemoryStream(File.ReadAllBytes(path)))
            using (var image = Image.FromStream(stream, false, true))
            {
                var orientation = ReadOrientation(image);
                var frames = image.RawFormat.Guid == ImageFormat.Tiff.Guid
                    ? image.GetFrameCount(FrameDimension.Page)
                    : 1;
                try
                {
                    for (var i = 0; i < frames; i++)
                    {
                        if (frames > 1) image.SelectActiveFrame(FrameDimension.Page, i);
                        var page = OnWhite(image);
                        page.RotateFlip(Upright(orientation));
                        pages.Add(page);
                    }
                }
                catch
                {
                    foreach (var p in pages) p.Dispose();
                    throw;
                }
            }
            return pages;
        }

        internal static List<Bitmap> LoadWithWic(string path)
        {
            var pages = new List<Bitmap>();
            using (var stream = new MemoryStream(File.ReadAllBytes(path)))
            {
                var decoder = System.Windows.Media.Imaging.BitmapDecoder.Create(
                    stream,
                    System.Windows.Media.Imaging.BitmapCreateOptions.IgnoreColorProfile,
                    System.Windows.Media.Imaging.BitmapCacheOption.OnLoad);
                try
                {
                    foreach (var frame in decoder.Frames)
                    {
                        var bgr = new System.Windows.Media.Imaging.FormatConvertedBitmap(
                            frame, System.Windows.Media.PixelFormats.Bgr24, null, 0);
                        var buffer = new PixelBuffer(bgr.PixelWidth, bgr.PixelHeight);
                        bgr.CopyPixels(buffer.Data, buffer.Width * 3, 0);
                        var page = buffer.ToBitmap();
                        page.RotateFlip(Upright(WicOrientation(frame)));
                        pages.Add(page);
                    }
                }
                catch
                {
                    foreach (var p in pages) p.Dispose();
                    throw;
                }
            }
            return pages;
        }

        private static int ReadOrientation(Image image)
        {
            if (Array.IndexOf(image.PropertyIdList, ExifOrientation) < 0) return 1;
            var value = image.GetPropertyItem(ExifOrientation).Value;
            return value != null && value.Length >= 2 ? BitConverter.ToUInt16(value, 0) : 1;
        }

        private static int WicOrientation(System.Windows.Media.Imaging.BitmapFrame frame)
        {
            try
            {
                var metadata = frame.Metadata as System.Windows.Media.Imaging.BitmapMetadata;
                var value = metadata == null ? null : metadata.GetQuery("System.Photo.Orientation");
                return value is ushort ? (ushort)value : 1;
            }
            catch (NotSupportedException)
            {
                return 1;
            }
            catch (InvalidOperationException)
            {
                return 1;
            }
        }

        // Normalizes any format (palette, alpha, CMYK, 16-bit) to 24-bit on white.
        private static Bitmap OnWhite(Image image)
        {
            var page = new Bitmap(image.Width, image.Height, PixelFormat.Format24bppRgb);
            page.SetResolution(96, 96);
            using (var g = Graphics.FromImage(page))
            {
                g.Clear(Color.White);
                g.InterpolationMode = InterpolationMode.NearestNeighbor;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                g.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height));
            }
            return page;
        }
    }
}
```

`OutlookAI.csproj`: after `<Compile Include="Services\QuickTools\Homography.cs" />` add:

```xml
    <Compile Include="Services\QuickTools\ImageLoader.cs" />
    <Compile Include="Services\QuickTools\JpegFiles.cs" />
```

after `<Compile Include="Services\QuickTools\PagePlan.cs" />` add:

```xml
    <Compile Include="Services\QuickTools\PixelBuffer.cs" />
```

and next to `<Reference Include="System.Xml.Linq" />` add:

```xml
    <Reference Include="PresentationCore" />
    <Reference Include="System.Xaml" />
    <Reference Include="WindowsBase" />
```

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `ImageLoaderTests`
Expected: all pass (9 theory cases + 8 facts). If `LoadPages_TurnsAPhoneShotUpright` fails because GDI+ didn't write the orientation tag, check with `Image.FromFile(path).PropertyIdList` that 0x0112 is present after `JpegFiles.Save`; GDI+ keeps property items set with `SetPropertyItem` when encoding JPEG.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools/ImageLoaderTests.cs
git commit -m "feat(quick-tools): load images upright and keep pages as JPEG files" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 5: Perspective warp with the page's true aspect ratio

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/PageAspect.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/PerspectiveWarp.cs`
- Create: `VSTO2/OutlookAI.Tests/Helpers/SyntheticPages.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Modify: `docs/superpowers/specs/2026-10-13-quick-tools-images-to-pdf-design.md` (§ 6.4 step 1)
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/PerspectiveWarpTests.cs`

**Interfaces:**
- Consumes: `PointD`, `Homography` (Task 1), `PixelBuffer` (Task 4).
- Produces: `static class PageAspect { double Estimate(PointD[] cornersPx, int imageWidth, int imageHeight); double EdgeRatio(PointD[] cornersPx); }` (width ÷ height); `static class PerspectiveWarp { const int MaxLongEdge = 3300, MaxShortEdge = 2550; Size OutputSize(PointD[] cornersPx, int imageWidth, int imageHeight); PixelBuffer Warp(PixelBuffer src, PointD[] cornersPx); }`; test helper `SyntheticPages { Color DimPaper, DimInk, Desk; Bitmap Page(int w, int h, Color paper, Color ink, bool marker = false, Color? stamp = null); Bitmap Photograph(Bitmap page, int canvasW, int canvasH, PointD[] cornersPx, Color background); PointD DarkCentroid(PixelBuffer px, int threshold); }` (marker: a black square, top-left at (0.1 W, 0.1 H), 0.1 W wide).

- [ ] **Step 1: Write the test helper**

`VSTO2/OutlookAI.Tests/Helpers/SyntheticPages.cs`:

```csharp
using System.Drawing;
using System.Drawing.Imaging;
using OutlookAI.Services.QuickTools;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>Fake document photos with known geometry (fake text only).</summary>
    internal static class SyntheticPages
    {
        public static readonly Color DimPaper = Color.FromArgb(150, 142, 110);
        public static readonly Color DimInk = Color.FromArgb(100, 94, 78);
        public static readonly Color Desk = Color.FromArgb(58, 46, 38);

        /// <summary>
        /// Text lines from 30% to ~88% of the height. With marker, a black square
        /// whose top-left is at (0.1 W, 0.1 H), 0.1 W wide. With stamp, a line of
        /// bold text in that color at (0.55 W, 0.12 H).
        /// </summary>
        public static Bitmap Page(int width, int height, Color paper, Color ink, bool marker = false, Color? stamp = null)
        {
            var page = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(page))
            using (var inkBrush = new SolidBrush(ink))
            using (var font = new Font("Times New Roman", height / 40f, GraphicsUnit.Pixel))
            {
                g.Clear(paper);
                if (marker) g.FillRectangle(Brushes.Black, width * 0.1f, height * 0.1f, width * 0.1f, width * 0.1f);
                if (stamp.HasValue)
                {
                    using (var stampBrush = new SolidBrush(stamp.Value))
                    using (var bold = new Font("Arial", height / 30f, FontStyle.Bold, GraphicsUnit.Pixel))
                    {
                        g.DrawString("RECEIVED OCT 13 2026", bold, stampBrush, width * 0.55f, height * 0.12f);
                    }
                }
                for (var line = 0; line < 12; line++)
                {
                    g.DrawString("Lease line " + (line + 1) + ": tenant John Example, 555-0187",
                        font, inkBrush, width * 0.1f, height * (0.3f + line * 0.05f));
                }
            }
            return page;
        }

        /// <summary>The page as if photographed: its corners at these canvas pixels, the rest background.</summary>
        public static Bitmap Photograph(Bitmap page, int canvasWidth, int canvasHeight, PointD[] cornersPx, Color background)
        {
            var src = PixelBuffer.FromBitmap(page);
            var toPage = Homography.FromPoints(cornersPx, new[]
            {
                new PointD(0, 0), new PointD(page.Width, 0), new PointD(page.Width, page.Height), new PointD(0, page.Height),
            });
            var dst = new PixelBuffer(canvasWidth, canvasHeight);
            for (var y = 0; y < canvasHeight; y++)
            {
                for (var x = 0; x < canvasWidth; x++)
                {
                    var p = toPage.Map(x + 0.5, y + 0.5);
                    var o = (y * canvasWidth + x) * 3;
                    if (p.X >= 0 && p.X < page.Width && p.Y >= 0 && p.Y < page.Height)
                    {
                        var i = ((int)p.Y * page.Width + (int)p.X) * 3;
                        dst.Data[o] = src.Data[i];
                        dst.Data[o + 1] = src.Data[i + 1];
                        dst.Data[o + 2] = src.Data[i + 2];
                    }
                    else
                    {
                        dst.Data[o] = background.B;
                        dst.Data[o + 1] = background.G;
                        dst.Data[o + 2] = background.R;
                    }
                }
            }
            return dst.ToBitmap();
        }

        /// <summary>Centroid of pixels darker than threshold (all channels), as fractions of the image.</summary>
        public static PointD DarkCentroid(PixelBuffer px, int threshold)
        {
            double sx = 0, sy = 0;
            long n = 0;
            for (var y = 0; y < px.Height; y++)
            {
                for (var x = 0; x < px.Width; x++)
                {
                    var i = (y * px.Width + x) * 3;
                    if (px.Data[i] < threshold && px.Data[i + 1] < threshold && px.Data[i + 2] < threshold)
                    {
                        sx += x;
                        sy += y;
                        n++;
                    }
                }
            }
            return n == 0 ? new PointD(double.NaN, double.NaN) : new PointD(sx / n / px.Width, sy / n / px.Height);
        }
    }
}
```

- [ ] **Step 2: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/PerspectiveWarpTests.cs`:

```csharp
using System;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PerspectiveWarpTests
    {
        // A 0.85 x 1.10 page (Letter, portrait) seen by a pinhole camera.
        private static PointD[] Projected(double ax, double ay, double az, double distance, int width, int height, double focal)
        {
            double cx = Math.Cos(ax), sx = Math.Sin(ax), cy = Math.Cos(ay), sy = Math.Sin(ay), cz = Math.Cos(az), sz = Math.Sin(az);
            double[,] r =
            {
                { cz * cy, cz * sy * sx - sz * cx, cz * sy * cx + sz * sx },
                { sz * cy, sz * sy * sx + cz * cx, sz * sy * cx - cz * sx },
                { -sy, cy * sx, cy * cx },
            };
            var corners = new[] { new[] { -0.425, -0.55 }, new[] { 0.425, -0.55 }, new[] { 0.425, 0.55 }, new[] { -0.425, 0.55 } };
            var result = new PointD[4];
            for (var i = 0; i < 4; i++)
            {
                double px = corners[i][0], py = corners[i][1];
                var x = r[0, 0] * px + r[0, 1] * py + 0.05;
                var y = r[1, 0] * px + r[1, 1] * py - 0.03;
                var z = r[2, 0] * px + r[2, 1] * py + distance;
                result[i] = new PointD(focal * x / z + width / 2.0, focal * y / z + height / 2.0);
            }
            return result;
        }

        [Theory]
        [InlineData(0.35, 0.2, 0.05, 3.0)]
        [InlineData(-0.5, 0.1, -0.1, 2.5)]
        [InlineData(0.6, -0.4, 0.2, 3.5)]
        public void PageAspect_RecoversTheTrueProportions_OfAnAngledPhoto(double ax, double ay, double az, double distance)
        {
            var corners = Projected(ax, ay, az, distance, 4032, 3024, 3000);

            Assert.Equal(0.85 / 1.10, PageAspect.Estimate(corners, 4032, 3024), 2);
            // The longer-edges rule is visibly off for these.
            Assert.True(Math.Abs(PageAspect.EdgeRatio(corners) - 0.85 / 1.10) > 0.04);
        }

        [Fact]
        public void PageAspect_NoPerspective_UsesTheEdges()
        {
            var corners = new[] { new PointD(100, 100), new PointD(950, 100), new PointD(950, 1200), new PointD(100, 1200) };

            Assert.Equal(850.0 / 1100.0, PageAspect.Estimate(corners, 1600, 1600), 6);
        }

        [Fact]
        public void PageAspect_ImplausibleCamera_UsesTheEdges()
        {
            // An outline no real camera produces (focal length ~10x the image).
            var corners = new[] { new PointD(470, 140), new PointD(1150, 190), new PointD(1290, 1080), new PointD(330, 1010) };

            Assert.Equal(PageAspect.EdgeRatio(corners), PageAspect.Estimate(corners, 1600, 1200), 9);
        }

        [Fact]
        public void OutputSize_CapsHugePhotos()
        {
            var full = new[] { new PointD(0, 0), new PointD(6000, 0), new PointD(6000, 4500), new PointD(0, 4500) };

            var size = PerspectiveWarp.OutputSize(full, 6000, 4500);

            Assert.True(size.Width <= PerspectiveWarp.MaxLongEdge && size.Height <= PerspectiveWarp.MaxShortEdge);
            Assert.Equal(3300, size.Width);
        }

        [Fact]
        public void Warp_FlattensAPhotographedPage_KeepingItsLayout()
        {
            var corners = new[] { new PointD(470, 140), new PointD(1150, 190), new PointD(1290, 1080), new PointD(330, 1010) };
            using (var page = SyntheticPages.Page(850, 1100, SyntheticPages.DimPaper, SyntheticPages.DimInk, marker: true))
            using (var photo = SyntheticPages.Photograph(page, 1600, 1200, corners, SyntheticPages.Desk))
            {
                var flat = PerspectiveWarp.Warp(PixelBuffer.FromBitmap(photo), corners);

                // The black marker sits where it does on the page: centre (0.15, 0.1 + 0.05 * 850/1100).
                var marker = SyntheticPages.DarkCentroid(flat, 40);
                Assert.InRange(marker.X, 0.13, 0.17);
                Assert.InRange(marker.Y, 0.118, 0.158);
                // No desk left at the edges: the corners are paper-colored.
                var corner = (2 * flat.Width + 2) * 3;
                Assert.InRange((int)flat.Data[corner + 2], SyntheticPages.DimPaper.R - 20, SyntheticPages.DimPaper.R + 20);
            }
        }
    }
}
```

- [ ] **Step 3: Run to verify they fail**

Run: Build
Expected: compile errors `CS0103 'PageAspect'`, `'PerspectiveWarp'`.

- [ ] **Step 4: Implement**

`VSTO2/OutlookAI/Services/QuickTools/PageAspect.cs`:

```csharp
using System;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// Width ÷ height of the real page behind a photographed outline (TL, TR,
    /// BR, BL in pixels), assuming the optical center is the image center
    /// (Zhang &amp; He, "Whiteboard scanning and image enhancement"). Falls back to
    /// the longer opposite edges when the photo shows no usable perspective.
    /// </summary>
    internal static class PageAspect
    {
        private const double MinRatio = 0.2;
        private const double MaxRatio = 5.0;

        public static double Estimate(PointD[] c, int imageWidth, int imageHeight)
        {
            var fallback = EdgeRatio(c);
            double u0 = imageWidth / 2.0, v0 = imageHeight / 2.0;
            var m1 = V(c[0], u0, v0);
            var m2 = V(c[1], u0, v0);
            var m4 = V(c[2], u0, v0);
            var m3 = V(c[3], u0, v0);

            var d2 = Dot(Cross(m2, m4), m3);
            var d3 = Dot(Cross(m3, m4), m2);
            if (Math.Abs(d2) < 1e-9 || Math.Abs(d3) < 1e-9) return fallback;
            var c14 = Cross(m1, m4);
            var k2 = Dot(c14, m3) / d2;
            var k3 = Dot(c14, m2) / d3;
            var n2 = new[] { k2 * m2[0] - m1[0], k2 * m2[1] - m1[1], k2 * m2[2] - m1[2] };
            var n3 = new[] { k3 * m3[0] - m1[0], k3 * m3[1] - m1[1], k3 * m3[2] - m1[2] };

            var den = n2[2] * n3[2];
            if (Math.Abs(den) < 1e-12) return fallback; // no perspective: the edges are right
            var f2 = -(n2[0] * n3[0] + n2[1] * n3[1]) / den;
            var maxEdge = Math.Max(imageWidth, imageHeight);
            // A real camera's focal length is a fraction to a few times the image size.
            if (f2 <= 0 || f2 < Sq(0.25 * maxEdge) || f2 > Sq(4.0 * maxEdge)) return fallback;

            var ratio = Math.Sqrt(
                ((Sq(n2[0]) + Sq(n2[1])) / f2 + Sq(n2[2]))
                / ((Sq(n3[0]) + Sq(n3[1])) / f2 + Sq(n3[2])));
            return ratio >= MinRatio && ratio <= MaxRatio ? ratio : fallback;
        }

        public static double EdgeRatio(PointD[] c) =>
            Math.Max(Distance(c[0], c[1]), Distance(c[3], c[2]))
            / Math.Max(Distance(c[0], c[3]), Distance(c[1], c[2]));

        private static double[] V(PointD p, double u0, double v0) => new[] { p.X - u0, p.Y - v0, 1.0 };

        private static double[] Cross(double[] a, double[] b) => new[]
        {
            a[1] * b[2] - a[2] * b[1],
            a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0],
        };

        private static double Dot(double[] a, double[] b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];

        private static double Sq(double v) => v * v;

        private static double Distance(PointD a, PointD b) => Math.Sqrt(Sq(a.X - b.X) + Sq(a.Y - b.Y));
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/PerspectiveWarp.cs`:

```csharp
using System;
using System.Drawing;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>Maps a photographed page (4 corners in pixels) onto an upright rectangle.</summary>
    internal static class PerspectiveWarp
    {
        public const int MaxLongEdge = 3300;
        public const int MaxShortEdge = 2550;

        /// <summary>
        /// Height from the longer of the left and right edges, width from the
        /// page's estimated aspect ratio; scaled down (never up) to fit 2550 x 3300
        /// (or 3300 x 2550 when wider than tall).
        /// </summary>
        public static Size OutputSize(PointD[] cornersPx, int imageWidth, int imageHeight)
        {
            var height = Math.Max(Distance(cornersPx[0], cornersPx[3]), Distance(cornersPx[1], cornersPx[2]));
            var width = height * PageAspect.Estimate(cornersPx, imageWidth, imageHeight);
            double maxW = width >= height ? MaxLongEdge : MaxShortEdge;
            double maxH = width >= height ? MaxShortEdge : MaxLongEdge;
            var scale = Math.Min(1.0, Math.Min(maxW / width, maxH / height));
            return new Size(Math.Max(1, (int)Math.Round(width * scale)), Math.Max(1, (int)Math.Round(height * scale)));
        }

        public static PixelBuffer Warp(PixelBuffer src, PointD[] cornersPx)
        {
            var size = OutputSize(cornersPx, src.Width, src.Height);
            int outW = size.Width, outH = size.Height;
            var map = Homography.FromPoints(
                new[] { new PointD(0, 0), new PointD(outW, 0), new PointD(outW, outH), new PointD(0, outH) },
                cornersPx);
            var c = map.Coefficients;
            var dst = new PixelBuffer(outW, outH);
            var s = src.Data;
            var d = dst.Data;
            for (var y = 0; y < outH; y++)
            {
                var fy = y + 0.5;
                // Along a row the numerators and the denominator are linear in x.
                var nx = c[0] * 0.5 + c[1] * fy + c[2];
                var ny = c[3] * 0.5 + c[4] * fy + c[5];
                var nw = c[6] * 0.5 + c[7] * fy + c[8];
                var o = y * outW * 3;
                for (var x = 0; x < outW; x++, o += 3)
                {
                    Sample(s, src.Width, src.Height, nx / nw - 0.5, ny / nw - 0.5, d, o);
                    nx += c[0];
                    ny += c[3];
                    nw += c[6];
                }
            }
            return dst;
        }

        private static void Sample(byte[] s, int w, int h, double x, double y, byte[] d, int o)
        {
            if (x < 0) x = 0; else if (x > w - 1) x = w - 1;
            if (y < 0) y = 0; else if (y > h - 1) y = h - 1;
            var x0 = (int)x;
            var y0 = (int)y;
            var x1 = x0 + 1 < w ? x0 + 1 : x0;
            var y1 = y0 + 1 < h ? y0 + 1 : y0;
            var fx = x - x0;
            var fy = y - y0;
            var i00 = (y0 * w + x0) * 3;
            var i10 = (y0 * w + x1) * 3;
            var i01 = (y1 * w + x0) * 3;
            var i11 = (y1 * w + x1) * 3;
            for (var k = 0; k < 3; k++)
            {
                var top = s[i00 + k] + (s[i10 + k] - s[i00 + k]) * fx;
                var bottom = s[i01 + k] + (s[i11 + k] - s[i01 + k]) * fx;
                d[o + k] = (byte)(top + (bottom - top) * fy + 0.5);
            }
        }

        private static double Distance(PointD a, PointD b) =>
            Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
    }
}
```

`OutlookAI.csproj`: after `<Compile Include="Services\QuickTools\JpegFiles.cs" />` add `<Compile Include="Services\QuickTools\PageAspect.cs" />`; after `<Compile Include="Services\QuickTools\PagePlan.cs" />` add `<Compile Include="Services\QuickTools\PerspectiveWarp.cs" />`.

In the spec, § 6.4 step 1, replace "Width is the longer of the top and bottom edges, height the longer of the left and right edges;" with:

```markdown
   Height is the longer of the left and right edges; width is the height times
   the page's aspect ratio, estimated from the perspective (Zhang & He,
   assuming the optical center is the image center) and falling back to the
   longer of the top and bottom edges when the photo shows no usable
   perspective or the estimate is implausible (focal length outside 0.25–4×
   the image size, or a ratio outside 0.2–5);
```

- [ ] **Step 5: Run to verify they pass**

Run: Build; Test `PerspectiveWarpTests`
Expected: 7 passed (3 theory cases + 4 facts).

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests docs/superpowers/specs/2026-10-13-quick-tools-images-to-pdf-design.md
git commit -m "feat(quick-tools): flatten photographed pages with their true proportions" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 6: Page look (white balance, even light, levels, sharpen)

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/PageLook.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/PageLookTests.cs`

**Interfaces:**
- Consumes: `PixelBuffer` (Task 4), `CleanupStrength` (Task 2), `SyntheticPages` (Task 5).
- Produces: `sealed class LookProfile { bool EvenLight; double BlackPercentile; double WhitePercentile; double Gamma; double Sharpen; }`; `static class PageLook { LookProfile Light, Normal, Strong; LookProfile For(CleanupStrength); void Apply(PixelBuffer px, LookProfile profile, bool grayscale); internal byte[] Luminance(PixelBuffer px); internal int Percentile(byte[] values, double percent); }`.

Order of operations (fixed): color mode white-balances on the paper first; then light evening (if the profile has it), levels from the luminance's black and white percentiles, sharpening. Gamma below 1 darkens midtones: `out = x^(1/gamma)`.

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/PageLookTests.cs`:

```csharp
using System.Drawing;
using System.Linq;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PageLookTests
    {
        private static PixelBuffer DimPage(Color? stamp = null)
        {
            using (var page = SyntheticPages.Page(600, 800, SyntheticPages.DimPaper, SyntheticPages.DimInk, stamp: stamp))
            {
                return PixelBuffer.FromBitmap(page);
            }
        }

        private static byte[] Region(PixelBuffer px, double x0, double y0, double x1, double y1)
        {
            var lum = PageLook.Luminance(px);
            return Enumerable.Range((int)(y0 * px.Height), (int)((y1 - y0) * px.Height))
                .SelectMany(y => Enumerable.Range((int)(x0 * px.Width), (int)((x1 - x0) * px.Width)).Select(x => lum[y * px.Width + x]))
                .ToArray();
        }

        [Fact]
        public void Profiles_AreTheSpecDefaults()
        {
            Assert.False(PageLook.For(CleanupStrength.Light).EvenLight);
            Assert.Equal(98.0, PageLook.For(CleanupStrength.Normal).WhitePercentile);
            Assert.Equal(0.85, PageLook.For(CleanupStrength.Strong).Gamma);
            Assert.Equal(0.8, PageLook.For(CleanupStrength.Strong).Sharpen);
        }

        [Fact]
        public void DimTintedPage_GetsAWhiteBackground_AndDarkText()
        {
            var px = DimPage();

            PageLook.Apply(px, PageLook.Normal, grayscale: true);

            Assert.True(PageLook.Percentile(Region(px, 0.1, 0.9, 0.9, 0.97), 50) > 235, "background not white");
            Assert.True(PageLook.Percentile(Region(px, 0.1, 0.3, 0.9, 0.88), 1) < 60, "text not dark");
        }

        [Fact]
        public void Grayscale_IsGray()
        {
            var px = DimPage(Color.FromArgb(200, 40, 40));

            PageLook.Apply(px, PageLook.Normal, grayscale: true);

            for (var i = 0; i < px.Data.Length; i += 3)
            {
                Assert.True(px.Data[i] == px.Data[i + 1] && px.Data[i + 1] == px.Data[i + 2]);
            }
        }

        [Fact]
        public void ColorMode_KeepsARedStampRed_AndTheTintNeutral()
        {
            var px = DimPage(Color.FromArgb(200, 40, 40));

            PageLook.Apply(px, PageLook.Normal, grayscale: false);

            var red = 0;
            for (var i = 0; i < px.Data.Length; i += 3)
            {
                if (px.Data[i + 2] - px.Data[i + 1] > 60) red++;
            }
            Assert.True(red > 50, "stamp lost its color");
            // Paper: neutral white, not yellow.
            var o = ((int)(0.95 * px.Height) * px.Width + px.Width / 2) * 3;
            Assert.True(px.Data[o] > 230 && px.Data[o + 1] > 230 && px.Data[o + 2] > 230);
        }

        [Fact]
        public void FlatImage_StaysUniform()
        {
            var px = new PixelBuffer(200, 200);
            for (var i = 0; i < px.Data.Length; i++) px.Data[i] = 120;

            PageLook.Apply(px, PageLook.Strong, grayscale: true);

            Assert.True(px.Data.All(v => v == px.Data[0]));
        }

        [Fact]
        public void Apply_IsDeterministic()
        {
            var a = DimPage();
            var b = a.Clone();

            PageLook.Apply(a, PageLook.Strong, grayscale: false);
            PageLook.Apply(b, PageLook.Strong, grayscale: false);

            Assert.Equal(a.Data, b.Data);
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile error `CS0103 'PageLook'`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/QuickTools/PageLook.cs`:

```csharp
using System;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class LookProfile
    {
        public LookProfile(bool evenLight, double blackPercentile, double whitePercentile, double gamma, double sharpen)
        {
            EvenLight = evenLight;
            BlackPercentile = blackPercentile;
            WhitePercentile = whitePercentile;
            Gamma = gamma;
            Sharpen = sharpen;
        }

        public bool EvenLight { get; }
        public double BlackPercentile { get; }
        public double WhitePercentile { get; }
        /// <summary>Below 1 darkens midtones: out = x^(1/Gamma).</summary>
        public double Gamma { get; }
        public double Sharpen { get; }
    }

    /// <summary>
    /// Deterministic document cleanup: (color only) white balance on the paper,
    /// light evening, levels from the luminance's black and white points, a light
    /// unsharp mask.
    /// </summary>
    internal static class PageLook
    {
        public static readonly LookProfile Light = new LookProfile(false, 0.5, 99.5, 1.00, 0.3);
        public static readonly LookProfile Normal = new LookProfile(true, 1, 98, 0.95, 0.5);
        public static readonly LookProfile Strong = new LookProfile(true, 2, 95, 0.85, 0.8);

        private const int Block = 8;
        private const int SharpenThreshold = 2;

        public static LookProfile For(CleanupStrength strength)
        {
            switch (strength)
            {
                case CleanupStrength.Light: return Light;
                case CleanupStrength.Strong: return Strong;
                default: return Normal;
            }
        }

        public static void Apply(PixelBuffer px, LookProfile profile, bool grayscale)
        {
            var n = px.Width * px.Height;
            var d = px.Data;
            if (!grayscale) WhiteBalance(px);
            var lum = Luminance(px);

            byte[][] channels;
            if (grayscale)
            {
                channels = new[] { lum };
            }
            else
            {
                channels = new[] { new byte[n], new byte[n], new byte[n] };
                for (var i = 0; i < n; i++)
                {
                    channels[0][i] = d[3 * i];
                    channels[1][i] = d[3 * i + 1];
                    channels[2][i] = d[3 * i + 2];
                }
            }

            if (profile.EvenLight) EvenLight(channels, lum, px.Width, px.Height, grayscale);

            var black = Percentile(lum, profile.BlackPercentile);
            var white = Percentile(lum, profile.WhitePercentile);
            if (white - black >= 10)
            {
                var lut = Levels(black, white, profile.Gamma);
                foreach (var c in channels)
                {
                    for (var i = 0; i < n; i++) c[i] = lut[c[i]];
                }
            }

            if (profile.Sharpen > 0)
            {
                foreach (var c in channels) Sharpen(c, px.Width, px.Height, profile.Sharpen);
            }

            for (var i = 0; i < n; i++)
            {
                if (grayscale)
                {
                    d[3 * i] = d[3 * i + 1] = d[3 * i + 2] = lum[i];
                }
                else
                {
                    d[3 * i] = channels[0][i];
                    d[3 * i + 1] = channels[1][i];
                    d[3 * i + 2] = channels[2][i];
                }
            }
        }

        internal static byte[] Luminance(PixelBuffer px)
        {
            var n = px.Width * px.Height;
            var d = px.Data;
            var lum = new byte[n];
            for (var i = 0; i < n; i++)
            {
                lum[i] = (byte)((114 * d[3 * i] + 587 * d[3 * i + 1] + 299 * d[3 * i + 2] + 500) / 1000);
            }
            return lum;
        }

        internal static int Percentile(byte[] values, double percent)
        {
            var histogram = new long[256];
            foreach (var v in values) histogram[v]++;
            var target = (long)Math.Ceiling(values.Length * percent / 100.0);
            if (target < 1) target = 1;
            long seen = 0;
            for (var v = 0; v < 256; v++)
            {
                seen += histogram[v];
                if (seen >= target) return v;
            }
            return 255;
        }

        // Scales each channel so the paper (median of the brightest 10%) becomes neutral.
        private static void WhiteBalance(PixelBuffer px)
        {
            var lum = Luminance(px);
            var threshold = Percentile(lum, 90);
            var hb = new long[256];
            var hg = new long[256];
            var hr = new long[256];
            long count = 0;
            var d = px.Data;
            for (var i = 0; i < lum.Length; i++)
            {
                if (lum[i] < threshold) continue;
                hb[d[3 * i]]++;
                hg[d[3 * i + 1]]++;
                hr[d[3 * i + 2]]++;
                count++;
            }
            if (count == 0) return;
            double mb = Median(hb, count), mg = Median(hg, count), mr = Median(hr, count);
            var paper = 0.114 * mb + 0.587 * mg + 0.299 * mr;
            var lutB = Gain(paper / Math.Max(mb, 1));
            var lutG = Gain(paper / Math.Max(mg, 1));
            var lutR = Gain(paper / Math.Max(mr, 1));
            for (var i = 0; i < lum.Length; i++)
            {
                d[3 * i] = lutB[d[3 * i]];
                d[3 * i + 1] = lutG[d[3 * i + 1]];
                d[3 * i + 2] = lutR[d[3 * i + 2]];
            }
        }

        private static double Median(long[] histogram, long count)
        {
            long seen = 0;
            for (var v = 0; v < 256; v++)
            {
                seen += histogram[v];
                if (seen * 2 >= count) return v;
            }
            return 255;
        }

        private static byte[] Gain(double gain)
        {
            gain = Math.Max(0.5, Math.Min(2.0, gain));
            var lut = new byte[256];
            for (var v = 0; v < 256; v++) lut[v] = Clamp((float)(v * gain));
            return lut;
        }

        // Divides by the paper brightness around each pixel: the max of each 8x8
        // block (text is darker than paper), box-blurred with a radius of 1/20 of
        // the short side, bilinear between blocks.
        private static void EvenLight(byte[][] channels, byte[] lum, int w, int h, bool lumIsChannel)
        {
            var bw = (w + Block - 1) / Block;
            var bh = (h + Block - 1) / Block;
            var background = new float[bw * bh];
            for (var by = 0; by < bh; by++)
            {
                for (var bx = 0; bx < bw; bx++)
                {
                    var max = 0;
                    var yEnd = Math.Min(h, (by + 1) * Block);
                    var xEnd = Math.Min(w, (bx + 1) * Block);
                    for (var y = by * Block; y < yEnd; y++)
                    {
                        for (var x = bx * Block; x < xEnd; x++)
                        {
                            if (lum[y * w + x] > max) max = lum[y * w + x];
                        }
                    }
                    background[by * bw + bx] = max;
                }
            }
            BoxBlur(background, bw, bh, Math.Max(1, (int)Math.Round(Math.Min(w, h) / 20.0 / Block)));

            for (var y = 0; y < h; y++)
            {
                var gy = (y + 0.5) / Block - 0.5;
                for (var x = 0; x < w; x++)
                {
                    var gx = (x + 0.5) / Block - 0.5;
                    var gain = 255f / Math.Max(Bilinear(background, bw, bh, gx, gy), 16f);
                    var i = y * w + x;
                    foreach (var c in channels) c[i] = Clamp(c[i] * gain);
                    if (!lumIsChannel) lum[i] = Clamp(lum[i] * gain);
                }
            }
        }

        private static void BoxBlur(float[] a, int w, int h, int r)
        {
            var tmp = new float[a.Length];
            for (var y = 0; y < h; y++)
            {
                for (var x = 0; x < w; x++)
                {
                    float sum = 0;
                    var count = 0;
                    for (var k = -r; k <= r; k++)
                    {
                        var xx = x + k;
                        if (xx < 0 || xx >= w) continue;
                        sum += a[y * w + xx];
                        count++;
                    }
                    tmp[y * w + x] = sum / count;
                }
            }
            for (var y = 0; y < h; y++)
            {
                for (var x = 0; x < w; x++)
                {
                    float sum = 0;
                    var count = 0;
                    for (var k = -r; k <= r; k++)
                    {
                        var yy = y + k;
                        if (yy < 0 || yy >= h) continue;
                        sum += tmp[yy * w + x];
                        count++;
                    }
                    a[y * w + x] = sum / count;
                }
            }
        }

        private static float Bilinear(float[] g, int w, int h, double x, double y)
        {
            if (x < 0) x = 0; else if (x > w - 1) x = w - 1;
            if (y < 0) y = 0; else if (y > h - 1) y = h - 1;
            var x0 = (int)x;
            var y0 = (int)y;
            var x1 = Math.Min(x0 + 1, w - 1);
            var y1 = Math.Min(y0 + 1, h - 1);
            var fx = (float)(x - x0);
            var fy = (float)(y - y0);
            var top = g[y0 * w + x0] + (g[y0 * w + x1] - g[y0 * w + x0]) * fx;
            var bottom = g[y1 * w + x0] + (g[y1 * w + x1] - g[y1 * w + x0]) * fx;
            return top + (bottom - top) * fy;
        }

        private static byte[] Levels(int black, int white, double gamma)
        {
            var lut = new byte[256];
            for (var v = 0; v < 256; v++)
            {
                var x = (v - black) / (double)(white - black);
                if (x < 0) x = 0; else if (x > 1) x = 1;
                lut[v] = (byte)Math.Round(255 * Math.Pow(x, 1 / gamma));
            }
            return lut;
        }

        // Unsharp mask: v + amount * (v - 3x3 box blur), ignoring differences of 2 or less.
        private static void Sharpen(byte[] c, int w, int h, double amount)
        {
            var blur = new byte[c.Length];
            for (var y = 0; y < h; y++)
            {
                for (var x = 0; x < w; x++)
                {
                    var sum = 0;
                    for (var dy = -1; dy <= 1; dy++)
                    {
                        var yy = Math.Min(h - 1, Math.Max(0, y + dy));
                        for (var dx = -1; dx <= 1; dx++)
                        {
                            var xx = Math.Min(w - 1, Math.Max(0, x + dx));
                            sum += c[yy * w + xx];
                        }
                    }
                    blur[y * w + x] = (byte)((sum + 4) / 9);
                }
            }
            for (var i = 0; i < c.Length; i++)
            {
                var diff = c[i] - blur[i];
                if (Math.Abs(diff) > SharpenThreshold) c[i] = Clamp((float)(c[i] + amount * diff));
            }
        }

        private static byte Clamp(float v) => v <= 0 ? (byte)0 : v >= 255 ? (byte)255 : (byte)(v + 0.5f);
    }
}
```

`OutlookAI.csproj`: after `<Compile Include="Services\QuickTools\PageAspect.cs" />` add `<Compile Include="Services\QuickTools\PageLook.cs" />`.

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `PageLookTests`
Expected: 6 passed. If `DimTintedPage…` fails on the text check, print `Percentile(Region(…), 1)` — the darkest 1% of the text area must reach the black point; do not loosen the thresholds without understanding why.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools/PageLook.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools/PageLookTests.cs
git commit -m "feat(quick-tools): deterministic page cleanup (white balance, evening, levels, sharpen)" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 7: PageCleaner (file in, cleaned JPEG out)

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/PageCleaner.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/PageCleanerTests.cs`

**Interfaces:**
- Consumes: `PagePlan`, `PageColor` (Task 2), `JpegFiles`, `PixelBuffer` (Task 4), `PerspectiveWarp` (Task 5), `PageLook` (Task 6).
- Produces: `static class PageCleaner { void WriteCleaned(string uprightPath, PagePlan plan, PageColor color, string outputPath); void WriteRotated(string sourceJpeg, int quarterTurns, string outputPath); internal RotateFlipType QuarterTurns(int turns); }` (turns clockwise; −1 ≡ 3).

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/PageCleanerTests.cs`:

```csharp
using System;
using System.Drawing;
using System.IO;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PageCleanerTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-cleaner", Path.GetRandomFileName());

        public PageCleanerTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private string Photo(out PointD[] cornersPx)
        {
            cornersPx = new[] { new PointD(470, 140), new PointD(1150, 190), new PointD(1290, 1080), new PointD(330, 1010) };
            var path = Path.Combine(_dir, "upright.jpg");
            using (var page = SyntheticPages.Page(850, 1100, SyntheticPages.DimPaper, SyntheticPages.DimInk, marker: true))
            using (var photo = SyntheticPages.Photograph(page, 1600, 1200, cornersPx, SyntheticPages.Desk))
            {
                JpegFiles.Save(photo, path, JpegFiles.OriginalQuality);
            }
            return path;
        }

        private static PagePlan PlanFor(PointD[] cornersPx, int width, int height, string strength)
        {
            var corners = "[";
            for (var i = 0; i < 4; i++)
            {
                corners += (i > 0 ? "," : "") + "[" + (cornersPx[i].X / width).ToString("R", System.Globalization.CultureInfo.InvariantCulture)
                    + "," + (cornersPx[i].Y / height).ToString("R", System.Globalization.CultureInfo.InvariantCulture) + "]";
            }
            return PagePlan.FromJson("{\"kind\":\"document_page\",\"reason\":\"\",\"corners\":" + corners
                + "],\"color\":\"grayscale\",\"strength\":\"" + strength + "\"}");
        }

        [Fact]
        public void WriteCleaned_FlattensAndCleansThePage()
        {
            PointD[] cornersPx;
            var upright = Photo(out cornersPx);
            var output = Path.Combine(_dir, "cleaned.jpg");

            PageCleaner.WriteCleaned(upright, PlanFor(cornersPx, 1600, 1200, "light"), PageColor.Grayscale, output);

            using (var cleaned = JpegFiles.Load(output))
            {
                var px = PixelBuffer.FromBitmap(cleaned);
                var marker = SyntheticPages.DarkCentroid(px, 40);
                Assert.InRange(marker.X, 0.13, 0.17);
                Assert.InRange(marker.Y, 0.118, 0.158);
                // Bottom margin of the page: paper, now near white.
                var o = ((int)(0.95 * px.Height) * px.Width + px.Width / 2) * 3;
                Assert.True(px.Data[o] > 200);
            }
        }

        [Fact]
        public void WriteCleaned_FallbackPlan_KeepsTheWholeImage()
        {
            var upright = Path.Combine(_dir, "whole.jpg");
            using (var bitmap = new Bitmap(400, 300)) JpegFiles.Save(bitmap, upright, 95);
            var output = Path.Combine(_dir, "whole-cleaned.jpg");

            PageCleaner.WriteCleaned(upright, PagePlan.Fallback(), PageColor.Grayscale, output);

            using (var cleaned = JpegFiles.Load(output))
            {
                Assert.Equal(new Size(400, 300), cleaned.Size);
            }
        }

        [Fact]
        public void WriteRotated_TurnsClockwise()
        {
            var source = Path.Combine(_dir, "wide.jpg");
            using (var bitmap = new Bitmap(40, 20)) JpegFiles.Save(bitmap, source, 95);
            var output = Path.Combine(_dir, "tall.jpg");

            PageCleaner.WriteRotated(source, 1, output);

            using (var rotated = JpegFiles.Load(output))
            {
                Assert.Equal(new Size(20, 40), rotated.Size);
            }
        }

        [Theory]
        [InlineData(0, RotateFlipType.RotateNoneFlipNone)]
        [InlineData(1, RotateFlipType.Rotate90FlipNone)]
        [InlineData(2, RotateFlipType.Rotate180FlipNone)]
        [InlineData(3, RotateFlipType.Rotate270FlipNone)]
        [InlineData(-1, RotateFlipType.Rotate270FlipNone)]
        [InlineData(5, RotateFlipType.Rotate90FlipNone)]
        public void QuarterTurns_WrapAround(int turns, RotateFlipType expected)
        {
            Assert.Equal(expected, PageCleaner.QuarterTurns(turns));
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile error `CS0103 'PageCleaner'`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/QuickTools/PageCleaner.cs`:

```csharp
using System.Drawing;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// File in, JPEG file out, so nothing full-size outlives the call (32-bit
    /// Outlook has 2 GB of address space).
    /// </summary>
    internal static class PageCleaner
    {
        public static void WriteCleaned(string uprightPath, PagePlan plan, PageColor color, string outputPath)
        {
            PixelBuffer src;
            using (var bitmap = JpegFiles.Load(uprightPath))
            {
                src = PixelBuffer.FromBitmap(bitmap);
            }
            var corners = new PointD[4];
            for (var i = 0; i < 4; i++)
            {
                corners[i] = new PointD(plan.Corners[i].X * src.Width, plan.Corners[i].Y * src.Height);
            }
            var page = PerspectiveWarp.Warp(src, corners);
            src = null;
            PageLook.Apply(page, PageLook.For(plan.Strength), color == PageColor.Grayscale);
            using (var output = page.ToBitmap())
            {
                JpegFiles.Save(output, outputPath, JpegFiles.PageQuality);
            }
        }

        public static void WriteRotated(string sourceJpeg, int quarterTurns, string outputPath)
        {
            using (var bitmap = JpegFiles.Load(sourceJpeg))
            {
                bitmap.RotateFlip(QuarterTurns(quarterTurns));
                JpegFiles.Save(bitmap, outputPath, JpegFiles.PageQuality);
            }
        }

        internal static RotateFlipType QuarterTurns(int turns)
        {
            switch (((turns % 4) + 4) % 4)
            {
                case 1: return RotateFlipType.Rotate90FlipNone;
                case 2: return RotateFlipType.Rotate180FlipNone;
                case 3: return RotateFlipType.Rotate270FlipNone;
                default: return RotateFlipType.RotateNoneFlipNone;
            }
        }
    }
}
```

`OutlookAI.csproj`: after `<Compile Include="Services\QuickTools\PageAspect.cs" />` add `<Compile Include="Services\QuickTools\PageCleaner.cs" />`.

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `PageCleanerTests`
Expected: 9 passed.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools/PageCleaner.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools/PageCleanerTests.cs
git commit -m "feat(quick-tools): clean a page from its file into a JPEG" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 8: PDF writer and saving to the chosen place

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/JpegHeader.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/ImagePdfWriter.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/ImagesToPdfOutput.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/ImagePdfWriterTests.cs`

**Interfaces:**
- Consumes: `JpegFiles` (Task 4).
- Produces: `static class JpegHeader { bool TryRead(byte[] jpeg, out int width, out int height, out int components); }`; `static class ImagePdfWriter { void Write(IReadOnlyList<string> jpegPaths, string title, string producer, Stream output); void WriteFile(IReadOnlyList<string> jpegPaths, string title, string producer, string path); internal string TextString(string text); }`; `static class ImagesToPdfOutput { string Publish(string tempPdf, string destination); /* null = saved, else a message for the user */ }`.

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/ImagePdfWriterTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class ImagePdfWriterTests : IDisposable
    {
        private static readonly Encoding Latin1 = Encoding.GetEncoding(28591);
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-pdf", Path.GetRandomFileName());

        public ImagePdfWriterTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private string Jpeg(string name, int width, int height)
        {
            var path = Path.Combine(_dir, name);
            using (var bitmap = new Bitmap(width, height))
            {
                using (var g = Graphics.FromImage(bitmap)) g.Clear(Color.White);
                JpegFiles.Save(bitmap, path, 85);
            }
            return path;
        }

        private string Pdf(IReadOnlyList<string> pages, string title = "Lease docs")
        {
            using (var stream = new MemoryStream())
            {
                ImagePdfWriter.Write(pages, title, "OutlookAI v2.3.0", stream);
                return Latin1.GetString(stream.ToArray());
            }
        }

        [Fact]
        public void JpegHeader_ReadsSizeAndComponents()
        {
            int w, h, c;
            Assert.True(JpegHeader.TryRead(File.ReadAllBytes(Jpeg("a.jpg", 300, 400)), out w, out h, out c));
            Assert.Equal(300, w);
            Assert.Equal(400, h);
            Assert.Equal(3, c);
        }

        [Fact]
        public void JpegHeader_RejectsOtherBytes()
        {
            int w, h, c;
            Assert.False(JpegHeader.TryRead(Encoding.ASCII.GetBytes("not a jpeg"), out w, out h, out c));
            Assert.False(JpegHeader.TryRead(new byte[] { 0xFF, 0xD8, 0xFF }, out w, out h, out c));
        }

        [Fact]
        public void Write_HasOnePagePerImage_InTheRightOrientation()
        {
            var pdf = Pdf(new[] { Jpeg("p.jpg", 300, 400), Jpeg("l.jpg", 400, 300) });

            Assert.StartsWith("%PDF-1.4", pdf);
            Assert.Contains("/Count 2", pdf);
            Assert.Contains("/MediaBox [0 0 612 792]", pdf);
            Assert.Contains("/MediaBox [0 0 792 612]", pdf);
            Assert.Contains("/Width 300 /Height 400", pdf);
            // Fitted inside 18 pt margins and centered: 300x400 px at 1.89 pt/px.
            Assert.Contains("q 567 0 0 756 22.5 18 cm /Im1 Do Q", pdf);
            Assert.EndsWith("%%EOF\n", pdf);
        }

        [Fact]
        public void Write_XrefOffsetsPointAtTheirObjects()
        {
            var pdf = Pdf(new[] { Jpeg("p.jpg", 300, 400), Jpeg("l.jpg", 400, 300) });

            var startxref = int.Parse(Regex.Match(pdf, @"startxref\n(\d+)\n%%EOF").Groups[1].Value);
            Assert.Equal("xref", pdf.Substring(startxref, 4));
            var entries = Regex.Matches(pdf.Substring(startxref), @"(\d{10}) 00000 n ").Cast<Match>().ToList();
            Assert.Equal(9, entries.Count); // catalog, pages, info, 2 x (page, image, content)
            for (var i = 0; i < entries.Count; i++)
            {
                var offset = int.Parse(entries[i].Groups[1].Value);
                Assert.StartsWith((i + 1) + " 0 obj", pdf.Substring(offset));
            }
        }

        [Fact]
        public void Write_EmbedsTheJpegBytesAsIs()
        {
            var path = Jpeg("p.jpg", 300, 400);
            var bytes = File.ReadAllBytes(path);

            var pdf = Pdf(new[] { path });

            var start = pdf.IndexOf("/Length " + bytes.Length + " >>\nstream\n", StringComparison.Ordinal);
            Assert.True(start > 0);
            var data = pdf.IndexOf("stream\n", start, StringComparison.Ordinal) + "stream\n".Length;
            Assert.Equal(bytes, Latin1.GetBytes(pdf.Substring(data, bytes.Length)));
        }

        [Fact]
        public void Write_TitleCanHoldAnyCharacters()
        {
            var pdf = Pdf(new[] { Jpeg("p.jpg", 30, 40) }, title: "Résumé");

            Assert.Contains("/Title <FEFF005200E9", pdf);
        }

        [Fact]
        public void Write_NeedsAtLeastOnePage_AndOnlyJpegs()
        {
            Assert.Throws<ArgumentException>(() => Pdf(new string[0]));
            var notJpeg = Path.Combine(_dir, "x.jpg");
            File.WriteAllText(notJpeg, "nope");
            Assert.Throws<InvalidDataException>(() => Pdf(new[] { notJpeg }));
        }

        [Fact]
        public void Publish_CopiesTheFile()
        {
            var temp = Path.Combine(_dir, "temp.pdf");
            File.WriteAllText(temp, "%PDF-1.4");
            var destination = Path.Combine(_dir, "out.pdf");

            Assert.Null(ImagesToPdfOutput.Publish(temp, destination));
            Assert.Equal("%PDF-1.4", File.ReadAllText(destination));
        }

        [Fact]
        public void Publish_LockedDestination_SaysCloseIt()
        {
            var temp = Path.Combine(_dir, "temp.pdf");
            File.WriteAllText(temp, "%PDF-1.4");
            var destination = Path.Combine(_dir, "open-in-viewer.pdf");
            File.WriteAllText(destination, "old");
            using (new FileStream(destination, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var message = ImagesToPdfOutput.Publish(temp, destination);

                Assert.NotNull(message);
                Assert.Contains("close it", message);
            }
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile errors `CS0103 'JpegHeader'`, `'ImagePdfWriter'`, `'ImagesToPdfOutput'`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/QuickTools/JpegHeader.cs`:

```csharp
namespace OutlookAI.Services.QuickTools
{
    /// <summary>Size and component count from a JPEG's frame header.</summary>
    internal static class JpegHeader
    {
        public static bool TryRead(byte[] jpeg, out int width, out int height, out int components)
        {
            width = height = components = 0;
            if (jpeg == null || jpeg.Length < 4 || jpeg[0] != 0xFF || jpeg[1] != 0xD8) return false;
            var i = 2;
            while (i + 3 < jpeg.Length)
            {
                if (jpeg[i] != 0xFF) { i++; continue; }
                var marker = jpeg[i + 1];
                if (marker == 0xFF) { i++; continue; }
                i += 2;
                if (marker == 0xD8 || marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7)) continue;
                if (marker == 0xD9 || marker == 0xDA) return false;
                if (i + 1 >= jpeg.Length) return false;
                var length = (jpeg[i] << 8) | jpeg[i + 1];
                if (length < 2 || i + length > jpeg.Length) return false;
                var isFrame = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
                if (isFrame)
                {
                    if (length < 8) return false;
                    height = (jpeg[i + 3] << 8) | jpeg[i + 4];
                    width = (jpeg[i + 5] << 8) | jpeg[i + 6];
                    components = jpeg[i + 7];
                    return width > 0 && height > 0;
                }
                i += length;
            }
            return false;
        }
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/ImagePdfWriter.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// A PDF 1.4 with one JPEG per Letter page (landscape when wider than tall),
    /// fitted inside 18 pt margins. The JPEG bytes are embedded as they are, one
    /// file in memory at a time.
    /// </summary>
    internal static class ImagePdfWriter
    {
        public const double LetterShort = 612;
        public const double LetterLong = 792;
        public const double Margin = 18;

        public static void WriteFile(IReadOnlyList<string> jpegPaths, string title, string producer, string path)
        {
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                Write(jpegPaths, title, producer, stream);
            }
        }

        public static void Write(IReadOnlyList<string> jpegPaths, string title, string producer, Stream output)
        {
            if (jpegPaths == null || jpegPaths.Count == 0)
                throw new ArgumentException("At least one page is needed.", nameof(jpegPaths));

            var pdf = new PdfOut(output);
            var count = jpegPaths.Count;
            var size = 4 + 3 * count;
            var offsets = new long[size];

            pdf.Write("%PDF-1.4\n");
            pdf.Write(new byte[] { (byte)'%', 0xE2, 0xE3, 0xCF, 0xD3, (byte)'\n' });

            offsets[1] = pdf.Position;
            pdf.Write("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

            var kids = new StringBuilder();
            for (var p = 0; p < count; p++) kids.Append(p == 0 ? "" : " ").Append(PageObject(p)).Append(" 0 R");
            offsets[2] = pdf.Position;
            pdf.Write("2 0 obj\n<< /Type /Pages /Kids [" + kids + "] /Count " + count + " >>\nendobj\n");

            offsets[3] = pdf.Position;
            pdf.Write("3 0 obj\n<< /Title " + TextString(title ?? "") + " /Producer " + TextString(producer ?? "") + " >>\nendobj\n");

            for (var p = 0; p < count; p++)
            {
                var jpeg = File.ReadAllBytes(jpegPaths[p]);
                int width, height, components;
                if (!JpegHeader.TryRead(jpeg, out width, out height, out components))
                    throw new InvalidDataException("Page " + (p + 1) + " is not a JPEG.");
                var colorSpace = components == 1 ? "/DeviceGray" : components == 3 ? "/DeviceRGB" : null;
                if (colorSpace == null)
                    throw new NotSupportedException("Page " + (p + 1) + " has " + components + " color components.");

                var landscape = width > height;
                var pageW = landscape ? LetterLong : LetterShort;
                var pageH = landscape ? LetterShort : LetterLong;
                var scale = Math.Min((pageW - 2 * Margin) / width, (pageH - 2 * Margin) / height);
                var drawW = width * scale;
                var drawH = height * scale;
                var content = "q " + N(drawW) + " 0 0 " + N(drawH) + " " + N((pageW - drawW) / 2) + " " + N((pageH - drawH) / 2) + " cm /Im1 Do Q";

                var pageObject = PageObject(p);
                offsets[pageObject] = pdf.Position;
                pdf.Write(pageObject + " 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + N(pageW) + " " + N(pageH)
                    + "] /Resources << /XObject << /Im1 " + (pageObject + 1) + " 0 R >> >> /Contents " + (pageObject + 2) + " 0 R >>\nendobj\n");

                offsets[pageObject + 1] = pdf.Position;
                pdf.Write((pageObject + 1) + " 0 obj\n<< /Type /XObject /Subtype /Image /Width " + width + " /Height " + height
                    + " /ColorSpace " + colorSpace + " /BitsPerComponent 8 /Filter /DCTDecode /Length " + jpeg.Length + " >>\nstream\n");
                pdf.Write(jpeg);
                pdf.Write("\nendstream\nendobj\n");

                offsets[pageObject + 2] = pdf.Position;
                pdf.Write((pageObject + 2) + " 0 obj\n<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream\nendobj\n");
            }

            var xref = pdf.Position;
            var table = new StringBuilder("xref\n0 " + size + "\n0000000000 65535 f \n");
            for (var o = 1; o < size; o++)
            {
                table.Append(offsets[o].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
            }
            pdf.Write(table.ToString());
            pdf.Write("trailer\n<< /Size " + size + " /Root 1 0 R /Info 3 0 R >>\nstartxref\n"
                + xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        }

        /// <summary>A PDF text string as UTF-16BE hex with a byte-order mark (any characters).</summary>
        internal static string TextString(string text)
        {
            var sb = new StringBuilder("<FEFF");
            foreach (var b in Encoding.BigEndianUnicode.GetBytes(text)) sb.Append(b.ToString("X2", CultureInfo.InvariantCulture));
            return sb.Append('>').ToString();
        }

        private static int PageObject(int index) => 4 + 3 * index;

        private static string N(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

        private sealed class PdfOut
        {
            private readonly Stream _stream;

            public PdfOut(Stream stream)
            {
                _stream = stream;
            }

            public long Position { get; private set; }

            public void Write(string ascii) => Write(Encoding.ASCII.GetBytes(ascii));

            public void Write(byte[] bytes)
            {
                _stream.Write(bytes, 0, bytes.Length);
                Position += bytes.Length;
            }
        }
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/ImagesToPdfOutput.cs`:

```csharp
using System;
using System.IO;

namespace OutlookAI.Services.QuickTools
{
    internal static class ImagesToPdfOutput
    {
        /// <summary>
        /// Copies the finished PDF (written locally first) to where the user chose,
        /// so a network share never holds a half-written file. Null when saved,
        /// otherwise a message for the user.
        /// </summary>
        public static string Publish(string tempPdf, string destination)
        {
            try
            {
                File.Copy(tempPdf, destination, overwrite: true);
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return "You don't have permission to save in that folder. Choose another folder.";
            }
            catch (IOException ex)
            {
                return "The PDF couldn't be saved there. If the file is open in another program, close it and try again.\n\n" + ex.Message;
            }
        }
    }
}
```

`OutlookAI.csproj`: after `<Compile Include="Services\QuickTools\ImageLoader.cs" />` add:

```xml
    <Compile Include="Services\QuickTools\ImagePdfWriter.cs" />
    <Compile Include="Services\QuickTools\ImagesToPdfOutput.cs" />
```

and after `<Compile Include="Services\QuickTools\JpegFiles.cs" />` add `<Compile Include="Services\QuickTools\JpegHeader.cs" />`. (Alphabetical: `ImageLoader`, `ImagePdfWriter`, `ImagesToPdfOutput`, `JpegFiles`, `JpegHeader`.)

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `ImagePdfWriterTests`
Expected: 9 passed. Optional manual check: write a PDF from two real phone photos to the scratchpad and open it in a PDF viewer.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools/ImagePdfWriterTests.cs
git commit -m "feat(quick-tools): write a PDF of JPEG pages and save it safely" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 9: Collecting the email's images

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/ImageAttachmentOrder.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/EmailImageCollector.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/ImageAttachmentOrderTests.cs`

**Interfaces:**
- Produces: `sealed class AttachmentInfo { int Index /* 1-based */; string FileName; string MimeTag; string ContentId; bool IsByValue; }`; `static class ImageAttachmentOrder { string[] ImageExtensions; bool IsImage(AttachmentInfo); bool IsInline(AttachmentInfo, string html); List<AttachmentInfo> Order(IEnumerable<AttachmentInfo>, string html); string SafeExtension(string fileName); }`; `sealed class CollectedImage(string path, string fileName, bool isInline)`; `sealed class CollectedImages(IReadOnlyList<CollectedImage> images, bool isRichText)`; `static class EmailImageCollector { CollectedImages Collect(Outlook.MailItem mail, string runFolder); }` (UI thread only).

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/ImageAttachmentOrderTests.cs`:

```csharp
using System.Linq;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class ImageAttachmentOrderTests
    {
        private static AttachmentInfo A(int index, string name, string cid = null, string mime = null, bool byValue = true) =>
            new AttachmentInfo { Index = index, FileName = name, ContentId = cid, MimeTag = mime, IsByValue = byValue };

        private const string Html =
            "<p>Here are the pages</p><img src=\"cid:image002.jpg@01DB1234\"><br><IMG SRC=\"CID:image001.png@01DB1234\">";

        [Fact]
        public void Order_InlineImagesByBodyPosition_ThenTheRestByIndex()
        {
            var ordered = ImageAttachmentOrder.Order(new[]
            {
                A(1, "scan-b.jpg"),
                A(2, "image001.png", "<image001.png@01DB1234>"),
                A(3, "image002.jpg", "image002.jpg@01DB1234"),
                A(4, "scan-a.jpg"),
            }, Html);

            Assert.Equal(new[] { 3, 2, 1, 4 }, ordered.Select(a => a.Index));
        }

        [Fact]
        public void Order_KeepsOnlyImagesByValue()
        {
            var ordered = ImageAttachmentOrder.Order(new[]
            {
                A(1, "contract.pdf"),
                A(2, "photo", mime: "image/jpeg"),
                A(3, "notes.docx"),
                A(4, "phone.HEIC"),
                A(5, "linked.jpg", byValue: false),
            }, null);

            Assert.Equal(new[] { 2, 4 }, ordered.Select(a => a.Index));
        }

        [Fact]
        public void Order_WithoutABody_UsesAttachmentOrder()
        {
            var ordered = ImageAttachmentOrder.Order(new[] { A(2, "b.png", "x@y"), A(1, "a.png", "z@y") }, null);

            Assert.Equal(new[] { 1, 2 }, ordered.Select(a => a.Index));
        }

        [Fact]
        public void IsInline_OnlyWhenTheBodyReferencesIt()
        {
            Assert.True(ImageAttachmentOrder.IsInline(A(2, "image001.png", "<image001.png@01DB1234>"), Html));
            Assert.False(ImageAttachmentOrder.IsInline(A(9, "other.png", "other@01DB1234"), Html));
        }

        [Theory]
        [InlineData("Scan.JPG", ".jpg")]
        [InlineData("photo.heic", ".heic")]
        [InlineData("odd.exe", ".img")]
        [InlineData(null, ".img")]
        public void SafeExtension_KeepsOnlyImageExtensions(string name, string expected)
        {
            Assert.Equal(expected, ImageAttachmentOrder.SafeExtension(name));
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile errors `CS0246 'AttachmentInfo'`, `CS0103 'ImageAttachmentOrder'`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/QuickTools/ImageAttachmentOrder.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class AttachmentInfo
    {
        public int Index { get; set; }
        public string FileName { get; set; }
        public string MimeTag { get; set; }
        public string ContentId { get; set; }
        public bool IsByValue { get; set; }
    }

    /// <summary>Which attachments are images, and in what order they become pages.</summary>
    internal static class ImageAttachmentOrder
    {
        public static readonly string[] ImageExtensions =
            { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".heif" };

        public static bool IsImage(AttachmentInfo a) =>
            a.IsByValue
            && (ImageExtensions.Contains(Extension(a.FileName))
                || (a.MimeTag ?? "").StartsWith("image/", StringComparison.OrdinalIgnoreCase));

        public static bool IsInline(AttachmentInfo a, string html) => Position(a, html) >= 0;

        /// <summary>Images the body shows, in the order it shows them; then the rest by attachment order.</summary>
        public static List<AttachmentInfo> Order(IEnumerable<AttachmentInfo> attachments, string html)
        {
            var images = attachments.Where(IsImage).ToList();
            var inline = images.Where(a => Position(a, html) >= 0).OrderBy(a => Position(a, html));
            var rest = images.Where(a => Position(a, html) < 0).OrderBy(a => a.Index);
            return inline.Concat(rest).ToList();
        }

        public static string SafeExtension(string fileName)
        {
            var ext = Extension(fileName);
            return ImageExtensions.Contains(ext) ? ext : ".img";
        }

        private static int Position(AttachmentInfo a, string html)
        {
            var id = (a.ContentId ?? "").Trim().TrimStart('<').TrimEnd('>');
            if (id.Length == 0 || string.IsNullOrEmpty(html)) return -1;
            return html.IndexOf("cid:" + id, StringComparison.OrdinalIgnoreCase);
        }

        private static string Extension(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "";
            try
            {
                return Path.GetExtension(fileName).ToLowerInvariant();
            }
            catch (ArgumentException)
            {
                return "";
            }
        }
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/EmailImageCollector.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using OutlookAI.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class CollectedImage
    {
        public CollectedImage(string path, string fileName, bool isInline)
        {
            Path = path;
            FileName = fileName ?? "";
            IsInline = isInline;
        }

        public string Path { get; }
        public string FileName { get; }
        public bool IsInline { get; }
    }

    internal sealed class CollectedImages
    {
        public CollectedImages(IReadOnlyList<CollectedImage> images, bool isRichText)
        {
            Images = images;
            IsRichText = isRichText;
        }

        public IReadOnlyList<CollectedImage> Images { get; }

        /// <summary>RTF emails carry pasted images as OLE attachments, which this version skips.</summary>
        public bool IsRichText { get; }
    }

    internal static class EmailImageCollector
    {
        private const string MimeTagProperty = "http://schemas.microsoft.com/mapi/proptag/0x370E001F";
        private const string ContentIdProperty = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";

        /// <summary>Saves the email's image attachments into the run folder, in page order. UI thread only.</summary>
        public static CollectedImages Collect(Outlook.MailItem mail, string runFolder)
        {
            var attachments = mail.Attachments;
            var infos = new List<AttachmentInfo>();
            for (var i = 1; i <= attachments.Count; i++)
            {
                var attachment = attachments[i];
                infos.Add(new AttachmentInfo
                {
                    Index = i,
                    FileName = Safe(() => attachment.FileName),
                    IsByValue = Safe(() => attachment.Type) == Outlook.OlAttachmentType.olByValue,
                    MimeTag = Property(attachment, MimeTagProperty),
                    ContentId = Property(attachment, ContentIdProperty),
                });
            }

            var html = Safe(() => mail.HTMLBody);
            var images = new List<CollectedImage>();
            var n = 0;
            foreach (var info in ImageAttachmentOrder.Order(infos, html))
            {
                n++;
                var path = Path.Combine(runFolder, "image-" + n.ToString("D2", CultureInfo.InvariantCulture)
                    + ImageAttachmentOrder.SafeExtension(info.FileName));
                try
                {
                    attachments[info.Index].SaveAsFile(path);
                }
                catch (COMException ex)
                {
                    TraceLog.Write("QuickTools: SaveAsFile failed (0x" + ex.HResult.ToString("X8", CultureInfo.InvariantCulture) + ")", "QuickTools");
                    continue;
                }
                images.Add(new CollectedImage(path, info.FileName, ImageAttachmentOrder.IsInline(info, html)));
            }

            var richText = Safe(() => mail.BodyFormat) == Outlook.OlBodyFormat.olFormatRichText;
            return new CollectedImages(images, richText);
        }

        private static string Property(Outlook.Attachment attachment, string tag)
        {
            try
            {
                return attachment.PropertyAccessor.GetProperty(tag) as string;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static T Safe<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (COMException)
            {
                return default(T);
            }
        }
    }
}
```

`OutlookAI.csproj`: insert `<Compile Include="Services\QuickTools\EmailImageCollector.cs" />` just before `<Compile Include="Services\QuickTools\Geometry.cs" />`, and `<Compile Include="Services\QuickTools\ImageAttachmentOrder.cs" />` just before `<Compile Include="Services\QuickTools\ImageLoader.cs" />`.

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `ImageAttachmentOrderTests`
Expected: 8 passed.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools/ImageAttachmentOrderTests.cs
git commit -m "feat(quick-tools): collect an email's images in page order" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 10: PagePlanner (the AI call per image)

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/PagePlanner.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/PagePlannerTests.cs`

**Interfaces:**
- Consumes: `IStructuredCompletion`, `CodexBackendException` (Task 3), `NotSignedInException` (existing, `Services/CodexAuthService.cs`), `PagePlan` (Task 2), `JpegFiles` (Task 4).
- Produces: `sealed class PagePlanner { const int MaxConcurrent = 3, AiMaxEdge = 1600; PagePlanner(IStructuredCompletion ai, Func<TimeSpan, CancellationToken, Task> delay = null, TimeSpan? timeout = null); Task<PagePlan> PlanAsync(byte[] jpeg, CancellationToken ct); static byte[] PrepareForAi(Image page); internal static TimeSpan RetryDelay(TimeSpan? retryAfter); internal const string Instructions, Request; }`. `PlanAsync` throws only `OperationCanceledException` (run cancelled) and `NotSignedInException`; every other failure returns `PagePlan.Fallback()`.

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/PagePlannerTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PagePlannerTests
    {
        private const string Plan =
            "{\"kind\":\"document_page\",\"reason\":\"page\",\"corners\":[[0.1,0.1],[0.9,0.1],[0.9,0.9],[0.1,0.9]],\"color\":\"grayscale\",\"strength\":\"normal\"}";

        private sealed class FakeAi : IStructuredCompletion
        {
            public readonly Queue<Func<CancellationToken, Task<string>>> Answers = new Queue<Func<CancellationToken, Task<string>>>();
            public readonly List<IReadOnlyList<JObject>> Parts = new List<IReadOnlyList<JObject>>();
            public string Effort;
            public string SchemaName;
            public int Calls;

            public Task<string> CompleteStructuredAsync(string instructions, IReadOnlyList<JObject> contentParts, string schemaName, JObject schema, string effort, CancellationToken cancellationToken)
            {
                Interlocked.Increment(ref Calls);
                lock (Parts) Parts.Add(contentParts);
                Effort = effort;
                SchemaName = schemaName;
                Func<CancellationToken, Task<string>> answer;
                lock (Answers) answer = Answers.Dequeue();
                return answer(cancellationToken);
            }
        }

        private static Func<CancellationToken, Task<string>> Returns(string text) => _ => Task.FromResult(text);

        private static Func<CancellationToken, Task<string>> Throws(Exception ex) => _ =>
        {
            var tcs = new TaskCompletionSource<string>();
            tcs.SetException(ex);
            return tcs.Task;
        };

        [Fact]
        public async Task PlanAsync_SendsTheImageAsADataUrl_AndParsesThePlan()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Returns(Plan));

            var plan = await new PagePlanner(ai).PlanAsync(new byte[] { 1, 2, 3 }, CancellationToken.None);

            Assert.True(plan.Analyzed);
            Assert.Equal("data:image/jpeg;base64,AQID", (string)ai.Parts[0][1]["image_url"]);
            Assert.Equal("input_text", (string)ai.Parts[0][0]["type"]);
            Assert.Equal("Low", ai.Effort);
            Assert.Equal(PagePlan.SchemaName, ai.SchemaName);
        }

        [Fact]
        public async Task PlanAsync_RetriesOnceAfterARateLimit_WaitingRetryAfter()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Throws(new CodexBackendException((HttpStatusCode)429, TimeSpan.FromSeconds(7), "slow down")));
            ai.Answers.Enqueue(Returns(Plan));
            var waited = new List<TimeSpan>();

            var plan = await new PagePlanner(ai, (d, ct) => { waited.Add(d); return Task.CompletedTask; })
                .PlanAsync(new byte[] { 1 }, CancellationToken.None);

            Assert.True(plan.Analyzed);
            Assert.Equal(new[] { TimeSpan.FromSeconds(7) }, waited);
            Assert.Equal(2, ai.Calls);
        }

        [Fact]
        public async Task PlanAsync_SecondRateLimit_FallsBack()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Throws(new CodexBackendException((HttpStatusCode)429, null, "slow down")));
            ai.Answers.Enqueue(Throws(new CodexBackendException((HttpStatusCode)429, null, "slow down")));

            var plan = await new PagePlanner(ai, (d, ct) => Task.CompletedTask).PlanAsync(new byte[] { 1 }, CancellationToken.None);

            Assert.False(plan.Analyzed);
            Assert.Equal(2, ai.Calls);
        }

        [Fact]
        public void RetryDelay_IsCappedAndDefaulted()
        {
            Assert.Equal(TimeSpan.FromSeconds(5), PagePlanner.RetryDelay(null));
            Assert.Equal(TimeSpan.FromSeconds(20), PagePlanner.RetryDelay(TimeSpan.FromSeconds(60)));
            Assert.Equal(TimeSpan.Zero, PagePlanner.RetryDelay(TimeSpan.FromSeconds(-3)));
        }

        [Fact]
        public async Task PlanAsync_BackendError_FallsBack()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Throws(new CodexBackendException(HttpStatusCode.InternalServerError, null, "boom")));

            var plan = await new PagePlanner(ai).PlanAsync(new byte[] { 1 }, CancellationToken.None);

            Assert.False(plan.Analyzed);
        }

        [Fact]
        public async Task PlanAsync_Timeout_FallsBack()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(async ct => { await Task.Delay(Timeout.Infinite, ct); return Plan; });

            var plan = await new PagePlanner(ai, timeout: TimeSpan.FromMilliseconds(50)).PlanAsync(new byte[] { 1 }, CancellationToken.None);

            Assert.False(plan.Analyzed);
        }

        [Fact]
        public async Task PlanAsync_NotSignedIn_Throws()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Throws(new NotSignedInException("not signed in", "Not signed in")));

            await Assert.ThrowsAsync<NotSignedInException>(() => new PagePlanner(ai).PlanAsync(new byte[] { 1 }, CancellationToken.None));
        }

        [Fact]
        public async Task PlanAsync_Cancelled_MakesNoCall()
        {
            var ai = new FakeAi();
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();

                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new PagePlanner(ai).PlanAsync(new byte[] { 1 }, cts.Token));
            }
            Assert.Equal(0, ai.Calls);
        }

        [Fact]
        public async Task PlanAsync_RunsAtMostThreeAtOnce()
        {
            var ai = new FakeAi();
            var release = new TaskCompletionSource<string>();
            var running = 0;
            var peak = 0;
            for (var i = 0; i < 5; i++)
            {
                ai.Answers.Enqueue(async ct =>
                {
                    var now = Interlocked.Increment(ref running);
                    lock (ai) peak = Math.Max(peak, now);
                    var text = await release.Task;
                    Interlocked.Decrement(ref running);
                    return text;
                });
            }
            var planner = new PagePlanner(ai);

            var tasks = Enumerable.Range(0, 5).Select(_ => Task.Run(() => planner.PlanAsync(new byte[] { 1 }, CancellationToken.None))).ToList();
            await Task.Delay(300);
            release.SetResult(Plan);
            await Task.WhenAll(tasks);

            Assert.Equal(PagePlanner.MaxConcurrent, peak);
        }

        [Fact]
        public void Instructions_CountIdsAsDocuments_AndNeverAlterContent()
        {
            Assert.Contains("ID or passport", PagePlanner.Instructions);
            Assert.Contains("never alter", PagePlanner.Instructions);
            Assert.Contains("top-left, top-right, bottom-right, bottom-left", PagePlanner.Instructions);
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile error `CS0246 'PagePlanner'`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/QuickTools/PagePlanner.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>Asks the AI for each image's page plan, 3 at a time.</summary>
    internal sealed class PagePlanner
    {
        public const int MaxConcurrent = 3;
        public const int AiMaxEdge = 1600;
        public static readonly TimeSpan PerImageTimeout = TimeSpan.FromSeconds(90);
        private static readonly TimeSpan DefaultRetryDelay = TimeSpan.FromSeconds(5);
        private static readonly TimeSpan MaxRetryDelay = TimeSpan.FromSeconds(20);

        internal const string Instructions =
            "You plan the cleanup of a photographed or scanned document for a program that applies your answer. "
            + "You never alter or retype content: the program only straightens, crops and adjusts exposure.\n"
            + "Decide whether the image is a document: a photo or scan of a page, card, ID or passport, form, receipt or screen, "
            + "even when it shows a person. Logos, signature images, decorations and photos with no document in them are not documents.\n"
            + "For a document, give the four corners of the document itself (not the background) as fractions of the image width "
            + "and height, in the order top-left, top-right, bottom-right, bottom-left as the text reads, so a sideways or "
            + "upside-down photo comes out upright. If the document fills the image or its edges aren't visible, give the image "
            + "corners in reading order. For anything else, give [[0,0],[1,0],[1,1],[0,1]].\n"
            + "Choose \"grayscale\" unless color carries meaning (stamps, highlights, colored ink, photos printed on the page), then \"color\".\n"
            + "Choose the strength from how dim, tinted or unevenly lit the image is: \"light\" for a clean scan, \"normal\" for a "
            + "typical phone photo, \"strong\" for a dark or shadowed one. Choose \"light\" when the document has large dark or "
            + "colored areas (a photo on an ID card, black bars), which a stronger cleanup would wash out.\n"
            + "Keep \"reason\" short: what the image is, e.g. \"lease page 2\" or \"company logo\".";

        internal const string Request = "Plan the cleanup of this image.";

        private readonly IStructuredCompletion _ai;
        private readonly Func<TimeSpan, CancellationToken, Task> _delay;
        private readonly TimeSpan _timeout;
        private readonly SemaphoreSlim _gate = new SemaphoreSlim(MaxConcurrent);

        public PagePlanner(IStructuredCompletion ai, Func<TimeSpan, CancellationToken, Task> delay = null, TimeSpan? timeout = null)
        {
            _ai = ai ?? throw new ArgumentNullException(nameof(ai));
            _delay = delay ?? Task.Delay;
            _timeout = timeout ?? PerImageTimeout;
        }

        /// <summary>The copy sent to the AI: at most 1600 px on the long edge, JPEG 85.</summary>
        public static byte[] PrepareForAi(Image page)
        {
            using (var small = JpegFiles.Scaled(page, AiMaxEdge))
            {
                return JpegFiles.Encode(small, JpegFiles.PageQuality);
            }
        }

        /// <summary>
        /// The image's plan, or the fallback plan when the AI fails, times out or
        /// answers implausibly. Throws only when the run is cancelled or nobody is
        /// signed in.
        /// </summary>
        public async Task<PagePlan> PlanAsync(byte[] jpeg, CancellationToken cancellationToken)
        {
            var parts = new List<JObject>
            {
                new JObject(new JProperty("type", "input_text"), new JProperty("text", Request)),
                new JObject(
                    new JProperty("type", "input_image"),
                    new JProperty("image_url", "data:image/jpeg;base64," + Convert.ToBase64String(jpeg))),
            };

            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                for (var attempt = 0; ; attempt++)
                {
                    try
                    {
                        using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                        {
                            timeout.CancelAfter(_timeout);
                            var text = await _ai.CompleteStructuredAsync(
                                Instructions, parts, PagePlan.SchemaName, PagePlan.Schema(), "Low", timeout.Token).ConfigureAwait(false);
                            return PagePlan.FromJson(text);
                        }
                    }
                    catch (CodexBackendException ex) when ((int)ex.StatusCode == 429 && attempt == 0)
                    {
                        await _delay(RetryDelay(ex.RetryAfter), cancellationToken).ConfigureAwait(false);
                    }
                }
            }
            catch (NotSignedInException)
            {
                throw;
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                TraceLog.Write("QuickTools: planning failed (" + ex.GetType().Name + ")", "QuickTools");
                return PagePlan.Fallback();
            }
            finally
            {
                _gate.Release();
            }
        }

        internal static TimeSpan RetryDelay(TimeSpan? retryAfter)
        {
            if (!retryAfter.HasValue) return DefaultRetryDelay;
            if (retryAfter.Value < TimeSpan.Zero) return TimeSpan.Zero;
            return retryAfter.Value > MaxRetryDelay ? MaxRetryDelay : retryAfter.Value;
        }
    }
}
```

`OutlookAI.csproj`: after `<Compile Include="Services\QuickTools\PagePlan.cs" />` add `<Compile Include="Services\QuickTools\PagePlanner.cs" />`.

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `PagePlannerTests`
Expected: 10 passed.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools/PagePlanner.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools/PagePlannerTests.cs
git commit -m "feat(quick-tools): ask the AI for each page's plan, with retry and fallback" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 11: Run folder, review pages, and the pipeline

**Files:**
- Create: `VSTO2/OutlookAI/Services/QuickTools/QuickToolsRunFolder.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/ReviewPage.cs`
- Create: `VSTO2/OutlookAI/Services/QuickTools/ImagesToPdfPipeline.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/ImagesToPdfPipelineTests.cs`

**Interfaces:**
- Consumes: `ImageLoader`, `JpegFiles` (Task 4), `PageCleaner` (Task 7), `PagePlanner.PrepareForAi`, `PagePlanner.AiMaxEdge` (Task 10), `CollectedImage` (Task 9), `PagePlan`, `PageKind`, `PageColor` (Task 2).
- Produces:
  - `static class QuickToolsRunFolder { string Root { get; set; } /* test seam */; string Create(DateTime nowUtc); void Delete(string dir); internal void PurgeOlderThan(DateTime cutoffUtc); }`
  - `sealed class ReviewPage { ReviewPage(int number, string label); int Number; string Label; bool Readable; bool OverLimit; string UprightPath; string CleanedPath; PagePlan Plan; bool Include; int QuarterTurns; bool UseOriginal; PageColor? ColorOverride; PageColor EffectiveColor; string Status; }`
  - `sealed class PipelineResult { IReadOnlyList<ReviewPage> Pages; int SkippedSmall; int NotProcessed; bool NoneAnalyzed; }`
  - `sealed class ImagesToPdfPipeline { const int MinShortSide = 200, MaxPages = 30, MaxParallelCleaning = 2; ImagesToPdfPipeline(Func<byte[], CancellationToken, Task<PagePlan>> plan, string runFolder); Task<PipelineResult> RunAsync(IReadOnlyList<CollectedImage> images, IProgress<string> progress, CancellationToken ct); static void Reclean(ReviewPage page); static List<string> FinalJpegs(IEnumerable<ReviewPage> chosenInOrder, string runFolder); }`

Status texts (spec § 4.3): `"page"`, `"not a document: <reason>"` (or `"not a document"` when the reason is empty), `"AI couldn't analyze this one"`, `"can't read this file"`, `"not processed (limit 30)"`.

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/Services/QuickTools/ImagesToPdfPipelineTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class ImagesToPdfPipelineTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-pipeline", Path.GetRandomFileName());

        public ImagesToPdfPipelineTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private CollectedImage AttachedImage(string name, int width, int height)
        {
            var path = Path.Combine(_dir, name);
            using (var page = SyntheticPages.Page(width, height, SyntheticPages.DimPaper, SyntheticPages.DimInk))
            {
                JpegFiles.Save(page, path, 90);
            }
            return new CollectedImage(path, name, isInline: true);
        }

        private static string PlanJson(string kind, string reason = "") =>
            "{\"kind\":\"" + kind + "\",\"reason\":\"" + reason + "\",\"corners\":[[0,0],[1,0],[1,1],[0,1]],\"color\":\"grayscale\",\"strength\":\"normal\"}";

        // Plans by the AI copy's width (concurrency makes call order nondeterministic).
        private static Func<byte[], CancellationToken, Task<PagePlan>> PlansByWidth(Func<int, string> json) => (jpeg, ct) =>
        {
            int w, h, c;
            JpegHeader.TryRead(jpeg, out w, out h, out c);
            return Task.FromResult(PagePlan.FromJson(json(w)));
        };

        [Fact]
        public async Task RunAsync_CleansEachPage_AndIncludesOnlyDocuments()
        {
            var images = new[] { AttachedImage("scan.jpg", 850, 1100), AttachedImage("logo.jpg", 400, 300) };
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => w == 400 ? PlanJson("not_document", "company logo") : PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(images, null, CancellationToken.None);

            Assert.Equal(2, result.Pages.Count);
            Assert.True(result.Pages[0].Include);
            Assert.Equal("page", result.Pages[0].Status);
            Assert.True(File.Exists(result.Pages[0].CleanedPath));
            Assert.True(File.Exists(result.Pages[0].UprightPath));
            Assert.False(result.Pages[1].Include);
            Assert.Equal("not a document: company logo", result.Pages[1].Status);
            Assert.False(result.NoneAnalyzed);
        }

        [Fact]
        public async Task RunAsync_SkipsSmallImages()
        {
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(new[] { AttachedImage("icon.jpg", 150, 150) }, null, CancellationToken.None);

            Assert.Empty(result.Pages);
            Assert.Equal(1, result.SkippedSmall);
        }

        [Fact]
        public async Task RunAsync_StopsAt30Pages()
        {
            var images = Enumerable.Range(1, 31).Select(i => AttachedImage("p" + i + ".jpg", 220, 220)).ToList();
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(images, null, CancellationToken.None);

            Assert.Equal(30, result.Pages.Count(p => p.Readable));
            Assert.Equal(1, result.NotProcessed);
            Assert.Equal("not processed (limit 30)", result.Pages.Last().Status);
        }

        [Fact]
        public async Task RunAsync_UnreadableFile_IsListed()
        {
            var bad = Path.Combine(_dir, "broken.jpg");
            File.WriteAllText(bad, "nope");
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(new[] { new CollectedImage(bad, "broken.jpg", false) }, null, CancellationToken.None);

            var page = Assert.Single(result.Pages);
            Assert.False(page.Readable);
            Assert.False(page.Include);
            Assert.Equal("can't read this file", page.Status);
        }

        [Fact]
        public async Task RunAsync_NoneAnalyzed_WhenEveryPlanFellBack()
        {
            var pipeline = new ImagesToPdfPipeline((jpeg, ct) => Task.FromResult(PagePlan.Fallback()), _dir);

            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400) }, null, CancellationToken.None);

            Assert.True(result.NoneAnalyzed);
            Assert.Equal("AI couldn't analyze this one", result.Pages[0].Status);
            Assert.True(result.Pages[0].Include);
        }

        [Fact]
        public async Task RunAsync_Cancelled_Throws()
        {
            using (var cts = new CancellationTokenSource())
            {
                var calls = 0;
                var pipeline = new ImagesToPdfPipeline((jpeg, ct) =>
                {
                    Interlocked.Increment(ref calls);
                    cts.Cancel();
                    ct.ThrowIfCancellationRequested();
                    return Task.FromResult(PagePlan.Fallback());
                }, _dir);

                await Assert.ThrowsAnyAsync<OperationCanceledException>(
                    () => pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400), AttachedImage("b.jpg", 300, 400) }, null, cts.Token));
                Assert.True(calls >= 1);
            }
        }

        [Fact]
        public async Task Reclean_InColor_IsNoLongerGray()
        {
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);
            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400) }, null, CancellationToken.None);
            var page = result.Pages[0];
            using (var stamped = SyntheticPages.Page(300, 400, SyntheticPages.DimPaper, SyntheticPages.DimInk, stamp: Color.FromArgb(200, 40, 40)))
            {
                JpegFiles.Save(stamped, page.UprightPath, 95);
            }

            page.ColorOverride = PageColor.Color;
            ImagesToPdfPipeline.Reclean(page);

            using (var cleaned = JpegFiles.Load(page.CleanedPath))
            {
                var px = PixelBuffer.FromBitmap(cleaned);
                var colored = 0;
                for (var i = 0; i < px.Data.Length; i += 3) if (px.Data[i + 2] - px.Data[i + 1] > 60) colored++;
                Assert.True(colored > 20);
            }
        }

        [Fact]
        public async Task FinalJpegs_RespectOrderRotationAndOriginals()
        {
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);
            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400), AttachedImage("b.jpg", 400, 300) }, null, CancellationToken.None);
            var first = result.Pages[0];
            var second = result.Pages[1];
            second.QuarterTurns = 1;
            first.UseOriginal = true;

            var files = ImagesToPdfPipeline.FinalJpegs(new[] { second, first }, _dir);

            Assert.Equal(2, files.Count);
            using (var rotated = JpegFiles.Load(files[0])) Assert.True(rotated.Height > rotated.Width);
            Assert.Equal(first.UprightPath, files[1]);
        }

        [Fact]
        public void RunFolder_IsCreated_AndOldOnesArePurged()
        {
            var root = Path.Combine(_dir, "root");
            var saved = QuickToolsRunFolder.Root;
            QuickToolsRunFolder.Root = root;
            try
            {
                var old = QuickToolsRunFolder.Create(new DateTime(2026, 10, 1, 9, 0, 0, DateTimeKind.Utc));
                Directory.SetCreationTimeUtc(old, new DateTime(2026, 10, 1, 9, 0, 0, DateTimeKind.Utc));

                var current = QuickToolsRunFolder.Create(new DateTime(2026, 10, 13, 9, 0, 0, DateTimeKind.Utc));

                Assert.True(Directory.Exists(current));
                Assert.False(Directory.Exists(old));
                QuickToolsRunFolder.Delete(current);
                Assert.False(Directory.Exists(current));
                QuickToolsRunFolder.Delete(current); // already gone: no exception
            }
            finally
            {
                QuickToolsRunFolder.Root = saved;
            }
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile errors `CS0246 'ImagesToPdfPipeline'`, `'ReviewPage'`, `CS0103 'QuickToolsRunFolder'`.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/QuickTools/QuickToolsRunFolder.cs`:

```csharp
using System;
using System.Globalization;
using System.IO;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// One temp folder per run under %LOCALAPPDATA%\OutlookAI\QuickTools, deleted
    /// when the run ends; leftovers older than a day are purged on the next run.
    /// </summary>
    internal static class QuickToolsRunFolder
    {
        // Test seam.
        public static string Root { get; set; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI", "QuickTools");

        public static string Create(DateTime nowUtc)
        {
            PurgeOlderThan(nowUtc.AddDays(-1));
            var dir = Path.Combine(Root, nowUtc.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)
                + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(dir);
            return dir;
        }

        public static void Delete(string dir)
        {
            try
            {
                if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        internal static void PurgeOlderThan(DateTime cutoffUtc)
        {
            if (!Directory.Exists(Root)) return;
            foreach (var dir in Directory.GetDirectories(Root))
            {
                if (Directory.GetCreationTimeUtc(dir) < cutoffUtc) Delete(dir);
            }
        }
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/ReviewPage.cs`:

```csharp
namespace OutlookAI.Services.QuickTools
{
    /// <summary>One page in the review window: files on disk plus the user's choices.</summary>
    internal sealed class ReviewPage
    {
        public ReviewPage(int number, string label)
        {
            Number = number;
            Label = label ?? "";
            Plan = PagePlan.Fallback();
        }

        public int Number { get; }

        /// <summary>The attachment's file name, plus "(page n)" for a multi-page file.</summary>
        public string Label { get; }

        public bool Readable { get; set; }
        public bool OverLimit { get; set; }
        public string UprightPath { get; set; }
        public string CleanedPath { get; set; }
        public PagePlan Plan { get; set; }
        public bool Include { get; set; }
        public int QuarterTurns { get; set; }
        public bool UseOriginal { get; set; }
        public PageColor? ColorOverride { get; set; }

        public PageColor EffectiveColor => ColorOverride ?? Plan.Color;

        public string Status
        {
            get
            {
                if (OverLimit) return "not processed (limit " + ImagesToPdfPipeline.MaxPages + ")";
                if (!Readable) return "can't read this file";
                if (!Plan.Analyzed) return "AI couldn't analyze this one";
                if (Plan.Kind == PageKind.NotDocument)
                    return string.IsNullOrEmpty(Plan.Reason) ? "not a document" : "not a document: " + Plan.Reason;
                return "page";
            }
        }
    }
}
```

`VSTO2/OutlookAI/Services/QuickTools/ImagesToPdfPipeline.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class PipelineResult
    {
        public PipelineResult(IReadOnlyList<ReviewPage> pages, int skippedSmall, int notProcessed)
        {
            Pages = pages;
            SkippedSmall = skippedSmall;
            NotProcessed = notProcessed;
        }

        public IReadOnlyList<ReviewPage> Pages { get; }
        public int SkippedSmall { get; }
        public int NotProcessed { get; }

        /// <summary>Every readable page fell back (AI unreachable or every answer unusable).</summary>
        public bool NoneAnalyzed =>
            Pages.Any(p => p.Readable) && Pages.Where(p => p.Readable).All(p => !p.Plan.Analyzed);
    }

    /// <summary>
    /// Decode → plan (AI) → clean, writing every image as files in the run folder.
    /// Runs off the UI thread; at most one full-size image per worker in memory.
    /// </summary>
    internal sealed class ImagesToPdfPipeline
    {
        public const int MinShortSide = 200;
        public const int MaxPages = 30;
        public const int MaxParallelCleaning = 2;

        private readonly Func<byte[], CancellationToken, Task<PagePlan>> _plan;
        private readonly string _runFolder;

        public ImagesToPdfPipeline(Func<byte[], CancellationToken, Task<PagePlan>> plan, string runFolder)
        {
            _plan = plan ?? throw new ArgumentNullException(nameof(plan));
            _runFolder = runFolder ?? throw new ArgumentNullException(nameof(runFolder));
        }

        public async Task<PipelineResult> RunAsync(IReadOnlyList<CollectedImage> images, IProgress<string> progress, CancellationToken cancellationToken)
        {
            progress?.Report("Reading images\u2026");
            var pages = new List<ReviewPage>();
            var aiCopies = new Dictionary<ReviewPage, byte[]>();
            var skipped = 0;
            var notProcessed = 0;

            foreach (var image in images)
            {
                cancellationToken.ThrowIfCancellationRequested();
                List<Bitmap> decoded;
                try
                {
                    decoded = ImageLoader.LoadPages(image.Path);
                }
                catch (Exception ex) when (ex is UnreadableImageException || ex is IOException || ex is OutOfMemoryException)
                {
                    TraceLog.Write("QuickTools: unreadable image (" + ex.GetType().Name + ")", "QuickTools");
                    pages.Add(new ReviewPage(pages.Count + 1, image.FileName));
                    continue;
                }

                try
                {
                    for (var f = 0; f < decoded.Count; f++)
                    {
                        var bitmap = decoded[f];
                        if (Math.Min(bitmap.Width, bitmap.Height) < MinShortSide)
                        {
                            skipped++;
                            continue;
                        }
                        var label = decoded.Count > 1 ? image.FileName + " (page " + (f + 1) + ")" : image.FileName;
                        if (aiCopies.Count >= MaxPages)
                        {
                            notProcessed++;
                            pages.Add(new ReviewPage(pages.Count + 1, label) { OverLimit = true });
                            continue;
                        }
                        var page = new ReviewPage(pages.Count + 1, label) { Readable = true };
                        page.UprightPath = Path.Combine(_runFolder, "upright-" + page.Number.ToString("D2", CultureInfo.InvariantCulture) + ".jpg");
                        JpegFiles.Save(bitmap, page.UprightPath, JpegFiles.OriginalQuality);
                        aiCopies[page] = PagePlanner.PrepareForAi(bitmap);
                        pages.Add(page);
                    }
                }
                finally
                {
                    foreach (var bitmap in decoded) bitmap.Dispose();
                }
            }

            var readable = pages.Where(p => p.Readable).ToList();
            var done = 0;
            if (readable.Count > 0) progress?.Report("Analyzing image 1 of " + readable.Count + "\u2026");
            using (var cleaning = new SemaphoreSlim(MaxParallelCleaning))
            {
                var work = readable.Select(async page =>
                {
                    var plan = await _plan(aiCopies[page], cancellationToken).ConfigureAwait(false);
                    page.Plan = plan;
                    page.Include = plan.Kind == PageKind.DocumentPage;
                    await cleaning.WaitAsync(cancellationToken).ConfigureAwait(false);
                    try
                    {
                        page.CleanedPath = Path.Combine(_runFolder, "cleaned-" + page.Number.ToString("D2", CultureInfo.InvariantCulture) + ".jpg");
                        await Task.Run(() => PageCleaner.WriteCleaned(page.UprightPath, plan, page.EffectiveColor, page.CleanedPath), cancellationToken).ConfigureAwait(false);
                    }
                    catch (Exception ex) when (!(ex is OperationCanceledException))
                    {
                        TraceLog.Write("QuickTools: cleanup failed (" + ex.GetType().Name + ")", "QuickTools");
                        page.Readable = false;
                        page.Include = false;
                    }
                    finally
                    {
                        cleaning.Release();
                    }
                    var finished = Interlocked.Increment(ref done);
                    progress?.Report(finished < readable.Count
                        ? "Analyzing image " + (finished + 1) + " of " + readable.Count + "\u2026"
                        : "Finishing\u2026");
                }).ToList();
                await Task.WhenAll(work).ConfigureAwait(false);
            }

            return new PipelineResult(pages, skipped, notProcessed);
        }

        /// <summary>Re-cleans one page after the review changed its color.</summary>
        public static void Reclean(ReviewPage page)
        {
            PageCleaner.WriteCleaned(page.UprightPath, page.Plan, page.EffectiveColor, page.CleanedPath);
        }

        /// <summary>The JPEG files for the PDF, in order: rotated copies where the review turned a page.</summary>
        public static List<string> FinalJpegs(IEnumerable<ReviewPage> chosenInOrder, string runFolder)
        {
            var files = new List<string>();
            foreach (var page in chosenInOrder)
            {
                var source = page.UseOriginal ? page.UprightPath : page.CleanedPath;
                if (page.QuarterTurns % 4 == 0)
                {
                    files.Add(source);
                    continue;
                }
                var rotated = Path.Combine(runFolder, "final-" + page.Number.ToString("D2", CultureInfo.InvariantCulture) + ".jpg");
                PageCleaner.WriteRotated(source, page.QuarterTurns, rotated);
                files.Add(rotated);
            }
            return files;
        }
    }
}
```

`OutlookAI.csproj`: after `<Compile Include="Services\QuickTools\ImagesToPdfOutput.cs" />` add `<Compile Include="Services\QuickTools\ImagesToPdfPipeline.cs" />`, and after `<Compile Include="Services\QuickTools\PixelBuffer.cs" />` add:

```xml
    <Compile Include="Services\QuickTools\QuickToolsRunFolder.cs" />
    <Compile Include="Services\QuickTools\ReviewPage.cs" />
```

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `ImagesToPdfPipelineTests`
Expected: 9 passed.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/QuickTools VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/Services/QuickTools/ImagesToPdfPipelineTests.cs
git commit -m "feat(quick-tools): run decode, AI plan and cleanup into a review list" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 12: Review window

**Files:**
- Create: `VSTO2/OutlookAI/QuickTools/ImagesToPdfReviewForm.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/QuickTools/ImagesToPdfReviewFormTests.cs`

**Interfaces:**
- Consumes: `ReviewPage`, `PageColor`, `PageCleaner.QuarterTurns`, `JpegFiles.LoadScaled` (Tasks 2, 4, 7, 11).
- Produces: `sealed class ImagesToPdfReviewForm : Form` (namespace `OutlookAI.QuickTools`) with constructor `(string heading, string note, IReadOnlyList<ReviewPage> pages, Func<ReviewPage, Task> reclean, Func<IWin32Window, IReadOnlyList<ReviewPage>, bool> save, Func<string, int, Bitmap> loadScaled = null)` and, for tests, `internal IReadOnlyList<ReviewPage> Pages; internal bool CanSave; internal void SelectPage(int index); internal void SetIncluded(int index, bool include); internal void MovePage(int delta); internal void Rotate(int quarterTurns); internal void SetUseOriginal(bool useOriginal); internal Task SetColorAsync(int choice /* 0 Auto, 1 Color, 2 Grayscale */); internal void SavePdf();`. `save` returns true when the PDF was saved (the window then closes with `DialogResult.OK`).

- [ ] **Step 1: Write the failing tests**

`VSTO2/OutlookAI.Tests/QuickTools/ImagesToPdfReviewFormTests.cs`:

```csharp
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.QuickTools;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.QuickTools
{
    public class ImagesToPdfReviewFormTests
    {
        private static ReviewPage Page(int number, string kind = "document_page", bool readable = true)
        {
            var page = new ReviewPage(number, "image" + number + ".jpg")
            {
                Readable = readable,
                Plan = PagePlan.FromJson("{\"kind\":\"" + kind + "\",\"reason\":\"company logo\",\"corners\":[[0,0],[1,0],[1,1],[0,1]],\"color\":\"grayscale\",\"strength\":\"normal\"}"),
                UprightPath = "upright.jpg",
                CleanedPath = "cleaned.jpg",
            };
            page.Include = readable && page.Plan.Kind == PageKind.DocumentPage;
            return page;
        }

        private static ImagesToPdfReviewForm NewForm(IReadOnlyList<ReviewPage> pages,
            System.Func<IWin32Window, IReadOnlyList<ReviewPage>, bool> save = null,
            System.Func<ReviewPage, Task> reclean = null) =>
            new ImagesToPdfReviewForm("\"Lease docs\" from Jane Example", null, pages,
                reclean ?? (p => Task.CompletedTask), save ?? ((o, p) => true), (path, size) => new Bitmap(20, 30));

        [Fact]
        public void Opens_WithDocumentsIncluded_AndOthersNot()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1), Page(2, "not_document") }))
                {
                    Assert.True(form.Pages[0].Include);
                    Assert.False(form.Pages[1].Include);
                    Assert.True(form.CanSave);
                }
            });
        }

        [Fact]
        public void Save_IsDisabled_WhenNoPageIsIncluded()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1), Page(2) }))
                {
                    form.SetIncluded(0, false);
                    form.SetIncluded(1, false);

                    Assert.False(form.CanSave);
                }
            });
        }

        [Fact]
        public void UnreadablePage_CantBeIncluded()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1, readable: false) }))
                {
                    form.SetIncluded(0, true);

                    Assert.False(form.Pages[0].Include);
                    Assert.False(form.CanSave);
                }
            });
        }

        [Fact]
        public void RotateAndOriginal_ApplyToTheSelectedPageOnly()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1), Page(2) }))
                {
                    form.SelectPage(0);
                    form.Rotate(1);
                    form.SetUseOriginal(true);

                    Assert.Equal(1, form.Pages[0].QuarterTurns);
                    Assert.True(form.Pages[0].UseOriginal);
                    Assert.Equal(0, form.Pages[1].QuarterTurns);
                    Assert.False(form.Pages[1].UseOriginal);
                    form.Rotate(3);
                    Assert.Equal(0, form.Pages[0].QuarterTurns);
                }
            });
        }

        [Fact]
        public void ColorChange_RecleansThatPage()
        {
            Sta.Run(() =>
            {
                var recleaned = new List<ReviewPage>();
                using (var form = NewForm(new[] { Page(1), Page(2) }, reclean: p => { recleaned.Add(p); return Task.CompletedTask; }))
                {
                    form.SelectPage(1);
                    form.SetColorAsync(1).GetAwaiter().GetResult();
                    form.SetColorAsync(1).GetAwaiter().GetResult(); // no change: no second clean

                    Assert.Equal(PageColor.Color, form.Pages[1].ColorOverride);
                    Assert.Same(form.Pages[1], Assert.Single(recleaned));
                }
            });
        }

        [Fact]
        public void Save_PassesTheIncludedPagesInOrder_AndCloses()
        {
            Sta.Run(() =>
            {
                IReadOnlyList<ReviewPage> saved = null;
                var a = Page(1);
                var b = Page(2);
                var c = Page(3, "not_document");
                using (var form = NewForm(new[] { a, b, c }, save: (owner, pages) => { saved = pages; return true; }))
                {
                    form.SelectPage(1);
                    form.MovePage(-1);
                    form.SavePdf();

                    Assert.Equal(new[] { b, a }, saved.ToArray());
                    Assert.Equal(DialogResult.OK, form.DialogResult);
                }
            });
        }

        [Fact]
        public void Save_StaysOpen_WhenSavingFails()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1) }, save: (owner, pages) => false))
                {
                    form.SavePdf();

                    Assert.Equal(DialogResult.None, form.DialogResult);
                    Assert.False(form.IsDisposed);
                }
            });
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build
Expected: compile error `CS0246 'ImagesToPdfReviewForm'` / namespace `OutlookAI.QuickTools` not found.

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/QuickTools/ImagesToPdfReviewForm.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Services.QuickTools;

namespace OutlookAI.QuickTools
{
    /// <summary>
    /// The pages before they become a PDF: include or drop, reorder, rotate,
    /// use the original, change color. Holds file paths, 160 px thumbnails and
    /// one preview, never full-size images.
    /// </summary>
    internal sealed class ImagesToPdfReviewForm : Form
    {
        internal const int ThumbSize = 160;
        internal const int PreviewSize = 1200;
        private const string Caption = "Clean up images \u2192 PDF";

        private readonly List<ReviewPage> _pages;
        private readonly Func<ReviewPage, Task> _reclean;
        private readonly Func<IWin32Window, IReadOnlyList<ReviewPage>, bool> _save;
        private readonly Func<string, int, Bitmap> _loadScaled;

        private readonly ListView _list;
        private readonly ImageList _thumbs;
        private readonly PictureBox _preview;
        private readonly Button _up;
        private readonly Button _down;
        private readonly Button _rotateLeft;
        private readonly Button _rotateRight;
        private readonly Button _cancel;
        private readonly Button _savePdf;
        private readonly RadioButton _cleaned;
        private readonly RadioButton _original;
        private readonly ComboBox _color;
        private int _selected = -1;
        private bool _updating;
        private bool _busy;

        public ImagesToPdfReviewForm(
            string heading,
            string note,
            IReadOnlyList<ReviewPage> pages,
            Func<ReviewPage, Task> reclean,
            Func<IWin32Window, IReadOnlyList<ReviewPage>, bool> save,
            Func<string, int, Bitmap> loadScaled = null)
        {
            _pages = pages.ToList();
            _reclean = reclean;
            _save = save;
            _loadScaled = loadScaled ?? JpegFiles.LoadScaled;

            Text = Caption + " \u2014 " + heading;
            Size = new Size(960, 680);
            MinimumSize = new Size(760, 520);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(250, 249, 248);
            ShowInTaskbar = false;
            MinimizeBox = false;

            _thumbs = new ImageList { ImageSize = new Size(ThumbSize, ThumbSize), ColorDepth = ColorDepth.Depth24Bit };
            _list = new ListView
            {
                Dock = DockStyle.Left,
                Width = 250,
                View = View.LargeIcon,
                CheckBoxes = true,
                MultiSelect = false,
                HideSelection = false,
                LargeImageList = _thumbs,
            };
            _preview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(64, 64, 64) };

            var pageControls = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(6), WrapContents = false };
            _up = MakeButton("Move up", 80);
            _down = MakeButton("Move down", 90);
            _rotateLeft = MakeButton("Rotate left", 90);
            _rotateRight = MakeButton("Rotate right", 95);
            _cleaned = new RadioButton { Text = "Cleaned", AutoSize = true, Checked = true, Margin = new Padding(16, 8, 0, 0) };
            _original = new RadioButton { Text = "Original", AutoSize = true, Margin = new Padding(8, 8, 0, 0) };
            var colorLabel = new Label { Text = "Color:", AutoSize = true, Margin = new Padding(16, 10, 0, 0) };
            _color = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100, Margin = new Padding(4, 6, 0, 0) };
            _color.Items.AddRange(new object[] { "Auto", "Color", "Grayscale" });
            pageControls.Controls.AddRange(new Control[] { _up, _down, _rotateLeft, _rotateRight, _cleaned, _original, colorLabel, _color });

            var dialogButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(6), FlowDirection = FlowDirection.RightToLeft };
            _savePdf = MakeButton("Save PDF\u2026", 110);
            _cancel = MakeButton("Cancel", 90);
            dialogButtons.Controls.AddRange(new Control[] { _savePdf, _cancel });

            Controls.Add(_preview);
            Controls.Add(_list);
            Controls.Add(pageControls);
            Controls.Add(dialogButtons);
            if (!string.IsNullOrEmpty(note))
            {
                Controls.Add(new Label { Text = note, Dock = DockStyle.Top, Height = 28, ForeColor = Color.DarkRed, Padding = new Padding(8, 6, 8, 0) });
            }

            foreach (var page in _pages)
            {
                _thumbs.Images.Add(Key(page), Thumbnail(page));
                _list.Items.Add(new ListViewItem(page.Number + ". " + page.Status, Key(page)) { Checked = page.Include, Tag = page });
            }

            _list.ItemCheck += (s, e) =>
            {
                if (_updating) return;
                if (!((ReviewPage)_list.Items[e.Index].Tag).Readable) e.NewValue = CheckState.Unchecked;
            };
            _list.ItemChecked += (s, e) =>
            {
                if (_updating) return;
                ((ReviewPage)e.Item.Tag).Include = e.Item.Checked;
                UpdateButtons();
            };
            _list.SelectedIndexChanged += (s, e) =>
            {
                if (_updating || _list.SelectedIndices.Count == 0) return;
                SelectPage(_list.SelectedIndices[0]);
            };
            _up.Click += (s, e) => MovePage(-1);
            _down.Click += (s, e) => MovePage(1);
            _rotateLeft.Click += (s, e) => Rotate(3);
            _rotateRight.Click += (s, e) => Rotate(1);
            _cleaned.CheckedChanged += (s, e) => { if (!_updating && _cleaned.Checked) SetUseOriginal(false); };
            _original.CheckedChanged += (s, e) => { if (!_updating && _original.Checked) SetUseOriginal(true); };
            _color.SelectedIndexChanged += async (s, e) => { if (!_updating) await SetColorAsync(_color.SelectedIndex); };
            _savePdf.Click += (s, e) => SavePdf();
            _cancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            if (_pages.Count > 0) SelectPage(0);
            UpdateButtons();
        }

        internal IReadOnlyList<ReviewPage> Pages => _pages;

        internal bool CanSave => _savePdf.Enabled;

        private ReviewPage Selected => _selected >= 0 && _selected < _pages.Count ? _pages[_selected] : null;

        internal void SelectPage(int index)
        {
            if (index < 0 || index >= _pages.Count) return;
            _selected = index;
            var page = _pages[index];
            _updating = true;
            _list.Items[index].Selected = true;
            _cleaned.Checked = !page.UseOriginal;
            _original.Checked = page.UseOriginal;
            _color.SelectedIndex = page.ColorOverride == null ? 0 : page.ColorOverride == PageColor.Color ? 1 : 2;
            _updating = false;
            ShowPreview(page);
            UpdateButtons();
        }

        internal void SetIncluded(int index, bool include)
        {
            var page = _pages[index];
            page.Include = include && page.Readable;
            _updating = true;
            _list.Items[index].Checked = page.Include;
            _updating = false;
            UpdateButtons();
        }

        internal void MovePage(int delta)
        {
            var from = _selected;
            var to = from + delta;
            if (from < 0 || to < 0 || to >= _pages.Count) return;
            var page = _pages[from];
            _pages.RemoveAt(from);
            _pages.Insert(to, page);
            _updating = true;
            var item = _list.Items[from];
            _list.Items.RemoveAt(from);
            _list.Items.Insert(to, item);
            item.Checked = page.Include;
            _updating = false;
            SelectPage(to);
        }

        internal void Rotate(int quarterTurns)
        {
            var page = Selected;
            if (page == null || !page.Readable) return;
            page.QuarterTurns = (page.QuarterTurns + quarterTurns) % 4;
            RefreshPage(page);
        }

        internal void SetUseOriginal(bool useOriginal)
        {
            var page = Selected;
            if (page == null || !page.Readable || page.UseOriginal == useOriginal) return;
            page.UseOriginal = useOriginal;
            RefreshPage(page);
        }

        internal async Task SetColorAsync(int choice)
        {
            var page = Selected;
            if (page == null || !page.Readable) return;
            var wanted = choice == 1 ? (PageColor?)PageColor.Color : choice == 2 ? PageColor.Grayscale : (PageColor?)null;
            if (wanted == page.ColorOverride) return;
            var previous = page.ColorOverride;
            var before = page.EffectiveColor;
            page.ColorOverride = wanted;
            if (page.EffectiveColor == before) return;

            _busy = true;
            UpdateButtons();
            try
            {
                await _reclean(page);
            }
            catch (Exception ex)
            {
                page.ColorOverride = previous;
                MessageBox.Show(this, "This page couldn't be redone in that color.\n\n" + ex.Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                _busy = false;
                UpdateButtons();
            }
            RefreshPage(page);
        }

        internal void SavePdf()
        {
            var chosen = _pages.Where(p => p.Include && p.Readable).ToList();
            if (chosen.Count == 0 || _busy) return;
            if (_save(this, chosen))
            {
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _preview.Image?.Dispose();
                _thumbs.Dispose();
            }
            base.Dispose(disposing);
        }

        private void RefreshPage(ReviewPage page)
        {
            var key = Key(page);
            _thumbs.Images.RemoveByKey(key);
            _thumbs.Images.Add(key, Thumbnail(page));
            var index = _pages.IndexOf(page);
            _list.Items[index].ImageKey = key;
            if (index == _selected) ShowPreview(page);
        }

        private void ShowPreview(ReviewPage page)
        {
            var old = _preview.Image;
            _preview.Image = Render(page, PreviewSize);
            old?.Dispose();
        }

        private Bitmap Thumbnail(ReviewPage page)
        {
            var thumb = new Bitmap(ThumbSize, ThumbSize);
            using (var g = Graphics.FromImage(thumb))
            {
                g.Clear(Color.Gainsboro);
                using (var image = Render(page, ThumbSize))
                {
                    if (image != null)
                    {
                        var scale = Math.Min((double)ThumbSize / image.Width, (double)ThumbSize / image.Height);
                        var w = (int)(image.Width * scale);
                        var h = (int)(image.Height * scale);
                        g.DrawImage(image, (ThumbSize - w) / 2, (ThumbSize - h) / 2, w, h);
                    }
                }
            }
            return thumb;
        }

        private Bitmap Render(ReviewPage page, int maxEdge)
        {
            if (!page.Readable) return null;
            var path = page.UseOriginal ? page.UprightPath : page.CleanedPath;
            try
            {
                var bitmap = _loadScaled(path, maxEdge);
                bitmap.RotateFlip(PageCleaner.QuarterTurns(page.QuarterTurns));
                return bitmap;
            }
            catch (Exception ex) when (ex is System.IO.IOException || ex is ArgumentException || ex is OutOfMemoryException)
            {
                return null;
            }
        }

        private void UpdateButtons()
        {
            var page = Selected;
            var editable = !_busy && page != null && page.Readable;
            _up.Enabled = !_busy && _selected > 0;
            _down.Enabled = !_busy && _selected >= 0 && _selected < _pages.Count - 1;
            _rotateLeft.Enabled = _rotateRight.Enabled = _cleaned.Enabled = _original.Enabled = _color.Enabled = editable;
            _savePdf.Enabled = !_busy && _pages.Any(p => p.Include && p.Readable);
        }

        private static string Key(ReviewPage page) => page.Number.ToString(System.Globalization.CultureInfo.InvariantCulture);

        private static Button MakeButton(string text, int width) => new Button
        {
            Text = text,
            Width = width,
            Height = 28,
            ForeColor = Color.Black,
            BackColor = SystemColors.ButtonFace,
            UseVisualStyleBackColor = false,
            Margin = new Padding(4),
        };
    }
}
```

`OutlookAI.csproj`: before `<Compile Include="Ribbon.cs" />` add:

```xml
    <Compile Include="QuickTools\ImagesToPdfReviewForm.cs">
      <SubType>Form</SubType>
    </Compile>
```

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `ImagesToPdfReviewFormTests`
Expected: 7 passed.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/QuickTools VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/QuickTools
git commit -m "feat(quick-tools): review window for the cleaned pages" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 13: Command, progress window and ribbon wiring

**Files:**
- Create: `VSTO2/OutlookAI/QuickTools/QuickToolsProgressForm.cs`
- Create: `VSTO2/OutlookAI/QuickTools/OwnerWindow.cs`
- Create: `VSTO2/OutlookAI/QuickTools/ImagesToPdfCommand.cs`
- Modify: `VSTO2/OutlookAI/Services/CodexAuthService.cs` (message constant)
- Modify: `VSTO2/OutlookAI/ThisAddIn.cs` (`RunImagesToPdf`)
- Modify: `VSTO2/OutlookAI/Ribbon.cs`, `VSTO2/OutlookAI/Ribbon.xml`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Test: `VSTO2/OutlookAI.Tests/RibbonTests.cs`, `VSTO2/OutlookAI.Tests/QuickTools/ImagesToPdfCommand_SourceTests.cs`

**Interfaces:**
- Consumes: everything above; `Globals.ThisAddIn.AuthService`, `.ChatService`, `.ExportPathResolver`; `ExportFilenameSanitizer.Build`; `UpdateManifest.LoadFromInstallDir().Tag`; `TraceLog`.
- Produces: `CodexAuthService.NotSignedInMessage` (public const string); `ThisAddIn.RunImagesToPdf(object context)`; `Ribbon.OnImagesToPdfClick(IRibbonControl)`; `static class ImagesToPdfCommand { static async void Run(object context); }`.

- [ ] **Step 1: Write the failing tests**

In `VSTO2/OutlookAI.Tests/RibbonTests.cs`, replace `AiAssistantGroup` with a version that also maps the opened-email ribbon, add the read tab to the Settings theory, and add the Quick Tools tests:

```csharp
        private static string RibbonIdFor(string tab) =>
            tab == "TabMail" ? "Microsoft.Outlook.Explorer"
            : tab == "TabReadMessage" ? "Microsoft.Outlook.Mail.Read"
            : "Microsoft.Outlook.Mail.Compose";

        private static XElement AiAssistantGroup(string tab) =>
            Served(RibbonIdFor(tab))
                .Descendants(Ui + "tab").Single(t => (string)t.Attribute("idMso") == tab)
                .Elements(Ui + "group").Single(g => (string)g.Attribute("label") == "AI Assistant");
```

Add `[InlineData("TabReadMessage")]` to `SettingsButton_IsInTheAiAssistantGroup`, then:

```csharp
        [Theory]
        [InlineData("TabMail")]
        [InlineData("TabReadMessage")]
        public void QuickToolsMenu_HasCleanUpImagesToPdf(string tab)
        {
            var menu = AiAssistantGroup(tab).Elements(Ui + "menu")
                .SingleOrDefault(m => (string)m.Attribute("label") == "Quick Tools");
            Assert.NotNull(menu);
            Assert.Equal("ControlToolboxOutlook", (string)menu.Attribute("imageMso"));

            var item = menu.Elements(Ui + "button").SingleOrDefault(b => (string)b.Attribute("onAction") == "OnImagesToPdfClick");
            Assert.NotNull(item);
            Assert.Equal("Clean up images \u2192 PDF", (string)item.Attribute("label"));
            Assert.Equal("PictureCrop", (string)item.Attribute("imageMso"));
        }

        [Fact]
        public void ComposeWindow_HasNoQuickTools()
        {
            Assert.Empty(AiAssistantGroup("TabNewMailMessage").Elements(Ui + "menu"));
        }
```

`VSTO2/OutlookAI.Tests/QuickTools/ImagesToPdfCommand_SourceTests.cs`:

```csharp
using System;
using System.IO;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.QuickTools
{
    // The command drives Outlook COM and dialogs, so these pin its structure.
    public class ImagesToPdfCommand_SourceTests
    {
        private static string Source =>
            File.ReadAllText(RepoFiles.Find("OutlookAI", "QuickTools", "ImagesToPdfCommand.cs")).Replace("\r\n", "\n");

        // Decoding big photos on Outlook's UI thread would freeze it.
        [Fact]
        public void Pipeline_RunsOffTheUiThread()
        {
            Assert.Contains("await Task.Run(() => pipeline.RunAsync(", Source);
        }

        [Fact]
        public void RunFolder_IsDeletedHoweverTheRunEnds()
        {
            var src = Source;
            var finallyAt = src.LastIndexOf("finally", StringComparison.Ordinal);
            Assert.True(finallyAt > 0);
            Assert.True(src.IndexOf("QuickToolsRunFolder.Delete(runFolder)", finallyAt, StringComparison.Ordinal) > finallyAt);
        }

        [Fact]
        public void Cancel_EndsTheRunQuietly()
        {
            Assert.Contains("catch (OperationCanceledException) when (cts.IsCancellationRequested)", Source);
        }

        // The progress window is created before the first await, which installs
        // the WinForms synchronization context on a fresh Outlook UI thread.
        [Fact]
        public void ProgressWindow_ExistsBeforeTheFirstAwait()
        {
            var src = Source;
            Assert.True(src.IndexOf("new QuickToolsProgressForm(", StringComparison.Ordinal)
                < src.IndexOf("await ", StringComparison.Ordinal));
        }
    }
}
```

- [ ] **Step 2: Run to verify they fail**

Run: Build; Test `RibbonTests|FullyQualifiedName~ImagesToPdfCommand_SourceTests` (filter: `/TestCaseFilter:"FullyQualifiedName~RibbonTests|FullyQualifiedName~ImagesToPdfCommand_SourceTests"`)
Expected: the read-tab cases fail (`GetCustomUI` returns null for `Microsoft.Outlook.Mail.Read` → `ArgumentNullException`), the Quick Tools cases fail (no menu), the source tests fail (`FileNotFoundException`).

- [ ] **Step 3: Implement**

`VSTO2/OutlookAI/Services/CodexAuthService.cs`: add to the `CodexAuthService` class

```csharp
        public const string NotSignedInMessage =
            "OutlookAI is not signed in to ChatGPT. An admin signs in once for everyone: Settings in the AI Assistant group on the ribbon.";
```

and use it at the existing throw site:

```csharp
                    throw new NotSignedInException(NotSignedInMessage, "Not signed in");
```

`VSTO2/OutlookAI/QuickTools/OwnerWindow.cs`:

```csharp
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OutlookAI.QuickTools
{
    /// <summary>An Outlook Explorer or Inspector as the owner of our windows.</summary>
    internal sealed class OwnerWindow : IWin32Window
    {
        [ComImport]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleWindow
        {
            void GetWindow(out IntPtr phwnd);

            void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
        }

        private OwnerWindow(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; }

        public static IWin32Window For(object outlookWindow)
        {
            var ole = outlookWindow as IOleWindow;
            if (ole == null) return null;
            try
            {
                IntPtr handle;
                ole.GetWindow(out handle);
                return handle == IntPtr.Zero ? null : new OwnerWindow(handle);
            }
            catch (COMException)
            {
                return null;
            }
        }
    }
}
```

`VSTO2/OutlookAI/QuickTools/QuickToolsProgressForm.cs`:

```csharp
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace OutlookAI.QuickTools
{
    internal sealed class QuickToolsProgressForm : Form
    {
        private readonly Label _status;
        private readonly Button _cancel;
        private readonly CancellationTokenSource _cts;
        private bool _finished;

        public QuickToolsProgressForm(CancellationTokenSource cts)
        {
            _cts = cts;
            Text = "Clean up images \u2192 PDF";
            Size = new Size(400, 140);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(250, 249, 248);

            _status = new Label { Text = "Reading images\u2026", Location = new Point(20, 20), Size = new Size(350, 24) };
            _cancel = new Button
            {
                Text = "Cancel",
                Location = new Point(280, 56),
                Width = 90,
                Height = 28,
                ForeColor = Color.Black,
                BackColor = SystemColors.ButtonFace,
                UseVisualStyleBackColor = false,
            };
            _cancel.Click += (s, e) => Cancel();
            // The window's X cancels the run; only Finish() (the run ending) closes it.
            FormClosing += (s, e) =>
            {
                if (_finished) return;
                e.Cancel = true;
                Cancel();
            };
            Controls.AddRange(new Control[] { _status, _cancel });
        }

        public void SetStatus(string text)
        {
            if (!IsDisposed) _status.Text = text;
        }

        public void Finish()
        {
            _finished = true;
            if (!IsDisposed) Close();
        }

        private void Cancel()
        {
            if (_cts.IsCancellationRequested) return;
            _cancel.Enabled = false;
            _status.Text = "Cancelling\u2026";
            _cts.Cancel();
        }
    }
}
```

`VSTO2/OutlookAI/QuickTools/ImagesToPdfCommand.cs`:

```csharp
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Export;
using OutlookAI.Services.QuickTools;
using OutlookAI.Services.Updates;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.QuickTools
{
    /// <summary>Quick Tools → Clean up images → PDF, from the Inbox or an opened email.</summary>
    internal static class ImagesToPdfCommand
    {
        private const string Caption = "Clean up images \u2192 PDF";
        private const string NoImages = "No images found in this email.";
        private const string NoImagesRichText = "No images found. Images pasted into Rich Text emails aren't supported yet.";

        // The progress or review window of the run in progress (one run at a time).
        private static Form _active;

        public static async void Run(object context)
        {
            if (_active != null)
            {
                try { _active.Activate(); } catch (ObjectDisposedException) { }
                return;
            }
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            var owner = OwnerWindow.For(context);
            string runFolder = null;
            try
            {
                var mail = FindMail(context, owner);
                if (mail == null) return;

                var addIn = Globals.ThisAddIn;
                var auth = addIn?.AuthService;
                var chat = addIn?.ChatService;
                if (auth == null || chat == null)
                {
                    Info(owner, "OutlookAI isn't ready. Restart Outlook and try again.");
                    return;
                }
                if (auth.GetStatus().State != AuthState.Authenticated)
                {
                    Info(owner, CodexAuthService.NotSignedInMessage);
                    return;
                }

                var subject = Safe(() => mail.Subject) ?? "";
                var sender = Safe(() => mail.SenderName) ?? "";
                runFolder = QuickToolsRunFolder.Create(DateTime.UtcNow);
                var collected = EmailImageCollector.Collect(mail, runFolder);
                if (collected.Images.Count == 0)
                {
                    Info(owner, collected.IsRichText ? NoImagesRichText : NoImages);
                    return;
                }

                PipelineResult result;
                using (var cts = new CancellationTokenSource())
                using (var progressForm = new QuickToolsProgressForm(cts))
                {
                    _active = progressForm;
                    progressForm.Show(owner);
                    var progress = new Progress<string>(progressForm.SetStatus);
                    var planner = new PagePlanner(chat);
                    var pipeline = new ImagesToPdfPipeline(planner.PlanAsync, runFolder);
                    var started = Stopwatch.StartNew();
                    try
                    {
                        result = await Task.Run(() => pipeline.RunAsync(collected.Images, progress, cts.Token));
                    }
                    catch (OperationCanceledException) when (cts.IsCancellationRequested)
                    {
                        TraceLog.Write("QuickTools: cancelled", "QuickTools");
                        return;
                    }
                    finally
                    {
                        _active = null;
                        progressForm.Finish();
                    }
                    TraceLog.Write("QuickTools: " + result.Pages.Count + " pages, " + result.SkippedSmall + " small skipped, "
                        + started.Elapsed.TotalSeconds.ToString("F1", CultureInfo.InvariantCulture) + " s", "QuickTools");
                }

                if (result.Pages.Count == 0)
                {
                    Info(owner, NoImages + (result.SkippedSmall > 0 ? " (" + result.SkippedSmall + " small images skipped.)" : ""));
                    return;
                }

                var note = result.NoneAnalyzed
                    ? "ChatGPT couldn't analyze these images, so they have basic cleanup only. Check each page before saving."
                    : null;
                var heading = "\"" + subject + "\" from " + sender;
                var folder = runFolder;
                using (var review = new ImagesToPdfReviewForm(heading, note, result.Pages,
                    page => Task.Run(() => ImagesToPdfPipeline.Reclean(page)),
                    (window, pages) => SavePdf(window, pages, subject, sender, folder)))
                {
                    _active = review;
                    try
                    {
                        review.ShowDialog(owner);
                    }
                    finally
                    {
                        _active = null;
                    }
                }
            }
            catch (NotSignedInException ex)
            {
                Info(owner, ex.Message);
            }
            catch (Exception ex)
            {
                TraceLog.Write("QuickTools: failed (" + ex.GetType().Name + ")", "QuickTools");
                MessageBox.Show(owner, "Error: " + ex.Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (runFolder != null) QuickToolsRunFolder.Delete(runFolder);
            }
        }

        private static Outlook.MailItem FindMail(object context, IWin32Window owner)
        {
            object item = null;
            var explorer = context as Outlook.Explorer;
            if (explorer != null)
            {
                Outlook.Selection selection = null;
                try { selection = explorer.Selection; } catch (COMException) { }
                if (selection == null || selection.Count != 1)
                {
                    Info(owner, "Select one email first.");
                    return null;
                }
                item = selection[1];
            }
            else
            {
                var inspector = context as Outlook.Inspector;
                if (inspector != null)
                {
                    try { item = inspector.CurrentItem; } catch (COMException) { }
                }
            }

            var mail = item as Outlook.MailItem;
            if (mail == null) Info(owner, item == null ? "Select one email first." : "This works on email messages.");
            return mail;
        }

        private static bool SavePdf(IWin32Window owner, System.Collections.Generic.IReadOnlyList<ReviewPage> pages, string subject, string sender, string runFolder)
        {
            var resolver = Globals.ThisAddIn?.ExportPathResolver ?? new ExportPathResolver();
            string folder = null;
            try
            {
                resolver.EnsureExists();
                folder = resolver.ResolveBaseDir();
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "Save PDF";
                dialog.Filter = "PDF files (*.pdf)|*.pdf";
                dialog.DefaultExt = "pdf";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                if (folder != null) dialog.InitialDirectory = folder;
                dialog.FileName = ExportFilenameSanitizer.Build(sender + " - " + subject, ".pdf", DateTimeOffset.Now,
                    name => folder != null && File.Exists(Path.Combine(folder, name)));
                if (dialog.ShowDialog(owner) != DialogResult.OK) return false;

                var temp = Path.Combine(runFolder, "output.pdf");
                ImagePdfWriter.WriteFile(ImagesToPdfPipeline.FinalJpegs(pages, runFolder), subject, Producer(), temp);
                var error = ImagesToPdfOutput.Publish(temp, dialog.FileName);
                if (error != null)
                {
                    MessageBox.Show(owner, error, Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                TraceLog.Write("QuickTools: saved " + pages.Count + " pages", "QuickTools");
                try
                {
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
                catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException)
                {
                    Info(owner, "Saved to " + dialog.FileName + ", but it couldn't be opened automatically.");
                }
                return true;
            }
        }

        private static string Producer()
        {
            try
            {
                var tag = UpdateManifest.LoadFromInstallDir().Tag;
                return string.IsNullOrEmpty(tag) ? "OutlookAI" : "OutlookAI " + tag;
            }
            catch (Exception)
            {
                return "OutlookAI";
            }
        }

        private static void Info(IWin32Window owner, string message) =>
            MessageBox.Show(owner, message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);

        private static T Safe<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (COMException)
            {
                return default(T);
            }
        }
    }
}
```

`VSTO2/OutlookAI/ThisAddIn.cs`, after `ShowSettings()`:

```csharp
        /// <summary>
        /// Quick Tools → Clean up images → PDF (Home tab and opened emails).
        /// </summary>
        public void RunImagesToPdf(object context)
        {
            using (TraceLog.Scope("RunImagesToPdf", "ThisAddIn"))
                OutlookAI.QuickTools.ImagesToPdfCommand.Run(context);
        }
```

`VSTO2/OutlookAI/Ribbon.cs`: in `GetCustomUI`, serve the opened-email ribbon too:

```csharp
            if (ribbonID == "Microsoft.Outlook.Mail.Compose" ||
                ribbonID == "Microsoft.Outlook.Explorer" ||
                ribbonID == "Microsoft.Outlook.Mail.Read")
```

and add the handler after `OnSettingsClick`:

```csharp
        public void OnImagesToPdfClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.RunImagesToPdf(control.Context);
        }
```

`VSTO2/OutlookAI/Ribbon.xml`: inside the `TabMail` group, after the `btnSettingsExplorer` button:

```xml
					<menu id="mnuQuickToolsExplorer"
						  label="Quick Tools"
						  size="large"
						  imageMso="ControlToolboxOutlook"
						  supertip="Quick tools for the selected email.">
						<button id="btnImagesToPdfExplorer"
								label="Clean up images &#x2192; PDF"
								imageMso="PictureCrop"
								onAction="OnImagesToPdfClick"
								supertip="Straighten, crop and clean up the images in the selected email, review them, and save them as one PDF." />
					</menu>
```

and a new tab after the `TabMail` tab (before `</tabs>`):

```xml
			<!-- Opened email (read inspector): Quick Tools and Settings. -->
			<tab idMso="TabReadMessage">
				<group id="AIAssistantReadGroup" label="AI Assistant">
					<menu id="mnuQuickToolsRead"
						  label="Quick Tools"
						  size="large"
						  imageMso="ControlToolboxOutlook"
						  supertip="Quick tools for this email.">
						<button id="btnImagesToPdfRead"
								label="Clean up images &#x2192; PDF"
								imageMso="PictureCrop"
								onAction="OnImagesToPdfClick"
								supertip="Straighten, crop and clean up the images in this email, review them, and save them as one PDF." />
					</menu>
					<button id="btnSettingsRead"
							label="Settings"
							size="large"
							onAction="OnSettingsClick"
							supertip="Open OutlookAI Settings (admin password): ChatGPT sign-in, model, reasoning effort, write tools and updates."
							imageMso="CurrentViewSettings" />
				</group>
			</tab>
```

`OutlookAI.csproj`, before `<Compile Include="QuickTools\ImagesToPdfReviewForm.cs">` add `<Compile Include="QuickTools\ImagesToPdfCommand.cs" />`, and after the review form's entry add:

```xml
    <Compile Include="QuickTools\OwnerWindow.cs" />
    <Compile Include="QuickTools\QuickToolsProgressForm.cs">
      <SubType>Form</SubType>
    </Compile>
```

- [ ] **Step 4: Run to verify they pass**

Run: Build; Test `RibbonTests|FullyQualifiedName~ImagesToPdfCommand_SourceTests|FullyQualifiedName~CodexAuthServiceTests`
Expected: all pass (the existing `CodexAuthServiceTests` still find the ribbon text in the message).

- [ ] **Step 5: Full suite**

Run: Build; Test (no filter)
Expected: all pass (908 + the new tests).

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI VSTO2/OutlookAI.Tests
git commit -m "feat(quick-tools): Clean up images to PDF on the Inbox and opened-email ribbons" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 14: Opt-in live test against the real backend

**Files:**
- Test: `VSTO2/OutlookAI.Tests/Services/QuickTools/PagePlannerLiveTests.cs`

**Interfaces:**
- Consumes: `PagePlanner`, `PagePlanner.PrepareForAi` (Task 10), `SyntheticPages` (Task 5), `CodexAuthService(string)`, `CodexChatService(CodexAuthService)`.

xUnit 2 has no runtime skip, so the test returns early (and passes) unless `OUTLOOKAI_LIVE_TESTS=1`. It uses the real token file, never a copy: `CodexAuthService` saves refreshed tokens to the file it was given and refresh tokens are single-use.

- [ ] **Step 1: Write the test**

```csharp
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    [Collection("Config")]
    public class PagePlannerLiveTests : IDisposable
    {
        private readonly ConfigStateScope _scope = new ConfigStateScope();

        public void Dispose()
        {
            _scope.Dispose();
        }

        [Fact]
        public async Task Live_FindsTheCornersOfAPhotographedPage()
        {
            if (Environment.GetEnvironmentVariable("OUTLOOKAI_LIVE_TESTS") != "1") return;
            var authPath = Environment.GetEnvironmentVariable("OUTLOOKAI_LIVE_AUTH")
                ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "auth.json");
            Config.Model = "gpt-5.6-sol";
            var corners = new[] { new PointD(470, 140), new PointD(1150, 190), new PointD(1290, 1080), new PointD(330, 1010) };

            using (var page = SyntheticPages.Page(850, 1100, SyntheticPages.DimPaper, SyntheticPages.DimInk, marker: true))
            using (var photo = SyntheticPages.Photograph(page, 1600, 1200, corners, SyntheticPages.Desk))
            using (var auth = new CodexAuthService(authPath))
            using (var chat = new CodexChatService(auth))
            {
                var plan = await new PagePlanner(chat).PlanAsync(PagePlanner.PrepareForAi(photo), CancellationToken.None);

                Assert.True(plan.Analyzed);
                Assert.Equal(PageKind.DocumentPage, plan.Kind);
                for (var i = 0; i < 4; i++)
                {
                    Assert.InRange(plan.Corners[i].X, corners[i].X / 1600 - 0.03, corners[i].X / 1600 + 0.03);
                    Assert.InRange(plan.Corners[i].Y, corners[i].Y / 1200 - 0.03, corners[i].Y / 1200 + 0.03);
                }
            }
        }
    }
}
```

- [ ] **Step 2: Run it offline (expected: passes without calling anything)**

Run: Build; Test `PagePlannerLiveTests`
Expected: 1 passed.

- [ ] **Step 3: Run it live once**

Run (PowerShell): `$env:OUTLOOKAI_LIVE_TESTS = "1"`, then Test `PagePlannerLiveTests`, then `Remove-Item Env:OUTLOOKAI_LIVE_TESTS`.
Expected: 1 passed in a few seconds. A failure here means a real backend change (images, schema, or model): read the assertion, don't widen the tolerance.

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI.Tests/Services/QuickTools/PagePlannerLiveTests.cs
git commit -m "test(quick-tools): opt-in live check of page planning" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 15: Docs

**Files:**
- Modify: `README.md` (feature section after Settings; Status list; tests badge and Contributing count)
- Modify: `Deploy/README.txt` (new QUICK TOOLS section; ASCII only)
- Modify: `docs/superpowers/checklists/phase-2-and-3a-smoke.md` (Quick Tools smoke)
- Modify: `CLAUDE.md` (architecture bullet, lookup row)

- [ ] **Step 1: README**

After the `### ⚙️ Settings` section, add:

```markdown
### 🧰 Quick Tools: Clean up images → PDF

Clients often send scans as phone photos pasted into an email. **Quick Tools →
Clean up images → PDF** (Home tab, or the Message tab of an opened email)
takes the email's images, pasted in or attached, and:

- has ChatGPT judge each one (a document page, or a logo or signature image
  to leave out) and find the page's four corners;
- flattens each page from those corners to its real proportions, evens out
  shadows, sets the paper to white and the ink to black, and sharpens it. The
  AI never redraws a page, so it can't change a word;
- shows a review window to drop, reorder or rotate pages, switch a page back
  to the original, or keep it in color;
- saves one PDF where you choose (starting in the Reports folder) and opens it.

Images are sent to ChatGPT (downscaled) the same way email text is. Pages of
PDF attachments and images pasted into Rich Text emails aren't handled yet.
```

In "Status and roadmap", add after the v2.2.3 line:

```markdown
- **v2.3.0:** Quick Tools → Clean up images → PDF.
```

Update the tests badge and the "# Test (N tests, all passing)" line in Contributing to the new total from Task 13 Step 5.

- [ ] **Step 2: Deploy/README.txt (ASCII only)**

Add a section after UPDATES:

```text
QUICK TOOLS
-----------
Quick Tools -> "Clean up images -> PDF" (Home tab on the selected email, or
the Message tab of an opened email) turns the images in an email (pasted in
or attached) into one PDF of flat, upright, high-contrast pages:

- ChatGPT judges each image (document page, or a logo/signature to leave
  out) and finds the page's corners. It never redraws a page.
- The add-in flattens and cleans each page, then shows a review window
  (drop, reorder, rotate, original vs cleaned, color).
- Save PDF asks where to save (starting in the Reports folder) and opens it.

Images are sent to ChatGPT, downscaled, through the same sign-in as the
rest of OutlookAI, like email text already is. Working copies go to
%LOCALAPPDATA%\OutlookAI\QuickTools and are deleted when the window closes.
Any user can run it; it doesn't need the admin password.
```

Verify: `python -I -c "d=open(r'Deploy/README.txt','rb').read(); print(sum(b>127 for b in d))"` prints `0`.

- [ ] **Step 3: Smoke checklist**

Append to `docs/superpowers/checklists/phase-2-and-3a-smoke.md`:

```markdown
---

# Quick Tools Smoke (Clean up images → PDF)

1. [ ] Home tab → AI Assistant → Quick Tools → Clean up images → PDF with no
   email selected → "Select one email first."
2. [ ] Select an email with 2–3 pasted phone photos of paper, one attached
   JPEG and a signature logo → run it. Progress shows "Analyzing image n of m";
   the review window lists every image, with the logo unchecked and labelled.
3. [ ] Pages are flat, upright and readable (paper white, ink black).
4. [ ] Move a page up, rotate one, switch one to Original, set one to Color
   (it re-cleans). Save PDF… → the dialog starts in the Reports folder; save
   to a network folder; the PDF opens with the pages in that order.
5. [ ] Open the PDF in a viewer, run again and save over it → a "close it"
   message; the review window stays open.
6. [ ] Double-click an email to open it → Message tab → AI Assistant →
   Quick Tools works there too; Settings opens from that group.
7. [ ] Cancel during "Analyzing…" → nothing else appears;
   `%LOCALAPPDATA%\OutlookAI\QuickTools` has no folder left for that run.
8. [ ] `%LOCALAPPDATA%\OutlookAI\trace.log` has QuickTools lines with counts
   and durations only (no subjects, file names or AI text).
```

- [ ] **Step 4: CLAUDE.md**

In Architecture, after the `Services/Updates/` mention, add `Services/QuickTools/` + `QuickTools/` (Clean up images → PDF: image collection, AI page plans, perspective cleanup, review window, PDF writer). In the lookup table add:

```markdown
| Quick Tools (images → PDF) | `docs/superpowers/specs/2026-10-13-quick-tools-images-to-pdf-design.md` | `Services/QuickTools/`, `QuickTools/` |
```

- [ ] **Step 5: Commit**

```powershell
git add README.md Deploy/README.txt docs/superpowers/checklists/phase-2-and-3a-smoke.md CLAUDE.md
git commit -m "docs: Quick Tools, clean up images to PDF" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
```

---

### Task 16: Review, release v2.3.0, handoff

- [ ] **Step 1: Full verification**

Run: Build; Test (no filter). Expected: all pass. Record the count.

- [ ] **Step 2: Manual check in a running Outlook (dev box)**

The ribbon icons, the review window layout and the real-photo results can't be unit-tested. Ask the user to run the Quick Tools smoke (Task 15 Step 3) after installing, or run it on the dev box if the user asks; do not claim it was checked otherwise.

- [ ] **Step 3: Reviews**

Run the `code-review` skill at `high` on the branch, and the Codex adversarial review from Git Bash:

```bash
node "/c/Users/MDASR/.claude/plugins/cache/openai-codex/codex/1.0.6/scripts/codex-companion.mjs" adversarial-review --wait --scope branch --base master "OutlookAI v2.3.0 Quick Tools: email images to a cleaned PDF (VSTO, C# 7.3, .NET 4.7.2). Look for: memory use under 32-bit Outlook with many large photos, Outlook COM calls off the UI thread, anything that could alter document content, a page silently dropped or mirrored, temp files left behind or leaked to the trace log, ribbon XML Office would reject, and PDF structure errors."
``` Fix confirmed findings test-first (each fix with a test that fails without it), re-run, repeat Codex until it approves.

- [ ] **Step 4: Merge and tag**

```powershell
git checkout master
git merge --no-ff feat/quick-tools-images-to-pdf -m "Merge: Quick Tools, clean up images to PDF (v2.3.0)" -m "Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01GTc94TjMbrRMqHy7Rp8zwu"
git push origin master
git tag -a v2.3.0 -m "OutlookAI v2.3.0 - Quick Tools: clean up images to PDF"
git push origin v2.3.0
```

(`Remove-Item Env:GITHUB_TOKEN` first if `gh`/`git push` uses the wrong account.)

- [ ] **Step 5: Build the bundle and check it**

```powershell
.\Deploy\Make-ReleaseZip.ps1 -Tag v2.3.0 -OutDir out
```

Check the zip: `version.json` tag `v2.3.0` and the merge commit; `OutlookAI.dll` contains the UTF-16 strings `Clean up images` and `ControlToolboxOutlook` (embedded ribbon XML is UTF-8: search both encodings); `README.txt` contains `QUICK TOOLS`; the `.sha256` matches.

- [ ] **Step 6: Release**

Write release notes (what it does, the review window, the privacy line, the limits: no PDF attachments, no images pasted into RTF emails; updating via Settings → Updates) to `$notes = "$env:TEMP\outlookai-release-notes-v2.3.0.md"`, then:

```powershell
gh release create v2.3.0 --title v2.3.0 --notes-file $notes out\OutlookAI-v2.3.0-RDS-Deploy.zip out\OutlookAI-v2.3.0-RDS-Deploy.zip.sha256
$run = gh run list --workflow release.yml --limit 1 --json databaseId -q '.[0].databaseId'
gh run watch $run --exit-status
gh api repos/kirklandsig/OutlookAI/releases/latest -q .tag_name   # v2.3.0
Remove-Item -Recurse -Force out
git branch -d feat/quick-tools-images-to-pdf
```

- [ ] **Step 7: Handoff and memory**

Update `handoff.md` (snapshot table: latest commit, tests, release, SHA256, pending smoke; § 5 a "Quick Tools (v2.3.0)" section; § 12 a new item; § 16 releases) and the project memory (`project_settings_ribbon.md` or a new `project_quick_tools.md` with an index line in `MEMORY.md`).
