# Phase 5: Excel & PDF Exports — Design

Branch: `feature/codex-oauth-migration`
Author session: 2026-05-18
Status: Proposed (pending user review)

## Why

Phase 4 shipped the Reports pane with templated chips that drive markdown reports back to the user in chat. The user's first real test prompt was:

> *"Can you generate an Excel report detailing all of the quotes I received from IT Creations for servers."*

The model produced a markdown table — useful, but it isn't an Excel file. The user clarified that PDF should be a feature on **any** chat output, while Excel is **specific** to tabular data. Without real file export, the Reports pane stops short of the workflow it implies: a report you can hand to a stakeholder.

## Goal

Ship two new model-callable tools and one UI affordance:

1. `outlook_export_excel` — model passes typed columns + rows; tool produces a styled `.xlsx`.
2. `outlook_export_pdf` — model passes markdown content; tool renders it through an isolated print template to a `.pdf`.
3. **Per-message "Save as PDF" button** on every assistant message — zero-cost one-click export of any output, calling `outlook_export_pdf` directly via the WebView bridge without a model round-trip.

Both tools save to `~\Documents\OutlookAI\Reports\` with auto-generated, timestamped filenames. Tool results surface as inline **file cards** with `Open` and `Show in folder` actions.

## Non-goals

- Settings UI for picking a custom Reports folder.
- Excel multi-sheet workbooks.
- Excel summary row (sum / avg / count footer).
- PDF page numbers, footer, table of contents, embedded images.
- Syntax highlighting in PDF code blocks.
- PDF metadata fields (Title / Author / Subject embedded in the file).
- CSV, `.docx`, or any other format.
- Email-the-file action (attach to a new draft).
- Drag-and-drop file card out of the chat window.

These are explicit follow-up candidates. Phase 5 ships the two formats, the per-message button, and the file-card UI.

## Architecture

Two single-purpose tools, one off-screen WebView2 renderer, one shared markdown module factored out of `chat.js`, and a small extension of `WebViewBridge` for the per-message button + file-action callbacks.

```
outlook_export_excel (tool)            outlook_export_pdf (tool)
        │                                       │
        ▼                                       ▼
LiveOutlookSurface.ExportExcel         LiveOutlookSurface.ExportPdf
        │                                       │
        ├─ ExportPathResolver                   ├─ ExportPathResolver
        ├─ ExportFilenameSanitizer              ├─ ExportFilenameSanitizer
        ├─ ExcelWorkbookBuilder (pure)          ├─ PrintTemplateRenderer (pure)
        │     └─ ClosedXML XLWorkbook           │     └─ string composition
        └─ XLWorkbook.SaveAs                    └─ PdfRenderer
                                                      ├─ off-screen WebView2 (cached)
                                                      ├─ NavigateToString(html)
                                                      └─ PrintToPdfAsync(path, settings)

both return { result_type:"file_saved", path, file_url, format, bytes, filename }
        │
        ▼
chat.js renderer detects result_type
        │
        ▼
appendFileCardToMessage(messageId, fileInfo)
        │
        ▼
FileCard with [Open] [Show in folder] → WebViewBridge.OpenFile / RevealInExplorer
                                          (path policy: must be under Reports dir)


per-message PDF button (every assistant message)
        │
        ▼
chat.js → chrome.webview.hostObjects.outlookai.exportPdf(hint, markdown)
        │
        ▼
WebViewBridge.ExportPdf → LiveOutlookSurface.ExportPdf (no model)
        │
        ▼
same FileCard rendered inline beneath the originating message
```

**Threading:** ClosedXML runs on any thread; the surface marshals via `OutlookThreadMarshaller` for consistency. WebView2 is STA and runs on the UI thread (same marshaller).

**Bridge security:** `WebViewBridge.OpenFile` and `RevealInExplorer` both pass the path through `IExportPathPolicy.RequireInsideReportsDir` — paths outside `~\Documents\OutlookAI\Reports\` are rejected. Prevents a malicious tool result or markdown injection from launching arbitrary executables.

## Components — new

### `Services/Export/ExportPathResolver.cs`

Resolves the canonical base directory `Path.Combine(Environment.GetFolderPath(MyDocuments), "OutlookAI", "Reports")`, creates it if missing, returns the absolute path. Caches the resolved path for the process lifetime.

### `Services/Export/ExportFilenameSanitizer.cs`

Pure helper. `Build(filenameHint, extension)` returns a safe filename:

- Sanitizes invalid Windows characters (`\ / : * ? " < > |`) → `-`.
- Trims trailing dots / spaces (Windows reserved quirk).
- Truncates the hint to 80 chars after sanitization.
- Falls back to `"OutlookAI-Report"` when hint is empty / whitespace-only.
- Appends `-yyyy-MM-dd-HHmm` timestamp before extension.
- If the resulting path already exists, appends `-2`, `-3`, ... before the extension until unique.

### `Services/Export/IExportPathPolicy.cs` + `ExportPathPolicy.cs`

Security gate for `OpenFile` and `RevealInExplorer` bridge methods. `RequireInsideReportsDir(path)` normalizes via `Path.GetFullPath`, compares case-insensitively against the canonical Reports dir prefix, throws `UnauthorizedExportPathException` if outside. Rejects UNC paths.

### `Services/Export/ExcelColumnType.cs` + type coercion

Enum: `Text`, `Date`, `DateTime`, `Number`, `Currency`, `Boolean`. Helper `Coerce(value, type)` returns the typed object for ClosedXML, or the raw string as a fallback when the input is unparseable (e.g. `"yesterday"` for a date column becomes a text cell, not an error).

### `Services/Export/ExcelWorkbookBuilder.cs`

Pure builder, no I/O. `Build(sheetName, columns, rows)` returns an `XLWorkbook` with:

- Bold header row, gray fill (`#F3F4F6`), frozen at row 1.
- AutoFilter applied over the header range.
- Columns auto-sized to content (`AdjustToContents()`).
- Currency cells formatted `$#,##0.00`.
- Date cells formatted `yyyy-mm-dd`.
- DateTime cells formatted `yyyy-mm-dd hh:mm`.
- Number cells default number format.
- Boolean cells `TRUE` / `FALSE`.

### `Services/Tools/ExportExcelArgs.cs` + `ExportExcelArgsParser.cs`

Tool args record + JSON parser. Validates:

- `columns` non-empty.
- `columns[].type` ∈ enum.
- `rows.length` ≤ **10 000**.
- Each row has exactly `columns.length` cells.
- `filename_hint` falls back to default if missing.

Throws `ToolArgValidationException` with the structured detail message documented in the design.

### `Services/Tools/OutlookExportExcelTool.cs`

Implements `IOutlookTool`. Parses args, calls `IOutlookSurface.ExportExcel(args, ct)`, returns `{ result_type: "file_saved", ... }` envelope or `{ error, detail }` on `ToolArgValidationException` / `ExportException`.

### `Services/Export/PrintTemplateRenderer.cs`

Pure renderer. `Render(title, subtitle, markdown, generatedAt)` returns full HTML string:

- Loads `print-template.html` from install dir once and caches it.
- HTML-encodes `title` and `subtitle`.
- JSON-encodes markdown and emits a `<script>window.__OUTLOOKAI_MD__ = "..."</script>` block.
- Strips inline images from markdown before injection: `![alt](url)` → `[image: alt]`.
- Replaces template tokens: `__TITLE_TEXT__`, `__SUBTITLE_TEXT__`, `__GENERATED_AT__`.

No DOM, no I/O. Fully unit-testable.

### `Services/Export/PdfRenderer.cs`

Owns the off-screen WebView2 instance + invisible host `Form`. Lazy-initialized on first use; cached until `Dispose`.

```csharp
public async Task<long> RenderAsync(string html, string outputPath, CancellationToken ct);
```

- `SemaphoreSlim(1, 1)` gate — one PDF at a time.
- `EnsureInitializedAsync()` first call: create form + WebView2, `EnsureCoreWebView2Async`, map install dir as virtual host (so `<script src="markdown.js">` resolves).
- `NavigateToString(html)` + await `NavigationCompleted`.
- Poll `document.body.dataset.renderState === "ready"` via `ExecuteScriptAsync` (up to 5s) to ensure markdown render finished.
- `PrintToPdfAsync(outputPath, settings)` with: `ShouldPrintBackgrounds=true`, `ShouldPrintHeaderAndFooter=false`, `Orientation=Portrait`, `ScaleFactor=1.0`.
- Returns the produced PDF byte size.

Cancellation cancels during nav-wait and render-ready poll. `PrintToPdfAsync` itself is uncancellable (WebView2 limit), but typical duration <500ms.

### `Services/Tools/ExportPdfArgs.cs` + `ExportPdfArgsParser.cs`

Validates:

- `content_markdown` non-empty.
- `content_markdown.length` ≤ **250 000** chars.
- `title` clamped to 200 chars; `subtitle` clamped to 400.
- `filename_hint` falls back to title or `"OutlookAI-Report"`.

### `Services/Tools/OutlookExportPdfTool.cs`

Implements `IOutlookTool`. Same pattern as Excel tool. Returns `{ result_type: "file_saved", ... }` or error envelope.

### `WebUI/markdown.js`

The existing chat markdown renderer (ATX headers, GFM tables, blockquotes, `<br>`) extracted from `chat.js` into a standalone module exporting `renderMarkdown(source)`. Pure function. Loaded by both `chat.html` (existing chat) and the new `print-template.html` (PDF rendering) so the two surfaces stay byte-identical for any markdown the model produces.

### `WebUI/print-template.html`

Skeleton HTML for PDF rendering:

```html
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>__TITLE_TEXT__</title>
  <link rel="stylesheet" href="print-styles.css">
  <script src="markdown.js"></script>
</head>
<body>
  <header class="doc-header">
    <h1 class="doc-title">__TITLE_TEXT__</h1>
    <p class="doc-subtitle">__SUBTITLE_TEXT__</p>
    <p class="doc-meta">Generated by OutlookAI · __GENERATED_AT__</p>
  </header>
  <main id="content"></main>
  <script>
    const html = renderMarkdown(window.__OUTLOOKAI_MD__ || "");
    document.getElementById("content").innerHTML = html;
    document.body.dataset.renderState = "ready";
  </script>
</body>
</html>
```

### `WebUI/print-styles.css`

A4 page, 0.5" margins via `@page`. `thead { display: table-header-group }` so long table headers repeat across page breaks. `tr { page-break-inside: avoid }`. Calibri 11pt body, color palette matches the Office UI (`#2b579a` headings, `#1f3a5f` body headings, `#f6f8fa` code background). Print-friendly (no shadows, hover states, transitions).

### `WebUI` file-card markup + styles

Added to `chat.js` and `styles.css`. A `.file-card` element with icon, filename + size, two action buttons. Distinct icons per format (`xlsx` / `pdf`). Action buttons call `chrome.webview.hostObjects.outlookai.openFile(path)` / `revealInExplorer(path)`.

### Per-message PDF button

Small button (📄) absolutely positioned top-right of every assistant message, opacity 0 unless message is hovered or button is focused. Click → `handleExportPdf(messageId)` → bridge call → file card rendered inline in the message's `.msg-attachments` container.

## Components — modified

### `Services/Tools/IOutlookSurface.cs` + `LiveOutlookSurface.cs`

Two new methods:

```csharp
Task<FileSavedResult> ExportExcel(ExportExcelArgs args, CancellationToken ct);
Task<FileSavedResult> ExportPdf(ExportPdfArgs args, CancellationToken ct);
```

`FileSavedResult` is `{ string Path, string FileUrl, string Format, long Bytes, string Filename }`.

### `Services/Tools/ToolCatalogSchema.cs`

Two new schema entries with steering descriptions:

- `outlook_export_excel`: "Use when the user asks for a spreadsheet, Excel file, xlsx, or tabular export. Pass typed columns and rows. Best for vendor lists, message tables, sender breakdowns. Don't use for prose."
- `outlook_export_pdf`: "Use when the user asks for a PDF, printable report, or document export. Pass polished markdown (title, headings, tables). Best for shareable narrative reports."

Both descriptions teach the model when to pick Excel vs PDF. Tests pin the steering text.

### `ThisAddIn.cs`

Registers both tools in the tool host alongside existing ones. Provides `PdfRenderer` as a singleton so it's shared across panes and disposed on shutdown.

### `WebUI/chat.js`

- Extract markdown renderer to `markdown.js` (imported via `<script src="markdown.js">` in `chat.html`).
- Add `appendFileCardToMessage(messageId, fileInfo)`.
- Add tool-result handling for `result_type === "file_saved"` (routes to `appendFileCardToMessage` instead of inline JSON).
- Add per-message PDF button to `renderAssistantMessage`.
- Add `handleExportPdf(messageId)` flow with spinner / disable state.
- Add error-card rendering for export-error envelopes.

### `WebUI/styles.css`

- `.msg-action` button positioned absolute, hover-revealed.
- `.file-card` and `.file-card-*` element styles.
- `.error-card` styles for export errors with optional Retry button.

### `Services/Tools/WebViewBridge.cs` (or equivalent host-object class)

Three new methods, all `[ComVisible(true)]`:

- `ExportPdf(string filenameHint, string contentMarkdown)` → calls `LiveOutlookSurface.ExportPdf` with synthesized args + `CancellationToken.None`.
- `OpenFile(string path)` → `IExportPathPolicy.RequireInsideReportsDir`, then `Process.Start(new ProcessStartInfo(path) { UseShellExecute = true })`.
- `RevealInExplorer(string path)` → policy check, `Process.Start("explorer.exe", $"/select,\"{path}\"")`.

## Data flow

### Excel (model-initiated only)

1. User asks "Make an Excel of X."
2. Model dispatches `outlook_search_messages` and optionally `outlook_read_messages`.
3. Model dispatches `outlook_export_excel` with `{ filename_hint, sheet_name, columns, rows }` projected from the gathered data.
4. `OutlookExportExcelTool` parses args, calls `LiveOutlookSurface.ExportExcel`.
5. `ExportPathResolver` ensures `~\Documents\OutlookAI\Reports\` exists.
6. `ExportFilenameSanitizer.Build(filename_hint, ".xlsx")` produces the final path.
7. `ExcelWorkbookBuilder.Build(...)` returns an in-memory `XLWorkbook`.
8. `XLWorkbook.SaveAs(path)` writes the file.
9. Tool returns `{ result_type: "file_saved", path, file_url, format: "xlsx", bytes, filename }`.
10. `chat.js` renders a FileCard with `Open` and `Show in folder` actions.
11. Model continues with a short prose summary above the card.

### PDF — model path

1. User asks for a PDF (or "save that as a polished report").
2. Model optionally composes fresh export-tuned markdown (option C from brainstorm).
3. Model dispatches `outlook_export_pdf` with `{ filename_hint, content_markdown, title?, subtitle? }`.
4. `OutlookExportPdfTool` parses args, calls `LiveOutlookSurface.ExportPdf`.
5. `PrintTemplateRenderer.Render(...)` composes the full HTML.
6. `PdfRenderer.RenderAsync(html, outputPath, ct)`:
   - Lazy-inits off-screen WebView2 if first call.
   - `NavigateToString(html)`; awaits navigation + render-ready.
   - `PrintToPdfAsync(outputPath, settings)`.
7. Tool returns the same file-saved envelope.
8. Same FileCard render as Excel.

### PDF — per-message button

1. User clicks `📄` button on an assistant message.
2. `chat.js` reads the message's raw markdown from `conversationStore`.
3. Derives `filename_hint` from the first heading or `"OutlookAI Report"`.
4. Calls `chrome.webview.hostObjects.outlookai.exportPdf(hint, markdown)`.
5. `WebViewBridge.ExportPdf` synthesizes `ExportPdfArgs` and calls the surface directly.
6. Same renderer pipeline as the model path.
7. FileCard rendered **inline within the originating message** (not as a new chat turn).

## Error handling

### Validation errors (pre-execution)

| Code | Trigger |
|---|---|
| `invalid_args` | Required field missing / wrong type / unknown column type |
| `too_many_rows` | Excel rows > 10 000 |
| `content_too_large` | PDF markdown > 250 000 chars |
| `row_shape_mismatch` | Excel row has wrong cell count |

### Runtime errors

| Code | Trigger |
|---|---|
| `path_unavailable` | Reports dir cannot be created / not writable |
| `file_locked` | Output file is open in another process (retried once with `-2` suffix before bubbling) |
| `disk_full` | `IOException` during save |
| `excel_build_failed` | ClosedXML throws |
| `pdf_render_failed` | WebView2 navigation `IsSuccess=false` |
| `pdf_print_failed` | `PrintToPdfAsync` throws |
| `pdf_render_timeout` | Render-ready signal didn't fire within 5s |
| `webview2_missing` | WebView2 runtime not installed (user-facing: link to Evergreen installer) |
| `cancelled` | Tool was cancelled mid-execution |

All errors are caught and translated into a structured `{ error: <code>, detail: <text> }` envelope. The chat renderer maps each code to a friendly error card; retryable codes (`file_locked`, `webview2_missing`, `path_timeout`) include a `Retry` button.

### Concurrency

- `PdfRenderer` semaphore (1, 1) — one PDF at a time.
- Excel has no semaphore (workbooks are independent; collision suffix handles same-filename races).
- Per-message PDF button is disabled while in-flight (cosmetic + prevents double-submit).

### Cancellation

- Tool path: same token plumbing as search tools (Stop button or shutdown).
- Bridge path: `CancellationToken.None`; one-click is uncancellable (≤500ms typical).

## Edge cases

| Case | Behavior |
|---|---|
| `filename_hint` empty / whitespace | Falls back to `"OutlookAI-Report-{format}"` |
| `filename_hint` 200+ chars | Truncated to 80 chars |
| `filename_hint` ending in dot/space | Trimmed |
| Same-minute filename collision | Suffix `-2`, `-3`, ... before extension |
| Reports dir is a file (sabotage) | `path_unavailable` error |
| Reports dir is a UNC path that's slow | 5s timeout → `path_timeout` |
| Markdown with `<script>` tags | `markdown.js` escapes HTML in non-code content |
| Markdown with inline images | Stripped to `[image: alt]` in v1 |
| Excel row with `null` / `undefined` | Empty cell |
| Excel currency string `"₹1,200"` | Falls back to text cell |
| Excel date string `"yesterday"` | Falls back to text cell |
| WebView2 crashes mid-export | `ProcessFailed` → recreate on next call; current call fails |
| Double-click PDF button | Button disabled mid-flight; second click no-ops |
| User deletes file while card visible | Open/Show-in-folder show toast "File no longer exists" |
| Bridge `hostObjects.outlookai` unavailable | Button shows toast "Export bridge not available — restart Outlook" |

## Tracing

Both tools emit structured events to `%LOCALAPPDATA%\OutlookAI\trace.log`:

```
export.excel.start    hint="..." rows=N cols=N
export.excel.built    workbook_bytes=N elapsed_ms=N
export.excel.saved    path="..." elapsed_ms=N
export.pdf.start      hint="..." md_chars=N
export.pdf.template_rendered  html_chars=N elapsed_ms=N
export.pdf.nav_complete       elapsed_ms=N
export.pdf.printed    path="..." pdf_bytes=N total_elapsed_ms=N
export.bridge.open    path="..."
export.bridge.reveal  path="..."
export.<*>.failed     err="<code>" retried=<bool>
```

## Testing strategy

| Layer | Coverage | Where |
|---|---|---|
| Pure helpers | Filename sanitization, path resolution, path policy, type coercion, template composition, image stripping, collision suffix | xUnit |
| Workbook composition | `ExcelWorkbookBuilder` produces correct cell types, formatting, autofilter, frozen header | xUnit against in-memory `XLWorkbook` |
| Arg parsers | Accept valid args, reject invalid with structured errors | xUnit |
| Tool wrappers | Translate args → surface call → envelope; cancellation propagates; errors translate | xUnit + `FakeOutlookSurface` |
| Path policy | Accepts canonical paths, rejects traversal / UNC / outside-dir | xUnit (security-critical) |
| Schema descriptions | Tools teach the model when to use which | xUnit on `ToolCatalogSchema` |
| Surface integration (Excel) | `LiveOutlookSurface.ExportExcel` writes a real `.xlsx` | manual smoke |
| Surface integration (PDF) | `LiveOutlookSurface.ExportPdf` produces a real `.pdf` | manual smoke |
| WebView bridge | Per-message button click produces a file card; Open/Show-in-folder work | manual smoke |

WebView2 cannot be unit-tested from `net472` xUnit (STA + COM + native), so `PdfRenderer` is integration-only. `PrintTemplateRenderer` (pure C# composition) is fully unit-tested.

Target test count after Phase 5: **~413 passing** (317 baseline + ~96 new).

## Implementation order

One commit per step. TDD where the gate is a unit-test count increase.

| # | Step | Test delta | Commit message |
|---|------|-----------:|----------------|
| 1 | Add ClosedXML 0.102.x NuGet dep | 0 | `chore(deps): add ClosedXML for Excel export` |
| 2 | TDD `ExportFilenameSanitizer` | +5 | `feat(export): add filename sanitizer with collision handling` |
| 3 | TDD `ExportPathResolver` + `ExportPathPolicy` | +14 | `feat(export): add reports directory resolver and path policy` |
| 4 | TDD `ExcelColumnType` + coercion helpers | +12 | `feat(export): add Excel column type coercion` |
| 5 | TDD `ExcelWorkbookBuilder` | +10 | `feat(export): add ExcelWorkbookBuilder with frozen header and autofilter` |
| 6 | TDD `ExportExcelArgs` + parser | +12 | `feat(export): add Excel args parser with strict validation` |
| 7 | Implement `LiveOutlookSurface.ExportExcel` | smoke | `feat(export): wire LiveOutlookSurface.ExportExcel` |
| 8 | TDD `OutlookExportExcelTool` | +8 | `feat(tools): add outlook_export_excel tool` |
| 9 | Add Excel schema + steering | +2 | `feat(schema): teach model when to use Excel export` |
| 10 | Register Excel tool in `ThisAddIn` | manual | `feat(host): register Excel export tool` |
| 11 | Factor `markdown.js` out of `chat.js` | smoke | `refactor(webui): extract markdown renderer to shared module` |
| 12 | TDD `PrintTemplateRenderer` | +12 | `feat(export): add PDF print template renderer` |
| 13 | Add `print-template.html` + `print-styles.css` content files | smoke | `feat(export): add PDF print template and styles` |
| 14 | Implement `PdfRenderer` (off-screen WebView2) | smoke | `feat(export): add WebView2-backed PdfRenderer` |
| 15 | TDD `ExportPdfArgs` + parser | +6 | `feat(export): add PDF args parser` |
| 16 | Implement `LiveOutlookSurface.ExportPdf` | smoke | `feat(export): wire LiveOutlookSurface.ExportPdf` |
| 17 | TDD `OutlookExportPdfTool` | +6 | `feat(tools): add outlook_export_pdf tool` |
| 18 | Add PDF schema + steering | +2 | `feat(schema): teach model when to use PDF export` |
| 19 | Register PDF tool in `ThisAddIn` | manual | `feat(host): register PDF export tool` |
| 20 | Extend `WebViewBridge` with `ExportPdf` + `OpenFile` + `RevealInExplorer` (path policy enforced) | +4 | `feat(bridge): add export PDF and file-action methods` |
| 21 | Add `result_type:"file_saved"` envelope handling + `appendFileCardToMessage` in `chat.js` | manual visual | `feat(webui): render file-saved tool results as file card` |
| 22 | Add per-message PDF button + hover styles + `handleExportPdf` flow | manual visual | `feat(webui): add per-message Save as PDF button` |
| 23 | Add error-card rendering for `{error}` envelopes with optional Retry action | manual visual | `feat(webui): render export errors as actionable error card` |
| 24 | Verify `conversationStore` retains raw markdown; add regression test if it doesn't | +1 | `fix(store): preserve raw markdown for assistant messages` (skip if already correct) |
| 25 | Update plan/spec docs as DONE; full test run | total ≈ 413 | `docs(superpowers): mark Phase 5 exports complete` |
| 26 | Publish Release + install elevated | hash match | (deployment, no commit) |
| 27 | End-to-end smoke (Excel + PDF + button) | trace + visual | (no commit) |
| 28 | Push branch | — | `git push origin feature/codex-oauth-migration` |

**Approximate commit count: 22-24.**

## Smoke acceptance criteria

1. Open Inbox Copilot → ask *"Make an Excel of my unread messages from today"* → file card appears within ~3s → `Open` launches Excel with bold + frozen header + autofilter + correct column types.
2. Ask *"Save that as a PDF"* → file card appears → `Open` shows a clean A4 PDF with header bar, our title, table preserved (header repeats across page breaks), no chat chrome.
3. Open any assistant chat output → hover → `📄` button appears → click → spinner → file card appears below the message → `Open` works.
4. Path-traversal test: synthesize a malicious `openFile` call with `..\..\Windows\System32\cmd.exe` → bridge rejects with `UnauthorizedExportPathException`.
5. Trace log shows clean `export.*` events; no exceptions.

## Risks

| Risk | Mitigation |
|---|---|
| WebView2 off-screen instance behaves differently than the visible one used for chat | Fallback: render the print template inside the existing visible WebView2 (hidden iframe) and `PrintToPdfAsync` that. Hackier but proven. |
| ClosedXML version conflict with existing deps | Pin to 0.102.x (current stable, no Microsoft.IO.Recyclable conflict on net472) |
| `PrintToPdfAsync` produces different output across WebView2 runtime versions | Standardize @page CSS controls all sizing; runtime-version-driven differences should be cosmetic |
| Large Excel rows (10k) hit ClosedXML memory ceiling | 10k cap chosen with 50KB per row headroom → well under default 2GB process limit |
| File card layout breaks in narrow task pane | Max-width: 480px on card; min content width ~280px tested visually |
| Per-message button visual clutter | Hover-only opacity 0 → 1; standard convention; can be moved to message footer if user finds it noisy |

## Out of scope (explicit Phase 6+ candidates)

- Settings UI for default Reports folder
- Excel multi-sheet
- Excel summary rows (sum/avg/count)
- PDF page numbers, footer, ToC
- PDF embedded images
- PDF metadata (Title/Author file properties)
- Code block syntax highlighting in PDF
- CSV format
- Word (.docx) format
- Email-the-file action
- Drag-and-drop file card to Explorer
- Per-pane export setting (enable/disable on Reports vs Inbox Copilot)
