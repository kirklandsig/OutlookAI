# Quick Tools: Clean up images → PDF (design)

**Date:** 2026-10-13 · **Target release:** v2.3.0 · **Status:** design approved in conversation, spec under review

## 1. Goal

Clients send scans and documents as images inside emails (pasted into the
body or attached), usually phone photos: dim, tinted, tilted, shot at an
angle, with the desk around the page. Staff need them legible and in one PDF.

One ribbon command takes the selected (or open) email's images, has the AI
work out what each one is and where the page is, cleans each page up with
deterministic image processing, lets the user review the result, and saves a
single PDF.

**Success:** for a typical email of 1–10 phone photos of paper, the user gets a
PDF of flat, upright, high-contrast pages with logos and signature images left
out, from two clicks (Clean up images → PDF, then Save PDF…), and can fix any
page the AI got wrong without leaving the window.

## 2. Decisions (from the brainstorming conversation)

| Question | Decision |
|---|---|
| Review before saving? | Yes: a review window (drop, reorder, rotate, Original vs Cleaned, color). |
| Which images? | Embedded (pasted in the body) **and** attached image files. Not pages of PDF attachments. |
| Where is the button? | Home tab (selected email) and the opened-email window. Not compose. |
| Output | Save As dialog starting in the Reports folder, then the PDF opens. |
| Image type | Mostly phone photos → perspective correction from 4 AI-found corners. |
| Cleanup approach | **A:** AI decides geometry and judgment (page or not, corners, color vs grayscale, strength); the add-in applies fixed, deterministic image processing. The AI never produces pixels. |
| Who can use it | Every user (no admin password): it reads mail and writes a file the user picks. |
| Scope | One email per run. |

## 3. Feasibility probe (2026-10-13, throwaway scripts, synthetic images)

Against `chatgpt.com/backend-api/codex/responses` with a ChatGPT-subscription
token, `input_image` as a base64 JPEG data URL:

- gpt-5.6-sol, gpt-6-astra and gpt-5.5 all accepted images (HTTP 200, 4–7 s).
- Tilted page (4°): all three returned rotation −4° and the page box within
  0.3% of the truth, and read the first line correctly.
- Page photographed at an angle (trapezoid): all three returned the 4 corners
  within ~0.2% of the truth, in reading order.
- `text.format` with a strict `json_schema` works on this backend; without it
  gpt-5.5 returned a different JSON shape.

Real photos (shadows, clutter, curled pages, fingers) will be harder than the
synthetic ones; the review window and the plan validation (§ 6.3) cover that.

## 4. User experience

### 4.1 Ribbon

- **Home tab (`TabMail`), AI Assistant group:** a large **Quick Tools** menu
  (`imageMso="ControlToolboxOutlook"`) after Settings, containing one button,
  **Clean up images → PDF** (`imageMso="PictureCrop"`). The menu leaves room
  for later tools.
- **Opened email (`TabReadMessage`):** a new AI Assistant group (appended to
  the tab, no `insertAfterMso`) with the same Quick Tools menu and a Settings
  button. `Ribbon.GetCustomUI` also serves `Microsoft.Outlook.Mail.Read`. AI
  Assistant itself is not added there (its pane is the compose pane).
- All control ids are unique across the XML (Office drops the whole ribbon
  otherwise); `RibbonTests` cover the new groups.

### 4.2 Flow

1. The user clicks **Clean up images → PDF**.
   - Inbox: needs exactly one selected item; read window: its current item.
   - Not a mail item, or zero/several selected → "Select one email first." (or
     "This works on email messages.").
   - Not signed in → the usual `NotSignedInException` message (Settings on the
     ribbon), before any work.
2. The images are collected (§ 6.2). None usable → "No images found in this
   email." (with "N small images skipped" when that is why).
3. A progress window ("Analyzing image 2 of 5…", Cancel) stays up while each
   image is planned by the AI and cleaned. Cancel stops everything and shows
   nothing further.
4. The review window opens (§ 4.3).
5. **Save PDF…** → Save As dialog → the PDF is written and opened.

### 4.3 Review window (`ImagesToPdfReviewForm`, WinForms)

```
┌ Clean up images → PDF — "Lease docs" from Jane Example ─────────────┐
│ ┌────────────────┐ ┌───────────────────────────────────────────┐  │
│ │ [x] 1  ▣ page   │ │                                           │  │
│ │ [x] 2  ▣ page   │ │      large preview of the selected page   │  │
│ │ [ ] 3  ▣ logo?  │ │                                           │  │
│ │ [x] 4  ▣ page   │ │                                           │  │
│ └────────────────┘ └───────────────────────────────────────────┘  │
│  [▲][▼]  [⟲][⟳]   (•) Cleaned  ( ) Original   Color: [Auto ▼]      │
│                                      [Cancel]   [Save PDF…]       │
└───────────────────────────────────────────────────────────────────┘
```

- **Page list:** thumbnail, include checkbox, position, and a label:
  - "page": an AI-judged document (a page, card, ID, form, receipt or screen),
    included by default.
  - "not a document: <reason>" (logo, signature image, decoration, a photo
    with no document in it): excluded by default but listed, so nothing
    disappears silently.
  - "AI couldn't analyze this one": basic cleanup only (§ 6.5), included.
  - "can't read this file": listed, can't be included.
- **Per page:** move up/down, rotate 90° either way, **Cleaned** / **Original**
  (the untouched image, upright per EXIF), **Color: Auto / Color / Grayscale**
  (Auto = the AI's choice). Changing color re-runs the local cleanup; no new AI
  call.
- **Memory:** the form never holds full-size images. Each page's original
  and cleaned versions live as files in the run folder; the form keeps their
  paths, a thumbnail (160 px) and, for the selected page, a preview (at most
  1200 px). A color change re-cleans just that page from its file.
- **Save PDF…** is disabled while no page is included.
- Buttons use explicit `ForeColor`/`BackColor` and
  `UseVisualStyleBackColor = false` (Server 2025).

### 4.4 Saving

- `SaveFileDialog`: initial folder `ExportPathResolver.ResolveBaseDir()` (the
  Reports folder, including the redirected-Documents fallback; created if
  missing). Default name from `ExportFilenameSanitizer.Build`, with the hint
  `"<sender display name> - <subject>"` and extension `.pdf`. Overwrite prompt on.
- The user may pick any folder, including network shares. A user-chosen
  Save As path is the user's own choice; `IExportPathPolicy` keeps governing
  paths chosen by the model, which this feature never does.
- The PDF is written to the run's temp folder first, then copied to the chosen
  path (overwrite), so a share never holds a half-written file. Then it opens
  with the shell's default app; if that fails, a message gives the path.

## 5. Architecture

```
Ribbon.OnImagesToPdfClick ─► ThisAddIn.RunImagesToPdf ─► ImagesToPdfCommand
   ├─ EmailImageCollector   (Outlook COM, UI thread)  → temp files + order
   ├─ PagePlanner           (AI, background, 3 at a time) → PagePlan per image
   ├─ PageCleaner           (pure, background)        → cleaned JPEG file per page
   ├─ ImagesToPdfReviewForm (UI)                      → final page list (paths + thumbnails)
   └─ ImagePdfWriter        (pure)                    → PDF from JPEG files → temp → Save As target
```

New files under `VSTO2/OutlookAI/Services/QuickTools/` (plus the form under
`QuickTools/`), each added as a `<Compile Include>`:

| Unit | Responsibility | Depends on |
|---|---|---|
| `ImagesToPdfCommand` | Find the email, run the steps with progress and cancel, own the run's temp folder, show review, save. | the units below, `OutlookThreadMarshaller`, `ExportPathResolver` |
| `EmailImageCollector` | Pick and order image attachments, save temp copies. COM part thin; ordering and filtering are pure functions. | Outlook interop |
| `ImageLoader` | Decode a file into upright bitmaps: EXIF orientation applied; multi-page TIFF → one bitmap per frame; GIF → first frame; WebP/HEIC through WIC when Windows has the codec. | System.Drawing, PresentationCore (WIC) |
| `PagePlanner` | Downscale, ask the AI, parse and validate the plan. | `CodexChatService` (new structured call), `Config` |
| `PagePlan` | The validated plan for one image (kind, reason, corners, color, strength, analyzed flag). | — |
| `PageCleaner` | Perspective warp + look pipeline + sharpening, on pixel arrays; writes the result as a JPEG file. | System.Drawing (LockBits, JPEG encode) |
| `Homography` | Solve the 3×3 projective map from 4 point pairs; map points. | — |
| `ImagePdfWriter` | Write a PDF with one JPEG file per page, embedding the file's bytes. | — |
| `ImagesToPdfReviewForm` | The review window (paths, thumbnails, one preview). | the above |

No new NuGet packages. `PresentationCore`/`WindowsBase` are .NET Framework
assemblies (reference only), used for WIC decoding.

**Threading:** the ribbon callback runs on the UI thread. Attachment access
(`Attachments`, `SaveAsFile`, `HTMLBody`) stays on the UI thread. Decoding,
AI calls and cleanup run on the thread pool; results come back via
`async`/`await` on the UI synchronization context. One run at a time: a second
click while a run is open brings that window to the front.

**Memory:** 32-bit Outlook still exists in the fleet (2 GB address space), so
at most one full-size image is decoded per worker at a time and every bitmap
is disposed as soon as its file is written. AI requests use the downscaled
copies only. Cleaning runs at most 2 pages in parallel.

## 6. Details

### 6.1 Finding the email

- Inbox: `Explorer.Selection.Count == 1` and `Selection[1] is MailItem`.
- Read window: `Inspector.CurrentItem is MailItem` (the ribbon control's
  `Context` gives the Explorer or Inspector).

### 6.2 Collecting images

- Candidates: `olByValue` attachments whose extension is one of `.jpg .jpeg
  .png .gif .bmp .tif .tiff .webp .heic .heif`, or whose `PR_ATTACH_MIME_TAG`
  (`http://schemas.microsoft.com/mapi/proptag/0x370E001F`, read through
  `Attachment.PropertyAccessor`) starts with `image/`. Embedded items and
  OLE attachments are ignored.
- Order: inline images first, in the order their `PR_ATTACH_CONTENT_ID`
  (`…/proptag/0x3712001F`) first appears as `cid:` in `HTMLBody`. Then the
  remaining image attachments in attachment order. A missing property is
  treated as absent, never as an error.
- Saved with `SaveAsFile` to `%LOCALAPPDATA%\OutlookAI\QuickTools\<run id>\`,
  with sanitized unique names.
- After decoding: images whose short side is under 200 px are skipped (icons,
  spacers, most signature logos) and counted.
- Rich Text (RTF) emails carry pasted images as OLE attachments, which this
  version ignores. When such an email yields no images, the message says so:
  "No images found. Images pasted into Rich Text emails aren't supported yet." At most 30 pages per run; extra
  ones are listed as "not processed (limit 30)".
- The run folder is deleted when the run ends (saved, cancelled or failed).
  Run folders older than a day are purged at the start of the next run.

### 6.3 Planning (AI)

- **Input to the AI:** the upright image, downscaled to at most 1600 px on the
  long edge, JPEG quality 85, as a data URL.
- **Request:** a new `CodexChatService` method,
  `CompleteStructuredAsync(instructions, IReadOnlyList<JObject> contentParts, string schemaName, JObject schema, string effort, CancellationToken)`,
  using the same endpoint, headers, token refresh, SSE parsing and error
  handling as the existing calls. Body: `model` =
  `Config.ModelCatalog.ResolveEffectiveModel(Config.Model, now)`;
  `reasoning.effort` = `ResolveWireEffort(model, "Low")` (omitted when the
  model doesn't take it); `input` = one user message with an `input_text` and
  an `input_image`; `text.format` = strict `json_schema`; `store: false`;
  `stream: true`. Returns the concatenated output text.
- **Instructions (summary):** you plan cleanup of photographed or scanned
  documents for a program that applies your answer; you never alter content;
  judge whether the image is a document page; give the page's 4 corners in the
  order the text reads (so a sideways or upside-down photo comes out upright);
  choose grayscale unless color carries meaning (stamps, highlights, colored
  ink, photos within the page); choose cleanup strength from how dim or
  uneven the image is. A *document* is a photo or scan of a page, card, ID or
  passport, form, receipt or screen, even when it shows a person;
  *not_document* is a logo, signature image, decoration, or a photo with no
  document in it.
- **Schema** (`page_plan`, strict, all fields required):

  ```json
  {
    "kind": "document_page | not_document",
    "reason": "string, short (e.g. 'company logo', 'signature image')",
    "corners": [[x, y], [x, y], [x, y], [x, y]],
    "color": "grayscale | color",
    "strength": "light | normal | strong"
  }
  ```

  `corners` are fractions of the image width and height: top-left,
  top-right, bottom-right, bottom-left as the text reads. For a
  `not_document`, the full image is expected.
- **Validation** (`PagePlan.FromJson`): each coordinate within [−0.02, 1.02]
  (then clamped to [0, 1]); the quadrilateral is convex and not
  self-intersecting; its area is at least 5% of the image. A plan that fails
  validation keeps its kind and color but uses the full image as corners, and
  is marked so the review shows "AI couldn't analyze this one".
- **Concurrency and limits:** 3 requests at a time (`SemaphoreSlim`); 90 s per
  image; one retry after a 429 (rate limited), waiting `Retry-After` when
  given (capped at 20 s) or 5 s otherwise; the run's `CancellationToken`
  cancels all of them.

### 6.4 Cleaning (`PageCleaner`, deterministic)

Input: the upright original (full resolution, loaded from its file) and the
plan (with the review's color override). Output: a JPEG file (quality 85) in
the run folder; the bitmaps are disposed before returning.

1. **Perspective warp:** map the 4 corners (in source pixels) to an output
   rectangle. Height is the longer of the left and right edges; width is the height times
   the page's aspect ratio, estimated from the perspective (Zhang & He,
   assuming the optical center is the image center) and falling back to the
   longer of the top and bottom edges when the photo shows no usable
   perspective or the estimate is implausible (focal length outside 0.25–4×
   the image size, or a ratio outside 0.2–5); scaled down (never up) to fit within
   2550 × 3300 px (or 3300 × 2550 when wider than tall). Inverse mapping
   through `Homography` with bilinear sampling.
2. **Look**, by strength:

   | Strength | Light evening | Black point | White point | Gamma | Sharpen |
   |---|---|---|---|---|---|
   | light | off | 0.5th percentile | 99.5th percentile | 1.00 | 0.3 |
   | normal | on | 1st percentile | 98th percentile | 0.95 | 0.5 |
   | strong | on | 2nd percentile | 95th percentile | 0.85 | 0.8 |

   - *Light evening:* divide each pixel by a heavily blurred copy of the
     luminance (computed at 1/8 scale, box blur with radius 1/20 of the short
     side, upsampled), which removes shadows and lighting gradients.
   - *Grayscale:* Rec. 601 luminance, then levels (black point → 0, white point
     → 255, gamma).
   - *Color:* white balance first (scale each channel so the paper color, the
     median of pixels above the 90th luminance percentile, becomes neutral), then
     the same levels on all channels using the luminance black and white points.
   - *Sharpen:* unsharp mask with a 3 × 3 blur, the amount above, threshold 2.
3. The review's quarter turns are applied last (`RotateFlip`) before the
   JPEG is written; "Original" pages are written the same way from the
   upright original, rotated but otherwise untouched.

The exact numbers are defaults in one place (`PageCleaner.Profiles`) and are
covered by tests; tuning them later doesn't change the design.

### 6.5 When something fails

- One image's AI call fails, times out or returns an invalid plan: that page
  is cleaned with the full image as corners, `normal` strength and grayscale,
  and labelled "AI couldn't analyze this one". The others continue.
- A file can't be decoded: listed as "can't read this file".
- `NotSignedInException` during the run: the run stops with its message.
- All AI calls fail (network): the review still opens with basic cleanup and
  a note at the top.
- Unexpected exceptions: trace-logged and shown in an error box, like the
  other ribbon commands; the run folder is still deleted.

### 6.6 PDF (`ImagePdfWriter`)

- Input: the final list of page JPEG files. Only one file's bytes are held in
  memory at a time while writing.
- PDF 1.4, written directly: catalog, pages, one page object per image, one
  image XObject per page (`/Filter /DCTDecode`, `/ColorSpace /DeviceRGB`,
  `/BitsPerComponent 8`; the JPEG file's bytes as-is, width and height read
  from its header), a content stream drawing it,
  an xref table with exact byte offsets, a trailer, and an Info dictionary
  (`/Title` = email subject; `/Producer` = `OutlookAI <version>`; non-ASCII
  text as UTF-16BE hex strings).
- Pages: Letter (612 × 792 pt), landscape (792 × 612) when the image is wider
  than tall; the image is fitted inside 18 pt margins, centered, aspect kept.
- Each page is encoded as a 24-bit JPEG at quality 85 (grayscale pages too,
  for simplicity).

### 6.7 Privacy and logging

- Image bytes (downscaled) go to ChatGPT through the same sign-in as the rest
  of OutlookAI, like email text already does. `Deploy/README.txt` and the
  README say so in one line.
- The trace log gets counts, sizes, durations, model and error types only:
  never image data, file names from the email, or the AI's answers.
- Temp copies are deleted as described in § 6.2.

## 7. Testing

**Unit tests** (xUnit, no Outlook, no network):

- `EmailImageCollector` ordering/filtering: inline-by-`cid:`-order then the
  rest; type detection by extension and MIME tag; small-image skip count; the
  30-page cap.
- `ImageLoader`: EXIF orientations 1–8 come out upright; a 3-frame TIFF gives
  3 bitmaps; an unreadable file gives "can't read".
- `PagePlan.FromJson`: valid plan; out-of-range, concave, self-intersecting
  and tiny quadrilaterals fall back to the full image and are flagged;
  malformed JSON is flagged.
- `Homography`: the 4 points map exactly; a known trapezoid maps to the unit
  square.
- `PageCleaner`: a synthetic trapezoid page with known corners comes out as a
  rectangle with a marker where expected; a dim, tinted page comes out with a
  near-white background and dark text (luminance thresholds); grayscale vs
  color output; the size cap; determinism (same input, same bytes).
- `ImagePdfWriter`: header, xref offsets point at their objects, page count,
  MediaBox per orientation, image width/height, the embedded JPEG decodes;
  Info title escaping.
- `CodexChatService.CompleteStructuredAsync` against `FakeHttpMessageHandler`:
  the body has `input_image` with a data URL, strict `text.format`, the
  resolved model and effort; the SSE output text is returned; errors map like
  the existing calls.
- `RibbonTests`: the Home and read-window Quick Tools menus, handlers, unique
  ids, `GetCustomUI` serving `Microsoft.Outlook.Mail.Read`.

**Opt-in live test** (skipped unless `OUTLOOKAI_LIVE_TESTS=1`): sends the two
synthetic probe images to the real backend through `PagePlanner` and checks
the plan against the known corners. It uses the real token file
(`OUTLOOKAI_LIVE_AUTH`, default `~/.codex/auth.json`), never a copy:
`CodexAuthService` saves refreshed tokens to the file it was given, and
refresh tokens are single-use, so refreshing a copy would sign Codex CLI out.

**Manual smoke** (added to the smoke checklist): a real email with pasted
phone photos and an attached JPEG, from the Inbox and from an opened email;
drop a logo, reorder, flip one page to Original, switch one to Color, Save PDF
to a network folder; the PDF opens and is legible.

## 8. Docs and release

- README: a Quick Tools section; Deploy/README.txt: Quick Tools and the
  privacy line; smoke checklist; CLAUDE.md lookup table row; handoff.md.
- Release v2.3.0 with the usual review flow (code review, Codex adversarial
  passes), installed through Settings → Updates.

## 9. Out of scope (possible follow-ups)

- Dragging the corners by hand in the review window.
- Several emails in one run.
- Pages from PDF attachments.
- Images pasted into Rich Text (RTF) emails (OLE attachments).
- Saving the PDF as an attachment to a reply.
- OCR or a text layer in the PDF.
- Quick Tools in compose windows.
