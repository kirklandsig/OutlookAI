# README and Launch Positioning - Design

Branch: `feature/codex-oauth-migration`
Author session: 2026-05-19
Status: Proposed (pending user review)

## Why

OutlookAI has grown well past the Phase 2 launch the current `README.md`
describes. The branch about to merge to `master` ships:

- WebView2 chat in the compose pane.
- Inbox Copilot taskpane with selection-aware actions.
- Inbox Reports taskpane with templated chips and markdown reports.
- Excel and PDF export tools, per-message Save as PDF, file cards with
  Open / Show in folder, and a path-policy security gate.
- A 12-tool mailbox surface, admin-gated write tools, model and reasoning
  effort pickers, and OAuth via the user's ChatGPT subscription.

The current `README.md` still says Phase 2, 94/94 tests, 10 tools, and "not
merged to master." It also undersells the product as a generic compose helper
when the real story is a full Outlook AI integration that competes with the
paid market.

The goal of this work is to refresh the launch positioning before merging
`feature/codex-oauth-migration` to `master`, so the first impression of the
repository on `master` matches what is actually shipped.

## Goal

1. Rewrite `README.md` as a product landing page that frames OutlookAI as
   "the free, open-source Outlook AI add-in powered by your existing ChatGPT
   subscription."
2. Tell the full story: Chat, Inbox Copilot, Inbox Reports, Excel/PDF
   exports, voice, write-tool permissions, RDS/Server deployment.
3. Refresh `Deploy/README.txt` (and add a short `docs/Install.md` if useful)
   to cover three deployment shapes: single workstation, multi-user RDS /
   Terminal Server, and IT-managed image / silent install.
4. Keep all claims factual and dated to what is actually shipped on the
   branch about to merge.
5. Set explicit "still in development" expectations and link to the open PR.

## Non-goals

- New product features. This is a documentation/positioning pass only.
- A separate marketing site or landing page outside the repository.
- Changing the OAuth flow, the tool catalog, or any installer behavior.
- Removing or reworking existing phase plans/specs under
  `docs/superpowers/`. Those remain as historical record.
- Adding screenshots or video assets in this pass. Image work can come
  later in a follow-up.

These are explicit follow-up candidates.

## Audience and messaging

**Primary audience:** Outlook power users and small/medium businesses already
paying for ChatGPT, who currently overpay for paid Outlook AI add-ins such as
GPT for Outlook, Mailbutler, Lavender, Compose AI, OtterMail, Boomerang
Respondable, Mailmaestro, EmailTree, and SaneBox AI.

**Secondary audience:** IT admins deploying to RDS / Terminal Server, and
open-source developers and contributors.

**Core message:** You already pay for ChatGPT. OutlookAI lets you use it inside
Outlook for free, with real mailbox tools, real reports, and real exports - not
just "rewrite this email."

**Tone:** Confident, factual, no hype. Always disclose what is still in
development and what is not yet merged. Prices and feature claims about
competitors are stamped with "public list prices as of writing" and are not
guarantees.

## README structure

The new `README.md` is a top-to-bottom landing page. Section order:

1. **Hero.** Title and one-line description. Badges for license, build, tests
   (`546/546`), and branch status. Two-sentence pitch focused on "free + uses
   your own ChatGPT subscription + runs inside Outlook." A clearly-marked
   "still in development" banner that links to the open PR for the merge to
   `master`.
2. **TL;DR.** Four to six bullets covering: free, OAuth via ChatGPT,
   in-process (no proxy), tool calling on real Outlook data, Inbox Copilot +
   Reports + Exports, voice.
3. **What you get.** Feature deep-dive for the surfaces actually shipped on
   this branch:
   - Chat tab in the compose pane (WebView2, streaming, multi-round tool
     dispatch).
   - Inbox Copilot taskpane (selection-aware actions, multi-round chat).
   - Inbox Reports taskpane (templated chips, markdown reports, action
     items).
   - Excel and PDF exports, with file cards offering Open / Show in folder.
   - Per-message Save as PDF in chat and reports.
   - 15 model-callable Outlook tools: 11 read/utility tools always on, plus
     4 admin-gated safe-write tools.
   - Voice input via the OpenAI Realtime WebSocket.
   - Settings UI: model picker, reasoning effort, write-tool gates.
4. **How it compares.** Refreshed comparison table vs GPT for Outlook,
   Mailbutler, Lavender, OtterMail, Compose AI, Boomerang Respondable,
   Mailmaestro, EmailTree, and SaneBox AI. Columns include price, OSS, OAuth
   via personal ChatGPT, mailbox tool calling, in-process (no proxy), exports
   (Excel/PDF), reports, voice, write-permission gating, RDS-ready, and
   telemetry.
5. **Why OutlookAI.** BYO ChatGPT subscription with no per-seat fee. Runs
   entirely inside `Outlook.exe`, no proxy or SaaS middle layer. Real tools,
   real exports, real reports - not just compose helpers. Open source under
   MIT, auditable end-to-end.
6. **Install (single workstation).** A 60-second snippet that clones, builds
   Release, and runs `Install-OutlookAI.ps1 -SourcePath` elevated. Link to the
   deployment guide for everything else.
7. **Deployment guide link.** Pointer to `Deploy/README.txt` (and
   `docs/Install.md` if added) covering the three install shapes below.
8. **Architecture.** Brief description of the runtime layout: Outlook process
   to WebView2 chat UI, `CodexChatService` to the tool catalog, and export
   services for Excel/PDF rendering. Names key components such as
   `LiveOutlookSurface`, `OutlookToolHost`, `ExportBridge`,
   `PrintTemplateRenderer`/`PdfRenderer`, and `ExcelWorkbookBuilder`.
9. **Security model.** OAuth via ChatGPT, where tokens live (machine + per
   user), RDS shared-credential risk, rotation steps, path policy on file
   actions, no send/delete/move-to-deleted tools, no telemetry, and the two
   outbound endpoints in use.
10. **FAQ + competitor framing.** Honest, short answers to "Is there a free
    alternative to GPT for Outlook?", "Can I use my ChatGPT Plus subscription
    inside Outlook?", "How does this compare to Mailbutler / Lavender / etc.?",
    Outlook desktop only, OWA / Mac not in scope today.
11. **Status and roadmap.** Active development on
    `feature/codex-oauth-migration`, merging to `master` now. List of what is
    shipped: Chat, Inbox Copilot, Inbox Reports, Excel/PDF exports. Known
    gaps: multi-sheet Excel, page numbers in PDF, settings UI for custom
    Reports folder, Mac / OWA.
12. **Contributing.** Build and test instructions, branch model, spec/plan
    workflow under `docs/superpowers/`.
13. **License.** MIT.

## Deployment doc shape

`Deploy/README.txt` is rewritten around three install shapes, with the current
OAuth + RDS content folded in.

### Single workstation (developer or power user)

- Build from source or use the latest Release zip.
- Run `Install-OutlookAI.ps1 -SourcePath <staging>` elevated.
- Sign in once with the ChatGPT account that will be billed.

### Multi-user RDS / Terminal Server (primary deployment target)

- Same installer, run on the server image.
- Shared OAuth credential at `C:\ProgramData\OutlookAI\auth.json` with
  `Authenticated Users: Modify`.
- Explicit accepted-risk section that documents what a malicious local user
  could do with the shared credential.
- Rotation steps via the OutlookAI Settings UI (Sign Out / Sign In) and
  ChatGPT account-side session revocation.
- Per-user `%APPDATA%\OutlookAI\config.xml` handling for the v2 layout.

### IT-managed image / silent install

- Bake `Install-OutlookAI.ps1` into the image, calling it with explicit
  `-SourcePath` and noting that `#Requires -RunAsAdministrator` enforces
  elevation.
- Pre-stage the WebView2 Evergreen runtime using the vendored
  `Deploy/Fetch-WebView2Bootstrapper.ps1` so the installer does not need
  internet at install time.
- Post-install verification: compare SHA256 of installed
  `C:\Program Files\OutlookAI\OutlookAI.dll` against the staged published
  build.

Common sections for all three shapes: prerequisites, what the installer does,
verification, troubleshooting, rollback, uninstall.

A short `docs/Install.md` may be added that links from `README.md` to the same
three shapes if `Deploy/README.txt` grows past a comfortable size. The
canonical content lives in one place to avoid drift; the other location links
to it.

## Honesty constraints

The new docs must keep claims factual.

- Test count cited as `546/546`, matching the current `feature/codex-oauth-migration` branch verification.
- "Still in development" banner at the top of `README.md`, with a link to the
  open PR for the merge to `master`.
- Competitor comparison table is labeled "public list prices as of writing"
  and lists the actual feature gating each competitor advertises. No
  guarantees about competitor behavior.
- Voice section is labeled as Phase 1, present but not the main draw.
- Do not claim feature parity for anything not actually wired (for example,
  multi-sheet Excel, PDF page numbers, settings UI for a custom Reports
  folder).
- Security section explicitly calls out the RDS shared-credential risk and
  links to the rotation steps.
- Tool count is `15`: 11 read/utility tools always on
  (`outlook_get_current_compose_state`, `outlook_get_current_selection`,
  `outlook_list_folders`, `outlook_search_messages`, `outlook_read_message`,
  `outlook_read_messages`, `outlook_count_messages`,
  `outlook_aggregate_messages`, `outlook_list_recent_threads_with`,
  `outlook_export_excel`, `outlook_export_pdf`) plus 4 admin-gated safe-write
  tools (`outlook_create_draft`, `outlook_mark_as_read`,
  `outlook_flag_message`, `outlook_set_category`). The README lists every
  tool name so the count is verifiable.
- Reports surface is described as "templated chips that drive markdown
  reports back to chat," matching `InboxReportsController` behavior rather
  than over-claiming.

## Out of scope for this pass

- Adding any new product features.
- Screenshots, videos, or marketing imagery.
- A separate web landing page.
- Reworking the OAuth flow, the tool catalog, the installer, or any of the
  phase-history documents under `docs/superpowers/`.

## Acceptance criteria

- `README.md` follows the structure above and replaces the Phase 2-era
  framing on the branch about to merge.
- `Deploy/README.txt` covers single workstation, RDS / Terminal Server, and
  IT-managed image installs without losing any current OAuth or shared
  credential warnings.
- All counts, version numbers, and "what is shipped" claims match the
  feature branch state at merge time (`546/546` tests, 15 tools (11 always
  on + 4 admin-gated), exports shipped, Inbox Copilot + Inbox Reports
  shipped, multi-sheet Excel and PDF page numbers explicitly listed as not
  yet shipped).
- The merge banner at the top of `README.md` links to the open PR. After
  the merge to `master`, that banner is updated to a "v2 shipped" callout in
  a follow-up commit on `master`.
- No new product code is touched. Tests, build, and installed DLL hash
  remain green and unchanged from the current branch tip.
