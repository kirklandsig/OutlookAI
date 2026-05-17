# Phase 2 Manual Smoke Checklist

Source of truth: `docs/superpowers/specs/2026-05-15-phase-2-tool-calling-and-compose-chat-design.md` § 6.

Run this every time the install bundle is refreshed. Each item should pass on
a clean Outlook session (close Outlook → run the installer → reopen Outlook).

## Pre-flight

- [ ] Build is clean and the test suite is green:

  ```powershell
  & "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
  dotnet restore VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
  & "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /t:Rebuild /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
  & "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
  ```

- [ ] Installer ran without errors. Step `[8/10]` either reports
  "WebView2 Runtime detected" or successfully installed the runtime from the
  vendored bootstrapper.
- [ ] Outlook starts; ribbon shows the `AI Assistant` button.

## Actions tab

1. [ ] Open Outlook → click a message in the Inbox → **Reply**. Open the AI Assistant pane.
2. [ ] Tabs visible: ✨ Actions, 💬 Chat, 🎭 Variants. Default selection: Actions.
3. [ ] Pane width ~340 px; no horizontal scroll on the Actions content.
4. [ ] Type a short draft in the compose window. Click **Proofread**.
   - [ ] Status shows `Processing...` then `Done! Review the result below.`
   - [ ] Result text appears in the Result panel.
   - [ ] Cancel button is hidden after completion.
5. [ ] Click **Draft Email** with instructions (e.g. "thank you for the meeting").
   - [ ] Tool-call mini-strip MAY surface zero or one `outlook_get_current_compose_state` line.
   - [ ] Result appears; `Insert` and `Replace` both work against the compose body.
6. [ ] Cancel mid-turn: start a long draft → click **Cancel** while streaming.
   - [ ] Status: `Cancelled.` or `Stopped. Partial result shown.`
   - [ ] No exception dialog.

## Chat tab

7. [ ] Switch to **Chat**. WebView2 panel renders with the context strip at the top
   (subject + recipients + thread topic).
8. [ ] If WebView2 is missing, a friendly fallback label shows the download URL.
9. [ ] Type "Summarize this thread" + Enter.
   - [ ] Tool cards appear: `outlook_get_current_compose_state` →
     `outlook_read_message` (one per recent message in the reply chain).
   - [ ] Cards show check (✓) on success, warning (⚠) on failure.
   - [ ] Assistant message streams in below the tool cards.
10. [ ] Type "Has [some recipient] replied to my proposal?".
    - [ ] Tool cards: `outlook_list_recent_threads_with` followed by
      `outlook_search_messages`.
11. [ ] Type "Mark all newsletter emails as read".
    - [ ] Tool cards: `outlook_search_messages` then repeated `outlook_mark_as_read`.
    - [ ] **Audit row** appears for each successful write.
    - [ ] Mailbox UnRead counts update accordingly.
12. [ ] Type a long question, click **Stop** mid-stream.
    - [ ] Partial assistant message has the `· stopped` italic suffix.
    - [ ] Composer is enabled again. Follow-up turn works normally.
13. [ ] Click **Clear**. Messages area empties. ConversationStore is reset.
14. [ ] Click **Copy**. Clipboard has the conversation in the
    `[role]\n<content>\n` / `[tool call]` / `[tool result]` shape.

## Variants tab

15. [ ] Switch to **Variants**. UI shows intent textbox, count `3`, blank
    reasoning, Generate / Regenerate-all / Cancel.
16. [ ] Type "Polite follow-up about Q4 report". Click **Generate**.
    - [ ] Status: `Generating...` → `Done. 3 variant(s) ready.`
    - [ ] 3 cards with distinct color-coded tone chips, char counts,
      first-3-lines preview, Insert / Replace / Regenerate per card.
17. [ ] Click one card's **Regenerate** button.
    - [ ] Only that card changes; the others stay put.
    - [ ] The replacement has the same tone label as the original.
18. [ ] Click another card's **Insert** → compose body gets the variant text
    prepended.
19. [ ] Click a third card's **Replace** → compose body becomes that variant.

## Settings + admin

20. [ ] Open Settings (gear icon on Actions tab) → enter admin password.
21. [ ] Change Model dropdown to `gpt-5.5-pro`. Reasoning dropdown re-filters
    (`None`, `Minimal`, `Low`, `Medium`, `High` all visible).
22. [ ] Switch Model to `gpt-4.1-mini`. Reasoning dropdown collapses to `None`
    only.
23. [ ] Uncheck `outlook_mark_as_read` in the write-tools checklist.
    Click **Save AI Settings**. `Saved.` indicator appears, auto-hides ~2.5 s.
24. [ ] Switch back to Chat tab → type "Mark this as read". The tool card
    should NOT include `outlook_mark_as_read` (model gets a tool error or
    works around it).
25. [ ] Re-open Settings → confirm Model + Reasoning + checklist persisted.

## Persistence + lifecycle

26. [ ] Close the compose window. Reopen Reply on the same message.
27. [ ] Chat tab: conversation history is **empty** (per spec, history is
    per-Inspector and clears on close).
28. [ ] Variants tab: cards are **empty**.
29. [ ] Settings (Model, Reasoning, EnabledWriteTools) **still persisted** via
    `%APPDATA%\OutlookAI\config.xml`.

## Voice (still wired from Phase 1)

30. [ ] On Actions tab, click the red mic next to the Draft prompt.
    - [ ] Recording indicator appears; speak for 2-3 seconds.
    - [ ] Click again. Status: `Transcribing...` → text appears in the prompt.

## Performance gates (Spec § 6)

- [ ] First chat-tab activation < 800 ms cold (WebView2 init).
- [ ] First streamed token visible < 2 s p50, < 4 s p95.
- [ ] `outlook_search_messages(max=25)` returns in < 1.5 s on a 50k+ message
  mailbox.
- [ ] No frame drops during sustained token streaming.
- [ ] No Outlook UI freeze for > 100 ms during a tool call.
- [ ] After 100 chat turns the add-in's memory delta is < 50 MB (GC settled).

## Cleanup

- [ ] Run `Deploy\Uninstall-OutlookAI.ps1` (when applicable). Verify
  `C:\Program Files\OutlookAI` removed, `%LOCALAPPDATA%\OutlookAI\WebUI` and
  `%LOCALAPPDATA%\OutlookAI\WebView2Data` cleaned up on next user logon.

---

If anything fails: file a follow-up bug task on the same branch and re-run
the smoke pass before tagging the Phase 2 release.

---

# Phase 3a Manual Smoke (Inbox Copilot)

Run on a fresh Outlook session **after** the Phase 2 smoke section above
passes.

## Pre-flight

- [ ] Re-publish + reinstall via the elevated installer one-liner.
- [ ] Verify the AI Assistant button is on both ribbons:
  - Open a compose window → the AI Assistant group on the compose
    ribbon is unchanged.
  - Close it. Look at the main Outlook Home tab (TabMail) → an
    AI Assistant group now appears far right, after Move.

## Pane lifecycle

1. [ ] Click AI Assistant from the Inbox view → a 340-px pane opens on
   the right. Single chat surface, no tabs.
2. [ ] Context strip line 1 shows `In: Inbox (<n> unread)`.
3. [ ] Selection line is hidden when no message is selected.
4. [ ] Click a message in the reading pane → context strip refreshes,
   shows `Selected: <subject> — <from>`. Chip row refreshes to add
   `Summarize this thread` and `Draft a reply`.
5. [ ] Ctrl-click a second message → chip row updates to multi-select
   variants `Summarize all selected` / `Triage selected`.
6. [ ] Switch folders in the navigation pane → context strip's folder
   line updates.
7. [ ] Open a second Explorer (File → Open & Export → New Window) →
   click AI Assistant → a second, independent pane opens. Conversations
   are NOT shared between the two.
8. [ ] Close the first Explorer → its pane disappears. The second
   pane + conversation are untouched.

## Functional checks

9. [ ] Click the `What needs my attention?` chip → prompt fills the
   textarea AND auto-sends. Streaming response appears. Tool cards
   reflect the actual search/list calls the model made.
10. [ ] Type "Find unread emails from <a known sender> with attachments
    from the last 7 days." → the model issues exactly ONE
    `outlook_search_messages` call with all four structured fields
    populated. Verify via `%LOCALAPPDATA%\OutlookAI\trace.log` looking
    for the `BuildRunTurnRequest` line and the function_call argument
    payload.
11. [ ] Select one message in the reading pane, click `Summarize this
    thread` → response mentions the actual subject line and at least
    one detail from the body. The model likely calls
    `outlook_get_current_selection` first (visible in tool cards), then
    `outlook_read_message` for other thread items.
12. [ ] Stop mid-stream → partial assistant message keeps the
    "stopped" badge. Composer re-enables.
13. [ ] Clear button empties the chat.
14. [ ] Copy button puts the conversation on the clipboard
    (paste into Notepad to verify).

## Settings / model awareness

15. [ ] Open Settings → change Model to `gpt-5.4` → save. Close the
    pane (X). Reopen via the ribbon button. The reasoning dropdown now
    includes `Minimal` (gpt-5.4 supports it; gpt-5.5 does not).
16. [ ] Change Model back to `gpt-5.5` and confirm the dropdown drops
    `Minimal` on the next pane open.

## Lifecycle / cleanup

17. [ ] Close all Explorers → the per-Explorer panes + conversations
    are gone.
18. [ ] Restart Outlook → no leftover state visible. (Phase 3a is
    deliberately non-persistent.)

If any step fails, capture the trace log + screenshot and file as a
follow-up. Do NOT merge to master until every step passes.
