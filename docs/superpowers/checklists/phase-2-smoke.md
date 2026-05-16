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
