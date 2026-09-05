# Web parity — `Passage.Web` vs `Passage.App.Linux`

The work queue for bringing the Blazor Server app up to the Linux port's
feature and visual set. **This file is the source of truth for remaining
work.** Update the row in the same session as the change.

Line ranges are the point of this document. Read the ranges named here with
`sed -n 'X,Yp' <file>`; do not read `MainWindowViewModel.cs` (2,797 lines) or
`MainWindow.axaml.cs` (1,348 lines) in full. Never read
`Passage/Passage.Web/wwwroot/lib/`.

Paths below are relative to the repository root. `VM` =
`Passage/Passage.App.Linux/ViewModels/MainWindowViewModel.cs`. `MW.cs` =
`Passage/Passage.App.Linux/Views/MainWindow.axaml.cs`. `MW.axaml` =
`Passage/Passage.App.Linux/Views/MainWindow.axaml`. `Editor.razor` =
`Passage/Passage.Web/Components/Pages/Editor.razor`.

Status values: `missing` · `partial` · `done` · `verified` (done and manually
checked against the Linux build).

---

## Tier 1 — wiring, not porting

The shared services already exist in `Passage.Core` and the web app simply
never calls them. Highest value per token in the project. Do these first.

### 1.1 Session state restore — `done`

- **Linux:** VM 2398–2483 (`LoadSessionState`, `SaveSessionNow`);
  MW.cs 699–773 (`OnOpened`), 774–795 (`OnClosing`)
- **Shared code already present:** `Passage/Passage.Core/Services/SessionStorage.cs`
- **Web:** `Editor.razor` 371–389 (`OnInitialized` / `OnAfterRenderAsync`)
- **Notes:** Restores last document, zoom and (per FUTURE_IMPROVEMENTS 12) not
  yet caret. The desktop app has one user; the web app does not — persist
  per-browser in `localStorage` via `passage.js`, **not** globally on the data
  volume, or two browsers will fight over it. Caret restore is a free addition
  here since `OnCaretMoved` already exists (`Editor.razor` 403–410).
- **Done:** Restores last document, caret line, editor font size and preview
  zoom. State lives under the `passage.session.v1` localStorage key, written by
  `scheduleSessionSave` / `loadSession` / `setSessionDocument` in `passage.js`.
  `SessionStorage.cs` is deliberately **not** used — it writes one file under
  the server's `LocalApplicationData`, shared by every client. The caret is
  tracked in JS, not pushed from Blazor, so a moving cursor costs no
  round-trips (CLAUDE.md trap 4). Razor pushes only document/font/zoom, via
  `PersistSessionAsync`. A stored document that no longer exists is pruned on
  load with a status message rather than an error. Rows 1.5 (Open Recent) and
  3.2 (line-type overrides) should extend the same key.

### 1.2 Crash recovery — `done`

- **Linux:** VM 2333–2397 (`StartRecoveryAutosave`, `StopRecoveryAutosave`,
  `SaveRecoverySnapshot`, `LoadRecoveryDocument`); VM 80 (`_recoveryTimer`);
  `Passage/Passage.App.Linux/Views/RecoveryPromptDialog.axaml` + `.axaml.cs`
  (whole files, 44 + 22 lines)
- **Shared code already present:** `Passage/Passage.Core/Services/RecoveryStorage.cs`
- **Web:** missing
- **Notes:** Distinct from the existing file autosave (`Editor.razor` 612–646),
  which writes the real file. Recovery snapshots are a separate crash-safety
  net. On the server, scope snapshots by browser session so one client's crash
  recovery is not offered to another.
- **Done:** Snapshots live in `localStorage` under `passage.recovery.v1`,
  written by `saveRecoverySnapshot` in `passage.js` on a 3s timer while dirty —
  the same interval as the Linux `_recoveryTimer`. **Deviation from the note
  above:** they are kept client-side, not on the server. A server-side snapshot
  needs a live Blazor circuit, which is exactly what a crash removes, so edits
  made in the seconds before a kill would be lost; localStorage also scopes
  snapshots to one browser with no session-id plumbing and no way to hand one
  client's draft to another (there is no auth — trap 5).
  `CheckForRecoveryAsync` (`Editor.razor`) offers the snapshot only when it is
  newer than the file's `LastModifiedUtc`; an untitled snapshot has no file to
  compare against and always counts as newer. Prompt wording and button order
  match `RecoveryPromptDialog.axaml` exactly. Clearing is explicit — a real
  save (`setDirty(false)`) or Discard. The timer must **never** clear: a
  snapshot can be sitting in the open prompt, and a tick that wiped it would
  destroy the work being offered back.

### 1.3 Light/dark theme toggle — `done`

- **Linux:** VM 655–665 (`SetDarkTheme`), 666–676 (`SetLightTheme`);
  `Passage/Passage.App.Linux/App.axaml.cs` 173–264 (`Brush`,
  `LoadThemeResources`, palette); MW.axaml 302–305 (menu)
- **Web:** dark only — `Passage/Passage.Web/wwwroot/css/app.css` 1–70
  (`:root` custom properties, `--theme-background: #0D0D0C`)
- **Notes:** The CSS already routes colour through custom properties, so this
  is a second `:root[data-theme="light"]` block plus a toggle that persists to
  `localStorage`. Take the light palette from `App.axaml.cs` `LoadThemeResources`
  so the two apps match exactly. Do **not** invent colours; see
  `docs/Linux Port UI Overhaul — Visual Consistency with Windows.md`.
- **Done:** `:root` keeps the dark palette as the default; `:root[data-theme="light"]`
  holds the light one. All 26 shared values are copied verbatim from
  `LoadThemeResources` and verified equal at runtime. Note the UI-overhaul doc's
  cream/paper table was **superseded** — the shipped code is the grey monochrome
  scheme, and that is what was used.
  Nine hardcoded colours in `app.css` were routed through new tokens:
  `--selection-background`, `--backdrop`, `--page-background`/`--page-foreground`
  and five `--shadow-*` values (light leans on shadows, dark on hairline borders,
  matching the two schemes' comments). `--control-pressed-foreground` now colours
  text on accent fills, which `--theme-background` only approximated in light.
  The preview page is deliberately theme-invariant — it simulates printed paper —
  but is tokenised so nothing is hardcoded. Tokens with no CSS consumer yet
  (`--editor-page-border`, `--card-accent`, `--hierarchy-indicator`,
  `--drag-over-background`) are defined for both themes ready for Tier 2/3.
  The theme is stored under `passage.theme.v1` and applied by an inline script in
  `App.razor` **before first paint**, so a stored light theme does not flash dark
  on every load. Toggle lives in the topbar. Not done: Avalonia's startup system
  detection (`prefers-color-scheme`) — the web app defaults to dark as before.

### 1.4 Keyboard shortcuts — `partial` (blocked on Tier 2 features)

- **Linux:** MW.cs 146–179 (`AddKeyboardShortcuts`); MW.axaml 250–320
  (`InputGesture` on every menu item)
- **Web:** `Passage/Passage.Web/wwwroot/js/passage.js` 66 — `Ctrl-S` only
- **Notes:** Bind in CodeMirror's keymap client-side, not in Razor. Ctrl+N/O/W,
  Ctrl+Z/Y, Ctrl+F, Ctrl+G, Ctrl+M, zoom. Some rows below depend on this one;
  do it before them. Avalonia rejects `Ctrl+Plus` (FUTURE_IMPROVEMENTS 10) —
  the browser has no such limitation, so web zoom gestures can be complete.
- **Done:** Bound in CodeMirror's `extraKeys` (`passage.js`), client-side.
  Added `Ctrl-=` / `Ctrl-+` / `Ctrl--` for editor zoom, calling
  `OnZoomShortcut`. Verified the browser's own page zoom is suppressed
  (`defaultPrevented === true`), so the doc's note holds — web zoom gestures
  work where Avalonia's `Ctrl+Plus` does not. Already working and left alone:
  `Ctrl-S` (bound previously) and `Ctrl-Z` / `Ctrl-Y` / `Ctrl-Shift-Z`, which
  come free from CodeMirror's default PC keymap — all four verified.
- **Skipped, browser-reserved** — a page cannot intercept these, so they can
  never be bound on the web at all. This is a permanent gap, not a to-do:
  `Ctrl-N` (New), `Ctrl-O` (Open), `Ctrl-W` (Close), `Ctrl-Q` (Exit, also
  out of scope). The New button remains the only route.
- **Skipped, feature does not exist yet** — bind these when the row lands:
  `Ctrl-F` (2.1 Find), `Ctrl-G` (2.2 Go to Line), `Ctrl-Shift-G` (2.3 Go to
  Scene), `F1` (2.5 Syntax panel). Also `Ctrl-Shift-S` (Save As — the web app
  has no Save As command; the dialog only appears when saving an untitled
  document), `Ctrl-M` (write mode is derived from the file extension, there is
  no toggle to bind), and `Ctrl-0` (no reset-zoom command exists).

### 1.5 Open Recent / MRU — `done`

- **Linux:** VM 235–253 (`MaxRecentFiles`, `RecentFiles`, `AddRecentFile`),
  481–496 (`OpenRecent`); MW.axaml 253–263 (menu)
- **Web:** `Editor.razor` 53–74 (FILES panel), 491–507 (`RequestOpenFile`)
- **Notes:** Per-browser, alongside 1.1. Small.
- **Done:** A RECENT section at the top of the FILES panel, above ALL SCRIPTS.
  Stored as `recentFiles` inside the existing `passage.session.v1` key, as row
  1.1 anticipated — not a second key. `AddRecentFile` mirrors the Linux
  remove/insert-at-0/trim, capped at `MaxRecentFiles = 8`.
  Entries are pruned in two places: on load, against `Library.List()`, so a
  stale list never renders; and on click, via `Library.Exists`, which mirrors
  the Linux `OpenRecent` — status message, drop the entry, no throw — since
  another client may have deleted the file since the list was written. The
  app's own delete also drops the name.

### 1.6 Undo / Redo — `done`

- **Linux:** MW.cs 249–276 (`UndoEditor`, `RedoEditor`,
  `ResetEditorUndoHistory`, `RedrawEditor`)
- **Web:** CodeMirror maintains its own undo stack; likely already works via
  Ctrl+Z with no toolbar affordance.
- **Notes:** Check first, then decide. The desktop trap in FUTURE_IMPROVEMENTS
  1 applies here too: any programmatic full-text replacement resets the undo
  stack. Web code paths that rewrite `_content` wholesale (`ResetEditorAsync`,
  `Editor.razor` 559–569, and any Beat Board write-back from Tier 3) must use
  ranged replacements, or undo history dies on every card edit.
- **Audit result:** Undo/redo already worked — `Ctrl-Z` / `Ctrl-Y` /
  `Ctrl-Shift-Z` come free from CodeMirror's default PC keymap; verified, not
  assumed. Nothing needed binding.
  `passage.setContent` is the wholesale-replacement path: `setValue` +
  `clearHistory` + caret to 1,1 + scroll to top. Three callers were audited.
  `OpenFileAsync` and `StartNewFileAsync` are **correct** — a different
  document should not inherit undo history. `RestoreRecoveryAsync` is
  **acceptable** — the text genuinely differs.
- **Bug found and fixed:** `ConfirmSaveAsAsync` also went through it, so
  **Save As destroyed the entire undo stack and threw the caret back to line 1**
  — for an operation that only renames the file. This is FUTURE_IMPROVEMENTS 1's
  desktop trap reproduced on the web. Fixed with `passage.refreshHighlights`,
  which re-applies line classes against the document already in the editor
  (Save As can still flip screenplay/markdown mode via the extension, so the
  classification does need refreshing — the text does not). Verified: history
  and caret both survive a Save As.
  The old helper was renamed `PushHighlightsForCurrentVersionAsync` →
  `ReplaceEditorContentAsync`, because the old name hid the fact that it wiped
  the document — the next caller would have walked into the same trap.
- **Affordances:** ↶/↷ buttons added to the status bar beside A−/A+, since the
  keyboard path was confirmed working first. They drive `cm.undo()`/`cm.redo()`
  and hand focus back to the editor.
- **For Tier 3:** the Beat Board write-back in 3.1b must use ranged replacements
  and must not call `ReplaceEditorContentAsync`, or undo dies on every card edit.

---

## Tier 2 — self-contained UI ports

Each is one panel or dialog with a clear boundary. Roughly one session each.

### 2.1 Find / Replace — `missing`

- **Linux:** MW.cs 829–857 (`ShowFindReplaceDialog`), 858–995 (`FindNext`,
  `FindPrevious`, `FindText`, `FindTextIndex`, `IsWholeWord`,
  `SelectEditorRange`), 996–1149 (`ReplaceCurrent`, `ReplaceAll`,
  `SelectionMatchesSearch`, `MapCaretAfterReplaceAll`, `IsWholeWordStatic`);
  `Views/FindReplaceDialog.axaml` + `.axaml.cs` (41 + 41 lines); VM 622–630
- **Web:** missing
- **Notes:** Already checked — the vendored bundle is **CodeMirror 5.65.18**
  and contains neither `searchcursor.js`, `search.js` nor `dialog.js`
  (`getSearchCursor` and `openDialog` are both absent). So the cheap path is
  open: **vendor the three matching 5.65.18 addon files** into
  `wwwroot/lib/codemirror/addon/` and wire them up, rather than porting the
  320 lines of Avalonia search logic. Match the Linux dialog's option set
  (match case, whole word, replace, replace all). Do not upgrade CodeMirror,
  and do not load the addons from a CDN — the app is self-hosted on a LAN with
  no guaranteed internet egress. Find must stay client-side (CLAUDE.md
  trap 4).

### 2.2 Go to Line — `missing`

- **Linux:** MW.cs 805–814 (`ShowGoToLineDialog`); VM 677–685;
  `Views/GoToLineDialog.axaml` + `.axaml.cs` (25 + 28 lines)
- **Web:** `Editor.razor` 690–695 (`JumpToLineAsync`) — the navigation half
  already exists
- **Notes:** Only the prompt is missing. Smallest row in Tier 2; good first one.

### 2.3 Go to Scene — `missing`

- **Linux:** MW.cs 815–828 (`ShowGoToSceneDialog`); VM 686–694;
  `Views/GoToSceneDialog.axaml` + `.axaml.cs` (74 + 98 lines)
- **Web:** `Editor.razor` 690–695 (`JumpToLineAsync`); scene list is already
  computed for the OUTLINE panel
- **Notes:** Filterable scene picker over data the analysis already produces.

### 2.4 Scratchpad panel — `missing`

- **Linux:** MW.axaml 579–621 ("Scratch" tab); VM 262–267
  (`HasScratchpadItems`, `ScratchpadEmptyMessage`, `ScratchpadSearchText`),
  1136–1146 (`DeleteScratchpadCard`); MW.cs 389–396
  (`ScratchpadItem_DoubleTapped`)
- **Web:** sidebar has FILES / OUTLINE / NOTES / GOALS —
  `Editor.razor` 46–52 (tab strip), 53–92 (panel bodies)
- **Notes:** Linux sidebar is Outline / Notes / Scratch / Goal. The web app
  gained a FILES tab (server-side library) and lost Scratch. Adding it makes
  five tabs — check the tab strip still fits at the sidebar width before
  committing. The Scratchpad has its own search box, which
  FUTURE_IMPROVEMENTS 9 suggests copying for outline filtering later.

### 2.5 Syntax Quick Reference panel — `missing`

- **Linux:** `Views/SyntaxQuickReferencePanel.axaml` (310 lines, whole file) +
  `.axaml.cs` (38 lines); VM 695–703 (`ToggleSyntaxPanel`); MW.cs 332–380
  (`ToggleSyntaxPanel`, `SetSyntaxPanelVisible`)
- **Web:** missing
- **Notes:** Almost entirely static content. Cheapest high-visibility win in
  Tier 2 — read the `.axaml` once and transcribe the reference table into a
  Razor component with the same headings and order.

### 2.6 Title Page editor — `missing`

- **Linux:** VM 1005–1104 (`EditTitlePage`);
  `Views/TitlePageDialog.axaml` + `.axaml.cs` (122 + 30 lines);
  `ViewModels/TitlePageViewModel.cs` (37 lines, whole file); MW.axaml 268
- **Web:** recognises title pages for rendering only — `Editor.razor` 252
  (`page.IsTitlePage`), 341 (`TitlePageEntry`), 675 (`"sx-titlepage"`)
- **Notes:** The web app can display a title page but not author one. The
  Fountain title-page block is key/value pairs at the head of the document, so
  the dialog is a form that splices lines 1..n. Reuse the same field order as
  the Avalonia dialog so exports match.

---

## Tier 3 — substantial work

Do not attempt any of these in one session. Each bullet under "split into" is
a session.

### 3.1 Beat Board editing — `partial` (web board is read-only)

- **Linux — board build:** VM 1537–1602 (`UpdateBeatBoardCards`), 1603–1668
  (`RebuildBeatBoardLanes`), 2738–2799 (lane / group / card view models)
- **Linux — editing:** VM 704–732 (`CreateNewCard`), 733–767 (`StartCardEdit`,
  `CancelCardEdit`, `SaveCard`), 768–802 (`SyncBoardToScript`, expand/collapse),
  803–841 (`AddSceneToBlock`), 842–885 (`DeleteCard`), 886–952
  (`UpdateCardInScript`)
- **Linux — text splicing:** VM 2586–2666 (`GetBeatBoardCardLineRange`,
  `GetCardOwnLineRange`, `GetCardLineRange`), 2667–2677 (`CountEditorLines`),
  2678–2737 (`MoveBeatBoardCardText`)
- **Linux — drag & drop:** MW.cs 1282–1348 (`BeatBoardCard_PointerPressed`,
  `_DragOver`, `_Drop`); MW.axaml 724–848 (Beat Board tab)
- **Web:** `Editor.razor` 759–776 (`RenderBoardCard`), 193–285 (centre pane)
- **Split into:**
  - 3.1a Act lanes and Sequence groups (render only) — port VM 1603–1668
  - 3.1b Inline card edit + write-back — port VM 733–767, 886–952, 2586–2666
  - 3.1c Add scene to block — port VM 803–841
  - 3.1d Delete card with confirmation — port VM 842–885
  - 3.1e Drag-and-drop reorder — port VM 2678–2737 + MW.cs 1282–1348
- **Notes:** The splicing helpers (VM 2586–2737) are pure string/list logic
  with no Avalonia dependency. **Extract them to `Passage.Core` rather than
  reimplementing them in Razor** — that gives both frontends one implementation
  and finally makes them testable, which is FUTURE_IMPROVEMENTS 7. Do that in
  3.1b and the rest of the tier gets cheaper. Note this is one of the few
  sanctioned reasons to edit `Passage.Core`; the desktop apps must keep
  building afterwards.

### 3.2 "Classify As" line-type override — `missing`

- **Linux:** VM 82 (`_lineTypeOverrides`), 953–972 (the seven `ClassifyAs*`
  commands), 973–990 (`SetLineTypeOverride`), 991–1004
  (`GetLineNumberFromCaretIndex`), 2211–2250 (`GetLatestEffectiveLineType`);
  MW.axaml 640–723 (Script tab context menu)
- **Web:** missing
- **Notes:** Needs a CodeMirror context menu plus a server-side override map
  keyed by line number, and the overrides must survive edits above the line —
  check how `SetLineTypeOverride` handles renumbering before designing the web
  side. Overrides are session state, not file content, so they belong with 1.1.

### 3.3 In-editor page-break rules — `missing`

- **Linux:** `Views/ScreenplayPageRuler.cs` (93 lines, whole file); wired in
  MW.cs 44–145 (editor setup) and 180–212 (`ApplyEditorWriteMode`)
- **Web:** live screenplay indentation exists (`passage.js`), page rules do not
- **Notes:** Dashed rule + PAGE pill every 55 lines. Implement as a CodeMirror
  decoration layer client-side. This is the cheap version of
  FUTURE_IMPROVEMENTS 13 (true discrete pages), which is out of scope — do not
  let a session drift into rewriting the editor.

### 3.4 Autocomplete + Enter-continuation — `missing`

- **Linux:** VM 136–140 (`AutoCompleteSuggestions`, `_uniqueSceneHeadings`,
  `_uniqueCharacterNames`), 2251–2290 (`UpdateUniqueScreenplayElements`),
  2291–2332 (`UpdateSuggestions`), 75 (`EnterContinuationText`); MW.cs 594–628
  (`EditorBox_KeyDown`), 629–698 (`SuggestionsListBox_Tapped`,
  `PositionAutoCompletePopup`, `ApplyAutoCompleteSuggestion`)
- **Web:** missing
- **Notes:** Must be client-side (CLAUDE.md trap 4). Push the character-name
  and scene-heading lists to JS after each parse and let CodeMirror's hint
  layer do the rest — do not round-trip per keystroke.

---

## Verify, don't assume

Rows where the web app may already be at parity. Check before scheduling work.

- **Markdown mode panels** — Linux VM 1259–1336 (`ApplyMarkdownDocument`),
  1416–1425 (`ToggleWriteMode`). Web has markdown handling in
  `Passage/Passage.Web/Services/DocumentAnalyzer.cs`. FUTURE_IMPROVEMENTS 4
  describes this as broken on Linux, so the web app may be *ahead* here.
- **Word/page count status bar** — present in both.
- **Goals (document, session, timer, progress ring)** — present in both;
  `Editor.razor` 777–890.
- **Exports (PDF / text / fountain)** — both use `Passage.Export`.

## Out of scope

Desktop-only by nature; do not port. Native file dialogs and `Exit` (the web
app has a server-side library instead), `scripts/install-linux.sh`, Wayland /
XWayland handling (`Passage/Passage.App.Linux/Program.cs`), and
FUTURE_IMPROVEMENTS 13 (true discrete editable pages).

## Test coverage

`Passage/Passage.Tests` is a custom runner covering `Passage.Parser` only
(151 lines). Anything extracted into `Passage.Core` during 3.1b should arrive
with tests in the same session — that is the cheapest moment to close
FUTURE_IMPROVEMENTS 7.
