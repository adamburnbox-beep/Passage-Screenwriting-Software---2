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

### 1.4 Keyboard shortcuts — `partial` — everything bindable is bound

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
  no toggle to bind), and `Ctrl-0` (no reset-zoom command exists). The row
  stays `partial` permanently: the browser-reserved four can never be bound.

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

### 2.1 Find / Replace — `done`

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
- **Done, with one deviation from the session prompt.** Vendored
  **`searchcursor.js` only** (5.65.18, from cdnjs, committed to
  `wwwroot/lib/codemirror/` beside `placeholder.js` — the flat layout this repo
  already vendors addons in, not the upstream `addon/search/` nesting).
  `search.js` and `dialog.js` were downloaded, inspected and **not** vendored:
  `search.js` contains zero references to whole-word (verified by grep — the
  option does not exist in it), and decides case sensitivity from a smart-case
  heuristic (`query == query.toLowerCase()`) rather than an explicit checkbox.
  It therefore cannot produce the Avalonia option set the success criteria
  name, and its `openDialog` bar would be unthemed. Vendoring files that are
  then unused would be dead code (rule 2).
  So the option set is driven directly off `getSearchCursor`, which is what the
  addon is for — this is **not** the 320 lines of Avalonia index maths the
  prompt warned against.
  Whole word uses lookarounds, `(?<![A-Za-z0-9_])term(?![A-Za-z0-9_])`, not
  `\b`, so it behaves correctly when the term itself begins or ends with a
  non-word character. That mirrors `IsWholeWord`, which inspects the characters
  either side of the match and never the term.
  Wrap-around matches `FindText`: forward starts at the end of the selection
  and wraps to the top, backward starts at its head and wraps to the bottom.
  **Verified** against a crafted corpus — Replace All counts came out 8 / 6 / 4
  / 3 for substring, match-case, whole-word and both, matching the Avalonia
  semantics computed by hand, including `cat-like` counting as a whole word and
  `concatenated` not. Replace All is one undo unit, pushes through to the
  circuit, and the replaced text was confirmed on disk after a save.
- **UI deviation:** a docked bar under the topbar, not Avalonia's floating
  window — a modal would take focus from the editor it is searching, and there
  is no second window to put it in. Same fields, same two checkboxes, same
  buttons, plus Previous. `Ctrl-F` opens find, `Ctrl-H` opens it with replace
  shown; both browser defaults suppressed. Opening seeds the term from the
  selection.

### 2.2 Go to Line — `done`

- **Linux:** MW.cs 805–814 (`ShowGoToLineDialog`); VM 677–685;
  `Views/GoToLineDialog.axaml` + `.axaml.cs` (25 + 28 lines)
- **Web:** `Editor.razor` 690–695 (`JumpToLineAsync`) — the navigation half
  already exists
- **Notes:** Only the prompt is missing. Smallest row in Tier 2; good first one.
- **Done:** Modal matching `GoToLineDialog.axaml` — title "Go to Line", label
  "Line number:", placeholder "Enter line number...", Go and Cancel. Two
  deliberate differences from Avalonia: button order follows this app's other
  four modals (Cancel left, primary right) rather than Avalonia's Go-then-Cancel,
  since in-app consistency wins (rule 11); and bad input shows
  "Enter a line number of 1 or more." instead of Avalonia's silent refusal to
  close, which leaves the user guessing (rule 12).
  The input is focused on open — Avalonia focuses a new window's first control
  for free, a browser does not focus a freshly inserted node, so it is done
  explicitly. Out-of-range numbers clamp to the last line via
  `passage.scrollToLine`, matching `NavigateToLine`.
  Two entry points: `Ctrl-G` (bound now that the feature exists, per row 1.4)
  and the status bar's "Ln N", which is now a button — the web app has no
  Navigate menu to hang it off.

### 2.3 Go to Scene — `done`

- **Linux:** MW.cs 815–828 (`ShowGoToSceneDialog`); VM 686–694;
  `Views/GoToSceneDialog.axaml` + `.axaml.cs` (74 + 98 lines)
- **Web:** `Editor.razor` 690–695 (`JumpToLineAsync`); scene list is already
  computed for the OUTLINE panel
- **Notes:** Filterable scene picker over data the analysis already produces.
- **Done:** Modal matching `GoToSceneDialog.axaml` — title "Go To Scene",
  subtitle "Select a scene heading to jump to.", `line: heading` rows (the
  Avalonia `SceneJumpItem.DisplayText` format), Jump and Close, double-click
  and Enter to jump, first row selected on open.
  Scenes come from flattening `_analysis.Outline` on `Kind == "Scene"` — the
  same data the OUTLINE panel already renders, mirroring the Linux
  `FlattenScenes` over `OutlineRoots`. Nothing is reparsed.
  **Added beyond Avalonia**, per this row's own note: a filter box, focused on
  open, with ArrowUp/ArrowDown moving the selection and `scrollIntoView`
  keeping it visible. The selection is clamped when the filter narrows the
  list, or Jump would fire on an index that no longer exists. This is the
  filter FUTURE_IMPROVEMENTS 9 wants for the outline too — reuse it in 2.4.
  Bound to `Ctrl-Shift-G`, browser default suppressed.

### 2.4 Scratchpad panel — `blocked` — the Linux source is dead UI

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
- **Audit (do not port as specified):** The Linux Scratchpad does not work.
  `ScratchpadElements` is declared at VM 129 and **never populated** — there is
  no `.Add`, `.Insert` or assignment anywhere in `Passage.App.Linux`.
  `ScratchpadSearchText` is bound to the search box but **nothing filters by
  it**; typing only flips `ScratchpadEmptyMessage` between two strings, which
  makes it claim "No scratchpad cards match the current search" when nothing
  was ever searched. `DeleteScratchpadCard` has nothing to select and
  `ScratchpadItem_DoubleTapped` can never fire. Porting this row as written
  would add a fifth sidebar tab that is permanently empty, a search box that
  filters nothing, and a delete button that does nothing — speculative UI, and
  against rule 2.
- **Where the real feature lives:** `Passage.App` (WPF) has a working
  implementation — `MoveToScratchpadCommand` moves the current selection into
  the pad — and the shared model already exists as
  `Passage.Parser/ScreenplayModel.cs` 488 (`ScratchpadCardElement`, carrying a
  heading, description and a source line range). So the infrastructure is
  there; only the Linux port's populate path is missing.
- **Decision needed.** Options:
  1. **Defer until after 3.1** (recommended). "Moved scenes, notes, and loose
     ideas will appear here" describes cards moved off the Beat Board, so the
     populate path is 3.1d/3.1e. Build it once there is something to move.
  2. **Port from WPF instead.** Doable — the model is shared — but it needs a
     task that explicitly names `Passage.App`, which CLAUDE.md otherwise gates,
     and it is feature work, not the self-contained UI port this tier assumes.
  3. Port the empty shell for visual parity. Not recommended.

### 2.5 Syntax Quick Reference panel — `done`

- **Linux:** `Views/SyntaxQuickReferencePanel.axaml` (310 lines, whole file) +
  `.axaml.cs` (38 lines); VM 695–703 (`ToggleSyntaxPanel`); MW.cs 332–380
  (`ToggleSyntaxPanel`, `SetSyntaxPanelVisible`)
- **Web:** missing
- **Notes:** Almost entirely static content. Cheapest high-visibility win in
  Tier 2 — read the `.axaml` once and transcribe the reference table into a
  Razor component with the same headings and order.
- **Done:** `Components/SyntaxReferencePanel.razor` — a separate component, not
  more of `Editor.razor`. Same three cards in the same order (Sections,
  Synopses, Notes), same headings, blurbs and marker descriptions, transcribed
  verbatim. Right-hand dock like the Avalonia right panel; opened by `F1`
  (browser default suppressed, verified) or the topbar `?` button, since the
  web app has no View menu.
  **Fixed in passing:** the `.axaml` hardcodes each marker's dark-theme hex
  (`#4FC3F7`, `#FFB74D`, `#81C784`), so the Linux panel does not follow its own
  light theme. The web version routes them through `--syntax-section` /
  `--syntax-synopsis` / `--syntax-note` and is correct in both. Worth porting
  back — it is exactly the sweep the UI-overhaul doc's step 4 asks for.
  Copy buttons go through `passage.copyText`, which falls back from
  `navigator.clipboard` to a selection-based copy: the clipboard API needs a
  secure context and the app is served over plain HTTP on a LAN
  (`docs/web-app.md`), where it is simply absent. Unlike the Avalonia panel,
  which swallows clipboard errors, a failure is reported (rule 12).
- **Not ported:** the draggable splitter and remembered panel width
  (`SetSyntaxPanelVisible`, MW.cs 337–380). The panel is a fixed 320px. Raise
  it if resizing turns out to matter.

### 2.6 Title Page editor — `done`

- **Linux:** VM 1005–1104 (`EditTitlePage`);
  `Views/TitlePageDialog.axaml` + `.axaml.cs` (122 + 30 lines);
  `ViewModels/TitlePageViewModel.cs` (37 lines, whole file); MW.axaml 268
- **Web:** recognises title pages for rendering only — `Editor.razor` 252
  (`page.IsTitlePage`), 341 (`TitlePageEntry`), 675 (`"sx-titlepage"`)
- **Notes:** The web app can display a title page but not author one. The
  Fountain title-page block is key/value pairs at the head of the document, so
  the dialog is a form that splices lines 1..n. Reuse the same field order as
  the Avalonia dialog so exports match.
- **Done:** Form modal with the nine `TitlePageViewModel` fields, Credit
  defaulting to "written by" as it does there. Written in `EditTitlePage`'s
  **output** order — Title, Episode, Author, Credit, Source, Draft date,
  Revision, then Contact and Notes as `Label:` plus 4-space indented lines.
  That is not the order the view model declares them in (it has Credit before
  Author); the write order is what exports depend on, so that is the one
  copied. Empty fields are omitted, as there.
  `DocumentAnalysis` gained `TitlePageBodyStart` — the splice needs the
  boundary `TitlePageData.BodyStartLineIndex` already computes, and the web
  analysis was only projecting `Entries`. Delete drops lines `0..BodyStart`,
  matching the Avalonia Delete button.
  **Verified** by writing a block and reopening it: every field round-tripped
  through the parser, multi-line Contact included, and Delete left the body
  untouched.
- **Entry point:** the Export dropdown became **Document ▾** and now leads with
  "Title Page…". The web app has no File menu, and this is the closest thing.
- **Known cost:** a whole-document rewrite, so it clears the undo stack — the
  same trade-off `EditTitlePage` makes by assigning `EditorContent` outright.
  Splicing only the header range would preserve it; worth doing if 3.1b's
  ranged-replacement helpers get extracted to `Passage.Core`.

---

## Tier 3 — substantial work

Do not attempt any of these in one session. Each bullet under "split into" is
a session.

### 3.1 Beat Board editing — `done` (3.1a–3.1e; see the lane-drag limit below)

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
  - 3.1a Act lanes and Sequence groups (render only) — port VM 1603–1668 — **done**
  - 3.1b Inline card edit + write-back — port VM 733–767, 886–952, 2586–2666 — **done**
  - 3.1c Add scene to block — port VM 803–841 — **done**
  - 3.1d Delete card with confirmation — port VM 842–885 — **done**
  - 3.1e Drag-and-drop reorder — port VM 2678–2737 + MW.cs 1282–1348 — **done (cards only)**
- **3.1e done for cards.** `MoveBeatBoardCardText` is now
  `BeatBoardText.TryPlanMove`, which **returns the edit instead of applying it**:
  a move never changes the line count, so the affected span maps one-to-one onto
  its replacement and the whole reorder becomes a single contiguous splice. That
  keeps a drag to one undo step — verified, undo went 0 → 1 across a move and one
  undo restored the original order. `MainWindowViewModel` was repointed at it, so
  the refusals (onto itself, or a block onto a card nested inside it) are shared
  rather than duplicated. Seven tests now cover it.
  **Platform difference:** the desktop takes before/after from the target's
  *horizontal* midpoint; the web board stacks cards in a column, so this uses the
  *vertical* midpoint. It is measured in the browser, since only there is the
  card's real height known, and Blazor asks once on drop rather than per
  dragover.
- **Limit — lanes and groups are not draggable.** Acts and Sequences render as
  lane and group headers, not as cards in the flow, so they can be edited,
  added to and deleted, but not dragged. Reordering acts means reordering
  columns, which is a layout question this row did not cover. `TryPlanMove`
  already handles block moves and is unit-tested for them, so the logic is
  ready whenever the UI is.

- **3.1c/3.1d done.** Both go through `BeatBoardText.GetCardLineRange` and new
  ranged JS primitives — `insertLinesAt` and `deleteLineRange` — so neither
  clears the undo stack. Verified: undo depth rose across an add and a delete,
  and a single undo restored a deleted act whole.
  Add inserts after the container's **full block** (`includeNestedBlock: true`),
  so "+" on an act lands at the end of that act rather than under its heading,
  and writes the same three lines as `AddSceneToBlock`.
  Delete takes a Note's own lines and everything else's whole block, always
  confirms, and states the scope and line count in the Avalonia wording
  ("the act X and everything nested inside it (9 lines)").
  `deleteLineRange` takes the line break with the range so no blank line is left
  behind.
- **Gap closed while doing 3.1d:** 3.1a drew Acts and Sequences as lane and
  group headers, which left them with no edit or delete affordance — but on the
  desktop board they are cards like any other, and 3.1d's own wording covers
  them. The card editor was pulled out into a shared `RenderCardEditor`
  fragment, so a lane or group header swaps into the same form. Acts and
  sequences can now be renamed, re-typed and deleted.

- **3.1b web half done — ranged write-back.** Cards carry the parser's `Guid`
  through `OutlineNode` → `BoardCard`, and `DocumentAnalysis` now also exposes
  `Elements` and `LineCount`, which is what `BeatBoardText.GetCardLineRange`
  needs. An inline editor on each card (type, heading, synopsis) writes back via
  `BeatBoardText.BuildCardLines` and a new `passage.replaceLineRange`, which
  splices **only the card's own line range** with `cm.replaceRange`.
  That is the point: a `cm.replaceRange` is an ordinary edit, so it **joins the
  undo history and leaves the caret alone**, where a whole-document rewrite
  clears both. Verified — undo depth went 1 → 2 across a card edit, the caret
  stayed at line 3, and a single undo reverted the whole edit.
  `DocumentAnalyzer` gained `SuppressCardSynopses`, mirroring the desktop board
  build: synopsis lines shown as a card's description are marked
  `IsSuppressed`, which is what lets the card's own range swallow them. Without
  it, editing a description would leave the old `= ` lines behind — see the
  suppression note under 3.1b's extraction.
  Cards are written with `[[id:…]]`; the parser's `ExtractId` reads it back and
  strips it from display, so a card keeps its identity across a reparse and can
  be edited repeatedly. Verified by editing the same card twice and confirming
  the result on disk.
  Editing a card whose id has vanished from the document reports it rather than
  splicing at a stale offset.

- **3.1a done:** `BuildBoardLanes` now mirrors `RebuildBeatBoardLanes`: a flat,
  document-order card walk grouped into Act lanes containing Sequence groups
  containing cards, with implicit lanes/groups for cards that appear before any
  Act or Sequence heading. New `BoardLane` / `BoardGroup` records in
  `DocumentAnalyzer`; the old flat `BoardLanes` list of top-level Act/Sequence
  nodes is gone, as is `ToBoardCard`.
  Collapse/expand is included because the ported function exists to preserve
  `IsExpanded` across rebuilds — rendering it without a way to set it would
  leave a dead flag. Collapsed state is keyed by **heading text**, not the Linux
  card `Id`: the web re-parses on every keystroke and has no stable card
  identity, and line numbers move constantly. Verified that collapsing a lane
  and a group survives a re-parse.
  Still read-only — no card editing, adding, deleting or dragging yet.

- **3.1b extraction done (web card editing still outstanding):** The helpers now
  live in **`Passage/Passage.Parser/BeatBoardText.cs`** — `GetCardLineRange`,
  `BuildCardLines`, `ReplaceLines` — as pure static methods with no UI type in
  any signature.
  **Deviation:** the session prompt said `Passage.Core`. That is not possible:
  these operate on `ScreenplayElement`/`SectionElement`/`SceneHeadingElement`,
  and **`Passage.Parser` already depends on `Passage.Core`**, so putting them in
  Core would make the reference circular. `Passage.Parser` is referenced by
  Web, Linux and WPF alike, so the goal — one implementation, testable, shared —
  is met either way. No new `ProjectReference` was needed, so the Dockerfile
  copy list (trap 2) is untouched.
  `MainWindowViewModel` now delegates to it from `GetCardLineRange` and
  `UpdateCardInScript`; both keep their old signatures, so behaviour is
  unchanged. Verified `dotnet build Passage.Linux.slnf` still passes
  (0 warnings, 0 errors — the one pre-existing CS8601 in `MainWindow.axaml.cs`
  is untouched and unrelated).
  Five tests added to `Passage.Tests`, up from 5 to 10 total: card own range,
  nested section blocks, scene block extent, out-of-range input, and
  build/splice round-trip.
  **Behaviour worth knowing:** trailing synopsis lines are absorbed into a card's
  range *only once they are marked `IsSuppressed`*, which is what the board
  build does when it folds a synopsis into a card's description. On a raw parse
  nothing is suppressed, so a synopsis is a card in its own right. The tests
  cover both sides; this was found by probing the parser rather than assumed.

- **Notes:** The splicing helpers (VM 2586–2737) are pure string/list logic
  with no Avalonia dependency. **Extract them to `Passage.Core` rather than
  reimplementing them in Razor** — that gives both frontends one implementation
  and finally makes them testable, which is FUTURE_IMPROVEMENTS 7. Do that in
  3.1b and the rest of the tier gets cheaper. Note this is one of the few
  sanctioned reasons to edit `Passage.Core`; the desktop apps must keep
  building afterwards.

### 3.2 "Classify As" line-type override — `missing`

- **Linux:** VM 82 (`_lineTypeOverrides`), 915–933 (the seven `ClassifyAs*`
  commands), 935–951 (`SetLineTypeOverride`), 953–964
  (`GetLineNumberFromCaretIndex`), 2173–2211 (`GetLatestEffectiveLineType`);
  MW.axaml 640–723 (Script tab context menu)
  *(ranges re-checked after the 3.1b/3.1e extractions shortened the view model)*
- **Web:** missing
- **Notes:** Needs a CodeMirror context menu plus a server-side override map
  keyed by line number, and the overrides must survive edits above the line —
  check how `SetLineTypeOverride` handles renumbering before designing the web
  side. Overrides are session state, not file content, so they belong with 1.1.

### 3.3 In-editor page-break rules — `done`

- **Linux:** `Views/ScreenplayPageRuler.cs` (93 lines, whole file); wired in
  MW.cs 44–145 (editor setup) and 180–212 (`ApplyEditorWriteMode`)
- **Web:** live screenplay indentation exists (`passage.js`), page rules do not
- **Notes:** Dashed rule + PAGE pill every 55 lines. Implement as a CodeMirror
  decoration layer client-side. This is the cheap version of
  FUTURE_IMPROVEMENTS 13 (true discrete pages), which is out of scope — do not
  let a session drift into rewriting the editor.
- **Done:** `passage.js` gains a `.page-rules` overlay inside CodeMirror's
  sizer, redrawn on `changes` and `refresh`. Dashed rule plus a "PAGE N" pill at
  the right margin, snapped to a line boundary via
  `lineAtHeight`/`heightAtLine` so it sits between lines rather than through
  them. Colours come from the theme tokens, so it follows light and dark.
  **Measured in visual lines, not document lines.** The desktop editor has
  `WordWrap="True"`, so `ScreenplayPageRuler` works in visual space — every 55
  line *heights*. This editor wraps too, and with wrapping the two diverge, so
  the same visual measure is used. Verified: at a 22.5px line height the
  boundaries computed to 1237.5px and 2475px and snapped to 1232.5px and 2470px.
  Suppressed in the same two cases as the desktop: nothing is drawn for a
  document shorter than one page, and nothing in Markdown mode — matching
  `ApplyEditorWriteMode`, which adds the ruler for screenplays and removes it
  otherwise. All three cases verified (markdown 0 rules, screenplay 2, short
  script 0).
  The overlay re-creates itself if CodeMirror ever rebuilds the sizer, so it
  cannot silently disappear. Purely visual — the document is untouched and the
  Preview tab remains the exact layout. No editor rewrite (FUTURE_IMPROVEMENTS
  13 stays out of scope).

### 3.4 Autocomplete + Enter-continuation — `done` (autocomplete; see the Enter note)

- **Linux:** VM 136–140 (`AutoCompleteSuggestions`, `_uniqueSceneHeadings`,
  `_uniqueCharacterNames`), 2213–2251 (`UpdateUniqueScreenplayElements`),
  2253–2293 (`UpdateSuggestions`), 75 (`EnterContinuationText`); MW.cs 594–628
  (`EditorBox_KeyDown`), 629–698 (`SuggestionsListBox_Tapped`,
  `PositionAutoCompletePopup`, `ApplyAutoCompleteSuggestion`)
  *(ranges re-checked after the 3.1b/3.1e extractions shortened the view model)*
- **Web:** missing
- **Notes:** Must be client-side (CLAUDE.md trap 4). Push the character-name
  and scene-heading lists to JS after each parse and let CodeMirror's hint
  layer do the rest — do not round-trip per keystroke.
- **Done:** Vendored CodeMirror 5.65.18's `show-hint` addon (js + css) beside
  `searchcursor.js`, same approach as 2.1. The unique scene headings and
  character names are collected in `DocumentAnalyzer` (porting
  `UpdateUniqueScreenplayElements`, including taking names from Dialogue
  elements as well as Character ones) and pushed once per parse. Prefix
  matching, ordering and the 10-item cap happen in `passage.js`, mirroring
  `UpdateSuggestions` — **typing never round-trips** (trap 4).
- **Design note — where the classification lives.** The desktop decides which
  list a line wants with `GetLatestEffectiveLineType`, which falls back to
  `TextAnalysis.IsLiveCharacterCueCandidate` for lines the parse has no element
  for. That fallback composes four shared helpers; reimplementing it in JS would
  have been a real divergence risk. Instead `DocumentAnalysis` carries a
  per-line `SuggestionKinds` array computed server-side with the actual shared
  helpers, and the client only picks a list from it. Classification stays single
  and shared; matching stays client-side. The INT./EXT. override is applied on
  the client because it is instant, and the desktop applies it ahead of the
  parse too.
- **Verified:** typing `INT. KITCHEN` offered both matching headings; `EXT. GAR`
  offered the one; `MAR` offered `MARGARET` and not the other characters;
  `pick` applied the completion and replaced the whole line, as
  `ApplyAutoCompleteSuggestion` does.
- **Not verified — accepting with Enter/Tab.** The browser harness used for
  testing does not deliver synthetic Enter or Tab to CodeMirror at all (a
  Return there does not even insert a newline), so the accept keystroke could
  not be exercised end to end. The binding is show-hint's stock keymap and the
  underlying `pick` was confirmed working directly. **Worth a manual check.**
- **Enter-continuation was not ported — there is nothing to port.**
  `EnterContinuationText` (VM 75) is `=> "NewLine"` and is **referenced nowhere**
  in `Passage.App.Linux`. Like the Scratchpad in 2.4, it is a declared-but-unused
  property, so half this row's title has no implementation behind it. If
  Enter-continuation is wanted (Enter after a character cue dropping into
  dialogue, and so on), it is new feature work, not a port.

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
