# Future Improvements — Linux Port (Passage.App.Linux)

Findings from a review of the Linux (Avalonia) port on 2026-07-02. Written for a
future agent/contributor: each item states the gap, where the code lives, and a
suggested approach. Items are ordered roughly by user impact.

> **Status update (2026-07-02, later the same day):** items 1–6 and 8 are DONE
> (see commit "Work through FUTURE_IMPROVEMENTS"), along with the page-count
> status-bar readout from item 12. Still open: item 7 (tests), item 9 (outline
> filter), item 10 (informational), item 11 (theme-dictionary refactor; the
> VM-side duplication was already removed), the rest of item 12, and item 13.
> The per-item notes below are kept for context on what was built and why.

## 1. Undo / Redo are empty stubs
- **Where:** `Passage/Passage.App.Linux/ViewModels/MainWindowViewModel.cs` —
  `Undo()` / `Redo()` command bodies are empty; Ctrl+Z / Ctrl+Y do nothing.
- **Approach:** AvaloniaEdit's `TextEditor` already maintains an undo stack
  (`editor.Undo()` / `editor.Redo()` / `editor.CanUndo`). Route the commands to
  the editor (the VM already reaches the window via `_window is Views.MainWindow`).
  Watch out for the programmatic `editorBox.Text = ...` sync in
  `MainWindow.axaml.cs` — full-text replacement resets the undo stack, so
  prefer `Document.Replace` diffs for VM-driven edits (drag-drop moves, card
  saves, title-page edits) to keep undo history intact.

## 2. File-save/open errors are silent
- **Where:** `MainWindowViewModel.Save()`, `SaveAs()`, `Open()` — exceptions go
  to `Debug.WriteLine` only. Violates PROJECT_RULES rule 12 ("fail visibly").
- **Approach:** Set `StatusMessage` at minimum; ideally show a dialog for
  save failures (data-loss risk). There is a `RecoveryPromptDialog` to use as a
  styling reference for a simple message dialog.

## 3. No unsaved-changes confirmation on New / Close / Exit / window close
- **Where:** `New()`, `Close()`, `Exit()` in the VM; `OnClosing` in
  `MainWindow.axaml.cs` saves session state but never prompts.
- **Approach:** When `IsDirty`, prompt Save / Discard / Cancel before clearing
  the document or closing the window.

## 4. Markdown mode is editor-only
- **Where:** Mode toggle is `WriteMode` in `MainWindowViewModel`; the editor
  swaps colorizers in `MainWindow.ApplyEditorWriteMode`. But the Outline,
  Notes, Beat Board, and Page Preview panels still run the Fountain parser on
  Markdown documents, producing junk or empty panels.
- **Approach:** In Markdown mode, build the Outline from `#`-heading levels
  (a tiny line-based scanner is enough — no need for a full Markdown AST) and
  either hide the Beat Board / Page Preview tabs or show a "Screenplay mode
  only" placeholder. Export menu should likewise be filtered by mode.

## 5. Beat Board: per-lane collapse and quick-add
- **Where:** `MainWindow.axaml` Beat Board tab; lanes/groups built in
  `MainWindowViewModel.RebuildBeatBoardLanes`.
- **Approach:** Now that Acts render as lanes and Sequences as groups, add an
  expand/collapse chevron per lane (persist like outline expansion keys) and an
  "+ Scene" affordance inside each group that inserts a scene heading at the
  right line (the line-range helper `GetBeatBoardCardLineRange` already knows
  block extents).

## 6. Card deletion from the Beat Board
- There is no way to delete a card/element from the board; users must edit the
  script text. Add a delete action (with confirmation) that removes the card's
  line range via the same text-splice pattern as `MoveBeatBoardCardText`.

## 7. Test coverage stops at the parser
- **Where:** `Passage/Passage.Tests` (custom runner, no framework) covers
  `Passage.Parser` only.
- **Approach:** The highest-value untested logic is in the Linux VM:
  `GetBeatBoardCardLineRange` (incl. section-block extension),
  `MoveBeatBoardCardText` / `MoveOutlineNodeText` splicing,
  `RebuildBeatBoardLanes` grouping, and `UpdateCardInScript`. All are pure
  string/list manipulation — easily testable if extracted to a helper class or
  by instantiating the VM with `null` window.

## 8. Recent files / file associations
- No "Open Recent" menu. Session restore only reopens the last document.
  Persist an MRU list in `SessionStorage` and surface it under File.

## 9. Find/Replace does not cover the Beat Board or Notes
- `FindReplaceDialog` searches the editor text only, which is acceptable, but
  a "find in outline/scenes" quick filter on the Workspace panel would help
  large scripts (the Scratchpad tab already has a search box pattern to copy).

## 10. Keyboard shortcut collisions to review
- Ctrl+M is now write-mode toggle; check against future keybindings.
- Zoom gestures (`Ctrl+OemPlus` / `Ctrl+OemMinus` / `Ctrl+D0`) have no visible
  menu hint because Avalonia's `KeyGesture.Parse` rejects "Ctrl+Plus"; if hints
  are wanted, bind display strings manually rather than `InputGesture`.

## 11. Theme palette definition is duplicated
- The palette exists in `App.axaml` (startup defaults) and
  `App.axaml.cs LoadThemeResources`. Keep them in sync when changing colors;
  longer-term, move both into `ResourceDictionary.ThemeDictionaries` so
  `RequestedThemeVariant` swaps them natively and the code-behind shrinks.

## 12. Editor niceties
- Spell check (e.g. via `WeCantSpell.Hunspell`) — a screenwriting staple.
- Session-scoped cursor position restore (session saves zoom but not caret).
- Smooth caret/scroll animation options; typewriter scrolling mode.
- Dual dialogue, page-count in status bar (estimator already exists:
  `ScreenplayPageEstimator`).

## 13. Windows/Linux feature parity audit
- The WPF app (`Passage/Passage.App`) has visuals/utilities not yet ported
  (`Visuals/`, `Utilities/` folders). Diff the two `MainWindowViewModel`s and
  list gaps — e.g. any print support, statistics, or reports present there.
