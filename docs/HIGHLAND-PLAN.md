# Highland-informed upgrades — `Passage.Web`

The work queue for features identified in the Highland 2 teardown
(`docs/reference/HIGHLAND2_TEARDOWN.md`). Same scheme as `docs/WEB-PARITY.md`:
**this file is the source of truth for this work; update the row in the same
session as the change.**

Each row names the teardown section(s) to read. Read **only** those sections
(`sed -n '/^## 12\./,/^## 13\./p' docs/reference/HIGHLAND2_TEARDOWN.md`
style), never the whole reference. The reference is a behaviour spec: do not
reproduce Highland code, assets or branding.

Status values: `todo` · `partial` · `done` · `verified` (done and checked in
the running app).

Paths are relative to the repository root. `Layout` =
`Passage/Passage.Export/ScreenplayLayoutBuilder.cs`. `Analyzer` =
`Passage/Passage.Web/Services/DocumentAnalyzer.cs`. `Editor.razor` =
`Passage/Passage.Web/Components/Pages/Editor.razor` (5,170 lines: read ranges
only). `passage.js` = `Passage/Passage.Web/wwwroot/js/passage.js`.

---

## Standing decision — shared-library changes (read before any H1 row)

`ScreenplayLayoutBuilder.BuildPages` is called by the web app (Analyzer ~280),
by both desktop apps (Linux VM ~1301, WPF MainWindow ~2530) **and** by
`ScreenplayPdfExporter.Export` (`Passage/Passage.Export/ScreenplayPdfExporter.cs`
~21), which every frontend reaches through `ExporterCatalog` (the web via
`ExportService`). Web PDF output therefore only gets new pagination if the web
passes options into a PDF export path the desktop catalog instance does not use.
CLAUDE.md forbids changing desktop behaviour from shared code. Therefore:

- New pagination behaviour is **opt-in**: add an optional
  `PaginationOptions? options = null` parameter (or an overload).
  `null` must reproduce today's output byte-for-byte. Only `Passage.Web`
  passes options.
- Prove it: a test that builds pages for a fixed script with `null` options
  and asserts the same line/page output as before the change.
- New exporters (FDX, `.highland`) are **not** added to
  `ExporterCatalog.DefaultExporters` (that list feeds the desktop menus).
  `ExportService` in the web app references them directly.

If Adam later decides the desktop apps should get the new pagination too, that
is a separate, explicit decision. Do not make it implicitly.

---

## Tier H1 — Pagination that matches industry output

The biggest real gap. Today `Layout.Paginate` (267–284) cuts every 54 lines
regardless of content: scene headings can be orphaned at page bottoms,
dialogue splits mid-sentence with no (MORE)/(CONT'D).

### H1.1 Keep-together grouping — `done`

- **Reference:** teardown §11 (grouping), §12 (algorithm)
- **Passage:** `Layout` 147–173 (`BuildBodyLines`), 184–257
  (`BuildRenderableBlocks`, `TryCreateRenderableBlock`), 267–284 (`Paginate`)
- **Do:** Build groups before paginating, then fit groups, not lines:
  scene heading stays with the next element's first two lines; character cue +
  parenthetical stay with at least two dialogue lines; a transition stays with
  the element before it; section/synopsis/notes never start a page alone.
  Add a small overflow tolerance (Highland's `cheatPageBreakByAmount`):
  allow ≤1 line past the bottom margin rather than push a one-line orphan.
- **Success:** the H1 test script (create `Passage.Tests` fixtures) never ends a
  page on a scene heading or a bare character cue; `null` options output is
  unchanged (Standing decision test).
- **Landed:** `PaginationOptions` (`Passage.Export/PaginationOptions.cs`) and
  `ScreenplayLayoutBuilder.PaginateKeepTogether`; `ScreenplayPdfExporter` takes
  optional options. The web passes `WebPagination.Options` from
  `DocumentAnalyzer` (preview) and from `ExportService`, which swaps in its own
  PDF exporter instead of the shared catalog instance. A parenthetical is glued
  to the next element too. Section/synopsis/notes are dropped by
  `BuildRenderableBlocks` before layout, so that rule has nothing to act on.
  Tests: `Passage.Tests/Fixtures/H1Pagination.*` (golden captured before the change).

### H1.2 (MORE) / (CONT'D) dialogue splitting — `done`

- **Reference:** teardown §12 ("Splitting dialogue"), §13
- **Passage:** same `Layout` ranges; options object from H1.1
- **Do:** When a dialogue group must split: break only at a sentence boundary
  (fallback: line boundary if a single sentence exceeds the space), put
  `(MORE)` at the cue indent at the bottom, repeat the cue with ` (CONT'D)`
  at the top of the next page. Never split a parenthetical. Both strings
  configurable in options; defaults `(MORE)` and `(CONT'D)` (straight
  apostrophe — Passage convention; note Highland uses U+2019).
  Dual dialogue is never split; it moves whole.
- **Success:** fixture with a 70-line speech produces the MORE/CONT'D pair at
  a sentence end; preview and PDF show identical breaks.
- **Landed:** `PaginationOptions.MoreAndContinued` (default on), `MoreText`,
  `ContinuedText`; `ScreenplayLayoutBuilder.PaginateKeepTogether` glues each
  speech whole and `TrySplitSpeech` picks the break. A dialogue block may also
  break at its end when more of the speech follows. Preview and PDF share
  `BuildPages`, so they agree by construction; the test also checks the PDF
  streams. The parser hoists mid-speech parentheticals above the dialogue
  (and keeps the line in the dialogue text), so a speech has one dialogue
  block in practice. Tests: `Passage.Tests/Fixtures/H1MoreContd.fountain`.

### H1.3 Scene numbers — `done`

- **Reference:** teardown §3 (scene number row), Part 1 §5 Tools menu
- **Passage:** `Passage/Passage.Parser/FountainParser.cs` (scene heading
  parsing; grep `SceneHeading`), `Layout` scene-heading rendering
- **Do:** Parse a trailing `#…#` on scene headings into a `SceneNumber`
  (strip it from heading text). Render numbers in both margins in preview and
  PDF. Add "Add scene numbers" / "Remove scene numbers" commands that rewrite
  the text (numbers kept stable for already-numbered scenes).
- **Shared-code note:** parser change lands in desktop too. Acceptable only if
  headings *without* `#…#` parse exactly as before — add that test.
- **Success:** `INT. HOUSE - DAY #12#` previews as `12  INT. HOUSE - DAY  12`.
- **Landed:** `SceneHeadingElement.SceneNumber` (parser strips a trailing
  `#[A-Za-z0-9.-]+#` after the `[[id:…]]` comment, on plain, forced and
  overridden headings; a `#` mid-heading is left alone).
  `LayoutLine.SceneNumber` is set on the heading's first line; the PDF draws
  it at `LeftSceneNumberX` / `RightSceneNumberX`, the preview as absolutely
  placed spans, plain-text export as `12  HEADING  12`. Commands in the script
  menu: `SceneNumbers.Add` / `Remove` (`Passage.Web/Services/SceneNumbers.cs`)
  return line edits that `passage.replaceLines` applies as one undo step,
  skipping (and reporting) any line that changed since. Add never moves an
  existing number: a gap continues on (13 → 14), a scene between 5 and 6
  becomes 5A, then 5B. Desktop: headings with `#…#` no longer show the hashes
  in their text; desktop PDF and text export gain the margin numbers (shared
  exporters), but the Avalonia on-screen preview ignores `SceneNumber`. Tests: `Fixtures/H1SceneHeadings.*` (golden captured before
  the change) plus parse, layout/PDF and Add/Remove tests.

### H1.4 Auto-(CONT'D) in the editor — `todo`

- **Reference:** teardown §13
- **Passage:** `passage.js` (autocomplete/hint layer from WEB-PARITY 3.4);
  Analyzer per-line `SuggestionKinds`
- **Do:** Client-side only (CLAUDE.md trap 4). When the writer completes a
  character cue that matches the previous speaker in the same scene with only
  action between, offer ` (CONT'D)`. Compare names with `(V.O.)`, `(O.S.)`,
  `(CONT'D)` stripped. Typing `(CONT` completes to `(CONT'D)`. A toggle in
  settings; default on.
- **Success:** no Blazor round-trip per keystroke (check the WebSocket frames);
  offer appears only for the same-speaker case.

---

## Tier H2 — Interop

### H2.1 `.highland` import and export — `todo`

- **Reference:** teardown Part 1 §2
- **Passage:** `Passage/Passage.Web/Services/ScriptLibrary.cs` (open/save),
  `ExportService.cs`
- **Do:** Import: unzip, read `*.textbundle/text.fountain` (or `text.md` →
  Markdown mode), ignore unknown sidecars, surface any skipped content
  (PROJECT_RULES 12). Export: zip a TextBundle with `text.fountain`,
  `info.json` (`{"version":2,"type":"com.quoteunquoteapps.fountain","app.passage":{…}}`),
  and empty `characters.json`/`sprints.json`. Use
  `System.IO.Compression` (no new package).
- **Success:** a Highland sample file opens with identical text; Passage's
  export opens in Highland (Adam to verify on a Mac, or mark `done` not
  `verified`).

### H2.2 FDX (Final Draft) export — `todo`

- **Reference:** teardown Part 1 §5 (export list)
- **Passage:** new exporter class; wire via `ExportService` only (Standing
  decision)
- **Do:** Map Passage elements to FDX `Paragraph Type` values (Scene Heading,
  Action, Character, Parenthetical, Dialogue, Transition, Shot, General),
  dual dialogue as `DualDialogue`, title page to `TitlePage`. Scene numbers
  from H1.3 as `Number` attribute if present.
- **Success:** output validates as XML and opens in Final Draft / Fade In /
  WriterSolo with correct element types.

---

## Tier H3 — Writer tools

### H3.1 Character panel: list, rename, highlight — `todo`

- **Reference:** teardown Part 1 §6 (Character panel, Character Highlighting)
- **Passage:** Analyzer (unique character names already collected for 3.4);
  storage pattern of `SlateStore.cs`
- **Do:** Panel listing characters with line count and word count. Rename
  everywhere: rewrite cue lines only (preserve extensions like `(V.O.)`), one
  undoable edit. Highlight: pick a colour per character, applied as a
  CodeMirror mark on that character's dialogue blocks (client-side). Persist
  colours per document using the existing per-document store pattern.
- **Success:** rename is a single undo step; highlights survive reload.

### H3.2 Dialogue balance report (gender analysis) — `todo` · depends on H3.1

- **Reference:** teardown §15 (Gender report)
- **Do:** In the character panel, tag each character Male / Female /
  Unspecified (stored with H3.1 data). Report: character count by tag; lines
  spoken and words spoken by tag; scene co-occurrence counts per tag pair;
  per-character table (name, lines, words). On-screen first; PDF optional.
  Use theme tokens for colours (do not invent palette — Design Principles).
- **Success:** totals reconcile with the per-character table.

### H3.3 Typewriter scrolling and focus line — `todo`

- **Reference:** teardown §15 (editor view settings)
- **Passage:** `passage.js`
- **Do:** Client-side: keep the caret line vertically centred while typing;
  optional current-line highlight using an existing theme token. Both toggles
  persisted in the `passage.session.v1` localStorage key (WEB-PARITY 1.1).
- **Success:** no layout jank on long documents; toggles survive reload.

### H3.4 Reading time — `todo`

- **Reference:** Part 1 §6 (Statistics pane)
- **Passage:** `Passage/Passage.Core/TextAnalysis.cs`, stats display in
  `Editor.razor`
- **Do:** Add reading time beside word/page counts. Screenplays: 1 page ≈
  1 minute (use the existing page estimate); Markdown: words ÷ a configurable
  words-per-minute (default 230). Web-only display; do not change
  `TextAnalysis` outputs used by desktop (add a new method).

---

## Tier H4 — Power features

### H4.1 Header / footer directives in PDF — `todo`

- **Reference:** Part 1 §4 (Directives table)
- **Do:** Parse `{{HEADER: a | b | c}}` / `{{FOOTER: …}}` (left/centre/right
  slots; header default right, footer default centre), tokens `%p`, `%pp`,
  `%none`. Takes effect from the page it appears on. Directive lines never
  print. Web pagination path only (Standing decision).

### H4.2 Named markers + Jump to — `todo`

- **Reference:** Part 1 §4 (`{{%m: name}}`), §5 Tools menu
- **Do:** Treat `{{%m}}` / `{{%m: name}}` as non-printing markers; list named
  markers in the outline; Ctrl+J jump palette (client-side filter).

### H4.3 Outline filters — `todo`

- **Reference:** Part 1 §6 (Navigator), §15 (filter list)
- **Do:** Toggle visibility of sections, scenes, synopses, notes, markers in
  the outline; persist per browser.

### H4.4 Revision mode — `todo` · needs a design decision first

- **Reference:** §15 (Revisions), Part 1 §5
- **Do:** Session 1 is a **design note only**: how revised ranges are stored
  (sidecar vs inline), how they survive edits, colour sets (WGA sequence),
  and margin asterisks in PDF. Implement only after Adam approves the note.

---

## Not going in

- **Custom colour themes / Highland theme import.** Conflicts with
  `docs/Design Principles.md` and the fixed light/dark palette (WEB-PARITY 1.3).
- **Sprints, goals, milestones.** Passage already has Goals and Slate burst
  timers.
- **Bin, Scratchpad.** Scratchpad is blocked in WEB-PARITY 2.4; Slate covers
  idea capture.
- **Non-screenplay templates (Novel, MLA, Bulletin…).** Markdown mode covers
  prose; revisit only on request.
- **Locked pages / A-pages.** Worth doing eventually (Highland lacks them), but
  only after H1 and H4.4 exist.
