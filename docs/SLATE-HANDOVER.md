# Writer's Tools handover

## What this doc is

State of play for the Writer's Tools in `Passage.Web` (internally "Slate", after
the vault system it comes from), written for the next agent. Read this, then
`CLAUDE.md`, `PROJECT_RULES.md` and `docs/SLATE-REDO-PLAN.md` (the plan for the
current build), in that order, before touching anything. As of the end of row
T8 (2026-10-25) the redo is finished: every row T1–T8 is `done`.

`docs/SLATE-PLAN.md` is the plan for the **replaced** build. Read its *Rule 0*
and decisions *B*, *D* and *E* (still binding) and nothing else of it; its UI
sections describe worksheets and markup that no longer exist.

---

## The one-paragraph version

The writer (Adam) has a set of writing-block exercises in his Obsidian vault
("the Slate"). The tools are a right-hand dock, **Writer's Tools**, that opens
a **deck of exercises**: twelve cards under a "what have you got?" chip row,
topped by up to three *Right now* suggestions worked out from the caret and the
project. An exercise is **data** (a definition in C#), a run is **generic**
(answers keyed by step), and **one runner** renders any of them, one question
at a time, anchored to something in the project (a character, a scene, a
bracket, an inbox beat, two Board cards, the Shape). Every answer can be sent
somewhere real: the Beat inbox, a synopsis, a bracket, a note page, the Board,
the Shape, the clipboard. The one exception is the **Ideation burst**, a
full-screen overlay that reads nothing from the project. Rule 0 (no required
fields, no progress, no history) outranks anything anyone asks for.

---

## Where things are

| Thing | Path |
| --- | --- |
| The plan for this build, with Done / Deviations notes per row | `docs/SLATE-REDO-PLAN.md` (**source of truth; update the row in the same session as any change**) |
| Rule 0, decisions B, D, E | `docs/SLATE-PLAN.md` (the rest is the replaced build) |
| Copy-paste prompt per row | `docs/SLATE-REDO-SESSIONS.md` |
| Sidecar store, `SlateDocument`, story data records, `IdeationDocument` | `Passage.Web/Services/SlateStore.cs` |
| v1 → v2 migration, and the old record types as private DTOs | `Passage.Web/Services/Tools/SlateMigration.cs` |
| The model: `Exercise`, `Step`, `StepKind`, `Shelf`, `AnchorKind`, `TargetKind` | `Services/Tools/Exercise.cs` |
| A run: `ExerciseRun { Id, ExerciseId, Anchor, Answers, Rounds, Passed, Sent }` | `Services/Tools/ExerciseRun.cs` |
| The deck: one static definition per exercise | `Services/Tools/ExerciseCatalog.cs` |
| Pure engine: current step, trail, loops, carry, branches, picks, dice | `Services/Tools/ExerciseEngine.cs` |
| Project lists for the pickers (cast, scenes, brackets, inbox beats, speakers) | `Services/Tools/ProjectContext.cs` |
| The ≤ 3 *Right now* suggestions | `Services/Tools/RightNow.cs` |
| Where an anchor starts a run's blank answers (Bridge's A and Z, Position's moment) | `Services/Tools/AnchorPrefill.cs` |
| Where answers go: inbox beat, note page, scene card, Board | `Services/Tools/ToolTargets.cs` |
| Shape answers ↔ `SplitRun` (what `SplitScript` writes from) | `Services/Tools/ShapeAnswers.cs`, `Services/SplitScript.cs` |
| The host interface the runner talks to | `Services/Tools/IToolHost.cs`; implemented by `EditorToolHost` in `Editor.razor` |
| Launcher, routing | `Components/Tools/ToolsDock.razor` |
| Runner: anchor strip, trail, step card, end panel, timer row | `Components/Tools/ExerciseRunner.razor`, `StepCard.razor`, `TrailLine.razor` |
| Pickers and the Send to menu | `Components/Tools/AnchorPicker.razor`, `SendToMenu.razor` |
| The Ideation overlay (own file, no project data) | `Components/Tools/IdeationOverlay.razor`, `Services/IdeationDealer.cs`, `.slate/ideation.json` |
| Alive / flat pills; the assist seam's only UI | `Components/AliveFlat.razor`, `Components/PartnerLines.razor` |
| The asks, one builder (`ForStep`) | `Services/Suggestions.cs` |
| The assist seam interface and null partner | `Passage.Core/Extensibility/IStoryPartner.cs`; registered in `Program.cs` |
| Placeholder scanner | `Passage.Parser/BracketScanner.cs` (+ `FountainMarkup.MaskOmissions`) |
| Synopsis placement ("scene under the caret") | `Services/SynopsisPlacement.cs` |
| Shape line in the status bar | `Services/ShapeLine.cs` |
| Client-side pieces | `wwwroot/js/passage.js` — `replaceInLine`, `insertLinesAt`, `scrollToLine(line, focus)`, burst timer `startBurst` / `stopBurst`, `getToolsPref` / `setToolsPref`, `restoreWhy`, `focusStep`, `applyWorkshopWidth`; `passage.editor` is the CodeMirror 5 instance |
| Styles | `wwwroot/css/app.css` from `/* ---- Writer's tools dock` (launcher, runner, step card, trail, anchor strip, ideation overlay) |
| Tests | `Passage.Tests/Program.cs` — `TestExerciseEngine*`, `TestCatalog*`, `TestSlateMigration*` (real v1 sidecars in `Passage.Tests/Fixtures/slate-v1/`), `TestSlateStore*`, `TestRightNow*`, `TestSplit*`, `TestIdeationDealer*` |
| Source material (the spec for every exercise's question text) | `/home/arosa/Sync/Obsidian Vault/Story/Slate/` — five live files in `Worksheets/`, `How the Slate Works.md`, `loosening-up-practice.md`, `engines.md` |

`Editor.razor` is ~6,300 lines. Read ranges, not the file. Line numbers drift;
the `// ----` section markers don't: `// ---- Writer's tools dock` (the TOOLS
button, width, open/close, Cast / Board / inbox entry points), `// ---- Fill:
brackets handed over one at a time`, `// ---- Burst timer and synopsis promotion`,
`// ---- The tools dock host` (the `EditorToolHost` wrappers, `ToolRightNow`,
`SendToolAsync` with every target), `// ---- The Shape: writing a Split run
into the script`, `// ---- Ideation burst` (just `_ideationOpen`). The dock shell
is `<aside class="workshop">`; the overlay is `<IdeationOverlay>` after the
status bar.

---

## How it fits together

- **The exercise is data.** `ExerciseCatalog` holds eleven `Exercise`
  definitions: a list of `Step`s of a fixed set of kinds (`Text`, `Lines`,
  `Pick`, `YesNo`, `AliveFlat`, `Branch`, `Dice`, `Slots`, `Stop`, `Loop`,
  `Summary`). Behaviour that differs between tools is expressed by step flags
  (`Clock`, `Opens`, `Source`, `Echo`, `ReadForwardKey`, `AltQuestion`, …), never
  by forking the runner. If a tool needs something the model can't say, add one
  step kind or flag **with a test**, not a special case in the runner.
- **The run is generic.** `ExerciseRun.Answers` is a string dictionary. Keys:
  round r > 0 is `key#r`, a `Lines` slot is `key[i]`, a `Slots` item is
  `key/item`. A run never stores a line number: anchors are a `Guid` where one
  exists (Board card `[[id:…]]`, beat), otherwise text matched at use time
  (bracket text, character name, scene heading).
- **`ExerciseEngine` is pure** and owns what is current, what the trail shows,
  loop expansion, carry (read live, never copied), branch visibility, picks and
  dice. Every rule in it has a unit test; add a test with any change.
- **The runner talks to `IToolHost`**, never to CodeMirror, `_content` or
  `_analysis`. `EditorToolHost` wraps methods `Editor.razor` already has, which
  is where every write trap below is handled once.
- **Sends** go through `IToolHost.SendAsync(TargetKind, ToolSend)`; each target
  reports on the status line what it did or why it didn't ("INT. KITCHEN is no
  longer in the script — nothing inserted").
- **Migration.** `slateVersion` 2 is current. `SlateStore.Load` always calls
  `SlateMigration.Migrate(document, rawJson)`, which reads the version from the
  raw text (any key case; missing means 1) and, below 2, converts the version 1
  lists (Chains, Fills, Split, Belief, Extend, Bridges, Positions, Revisions) into
  `Exercises`. The old record types survive only as private DTOs inside
  `SlateMigration`. Migration is in memory: the file is rewritten at version 2 on
  the next save (opening any run saves). A version 2 file is never migrated.
  `SplitRun` still exists (in `SplitScript.cs`, nine strings) because the Shape
  write is built from it.
- **The Ideation burst** belongs to no script. `IdeationOverlay` loads and saves
  `.slate/ideation.json` itself (`SlateStore.LoadIdeation` / `SaveIdeation`,
  skipped by the orphan sweep), deals each round's input from
  `IdeationDealer.Deal(lane)`, and runs the timer through the same
  `IToolHost.StartBurstAsync`. Its kept line can go to the Beat inbox as an
  untagged beat (the only thing it touches in the project, and it only writes).

### Adding an exercise

1. A static `Exercise` in `ExerciseCatalog` (and into `All`); questions from the
   vault worksheet, question verbatim, one hint, anything longer behind `Why`.
2. A card in `ToolsDock.Cards` (name, ≤ 8-word tagline, shelf, glyph).
3. Its `Targets` (what Send to offers), its `Anchor` kind (the Board card and
   Cast row Tools menus filter the catalog by it), and a `RightNow` rule if the
   caret or project should suggest it.
4. Tests: engine behaviour for any new rule, the catalog's own steps.
5. A row in `SLATE-REDO-PLAN.md`, then the checklist below by hand.

---

## Patterns every tool follows

- **Every write into the script is ranged**: `replaceInLine` (matches by text,
  returns false if gone), `insertLinesAt`, `replaceLineRange`. Never
  `ReplaceEditorContentAsync`. Mirror the edit into `_content` and
  `RunAnalysis()` afterwards (a second write in a row plans from stale text
  otherwise). Undo verified by hand every time.
- **Timers**: `passage.startBurst(rootId, readIn, seconds)` over the
  `textarea[data-slot]` elements under the root, in DOM order, starting at the
  first empty one; Enter or expiry advances; a `data-slot-multiline` slot keeps
  Enter as a line break; the display is `[data-burst-display]` inside the root.
  Blazor hears only `OnBurstEnded`. `_burstRunning` is one flag for the whole
  page: one run at a time. Nothing per-keystroke goes over the circuit.
- **Copy rule** (R-B): the question verbatim (≤ ~25 words), at most one muted
  hint line, the long explanation behind a collapsed **Why?** (`<details
  class="why" data-why>`, state remembered per browser by `passage.restoreWhy`).
  Stop callouts become an end-of-step panel with *Stop here* / *Keep going*.
  Adam notices generated-sounding lines within minutes.
- **Never a disabled button that reads a field just typed**: the change event
  and the tap travel together. Leave it enabled; put "nothing to do" in the
  status line.
- **Status lines are sentences and fail visibly** (rule 12). The bar wraps at
  phone width so they stay readable. A full-screen overlay has no status bar,
  so it shows its own note.
- **A partner site** is an `Offer = true` step: the runner shows `PartnerLines`
  under it and `Suggestions.ForStep` builds the ask from the trail. With
  `NullStoryPartner` nothing renders.

---

## Open observations — Adam's call, not the next agent's

- **The burst clock can scroll out of view on a phone** in a long step; a sticky
  timer row would fix it.
- **The burst timer's default seconds.** 30 (the vault's number) versus Adam's
  "10 seconds for three ideas". Configurable; the default was never settled.
- **The Ideation kept list doesn't remember what it sent to the inbox** after the
  overlay is closed, so *Send to Inbox* is offered again for an old line. A flag in
  `ideation.json` would fix it, but it is a log of what was sent, and Rule 0 says no.
- **Not built**: Casting Against Type, the five-layer build, structure remix
  (`SLATE-REDO-PLAN.md` §4.5); "scene cards before Z" for Extend; a Send to for the
  whole run rather than one answer.
- **T9, a real partner**, is blocked on Adam: whether, which model, where the
  key lives and what text leaves the machine. The seam exists (`IStoryPartner`,
  every `Lines` step can be a partner site). Needs its own plan.

---

## Environment — read this before running anything

- **The `dotnet` on `PATH` is broken.** `/usr/bin/dotnet` is a stray
  `dotnet-host-10.0` apt package with no runtime under it. The real SDK is
  `~/.dotnet/dotnet` (9.0 and 10.0). Use the full path or `export
  PATH="$HOME/.dotnet:$PATH"`.
- **Build:** `~/.dotnet/dotnet build Passage.Web.slnf` (must be warning-free —
  `TreatWarningsAsErrors`). **Tests:** `DOTNET_ROLL_FORWARD=Major
  ~/.dotnet/dotnet run --project Passage/Passage.Tests/Passage.Tests.csproj`
  (136 as of T8, all passing).
- **The Razor SDK 9 trap.** A line inside `@code` that starts with `<` builds
  locally on SDK 10 but breaks the container build in CI (SDK 9). Before a PR,
  build once pinned: a temporary `global.json` with `{"sdk":{"version":"9.0.317",
  "rollForward":"disable"}}` at the repo root, `dotnet publish
  Passage/Passage.Web/Passage.Web.csproj -c Release -o <scratch>`, then delete it.
- **Run as Development.** `Properties/launchSettings.json` sets it and port
  5210. Production outside a publish serves static files empty to browsers that
  ask for gzip — the page renders unstyled and never connects.
- **Adam runs his own server on 5210** and restarts it himself; it does not
  hot-reload. **Never use it for verification and never touch his data.** Run
  your own: copy the dev data to your scratchpad, add an untracked
  `.claude/launch.json` entry on a free port with `--no-launch-profile --urls
  http://localhost:<port> --Passage:DataDir=<scratch>/data`, start it with
  `preview_start`, and stop it when you're done.
- **Dev data** lives at `Passage/Passage.Web/bin/Debug/net9.0/data/` (there is
  no `/data` on this machine) with sidecars in `.slate/`. `fill-test.fountain`
  has two chains and a Push/Pull run; `split-test.fountain` a written shape and a
  Split, Extend, Bridge and Position run; `beat-test.fountain` five beats.
  Copies of the first two, as the old build wrote them, are the migration test
  fixtures. **Copy the directory before a verification pass**; the runners write
  into it.
- **More than one agent session may be working at once**, each in its own
  worktree. `git log --oneline -5` and `pgrep -af Passage.Web` first.
- **Git shows ~150 files modified.** Those are exec-bit flips from the
  filesystem, not content. Commit with `git -c core.fileMode=false add <paths>`;
  never `git add -A`. Commits end with the `Co-Authored-By` line the session
  gives you. One branch and one PR per row; merge `origin/main` before opening it.

---

## Verifying a change

The minimum bar, all checked by hand in a browser on your own server
(`SLATE-REDO-PLAN.md` §8):

1. Desktop width and 375px: the launcher is readable without scrolling past
   descriptions; the runner shows one question at a time; no horizontal overflow.
2. Type, ←, reopen: kept. Reload, reopen the script: kept (sidecar).
3. Keyboard only through three steps (real Enter key; the agent browser's
   "Return" sends an empty `key`, use `key: "Enter"`).
4. Every script write: undo restores exactly and the caret doesn't jump to line
   1 (test undo as the Traps section says).
5. Delete the anchor underneath an open run, then send: a status line, no
   exception, no wrong edit.
6. Rule 0: no counter, percentage, date, streak or "done" mark anywhere.
7. `document.querySelectorAll('.partner').length === 0` with the null partner.
8. Build warning-free (also once on SDK 9), tests pass, CI green.

If you touch the sidecar format, also open each script in the dev data and check
its `.slate/*.json` is written at `SlateVersion` 2.

---

## Traps that have already bitten

- **Prerendering runs `OnInitialized` twice.** Anything that sets a one-shot
  status message on load has to run in the `firstRender` block of
  `OnAfterRenderAsync`, or the second pass overwrites it.
- **The server's `_content` is a debounce behind the editor.** Never trust a
  column from the last parse for an edit; match by text on the line
  (`replaceInLine` does this). After a successful client-side edit, mirror it
  into `_content` and re-run the analysis so the next action sees it.
- **Jumping to a line focuses the editor** and steals focus from the runner —
  on a phone that raises the keyboard over nothing. Pass `false` as the second
  argument to `passage.scrollToLine` from a runner.
- **A button disabled until a bound field has a value needs two taps** when the
  writer types and taps straight away. Leave it enabled; status line.
- **Sidecar JSON keys are PascalCase, and a hand-written or old one may not be.**
  `a.fountain.json` in the dev data has `"slateVersion":1`. `SlateDocument.SlateVersion`
  defaults to the current version, so a sidecar with the key missing or cased
  differently would pass for current if the version were read from the deserialised
  object. `SlateMigration.Migrate` reads it from the raw text.
- **The dock body is one scroll container.** A JS call made inside a click
  handler runs before the new view renders; use a flag that `OnAfterRenderAsync`
  acts on.
- **A long status message at phone width.** The bar is one flex row; before the
  wrap rule it squeezed a sentence to 52px wide. If you add a status bar item,
  check 375px.
- **`SetStatus` must re-render the page.** The dock's own events don't re-render
  the status bar; `EditorToolHost.SetStatus` calls `StateHasChanged`.
- **Razor attributes can't hold a C# string literal in double quotes**
  (`@onclick="() => Foo("x")"` breaks the generated code). Use a `const` or a
  field.
- **Blazor `@bind` and a scripted `.click()`.** A programmatic click doesn't blur
  the focused textarea, so its change event never reaches the server. Use a real
  click when testing a button that reads a field just typed, or dispatch a
  `change` event on the field first.
- **The agent browser's "Return" key sends an empty `key`.** Use `key: "Enter"`
  when testing anything that listens for Enter. Its Enter also inserts no newline
  in a textarea — verify multiline handling by checking `defaultPrevented` on a
  dispatched `KeyboardEvent`.
- **The agent browser's `ctrl+z` never reaches CodeMirror 5.** CM5 keys its
  bindings off `keyCode`, and the synthetic key has none. To test undo, use the
  status-bar Undo button, or dispatch a `KeyboardEvent('keydown', {key: 'z',
  ctrlKey: true})` with `keyCode`/`which` defined as 90 on the focused editor
  textarea.
- **The agent browser's screenshots lag a render behind a click**, and `find`
  does not see everything in the full-screen overlay. Read the DOM
  (`javascript_tool`) for state and screenshot twice for layout.
- **A short burst timer expires between agent tool calls.** Each call takes a few
  seconds; set the seconds higher before testing slot order, or check
  `document.activeElement` in the same batch.
- **Scripted editor edits leave a recovery snapshot.** `replaceRange` then undo
  returns the text but leaves the buffer dirty; the next reload asks "Recover
  Document?" — Discard.
- **Save As does not exist** in the web app; the dialog only names untitled
  buffers. Don't build a cascade for it.
