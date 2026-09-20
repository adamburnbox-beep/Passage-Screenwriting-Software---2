# Writer's Tools handover

## What this doc is

State of play for the Writer's Tools work in `Passage.Web` (internally "Slate",
after the vault system it comes from), written for the next agent. Read this,
then `CLAUDE.md`, `PROJECT_RULES.md`, and `docs/SLATE-PLAN.md` — in that
order — before touching anything. Everything below is as of the
verification-pass commit on branch `web-scope`, 2026-09-20. PR #13
(web-scope → main) is merged; later commits go up on `web-scope` and need
a new PR.

---

## The one-paragraph version

The writer (Adam) has a set of writing-block exercises in his Obsidian vault
("the Slate"). We brought the *tools* — not the board, not the practice —
into the web app as a right-hand dock called **Writer's Tools**, plus one
full-screen overlay. **All six tool phases are built and verified**: the
foundations, Fill, the forward chain (WOAC and Character Flaw Brainstorm),
the Split family (A→Z→Split, Belief Split, Extend Backward), Bridge and
Position, Push/Pull and Lens, and Ideation burst. The only unbuilt row is
Phase 7, the model-assist seam, which the plan deliberately puts last and
argues against building until it is actually wanted. **There is no next tool
to build.** What comes next is Adam using the tools for real and reporting
back; the next agent's work is most likely fixes and copy changes from that,
not new runners.

---

## Where things are

| Thing | Path |
| --- | --- |
| The plan — phases, decisions, Rule 0, status and Done/Deviation notes per row | `docs/SLATE-PLAN.md` (**source of truth; update the row in the same session as any change**) |
| Copy-paste prompt per phase | `docs/SLATE-SESSIONS.md` |
| Sidecar store and every run's record type | `Passage/Passage.Web/Services/SlateStore.cs` |
| Placeholder scanner | `Passage/Passage.Parser/BracketScanner.cs` (+ `FountainMarkup.MaskOmissions`) |
| Synopsis placement (chain promotion, "scene under the caret") | `Passage/Passage.Web/Services/SynopsisPlacement.cs` |
| Split write-back and `= Turn: text` parsing | `Passage/Passage.Web/Services/SplitScript.cs` |
| Shape line | `Passage/Passage.Web/Services/ShapeLine.cs` |
| Alive / flat and Y / N pills | `Passage/Passage.Web/Components/AliveFlat.razor` |
| Dock markup and all handlers | `Passage/Passage.Web/Components/Pages/Editor.razor` — `<aside class="workshop">` ~389; tools view ~395; BRACKETS ~477; FILL ~510; chain list and runner ~548; Bridge / Position ~666; Revise ~876; Split / Belief / Extend ~1082; the ideation overlay `@if (_ideationOpen` ~1440; handlers from `// ---- Writer's tools dock` ~3200, then `// ---- Forward chain` ~3470, `// ---- The Split family` ~3700, `// ---- Bridge and Position` ~3930, `// ---- Push/Pull and Lens` ~4070, `// ---- Ideation burst` ~4190 |
| Client-side pieces | `Passage/Passage.Web/wwwroot/js/passage.js` — `applyWorkshopWidth` / `initWorkshopResize` ~125, `scrollToTop`, `replaceInLine` ~545, `insertLinesAt` ~561, `scrollToLine(line, focus)`, burst timer `startBurst` / `stopBurst` ~976 |
| Styles | `Passage/Passage.Web/wwwroot/css/app.css` from `/* ---- Writer's tools dock` ~954 (chain, split, burst, ideation overlay sections follow) |
| Tests | `Passage/Passage.Tests/Program.cs` — `TestBracketScanner*`, `TestSlateStore*`, `TestSynopsisPlacement`, `TestSplitScript*`, `TestShapeLine*` (the test project references `Passage.Web`) |
| Source material (the spec for every tool's text) | `/home/arosa/Sync/Obsidian Vault/Story/Slate/` — five live files in `Worksheets/` (the numbered Short/Medium/Long files are stubs), `How the Slate Works.md`, `loosening-up-practice.md`, `engines.md` |

`Editor.razor` is now ~4,400 lines. Read the ranges above, not the file.

---

## What's built, in one line each

Every row's full Done and Deviations notes are in `SLATE-PLAN.md`; this is
the map.

- **Phase 0 — foundations.** `SlateStore`: `<data>/.slate/<validated script
  name>.json`, `slateVersion: 1`, orphans pruned on the first interactive
  render with a status line, a corrupt sidecar left alone and saves refused.
  The dock: `TOOLS` button, drag-resizable, persisted in `passage.session.v1`;
  below 900px it covers the main area.
- **Phase 1 — Fill.** `BracketScanner` ranks `[SOMETHING …]` spans; "Hand
  me one"; the worksheet; **Fill it** replaces the span via `replaceInLine`
  and hands the next. A run is dropped on accept.
- **Phase 2 — forward chain.** `Chains` per script by seed line. WOAC round
  ≥ 2 stores no Want — it reads the previous Consequence live. The burst
  timer (decision E) in `passage.js`. "Read it back" promotes one line as a
  `=` synopsis under the nearest heading above the caret
  (`SynopsisPlacement`).
- **Phase 3 — Split family.** One `Split` / `Belief` / `Extend` run per
  script; belief cuts shared. Rungs behind "Keep going" buttons under the
  worksheet's stop-callouts. *Write the shape into the script* appends four
  acts / eight sequences with `= Turn: text` lines and brackets for unknown
  turns (`SplitScript`); *Update* fills only bracketed turns in place. Shape
  line in the status bar from `ShapeLine.Derive(BoardLanes)`.
- **Phase 4 — Bridge and Position.** `Bridges` by "A → Z", candidates as
  burst slots, wire read forward; `Positions` by the moment, slot select
  over `SplitScript.Turns` with the script's own turn text beneath, seven
  turns is the ceiling.
- **Phase 5 — Push/Pull and Lens.** `Revisions` by scene; *Use the scene
  under the caret*; six checks with Y/N; a **Y** opens that check's own Lens
  (seven lenses, `LensMinutes` clock on the fragment alone, multiline slot);
  diagnostic Belief Split and read-back behind buttons.
- **Phase 6 — Ideation burst.** `.slate/ideation.json`, belongs to no
  script, skipped by the orphan sweep. Full-screen overlay: roll or pick a
  lane, locked for the sitting; rounds of Input + timed Output; *Done for
  this sitting* files the one kept line and drops the rounds.

## Patterns every tool follows

Copy these, don't reinvent them.

- **Record type in `SlateDocument`**, saved on every `@bind:after="SaveSlate"`.
  Per-script tools that can run more than once per script (chains, bridges,
  positions, revisions) are a `List<>` shown as a list keyed by what the run
  is about — never a date — with "Start a new one" on top; empty runs are
  pruned in `SaveSlate`. Opening a tool with nothing kept goes straight into
  a new run. One-per-script tools (the Split family) open directly.
- **A `WorkshopView` enum value per view**, a "← back" that keeps state, and
  `ScrollWorkshopToTopAsync()` on open (a flag `OnAfterRenderAsync` acts on).
- **Every write into the script is ranged**: `replaceInLine` (matches by
  text, returns false if gone), `insertLinesAt`, `replaceLineRange`. Never
  `ReplaceEditorContentAsync`. Mirror the edit into `_content` and
  `RunAnalysis()` afterwards. Undo verified by hand every time.
- **Timers**: `passage.startBurst(rootId, readIn, seconds)` over the
  `textarea[data-slot]` elements under the root, in DOM order, starting at
  the first empty one; Enter or expiry advances; a `data-slot-multiline`
  slot keeps Enter as a line break; the display is `[data-burst-display]`
  inside the root. Blazor hears only `OnBurstEnded`. `_burstRunning` is one
  flag for the whole page — one run at a time.
- **Copy rule** (decision A): the worksheets' words, near-verbatim. Stop
  callouts are `.stop-callout`. Labels are the worksheet's labels. Adam
  notices generated-sounding lines within minutes.
- **Never a disabled button that reads a field just typed**: the change
  event and the tap travel together. Leave it enabled; put "nothing to do"
  in the status line.

---

## Verification pass, 2026-09-20

Every tool was walked through the six-point checklist below (Fill, WOAC,
Flaw, A→Z→Split, Belief Split, Extend Backward, Bridge, Position, Push/Pull
+ Lens, Ideation) at 375px and 1280px, in the agent browser against the
alt server on the dev data. Build warning-free, 28/28 tests. Everything
passed except one layout defect, fixed in the same session: at phone width
the status bar was `flex; nowrap`, so a runner's sentence-long status
("Filled in Plot point 1; Midpoint already has text in the script — edit
there") wrapped a word wide and grew the bar to 190px, a quarter of the
screen, taken from the dock. It now wraps with the message on its own row
(`app.css`, the `max-width: 900px` block). Two things noticed and left for
Adam to call:

- **The burst clock scrolls out of view on a phone.** `[data-burst-display]`
  sits in `.burst-row` at the top of the runner; once the burst focuses a
  slot in round 3+ of a chain, the writer can't see "read… 5" or the count.
  A sticky `.burst-row` inside `.workshop-body` while `_burstRunning` would
  fix it. Not done unasked — it changes how the runner looks.
- **Position keeps empty turns.** "Another turn" appends a turn; six empty
  ones from a click-happy session stay in the record and render as six
  empty forms. `SaveSlate` prunes empty *runs*; trailing empty turns could
  be pruned the same way. The dev sidecar for `split-test` has exactly this.

## What's next

Nothing is scheduled. In order of likelihood:

1. **Fixes and copy from real use.** Adam has not yet used any of this from
   a phone for a real session. Expect wording changes, field sizes, and the
   odd flow that reads fine but feels wrong in the hand. Treat each as a
   small session: read the row, change it, update the row's notes, commit.
2. **The burst timer's default seconds.** 30 (the vault's number) versus
   Adam's "10 seconds for three ideas". Configurable in every runner that
   uses it; the default was never settled with him.
3. **Phase 7 — the assist seam.** Only if Adam asks. The plan's own pushback
   (rule 2, speculative abstraction) still stands; the records are already
   the shape an assist would need.
4. **The scope question** in `SLATE-PLAN.md` "The honest risk": whether the
   board should come across from Obsidian after all. That is Adam's call
   after using the tools, not something to build unasked.

---

## Environment — read this before running anything

- **The `dotnet` on `PATH` is broken.** `/usr/bin/dotnet` is a stray
  `dotnet-host-10.0` apt package with no runtime under it. The real SDK is
  `~/.dotnet/dotnet` (9.0 and 10.0). In your Bash tool use the full path or
  `export PATH="$HOME/.dotnet:$PATH"`.
- **Build:** `dotnet build Passage.Web.slnf` (must be warning-free —
  `TreatWarningsAsErrors`). **Tests:** `DOTNET_ROLL_FORWARD=Major dotnet run
  --project Passage/Passage.Tests/Passage.Tests.csproj`.
- **Run as Development.** `Properties/launchSettings.json` sets it and port
  5210. Production outside a publish serves static files empty to browsers
  that ask for gzip — the page renders unstyled and never connects.
- **Adam runs his own server on 5210**, from his terminal, with the absolute
  path: `~/.dotnet/dotnet run --project ~/"Code
  Projects/passage-web/Passage/Passage.Web/Passage.Web.csproj"`. It does
  **not** hot-reload; he has to restart it to see a new build. Check with
  `pgrep -af Passage.Web` before assuming the port is free. **For your own
  verification use the agent browser's `preview_start` with
  `passage-web-alt`** (`.claude/launch.json`, untracked, port 5211) so his
  server is untouched; stop it when you're done or his next start fails
  with "address already in use". Both servers share the same dev data.
- **Dev data** lives at `Passage/Passage.Web/bin/Debug/net9.0/data/` (there is
  no `/data` on this machine). `fill-test.fountain` has four brackets and a
  Push/Pull run in its sidecar; `split-test.fountain` holds a written shape
  with six bracketed turns and Split, Bridge and Position runs. Sidecars are
  in `.slate/` beside them, `ideation.json` too.
- **Git shows ~150 files modified.** Those are exec-bit flips from the
  filesystem, not content. Commit with `git -c core.fileMode=false add
  <paths>` and `git -c core.fileMode=false commit`; never `git add -A`.
  Commits end with the `Co-Authored-By` line the session gives you. Push to
  `origin web-scope`. PR #13 (web-scope → main) is merged; open a new PR
  from `web-scope` when the next piece of work lands.

---

## Verifying a change

The minimum bar, all checked by hand in a browser:

1. Open it from the dock at phone width (375px) and at desktop width.
2. Type into it, press ←, reopen: the text is still there.
3. Reload the page and reopen the script: the text is still there (sidecar).
4. Any write into the script: Ctrl+Z restores the previous text exactly, the
   caret does not jump to line 1.
5. Edit the script underneath an open runner so its target is gone, then
   complete the runner: a status line, no exception, no wrong edit.
6. Nothing anywhere shows a percentage, a streak, a date, or a count of
   rounds presented as progress. Rule 0 wins over the row.

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
  on a phone that raises the keyboard over nothing. Pass `false` as the
  second argument to `passage.scrollToLine` from a runner.
- **A button disabled until a bound field has a value needs two taps** when
  the writer types and taps straight away. Leave it enabled; status line.
- **The dock body is one scroll container.** A JS call made inside a click
  handler runs before the new view renders; `ScrollWorkshopToTopAsync` sets
  a flag that `OnAfterRenderAsync` acts on.
- **Razor attributes can't hold a C# string literal in double quotes**
  (`@onclick="() => Foo("x")"` breaks the generated code). Use a `const` or
  a field.
- **Blazor `@bind` and a scripted `.click()`.** A programmatic click doesn't
  blur the focused textarea, so its change event never reaches the server.
  Use a real click (the `computer` tool) when testing a button that reads a
  field just typed.
- **The agent browser's "Return" key sends an empty `key`.** Use `key:
  "Enter"` when testing anything that listens for Enter. Its Enter also
  inserts no newline in a textarea — verify multiline handling by checking
  `defaultPrevented` on a dispatched `KeyboardEvent`, not by looking for the
  line break.
- **Save As does not exist** in the web app; the dialog only names untitled
  buffers. Don't build a cascade for it.
- **The agent browser's `ctrl+z` never reaches CodeMirror 5.** CM5 keys its
  bindings off `keyCode`, and the synthetic key has none. To test undo, use
  the status-bar Undo button, or dispatch a `KeyboardEvent('keydown', {key:
  'z', ctrlKey: true})` with `keyCode`/`which` defined as 90 on the focused
  editor textarea. `passage.editor` is the CM5 instance (`getLine`,
  `getCursor`, `setCursor`, `replaceRange`).
- **A 5-second burst timer expires between agent tool calls.** Each call
  takes a few seconds; read-in + write at the dev sidecar's 5 s is 10 s, so
  a slot "skips" between one call and the next. That is the clock advancing
  as designed, not a double Enter. Set the seconds higher before testing
  slot order, or check `document.activeElement` in the same batch.
- **Scripted editor edits leave a recovery snapshot.** `replaceRange` then
  undo returns the text but leaves the buffer dirty; the next reload asks
  "Recover Document?" — Discard.
