# Writer's Tools handover

## What this doc is

State of play for the Writer's Tools work in `Passage.Web` (internally "Slate",
after the vault system it comes from), written for the next agent. Read this,
then `CLAUDE.md`, `PROJECT_RULES.md`, and `docs/SLATE-PLAN.md` — in that
order — before touching anything. Everything below is as of the plan's
finishing commit on branch `web-scope`, 2026-09-20, which is on GitHub as
PR #14 (web-scope → main; PR #13 before it is merged).

---

## The one-paragraph version

The writer (Adam) has a set of writing-block exercises in his Obsidian vault
("the Slate"). We brought the *tools* — not the board, not the practice —
into the web app as a right-hand dock called **Writer's Tools**, plus one
full-screen overlay. **The plan is finished: every row in `SLATE-PLAN.md`
is `done` and every tool has been through the verification checklist
below.** Phases 0–6 are the tools themselves — foundations, Fill, the
forward chain (WOAC and Character Flaw Brainstorm), the Split family
(A→Z→Split, Belief Split, Extend Backward), Bridge and Position, Push/Pull
and Lens, Ideation burst. Phase 7 is the assist seam: an interface a model
can be plugged into later, invisible until one is. **There is nothing left
to build from the plan.** What comes next is Adam using the tools for real
and reporting back; the next agent's work is fixes and copy from that, or
the first real partner if he asks for one. Rule 0 in the plan (no required
fields, no progress, no history) outranks anything anyone asks for.

---

## Where things are

| Thing | Path |
| --- | --- |
| The plan — phases, decisions, Rule 0, Done/Deviation notes per row | `docs/SLATE-PLAN.md` (**source of truth; update the row in the same session as any change**) |
| Copy-paste prompt per phase | `docs/SLATE-SESSIONS.md` |
| Sidecar store and every run's record type | `Passage/Passage.Web/Services/SlateStore.cs` |
| Placeholder scanner | `Passage/Passage.Parser/BracketScanner.cs` (+ `FountainMarkup.MaskOmissions`) |
| Synopsis placement (chain promotion, "scene under the caret") | `Passage/Passage.Web/Services/SynopsisPlacement.cs` |
| Split write-back and `= Turn: text` parsing | `Passage/Passage.Web/Services/SplitScript.cs` |
| Shape line | `Passage/Passage.Web/Services/ShapeLine.cs` |
| Ideation's offline dealer | `Passage/Passage.Web/Services/IdeationDealer.cs` — `Deal(lane)` and the ten lanes' banks; called from `StartSitting`, `AddIdeationRound` and *Deal another* |
| The assist seam: interface, request, null partner | `Passage/Passage.Core/Extensibility/IStoryPartner.cs`; registered in `Passage/Passage.Web/Program.cs` |
| The asks, one builder per site | `Passage/Passage.Web/Services/Suggestions.cs` |
| The seam's only UI | `Passage/Passage.Web/Components/PartnerLines.razor` (renders nothing with the null partner) |
| Alive / flat and Y / N pills | `Passage/Passage.Web/Components/AliveFlat.razor` |
| Dock markup and all handlers | `Passage/Passage.Web/Components/Pages/Editor.razor` — `<aside class="workshop">` ~393; tools view, BRACKETS, FILL, chains, Bridge / Position, Revise, Split / Belief / Extend follow in that order to ~1440; the ideation overlay `@if (_ideationOpen` ~1448; handlers from `// ---- Writer's tools dock` ~3214, then `// ---- Forward chain` ~3479, `// ---- The Split family` ~3713, `// ---- Bridge and Position` ~3937, `// ---- Push/Pull and Lens` ~4085, `// ---- Ideation burst` ~4197, `// ---- The assist seam` ~4389 |
| Client-side pieces | `Passage/Passage.Web/wwwroot/js/passage.js` — `applyWorkshopWidth` / `initWorkshopResize` ~125, `scrollToTop`, `replaceInLine` ~545, `insertLinesAt` ~561, `scrollToLine(line, focus)`, burst timer `startBurst` / `stopBurst` ~976; `passage.editor` is the CodeMirror 5 instance |
| Styles | `Passage/Passage.Web/wwwroot/css/app.css` from `/* ---- Writer's tools dock` ~992 (chain, split, burst, partner, ideation overlay sections follow); the phone-width block is the second `@media (max-width: 900px)` |
| Tests | `Passage/Passage.Tests/Program.cs` — `TestBracketScanner*`, `TestSlateStore*`, `TestSynopsisPlacement`, `TestSplitScript*`, `TestShapeLine*`, `TestIdeationDealer*`, `TestSuggestions*`, `TestNullStoryPartner*` (the test project references `Passage.Web`) |
| Source material (the spec for every tool's text) | `/home/arosa/Sync/Obsidian Vault/Story/Slate/` — five live files in `Worksheets/` (the numbered Short/Medium/Long files are stubs), `How the Slate Works.md`, `loosening-up-practice.md`, `engines.md` |

`Editor.razor` is ~4,950 lines. Read the ranges above, not the file. The
line numbers drift; the `// ----` section markers don't.

---

## What's built, in one line each

Every row's full Done and Deviations notes are in `SLATE-PLAN.md`; this is
the map.

- **Phase 0 — foundations.** `SlateStore`: `<data>/.slate/<validated script
  name>.json`, `slateVersion: 1`, orphans pruned on the first interactive
  render with a status line, a corrupt sidecar left alone and saves refused.
  The dock: `TOOLS` button, drag-resizable, persisted in `passage.session.v1`;
  below 900px it covers the main area and the status bar wraps with the
  status message on its own row.
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
  lane, locked for the sitting; rounds of a dealt Input (`IdeationDealer`,
  re-dealable, editable) + timed Output; *Done for this sitting* files the
  one kept line and drops the rounds.
- **Phase 7 — the assist seam.** `IStoryPartner.SuggestAsync(request)` →
  lines. A request is the worksheet so far as labelled lines plus the
  field's own question. Six sites (Fill candidates, Split midpoint
  candidates, Bridge round, WOAC next answer, Flaw next question, Extend
  open link) show "Offer three" and the offered lines as taps — only when
  the registered partner is not the null one. A tap writes into the first
  empty target; the writer can take none. `NullStoryPartner` is the only
  implementation and is registered in `Program.cs`; nothing is sent
  anywhere.

---

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
- **Status lines are sentences and fail visibly** (rule 12): "no line in
  the script for Plot point 1", "[sic] is no longer on line 5 — nothing
  changed". The bar wraps at phone width so they stay readable.
- **A partner site is one line of markup**: `<PartnerLines
  Lines="PartnerLinesFor(key)" OnAsk="() => AskPartnerAsync(key)"
  OnUse="UsePartnerLine" />`, a key property that names the exact target
  (bracket, round, link), a builder in `Suggestions`, and a branch in
  `BuildPartnerRequest` / `PlacePartnerLine`. To test the wiring, register
  a throwaway partner that echoes the request, then delete it before
  committing.

---

## Open observations — Adam's call, not the next agent's

Found in the 2026-09-20 verification pass and deliberately left alone,
because each changes how a runner looks or what a record keeps:

- **The burst clock scrolls out of view on a phone.** `[data-burst-display]`
  sits in `.burst-row` at the top of the runner; once the burst focuses a
  slot in round 3+ of a chain, the writer can't see "read… 5" or the count.
  A sticky `.burst-row` inside `.workshop-body` while `_burstRunning` would
  fix it.
- **Position keeps empty turns.** "Another turn" appends a turn; six empty
  ones from a click-happy session stay in the record and render as six
  empty forms. `SaveSlate` prunes empty *runs*; trailing empty turns could
  be pruned the same way. The dev sidecar for `split-test` has exactly this.
- **The burst timer's default seconds.** 30 (the vault's number) versus
  Adam's "10 seconds for three ideas". Configurable in every runner that
  uses it; the default was never settled with him.

---

## What's next

Nothing is scheduled. In order of likelihood:

1. **Fixes and copy from real use.** Adam has not yet used any of this from
   a phone for a real session. Expect wording changes, field sizes, and the
   odd flow that reads fine but feels wrong in the hand. Treat each as a
   small session: read the row, change it, update the row's notes, commit.
2. **One of the open observations above**, if he calls it.
3. **The first real partner.** Only if Adam asks. It is one class
   implementing `IStoryPartner` in `Passage.Web`, its key in configuration,
   and swapping the registration line in `Program.cs` — nothing else in the
   app changes, because the six sites only check whether the partner is the
   null one. Keep the lines short and in the worksheet's register; the
   partner offers, never writes; a failure goes to the status line. That is
   the moment to add a config key, not before (plan, Phase 7 deviation).
4. **The scope question** in `SLATE-PLAN.md` "The honest risk": whether the
   board should come across from Obsidian after all. That is Adam's call
   after using the tools, not something to build unasked.

Whatever it is: one thing per session, the row in `SLATE-PLAN.md` updated in
the same commit, the checklist below run by hand.

---

## Environment — read this before running anything

- **The `dotnet` on `PATH` is broken.** `/usr/bin/dotnet` is a stray
  `dotnet-host-10.0` apt package with no runtime under it. The real SDK is
  `~/.dotnet/dotnet` (9.0 and 10.0). In your Bash tool use the full path or
  `export PATH="$HOME/.dotnet:$PATH"`.
- **Build:** `dotnet build Passage.Web.slnf` (must be warning-free —
  `TreatWarningsAsErrors`). **Tests:** `DOTNET_ROLL_FORWARD=Major dotnet run
  --project Passage/Passage.Tests/Passage.Tests.csproj` (30 as of this
  commit, all passing).
- **Run as Development.** `Properties/launchSettings.json` sets it and port
  5210. Production outside a publish serves static files empty to browsers
  that ask for gzip — the page renders unstyled and never connects.
- **Adam runs his own server on 5210**, from his terminal, with the absolute
  path: `~/.dotnet/dotnet run --project ~/"Code
  Projects/passage-web/Passage/Passage.Web/Passage.Web.csproj"`. It does
  **not** hot-reload; he has to restart it to see a new build. Check with
  `pgrep -af Passage.Web` before assuming a port is free. **For your own
  verification use the agent browser's `preview_start` with
  `passage-web-alt`** (`.claude/launch.json`, untracked, port 5211) so his
  server is untouched; stop it when you're done or the next start fails
  with "address already in use". Both servers share the same dev data.
- **More than one agent session may be working in this checkout at once.**
  It has happened: commits from a parallel session landed on `web-scope`
  between two of this session's, and its alt server was holding 5211 with a
  stale binary. Before you start: `git log --oneline -5`, `pgrep -af
  Passage.Web`, and if 5211 is held by a process running an old build, kill
  that process only — never the one on 5210.
- **Dev data** lives at `Passage/Passage.Web/bin/Debug/net9.0/data/` (there is
  no `/data` on this machine). `fill-test.fountain` has four brackets, two
  chains and a Push/Pull run in its sidecar; `split-test.fountain` holds a
  written shape with six bracketed turns and Split, Belief, Extend, Bridge
  and Position runs. Sidecars are in `.slate/` beside them, `ideation.json`
  too. **Back the directory up before a verification pass** (`cp -a` to
  your scratchpad) and restore it after — the runners write into it, and
  Adam's server reads the same files.
- **Git shows ~150 files modified.** Those are exec-bit flips from the
  filesystem, not content. Commit with `git -c core.fileMode=false add
  <paths>` and `git -c core.fileMode=false commit`; never `git add -A`.
  Commits end with the `Co-Authored-By` line the session gives you. Push to
  `origin web-scope`; PR #14 (web-scope → main) picks it up while it is
  open — update its description when something lands (`gh pr edit` can trip
  on a GitHub GraphQL deprecation; `gh api -X PATCH repos/<owner>/<repo>/pulls/14
  --input -` with `{"title","body"}` JSON works).

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

Every tool passed all six on 2026-09-20. If you touch the seam, add a
seventh: with `NullStoryPartner` registered, `document.querySelectorAll('.partner')`
is empty in every runner.

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
- **A long status message at phone width.** The bar is one flex row; before
  the wrap rule it squeezed a sentence to 52px wide and grew to a quarter of
  the screen. If you add a status bar item, check 375px.
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
- **The tool cards' group container matches their first line.** When
  clicking a tool by text from JS, `.tool-paths` (the group) has the same
  first line as its first card; click the card, or use the `find` tool's
  refs.
- **Save As does not exist** in the web app; the dialog only names untitled
  buffers. Don't build a cascade for it.
