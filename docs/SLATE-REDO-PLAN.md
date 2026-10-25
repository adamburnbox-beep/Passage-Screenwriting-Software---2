# Writer's Tools, redone — plan

Replaces the *UI and run model* of the Writer's Tools dock (internally
"Slate"). Same scheme as the other plans: numbered rows, one row per session,
status per row, this file is the source of truth and is updated in the same
commit as the change. Copy-paste prompts are in `docs/SLATE-REDO-SESSIONS.md`.

Status values: `todo` · `partial` · `done` · `verified` (done and checked in
the running app).

Read with it: `CLAUDE.md`, `PROJECT_RULES.md`, and **only** these parts of
`docs/SLATE-PLAN.md`: *Rule 0*, decisions *B*, *D* and *E*. The rest of
SLATE-PLAN describes the build being replaced; its Done notes are useful for
finding code, not for design.

---

## 1. What's wrong now (review, 2026-10-25)

Checked in the running app on `split-test.fountain` at desktop width.

1. **The launcher is an essay.** Ten tools, each with a 25–45 word
   description, under five headed groups, in a 340px column. You read
   ~450 words to pick one. Nothing about the current script influences what
   is offered.
2. **Every runner is a printed worksheet.** A→Z→Split opens with a preamble
   paragraph, then rung headings, a paragraph per rung, six labelled
   textareas, a stress question, alive/flat pills, a checkbox, a stop-callout,
   a "keep going" button and an "update the script" button with its own
   paragraph — all on screen at once. Push/Pull is worse (GSU table, six
   checks, eight lenses, diagnostic split, read-back). The writer is doing
   reading comprehension, not writing.
3. **The worksheets were written for a partner.** The vault's own
   description (`How the Slate Works.md`, *What I do*) is "hand you one thing
   … three options, expected to be rejected … write your answers back into
   the file". The dock kept the worksheet and dropped the partner, so each
   tool is a form with homework instructions and empty boxes.
4. **Almost no integration with the project.** Inputs are free text: you
   type a character's name, a scene, a start and end event, even when the
   script, Cast, Board and Beat inbox already hold them. Outputs are three
   paths (fill a bracket, insert one `=` line, write the Shape). Everything
   else stays in sidecar "working-out" that no other view shows. The story
   tools built since (Cast colours, storylines, tagging, Beat inbox, project
   notes) are not used at all.
5. **The code can't take a redesign in place.** Dock markup is
   `Editor.razor` ~1017–2092, the ideation overlay ~2122, handlers ~6016–7360:
   ~2,600 lines of bespoke markup and handlers per tool inside a 7,864-line
   page. Ten record types in `SlateStore.cs`, one per worksheet.

## 2. The idea in one paragraph

Turn the dock into a **deck of exercises**. Each exercise is a short
sequence of **one question at a time**, anchored to **something in the
project** (a character, a scene, a bracket, an inbox beat, two Board cards,
the story's shape), and every answer can be **sent somewhere real**: the Beat
inbox, a scene's synopsis, the bracket it came from, a character's note page,
the Shape. The launcher is a grid of small cards, filtered by one chip row
("what have you got?") and topped by up to three **"right now"** suggestions
derived from where the caret is and what the project holds. Exercises can
also be launched from the thing they're about: a Board card, a Cast row, an
inbox beat, a bracket under the caret. Rule 0 stays binding throughout.

---

## 3. Standing decisions

### R-A. Rule 0 still outranks everything

`SLATE-PLAN.md` Rule 0 applies to every row here, plus three readings of it
for the new UI, binding:

- **No step counter.** No "3 of 8", no dots, no progress bar in the runner.
  The trail of answered questions is the only record of where you are.
- **No completion state on launcher cards.** No ticks, no "done", no counts
  of runs. A card may list kept runs *by their anchor* ("Maggie",
  "INT. KITCHEN"), never by date or number.
- **Every step can be skipped and every exercise closed** at any point with
  no prompt. "Done for now" is always visible.

### R-B. Copy: the question verbatim, the explanation folded

*This changes SLATE-PLAN decision A's copy rule. Settled 2026-10-25 — see §7.*

- The **question** on each step keeps the worksheet's words (≤ ~25 words).
- At most **one hint line** under it (≤ ~15 words), muted.
- Anything longer — the *why*, the worked example, the "proof it works"
  column — goes behind a **"Why?"** disclosure, collapsed by default and
  remembered per browser (not per run).
- Stop-callouts keep their words but become an **end-of-step panel** with two
  buttons ("Stop here" / "Keep going"), not a paragraph between fields.
- Launcher cards: name + one line of ≤ 8 words. The long description moves to
  the exercise's intro "Why?".

### R-C. Exercises are data; one runner renders them all

No more bespoke markup per worksheet. An exercise is a definition (C#, in
`Passage.Web/Services/Tools/ExerciseCatalog.cs`), a run is generic
(answers keyed by step), and a single runner component renders any of them.
Behaviour that genuinely differs between tools (WOAC carry, Split writing the
Shape, Position over the seven turns, a Revise **Y** opening its Lens) is
expressed with a small, fixed set of step kinds and hooks — not by forking
the runner. If a tool needs something the model can't express, add one step
kind, with a test; do not special-case the runner.

### R-D. Anchors and outputs go through one host interface

The runner never touches CodeMirror, `_content` or `_analysis` directly. It
talks to an `IToolHost` that `Editor.razor` implements, which keeps every
existing trap handled in one place: ranged writes only (SLATE-PLAN D,
WEB-PARITY 1.6), mirror into `_content` and re-analyse, `scrollToLine(…,
false)` from a runner, status lines that fail visibly (rule 12).

### R-E. Storage: one generic run list, migrated once

`SlateDocument` gets `List<ExerciseRun> Exercises`. The ten per-tool record
types are **moved** into it, one type at a time, in the row that deletes that
type's old view (`SlateMigration.Migrate` converts the records, then clears
the old list, so a second run finds nothing). Migration is not gated by
version; `slateVersion` goes 1 → 2 only in T8, when the old record types are
deleted. `StoryData` is **not** touched (STORY-PLAN decision
A, and its note that "a Slate sidecar revamp is coming" — this is that
revamp for the worksheet runs only; flag any conflict, don't guess). A run
never stores a line number; anchors are a Board/beat `Guid` where one exists,
otherwise text matched at use time (bracket text, character name).

### R-F. Client/server split stays as it is

Burst timers stay in `passage.js` (`startBurst` / `stopBurst`, which already
work over `textarea[data-slot]` under a root id — the step card exposes its
inputs that way). Nothing per-keystroke goes over the circuit (CLAUDE.md
trap 4). Answers save on change, like today.

---

## 4. The design

### 4.1 Launcher (dock top level)

```
┌ WRITER'S TOOLS ─────────────────── × ┐
│ RIGHT NOW                            │
│ ▸ Fill  [SOMETHING forces her to…]   │  ← caret is in a bracket
│ ▸ Push/Pull  INT. LIGHTHOUSE – NIGHT │  ← caret is in a scene
│ ▸ Position  "The keeper's letter"    │  ← inbox has a beat
│                                      │
│ What have you got?                   │
│ (All)(Nothing)(An ending)(A moment)  │
│ (A scene going)(A dead scene)        │
│                                      │
│ ┌──────────┐ ┌──────────┐            │
│ │ ⚡ Burst  │ │ ⇄ Bridge │            │
│ │ Loosen up│ │ Connect  │            │
│ │ fast     │ │ two cards│            │
│ └──────────┘ └──────────┘  …         │
└──────────────────────────────────────┘
```

- **Right now** (max 3, hidden when empty), derived on the server from
  `_caretLine`, `_analysis` and `_slate` by a pure `RightNow.Suggest(...)`:
  caret inside a bracket → Fill it; caret inside a scene → Push/Pull,
  Scene Dials or Understudy on that scene (one of them, rotate on each
  open); inbox beats exist → Position the newest; no Shape and no `#`/`##`
  headings → A→Z→Split; a character in Cast with no note page → Flaw
  Brainstorm on the character with the most lines. "One thing, not a menu"
  (Rule 0.5) is honoured by showing at most three and making each one tap
  straight into a run with its anchor filled.
- **Chip row**: the five "what have you got?" rows from the vault's
  `0 — Which Worksheet`, short labels. Default *All*; selection remembered in
  `passage.session.v1`.
- **Cards**: two columns at ≥ 360px dock width, one below. Icon, name, ≤ 8
  word tagline, and a small anchor glyph (👤 character, 🎬 scene, [ ] bracket,
  ⇄ two points, ◇ shape, none). Under a card, kept runs as anchor chips;
  tapping a chip resumes it.

### 4.2 Runner (one question at a time)

```
┌ ← WOAC chain                   Done for now ┐
│ ┌ 👤 MAGGIE · wants the lighthouse sold ─ ↗ ┐│  ← anchor strip
│ └───────────────────────────────────────────┘│
│  Want      the lighthouse sold          ✎    │  ← trail: answered steps,
│  Obstacle  her brother won't sign       ✎    │    one line each, tap to edit
│                                              │
│  What does she DO about it?                  │  ← the question, 17px
│  Action, not feeling.            Why? ▾      │  ← hint + disclosure
│  ┌────────────────────────────────────────┐  │
│  │                                        │  │
│  └────────────────────────────────────────┘  │
│  [ Next ⏎ ]   Skip    ⏱ 30s     Send to ▾    │
└──────────────────────────────────────────────┘
```

- **Anchor strip** at the top: the thing this run is about, with its colour
  (Cast / storyline colour), a ↗ that jumps to it in the script without
  stealing focus, and a "change" affordance. A missing anchor (scene deleted
  underneath) renders as "INT. KITCHEN — no longer in the script" and the run
  still works (rule 12, SLATE-PLAN risk 3).
- **Trail**: previous answers as single lines; tap one to reopen that step.
  No numbers.
- **Step card**: the question, hint, *Why?*, the input(s) for the step kind,
  and the actions. Enter = Next on single-line steps; Shift+Enter breaks a
  line on multiline ones (the existing `data-slot-multiline` rule).
- **Timer** is a per-run toggle in the step card's action row, using the
  existing burst settings; with it on, Next is automatic on expiry (decision
  E, unchanged).
- **Send to ▾** on any answered step and on the end panel: the exercise's
  allowed targets (§4.4). After a send, the trail line shows a small "→ Inbox"
  / "→ synopsis" tag so the writer can see where it went — not a count.
- **Stop points** (the worksheets' stop-callouts) render as an end-of-step
  panel: the callout's words, *Stop here* (closes, run kept) and *Keep going*.
- **Partner seam**: a step flagged `Offer = true` shows `PartnerLines` under
  its input. With `NullStoryPartner` registered nothing renders (existing
  check: `document.querySelectorAll('.partner')` is empty).

Width: default dock width rises from 340 to 400px; the runner has to be
readable at 340 (narrowest drag) and at 375px phone width, where the dock
already covers the main area.

### 4.3 Step kinds (the whole vocabulary)

| Kind | Renders | Used by |
| --- | --- | --- |
| `Text` | one textarea (single or multiline) | almost everything |
| `Lines(n)` | n short slots, burst-able, "another" link if `Open` | Fill candidates, Bridge round, Split candidates, Ideation output |
| `Pick(fromKey)` | the lines of an earlier `Lines` step as tap targets + "none of these, write it" | Split chosen midpoint, Bridge picked link |
| `AliveFlat` / `YesNo` | the existing `AliveFlat` pills | Split stress test, Revise checks, Position |
| `Branch(on, value, group)` | not a widget: shows a step group only when an earlier `YesNo` has a value | Revise **Y** → its Lens; Position flat everywhere → Time-Skip Ladder |
| `Dice(table)` | a "Roll" button, the rolled row with its meaning, "roll again" | Scene Dials, Ideation lane |
| `Slots(source)` | one row per item from a project list (the seven turns, the speakers in a scene) with a field each | Position (turns), GSU (speakers), Belief cuts |
| `Stop(text)` | the end-of-step panel | every stop-callout |
| `Loop(group, carry)` | repeats a group; `carry` maps one step's answer into the next round's first step, *read live, never copied* (SLATE-PLAN Phase 2's WOAC rule) | WOAC, Extend Backward links, Bridge rounds |

`ExerciseEngine` (pure, in `Services/Tools/`) owns: which step is current,
what the trail shows, loop expansion, carry, branch visibility. Every rule in
that table is a unit test in `Passage.Tests` (rule 9: test the carry reads
live, a branch hides when its Y is cleared, a stop point doesn't block a
later step that already has text, etc.).

### 4.4 Anchors (inputs) and targets (outputs)

`ProjectContext` (pure, built from `_analysis` + `_slate` on demand) gives the
pickers their lists:

| Anchor | Source | Picker shows |
| --- | --- | --- |
| Character | `_analysis.Characters` + `Slate.Characters` colours | name chips in Cast colour, most lines first |
| Scene | Board cards of kind Scene (`BoardLanes`), `[[id:…]]` | heading + synopsis line; default = scene under caret |
| Bracket | `_analysis.Brackets` (ranked, ignore list) | the bracket text and its line |
| Inbox beat | `Slate.Beats` where `Page == 0` | beat title |
| Two points | two Board cards, or a card and the end of the Shape | two scene pickers |
| Shape | `SplitScript.Turns` read from the script | the seven turns with their current text |

Per scene, `StoryFilter.SceneSpeakers(_analysis.Elements)` gives who speaks —
GSU rows and Understudy's "least important character" come from it, not
from typing.

Targets, all through `IToolHost`:

| Target | Mechanism (existing code to reuse) |
| --- | --- |
| **Fill the bracket** | `passage.replaceInLine` by text (Fill's `AcceptFillAsync`) |
| **Synopsis under the anchored scene** (or under caret) | `SynopsisPlacement.Find` + `passage.insertLinesAt` |
| **New Beat in the inbox**, tagged with the run's character / storyline | add a `Beat` with `Page = 0` (the tray's `AddBeat(0)` path) + `StoryItems.ToggleTag` for character/storyline |
| **New scene card after the anchor** | `BeatBoardText.BuildCardLines("Scene", …)` + `BeatInbox.TryPlan` placement + `insertLinesAt` |
| **Append to a note page** (character's note, or "Workbench") | `Story.Notes`; create from the built-in *Character* template (`NoteTemplates`, id `…a001`) when the character has no page |
| **The Shape** | `SplitScript.BuildLines` / the existing `WriteShapeAsync` update path, via an adapter from answers to `SplitRun` |
| **Copy** | clipboard, for anything else |

Every target reports in the status line what it did or why it didn't
("INT. KITCHEN is no longer in the script — nothing inserted").

### 4.5 The deck

Existing tools ported, Phase 8 leftovers folded in where they belong, two
new ones from `engines.md` that need a project to be useful. Questions come
from the vault worksheets (`/home/arosa/Sync/Obsidian Vault/Story/Slate/
Worksheets/`) — read the section named, per R-B.

| Exercise | Shelf | Anchor | Steps (sketch) | Lands in |
| --- | --- | --- | --- | --- |
| **Ideation burst** | Nothing | none | `Dice(lane)` → loop: dealt input (editable) → `Lines` output, timed → keep one line | Inbox beat (untagged) or copy. Stays full-screen. |
| **Bridge** | Nothing | Two points | loop: `Lines(3)` timed → `Pick` → the wire read forward | Inbox beats, or scene cards between A and Z |
| **A→Z→Split** | An ending | Shape | ends (prefilled from the script) → *optional end checks* → `Stop` → `Lines` candidates → `Pick` midpoint → stress `AliveFlat` → `Stop` → PP1/PP2 → `Stop` → the other four; *Mirror Z* as an optional branch | The Shape (write / update) |
| **Belief Split** | An ending | Character | the five cuts as `Slots` over the turns, third-rail at the pinches; *Cast card* for opponent/supporting as an optional group | Character note; turns' synopses |
| **Extend Backward** | An ending | Shape end or a scene | Z (or *Obituary Z* branch) → loop: "what had to be true right before this?" with mystery-lens toggle | Inbox beats in order, or scene cards before Z |
| **Position** | A moment | Inbox beat or scene | `Slots(turns)`: before / after / alive-flat per turn → `Branch` flat everywhere → *Time-Skip Ladder* → shortlist | Moves the beat to the Board under the chosen turn's sequence (the existing Send-to-Board path) |
| **Fill** | A moment | Bracket | the bracket → `Lines(3)` with "why it's probably wrong" → what the wrongness points to → the real answer (escape hatch first: "if it showed up already, write it here") → optional *Stage it* group | The bracket; then *Next bracket* |
| **Scene Dials** *(new standalone; was Fill Phase 8 item)* | A scene going | Scene | `Slots(speakers)`: wants out of *this* scene → `Dice(dials)` → fragment | Synopsis under the scene, or copy |
| **WOAC chain** | A scene going | Character (+ scene) | loop: Want (carried) → Obstacle → Action → Consequence → read it back | Inbox beats per round, or synopsis |
| **Flaw Brainstorm** | A scene going | Character | the eight questions | Character note; inbox beats |
| **Push/Pull** | A dead scene | Scene | `Slots(speakers)` GSU → six `YesNo` checks, each **Y** `Branch`es into its Lens + fragment → read-back (incl. the two new Phase 8 questions) | Revision note on the scene; fragment to copy |
| **Understudy** *(new)* | A scene going | Scene | picks the speaker with fewest lines in the scene (changeable) → what they wanted walking in → five minutes before → what they want that has nothing to do with your plot → freewrite, timed | Inbox beats tagged with that character; optional new storyline |

Not in the deck: Casting Against Type (needs the vault's script library),
the five-layer build and structure remix (undecided in `engines.md`). Same
"Not going in" list as SLATE-PLAN otherwise.

### 4.6 Entry points outside the dock

- **Board card**: a "Tools" item in the card's existing action row (next to
  ✎) → menu of scene-anchored exercises.
- **Cast row** (CHARACTER panel): "Tools" → character-anchored exercises.
- **Inbox beat** (BEATS tray and BOARD strip): "Position this".
- **Bracket under the caret**: covered by *Right now*; no editor widget
  (keeps CodeMirror untouched).

Each opens the dock (if closed) straight into a new run with the anchor set.

---

## 5. Code layout

New, all in `Passage/Passage.Web/`:

| Path | What |
| --- | --- |
| `Services/Tools/Exercise.cs` | `Exercise`, `Step`, `StepKind`, `Shelf`, `AnchorKind`, `TargetKind` records |
| `Services/Tools/ExerciseCatalog.cs` | the deck: one static definition per exercise |
| `Services/Tools/ExerciseRun.cs` | `ExerciseRun { Id, ExerciseId, Anchor, Answers (Dictionary<string,string>), Rounds, Sent }`, `Anchor { Kind, TargetId?, Text }` |
| `Services/Tools/ExerciseEngine.cs` | pure: current step, trail, loops, carry, branches |
| `Services/Tools/ProjectContext.cs` | pure: anchor lists from analysis + sidecar |
| `Services/Tools/RightNow.cs` | pure: the ≤3 suggestions |
| `Services/Tools/SlateMigration.cs` | v1 → v2: each old record → `ExerciseRun` |
| `Services/Tools/IToolHost.cs` | the host interface (R-D) |
| `Components/Tools/ToolsDock.razor` | launcher + routing between launcher and runner |
| `Components/Tools/ExerciseRunner.razor` | anchor strip, trail, step card, end panel |
| `Components/Tools/StepCard.razor` | renders one step by kind |
| `Components/Tools/AnchorPicker.razor`, `SendToMenu.razor` | pickers and target menu |
| `Components/Tools/IdeationOverlay.razor` | the full-screen burst, moved out |
| `wwwroot/css/tools.css` (or a new section in `app.css` — match whichever the repo does for components; today it's all `app.css`) | styles |

`Editor.razor` keeps: the `<aside class="workshop">` shell, the TOOLS button,
the `IToolHost` implementation (thin wrappers over methods it already has),
and the entry-point buttons on Board / Cast / inbox. Every old dock view and
handler is deleted as its exercise is ported. Reuse, don't rewrite:
`AliveFlat.razor`, `PartnerLines.razor`, `SplitScript`, `SynopsisPlacement`,
`BracketScanner`, `IdeationDealer`, `BeatInbox`, `StoryItems`, `NoteTemplates`,
`passage.startBurst` / `replaceInLine` / `insertLinesAt` / `scrollToLine`.

Check `Passage.Web.csproj` needs nothing (no new `ProjectReference`, so the
Dockerfile COPY list — CLAUDE.md trap 2 — is untouched). If a row adds a
project reference anyway, it adds the Dockerfile line in the same commit.

---

## 6. Rows

One row per session, one branch and one PR per row, per CLAUDE.md. Every row
must leave the dock fully working: tools not yet ported stay reachable from
the new launcher (a "Classic" link on their card that opens the old view)
until their row deletes the old code.

### T1 — Engine, model, migration scaffold · `done`

- `Exercise.cs`, `ExerciseRun.cs`, `ExerciseEngine.cs` with every step kind
  in §4.3 implemented in the engine (rendering comes later).
- `SlateDocument.Exercises`; `SlateMigration` skeleton that moves **Chains**
  (WOAC + Flaw) only, leaves the rest untouched, and is idempotent by
  construction (old list cleared). Not yet called from `Load`; version stays 1.
  Old sidecars still load; a sidecar carrying `Exercises` still loads.
- Tests: engine (carry read live; loop appends a round when its last step is
  answered; branch shows/hides; stop point; skip), migration round-trip of a
  WOAC chain with three rounds and a Flaw run, version bump.
- **No UI change.** Success: build clean, tests pass, existing dock untouched.

**Done:** `Services/Tools/{Exercise,ExerciseRun,ExerciseEngine,SlateMigration}.cs`.
Engine: `Sequence` (loop expansion, branch visibility), `Current`, `Trail`,
`GetAnswer` (carry read live, never stored; `SetAnswer` refuses a carried
slot), loop round appended when the group's last step is answered in the last
round, `Skip`/`Continue`, `PickOptions`, `Roll`, `SetLine`/`SetItem`. Answer
keys: round r>0 `key#r`, Lines slot `key[i]`, Slots item `key/item`. Migrated
WOAC keys `want/obstacle/action/consequence/readback`; Flaw `q0..q7`,
`readback`. `SlateMigration.Migrate` moves Chains and clears the old list;
`SlateStore.Load` does not call it yet and `CurrentVersion` stays 1.
7 new tests (117 total, all pass).

**Deviations:**
- Migration is a move per type, not a version-gated copy (review of PR #61):
  gating on version would have blocked T4–T7 and lost edits made to Chains
  between T1 and T2. No `slateVersion` bump until T8.
- Skip and Stop-continue share one `Passed` list; one loop group per exercise;
  Branch inside a Loop group is not supported (nothing in the deck needs it).
- Branch is a marker step (`BranchGroup`, `OnKey`, `OnValue`, `AllSlots`) with
  the group's steps tagged `Group`; `AllSlots` covers Position's "flat
  everywhere".

### T2 — Launcher, runner shell, WOAC and Flaw ported · `done`

- `IToolHost` + Editor's implementation; `ToolsDock`, `ExerciseRunner`,
  `StepCard` for `Text`, `Lines`, `Loop`, `Stop`; trail; Done for now; *Why?*
  disclosure; timer via `startBurst` on the step card root.
- Launcher per §4.1 **without** *Right now* (T3). Cards for all twelve; the
  ten unported open their old view ("Classic").
- **Wire `SlateMigration.Migrate` into `SlateStore.Load`** (this row deletes
  the Chain views, so Chains move here). Each later row adds its own type to
  `Migrate` in the row that deletes that view.
- WOAC and Flaw defined in the catalog and running in the new runner. Anchor
  is free text for now. Old Chains / Chain views and handlers deleted.
- Success: a WOAC chain of three rounds by keyboard only (Enter between
  steps), carry visible, close and reopen keeps it, reload keeps it; phone
  width readable; verification checklist (§8) passes.

**Done:** `Components/Tools/{ToolsDock,ExerciseRunner,StepCard,TrailLine}.razor`,
`Services/Tools/{IToolHost,ExerciseCatalog}.cs`; Editor implements `IToolHost`
through a nested `EditorToolHost`. `SlateStore.Load` calls
`SlateMigration.Migrate` (in memory; the file is rewritten on the next save).
WOAC and Flaw are catalog entries; old Chains / Chain views, handlers and the
chain partner wiring are deleted. `Exercise` gained `Why`, `AnchorLabel`;
`Step` gained `Label`, `Send`; engine gained `IsCarried`, `LineCount`;
`Suggestions.ForStep` serves any `Offer` step (null partner renders nothing).
Runner: anchor is a free-text strip, trail with older rounds folded under
"Earlier", carried wants shown as "↓ carried", Enter = answer + next (client-side
in `passage.js`, no per-keystroke traffic), Enter or Next on a blank step skips,
a "Read it back" button reaches the closing step, "Put it in the script"
inserts a synopsis (the old promotion, until T3's Send to menu). Timer: ⏱ toggle
with seconds and Go/Stop on the step wrapper, which survives step changes.
Launcher: five-shelf chip row (remembered in `passage.session.v1`), cards, kept
runs under their card by anchor text (first answer if unnamed) with a remove ×.
"Why?" open state is one per-browser flag. Dock default width 340 → 400.
Tests: 2 new (119 total, all pass); the Load/Chain round-trip test was updated
for migration-on-load.

**Deviations:**
- The launcher has ten cards, not twelve, and eight "Classic" ones, not ten:
  Scene Dials and Understudy don't exist until T4/T7, and a card that opens
  nothing would fail visibly the wrong way. They get cards in their rows.
- `Lines` and the other kinds beyond `Text`/`Stop`/`Loop` render, but nothing
  in the catalog uses `Lines` yet, so it was not exercised in the browser;
  `Pick`, `YesNo`, `AliveFlat`, `Branch`, `Dice`, `Slots` show "isn't in the
  runner yet" until their rows.
- Added a minimal "Put it in the script" on the read-back so the Chain view's
  one output path survives T2; T3 replaces it with the Send to menu.
- Verification step 5 (anchor deleted underneath) doesn't apply yet: the anchor
  is free text until T3. Step 4 undo checked via the status-bar Undo only.
- Screenshots are in `docs/screenshots/slate-redo-t2/` (the browser pane
  captures at reduced size; Playwright's Chrome isn't installed here).

### T3 — Anchors, targets, Right now · `done`

- `ProjectContext`, `AnchorPicker`, anchor strip (with "no longer in the
  script"), `SendToMenu` with: inbox beat (tagged), synopsis under anchor,
  note page append (Character template), copy.
- WOAC anchored to a character (+ optional scene); Flaw anchored to a
  character, sends to the character's note.
- `RightNow.Suggest` with tests; *Right now* row on the launcher.
- Entry points: Cast row "Tools", Board card "Tools" (menus list only the
  exercises ported so far).
- Success: from the Cast panel, start Flaw on a character, answer two
  questions, send to note: the character's note page exists with both
  answers; a WOAC consequence sent to the inbox shows on BEATS and BOARD
  strip tagged with the character; undo of any script write restores exactly.

**Done:** `Services/Tools/{ProjectContext,ToolTargets,RightNow}.cs`,
`Components/Tools/{AnchorPicker,SendToMenu}.razor`. `Anchor` gained `SceneId` /
`SceneText` (a scene by id where its heading carries `[[id:…]]`, otherwise by
unique heading text; ambiguous or missing resolves to nothing and the status
line says so). `IToolHost` gained `Cast`, `Scenes`, `SceneUnderCaret`,
`SpeakersOf`, `ResolveScene`, `RightNow()`, `SendAsync(target, ToolSend)`,
`SynopsisTargetLabel(anchor)`, `JumpToAsync`; `PutSynopsisAsync` and
`Step.Send` are gone. The runner's anchor is a strip (colour dot, scene,
"not speaking in the script" / "no longer in the script", ↗, change) over a
picker (Cast chips in colour, typed name, scene select for WOAC, "use the scene
I'm in"). Send to ▾ sits under any Text/Lines step and offers the exercise's
`Targets`: Beat inbox (tagged with the character, title = first line, the rest
in the description), Synopsis (under the anchored scene, or the caret's when the
run names none), the character's note page (made from the Character template if
absent; answers go under a "## <exercise>" heading as `- label: answer`), Copy.
Sent marks show on the trail ("→ note, Inbox"). Right now sits above the chips
(≤ 3, recomputed on each open, rotating the scene suggestion). Cast rows and
Board scene cards have a Tools menu listing WOAC and Flaw; it opens the dock
into a new anchored run. Tests: 4 new (123 total, all pass).

**Deviations:**
- An answered Text step now holds the card until Next (Enter still does answer +
  Next in one press), because otherwise the step moved on the instant it was
  answered and Send to was unreachable on it. Timed rounds still advance by
  themselves. `passage.js`'s Enter now always taps Next.
- Send to is under the answer, not in the action row, and acts on one step's
  answer; there is no "send the whole run" yet.
- `RightNow` offers WOAC (on the scene's busiest speaker, keeping the scene) in
  the caret-in-a-scene rotation: with Push/Pull, Scene Dials and Understudy not
  yet ported that rule would otherwise never fire. Bracket-under-caret is "a
  live bracket on the caret's line" (the editor reports a line, not a column).
  Fill, Position and A→Z→Split rules are written and tested but stay quiet until
  their exercises exist (suggestions only offer ids in the catalog).
- Flaw from a Board card takes the scene's busiest speaker as its character and
  does not keep the scene; WOAC keeps both.
- `SetStatus` now re-renders the page: the dock's own events don't re-render the
  status bar, so T2's status lines were not showing. `TargetKind` is stored by
  name in the sidecar.
- Verified in the browser on a scratch copy (§8 steps 1–7 at desktop and 375px,
  keyboard-only WOAC round, undo restores the script exactly, anchored scene
  deleted underneath → status line, nothing inserted); the screenshot pane was
  unreliable at phone width, so 375px was checked by layout metrics (no
  horizontal overflow) and a partial screenshot.

### T4 — Fill and Scene Dials · `done`

- Fill as an exercise anchored to a bracket; *Next bracket* keeps the
  "hand me one" cursor semantics (`_handCursor`, ignore list); "Not a
  bracket" still adds to the ignore list. Optional *Stage it* group.
- Scene Dials as its own exercise (`Slots(speakers)`, `Dice`).
- Migration: Fills. Delete old Brackets / Fill views (keep the bracket list
  as a picker inside the anchor picker).
- Success: SLATE-PLAN Phase 1's success line, in the new runner: caret on a
  bracket → *Right now* → Fill → accept → script edited, undo intact, next
  bracket handed.

**Done:** `ExerciseCatalog.Fill` and `.SceneDials` (the dial steps are one shared
builder, so Fill's "Stage it" and the standalone exercise ask the same three
questions). `StepCard` renders `Slots` (a field per speaker in the run's scene,
via `Host.SpeakersFor`) and `Dice` (Roll / Roll again, die face, "Name — what you
do"). Engine: `SlotItems`, `Roll` never repeats the row it holds, and
`SkipUnansweredBefore` for the escape hatch. Fill: Lines(3) candidates, Lines(3)
reasons, points-to, the real answer, then a `Stop` and the Stage it steps. A
"The real answer" button (`Exercise.ShortcutKey`) jumps there from any step and
passes the unanswered questions before it. A bracket run's strip carries
**Fill it** (ranged `passage.replaceInLine`, run dropped, next bracket handed,
"— that was the last bracket" when none is left), **Not a bracket** (ignore list,
hands on the next) and **Next bracket** (`_handCursor++`, hands on the next). The
launcher's Fill card hands one bracket and Scene Dials starts on the scene under
the caret; the picker lists every bracket (left-out ones with ↩). Closing a
bracket's run puts it back, so Fill from the launcher offers a different one.
`SlateMigration.Migrate` moves Fills (blank ones are cleared, not moved; candidate
and reason rows become `cands[i]` / `why[i]`). Old Brackets / Fill views,
handlers, `FillPartnerKey` and the `Fills` clean-up in `SaveSlate` are deleted.
Right now offers Fill on a bracket on the caret's line and Scene Dials in the
scene rotation; the Board card Tools menu lists the Scene-anchored exercises. 4
new tests (127 total, all pass).

**Deviations:**
- Fill's targets are *Fill it* in the strip (not a Send to entry), so it works
  from any step, including after Stage it; Send to on Fill's steps offers
  Synopsis (under the scene the bracket sits in) and Copy.
- A bracket run with nothing written, and a scene run with nothing written, are
  `IsEmpty` however the anchor is set (`ExerciseRun.IsEmpty`), so handing out
  brackets leaves no kept-run chip behind. A typed character still keeps one.
- The answer is joined onto one line before it replaces the bracket (a bracket
  sits in one line; the old view passed newlines through).
- Fill is not in the Board card menu, and Cast menus now list character-anchored
  exercises only: Fill is bracket-anchored and §4.6 gives scene cards
  scene-anchored tools. Right now covers a bracket under the caret.
- The old `SceneDials` table in `Editor.razor` stays: Push/Pull's check 0 still
  points at it until T7 deletes it. `Suggestions.ForFill` and `FillRun` stay for
  T8's cleanup (migration reads `FillRun`).
- The dev data holds no saved Fill runs (the old view pruned empty ones), so the
  migration was checked on a seeded scratch copy (`[gate or door?]` with two
  candidates and a points-to): it shows as a kept run on the Fill card, anchored
  by its bracket, with its rows in place.
- Verified in the browser on that scratch copy (§8): Right now → Fill on the
  caret's bracket; candidates, reasons, real answer; Fill it edits the line, the
  status-bar Undo restores it exactly with the caret on that line; the next
  bracket is handed; Not a bracket; Next bracket then Fill from the launcher offers
  a different one; the bracket edited away underneath → "no longer in the
  script — nothing changed", run kept; Stage it from the Stop uses the bracket's
  scene speakers; Scene Dials on the caret's scene (MARA prefilled, dial rolled,
  sketch sent to Synopsis, undo exact); Board card Tools → Scene Dials; no
  counters or dates in the dock text; `.partner` count 0; 375px has no horizontal
  overflow. Keyboard-only was exercised with the real Enter key between Lines
  slots and steps; keystrokes sent faster than the round trip can race the
  re-render, as before.

### T5 — Split family · `done`

- A→Z→Split (with optional end checks and Mirror Z), Belief Split (with Cast
  card group), Extend Backward (with Obituary Z and mystery lens).
- Adapter answers → `SplitRun` for `SplitScript`; Write / Update the Shape
  unchanged in behaviour.
- Migration: Split, Belief, Extend. Delete old views.
- Success: a fresh script, Split to rung 1, write the Shape: Board shows the
  four acts / eight sequences with the midpoint; update after rung 2 fills
  only bracketed turns.

**Done:** `ExerciseCatalog.Split`, `.Belief` and `.Extend`; `Services/Tools/ShapeAnswers.cs`
is the adapter (`ToSplitRun`, `ToBeliefRun`, `SplitTurn`, `Prefill`), so the Shape
write is still `SplitScript.BuildLines` / the ranged fill-in of bracketed turns,
now `Editor.WriteShapeAsync(SplitRun)` behind `IToolHost.WriteShapeAsync`. New step
rendering in `StepCard`: `Pick` (tap a candidate, or a fixed "Name — gloss" list, or
write it), `YesNo` and `AliveFlat` (the `AliveFlat` pills), and `Branch` groups
through the engine as before. Engine additions, each tested: fixed `Options` on a
`Pick`, `Pick` / `IsPicked`, a per-round second wording (`AltQuestion`, switched by a
toggle, stored as `key~alt`, never an answer) for Extend's mystery lens, `Echo` (what
the link is "right before": Z, then the previous link, read live) and `ReadForward`
(the chain, last link first, Z last). Split: A, Z (prefilled from the script's
`= Starts:` / `= Ends:` when the run is fresh), optional end checks (irony, worst-case
pitch) and Mirror Z behind yes/no gates, candidates → Pick → stress, the optional
belief layer, then the three stop-callouts in the worksheet's words. Belief: shape,
A/Z, origin with size and the gut check (resize on *n*), the five cuts each with its
matched plot turn from the Split run, alive/flat per cut, then the Cast card as a loop
behind its own stop. Extend: Z, optional Obituary Z, a loop of links with the mystery
toggle and "Right before" line, "Read it forward" → alive/flat → weakest link. The
Split strip has **Write the shape into the script / Update the shape in the script**
(any rung). Launcher: Split and Extend pick up the run already kept for the Shape
rather than start a second; Right now's A→Z→Split rule is live. The Board card Tools
menu lists Extend with the scene named. `SlateMigration.Migrate` moves Split, Belief
and Extend (blank ones are cleared, not moved); old views, handlers, `WorkshopView`
members and the Split/Extend partner keys are deleted. 6 new tests, 1 updated (134
total, all pass).

**Deviations:**
- Shape targets: the write is a button in the Split strip (like Fill's *Fill it*), not a
  Send to entry; Split's Send to offers Copy only. Belief sends to the character's note
  page and Copy, not "turns' synopses" (the old tool never wrote cuts into the script, and
  a turn's synopsis is the script's once promoted — SLATE-PLAN decision B). Extend sends to
  inbox beats, one link at a time; "scene cards before Z" is not built (`SceneCard` has no
  host target yet).
- The Split belief layer keeps working as an optional branch inside Split. It and Belief
  Split used to share one record of five cuts; they are two exercises now, so on migration
  the cuts are copied into the Split run when the layer was on. A migrated run is not
  asked the new optional gates again, and stops above text already written are passed.
- The optional gates (end checks, Mirror Z, belief layer) are yes/no steps that Enter
  skips; the Cast card sits after Belief's cuts behind a stop (the worksheet puts it
  after the origin) because a loop must be last in the exercise. Its two opponent-only
  questions are skippable steps, and the last question ("why we'd watch") is the one
  that opens the next character.
- Turns, A and Z are single-line fields (Enter moves on); only the opposite ending and the
  eulogy are multiline, where Ctrl+Enter moves on.
- Belief's anchor is a character (one run per character); the old tool had one per script, so
  a migrated Belief run has no character until one is picked in the strip.
- Left for T6 / T7 / T8: `Position`'s `ScriptTurnHint`, Push/Pull's `BeliefShapes`, and
  `Suggestions.ForSplitMidpoint` / `ForExtendLink` and the `SplitRun` / `BeliefRun` /
  `ExtendRun` types (migration reads them) stay until those rows or T8's cleanup.
- Verified in the browser on a scratch copy of the dev data (§8): `split-test`'s Split and
  Extend runs show on their cards after migration (its Belief was empty, so none); on a fresh
  script Right now → A→Z→Split, keyboard through A, Z, the end checks gate (yes, one answer,
  Enter-skip the rest), Mirror Z skipped, three candidates, Pick, stress, belief layer skipped,
  stop-callout; *Write the shape* put in four acts and eight sequences with the midpoint and
  the Board showed them; status-bar Undo restored the empty script exactly and Redo the
  structure; Update after PP1 / PP2 filled only those two bracketed lines and left the
  midpoint alone, and two Undos restored it with the caret on the line; a line removed
  underneath → "no line in the script for Plot point 1", no exception. Extend's Z was
  prefilled from `= Ends:`, the echo and mystery toggle behaved, Read it forward listed the
  chain. Belief showed the matched plot turn from the Split run. 375px has no horizontal
  overflow; no counters or dates; `.partner` count 0; no server errors. Not exercised by
  hand: the Cast card's Send to, and a Belief run started from a Cast row.

### T6 — Bridge and Position · `done`

- Bridge anchored to two Board cards; Position anchored to an inbox beat or a
  scene, `Slots(turns)`, Time-Skip Ladder branch.
- Position's target: send the beat to the Board under the chosen turn's
  sequence (reuse the Send-to-Board path; same confirm wording about the page
  plan being dropped). Inbox "Position this" entry point.
- Migration: Bridges, Positions (prune trailing empty turns — handover open
  observation, now moot in the generic model). Delete old views.

**Done:** `ExerciseCatalog.Bridge` and `.Position`; `Services/Tools/AnchorPrefill.cs` starts the blank answers an
anchor can supply (Bridge's A and Z from the two cards' heading and synopsis, Position's moment from the beat or
scene). Bridge is anchored to two scenes (`Anchor` gained `EndSceneId` / `EndSceneText` for Z; A is the existing
`SceneId` / `SceneText`; `ProjectContext.ResolveEnd`): A, Z, then a loop of `Lines(3)` → `Pick` (a pick opens the
next round, which the writer can leave), then the wire read forward (new `Step.LeadKey`: A, the picks in order, Z).
It sends a link to the inbox, Copy, or a new scene card (`TargetKind.SceneCard`, built from `BeatInbox.SceneLines`,
placed with `BeatInbox.TryPlan` just before Z, or just after A when Z isn't named, by `passage.insertLinesAt`, so
undo restores it exactly). Position is `Slots(turns)` three times (before, after, alive / flat pills — new
`Step.Pills`), the script's own turn text under each slot, the Time-Skip Ladder as a `Branch` that opens only when
all seven turns are flat, a two-line shortlist and a `Pick` of the turn over the alive ones (all seven when none is
marked). Its target `TargetKind.Board` files the inbox beat as a scene under the sequence whose synopsis carries the
chosen turn's line (`ToolTargets.SequenceForTurn`), through `ConvertBeatToCardAsync`, the Send-to-Board path; a beat
still on the page plan asks first with the same modal, which now says "under <turn>". "Position this" is a button on
every inbox beat (BEATS tray and BOARD strip share `RenderBeat`) and Right now's Position rule is live; both resume
the run kept for that beat. `ScriptTurnHint` is folded into the host's `ScriptShapeValue`. The Bridge card is in the
Board card Tools menu (that scene as A). `SlateMigration.Migrate` moves Bridges and Positions (blank ones are cleared,
not moved; trailing empty rounds and turns are not carried; a turn with text but no slot named keeps its text on the
moment). Old views, handlers, `WorkshopView` members, `BridgeSlotsId`, and (now unused) the Editor's whole partner
block and `IStoryPartner` injection are deleted. Tests: 4 new, 2 updated for load-time migration (138 total, all pass).

**Deviations:**
- Bridge's cards are scenes (a Board scene card), not sequence or act cards; Z can't be "the end of the Shape" (type
  it in the step). A migrated Bridge has no cards, so its ends are text and its "scene card" target says "Name the
  two cards first" until they are picked.
- `ConvertBeatToCardAsync` now mirrors its insert into `_content` and re-analyses (as `PutSynopsisAsync` already did):
  a second write in a row would otherwise plan from stale text. Undo of that insert restores the script exactly; as
  with the tray's Send to Board, the beat stays out of the inbox after an undo.
- Position's Board target is offered only on a run about an inbox beat; a run about a scene, or a migrated one, has
  Copy only (moving a card that is already in the script is a drag, not a send). "Flat everywhere" is the engine's
  all-seven rule, so the Ladder opens only after every turn is marked flat.
- Position's three questions are three `Slots` steps over the seven turns, not one row per turn: before, then after,
  then the marks. A turn left blank is a turn not tried.
- Beats on the page plan (Page > 0) can't be the anchor from "Position this" (inbox beats only), so the confirm
  modal is reached only if a beat is moved onto the plan after its run began. It was checked by reading, not by hand.
- Verified in the browser on a scratch copy of the dev data (§8): `split-test`'s Bridge and Position runs show on
  their cards after migration; "Position this" on a BOARD-strip beat, Right now → Position, moment prefilled, a
  turn marked alive, the pick, Send to → under Plot point 1 / Pinch 1 (a scene under Sequence 2 / 3, beat out of the
  inbox), status-bar Undo restores the script exactly; all seven turns flat opens the Ladder; Bridge from the card
  pickers (A and Z prefilled), three links, pick, Beat inbox (shows on the BOARD strip), scene card before Z, Undo
  exact; step 5 for both (Z's heading renamed, then a sequence's turn line renamed: status line, script untouched,
  no exception); no counters or dates; `.partner` 0; 375px has no horizontal overflow; no server errors. Also built
  once with SDK 9.0.317 pinned (the Razor trap). Not exercised by hand: the page-plan confirm, and a Bridge run on
  sequence or act cards.

### T7 — Push/Pull and Understudy · `todo`

- Push/Pull anchored to a scene: GSU `Slots(speakers)` prefilled; six checks;
  each **Y** branches into its Lens with the `LensMinutes` clock; read-back
  with the two Phase 8 questions. Revision note appended to a note page named
  after the scene.
- Understudy (new): default character = fewest lines among the scene's
  speakers; sends tagged inbox beats.
- Migration: Revisions. Delete old views.

### T8 — Ideation overlay, cleanup, docs · `todo`

- Move the overlay to `IdeationOverlay.razor`, restyle into the step-card
  look; keep it full-screen and project-free (vault `writing-practice` rule).
  Its kept line can go to the inbox.
- Delete the old record types from `SlateStore.cs` once nothing reads them
  (migration keeps reading v1 JSON via private DTOs inside
  `SlateMigration`). Remove "Classic" links.
- Update `SLATE-HANDOVER.md` to describe the new structure; mark SLATE-PLAN
  Phase 8 `superseded` with a pointer to the rows here that absorbed each
  item.

### T9 — A real partner · `blocked on Adam`

Not part of this plan. Listed because it is the biggest remaining lever: the
worksheets were designed to be run with someone offering three wrong answers.
The seam already exists (`IStoryPartner`); R-C makes every `Lines` step a
partner site for free (`Offer = true`). Whether, which model, where the key
lives and what text leaves the machine are Adam's decisions, and would get
their own plan.

---

## 7. Decisions — settled 2026-10-25 (Adam accepted all four recommendations)

1. **Copy rule (R-B) — settled: fold.** Explanations go behind *Why?*; only
   the question and one hint stay on screen. This supersedes SLATE-PLAN
   decision A's near-verbatim copy rule for on-screen text; the question
   itself still keeps the worksheet's words.
2. **Old runs (R-E) — settled: migrate.** Existing sidecar runs are
   moved into `ExerciseRun`s (per type, in the row that deletes its view), not dropped.
3. **New exercises — settled: build both** Scene Dials (standalone) and
   Understudy.
4. **Dock width — settled: 400px default.**

---

## 8. Verification (every row, by hand, in the browser)

Use your own server with a scratch copy of the dev data (see
`SLATE-HANDOVER.md` → Environment: `~/.dotnet/dotnet`, never Adam's server on
5210, a `launch.json` entry with `--Passage:DataDir=<scratchpad>/data`).

1. Desktop width and 375px: launcher readable without scrolling past
   descriptions; runner shows one question at a time.
2. Type, ←, reopen: kept. Reload, reopen the script: kept (sidecar).
3. Keyboard only through three steps (real Enter key, `key: "Enter"` in the
   agent browser).
4. Every script write: undo restores exactly, caret doesn't jump to line 1
   (test undo as the handover's *Traps* section says — synthetic Ctrl+Z
   doesn't reach CM5).
5. Delete the anchor underneath an open run, then send: status line, no
   exception, no wrong edit.
6. Rule 0: no counter, percentage, date, streak or "done" mark anywhere.
7. `document.querySelectorAll('.partner').length === 0` with the null
   partner.
8. `dotnet build Passage.Web.slnf` warning-free; tests pass; CI green.
