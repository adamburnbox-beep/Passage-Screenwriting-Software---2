# Slate in Passage — implementation plan

Bringing the Slate engines and the bracket queue into `Passage.Web`.

**Scope, as decided:** engines and brackets only. The multi-story board,
Seeds, stages, last-touched, next-move notes, briefs, read-back/crit and
script sessions **stay in Obsidian**. Passage gets the tools, not the practice.

**Model access, as decided:** deterministic first, with a seam for a model
later. See Phase 7 for why the seam is built last and not first.

This document follows the shape of `docs/WEB-PARITY.md`: numbered rows, a
status per row, one row per session. `docs/SLATE-SESSIONS.md` holds the
copy-paste prompt per row, mirroring `WEB-SESSIONS.md`.

Source material lives in the vault at `Story/Slate/`: `Slate.md`,
`How the Slate Works.md`, `engines.md`, `loosening-up-practice.md`, and the
five live files in `Worksheets/` (the numbered Short/Medium/Long files are
stubs). **The worksheets are the spec for each runner's text.** Labels,
hints and stop-callouts come from them verbatim or near it — the 2026-09-20
review found the first runner had paraphrased them into something thinner,
and the writer noticed.

**Naming:** in the app this is *Writer's Tools* (topbar button `TOOLS`), not
Slate — it is a set of nudges for getting unstuck, not a system to announce.
`Slate` stays in code identifiers, the sidecar directory and these docs,
where it points back at the vault.

---

## Rule 0 — the constraint that outranks every feature below

Every worksheet in the vault is deliberately dumb. No required fields, no
score, no completion state, no history. The fifth pass on the worksheets
existed *because* the Short/Medium/Long split had quietly re-derived the
Woodshed's time-boxing from a different angle — the system's own stated rule
("a session ends when a notch lands, nothing stops you mid-flow") was being
contradicted by the packaging, not the content.

Software makes that failure mode much easier to fall into. A form wants
required fields. A runner wants a progress bar. A stored run wants a history
list, and a history list is a log, and a log is the Woodshed.

So, binding on every row below:

1. **No required fields.** Every runner is closable at any point, partial
   state kept, no warning, no "are you sure".
2. **No completion percentage, no round counter presented as progress**, no
   streak, no "last run" date shown anywhere.
3. **No history view.** Past runs are retrievable by the script they belong
   to, never as a chronological list of sessions.
4. **A timer that expires never blocks and never scores.** It advances. The
   rule from `loosening-up-practice.md` is literal: the only failure state is
   silence, and speaking after the buzzer still counts.
5. **One thing, not a menu.** Where the vault says Claude hands you one
   bracket, the app hands you one bracket. A list view exists, but the "hand
   me one" path must not route through it.

If a row below conflicts with this section, this section wins.

---

## Architecture decisions

### A. Where the engines live in the UI — a right-hand dock

The dock's top level is `0 — Which Worksheet.md`'s one question, *what have
you got?*, with its five rows and the paths under each. That is how the
writer finds the tool; a flat list of engine names is not.

**Copy rule, binding on every tool.** Write for someone who has never seen
the vault and may never have written anything: each tool says in plain
words what it is, what you do, and what you end up with. No filler lines,
no slogans, nothing that reads as generated — the writer will notice and
call it out (2026-09-20). The worksheets supply the mechanics; the
descriptions have to supply the *why*.

Current shell: `.app-shell > .main-area > aside.sidebar + main.center-pane`,
with SCRIPT / BOARD / PREVIEW as view tabs in the topbar.

Add `aside.workshop` on the right of `.main-area`: collapsible, resizable,
persisted alongside the other per-browser state in `passage.session.v1`.

Why a dock and not a fourth view tab: nine of the ten engines are run
*against* the script. Fill needs the bracket visible. Push/Pull needs the
scene. Position needs the turns. A view tab would hide the thing the tool is
about. The dock also means the engines work identically whether you're in
SCRIPT, BOARD or PREVIEW — no coupling to the view tabs at all.

The one exception is Ideation burst mode, which is deliberately *not* about
your script (`writing-practice`: invented or borrowed material only, never his
own characters). That gets a full-screen overlay with nothing else on it —
Phase 6.

### B. Where Slate data lives — hybrid, leaning hard on derived

Recommended, since you left this open:

| Data | Where | Why |
| --- | --- | --- |
| Brackets | **Derived** from script text on every parse. Zero storage. | They already exist in the text as `[SOMETHING ...]`. Storing them creates a second source of truth that goes stale the moment you edit the line by hand. |
| The shape line / seven turns | **Derived** from the existing section + Beat Board structure. Zero storage. | `DocumentAnalyzer` already produces `BoardLanes`. The shape line is a rendering of data the app has. |
| Worksheet runs (WOAC chains, Split rungs, Position scores, Push/Pull flags, Lens fragments) | **Sidecar JSON**, `/data/.slate/<scriptname>.json` | Not screenplay text. Putting it in the fountain file pollutes exports, page counts and the desktop apps' parse. JSON is diffable, backup-able and deletable without touching a script. |
| Ideation burst output | Sidecar, same file, separate key | Deliberately disposable. Never touches the script. |

The sidecar is a **staging area, not a second home for the story**. Every
runner's output has exactly one promotion path into the script (a filled
bracket, a synopsis line, a section heading), and once promoted the script is
the truth. The sidecar keeps the working-out, not the result.

**Why not a boneyard block in the fountain file:** it survives round-trips
through the desktop apps, which is genuinely attractive, but it puts several
KB of half-finished worksheet text inside a file whose page count and export
you care about, and one hand-edit in Obsidian corrupts it. The round-trip
benefit only matters if you run these engines from the Linux app, and the
Linux app is frozen reference per `CLAUDE.md`.

**Storage mechanics:** a new `SlateStore` singleton beside `ScriptLibrary`,
reusing `ScriptLibrary.TryValidateName` for the key so browser input still
can't escape the root. `/data/.slate/` is a dot-directory, so
`ScriptLibrary.List()`'s extension filter already hides it — no change needed
there. Schema carries `slateVersion: 1` from day one.

**Orphaning is a real cost and must be handled in Phase 0, not later:** the
app's own delete and Save As must cascade to the sidecar, and a sidecar whose
script no longer exists is pruned on load with a status message — the same
pattern `RecentFiles` already uses (WEB-PARITY 1.5). Files can also change on
the volume from outside the app; a sidecar that references a line number that
no longer holds a bracket must degrade to "not found", never throw.

### C. Brackets — scan, don't change the convention

Your brackets are single square brackets: `[SOMETHING forces her to stay]`.
Fountain's note syntax is double: `[[note]]`. Passage's NOTES sidebar already
indexes the double form via `NoteElement`.

Do **not** migrate the convention to `[[ ]]` to ride the existing
infrastructure. The single-bracket form is written into every file in the
vault and into `How the Slate Works.md`'s own definition. Change the app, not
the habit.

So: a new `BracketScanner` in `Passage.Parser`, beside `FountainMarkup.cs`.
Pure string logic, no UI types, fully testable — the same shape as
`BeatBoardText`. It scans Action, Dialogue, Section and Synopsis lines for
`[...]` spans, after `FountainMarkup.StripOmissions` has removed `[[notes]]`
and boneyard so those can't double-match.

**The false-positive problem is real and must not be papered over.** Single
square brackets are legal prose — `[sic]`, a stage direction, a citation. The
scanner will catch them. Mitigations, in order of preference:

1. Rank, don't filter: a bracket whose content is `ALL CAPS` or starts with a
   known keyword (`SOMETHING`, `CONDITION`, `TODO`) sorts to the top of the
   queue. Everything else is still listed, just lower.
2. A per-script ignore list in the sidecar, one click to dismiss.
3. Never auto-anything. The scanner suggests; nothing is rewritten without an
   explicit accept.

### D. Trap compliance

Two existing traps in this codebase apply to almost every row below.

**`CLAUDE.md` trap 4 — Blazor round-trips.** Burst timers must run in
`passage.js`, not in Razor. A 30-second timer that ticks over a WebSocket is
fine until the circuit blips, and this timer is the external-pressure
mechanism the whole loosening-up practice depends on — it has to be exact and
it has to survive a reconnect. Server round-trips are for saving the run.

**`WEB-PARITY` 1.6 — undo.** Every write-back into the script from a runner
must use a ranged replacement. Nothing may call `ReplaceEditorContentAsync`,
which wipes the CodeMirror undo stack and throws the caret to line 1. That
trap has already bitten this repo twice (Save As, and the Beat Board
write-back it was fixed for). Filling a bracket is a one-line replacement —
it should be the easiest case in the app to get right, and it will be the most
frequently used write path in the whole feature.

### E. Burst timers — ready, prompt, write; never a click in between

Binding on every timed round (Chain in Phase 2, Ideation in Phase 6, and any
"three ideas in ten seconds" style burst). The timer exists to supply
external pressure; anything that costs the writer a click or a glance once
it is running defeats it.

1. **Ready.** The round's input is focused *before* anything counts. The
   caret is already where the words go; the writer never has to place it.
2. **Read-in.** The prompt appears with a short countdown (about five
   seconds, configurable) so it can be read. Typing during the read-in is
   allowed and simply starts the round early.
3. **Write.** The real timer runs (ten seconds, ninety, whatever the
   worksheet says). Enter commits a line and moves focus to the next slot.
   No mouse needed for the whole round.
4. **Done.** The round ends when the slots are filled — the timer stops and
   is not shown again — or when the timer expires. Expiry does nothing but
   advance (rule 0.4): the fields stay editable, late text still counts, no
   red, no sound of failure, no count of misses.

All of this runs in `passage.js` (decision D, trap 4): focus, countdown and
advance are client-side, and the server sees only the saved round.

---

## Phases

Each row is one session, one feature, per `CLAUDE.md`'s workflow. Every phase
is shippable alone.

### Phase 0 — Foundations · `done`

Nothing user-visible except an empty dock. Worth its own session so no later
row has to invent storage mid-task.

- `SlateStore` service: read/write `/data/.slate/<name>.json`, `slateVersion: 1`,
  name validation reused from `ScriptLibrary`, cascade on delete and Save As,
  prune orphans on load.
- `aside.workshop` dock in `Editor.razor` + `app.css` grid change + collapse
  toggle, persisted in `passage.session.v1` (extend the existing key, do not
  add a second one — the precedent WEB-PARITY 1.1 and 1.5 both set).
- Empty dock renders a tool list; every tool disabled.
- Tests: `SlateStore` round-trip, invalid-name rejection, orphan pruning.

**Success:** dock opens and closes, survives reload, `dotnet build
Passage.Web.slnf` clean with `TreatWarningsAsErrors`.

- **Done:** `Passage/Passage.Web/Services/SlateStore.cs` — `Load` / `Save` /
  `Exists` / `Delete` / `PruneOrphans` over `<root>/.slate/<validated name>.json`
  (so `draft` and `draft.fountain` are one key). `Load` returns null for a
  missing sidecar and lets a corrupt one throw `JsonException` rather than
  reading it as empty — losing working-out silently would break rule 12.
  `SlateDocument` carries only `SlateVersion` for now; Phase 1 adds the first
  real field.
  Wiring in `Editor.razor`: dock markup at `<aside class="workshop">` (~389),
  `WorkshopTools` list (~627), `ToggleWorkshopAsync` / `PruneSlateOrphans`
  (~2066), delete cascade beside `Library.Delete` (~1217). Pruning runs on the
  first interactive render, not in `OnInitialized` — prerendering runs that
  twice and the second pass would find nothing to prune and lose the message.
  Dock open state rides in `passage.session.v1` as `workshopOpen`; width as
  `workshopWidth`, owned by `passage.js` (`applyWorkshopWidth`,
  `initWorkshopResize`, ~125) via the `--workshop-width` CSS variable, so a
  drag costs no round-trips. Below 900px the dock covers the main area instead
  of squeezing the editor, and the status bar wraps with the status message
  on its own full-width row — the runners' messages are sentences, and the
  one-row bar squeezed them a word wide and grew to a quarter of the screen
  (found in the 2026-09-20 verification pass). CSS under
  `/* ---- Slate workshop dock ---- */`.
  Tests in `Passage/Passage.Tests/Program.cs` (`TestSlateStore*`); the test
  project now references `Passage.Web` for them.
- **Deviation — no Save As cascade.** The web app has no Save As command
  (WEB-PARITY 1.4): the "Save script" dialog only names an untitled buffer,
  which cannot have a sidecar. Delete of the open file already drops both the
  file and its sidecar, so the buffer that becomes Untitled afterwards starts
  clean. If a real Save As is ever added, the cascade (copy the sidecar to the
  new name) goes in the same change.

### Phase 1 — Brackets and Fill · `done` · **use it for two weeks before Phase 2**

The highest-value row in the plan and the one that decides whether the rest is
worth building. See "The honest risk" below.

- `BracketScanner` in `Passage.Parser` + tests (ranking rule from decision C,
  omission-stripping, line numbers, empty/malformed input).
- `DocumentAnalyzer` surfaces `Brackets` alongside `Notes` and `BoardLanes`.
- **BRACKETS panel** in the dock: list, count, click to jump to line — the
  same interaction the NOTES panel already has.
- **"Hand me one"** — a single button that yields one bracket, not a list.
  Deterministic selection: highest-ranked bracket in the current script that
  isn't on the ignore list, round-robin so the same one doesn't repeat.
- **Fill runner** — the Path B worksheet, live: the bracket as it reads, three
  option rows with a "why it's probably wrong" column, the "what the wrongness
  points to" field, and the escape hatch verbatim from the worksheet — *if the
  real answer showed up on its own before you'd finished the options, write it
  here and stop.* No required fields (rule 0).
- **Accept → ranged replacement** of the bracket span. Undo survives.
  Verify this manually; do not assume.
- Chaining: on accept, offer the next bracket immediately. Filling brackets is
  meant to chain "for as long as you want to keep going" — the stop is the
  writer's to call, so there is no built-in stopping point and no prompt
  suggesting one.

**Success:** open a script with brackets in it, press one button, fill one
bracket from your phone, and the script is edited correctly with undo intact.

- **Done:** `Passage/Passage.Parser/BracketScanner.cs` + `FountainMarkup.MaskOmissions`
  (blanks notes/boneyard in place so positions hold). `Bracket.Text` is the
  span exactly as written, brackets and padding included, because that is
  what the replacement has to match. Rank 0 = all caps or opens with
  SOMETHING / CONDITION / TODO; rank 1 = everything else, still listed.
  Tests `TestBracketScanner*` in `Passage.Tests/Program.cs`.
  `DocumentAnalysis.Brackets` is derived on every parse, screenplay and
  markdown alike. Dock views in `Editor.razor`: BRACKETS panel (~445) with
  "Hand me one" above the list, and the FILL runner (~478); handlers under
  `LoadSlate` / `HandMeOneAsync` / `AcceptFillAsync` (~2173–2360). Accept goes
  through `passage.replaceInLine` (passage.js ~540), which matches the span by
  text on the named line and returns false if it is gone — verified by hand:
  undo restores the bracket, redo re-applies, caret stays on the line, and a
  bracket edited away underneath an open runner yields a status line and no
  edit. `SlateDocument` gained `IgnoredBrackets` (by text, not line — drift
  safe) and `Fills` (one `FillRun` per bracket text, saved on every field
  change, so closing keeps everything).
- **Deviations, all deliberate:** every line is scanned rather than only
  Action/Dialogue/Section/Synopsis — a bracketed cue like `[SOMEONE]` parses as
  a character and is still a placeholder, and there is no line type where a
  bracket cannot be one. A run is dropped from the sidecar on accept: the
  filled line is the truth and there is no view that could retrieve it
  (rule 0.3). "Not a bracket" from inside the runner hands the next one, the
  same as accept, so the phone path is one tap per bracket. The hand-me-one
  cursor advances only when a bracket is put back unfilled and resets on any
  queue change, so the top of the queue is always what follows a fill.
- **Untitled buffers:** the runner works but says the working-out is kept only
  once the script has a name; the first Save then writes the sidecar. A
  sidecar that fails to parse is left on disk and saves are refused with a
  status line until it is fixed, rather than overwritten.

### Phase 2 — Forward chain (WOAC + Character Flaw Brainstorm) · `done`

- **Chain runner**, Path A: a round table where each round's Consequence
  auto-populates the next round's Want. That auto-carry *is* the tool — it's
  the thing that removes having to invent the next question, which
  `writing-practice` names as the actual block. Get it exactly right.
- Open-ended: rounds add indefinitely, no cap, no suggested stopping point
  beyond the worksheet's own "5–7 is a pattern, not a rule" note.
- **Path B, Character Flaw Brainstorm**: the fixed eight-question chain. This
  one *is* fixed-length by design — the eight land on a shape together — so the
  UI may show all eight, but must still allow stopping at any one.
- **Burst timer**, in JS, per decision E: ready → read-in → write, input
  focused before anything counts; configurable, default off. Expiry
  advances the round and does nothing else.
- **Promotion path:** "read it back — what's the scene now, in one line?" →
  that line inserts into the script as a synopsis (`=`) under the current
  section, by ranged insert.

- **Done:** `SlateDocument` gained `Chains` (one `ChainRun` per chain, `Path`
  = `Woac` | `Flaw`, stored by name) and `Burst` settings (off by default,
  30 s when turned on, 5–600 s; per script, since it rides in the sidecar).
  A WOAC round stores Want only for round 1: every later Want is the
  previous round's Consequence *read live* in the view, never copied, so
  editing an earlier round moves the carry with it. Writing a consequence
  into the last round appends the next one. The Flaw path stores its eight
  answers; all eight show, none is required. Dock views in `Editor.razor`:
  `WorkshopView.Chains` (the chains kept for this script, listed by seed
  line — no dates, no numbering; empty ones are pruned on save) and
  `WorkshopView.Chain` (one runner, both paths); handlers under
  `// ---- Forward chain`. Opening a tool with nothing kept goes straight
  into a new chain. The burst is `passage.startBurst` / `stopBurst`
  (passage.js, after `focusEditor`): slots are the runner's
  `textarea[data-slot]` in DOM order, starting at the first empty one;
  focus → 5 s read-in (typing cuts it short) → write clock → Enter or expiry
  moves to the next slot; a consequence with no slot after it waits up to
  2 s (MutationObserver) for Blazor to render the next round, then ends.
  Tapping another slot re-targets the run; focus leaving the slots ends it.
  Blazor hears only `OnBurstEnded`. Promotion: `SynopsisPlacement.Find`
  (`Passage.Web/Services`, tested) picks the nearest section / scene /
  markdown heading at or above the caret and the slot after its existing
  `=` lines, or the caret itself when there is no heading — the runner says
  which before the button is pressed. Insert goes through
  `passage.insertLinesAt`; verified by hand: undo removes the line exactly,
  redo restores it, the caret shifts with the text. The chain stays in the
  sidecar after promotion; the list's `×` removes it. Tests
  `TestSlateStoreChainRoundTrip`, `TestSynopsisPlacement`.
- **Deviations, all deliberate:** the seed line ("Character + starting
  want" / "Character + the flaw") is not a timed slot — it is the writer's
  own starting point, not a question the chain asks. The promote button is
  never disabled: the field's change event and the tap arrive together, and
  a button that is disabled until the change lands swallows the tap. A
  blank line gets a status message instead. Fill's "Fill it" had the same
  problem and got the same fix in this phase. The read-in length is
  a constant (`BurstReadInSeconds`), not a setting: one number to set is
  enough, and 5 s read fine in practice.

### Phase 3 — The Split family · `done`

A→Z→Split, Belief Split and Extend Backward share one runner with three paths,
matching `Generate — Split, Belief, or Extend Backward.md`.

- **Rung ladder UI** with the worksheet's stop-callouts rendered as actual UI
  text, not decoration: *"Stop here and you've got a real shape — nothing
  below is owed."* The callouts are the anti-Woodshed mechanism; they are
  content, not filler.
- **A→Z→Split writes into the script structure.** This is the single best
  integration available in the whole plan: the seven turns the tool produces
  are exactly the Acts/Sequences/Scenes the Beat Board already renders.
  Completing rung 3 emits `#`/`##` sections with synopses, which means the
  Beat Board populates itself from a Split run. Split becomes generative of the
  app's existing central view rather than a parallel artifact.
- **Shape line** in the status bar: `▫●▫●▫–▫●▫–▫●▫●▫`, derived from
  `BoardLanes` + which turns have content. Fifteen slots, the three glyph
  states, fill in split order. Purely derived, zero storage. Deliberately not
  a progress bar — the vault is explicit about why.
- **Belief Split**'s matched-cut layer attaches its 12/25/50/75/88 ratios to
  whichever turns Split has produced, and offers the lighter Third-rail
  question at the two pinches rather than manufacturing a sixth and seventh
  named cut.
- **Extend Backward**: one link at a time, backward only. The "no jumping
  ahead to fill the middle from the front" rule is enforceable in UI — only
  the most recent link generates the next field. Mystery-lens variant as a
  toggle on the next link.

- **Done:** one run of each per script — `SlateDocument.Split`, `Belief`,
  `Extend` — because the worksheet runs them once per story, and the five
  belief cuts are one record (`BeliefRun.Cuts`, by ratio) that Split's
  optional layer and Belief Split both edit. Three dock views
  (`WorkshopView.Split` / `Belief` / `Extend`) in `Editor.razor`; handlers
  under `// ---- The Split family`. Every stop-callout is the worksheet's
  words in a `.stop-callout`, and the rung below each sits behind a "Keep
  going — rung N" button until the writer presses it or the rung has text
  (`_splitRung` / `_beliefRung`, session only — a reopened run shows what it
  holds and nothing about how far it got). Alive / flat and y / n are the
  `AliveFlat` component: two pills, press again to clear.
  **Write-back** is `SplitScript` (`Passage.Web/Services`, tested): eight
  `##` sequences under four `#` acts (1, 2A, 2B, 3), the seven turns as
  `= Turn name: text` synopsis lines on the sequence each one closes, A and
  Z as `= Starts:` / `= Ends:`, and every turn not found yet as a bracket
  (`[MIDPOINT]`, `[PINCH 1]`, …) so Fill hands it back. Headings carry
  `[[id:…]]` via `BeatBoardText.BuildCardLines`, the same path the board's
  own edits use, so the Board populates from a run at any rung. The first
  write appends the whole structure by `insertLinesAt`; every write after
  fills in only the turns the script still holds as brackets, by
  `replaceInLine` on the exact line text, and names any turn that already
  has text in the script rather than overwriting it (decision B: promoted
  text is the script's). Each field shows "In the script: …" when the
  script's line has real text. Verified by hand: undo steps back through
  both writes exactly, redo restores, the Board shows the lanes.
  **Shape line** is `ShapeLine.Derive(BoardLanes)` (tested) in the status
  bar, only when the script has a Sequence: `▪`/`▫` from whether the
  sequence holds a scene or section, `●`/`·`/`–` from the turn's synopsis
  value on that sequence (real text / bracket or absent / a dash). Nothing
  stored. Extend Backward appends the next link only when the last one has
  text; the mystery-lens toggle sits on that open link alone and swaps its
  question; the chain is shown read forward, start to Z, above the alive /
  flat mark. Tests `TestSplitScriptLinesAndParse`,
  `TestShapeLineDerivesFromLanes`, `TestSlateStoreSplitFamilyRoundTrip`.
- **Deviations, all deliberate:** write-back is offered at every rung, not
  only on completing rung 3 — the worksheet's own stop-callouts say each
  rung is a complete shape, and a rung-1 write gives the Board a split
  story with six brackets to fill, which is more use than a form that
  withholds the Board until all seven are typed. The plan's "emits `#`/`##`
  sections" became eight sequences under four acts because the vault
  defines the shape as eight sequences with seven joins; the acts are the
  frame the Board needs for lanes. A dropped pinch is marked in the script
  (`= Pinch 1: –`), not in the runner, because the shape line derives from
  the script alone. Belief Split's matched plot turn is shown as a hint
  from the Split run in the same sidecar; nothing is copied between them.

### Phase 4 — Bridge and Position · `done`

Two small runners, plausibly one session for both. (The Lens moved to Phase
5: the Revise worksheet has it as "pick one", applied only where a check is
flagged. The earlier "roll for a lens" idea was a misreading — rolling is
how Ideation picks a *lane*.)

- **Bridge** (`Generate — Ideation or Bridge.md`, Path B): A and Z entered
  first, non-skippable *as a matter of sequence*, not validation — the
  worksheet's own words are "don't skip this — the whole tool depends on
  it". Then rounds of three candidate links, 30 seconds each (decision E),
  with a "picked / half-picked" line per round. Close: "the wire, read start
  to finish". The slow card version is Treatment-stage work and is out of
  scope; the runner says so where the worksheet says so.
- **Position** (`Place or Build — Position or Fill.md`, Path A): "the
  moment" first, then turns of Slot / Needs before it / Forces after it /
  Alive or flat, with the worksheet's question shown: *if this moment sat
  right here, what would it need to be true right before it, and what would
  it force to happen right after?* The seven slots are the app's own seven
  turns. Close: a shortlist of one or two alive answers, and the line that
  Split or Fill picks the winner another day. The only place with a natural
  ceiling — seven slots, and the UI can say that.

- **Done:** `SlateDocument.Bridges` (`BridgeRun`: A, Z, rounds of three
  candidates + picked/half-picked, wire) and `Positions` (`PositionRun`:
  moment, up to seven turns of slot / before / after / alive, shortlist of
  two), kept per script and listed by what they are about — "A → Z" and the
  moment — the same shape as the chains; empty runs vanish on save. Views
  `WorkshopView.Bridges` / `Bridge` / `Positions` / `Position` in
  `Editor.razor`, handlers under `// ---- Bridge and Position`. Bridge: the
  ends come first on the page and the label says why, nothing validates;
  the three candidates of every round are burst slots (`#bridge-slots`,
  same `startBurst`, same per-script setting — the hint says the
  worksheet's 30 s a round is about 10 each); *Another round — the gap still
  needs more than one link* is an explicit button, not an auto-append, since
  the worksheet makes the next round conditional; the close shows the wire
  read forward (A, the picked lines, Z) above the wire field; the card
  version is named in the stop-callout as Treatment-stage work not built
  here, pointing at A→Z→Split for a whole story's gap. Position: the slot is
  a select over `SplitScript.Turns` (a slot already tried in another turn is
  marked), with "In the script: …" under it when the script's synopsis for
  that turn has text; alive / flat is the `AliveFlat` pills; *Another turn*
  disappears at seven with the worksheet's own ceiling line; the shortlist
  shows "Alive so far: …" above it. `StartBurstAsync` now takes the slots'
  root id. Test `TestSlateStoreBridgeAndPositionRoundTrip`.
- **Deviations, all deliberate:** the burst timer is per candidate, not per
  round of three, because decision E's Enter-moves-to-the-next-slot is the
  mechanism and a single 30-second clock over three fields would need a
  second timer design for one tool. No timer on Position — the worksheet
  has none.

### Phase 5 — Push/Pull · `done`

The only independent check in the system and the stated payout, so it gets its
own row.

- Setup first: the scene or sequence, and "what it's supposed to do in the
  story".
- Six checks in the worksheet's priority order — polarity, linkage, third
  rail, migration, escalation, two-test — each shown as its actual question,
  with an answer field and a Y/N flag. Run one and stop, or run all six.
- Scope selector: this scene (derived from caret position) or this stretch;
  the worksheet's "going deeper" note says a stretch is checked as a whole,
  not scene by scene.
- A flagged check offers **the Lens** on the flagged thing only, not the
  whole scene: the seven lenses as a pick-one list, each with its "what you
  actually do" line, timed 3–5 minutes (decision E). Output is a fragment in
  the sidecar, never inserted. Close: "did it move the flagged problem? Y/N —
  one line why".
- Clean checks are not punished with more work. That's in the worksheet and it
  should be in the UI's behaviour, not just its copy.
- **Going deeper, both optional and both after the checks:** the diagnostic
  Belief Split (A, Z, which named shape, muddled between two? seam?) and the
  read-back and crit (what's working / what still doesn't sit right / ready to
  move on or sit another day). These are fillable fields in the worksheet
  now, so they are in scope — the "not going in" line below is narrowed to
  crit that needs a reader.

- **Done:** `SlateDocument.Revisions` (`ReviseRun`: scene, purpose, the six
  `ReviseCheck`s by name — answer, Y/N flag, and the check's own Lens,
  fragment, moved Y/N and why — plus the diagnostic belief fields and the
  three read-back fields), kept per script and listed by the scene or
  stretch checked; the same list-or-open shape as the chains. Views
  `WorkshopView.Revisions` / `Revise`, handlers under `// ---- Push/Pull
  and Lens`. Scope: *Use the scene under the caret* fills the scene field
  from the nearest scene or section heading above the caret
  (`SynopsisPlacement.Find`), and the worksheet's "going deeper — the whole
  stretch" note sits under the field; no selector. Each check is its
  question, an answer and a Y/N flag (`AliveFlat` with Y / N labels). A
  **Y** reveals that check's Lens block and nothing else does: pick one of
  the seven (select, with its "what you actually do" line beneath), *Apply
  it timed — N minutes* (`BurstSettings.LensMinutes`, default 4, one
  `startBurst` on a root holding only that fragment field), the fragment,
  and *Did it move the flagged problem?* Y/N with one line why. The
  diagnostic Belief Split (A, Z, named shape, cleanly or muddled, the seam,
  a Lens at the seam) and the read-back and crit sit behind *Going deeper
  still* / *Ready to close for real* buttons unless they have text.
  `passage.js`: the burst display is now looked up inside the root first,
  and a `data-slot-multiline` slot keeps Enter as a line break — a
  four-minute rewrite is not a one-liner. Test `TestSlateStoreReviseRoundTrip`.
- **Deviations, all deliberate:** the Lens lives under each flagged check
  rather than as one block for the run, because the worksheet's own rule is
  "to the flagged thing only" and two flagged checks are two flagged
  things. Scope is a button plus the worksheet's note, not a scene/stretch
  toggle — a toggle would be state with nothing to do. The Lens minutes are
  their own setting beside the burst seconds so a Lens never inherits a
  ten-second clock from the chain.

### Phase 6 — Ideation burst mode · `done`

Full-screen overlay, the one engine that isn't about your script.

- Ten lanes, lane locked for the sitting, roller to pick ("don't deliberate
  — deliberating isn't the drill"). Each lane's prompt from the worksheet,
  shown once: "read it once, then start the round — don't re-read mid-round".
- 90-second rounds of Input → Output, timer in JS per decision E, "the only
  failure state is silence" as the only rule shown.
- Close: circle one line worth keeping. One. "You don't have to know why
  yet." Output goes to a sidecar keyed to no script at all.
- Lane 10 is a three-step Character Flaw Brainstorm; the full eight-question
  chain is Phase 2's, and the runner should point there rather than repeat it.
- **Resolved:** the worksheet's Situation collision lane is "two unrelated
  fragments — anything, don't cherry-pick", so it needs no corpus. Nothing
  from `Brainstorming/` is required.

- **Done:** its own file, `<root>/.slate/ideation.json` (`IdeationDocument`:
  `RoundSeconds` default 90, `Current` sitting — lane, rounds of Input /
  Output, kept line — and `Kept`, one line per finished sitting).
  `SlateStore.LoadIdeation` / `SaveIdeation`; `PruneOrphans` skips the file
  by name since it belongs to no script (tested,
  `TestSlateStoreIdeationIsKeptApart`). The overlay is `.ideation-overlay`
  at the end of `Editor.razor`, fixed full-screen over everything, opened
  from the dock's *Ideation burst* button and closed with × keeping the
  sitting; handlers under `// ---- Ideation burst`. No sitting: the lane
  pick — *Roll* (`Random.Shared`) or tap one of the ten — and below it the
  lines kept, each with an × to let it go. In a sitting: the lane's prompt
  verbatim, shown once with "read it once, then start the round"; lane 10
  points at the Character Flaw Brainstorm for the full chain; the only rule
  on the page is the stop-callout *The only failure state is silence*.
  Each round is Input (untimed), *Start the round* (one `startBurst` on
  `#ideation-round-N`, a multiline slot, the seconds as set) and Output;
  *Keep a line from this* copies the output into the kept line; *Next
  round* is a button. *Done for this sitting* files the kept line, drops
  the rounds and returns to the lane pick — the worksheet says don't review
  in the same sitting, so nothing invites it. Verified by hand at 375 px:
  roll, round with the clock on the output, close, reload, reopen — the
  sitting is there with its lane locked — then done.
- **Deviations, all deliberate:** kept lines are listed on the pick screen.
  They are lines, not sessions — no date, no lane, no count — and the
  worksheet's own "optionally skim later" needs somewhere to skim. Rounds
  are never listed after a sitting ends. No ten-minute session timer: the
  worksheet says nothing depends on it. The round length is one setting in
  the ideation file rather than the per-script burst seconds, because the
  overlay belongs to no script.

### Phase 7 — The assist seam · `not started`

- `IStoryPartner` in `Passage.Core`, one method:
  `Task<IReadOnlyList<string>> SuggestAsync(SuggestionRequest, CancellationToken)`.
- `NullStoryPartner` registered by default, returns empty, UI renders nothing
  extra. Config-gated, off by default, no egress in the default build.
- Plugs in at roughly six call sites, all of which have the same shape
  already: "given this partially-filled run, offer N candidates for field X".
  Fill's three-options-you-expect-to-reject is the canonical one.

**This is built last, and here is the pushback on building the seam first.**
`PROJECT_RULES` rule 2 forbids speculative abstractions for single-use code,
and rule 7 says to surface conflicts rather than average them. An interface
with exactly one no-op implementation, written before any real implementation
exists, is precisely what rule 2 names. Deferring it costs nothing, because
the seam is *data-shaped*, not interface-shaped: every runner already produces
a structured record, and every assist is the same request/response. If the
records are right — which Phases 1–6 have to get right anyway — the interface
is an afternoon whenever you want it, and zero speculative code ships in the
meantime.

---

## Not going in

Stated so a later session doesn't re-derive them as gaps:

- The Slate board, Seeds, stages, last-touched, next-move. Obsidian.
- Briefs, script sessions, and crit that needs a reader. (The Revise
  worksheet's own read-back fields are in — see Phase 5.)
- The Bridge card version, misbelief-first, five-layer scene build, structure
  remix. Undecided in `engines.md` itself, so undecided here.
- Anything that logs, counts, streaks or reviews. Rule 0.

## Risks

1. **The worksheets become forms.** The likeliest failure, and the one with
   precedent — the Short/Medium/Long trap was exactly this, caught only
   because you pushed back on it. Rule 0 is the mitigation and it needs to be
   re-read at the start of every session, not just this one.
2. **Bracket false positives** make the queue noisy and the "hand me one"
   button untrustworthy. Rank-don't-filter plus an ignore list, per decision C.
   If the queue hands you `[sic]` twice, you'll stop pressing the button.
3. **Sidecar drift.** Files change on the volume from outside the app. Every
   line reference must degrade gracefully.
4. **Undo death on write-back.** Already bitten this repo twice. Ranged
   replacements only, verified by hand.

## The honest risk

With the board staying in Obsidian, you still start every session in Obsidian
and move to Passage to do the work. Two apps, and the friction between them is
exactly where a session dies. The engines will get used only if the bracket
queue alone is good enough to make Passage the place the work happens.

Which is testable cheaply: **Phase 1 is that test.** Build it, use it for two
weeks, and see whether you actually open Passage for a Fill session, or keep
filling brackets in Obsidian because that's where the board already had you.
If it's the latter, Phases 2–6 are six sessions of building tools you'll open
from the wrong app — and the correct response is to revisit the scope decision
and bring the board across after all, not to keep building engines.

Don't build Phase 2 until Phase 1 has answered that.
