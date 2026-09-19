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
`How the Slate Works.md`, `engines.md`, `loosening-up-practice.md`, and the six
files in `Worksheets/`.

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
  of squeezing the editor. CSS under `/* ---- Slate workshop dock ---- */`.
  Tests in `Passage/Passage.Tests/Program.cs` (`TestSlateStore*`); the test
  project now references `Passage.Web` for them.
- **Deviation — no Save As cascade.** The web app has no Save As command
  (WEB-PARITY 1.4): the "Save script" dialog only names an untitled buffer,
  which cannot have a sidecar. Delete of the open file already drops both the
  file and its sidecar, so the buffer that becomes Untitled afterwards starts
  clean. If a real Save As is ever added, the cascade (copy the sidecar to the
  new name) goes in the same change.

### Phase 1 — Brackets and Fill · `not started` · **build this one first and stop**

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

### Phase 2 — Forward chain (WOAC + Character Flaw Brainstorm) · `not started`

- **Chain runner**, Path A: a round table where each round's Consequence
  auto-populates the next round's Want. That auto-carry *is* the tool — it's
  the thing that removes having to invent the next question, which
  `writing-practice` names as the actual block. Get it exactly right.
- Open-ended: rounds add indefinitely, no cap, no suggested stopping point
  beyond the worksheet's own "5–7 is a pattern, not a rule" note.
- **Path B, Character Flaw Brainstorm**: the fixed eight-question chain. This
  one *is* fixed-length by design — the eight land on a shape together — so the
  UI may show all eight, but must still allow stopping at any one.
- **Burst timer**, in JS: per-round countdown, configurable, default off.
  Expiry advances the round and does nothing else. No sound of failure, no
  colour change to red, no count of missed rounds.
- **Promotion path:** "read it back — what's the scene now, in one line?" →
  that line inserts into the script as a synopsis (`=`) under the current
  section, by ranged insert.

### Phase 3 — The Split family · `not started`

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

### Phase 4 — Bridge, Position, Lens · `not started`

Three small runners, plausibly one session each or one session for all three.

- **Bridge**: two fixed ends (entered first, non-skippable *as a matter of
  sequence*, not validation — the tool is meaningless without them), then
  rounds of three candidate links. The slow card version is Treatment-stage
  work and is out of scope; the runner should say so where the worksheet says
  so rather than silently omitting it.
- **Position**: the seven slots are the app's own seven turns, so this runner
  can read them live from the script instead of asking you to name them.
  Alive/flat per slot, shortlist of one or two out. The only place in Slate
  with a natural ceiling — seven slots, and the UI can say that.
- **The Lens**: seven lenses with a roller (the worksheet says roll if
  choosing feels like another way to stall — so the roller is the default
  affordance and picking manually is secondary). Applies to the current
  selection. Output is a fragment, and a fragment goes to the sidecar, never
  into the script automatically.

### Phase 5 — Push/Pull · `not started`

The only independent check in the system and the stated payout, so it gets its
own row.

- Six checks in the worksheet's priority order — polarity, linkage, third
  rail, migration, escalation, two-test — each with an answer field and a Y/N
  flag. Run one and stop, or run all six.
- Scope selector: this scene (derived from caret position) or this stretch.
- A flagged check offers a Lens **on the flagged thing only**, not the whole
  scene.
- Clean checks are not punished with more work. That's in the worksheet and it
  should be in the UI's behaviour, not just its copy.

### Phase 6 — Ideation burst mode · `not started`

Full-screen overlay, the one engine that isn't about your script.

- Ten lanes, lane locked for the sitting, roller to pick.
- 90-second rounds, timer in JS, "the only failure state is silence" as the
  only rule shown.
- Close: circle one line worth keeping. One. Output goes to a sidecar keyed to
  no script at all.
- **Open question:** the Collision prompt needs your `Brainstorming/` corpus
  (86 files, 3,655 pairs), which lives in Obsidian, not `/data`. Either a
  seeds file gets synced into `/data`, or Collision is dropped from the app
  version and stays a vault-side prompt. Worth deciding before this phase, not
  during it.

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
- Briefs, read-back and crit, script sessions. These need a person or a model
  and aren't engines.
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
