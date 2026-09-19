# Slate — session prompts

One prompt per phase of `docs/SLATE-PLAN.md`, in build order. Copy the block,
paste it, let it run, review the diff, commit, `/clear`, next. Same scheme as
`docs/WEB-SESSIONS.md`: each session reads only its own row.

Every prompt starts by re-reading Rule 0. That is deliberate — the plan's
first-listed risk is the worksheets turning into forms, and it is caught by
re-reading the rule, not by remembering it.

**Stop after Phase 1.** Use it for two weeks before scheduling Phase 2; the
plan's "honest risk" section says why.

---

## Phase 0 — foundations · done

```
Read CLAUDE.md, PROJECT_RULES.md, and in docs/SLATE-PLAN.md: Rule 0, the
Architecture decisions, and the Phase 0 row. Read only what those name.

Build the foundations and nothing user-visible beyond an empty dock:
- SlateStore beside ScriptLibrary: read/write /data/.slate/<name>.json with
  slateVersion 1, keys through ScriptLibrary.TryValidateName, cascade on the
  app's own delete, prune orphans on load with a status message.
- aside.workshop on the right of .main-area: collapsible, resizable, state in
  the existing passage.session.v1 key (not a second key). Every engine listed,
  every entry disabled.
- Tests: round-trip, invalid-name rejection, orphan pruning.

Success: dock opens, closes, resizes, and all three survive a reload.
`dotnet build Passage.Web.slnf` clean. Set Phase 0 to `done` and commit.
```

## Phase 1 — Brackets and Fill · done

```
Read CLAUDE.md, PROJECT_RULES.md, and in docs/SLATE-PLAN.md: Rule 0,
decisions C and D, and the Phase 1 row. Then the Phase 0 "Done" notes for
where the dock and SlateStore live. Read only the line ranges those name,
plus Passage/Passage.Parser/FountainMarkup.cs and the NOTES panel in
Editor.razor for the jump-to-line interaction to copy.

Build brackets and Fill:
- BracketScanner in Passage.Parser beside FountainMarkup, pure string logic,
  with tests: ranking (ALL CAPS / SOMETHING / CONDITION / TODO first),
  omission stripping so [[notes]] and boneyard never match, line numbers,
  empty and malformed input.
- DocumentAnalyzer surfaces Brackets.
- BRACKETS panel in the dock: list, count, click to jump.
- "Hand me one": one button, one bracket, highest-ranked not on the ignore
  list, round-robin. It must not route through the list.
- Fill runner as the Path B worksheet reads, escape hatch verbatim, no
  required fields, closable at any point with partial state kept.
- Accept = ranged replacement of the bracket span via the existing
  passage.replaceLineRange. Never ReplaceEditorContentAsync.
- On accept, offer the next bracket. No stopping point, no prompt suggesting one.
- Ignore list in the sidecar (first real SlateDocument field).

Success: open a script with brackets, press one button, fill one bracket at
phone width, and the script is edited correctly with undo intact — verify undo
by hand, do not assume. `dotnet build Passage.Web.slnf` clean, tests pass.
Set Phase 1 to `done` and commit. Then stop; do not start Phase 2.
```

## Phase 2 — Forward chain

```
Read CLAUDE.md, PROJECT_RULES.md, and in docs/SLATE-PLAN.md: Rule 0,
decision D, and the Phase 2 row, plus the Phase 0 and 1 "Done" notes for
where things live. Read only the line ranges those name.

Build the chain runner (Path A: each round's Consequence auto-fills the next
Want — this carry is the whole tool) and the fixed eight-question Character
Flaw Brainstorm (Path B; all eight may show, stopping at any one is fine).
Rounds are unbounded. The per-round burst timer runs in passage.js, default
off; expiry advances the round and does nothing else — no red, no count.
Promotion: "what's the scene now, in one line?" inserts a `=` synopsis under
the current section by ranged insert (passage.insertLinesAt).

Success: run a chain from a phone, close it half-done, reopen the script and
the rounds are still there; promote a line and undo still works. Build clean,
tests pass. Set Phase 2 to `done` and commit.
```

## Phase 3 — The Split family

```
Read CLAUDE.md, PROJECT_RULES.md, and in docs/SLATE-PLAN.md: Rule 0,
decisions B and D, and the Phase 3 row, plus earlier "Done" notes. Read
Passage/Passage.Parser/BeatBoardText.cs for the section-writing helpers.

One runner, three paths: A→Z→Split, Belief Split, Extend Backward. The
worksheet's stop-callouts are rendered as UI text. Completing Split rung 3
emits #/## sections with synopses by ranged insert so the Beat Board
populates from the run. Shape line in the status bar, derived from BoardLanes
only — it is not a progress bar and stores nothing. Extend Backward: only the
most recent link opens the next field; mystery-lens toggle on that link.

Success: a Split run produces sections the BOARD view shows; the shape line
reflects them; undo survives. Build clean, tests pass. Set Phase 3 to `done`
and commit.
```

## Phase 4 — Bridge, Position, Lens

```
Read CLAUDE.md, PROJECT_RULES.md, and in docs/SLATE-PLAN.md: Rule 0 and the
Phase 4 row, plus earlier "Done" notes. Read only what those name.

Three small runners. Bridge: two ends first (sequence, not validation), then
rounds of three links; say where the card version is out of scope rather than
omitting it silently. Position: slots read live from the script's seven
turns; alive/flat per slot; the UI may say seven is the ceiling. Lens: roller
is the default affordance, manual pick secondary; applies to the editor
selection; output is a fragment saved to the sidecar, never inserted.

Success: each runner opens, closes mid-way with state kept, and reopens.
Build clean, tests pass. Set Phase 4 to `done` and commit.
```

## Phase 5 — Push/Pull

```
Read CLAUDE.md, PROJECT_RULES.md, and in docs/SLATE-PLAN.md: Rule 0 and the
Phase 5 row, plus earlier "Done" notes. Read only what those name.

Six checks in the worksheet's order, each with an answer and a Y/N flag; run
one or all. Scope: this scene (from caret) or this stretch. A flagged check
offers a Lens on the flagged thing only. A clean check offers nothing more.

Success: run one check on the scene under the caret, flag it, get a Lens on
just that. Build clean, tests pass. Set Phase 5 to `done` and commit.
```

## Phase 6 — Ideation burst mode

```
Decide the Collision open question in docs/SLATE-PLAN.md Phase 6 before
starting. Then read CLAUDE.md, PROJECT_RULES.md, and in SLATE-PLAN.md: Rule 0,
decisions A and D, and the Phase 6 row.

Full-screen overlay with nothing else on it. Ten lanes, locked for the
sitting, roller to pick. 90-second rounds timed in passage.js; the only rule
shown is "the only failure state is silence". Close by circling one line to
keep. Output to a sidecar keyed to no script.

Success: a full sitting from a phone, timer survives a circuit reconnect.
Build clean. Set Phase 6 to `done` and commit.
```

## Phase 7 — The assist seam

```
Read CLAUDE.md, PROJECT_RULES.md, and in docs/SLATE-PLAN.md the Phase 7 row
and the Phase 1–6 "Done" notes for where each runner's record is built.

Add IStoryPartner to Passage.Core with the one SuggestAsync method,
NullStoryPartner registered by default, config-gated and off. Wire the call
sites the row names; with the null partner the UI must render nothing extra.
No egress in the default build.

Success: every runner behaves identically to before with the null partner.
Build clean, tests pass. Set Phase 7 to `done` and commit.
```
