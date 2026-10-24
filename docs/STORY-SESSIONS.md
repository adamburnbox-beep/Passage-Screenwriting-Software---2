# Story development tools — session prompts

One prompt per Claude Code session, in order. Before each one:
`git checkout main && git pull`. Paste the block, let it run, try the change in
the running app, merge the PR when CI is green, `/clear`, next.

Every prompt reads only `CLAUDE.md`, `PROJECT_RULES.md`, the standing
decisions it names and one row of `docs/STORY-PLAN.md`, plus the line ranges
that row lists. Do not merge two sessions into one. The behaviour reference is
Arc Studio Pro's story tools as described in the plan: never copy its code,
assets, names or branding.

Order and dependencies:

```
S0 ─ S1.1 ─┬─ S1.2 ─┬─ S1.4 ─ S1.5 ─┬─ S2.1
           │ S1.3 ──┘               ├─ S3.1  (also needs S1.2)
           │                        └─ S3.2 ─ S3.3
           ├─ S2.2   (independent of S1.2–S1.5)
           ├─ S2.3   (draws S1.4's bars/dots if present)
           └─ S2.4   (independent; can run any time)
S4.1 (design note) any time; implementation only after approval
```

---

## Session S0 — orientation (run once, ~5 minutes, no code)

```
Read CLAUDE.md, PROJECT_RULES.md and docs/STORY-PLAN.md (the whole plan file,
this once).

Then, without changing code:
1. Run `dotnet build Passage.Web.slnf` and the tests; report results.
2. Check five line ranges from the plan's table at the top against the code
   (grep first). Correct any that drifted in docs/STORY-PLAN.md.
3. Read standing decision A2 and tell me in five lines or fewer whether
   lazy [[id:…]] stamping is sound given how the parser assigns ids, or what
   you'd do instead.

Only file you may change: docs/STORY-PLAN.md (line-range corrections).
Branch: docs/story-s0. Commit, push, open a PR.
```

---

## Tier S1

### Session S1.1 — stable card ids + key beats

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decisions A, A2 and B and row S1.1. Read only the line ranges S1.1 names.

Implement lazy [[id:…]] stamping, duplicate-id detection, the key-beat star
on Board cards and beats, the "Key only" dimming toggle on BOARD and BEATS,
and the Detached list, as the row describes. Stamp by appending to the
heading line with passage.replaceInLine. Do not rebuild it with
BuildCardLines, which drops scene numbers. Passage.Parser and BeatBoardText
are read only.

Before writing code, state the StoryData/ItemTags shape you will add and how
detachment is computed, and wait for my OK (PROJECT_RULES 1 and 10).

Success: the row's success list, checked in the running app, including the
Ctrl+Z re-attach and an unchanged .fountain hash after starring a beat.
Tests: the row's list. `dotnet build Passage.Web.slnf` passes.
Then set S1.1's status, add a **Landed:** note, commit.
Branch: feat/story-s1-1-ids-key-beats. Push, open a PR.
Budget: ~90k tokens; stop and ask if trending past it.
```

### Session S1.2 — storylines and the ELEMENTS tab

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decision A and row S1.2. Read only the line ranges S1.2 names.

Add a fifth sidebar tab, ELEMENTS, with a Storylines section (name, Cast
palette colour, description; add, edit, delete with a counted confirm) and
the one-line pointer to CAST for character colours. Do not build a second
character list. Leave a placeholder for Locations (S1.3).

Success: the row's success list in the running app. Tests: the row's list.
`dotnet build Passage.Web.slnf` passes.
Then set S1.2's status, add a **Landed:** note, commit.
Branch: feat/story-s1-2-storylines. Push, open a PR.
Budget: ~60k tokens.
```

### Session S1.3 — locations

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decision C and row S1.3. Read only the line ranges S1.3 names.

Add Locations.Normalise exactly as decision C specifies, a derived location
list from the analyzer, and the Locations section of ELEMENTS with optional
colours and the "Not in the script" list. Write the normalisation tests
first, from decision C's table, and show them to me before the
implementation.

Success: the row's success list in the running app; typing in SCRIPT is not
slower. Tests: the row's list. `dotnet build Passage.Web.slnf` passes.
Then set S1.3's status, add a **Landed:** note, commit.
Branch: feat/story-s1-3-locations. Push, open a PR.
Budget: ~60k tokens.
```

### Session S1.4 — tagging popover, bars and dots

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decisions A, A2, D and G and row S1.4. Read only the line ranges S1.4 names.

Build one tag popover in passage.js, used by BOARD cards and BEATS beats
(grid and tray). The filter runs in the browser; only Enter calls the server.
Draw storylines as left-edge colour bars and characters as dots, and make
CAST's rename also rename tags.

Before writing code, state the selection model (decision G: click selects; it changes two existing
clicks, so say so in the PR), the JS↔server calls and their
payloads, and wait for my OK.

Success: the row's success list in the running app, including no WebSocket
frames while typing in the popover. Tests: the row's list.
`dotnet build Passage.Web.slnf` passes.
Then set S1.4's status, add a **Landed:** note, commit.
Branch: feat/story-s1-4-tagging. Push, open a PR.
Budget: ~130k tokens. Stop and ask at 150k; splitting display into its own
PR is the fallback.
```

### Session S1.5 — element filters

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decisions B, C and D and row S1.5. Read only the line ranges S1.5 names.

Add the filter bar to BOARD and BEATS (storylines, characters, locations, Key
only), combining as OR within a kind and AND across kinds. Non-matching items
dim; nothing hides. Persist per browser, as the outline filters do. A scene
matches a character by explicit tag or by that character speaking in it, as
the row says.

Success: the row's success list in the running app. Tests: the row's match
table. `dotnet build Passage.Web.slnf` passes.
Then set S1.5's status, add a **Landed:** note, commit.
Branch: feat/story-s1-5-filters. Push, open a PR.
Budget: ~70k tokens.
```

---

## Tier S2

### Session S2.1 — one inbox

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decisions D and E and row S2.1. Read only the line ranges S2.1 names.

Show the Beat Sheet's tray as an inbox strip on BOARD. Add a pointer-gesture
drop from the strip onto a card (JS, one call on release; do not copy the
Board's @ondragover pattern). Convert the beat into a scene card as decision
E describes, moving its tags and star, and add "Send to Board" to the beat
editor.

Before writing code, state the insert-point rule for before/after on an act,
a sequence and a scene, and wait for my OK.

Success: the row's success list in the running app. Tests: the row's list.
`dotnet build Passage.Web.slnf` passes.
Then set S2.1's status, add a **Landed:** note, commit.
Branch: feat/story-s2-1-inbox. Push, open a PR.
Budget: ~100k tokens.
```

### Session S2.2 — structure templates

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decision F and row S2.2. Read only the line ranges S2.2 names.

First extract InsertStructureAsync from WriteShapeAsync's first-time branch
with no behaviour change, and confirm the Shape tool still writes and
updates. Then add the four built-in templates (Passage's is
SplitScript.BuildLines of an empty run), the opt-in beat seeding, and "Save
this structure as a template" in .slate/templates.json. Make PruneOrphans
spare templates.json. Templates write only #, ##, ### and = lines. Use the
template names from decision F, crediting each source in its description.

Success: the row's success list in the running app, including an unchanged
PREVIEW page count after applying a template. Tests: the row's list.
`dotnet build Passage.Web.slnf` passes.
Then set S2.2's status, add a **Landed:** note, commit.
Branch: feat/story-s2-2-templates. Push, open a PR.
Budget: ~90k tokens.
```

### Session S2.3 — Board beside the script

```
Read CLAUDE.md, PROJECT_RULES.md and row S2.3 of docs/STORY-PLAN.md. Read
only the line ranges S2.3 names.

Add a "Beside" pane to SCRIPT: a compact Board list (and a Beats list behind
a switch). Clicking jumps the caret without leaving SCRIPT, and the card under
the caret is highlighted. The pane re-renders on parse, never per keystroke.
The width is dragged in JS and stored per browser.

Success: the row's success list in the running app. Tests: the row's list.
`dotnet build Passage.Web.slnf` passes.
Then set S2.3's status, add a **Landed:** note, commit.
Branch: feat/story-s2-3-beside. Push, open a PR.
Budget: ~80k tokens.
```

### Session S2.4 — expand all / collapse all

```
Read CLAUDE.md, PROJECT_RULES.md and row S2.4 of docs/STORY-PLAN.md. Read
only the line ranges S2.4 names.

Add "Descriptions: show/hide" and "Expand all / Collapse all" to the BOARD
toolbar. The description toggle persists per browser.

Success: the row's success list in the running app. Tests: the row's list.
`dotnet build Passage.Web.slnf` passes.
Then set S2.4's status, add a **Landed:** note, commit.
Branch: feat/story-s2-4-expand-collapse. Push, open a PR.
Budget: ~30k tokens.
```

---

## Tier S3

### Session S3.1 — arc view

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decisions A, B and D and row S3.1. Read only the line ranges S3.1 names.

Add an ARC main tab: an SVG chart with one line per storyline through its
tagged items, a named axis per storyline, and points dragged in JS (one call
on release). The x-axis offers Board order (the default) and Pages; items
with no page are listed, never guessed. Check nothing persists
_mainViewIndex before shifting the tab indices.

Before writing code, state the ArcPoint shape and the two x-axis orderings,
and wait for my OK.

Success: the row's success list in the running app. Tests: the row's list.
`dotnet build Passage.Web.slnf` passes.
Then set S3.1's status, add a **Landed:** note, commit.
Branch: feat/story-s3-1-arc. Push, open a PR.
Budget: ~120k tokens.
```

### Session S3.2 — outline export: Fountain and text

```
Read CLAUDE.md, PROJECT_RULES.md, and from docs/STORY-PLAN.md only standing
decision A and row S3.2. Read only the files and ranges S3.2 names.

Add "Export outline…" with field checkboxes, producing Fountain (tags as
[[notes]], so it re-opens as an outline) and plain text, including the
optional Beat Sheet section. Web-only: do not add it to
ExporterCatalog.DefaultExporters.

Success: the row's success list in the running app, including the re-open
round trip. Tests: golden outputs and the parse round trip.
`dotnet build Passage.Web.slnf` passes.
Then set S3.2's status, add a **Landed:** note, commit.
Branch: feat/story-s3-2-outline-export. Push, open a PR.
Budget: ~60k tokens.
```

### Session S3.3 — outline export: PDF

```
Read CLAUDE.md, PROJECT_RULES.md and row S3.3 of docs/STORY-PLAN.md. Read
only the ranges S3.3 names.

Tell me whether you'll take option (a) or (b) from the row, and why, and wait
for my OK; (b) changes a shared library. Then add PDF to "Export outline…"
with the same fields as S3.2.

Success: the row's success list in the running app. Tests: the row's list.
`dotnet build Passage.Web.slnf` passes.
Then set S3.3's status, add a **Landed:** note, commit.
Branch: feat/story-s3-3-outline-pdf. Push, open a PR.
Budget: ~80k tokens.
```

---

## Tier S4

### Session S4.1 — project notes design note (no code)

```
Read CLAUDE.md, PROJECT_RULES.md and row S4.1 of docs/STORY-PLAN.md, then
only the SLATE-PLAN.md and SlateStore.cs ranges it names.

Write the design note the row asks for, as a new subsection under S4.1 in
docs/STORY-PLAN.md: the scope question against SLATE-PLAN's decision, sidecar
vs companion .md files with a recommendation, the Slate/notes boundary, and
templates. Do not write feature code. If I approve the note, add an S4.2
implementation row and a matching session prompt here.

Only file you may change: docs/STORY-PLAN.md (and this file, on approval).
Branch: docs/story-s4-1-notes-design. Commit, push, open a PR.
Budget: ~30k tokens.
```
