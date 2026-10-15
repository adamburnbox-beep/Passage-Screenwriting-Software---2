# Highland-informed upgrades — session prompts

One prompt per Claude Code session, in order. Before each one:
`git checkout main && git pull`. Paste the block, let it run, try the change in
the running app, merge the PR when CI is green, `/clear`, next.

Every prompt reads only `CLAUDE.md`, `PROJECT_RULES.md`, one row of
`docs/HIGHLAND-PLAN.md`, and the teardown sections that row names. Do not merge
two sessions into one.

---

## Session H0 — orientation (run once, ~5 minutes, no code)

```
Read CLAUDE.md, PROJECT_RULES.md and docs/HIGHLAND-PLAN.md (the whole plan
file, this once). Do not read docs/reference/HIGHLAND2_TEARDOWN.md yet.

Then, without changing code:
1. Run `dotnet build Passage.Web.slnf` and the tests; report results.
2. Check the line ranges the plan cites for Layout (ScreenplayLayoutBuilder.cs),
   Analyzer (~280) and the desktop BuildPages call sites. Correct any that
   are wrong in docs/HIGHLAND-PLAN.md.
3. Read "Standing decision — shared-library changes" and tell me in five
   lines or fewer whether the opt-in PaginationOptions approach is sound
   given how BuildPages is used, or what you'd do instead.

Only file you may change: docs/HIGHLAND-PLAN.md (line-range corrections).
Branch: docs/highland-h0. Commit, push, open a PR.
```

---

## Tier H1

### Session H1.1 — keep-together pagination

```
Read CLAUDE.md, PROJECT_RULES.md and row H1.1 plus the "Standing decision"
section of docs/HIGHLAND-PLAN.md. From docs/reference/HIGHLAND2_TEARDOWN.md
read only sections 11 and 12. Read only the Layout line ranges the row names.

Implement opt-in keep-together pagination in Passage.Export as the row
describes. Passing null options must reproduce today's output exactly.
Passage.Web passes options so its preview and PDF use the new behaviour.

Before writing code, state your grouping rules and the test fixtures you'll
use, and wait for my OK (PROJECT_RULES 1 and 10).

Success: tests prove (a) null options = old output, (b) no page ends on a
scene heading or bare character cue in the fixture. `dotnet build
Passage.Web.slnf` passes. Set H1.1 to `done`, commit on branch
feat/highland-h1-1-pagination, push, open a PR.
```

### Session H1.2 — (MORE) / (CONT'D)

```
Read CLAUDE.md, PROJECT_RULES.md and row H1.2 of docs/HIGHLAND-PLAN.md. From
docs/reference/HIGHLAND2_TEARDOWN.md read only sections 12 and 13. Read the
PaginationOptions code added in H1.1.

Implement dialogue splitting with (MORE)/(CONT'D) per the row, behind the same
opt-in options. Sentence-boundary breaks first; never split parentheticals or
dual dialogue.

Success: a fixture with a long speech splits at a sentence end with the
correct MORE/CONT'D lines; preview and PDF agree; null-options test still
passes. Set H1.2 to `done`; branch feat/highland-h1-2-more-contd; PR.
```

### Session H1.3 — scene numbers

```
Read CLAUDE.md, PROJECT_RULES.md and row H1.3 of docs/HIGHLAND-PLAN.md. From
docs/reference/HIGHLAND2_TEARDOWN.md read only section 3's scene-number row.

Parse #…# scene numbers, render them in both margins (web preview + PDF), and
add Add/Remove Scene Numbers commands. This touches Passage.Parser: first add
a test proving headings without #…# parse exactly as before.

Success: `INT. HOUSE - DAY #12#` previews with 12 in both margins; Remove
strips them; Add numbers unnumbered scenes without renumbering existing ones.
Set H1.3 to `done`; branch feat/highland-h1-3-scene-numbers; PR.
```

### Session H1.4 — auto-(CONT'D) in the editor

```
Read CLAUDE.md, PROJECT_RULES.md and row H1.4 of docs/HIGHLAND-PLAN.md. From
docs/reference/HIGHLAND2_TEARDOWN.md read only section 13. In passage.js read
only the autocomplete/hint code from WEB-PARITY 3.4 (find it with grep).

Implement client-side auto-(CONT'D) per the row, with a settings toggle
(default on). No per-keystroke Blazor round-trips (CLAUDE.md trap 4).

Success: repeated speaker after action gets the offer; a different speaker or
a new scene does not. Set H1.4 to `done`; branch feat/highland-h1-4-auto-contd; PR.
```

---

## Tier H2

### Session H2.1 — .highland import/export

```
Read CLAUDE.md, PROJECT_RULES.md and row H2.1 of docs/HIGHLAND-PLAN.md. From
docs/reference/HIGHLAND2_TEARDOWN.md read only Part 1 section 2. Read
ScriptLibrary.cs and ExportService.cs.

Implement .highland import and export per the row using
System.IO.Compression. Do not register anything in
ExporterCatalog.DefaultExporters. Report every skipped sidecar visibly.

Success: round-trip test (export then import) yields identical text; a
hand-built Highland-style zip imports. Set H2.1 to `done`; branch
feat/highland-h2-1-highland-files; PR.
```

### Session H2.2 — FDX export

```
Read CLAUDE.md, PROJECT_RULES.md and row H2.2 of docs/HIGHLAND-PLAN.md. Read
IExporter.cs, ScreenplayExporter.cs and ExportService.cs.

Add an FDX exporter wired only through ExportService (not DefaultExporters).
Map element types per the row. Before coding, list the FDX element mapping
and wait for my OK.

Success: a fixture covering every element type exports to well-formed XML
with the right Paragraph Types (test asserts them). Set H2.2 to `done`;
branch feat/highland-h2-2-fdx; PR.
```

---

## Tier H3

### Session H3.1 — character panel

```
Read CLAUDE.md, PROJECT_RULES.md and row H3.1 of docs/HIGHLAND-PLAN.md. From
docs/reference/HIGHLAND2_TEARDOWN.md read only Part 1 section 6's Character
bullets. Read SlateStore.cs for the storage pattern and the Analyzer code that
collects unique character names.

Build the character panel: list with line/word counts, rename-everywhere as
one undoable edit, per-character highlight colour (client-side marks),
persisted per document. Propose the UI placement first and wait for my OK.

Success: rename is one undo step; highlights survive reload. Set H3.1 to
`done`; branch feat/highland-h3-1-characters; PR.
```

### Session H3.2 — dialogue balance report

```
Read CLAUDE.md, PROJECT_RULES.md and row H3.2 of docs/HIGHLAND-PLAN.md. From
docs/reference/HIGHLAND2_TEARDOWN.md read only section 15's Gender report
bullet. Read the H3.1 character storage code.

Add gender tags to the character panel and the on-screen report the row
describes, using theme tokens only.

Success: test proves the per-tag totals equal the sums of the per-character
rows. Set H3.2 to `done`; branch feat/highland-h3-2-dialogue-report; PR.
```

### Session H3.3 — typewriter scrolling + focus line

```
Read CLAUDE.md, PROJECT_RULES.md and row H3.3 of docs/HIGHLAND-PLAN.md. In
passage.js read only the session-state code (passage.session.v1) and the
CodeMirror setup.

Implement both toggles client-side per the row.

Success: caret stays centred while typing in a 100-page script without
visible jank; toggles survive reload. Set H3.3 to `done`; branch
feat/highland-h3-3-typewriter; PR.
```

### Session H3.4 — reading time

```
Read CLAUDE.md, PROJECT_RULES.md and row H3.4 of docs/HIGHLAND-PLAN.md. Read
TextAnalysis.cs and the stats display range in Editor.razor (find with grep).

Add reading time per the row via a new method; existing TextAnalysis outputs
must not change. Set H3.4 to `done`; branch feat/highland-h3-4-reading-time; PR.
```

---

## Tier H4

Use the same pattern for H4.1–H4.3: read CLAUDE.md, PROJECT_RULES.md, the row,
and the teardown sections it names; state the plan and wait for OK; implement;
test; set the row `done`; branch `feat/highland-h4-N-…`; PR.

### Session H4.4 — revision mode (design note only)

```
Read CLAUDE.md, PROJECT_RULES.md and row H4.4 of docs/HIGHLAND-PLAN.md. From
docs/reference/HIGHLAND2_TEARDOWN.md read only section 15's Revisions bullet
and Part 1 section 5's Revisions menu.

Write docs/REVISIONS-DESIGN.md (≤ 150 lines): storage options compared, how
marks survive edits, colour sequence, PDF stars, and your recommendation.
No code. Branch docs/highland-h4-4-design; PR.
```
