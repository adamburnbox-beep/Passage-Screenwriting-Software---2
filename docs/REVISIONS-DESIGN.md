# Revision mode — design note (HIGHLAND-PLAN H4.4, session 1)

Status: **proposal, awaiting Adam's approval.** No code until signed off.
Reference: teardown §15 (Revisions bullet), Part 1 §5 (Revisions menu).

## What Highland does (behaviour only)

- Menu: Revision Mode toggle; colour Red / Green / Blue / None; Mark
  Revision, Clear Revision, Clear All, Go to Next Revision.
- Print: revision colour, margin stars, **revised pages only**.
- Storage: index sets on the text (`revisionIndexes`, `deletions`) inside the
  `.highland` bundle: character offsets the editor keeps up to date.

## Constraints that shape the choice

1. Scripts are plain `.fountain` files on `/data`, also opened by the frozen
   desktop apps and by any text editor. Anything we add must not print or
   break parsing there.
2. Saving is last-write-wins, no auth (`docs/web-app.md`). Two tabs can race.
3. Per-keystroke work belongs in `passage.js`; the server sees text at parse.
4. Fountain has no revision syntax. FDX has one (`Revision` attributes), so
   interop is a later concern, not a format to copy now.
5. `LayoutLine` (`ScreenplayLayoutBuilder.cs:18`) carries no source-line
   index; PDF stars need a shared-library change (see §PDF stars).

## Storage options compared

| | A. Inline markup | B. Sidecar offsets (Highland's model) | C. Sidecar line-text keys | D. Sidecar baseline snapshot + diff |
|---|---|---|---|---|
| Shape | `{{%r: blue}}` on each revised line, or a boneyard block at EOF | `[{start,len,colour}]` in JSON | list of revised line strings / hashes | the text as it stood when each colour began; marks = diff(baseline, now) |
| Fountain file | polluted; desktop apps print `{{%r}}` (null options keep directives) | clean | clean | clean |
| Survives web edits | yes (travels with the line) | only if every edit maps offsets — CM can, but server/autosave races drift | yes, until a line is retyped identically elsewhere | yes — nothing is stored that can drift |
| Survives desktop / text-editor edits | partly (user may delete markers) | **no** — offsets silently wrong | mostly; duplicates ambiguous (`CUT TO:`, `JOHN`) | **yes** — diff sees every change wherever it was made |
| Deletions | need a separate marker | separate set, same drift | not representable | free: a deleted hunk is in the diff |
| Multiple colours | per-marker | per-range | per-key | one snapshot per colour; blame through the chain |
| Cost | trivial | trivial storage, fragile upkeep | small | ~one script copy per colour (≈100–200 KB); one line diff per snapshot per parse |

A is out on constraint 1. B is the Highland model, but it relies on a single
editor owning every change; Passage does not have that (desktop apps, text
editors, two tabs). C fails on screenplay text's many identical lines.
**D is the only option where marks cannot drift**, because marks are derived,
not stored.

## Recommended model (D)

**Sidecar.** Add `RevisionSets` to `SlateDocument` (`.slate/<script>.json`),
following H3.1's `Characters` precedent; the sidecar is already deleted with
the script. Do **not** reuse the name `Revisions` — that property already
holds Push/Pull and Lens runs (`SlateStore.cs:43`).

```
RevisionSets: [ { Colour: "Blue", Started: "2026-10-16", Baseline: "<full text>" }, ... ]
ForcedMarks:  [ { Text: "<line text>", Colour: "Blue" } ]   // Mark Revision only
```

**Derivation (server, at parse — never per keystroke).** Line diff (Myers,
LCS on lines) of each consecutive pair: `S0→S1`, …, `Sn→current`. Each current
line takes the colour of the latest set in which it was inserted or changed;
lines untouched since `S0` are unmarked. A deletion marks the next surviving
line (or the previous one at EOF). Blank lines and lines that do not print
(boneyard, `[[notes]]`, markers) are skipped. A small diff routine lives in
`Passage.Web/Services`; nothing in Core.

**Commands** (Revisions menu, mirroring Highland's set):

- *Start revision (colour)* — snapshot current text as the next baseline.
  The first start also snapshots `S0` (the locked white draft).
- *Revision mode off* — hides marks; storage untouched.
- *Clear revision* (selection) — rewrite the matching hunks of the latest
  baseline to equal current text, so those lines stop diffing.
- *Clear all* — latest baseline := current text.
- *Mark revision* (selection) — add text-keyed `ForcedMarks` for edits with no
  text change (rare; C's weakness is acceptable at this scale).
- *Go to next revision* — client-side over the pushed line list (like Ctrl+J).

## How marks survive edits

- **Between parses** a CodeMirror `StateField` holds the marked-line set and
  maps it through each transaction; any line touched while revision mode is on
  is marked immediately in the current colour. Zero round-trips while typing.
- **At each parse** the server recomputes from baselines and pushes the
  authoritative set (`setRevisionMarks`, same pattern as `setJumpTargets`).
  The client view converges; nothing is persisted from the client.
- **External edits** (desktop app, `vim`, a restored backup) are just more
  diff — the marks are right the next time the script opens.
- **Races:** baselines change only on explicit commands, so last-write-wins on
  the sidecar only matters if two tabs press *Start*/*Clear* at once.
- **Moves** show as delete + insert, so a moved scene is marked revised. That
  matches production practice.
- **Known cost:** a whitespace-only reformat marks every reformatted line.
  Trailing whitespace is normalised before diffing; nothing else is.

## Colour sequence

Default is the standard US production sequence (usually called the WGA
sequence; studios vary, so the list is data, not code):

White (original) → Blue → Pink → Yellow → Green → Goldenrod → Buff → Salmon →
Cherry → Double White → Double Blue → … → Double Cherry → Triple White …

Highland offers only Red/Green/Blue as screen highlight colours. This note
**departs from Highland on purpose**: the plan row asks for the production
sequence, and the colour name is what appears on pages. *Start revision*
offers the next colour in sequence; the user may pick another.

On screen, paper tints (Buff, Goldenrod, Pink) are illegible as text colour,
so marks render as a **left-gutter bar plus right-margin `*`** in an
accessible tint per colour, light and dark tokens in `wwwroot/css/app.css`; the colour
name shows on hover. Text colour is never changed.

## PDF stars

- An asterisk at the right margin (right edge of the 6" text column, ≈ 7.5"
  from the page's left edge), Courier 12, on **every printed line** of a
  revised paragraph. Word-level starring of only the wrapped lines that
  changed is a follow-up, not v1.
- Header: `<Colour> Rev. <date>` right-aligned, via the H4.1 header field.
- Shared-library change, per the standing decision: `PaginationOptions` gains
  `IReadOnlySet<int>? RevisedLines` (source line indices; null = today's
  output), and `LayoutLine` gains an optional `bool Revised = false` set when
  the source line is in the set. `ScreenplayPdfExporter` draws the star when
  `Revised` is true. Desktop callers pass no set, so their output is
  byte-identical — a test asserts this.
- *Print revision colour* (coloured text) is **not** proposed: tinted text
  on white paper is unreadable for half the sequence.
- *Revised pages only*: emit only pages containing a starred line. Without
  locked pages / A-pages ("Not going in" in HIGHLAND-PLAN) pages reflow, so
  this output is of limited production use. Proposed as a later row.

## Recommendation

Adopt **D: sidecar baseline snapshots, marks derived by line diff at parse,
mapped client-side between parses.** It is the only option that stays correct
when the file is edited outside the web app, keeps `.fountain` clean for the
desktop apps, and needs no offset bookkeeping.

Suggested split into sessions once approved:

1. Sidecar `RevisionSets` + diff/blame service + tests (external-edit,
   deletion, multi-colour, blank-line cases).
2. Editor: Revisions menu, gutter bar/star, `StateField` mapping, next-revision.
3. PDF stars + header, `PaginationOptions.RevisedLines`, desktop-unchanged test.
4. (Later) revised-pages-only; word-level stars; FDX `Revision` export.

## Questions for Adam

1. Accept the departure from Highland's Red/Green/Blue in favour of the
   production sequence?
2. Is `.slate/<script>.json` the right home, or should revisions get their own
   sidecar (`.revisions/`) since they are not a Writer's Tool?
3. Paragraph-level stars for v1 acceptable?
