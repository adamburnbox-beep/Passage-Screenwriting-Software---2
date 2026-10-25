# Story development tools — `Passage.Web`

The work queue for story-development features: story elements (storylines,
characters, locations), tagging, templates and the arc view on top of the
existing BOARD and BEATS tabs. Same scheme as `docs/HIGHLAND-PLAN.md`:
**this file is the source of truth for this work; update the row in the same
session as the change.**

Behaviour reference: Arc Studio Pro's story development tools, as described in
this plan. Nothing here copies Arc code, assets, feature names or branding,
and no row should either. Where this plan names a feature, the name is
Passage's own.

Status values: `todo` · `partial` · `done` · `verified` (done and checked in
the running app).

Paths are relative to the repository root. `Editor.razor` =
`Passage/Passage.Web/Components/Pages/Editor.razor` (6,088 lines: read ranges
only, and grep first, because ranges drift). `passage.js` =
`Passage/Passage.Web/wwwroot/js/passage.js`. `Analyzer` =
`Passage/Passage.Web/Services/DocumentAnalyzer.cs`. `Slate` =
`Passage/Passage.Web/Services/SlateStore.cs`. `Tests` =
`Passage/Passage.Tests/Program.cs`.

Line ranges as of `d0a0a72` (Beat Sheet merged, PR #40):

| What | Where |
| --- | --- |
| Sidebar tabs; CAST panel | `Editor.razor` 150–156; 195–306 |
| BOARD view | `Editor.razor` 411–496 |
| BEATS view (toolbar, tray, grid, scene track) | `Editor.razor` 497–574 |
| Beat Sheet handlers | `Editor.razor` 3025–3205 |
| Board drag / delete / add / edit handlers | `Editor.razor` 3207–3525 |
| Character rename | `Editor.razor` 3882–3960 |
| Outline filters (per-browser persistence pattern) | `Editor.razor` 4189 (`SetOutlineFiltersAsync`); `passage.js` ~1185 |
| Shape writer (`WriteShapeAsync`, `ScriptHasShape`) | `Editor.razor` 4845–5000 |
| `SaveSlate` | `Editor.razor` 4367 |
| `RenderBeat`, `RenderBeatEditor`, `RenderCardEditor`, `RenderBoardCard` | `Editor.razor` 5716–5847 |
| Jump palette (filter-in-browser pattern) | `passage.js` 280–370 |
| Beat drag / resize gesture | `passage.js` 789–985 |
| `SlateDocument`, `Beat`, `CharacterSettings` | `Slate` 12–76 |
| `SlateStore` (load, save, `PruneOrphans`) | `Slate` 398–551 |
| `BoardCard` / `BoardGroup` / `BoardLane`; board build | `Analyzer` 16–24; 504–572 |
| `BeatBoardText` (shared) | `Passage/Passage.Parser/BeatBoardText.cs` |
| `SplitScript.BuildLines` (the Shape's lines) | `Passage/Passage.Web/Services/SplitScript.cs` 55–80 |
| Element `Id` default | `Passage/Passage.Parser/ScreenplayModel.cs` 181 |
| `[[id:…]]` parsing | `Passage/Passage.Parser/FountainParser.cs` 897–940 |

---

## Standing decision A — storage

Everything this plan adds lives in the Slate sidecar
(`/data/.slate/<script>.json`), next to the Beat Sheet's `Beats`. The
`.fountain` file only ever holds structure and synopses: `#` / `##` / `###`
headings, scene headings, `=` lines, and the `[[id:…]]` markers it already
uses. Exports, page counts and the desktop apps' parse are untouched.

One new property on `SlateDocument`, so the story data is in one place:

```csharp
public StoryData Story { get; set; } = new();

public sealed class StoryData
{
    public List<Storyline> Storylines { get; set; } = new();       // S1.2
    public List<LocationSettings> Locations { get; set; } = new(); // S1.3, keyed by normalised name
    public List<ItemTags> Items { get; set; } = new();             // S1.1, S1.4: one per tagged card or beat
    public List<ArcPoint> Arc { get; set; } = new();               // S3.1
}

public sealed class ItemTags
{
    public Guid TargetId { get; set; }        // a Board card's [[id:…]] or a Beat.Id
    public string Kind { get; set; } = "";    // "card" or "beat"
    public string Label { get; set; } = "";   // last heading seen, so a detached item is recognisable
    public bool Key { get; set; }
    public List<Guid> Storylines { get; set; } = new();
    public List<string> Characters { get; set; } = new(); // bare cue names, as CharacterSettings
    public List<string> Locations { get; set; } = new();  // normalised names (decision C)
}
```

Each row adds only the part it needs.

- **Beat model unchanged, no migration.** Tags for beats and cards go in the
  same `Items` list, keyed by id. `Beat` keeps its own `Colour` (the beat's
  fill); storylines are drawn as bars on top. Adding properties with defaults
  needs no `SlateVersion` bump. That's the precedent from H3.1 (`Characters`)
  and PR #40 (`Beats`): an old sidecar loads with an empty `Story`. Note that
  `SlateStore.Load` never reads `SlateVersion` at all today, so a bump would
  do nothing until something checks it. Bump it only when a row renames or
  removes a stored property, and add the check in that row.
- **Deleting a beat cascades** its `ItemTags` and arc points: the delete is
  explicit, confirmed, and the beat is sidecar-owned. Deleting a card does
  not (decision B).
- **Character rename** (`Editor.razor` 3882–3960) must also rename the name
  in every `ItemTags.Characters`. That's a change to `RenameCharacterAsync`,
  made in S1.4.
- **Library-wide data** (saved structure templates in S2.2) goes in
  `.slate/templates.json`, not in a script's sidecar. `PruneOrphans`
  currently spares only `ideation.json`, so S2.2 must add `templates.json` to
  that exception with a test, or the first load deletes it.

**Settled (2026-10-25): the sidecar now also holds permanent story-planning
data.** `docs/SLATE-PLAN.md` decision B called it "a staging area, not a
second home for the story". The Beat Sheet and Cast colours had already broken
that, and Adam accepted the same for this plan. S1.1 adds one sentence to
SLATE-PLAN decision B saying so.

**A Slate sidecar revamp is coming** as its own dedicated session (not
scheduled yet). To keep that cheap, every row keeps its data inside the one
`StoryData` object and adds no new top-level `SlateDocument` properties. If the
revamp lands before a row is started, re-read decision A against it first and
flag any mismatch rather than guessing (rule 7).

**Risk that grows with this plan.** `PruneOrphans` deletes the sidecar of
any script that isn't in the library on load (it reports this, as rule 12
requires). A script renamed outside the app loses its storylines and tags.
That was tolerable for worksheet working-out. It's less so for a fully tagged
board. No row changes this. Flag it again if it bites.

---

## Standing decision A2 — every Board card needs a stable id (S1.1)

**Verified: they don't have one today.** `ScreenplayElement.Id` defaults to
`Guid.NewGuid()` (`ScreenplayModel.cs` 181). The parser overwrites it only
when the line carries `[[id:…]]` (`FountainParser.cs` 897–940). Cards made by
the Board's `+` or saved through its editor get a marker
(`BeatBoardText.BuildCardLines`). Every hand-typed heading gets a **new
random id on every parse**, so sidecar data keyed by it would detach on the
next keystroke. The comment at `Editor.razor` ~3500 already says the web
board "has no stable card identity".

The fix is the first row:

- **Stamp lazily.** A card gets `[[id:…]]` appended to its heading line the
  first time something attaches sidecar data to it (star, tag, arc point).
  Use one `passage.replaceInLine` edit, which is one undo step. Unused cards
  stay as the writer typed them. Stamping every card on Board open was
  rejected: it would rewrite a script nobody asked to change.
- **Append to the line; don't rebuild it.** Do not stamp through
  `BuildCardLines`: `BoardCard.Heading` has the scene number already stripped
  (H1.3), so rebuilding the line would drop `#12#`. The parser removes
  `[[id:…]]` wherever it sits and then reads a trailing `#…#`, so
  `INT. HOUSE - DAY #12# [[id:…]]` round-trips.
  (The same stripping means the Board's existing **card editor already drops a
  scene number on Save**: `SaveCardAsync` rebuilds the heading from
  `card.Heading`. That's an existing bug, found by reading
  `ScreenplayModel.cs` 219–221, not yet seen in the running app. It's out of
  scope here, but S1.1 must not repeat it.)
- **Duplicate ids.** Copying a heading copies its id, and both copies parse to
  the same Guid (`GetCardLineRange` then finds only the first). The analyzer
  flags every occurrence after the first. Those count as unstamped, so
  attaching to one gives it a fresh id, and the status bar says so.
- **Shared-library impact: none.** `[[id:…]]` is the desktop apps' own
  convention, and stamping is a web-side text edit. The parser and
  `BeatBoardText` are unchanged.

---

## Standing decision B — when a card leaves the script

Sidecar data whose card id no longer parses is **kept and shown, never
silently dropped** (rule 12).

- **Detached** = an `ItemTags` (or arc point) of kind `card` whose
  `TargetId` matches no element in the current parse.
- **Where it shows:** a `Detached (n)` button on the BOARD toolbar, visible
  only when n > 0. It lists each item's `Label` and what it carried (key,
  storylines, characters, locations). Each item has a **Discard** button, and
  the list has **Discard all**, which confirms and states the count.
- **Coming back is automatic.** If the id parses again (CodeMirror undo, a
  paste, an outside edit), the item re-attaches with nothing to do. This is
  the main reason not to prune.
- **The edge case.** The parse briefly loses a heading while the user is
  mid-edit (for example, deleting the `.` of a forced heading). The item
  detaches for that moment and comes back. Nothing is written, because
  detachment is computed on each parse and never stored.
- **`Label` refreshes** whenever the item is attached and its heading has
  changed. It is written only on saves that happen anyway, never one save per
  parse.

---

## Standing decision C — location names

Locations are **derived from scene headings on every parse**. Only their
colours are stored, keyed by normalised name.

To normalise a heading (uppercase, whitespace collapsed):

1. Strip the scene number (already done by the parser) and `[[id:…]]`.
2. Strip a leading prefix: `INT./EXT.`, `EXT./INT.`, `INT/EXT`, `I/E.`, `I/E`,
   `INT.`, `EXT.`, `EST.`, `INT`, `EXT` (with or without the period), or a
   forced `.`.
3. Split on ` - ` (and ` — `). Repeatedly drop the **last** segment while it
   is a time of day: `DAY`, `NIGHT`, `MORNING`, `AFTERNOON`, `EVENING`,
   `DAWN`, `DUSK`, `SUNRISE`, `SUNSET`, `LATER`, `MOMENTS LATER`,
   `CONTINUOUS`, `SAME`, `SAME TIME`, `MAGIC HOUR`. Anything else stays, so
   `HOUSE - KITCHEN - NIGHT` → `HOUSE - KITCHEN`.
4. Empty after stripping → not a location (for example, a forced
   `.FLASHBACK`).

Results:

- `INT. KITCHEN - DAY`, `INT. KITCHEN - NIGHT` and `EXT. KITCHEN - DAY` are
  one location, `KITCHEN`. INT/EXT is deliberately ignored: the reference
  groups by place, and splitting by side would double the list.
- **Sub-locations are separate** (`HOUSE - KITCHEN` ≠ `HOUSE`). Grouping by
  parent is not planned.
- **Board-created scenes** (`. NEW SCENE`) normalise to `NEW SCENE`. That's
  harmless: it is a location until the writer renames the scene.
- **Rule 7 note:** `Characters.BareName` is the precedent for normalising a
  name before using it as a key. Locations follow the same idea in a new
  `Locations.Normalise`, and don't reuse it.

Scene cards are **implicitly tagged** with their own heading's location, and
nothing is stored for that. An explicit location tag (S1.4) is for beats, acts
and sequences, or for a scene that also uses a second place.

---

## Standing decision D — client vs server (CLAUDE.md trap 4)

| Interaction | Where | Pattern to reuse |
| --- | --- | --- |
| Beat drag / resize | JS (exists) | `passage.js` 789–985 |
| Inbox beat → Board drop (S2.1) | JS gesture, one call on release | the beat gesture, plus `dropIsAfter`'s midpoint rule |
| Tag popover search and keyboard (S1.4) | JS, one call per Enter | the jump palette (`passage.js` 280–370, `setJumpTargets`) |
| Arc point drag (S3.1) | JS, one call on release | the beat gesture's pointer capture |
| Selection highlight (S1.4) | JS class toggle; server told on `#` / Enter only | — |
| Filters, key toggle, expand/collapse (S1.5, S2.4) | Server render on click (not per keystroke), state per browser | outline filters (H4.3) |
| Saves, script writes, exports | Server | `SaveSlate`, `passage.replaceInLine` / `insertLinesAt` |

**Existing conflict, out of scope:** the Board's card drag uses Blazor
`@ondragover` (`RenderBoardCard`), which sends a server event on every
`dragover`, many per second. That is exactly what trap 4 warns about. No row
here fixes it. S2.1 must not copy it, and the inbox drop uses the pointer
gesture instead. Moving the existing card drag to JS would be its own row if
Adam wants it.

---

## Standing decision E — one inbox (S2.1)

- **One list.** The inbox is the beats with `Page == 0`, the existing BEATS
  tray. BOARD shows the same list as a strip above the lanes. There is no
  second store.
- **Placing on the grid** (exists) gives the beat a page. It leaves the
  inbox and stays a beat.
- **Placing on the Board converts the beat into a card.** It writes
  `. TITLE [[id:new]]` plus its description as `=` lines (through
  `BuildCardLines("Scene", …)`, same as `+`) at the drop point, moves its
  `ItemTags` to the new card id (kind `card`), and deletes the beat. The
  script is now the truth, as SLATE-PLAN decision B says of anything
  promoted.
- **A beat is never in both views.** Once it is a scene, the BEATS script
  track already shows where it lands, so a second, planned position would
  be a second truth that drifts. A beat already on the grid can also be
  sent to the Board (a "Send to Board" button in the beat editor). The
  confirm says its page plan is dropped and the script track takes over.
- **What a dropped beat becomes:** a Scene, which is printable, matching the
  Board's `+`. The card editor can change it to a Section.
- **Undo across the two stores is not atomic.** CodeMirror undo removes the
  card, but it doesn't restore the beat. The card's tags then show as
  Detached (decision B), so nothing is lost silently. Say this in the row's
  PR.

Alternative, if Adam prefers it: link instead of converting (`Beat.CardId`).
The beat keeps its page and shows the card's text read-only. That keeps a
plan-vs-actual page view, at the cost of a `Beat` model change and two
displays of one text. Not recommended (rule 2).

---

## Standing decision F — structure templates and the Shape writer (S2.2)

There is one writer. Both the Shape tool and the templates produce lines with
`BeatBoardText.BuildCardLines` and insert them with `passage.insertLinesAt`.

- `SplitScript.BuildLines` already writes acts and sequences through
  `BuildCardLines`. S2.2 pulls the first-time insert branch of
  `WriteShapeAsync` (`Editor.razor` ~4877–4915: blank-script handling, insert
  at end, mirror `_content`, re-analyse) into one `InsertStructureAsync(lines,
  description)`. The Shape tool and the template menu both call it.
  `WriteShapeAsync`'s update path (filling brackets by slot label) is
  unchanged.
- **Passage's own 8-sequence template is the Shape with empty values**
  (`SplitScript.BuildLines(new SplitRun())`). After it is applied,
  `ScriptHasShape` is true and the Shape tool fills it in. Don't keep a second
  copy of the eight sequences.
- **Templates never write anything that prints.** They write only `#`, `##`,
  `###` and `=` lines, never scene headings, so applying one adds no pages.
  Named beats (for example a 15-beat sheet's "Catalyst") become `###`
  Section cards with a `=` prompt line.
- **Shared library: no change.** Template data and the new insert method live
  in `Passage.Web` (`Services/StructureTemplates.cs`).
- **When templates are offered:** only when the script has no `#` or `##`
  heading. The template is inserted at the end, as the Shape is. If the script
  already has structure, the menu item explains why it's unavailable rather
  than merging.
- **Seeding beats: recommended, as an opt-in checkbox, off by default.** Each
  template beat with a conventional position gets a Beat Sheet beat. Its page
  is scaled from a 110-page reference to `BeatSheetPages`. Beats whose title
  already exists are skipped and counted. Why: a fresh template has no scenes,
  so the BEATS grid shows nothing, not even act breaks. Seeded beats give it
  shape. Why opt-in: the beats duplicate the Board's `###` cards as planning
  until scenes are written.
- **Naming (settled):** the built-ins are "8 sequences / 7 turns (Passage)",
  "Paradigm (Syd Field)", "Save the Cat (Blake Snyder)" and "Five-act TV".
  Each template's description credits its source. Passage is pointing writers
  to the original work, not claiming the structure as its own.
---

## Standing decision G — selecting a card or beat (S1.4)

**Settled (2026-10-25): a single click selects.** The tag popover opens with
`#` on "the selected card or beat". Neither view had selection, and the
gesture conflicted with two existing clicks:

- A **Board card click jumps to the script** (`JumpToLineAsync`) and leaves
  the Board.
- A **beat click opens its editor**, which focuses a text input, so `#` would
  type a `#`.

**What changes:** a single click **selects** (ring highlight, done in JS),
and `#` opens the popover. The arrow keys move the selection within a view.
On a card, **Enter / double-click** jumps to the script, and the existing ✎
still edits. On a beat, **Enter / double-click** edits. Double-click on an
empty page still adds a beat. Escape clears the selection. S1.4's PR must say
plainly that both clicks changed. The tag-button fallback was rejected.
---

## Tier S1 — Story elements (foundation)

Order: identity first (nothing can be keyed until ids are stable), then the
elements that tags point at, then tagging, then filtering what's tagged.
Key beats move to the front (brief item 5) because a star is the smallest
possible attachment. It proves stamping, detaching and re-attaching before
the bigger rows rely on them.

### S1.1 Stable card ids + key beats — `done`

- **Read:** Standing decisions A, A2, B. `Editor.razor` BOARD view 411–496,
  BEATS view 497–574, `RenderBeat` / `RenderBoardCard` 5716–5847, Beat Sheet
  handlers 3025–3205. `Analyzer` 16–24 and 504–572. `FountainParser.cs`
  897–940 (read only). `Slate` 12–76.
- **Do:** Add `StoryData` with `Items` only. Add `EnsureCardIdAsync(card)` per
  A2, plus duplicate-id detection in the analyzer, surfaced as a status
  message. Add a ☆/★ toggle on scene, section, act and sequence cards and on
  beats (grid and tray). A star on a card stamps the card first. Add a
  "Key only" toggle on the BOARD toolbar (new; it holds `Detached (n)` too)
  and on the BEATS toolbar; it dims non-key items (S1.5 folds it into the
  filter bar). Add the Detached list per decision B. Add one sentence to SLATE-PLAN
  decision B: the sidecar also holds permanent story-planning data (Beat
  Sheet, Cast colours, story elements), see STORY-PLAN decision A.
- **Shared-library impact:** none. Parser and `BeatBoardText` are read only.
- **Success (running app):**
  - Star a hand-typed scene: its heading line gains `[[id:…]]` as one undo
    step, and the star survives typing elsewhere and a reload.
  - Delete that scene in SCRIPT: BOARD shows `Detached (1)` with its heading.
    Ctrl+Z in SCRIPT and the star comes back on the card.
  - Copy-paste a stamped heading and star the copy: the copy gets a fresh id,
    and the status bar says so.
  - Starring a beat writes nothing to the `.fountain` file (SHA-256 is
    unchanged).
  - Key only dims, and never hides.
- **Tests:**
  - Stamping a line with and without `#12#` round-trips both the id and the
    number.
  - A duplicate id is flagged on the second occurrence only.
  - A detached item is computed and never removed by load or save.
  - Deleting a beat removes its `ItemTags`.
  - An old sidecar with no `Story` loads.
- **Budget:** ~90k tokens.
- **Landed:** `StoryData`/`ItemTags` (`Items` only) in `SlateStore.cs`; pure
  rules in `Services/StoryItems.cs` (`StampLine`, `DuplicateIdLines`,
  `Detached`, `SetKey`, `RemoveBeat`, `RefreshLabels`). `DocumentAnalysis`
  gains `DuplicateIdLines` (1-based lines of repeat ids, second occurrence
  on). `Editor.razor`: `EnsureCardIdAsync` (one `replaceInLine`, appended, so
  `#12#` survives), stars on act/sequence heads, cards and beats, a new BOARD
  toolbar (Key only, `Detached (n)`), and "Key only" on BEATS (one shared,
  unsaved flag; dims, never hides). The duplicate message appears when you
  star the copy, not on every parse. Beat delete cascades its tags. Checked in
  the running app: star stamps as one undo step, Ctrl+Z after deleting the
  scene re-attaches, a pasted copy gets a fresh id and the status message,
  the star survives a reload, starring a beat leaves the `.fountain` SHA-256
  unchanged. Five tests added.

### S1.2 Storylines and the ELEMENTS tab — `done`

- **Read:** Standing decision A. `Editor.razor` sidebar tabs 150–156, CAST
  panel 195–306 (the palette and swatch pattern to match), `SaveSlate` 4367.
  `Slate` 12–76. `Passage/Passage.Web/Services/Characters.cs` 31–45
  (`Palette`).
- **Recommendation: a new fifth sidebar tab, ELEMENTS.** It is not a section
  of CAST. CAST is a character list with a balance report. Storylines and
  locations aren't characters, and nesting them would push CAST's report
  further down. ELEMENTS shows **Storylines**, **Locations** (S1.3), and one
  line, "Characters: colours are set in CAST →", that switches tab. There is
  no second character list (brief item 2).
- **Do:** Storyline = `Id`, `Name`, `Colour` (Cast palette), `Description`.
  Add, rename, recolour, describe and delete. Deleting a storyline that tags
  items confirms with the count, then removes it from those items. Each item's
  order is the list order (drag-reorder isn't needed). Starter names are
  offered as placeholder text only (A Story, B Story, Runner). Nothing is
  auto-created.
- **Shared-library impact:** none.
- **Success:** create "A Story" in red with a description. It survives a
  reload, and the `.fountain` file is unchanged. Deleting a used storyline
  states the number of items it will untag. (Untagging is testable once S1.4
  lands; the count logic is tested now.)
- **Tests:** round trip. Delete removes the storyline id from every
  `ItemTags.Storylines` and reports the count. An unknown palette colour is
  rejected, as in `SaveBeatEdit`.
- **Budget:** ~60k.
- **Landed:** `Storyline` and `StoryData.Storylines` in `SlateStore.cs`;
  `ItemTags.Storylines` (ids) added for S1.4. Pure rules in `StoryItems.cs`
  (`StorylineColour`, `StorylineUseCount`, `DeleteStoryline`); `SetKey` now
  keeps an item that still carries storylines. `Editor.razor`: fifth sidebar
  tab ELEMENTS (Storylines list with add, edit by Save/Cancel, counted
  delete confirm; Locations placeholder for S1.3; "Characters: colours are
  set in CAST →" switches tab). No second character list. Sidebar tab
  letter-spacing tightened so five tabs fit. Checked in the running app: "A
  Story" in red with a description survives a reload and the `.fountain` is
  unchanged. The counted untag message needs tagged items, so it is covered by
  the unit test until S1.4.

### S1.3 Locations — `done`

- **Read:** Standing decision C. `ScreenplayModel.cs` 397–410
  (`SceneHeadingElement`). `Analyzer` board build 504–572. The ELEMENTS tab
  from S1.2.
- **Do:** `Locations.Normalise(heading)` per C. The analyzer exposes a derived
  location list (name, scene count, first line). ELEMENTS → Locations lists
  them by first appearance. Clicking one expands its scenes, and clicking a
  scene jumps to it. Each location gets an optional colour (`LocationSettings`,
  only kept when set, like `CharacterSettings`). A colour whose location no
  longer appears is listed greyed under "Not in the script", with Discard
  (decision B in spirit).
- **Shared-library impact:** none (web-only service).
- **Success:**
  - `INT. KITCHEN - DAY`, `INT. KITCHEN - NIGHT` and `EXT. KITCHEN - DAY` show
    as one `KITCHEN` (3 scenes).
  - `HOUSE - KITCHEN - NIGHT` is its own entry.
  - Typing in SCRIPT doesn't stall (the derivation runs inside the existing
    parse).
- **Tests:** a normalisation table covering every prefix and time-of-day in
  C, a multi-dash sub-location, a forced heading, an empty result, and a
  heading with `#12#` and `[[id:…]]`.
- **Budget:** ~60k.
- **Landed:** `Services/Locations.cs`: `Normalise` per decision C and `Build`
  (name, first line, scenes in script order). `Normalise` also strips `#…#`
  and `[[id:…]]` itself; a prefix only counts at a word boundary (`EXTRA` is
  not `EXT`); a bare time of day normalises to empty. `DocumentAnalysis`
  gains `Locations`, built inside the existing parse. `LocationSettings` and
  `StoryData.Locations` in `SlateStore.cs`, kept only when a colour is set.
  `Editor.razor` ELEMENTS: Locations by first appearance with scene count,
  click to expand and jump to a scene, colour swatch, and a greyed "Not in the
  script" list with Discard. Checked in the running app: `KITCHEN` (3) and
  `HOUSE - KITCHEN` as separate entries, jump moves the caret, a coloured
  location disappears into "Not in the script" and Discard clears it. No
  per-keystroke work was added (list derived on parse only).

### S1.4 Tagging: selection, `#` popover, bars and dots — `done`

- **Read:** Standing decisions A, A2, D, G. `RenderBeat` / `RenderBoardCard`
  5716–5847. Jump palette `passage.js` 280–370. Character rename
  `Editor.razor` 3882–3960. `app.css`: grep `.board-card` and `.beat-card`
  only.
- **Do:**
  - Selection per G.
  - One popover, built once in `passage.js` and used on both views. It lists
    storylines, characters (from the parse) and locations (derived), each with
    its colour, pushed by a `setTagTargets` like `setJumpTargets`. Typing
    filters in the browser. Enter toggles the highlighted element on the
    selected item through one `[JSInvokable] OnTagToggled(targetId, kind,
    key)`. Esc closes.
  - The server stamps the card (A2) and saves.
  - Display: storylines as stacked colour bars on the item's left edge, and
    characters as colour dots (a character with no Cast colour gets a
    neutral dot and a tooltip). An explicit location shows as a small chip.
    A scene card's implicit location isn't drawn; its heading already says it.
  - Rename in CAST also renames tags (A).
- **Shared-library impact:** none.
- **Success:**
  - Select a card, press `#`, type `a st`, Enter: a red bar appears. The
    WebSocket shows no frames while typing in the popover, and one on Enter.
  - The same popover works on a grid beat and a tray beat.
  - Renaming a tagged character in CAST keeps the dot.
  - Tags survive a reload, and the `.fountain` file changes only by the id
    stamp.
- **Tests:**
  - Toggling adds, then removes.
  - Tagging an unstamped card stamps it once.
  - A rename rewrites tags and merges duplicates.
  - Characters are stored as bare names (`BOB (V.O.)` → `BOB`).
- **Budget:** ~130k. This is the largest row: stop and ask at 150k (rule 6)
  rather than finishing display or rename in a rush. If it trends over, split
  display into its own PR.
- **Landed:** `ItemTags` gains `Characters` and `Locations` (bare cue names;
  normalised locations) and `IsEmpty`. `StoryItems`: `ToggleTag`, `TagKeys`,
  `RenameCharacter` (rewrites and merges), `NeedsStamp`. `passage.js`: JS-only
  selection (ring class, arrow keys, Enter / double-click, Esc, re-applied by a
  `MutationObserver` after Blazor re-renders) and one popover used on BOARD
  cards, grid beats and tray beats, fed by `setTagTargets(targets, items)`.
  Server calls: `OnTagToggled(targetId, kind, key)` on Enter only (returns the
  id, stamping an unstamped card first), and `OnBeatActivated(id)` for
  Enter / double-click on a beat (a card reuses `OnJumpTarget`). Display:
  storyline bars, character dots (neutral ring + tooltip with no Cast colour),
  location chips. A card's `data-item-id` is empty until its heading carries
  `[[id:…]]`, because the parser gives unmarked headings a random id on every
  parse; the client addresses those as `line:N`. `RenameCharacterAsync` renames
  tags. **Both clicks changed** (decision G): a card click selects instead of
  jumping, a beat click selects instead of opening its editor. Checked in the
  running app: no WebSocket frames while typing in the popover; Enter on an
  already-stamped item is one call plus Blazor's render acks (3 frames);
  the first tag on an unstamped card adds the stamp's edit round trips (12).
  Tags persist across reload; the `.fountain` changed only by the stamp (and
  the rename I did). Known gap: undoing a CAST rename restores the script and
  colour but not the renamed tags, because tags are moved, not copied.

### S1.5 Element filters — `done`

- **Read:** Standing decisions B, C, D. Outline filters: `Editor.razor` 158–166
  and `SetOutlineFiltersAsync` 4189, `Services/OutlineFilters.cs`, and
  `passage.js` ~1185 (persistence). The BOARD and BEATS toolbars from S1.1.
- **Do:** A filter bar on BOARD and on BEATS with storylines, characters,
  locations and Key only (folded in from S1.1). Choices combine: OR within a
  kind, AND across kinds, with the rule stated in a tooltip. Items that don't
  match **dim** (`.dimmed`, reduced opacity, still clickable and
  draggable). They never hide, so layout and beat stacking don't jump.
- **Matching:**
  - A scene card matches a location by its own heading (implicit, per C) or
    by an explicit tag.
  - A scene card matches a character by an explicit tag **or** by that
    character speaking in the scene (derived from the parse, nothing
    stored; settled 2026-10-25). Why: tagging every speaking character by hand
    is busywork the parse already does. Explicit tags stay for characters who
    are present but silent, and for beats, acts and sequences. Speaking is
    matched on bare cue names (`Characters.BareName`), so `BOB (V.O.)` counts
    as `BOB`. Derived characters aren't drawn as dots; only explicit tags
    are, so the card stays readable.
  - An act or sequence lane matches if any card in it matches.
- **Persistence:** per browser, as outline filters, one set per view.
- **Shared-library impact:** none.
- **Success:** pick "B Story" + "KITCHEN" on BOARD: only B Story scenes in the
  kitchen are at full strength. The same filter on BEATS dims beats likewise,
  and the scene track follows. The filters survive a reload, and clearing them
  restores everything.
- **Tests:** a match table (OR within a kind, AND across kinds, implicit
  location, derived speaker (including a `(V.O.)` cue), a lane matching via a child, an
  empty filter matching everything).
- **Budget:** ~70k.
- **Landed:** `Services/StoryFilter.cs`: the `StoryFilter` record (storylines,
  characters, locations, `KeyOnly`; OR within a kind, AND across kinds, empty
  matches everything), `LaneMatches`, and `SceneSpeakers` (bare cue names per
  scene, keyed by heading line). `Editor.razor`: `RenderFilterBar` on both
  toolbars (three `<details>` menus, Key only, Clear; the tooltip states the
  rule) replaces the S1.1 `_keyOnly` toggle. `_boardFilter` and `_beatFilter`
  are separate and persist in the session key (`passage.setStoryFilter`,
  `BoardFilter` / `BeatFilter` on the snapshot). Dimming covers cards, act and
  sequence heads (a lane matches via itself or any card in it), beats, and the
  BEATS scene track. A scene's own heading counts as its location; a speaker
  counts as its character. Beats match on explicit tags only. Server render on
  click, no per-keystroke handlers. Tests: `Test Story Filter Match Table`,
  `Test Story Filter Scene Speakers`. Tried in the app on BOARD (ANN, then
  ANN + KITCHEN, survived a reload) and BEATS (BOB dims the other scenes in
  the track). Not tried in the app: a storyline pick, since no tagged scenes
  were set up; that path is covered by the match table only.

---

## Tier S2 — Board workflow

### S2.1 One inbox for BOARD and BEATS — `done`

- **Read:** Standing decisions D, E. BEATS tray `Editor.razor` 528–538. Beat
  handlers 3025–3205. Board add and edit handlers 3361–3480 (insert and
  `BuildCardLines` use). Beat gesture `passage.js` 789–985, and
  `dropIsAfter`.
- **Do:** Show the tray's list as an inbox strip on BOARD (double-click it to
  add a beat, as on BEATS). Add the pointer gesture from the strip onto a
  card: before or after by vertical midpoint, one call on release
  (`OnInboxBeatPlaced(beatId, targetCardId, after)`). The insert point is the
  end of the target's whole block when after, or its first line when before,
  the same ranges `TryPlanMove` uses. Convert per E, moving `ItemTags`. Add
  "Send to Board" in the beat editor, which confirms when the beat has a page.
- **Shared-library impact:** none. `BuildCardLines` and `GetCardLineRange`
  are called, not changed.
- **Success:**
  - A beat jotted on BEATS shows in BOARD's inbox.
  - Dragging it after a scene writes `. TITLE [[id:…]]` and its `=` lines
    there, keeps its tags and star, and it's gone from both inboxes.
  - No server frames while dragging.
  - Ctrl+Z in SCRIPT removes the card, and its tags show as Detached.
- **Tests:**
  - Conversion lines for a title with and without a description.
  - Tag transfer re-keys `TargetId` and `Kind`.
  - The insert line index for before and after on an act, a sequence and a
    scene.
  - A beat that is no longer in the sidecar is reported, not placed.
- **Budget:** ~100k.
- **Landed:** `Services/BeatInbox.cs`: `SceneLines` (via `BuildCardLines`; an
  untitled beat gets `UNTITLED BEAT`), `TryPlan` (before = the target's first
  line, after = the line past its whole block, `GetCardLineRange` with the
  nested block, so after an act or sequence lands past everything inside it;
  a blank line pads whichever side touches text), `TextAtEnd`, `MoveTags`
  (re-keys `TargetId`, `Kind`, `Label`, keeps star and tags). `Editor.razor`:
  the unassigned beats show as `.board-inbox` above the lanes on BOARD (same
  `RenderBeat`, double-click adds, the beat editor opens there);
  `data-drop-id` on stamped cards and on act and sequence heads;
  `OnInboxBeatPlaced(beatId, targetId, after)`; a **Send to Board** button in
  the beat editor (end of the script; confirms when the beat has a page, and
  sends the editor's unsaved text). `passage.js` `moveInboxBeat`: pointer
  gesture on `document`, class toggles only, one call on release. It does not
  capture the pointer, so the beat's star still takes clicks. A card with no
  `[[id:…]]` marker is not a drop target (nothing highlights). Tests: `Test
  Inbox Conversion Lines`, `Tag Transfer`, `Insert Points`, `Missing Target
  Is Reported`. Tried in the app: a beat jotted on BEATS shows in BOARD's
  inbox; dragged after a scene (and before a sequence), it wrote the scene
  and its `=` line, a star carried across, and it left the inbox; one
  WebSocket frame for the gesture; Ctrl+Z removed the card and `Detached (1)`
  appeared; Send to Board from the editor. Not tried: the page-plan confirm
  for a beat on the grid, and a storyline or character tag carried across
  (the tag move is covered by the unit test only). Seen, not changed: on the
  BEATS tray the star didn't take a click when tried, probably because
  `moveBeat` captures the pointer on press (not confirmed). Undo is not atomic across the two stores (decision E), as
  planned.

### S2.2 Structure templates — `done`

- **Read:** Standing decision F. Shape writer `Editor.razor` 4845–5000.
  `SplitScript.cs` 13–80. `BeatBoardText.BuildCardLines`. `Slate` 398–551
  (`PruneOrphans`, `IdeationFile`). Beat model `Slate` 56–66.
- **Do:**
  - `InsertStructureAsync`, extracted from `WriteShapeAsync` with no change
    in behaviour.
  - `StructureTemplates` with four built-ins: Passage's (from
    `SplitScript.BuildLines`), Paradigm (Syd Field), Save the Cat (Blake Snyder) and five-act
    TV. Each lists its `#` / `##` / `###` / `=` lines and, optionally, the
    page each beat sits at on a 110-page reference.
  - A BOARD menu, "Start from a template…", enabled only when there is no
    `#` or `##` (F), with a "Seed the Beat Sheet too" checkbox.
  - "Save this structure as a template" captures the current `#` / `##` /
    `###` headings and `=` lines (scenes are excluded) into
    `.slate/templates.json`. Custom templates can be renamed and deleted.
  - `PruneOrphans` spares `templates.json`.
- **Shared-library impact:** none.
- **Success:**
  - On an empty script, each built-in writes its skeleton. PREVIEW page count
    stays 0 or 1 (nothing prints), and BOARD shows the lanes.
  - Passage's template, then the Shape tool, fills the turns.
  - With "seed" ticked on a 90-page target, the beats land at scaled pages.
    A second run says "n beats already there".
  - A saved custom template reappears in another script.
  - The Shape tool's own "Write the shape" still works as before.
- **Tests:**
  - Each built-in's lines parse to the expected act, sequence and section
    count with zero printable elements.
  - Passage's template equals `SplitScript.BuildLines` of an empty run.
  - Page scaling and clamping.
  - The duplicate-title skip.
  - `PruneOrphans` keeps `templates.json`.
  - A save-as-template round trip excludes scenes.
- **Budget:** ~90k.
- **Landed:** `Services/StructureTemplates.cs` (`TemplateNode`,
  `StructureTemplate`, `TemplatesDocument`; `BuildLines`, `ReadNodes`,
  `HasStructure`, `ScalePage`, `SeedBeats`, `FromScript`). Passage's built-in
  is `ReadNodes(SplitScript.BuildLines(new SplitRun()))`, so there is one copy
  of the eight sequences; its sequences carry pages, its acts don't. The other
  three use `#` acts with `###` named beats and no `##`; the Board shows those
  as lanes of BEAT cards. `Editor.razor`: `InsertStructureAsync(lines,
  description)` extracted from `WriteShapeAsync` (same messages), a "Start from
  a template…" button on BOARD opening a dialog (built-ins, custom with
  rename and delete, "Seed the Beat Sheet too", "Save this structure as a
  template"). `SlateStore`: `LoadTemplates`/`SaveTemplates`, and `PruneOrphans`
  spares `templates.json`. Also `passage.focusEditor` now refreshes
  CodeMirror: lines inserted from BOARD (template or inbox) showed a blank
  editor on returning to SCRIPT until it was clicked. Tests: `Structure
  Templates Print Nothing`, `Passage Is The Shape`, `Page Scaling`, `Seed Skips
  Duplicates`, `Save Round Trip`, and `SlateStore Keeps Templates File`. Tried
  in the app: Save the Cat with seed ticked wrote the lanes and 15 beats, and
  PREVIEW stayed at 1 page; the menu then said why it was unavailable;
  a saved custom template was listed and applied after a reload; Passage's
  template then the Shape tool filled Starts and Ends; the Shape tool alone
  still wrote the shape first time and filled a turn on update. Not tried in
  the app: seeding on a 90-page target and the "already there" message (unit
  tests only); the template picks up no `Page` from a saved structure, so a
  custom template never seeds. Seen, not changed: a script named
  "templates" (or "ideation") has a sidecar path that is the library-wide
  file.

### S2.3 Board beside the script — `done`

- **Read:** `Editor.razor` 60–75 (main view tabs), 405–412 (editor surface,
  hidden when not SCRIPT), `SetMainView` 3010, `UpdateCaretElementLabel` 2987.
  `app.css`: grep `.main-area` and `.editor-surface` only.
- **Do:** A "Beside" toggle on SCRIPT opens a right-hand pane with a compact
  list: lanes and groups as headers, cards as one line each with
  bars/dots/star. A switch in the pane shows Beats instead: beats in page
  order, from the Beat Sheet's "beats panel inside the script editor"
  follow-up. Clicking a card calls `scrollToLine` and leaves SCRIPT showing.
  The card under the caret is highlighted, driven by the existing caret
  label update, not per keystroke. The pane width is draggable in JS and
  stored per browser.
- **Shared-library impact:** none.
- **Success:** clicking the third card puts the caret on its heading, and the
  pane stays open. Moving the caret into another scene moves the highlight.
  Typing speed is unchanged (pane re-render happens only when the parse
  completes). The pane state survives a reload.
- **Tests:** the compact list order matches the Board's document order. The
  caret-to-card lookup is checked for a line inside a scene, on a heading, and
  before any card.
- **Budget:** ~80k.
- **Landed:** A "Beside" button on SCRIPT's top bar opens a right-hand pane
  with a Board/Beats switch. `Services/BesideList.cs` holds the pure lists:
  Board rows in the Board's order (lane, group, cards with nesting) and Beats
  in page order, unassigned first. Clicking a row calls `passage.scrollToLine`
  and stays on SCRIPT. The current row comes from the caret line the caret
  callbacks already keep; rows are cached per parse (`_besideRowsFor`). Width
  is dragged in `passage.js` (`--beside-width`); open/mode/width ride the
  session key. Beats have no script line, so a beat jumps to the first scene
  on its page (else the last scene before it); the current beat is the first
  one covering the caret scene's page. Unassigned beats are listed but inert.
  Checked in the running app: third card click, highlight move, reload.
  Not done: the pane lists Board rows even when a filter dims them (dimmed
  the same way); location chips are hidden in the compact rows.

### S2.4 Expand all / collapse all — `done`

- **Read:** BOARD view 411–496. `_collapsedLanes` / `_collapsedGroups`
  `Editor.razor` ~3500–3530. Outline-filter persistence pattern.
- **Do:** BOARD toolbar buttons: "Descriptions: show/hide" (synopses on
  lanes, groups and cards) and "Expand all / Collapse all" (lanes and groups).
  The description toggle is a CSS class on `.board-surface`, persisted per
  browser.
- **Shared-library impact:** none.
- **Success:** hiding descriptions shrinks every card to its heading and
  survives a reload. Collapse all and then expand all restores every lane and
  group.
- **Tests:** collapse all collects every lane and group key (including the
  untitled ones).
- **Budget:** ~30k.
- **Landed:** BOARD toolbar gains "Descriptions: show/hide" and a single
  "Collapse all / Expand all" button (it reads "Expand all" once every lane
  and group is collapsed). `Services/BoardCollapse.cs` now owns the lane and
  group keys and `AllKeys`, which collects every key including the untitled
  lane and group. Descriptions are a `no-descriptions` class on
  `.board-surface` hiding lane, group and card synopses; the flag rides the
  session key (`passage.setBoardDescriptions`, `boardDescriptions`). Checked in
  the running app: collapse all then expand all restored 5 lanes, 7 groups and
  8 cards; hiding emptied every synopsis; the toggle read "show" after a
  reload. Not done: collapsed state itself is not persisted (unchanged); the
  lane synopsis stays visible on a collapsed lane when descriptions are shown.

---

## Tier S3 — Views

The brief's item 11 (Beat Sheet follow-ups) **has no row of its own**. The
Beat Sheet PR (#40) listed four out-of-scope items:

- **Storyline colours on beats:** S1.4.
- **A beats panel inside the script editor:** S2.3.
- **Structure templates, or seeding beats from the Shape tool:** S2.2. The
  seed comes from the template, and the Shape tool's own turns aren't seeded:
  they're sequence synopses, which the script track shows once scenes exist.
- **Export of the beat sheet:** S3.2.
- **Image uploads on beats:** Not going in.

Nothing is left over, so an empty row would only cost a session (rule 2).

### S3.1 Arc view — `done`

- **Read:** Standing decisions A, D. BEATS view 497–574 and `BeatSheet.cs`
  (`MapScenes`). Beat gesture `passage.js` 789–985. Storylines from S1.2,
  tags from S1.4.
- **Do:**
  - A fifth main tab, ARC (the PREVIEW index shifts again: check nothing
    persists `_mainViewIndex`, as PR #40 did).
  - An SVG chart. Each storyline is a line in its colour, joined through the
    items tagged with it. Untagged items have no point on that line.
  - Each storyline has one named axis (`Storyline.ArcAxis`, default
    "Tension"), editable in ELEMENTS. The value is 0–100, and an untouched
    point sits at 50 drawn hollow.
  - Points are dragged vertically in JS, with one `OnArcPointMoved(storylineId,
    targetId, value)` on release.
- **The x-axis: offer both, default Board order.**
  - **Board order:** scene cards and beats-converted-to-cards in script
    order, evenly spaced. Every card has a position, and it's the order the
    writer reasons about.
  - **Pages:** scene cards at their `MapScenes` page, plus grid beats at their
    start page. Tray beats and unplaced scenes are listed as "not on a page"
    under the chart and never guessed (rule 12).
  - The toggle is per browser. The arc values are the same in both modes,
    because they're keyed by item, not position.
- **Detached points** follow decision B: listed under the chart, never
  dropped.
- **Shared-library impact:** none.
- **Success:**
  - Tag four scenes "A Story", drag their points: a red line joins them, and
    it survives a reload.
  - Switch to Pages: the same values at page positions, and tray beats are
    listed as unplaced.
  - No server frames during a drag.
  - Filters from S1.5 dim lines and points here too.
- **Tests:** point ordering in both modes, the value clamp, unplaced items
  reported, and a deleted beat's points cascading.
- **Budget:** ~120k.
- **Landed:** ARC is main tab index 3 and PREVIEW moved to 4; nothing persisted
  `_mainViewIndex`, so the only edits were the tab buttons and the PREVIEW
  branch. `Services/ArcLayout.cs` is the pure layer: `Build` returns each
  storyline's points, the unplaced list and the detached list; `SetValue`
  clamps to 0–100 and refuses an item not tagged with the storyline.
  `ArcPoint` (storyline, target, kind, label, value) sits in `StoryData.Arc`
  and is stored only once dragged; an untouched point is 50, drawn hollow.
  `Storyline.ArcAxis` (default "Tension") is edited in ELEMENTS and named in
  the chart legend. Board order spaces every card evenly and lists beats as
  "not on the Board" (a beat has no Board slot until converted); Pages uses
  `MapScenes` pages and grid beats, and lists acts and sequences, scenes at
  page 0 and tray beats as "not on a page". The x-axis choice rides the
  session key (`passage.setArcMode`). The drag is `startArcDrag` in
  `passage.js`: pointer capture, no server frames until release (0 sent
  during 15 moves, then one `OnArcPointMoved`). The S1.5 board filter dims
  lines and points. Deleting a beat or a storyline removes its points. Checked
  in the running app: points survive a reload, Pages shows the same values,
  the axis edit shows in the legend. Not done: a point whose item loses the
  storyline tag is kept dormant and not listed (it returns if re-tagged);
  the SVG labels go through `MarkupString` because Razor reserves `<text>`.
  The `passage.js` 789–985 range named above is stale: the beat gesture is
  at ~1070–1260.

### S3.2 Outline export: Fountain and text — `done`

- **Read:** `Passage/Passage.Web/Services/ExportService.cs` (whole file,
  small). Export menu in `Editor.razor` (grep `ExportService`; read that range
  only). Standing decision A.
- **Do:**
  - "Export outline…" with field checkboxes: synopses, storylines,
    characters, locations, key beats only, and include the Beat Sheet (beats
    by page, then the inbox).
  - **Fountain:** `#` / `##` / `###` headings, scene headings and `=` lines,
    with tags as `[[Storylines: A Story, B Story]]` notes, so it re-opens in
    Passage as an outline.
  - **Plain text:** an indented outline.
  - Web-only. Not added to `ExporterCatalog.DefaultExporters`, because that
    list feeds the desktop menus (HIGHLAND-PLAN's standing decision, applied
    here).
- **Shared-library impact:** none.
- **Success:** export with synopses and storylines, then open the `.fountain`
  in Passage: the BOARD shows the same lanes and cards. "Key beats only"
  exports only starred cards and their act/sequence headers. Detached items
  are excluded and their count is stated in the dialog.
- **Tests:** golden output for a fixture script with every field combination
  that changes output. The export round-trips through the parser to the same
  board shape.
- **Landed:** Export ▾ → "Export outline…" opens a dialog with the six
  checkboxes (synopses, storylines, characters, locations, key beats only,
  Beat Sheet), a Fountain and a Plain text button, and the detached count.
  `Services/OutlineExport.cs` is the pure writer, fed the Board lanes, the
  `StoryData` and the beats; the download is the new `passage.downloadText`.
  Not added to `ExporterCatalog.DefaultExporters`. Choices a later row should
  know: tags are `[[Storylines: …]]` / `[[Characters: …]]` / `[[Locations: …]]`
  notes under the `=` lines, and `[[id:…]]` markers are not exported, so the
  file is a fresh outline rather than a copy that shares ids with the
  original; the stars are not exported either. Section cards deeper than `##`
  come out as `###`, as the Board already flattens them. The Beat Sheet is a
  `/* … */` boneyard in Fountain, so re-opening leaves the lanes untouched
  and the beats stay in the file for a human reader (they do not re-import).
  Key beats only keeps starred cards (and starred beats) plus the act and
  sequence headers above them; a lane or group with no star is dropped. Tests:
  three goldens (all fields in both formats; key-only with synopses and
  locations off), a per-field toggle check, and the parse round trip with and
  without the Beat Sheet. Tried in the running app: exported, saved as a new
  script and opened; the BOARD showed the same lanes and cards.
- **Budget:** ~60k.

### S3.3 Outline export: PDF — `done`

- **Read:** `Passage/Passage.Export/ScreenplayPdfExporter.cs` 31–216 (how the
  hand-rolled PDF is written: no PDF library is referenced). S3.2's outline
  model.
- **Decide first (rule 7):** `WritePdf` and its helpers are `private` in a
  shared library. There are two options:
  - **(a, recommended)** A small web-only `OutlinePdfWriter`, using the same
    built-in Courier fonts and the same text sanitising approach, written
    fresh.
  - **(b)** Make the shared writer's primitives `internal`/`public` for
    reuse. That's a shared-library change: it's additive and desktop output
    is unchanged, but it's still a change.
  Ask before (b).
- **Do:** the same fields as S3.2. Act/sequence headers in bold, a storyline
  colour swatch beside each card, and page numbers.
- **Success:** the PDF opens in a browser and in Evince. Long synopses wrap,
  and a 60-card board paginates without clipping.
- **Tests:** page count for a fixture, wrapping at the column width, and
  sanitising of `( ) \`.
- **Landed:** Took option (a), agreed before starting: `Services/OutlinePdfWriter.cs`
  is a web-only writer, so `Passage.Export` is untouched. "Export outline…"
  gains a **PDF** button (same six checkboxes); the download is the new
  `passage.downloadBase64`. To keep every format on the same cards,
  `OutlineExport.Build`'s traversal was extracted into `OutlineExport.Walk`
  (no output change: the S3.2 goldens still pass). Layout: Letter, Courier
  10pt; Act (14pt) and Sequence (12pt) headings and section cards are bold;
  scene cards are regular; synopses wrap at the column width; storyline
  swatches sit in the left gutter, one per tagged storyline (grey if the
  storyline has no palette colour); the page number is centred at the foot.
  A heading never strands at the foot of a page. Tests: page count (fixture
  is one page; 60 long cards paginate with every line inside the margins),
  wrapping and long-word splitting, `( ) \` escaping, and swatch counts.
  Tried in the running app: the PDF downloaded and `pdftotext`/`pdfinfo` read
  it cleanly (1 page, headings, synopses, page number); the swatches and a
  multi-page board were covered by the tests only (the scratch library had no
  tags), and Evince was not opened here.
- **Budget:** ~80k.

---

## Tier S4 — Notes

### S4.1 Project notes — `partial` · design note written, awaiting approval

- **Read:** `docs/SLATE-PLAN.md` lines 1–20 and decision B (~92–112) only.
  `Slate` 352–380 (`IdeationDocument`, the precedent for non-script data).
- **Session 1 is a design note only.** Implement only after Adam approves it.
- **The note must settle:**
  - **Scope (rule 7).** SLATE-PLAN's scope decision keeps "Seeds, stages…
    next-move notes, briefs" in Obsidian: "Passage gets the tools, not the
    practice". A per-script notes area for brainstorming and world-building
    sits close to that line. The note proposes where it stands, and Adam
    decides whether Passage should have notes at all.
  - **Where notes live.** The options are:
    - **(a)** a sidecar `Story.Notes` list of markdown pages;
    - **(b)** companion `.md` files in the library (Passage already edits
      Markdown documents). That gives the full editor, save, export and
      Obsidian compatibility for free, but adds file-list clutter and has no
      link to the script except by name.
    Recommend one.
  - **What belongs where.** Slate holds working-out that is promoted into the
    script (fills, chains, runs) and stays as it is. Notes hold reference
    material that is never promoted (a world sheet, a character worksheet).
    The Ideation `Kept` lines stay in Ideation.
  - **Templates.** Character worksheet, location sheet and logline/premise
    sheet are built in. Custom templates live in `.slate/templates.json` (from
    S2.2). The built-ins must be Passage's own wording.
- **Budget:** ~30k (design note only).

#### S4.1 design note (2026-10-25) — awaiting Adam's approval

**Recommendation in one line:** yes, add notes, as a small per-script area
stored in the sidecar (`Story.Notes`), edited in the existing Markdown editor,
with three built-in templates and an export-to-`.md` button. Nothing is built
until this is approved.

**1. Scope (rule 7).** SLATE-PLAN's scope line keeps "Seeds, stages…
next-move notes, briefs" in Obsidian: "Passage gets the tools, not the
practice". Those are *project-management* artefacts: where a story is in its
life, what to do next. A world sheet or a character worksheet is *reference
material for one script*. It's closer to Cast colours and the Beat Sheet, which
decision A (settled 2026-10-25) already moved into the sidecar, than to a
story pipeline. So notes don't contradict the line, but they do lean on it.
The honest cost is a second place a writer might keep this material, next to
Obsidian. Two guard rails keep it small:

- Notes are **per script**. No cross-script board, no stages, no status, no
  last-touched, no "next move" field. Those stay in Obsidian.
- Notes are **plain Markdown pages with a title**. No linking between notes,
  backlinks, tags or search across scripts. A writer who wants a graph has one.

**Adam decides:** whether Passage has notes at all. If not, delete S4.1; nothing
else in the plan depends on it. If yes, the line above is the proposed edge,
and it's worth one sentence in SLATE-PLAN's scope paragraph, as S1.1 did for
decision B.

**2. Where notes live.**

| | (a) Sidecar `Story.Notes` | (b) Companion `.md` files |
| --- | --- | --- |
| Link to the script | By construction: same sidecar, same lifetime | Name convention only (`Heist.notes.md`); breaks on rename |
| Library list | Untouched. `ScriptLibrary.List()` filters `.fountain/.md/.txt` in the root, so every companion shows up as a "script" | One extra entry per note unless `List()` learns to hide them (a change to a shared list) |
| Opening one | New small page editor reusing CodeMirror in Markdown mode | Free: it is already a Markdown document, with the full editor, but also page counts, goals, Board and Beats built for scripts |
| Save | `SaveSlate` (whole-sidecar, last-write-wins, like everything else in it) | `ScriptLibrary.Save`, one file per note |
| Obsidian compatibility | Export button writes `.md` (or a zip) | Native, files sit in the data volume |
| Rename/delete a script | Notes follow the sidecar. `PruneOrphans` deletes them with it (the existing risk, below) | Notes are orphaned, never deleted, never reconnected |
| Slate revamp | Inside `StoryData`, so the revamp moves it with everything else (decision A) | Unaffected, but a second storage convention to remember |

**Recommend (a).** The decisive points are the library list and the link. (b)
looks free but isn't: every note appears in the file list beside the scripts,
the editor treats it as a script, and the only tie to its script is a filename
that nothing keeps in sync. (a) keeps one home for "everything about this
script", which is what decision A already says.

(a)'s real weakness is **`PruneOrphans`**. Notes are prose, so losing them to a
rename outside the app hurts more than losing a tag. Mitigations, none of which
change `PruneOrphans` itself:

- Every notes page has **Download .md**, and the notes list has **Download all**
  (zip). This is also the Obsidian path.
- The existing prune report (rule 12) already names what was removed; S4.2
  adds the note count to that message when a pruned sidecar held notes.

Size is not a concern: notes are kilobytes and `SaveSlate` already rewrites the
sidecar whole. The editor saves on a debounce, not per keystroke, which stays
server-side because it is save, not per-keystroke logic (trap 4).

```csharp
// Added to StoryData (decision A). No new top-level SlateDocument property.
public List<NotePage> Notes { get; set; } = new();

public sealed class NotePage
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Title { get; set; } = "";
    public string Body { get; set; } = "";        // Markdown
    public string Template { get; set; } = "";    // built-in/custom id it started from, informational only
}
```

New properties with defaults need no `SlateVersion` bump (decision A).

**3. What belongs where.**

| Material | Home | Why |
| --- | --- | --- |
| Fills, chains, runs, rungs, scores (WOAC, Split, Position, Push/Pull, Lens) | **Slate**, unchanged | Working-out with one promotion path into the script |
| Ideation `Kept` lines | **Ideation**, unchanged | Deliberately disposable, library-wide, not tied to a script |
| World sheet, character worksheet, premise/logline sheet, research, reference | **Notes** | Never promoted into the script text |
| Character name, colour, tags | **Cast / ItemTags**, unchanged | Structured, drives the Board and Arc |
| Storyline names/colours | **Storylines**, unchanged | Structured |

Rule for the grey zone: *if it ends up as a line in the script, it is Slate; if
it only informs the writer, it is a note.* Notes never feed the parser, the
Board, exports or page counts. There is no promotion path and no "send to
script" button. (A later row could offer copy-to-clipboard; not part of S4.2.)
A note may be about a character by name, but it does not link to `Cast`:
renaming a character (`RenameCharacterAsync`) does not touch notes.

**4. Templates.** Three built-ins, in Passage's own wording, written as short
Markdown prompts (headings plus one-line questions), not copied from any
published worksheet: **Character** (who they are, want vs need, flaw, how they
talk, relationships, arc), **Location** (what it looks like and sounds like,
who owns it, what it means to the story, what can happen there only), and
**Logline & premise** (the one sentence, the question the story asks, the
protagonist, the pressure, the ending it is pointed at). Templates seed a new
page's `Body`. They are not a live link; editing the page never changes the
template.

Custom templates reuse **`.slate/templates.json`** (S2.2), which already
exempts itself from `PruneOrphans`. S2.2's file holds structure templates, so
S4.2 must read S2.2's schema first and add a sibling list (`NoteTemplates`)
rather than overload the structure entries; if S2.2's shape can't take it, flag
the mismatch before coding (rule 7). Built-ins live in a `NoteTemplates.cs`
service in `Passage.Web`, as `StructureTemplates.cs` does. "Save as template"
from an open page stores the title and body.

**5. UI sketch (for S4.2 to confirm against `Editor.razor`).** A **NOTES** entry
in the sidebar (a tab or a section of an existing one; S4.2 chooses after
reading the tab code), a list of pages with new/rename/delete, and a plain
Markdown editor on the selected page. Delete is confirmed (sidecar-owned data,
as with beats). No second CodeMirror configuration beyond switching the mode.

**Not in S4.2:** note-to-note links, search across notes, images, sharing,
per-note history, and any bridge to the Slate runners.

**Open for Adam:**
1. Notes at all, yes or no.
2. Sidecar (recommended) or companion files.
3. Whether the three built-in templates are the right set.

---

## Not going in

- **Comments and collaborators.** There is no auth, and saving is
  last-write-wins (`docs/web-app.md`, CLAUDE.md trap 5). Comments without
  identities would be unattributed.
- **Images on beats or cards.** Wait until storage limits on the data volume
  are decided: uploads would be the first unbounded writes to `/data`.
- **Season and multi-episode outlines.** These need multi-script projects,
  and the library is a flat list of scripts with one sidecar each.
- **AI features** (generated beats, loglines, arc suggestions). SLATE-PLAN
  keeps model access deterministic-first behind a seam built last. Nothing in
  this plan builds that seam.
- **Moving the existing Board card drag to JS** (decision D). It's a real
  trap-4 problem, but it isn't a story feature. It gets its own row if Adam
  wants it.
