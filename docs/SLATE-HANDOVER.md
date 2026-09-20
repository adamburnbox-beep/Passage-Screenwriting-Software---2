# Writer's Tools handover

## What this doc is

State of play for the Writer's Tools work in `Passage.Web` (internally "Slate",
after the vault system it comes from), written for the next agent. Read this,
then `CLAUDE.md`, `PROJECT_RULES.md`, and `docs/SLATE-PLAN.md` — in that
order — before touching anything. Everything below is as of the Phase 2
commit on branch `web-scope`, 2026-09-20.

---

## The one-paragraph version

The writer (Adam) has a set of writing-block exercises in his Obsidian vault
("the Slate"). We are bringing the *tools* — not the board, not the practice —
into the web app as a right-hand dock called **Writer's Tools**. Three of
seven phases are built: the foundations (sidecar storage, the dock), **Fill**,
which finds `[SOMETHING happens]` placeholders in a script and walks the
writer through resolving one, and the **forward chain** — the WOAC chain and
the Character Flaw Brainstorm, with the first burst timer. The remaining
tools are listed in the dock as "not built yet". The writer asked to be able
to test the others, so **building them, one per session in plan order, is
the next step** — unless he names a different one first.

---

## Where things are

| Thing | Path |
| --- | --- |
| The plan — phases, decisions, Rule 0, status per row | `docs/SLATE-PLAN.md` (**source of truth for remaining work; update the row in the same session as the change**) |
| Copy-paste prompt per phase | `docs/SLATE-SESSIONS.md` |
| Sidecar store | `Passage/Passage.Web/Services/SlateStore.cs` |
| Placeholder scanner | `Passage/Passage.Parser/BracketScanner.cs` (+ `FountainMarkup.MaskOmissions`) |
| Dock markup, views, handlers | `Passage/Passage.Web/Components/Pages/Editor.razor` — `<aside class="workshop">` ~389; tools view ~395; BRACKETS ~477; FILL ~510; chain list and runner ~545–660; handlers from `// ---- Writer's tools dock` ~2320, `// ---- Forward chain` ~2560 |
| Synopsis placement (promotion target) | `Passage/Passage.Web/Services/SynopsisPlacement.cs` |
| Client-side pieces | `Passage/Passage.Web/wwwroot/js/passage.js` — `applyWorkshopWidth` / `initWorkshopResize` ~125, `replaceInLine` ~540, `insertLinesAt` ~556, `scrollToLine(line, focus)`, burst timer `startBurst` / `stopBurst` after `focusEditor` ~960 |
| Styles | `Passage/Passage.Web/wwwroot/css/app.css` from `/* ---- Writer's tools dock` ~954 |
| Tests | `Passage/Passage.Tests/Program.cs` — `TestBracketScanner*`, `TestSlateStore*`, `TestSynopsisPlacement` (the test project references `Passage.Web`) |
| Source material (the spec for every tool's text) | `/home/arosa/Sync/Obsidian Vault/Story/Slate/` — five live files in `Worksheets/` (the numbered Short/Medium/Long files are stubs), `How the Slate Works.md`, `loosening-up-practice.md`, `engines.md` |

---

## What's built

### Phase 0 — foundations (`42bfa92`)

- `SlateStore`: `<data>/.slate/<validated script name>.json`, `slateVersion: 1`.
  Keys go through `ScriptLibrary.TryValidateName`. Delete cascades; orphans are
  pruned on the first interactive render with a status message. A sidecar
  that fails to parse is left alone and saves are refused with a status line.
- The dock: collapsible (`TOOLS` button in the topbar), drag-resizable, both
  persisted in the existing `passage.session.v1` localStorage key. Below 900px
  it covers the main area (the tools are meant to work from a phone).

### Phase 1 — Fill (`de4884d`, revised in `b15b95c`, `4cc1c84`)

- `BracketScanner.Scan(text)` → ranked `Bracket(LineIndex, Column, Text, Rank)`.
  Rank 0 = all caps or opens with SOMETHING / CONDITION / TODO; rank 1 = the
  rest, still listed. Every line is scanned. `Text` is the span as written,
  brackets and padding included — it is what the replacement matches.
- `DocumentAnalysis.Brackets`, derived on every parse, never stored.
- BRACKETS view: "Hand me one" (round-robin over non-ignored, ranked
  brackets; cursor resets on any queue change), then the list with jump /
  Fill / ignore per row. Ignored brackets are struck through, restorable.
- FILL view: the worksheet's text, three candidate rows, "what the wrongness
  points to", the escape-hatch answer field. Every field change saves the
  run to the sidecar keyed by bracket text. **Fill it** replaces the span via
  `passage.replaceInLine`, which matches by text on the named line and
  returns false if the bracket has gone — undo/redo verified by hand. On
  success the next bracket is handed straight away.

### Phase 2 — Forward chain

- `SlateDocument.Chains` (`ChainRun`: `Path` Woac | Flaw, `Seed`, `Rounds`
  or `Answers`, `ReadBack`) and `SlateDocument.Burst` (`Enabled`, `Seconds`).
- A WOAC round ≥ 2 has no stored Want: the view shows the previous round's
  Consequence live. Writing a consequence in the last round appends the next.
- Burst timer (decision E) is now built, in `passage.js`: slots are
  `textarea[data-slot]` in DOM order; 5 s read-in, then the clock; Enter or
  expiry advances; a MutationObserver waits up to 2 s for Blazor to add the
  next round. Off by default, 30 s when on, per script.
- Promotion inserts `= line` under the nearest heading above the caret via
  `insertLinesAt`; `SynopsisPlacement.Find` decides where and the runner
  says so before the tap. Undo verified.
- See the row's Done and Deviations notes in `SLATE-PLAN.md` for the rest,
  including why neither the promote button nor Fill's "Fill it" is ever
  disabled.

### Decisions made in conversation that the plan now records

- **Name:** "Writer's Tools" in the UI. `Slate` stays in identifiers, the
  sidecar directory and the doc filenames.
- **Dock top level** is *"what have you got"* — five situations, the tools
  under each, every tool with a plain-language description. See the **copy
  rule** in `SLATE-PLAN.md` decision A: written for someone who has never
  seen the vault and may never have written anything; no filler, no slogans.
  Adam called out one generated-sounding line within minutes of seeing it.
- **Decision E — burst timers:** ready (input focused *before* anything
  counts) → read-in countdown (~5s) → write (the real timer; Enter commits
  and moves to the next slot) → done (slots filled, or expiry, which only
  advances). All client-side. Built in Phase 2; Phase 6 should reuse
  `startBurst` rather than write another. Adam said "10 seconds for three
  ideas"; the vault says 30-second bursts — the default is 30 and the
  number is his to change in the runner; not yet settled with him.
- **The worksheets are the spec** for each runner's text, verbatim or near
  it. The 2026-09-20 review found the first runner had paraphrased them
  thinner and Adam noticed. Read the worksheet in the vault before building
  its tool.
- **Save As / .docx export:** dropped. Not needed.
- **Plan corrections from the worksheet review:** the Lens is pick-one inside
  Revise (Phase 5), not a rolled standalone tool; Situation collision needs no
  corpus; Revise's read-back fields are in scope.

---

## What's next

Adam wants to test the other tools. Build them **one per session, in this
order**, each with its own commit and its own row update in `SLATE-PLAN.md`:

1. **Phase 3 — A→Z→Split, Belief Split, Extend Backward.** Rung 3 of Split
   writes `#`/`##` sections into the script by ranged insert.
2. **Phase 4 — Bridge + Position.** Bridge's 30-second rounds use
   `startBurst`.
3. **Phase 5 — Push/Pull and Lens** (plus the optional diagnostic Belief
   Split and read-back fields).
4. **Phase 6 — Ideation burst**, full-screen overlay, on `startBurst`.
5. Phase 7, the model seam, is deliberately last and probably not wanted yet.

Each session: read Rule 0 again, read the worksheet in the vault, use the
prompt in `SLATE-SESSIONS.md`. Every runner: no required fields, closable at
any point with state kept, no progress indicators, no history view, and every
write into the script through a ranged operation (`replaceInLine`,
`insertLinesAt`, `replaceLineRange`) — never `ReplaceEditorContentAsync`.

Shape to copy: each tool is a `WorkshopView` with its record type in
`SlateDocument`, saved on every `@bind:after`, and a "← back" that keeps
state. The Fill runner is the reference for a tool aimed at one thing in
the script; the chain runner is the reference for a kept-per-script run
with a list, a timer and a promotion.

---

## Environment — read this before running anything

- **The `dotnet` on `PATH` is broken.** `/usr/bin/dotnet` is a stray
  `dotnet-host-10.0` apt package with no runtime under it. The real SDK is
  `~/.dotnet/dotnet` (9.0 and 10.0). Adam has since appended
  `~/.dotnet` to `PATH` in `.bashrc`, but in any shell that predates that,
  or in your Bash tool, use the full path or `export PATH="$HOME/.dotnet:$PATH"`.
- **Run as Development.** `Properties/launchSettings.json` (added `a79f460`)
  sets it and port 5210. Running as Production outside a publish serves
  static files as empty responses to browsers that ask for gzip — the page
  renders unstyled and never connects. The Docker build is unaffected.
- **A server may already be running.** Check with `pgrep -af Passage.Web`;
  `pkill -f Passage.Web` stops it. The Phase 2 session ran it through the
  agent browser's `preview_start` (`passage-web`), which stops with the
  session. It does **not** hot-reload — rebuild and restart after every
  change. Static files (`passage.js`, `app.css`) do pick up on reload.
- **Preview in the agent browser:** `.claude/launch.json` (untracked, has a
  machine-specific path) defines `passage-web` on 5210 and `passage-web-alt`
  on 5211 — use the alt one when Adam's server holds 5210.
- **Dev data** lives at `Passage/Passage.Web/bin/Debug/net9.0/data/` (there is
  no `/data` on this machine). `fill-test.fountain` there has four brackets
  for testing. Sidecars are in `.slate/` beside it.
- **Git shows ~140 files modified.** Those are exec-bit flips from the
  filesystem, not content. Commit with `git -c core.fileMode=false add <paths>`
  and `git -c core.fileMode=false commit`; never `git add -A`. Commits end
  with the `Co-Authored-By` line the session gives you.
- Build: `dotnet build Passage.Web.slnf` (must be warning-free —
  `TreatWarningsAsErrors`). Tests: `dotnet run --project
  Passage/Passage.Tests/Passage.Tests.csproj`.
- **The agent browser's "Return" key sends an empty `key`.** Use
  `key: "Enter"` when testing anything that listens for Enter, or the
  handler will look broken when it is not. Its `type` also lands a beat
  after a click that causes a Blazor round-trip; check state after a wait.

---

## Verifying a tool

The minimum bar for calling a runner done, all checked by hand in a browser:

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
- **Save As does not exist** in the web app; the dialog only names untitled
  buffers. Don't build a cascade for it.
- **A button disabled until a bound field has a value needs two taps** when
  the writer types and taps straight away: the field's change event and the
  click travel together and the click hits a still-disabled button. Leave
  the button enabled and put the "nothing to do" in the status line.
- **The dock body is one scroll container.** A view opened from the bottom
  of a long list opens scrolled to its own bottom unless it scrolls itself
  to `.workshop-nav` (`ScrollWorkshopToTopAsync`).
