# Web parity — session prompts

One prompt per Claude Code session, in recommended order. Copy the block,
paste it, let it run, review the diff, commit, `/clear`, next.

Every prompt assumes `CLAUDE.md`, `PROJECT_RULES.md` and `docs/WEB-PARITY.md`
are in the repo. Do not merge two sessions into one — the whole scheme depends
on each session reading only its own row.

---

## Session 0 — sanity check (run once, ~2 minutes)

```
Read CLAUDE.md and PROJECT_RULES.md. Do not read any other file yet.

Then verify the scaffolding is correct without changing behaviour:
1. Run `dotnet build Passage.Web.slnf` and report the result.
2. Run `dotnet run --project Passage/Passage.Tests/Passage.Tests.csproj` and
   report the result.
3. Confirm `docs/WEB-PARITY.md` exists and that three line ranges of your
   choice, drawn from three different rows, actually point at the code the
   document claims they do. Report any that are wrong.

Change no files except to correct a wrong line range in docs/WEB-PARITY.md.
```

---

## Tier 1

### Session 1.1 — session state restore

```
Read CLAUDE.md, PROJECT_RULES.md and section 1.1 of docs/WEB-PARITY.md.
Read only the line ranges that section names, plus
Passage/Passage.Core/Services/SessionStorage.cs.

Implement per-browser session restore in Passage.Web: last open document,
editor font size, preview zoom, and caret line. Persist in localStorage via
wwwroot/js/passage.js — NOT on the server data volume, since multiple browsers
share it.

Success: reload the page and the previous document reopens at the previous
caret line with the previous zoom. `dotnet build Passage.Web.slnf` passes.
Then set section 1.1's status to `done` and commit.
```

### Session 1.2 — crash recovery

```
Read CLAUDE.md, PROJECT_RULES.md and section 1.2 of docs/WEB-PARITY.md.
Read only the line ranges it names, plus
Passage/Passage.Core/Services/RecoveryStorage.cs.

Wire crash-recovery snapshots into Passage.Web, scoped per browser session so
one client's unsaved work is never offered to another. On load, if a snapshot
is newer than the file on disk, prompt to recover or discard — match the
wording of Views/RecoveryPromptDialog.axaml.

This is separate from the existing file autosave in Editor.razor; do not merge
them. Success: kill the container mid-edit, reload, get the prompt, recover the
text. `dotnet build Passage.Web.slnf` passes. Set 1.2 to `done` and commit.
```

### Session 1.3 — light/dark theme toggle

```
Read CLAUDE.md, PROJECT_RULES.md and section 1.3 of docs/WEB-PARITY.md.
Read only the line ranges it names, plus
docs/Linux Port UI Overhaul — Visual Consistency with Windows.md.

Add a light theme to Passage.Web as a second custom-property block, with a
toggle that persists per browser. Take the exact palette values from
App.axaml.cs LoadThemeResources — do not invent or approximate colours.

Report any colour in app.css that is hardcoded rather than routed through a
custom property, and fix those in the same change.

Success: toggling swaps every surface with no hardcoded colour left behind, and
both themes match the Avalonia app. Set 1.3 to `done` and commit.
```

### Session 1.4 — keyboard shortcuts

```
Read CLAUDE.md, PROJECT_RULES.md and section 1.4 of docs/WEB-PARITY.md.
Read only the line ranges it names.

Bind the Linux shortcut set in Passage.Web via CodeMirror's keymap in
wwwroot/js/passage.js, client-side. Bind only shortcuts whose feature already
exists in the web app; list the ones you skipped and why.

Do not add Razor-side @onkeydown handlers for editor keys — see CLAUDE.md
trap 4. Set 1.4 to `done` (or `partial` with the skipped list) and commit.
```

### Session 1.5 — Open Recent

```
Read CLAUDE.md, PROJECT_RULES.md and section 1.5 of docs/WEB-PARITY.md.
Read only the line ranges it names.

Add a per-browser recent-files list to the FILES sidebar panel, capped at 8 to
match MaxRecentFiles. Entries that no longer exist on the server are pruned on
load rather than erroring. Set 1.5 to `done` and commit.
```

### Session 1.6 — undo/redo audit

```
Read CLAUDE.md, PROJECT_RULES.md and section 1.6 of docs/WEB-PARITY.md.

First determine whether undo/redo already works in the web editor, and whether
any existing code path replaces the document wholesale and so destroys the undo
stack. Report findings before changing anything.

Then fix only what is actually broken, and add toolbar affordances only if the
keyboard path works. Update 1.6's status to reflect what you found and commit.
```

---

## Tier 2

Order within the tier is by ascending cost: 2.2, 2.5, 2.3, 2.4, 2.1, 2.6.

### Session 2.x — generic Tier 2 prompt

Substitute the section number and title.

```
Read CLAUDE.md, PROJECT_RULES.md and section <N> of docs/WEB-PARITY.md.
Read only the line ranges that section names, plus the web files it names.
Read nothing else — in particular not Passage.App/ and not
Passage/Passage.Web/wwwroot/lib/.

Port <feature> to Passage.Web. Match the Linux behaviour and visual treatment;
where the web platform makes an exact match wrong, say so and propose the
alternative before implementing it.

Success criteria: <fill in — what you will click to check it works>.
`dotnet build Passage.Web.slnf` passes. Then set section <N>'s status to `done`
and commit.
```

### Session 2.1 — Find/Replace

```
Read CLAUDE.md, PROJECT_RULES.md and section 2.1 of docs/WEB-PARITY.md.
Do not read codemirror.js.

The vendored build is CodeMirror 5.65.18 and does not include the search
addons. Vendor searchcursor.js, search.js and dialog.js from the matching
5.65.18 release into wwwroot/lib/codemirror/addon/, reference them from
App.razor, and wire find/replace into the editor with the same option set as
the Avalonia dialog: match case, whole word, replace, replace all.

Do not upgrade CodeMirror. Do not load the addons from a CDN. Do not port the
Avalonia search logic — it is only a fallback if the addons prove unworkable,
and if you reach that point, stop and tell me first.

Success: Ctrl+F opens find, Ctrl+H replace, whole-word and match-case behave
as they do in the Linux app. `dotnet build Passage.Web.slnf` passes. Set 2.1
to `done` and commit.
```

---

## Tier 3

Never run these as a single session. Use the sub-sessions listed in
docs/WEB-PARITY.md section 3.1, and treat 3.2–3.4 as one session each with an
explicit stop.

### Session 3.1b — the one that pays for the rest

```
Read CLAUDE.md, PROJECT_RULES.md and section 3.1 of docs/WEB-PARITY.md.
Read only VM 733–767, 886–952 and 2586–2666.

Extract the Beat Board line-range and splicing helpers into Passage.Core as a
pure, frontend-agnostic class. They have no Avalonia dependency. Then:
- repoint Passage.App.Linux at the extracted code (behaviour must not change),
- add tests for it in Passage/Passage.Tests covering block extents, nested
  section blocks, and out-of-range input,
- use it from Passage.Web to implement inline card editing with write-back.

This edits Passage.Core, which CLAUDE.md normally forbids — section 3.1
sanctions it for this task specifically. `dotnet build Passage.Web.slnf` AND
`dotnet build Passage.Linux.slnf` must both pass before you commit.

Checkpoint after the extraction and before the web work; do not do all three
parts without stopping.
```

### Session 3.x — generic Tier 3 prompt

```
Read CLAUDE.md, PROJECT_RULES.md and section <N> of docs/WEB-PARITY.md.
Read only the line ranges that section names.

Before writing any code, state your plan and your estimated diff size, and
stop. Under PROJECT_RULES rule 6 this task has a hard budget: if the work
trends past your estimate, stop and tell me rather than continuing.

<after approval> Implement <feature>. Success: <fill in>.
`dotnet build Passage.Web.slnf` passes. Set section <N> to `done` and commit.
```

---

## Housekeeping session (run every 4–5 feature sessions)

```
Read CLAUDE.md and docs/WEB-PARITY.md. Read no source file unless a check
below requires it.

1. Confirm every row marked `done` is actually implemented in Passage.Web.
2. Confirm the line ranges in every `missing` row still point at the right
   code — earlier commits may have shifted them.
3. Correct docs/WEB-PARITY.md where it has drifted.

Change nothing else. Commit as a docs-only change.
```
