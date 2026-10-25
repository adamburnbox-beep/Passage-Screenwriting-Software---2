# Writer's Tools redo — session prompts

One prompt per row of `docs/SLATE-REDO-PLAN.md`. Paste the shared preamble,
then the row's lines. One row per session.

## Shared preamble

```
You are implementing one row of docs/SLATE-REDO-PLAN.md in Passage.Web.

Read, in this order and nothing else up front: CLAUDE.md, PROJECT_RULES.md,
docs/SLATE-REDO-PLAN.md (all of it), and from docs/SLATE-PLAN.md only
"Rule 0" and decisions B, D, E. Then the "Environment" and "Traps that have
already bitten" sections of docs/SLATE-HANDOVER.md.

Rules for this session:
- Branch off a fresh main (git checkout main && git pull), one PR for the row.
- Editor.razor is ~7,900 lines: grep for section markers ("// ---- Writer's
  tools dock", "<aside class=\"workshop\"", "_ideationOpen") and read ranges.
  Never read wwwroot/lib/. Never edit Passage.App.Linux or Passage.App.
- Use ~/.dotnet/dotnet (the dotnet on PATH is broken). Build with
  `dotnet build Passage.Web.slnf` (TreatWarningsAsErrors), tests with
  `DOTNET_ROLL_FORWARD=Major dotnet run --project Passage/Passage.Tests/Passage.Tests.csproj`.
- Razor trap: a line inside @code that starts with `<` (e.g. a generic
  `<T>` on its own line) builds locally but breaks the SDK 9 container build.
- Verify in your own server with a scratch copy of the dev data, never on
  Adam's server (port 5210). Run the plan's §8 checklist by hand.
- The §7 decisions are settled; don't reopen them. Anything else the plan
  leaves to Adam, ask before building.
- If something in the plan conflicts with the code, stop and say which
  (PROJECT_RULES 7). Don't average.
- Finish: update the row's status and add Done / Deviations notes under it
  in docs/SLATE-REDO-PLAN.md, in the same commit.
```

## T1

```
Row T1 — engine, model, migration scaffold. No UI change. Read
Services/SlateStore.cs (ChainRun, ChainRound, SlateDocument, SlateStore.Load)
and the existing TestSlateStore* tests in Passage.Tests/Program.cs before
writing. Engine tests must check behaviour listed in the row, not just that
objects construct.
```

## T2

```
Row T2 — launcher, runner shell, WOAC and Flaw. Read the dock markup's
Tools view and Chains/Chain views, and the "// ---- Forward chain" handlers,
before deleting them. Take the WOAC and Flaw questions from the vault
worksheet "Place or Build — Forward Chain.md" and the existing FlawQuestions
array; long text goes behind Why? (plan R-B). Screenshot the launcher and a
runner at 1440px and 375px in the PR.
```

## T3

```
Row T3 — anchors, targets, Right now, Cast / Board entry points. Read
Services/BeatInbox.cs, Services/StoryItems.cs (ToggleTag), Services/NoteTemplates.cs,
SynopsisPlacement.cs, StoryFilter.SceneSpeakers, and Editor's AddBeat and
RenderBoardCard before writing IToolHost's target methods.
```

## T4

```
Row T4 — Fill and Scene Dials. Read the Brackets/Fill dock views and
HandMeOneAsync / AcceptFillAsync. Scene Dials text: engines.md § Scene Dials
and "Place or Build — Position or Fill.md" Path B.
```

## T5

```
Row T5 — Split family. Read the Split/Belief/Extend views, the
"// ---- The Split family" handlers, WriteShapeAsync and SplitScript.cs.
Optional groups (end checks, Mirror Z, Cast card, Obituary Z): the vault
worksheet "Generate — Split, Belief, or Extend Backward.md".
```

## T6

```
Row T6 — Bridge and Position. Read the Bridge/Position views and handlers,
and SendToBoardAsync / BeatInbox.TryPlan. Time-Skip Ladder: engines.md
§ Position.
```

## T7

```
Row T7 — Push/Pull and Understudy. Read the Revise views and the
"// ---- Push/Pull and Lens" handlers, SceneDials, LensMinutes. Read-back
questions: "Revise — Push-Pull and Lens.md". Understudy: engines.md
§ Understudy.
```

## T8

```
Row T8 — Ideation overlay, cleanup, docs. Read the overlay markup and the
"// ---- Ideation burst" handlers. After deleting old record types, confirm
a v1 sidecar from the dev data still migrates (copy one into a test).
```
