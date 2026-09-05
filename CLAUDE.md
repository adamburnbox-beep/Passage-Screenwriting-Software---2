# CLAUDE.md — Passage (web-focused scope)

Active development in this repo targets **`Passage.Web` only** (Blazor Server).
The desktop frontends are frozen reference implementations, not work targets.

Read `PROJECT_RULES.md` before writing code. It is the behaviour contract and
takes precedence over anything here. This file is scope and mechanics only.

## Scope rules

| Path | Rule |
| --- | --- |
| `Passage/Passage.Web/` | **Work target.** Edit freely. |
| `Passage/Passage.Core/`, `Passage.Parser/`, `Passage.Export/` | Shared with the desktop apps. Edit only when the task explicitly calls for it, and never in a way that changes desktop behaviour. A change here lands in all three frontends. |
| `Passage/Passage.App.Linux/` | **Read-only reference.** The source of truth for features being ported to the web. Never edit. |
| `Passage/Passage.App/` | **Read-only, rarely relevant.** WPF, 24k lines, cannot build on Linux. Do not read unless a task names it. |
| `Passage/Passage.Web/wwwroot/lib/` | **Never read.** Vendored CodeMirror, ~10k lines. Reading it will blow the context budget for no benefit. |

When porting a feature, read the **line ranges** named in `docs/WEB-PARITY.md`,
not whole files. `MainWindowViewModel.cs` is 2,797 lines and
`MainWindow.axaml.cs` is 1,348 — reading either in full is a budget failure
under PROJECT_RULES rule 6.

## Commands

```bash
# Build the web target (excludes WPF and Avalonia)
dotnet build Passage.Web.slnf

# Run locally
dotnet run --project Passage/Passage.Web/Passage.Web.csproj

# Tests
dotnet run --project Passage/Passage.Tests/Passage.Tests.csproj
# (with only the .NET 10 runtime installed, prefix: DOTNET_ROLL_FORWARD=Major)

# Container build, as deployed
docker build -t passage-web .
```

`dotnet build Passage.Web.slnf` passing is the minimum bar for any change.

## Traps

1. **Target frameworks are pinned.** `Passage.Core`, `Passage.Parser`,
   `Passage.Export`, `Passage.Web` and `Passage.Tests` are `net9.0`.
   `Passage.App.Linux` is `net10.0`; `Passage.App` is `net9.0-windows`.
   Do not "modernise" the shared libraries to net10 — the Dockerfile builds on
   `dotnet/sdk:9.0` and the container build will break while `dotnet build`
   still passes locally.

2. **The Dockerfile hardcodes its copy list.** If you add a `ProjectReference`
   to `Passage.Web.csproj`, add the matching `COPY` line to `Dockerfile` in the
   same change, or the image build fails.

3. **`TreatWarningsAsErrors` is enabled on `Passage.Web` only.** This is
   deliberate. Fix the warning; do not disable the setting.

4. **Blazor Server round-trips every interaction over a WebSocket.** Anything
   that must respond per-keystroke — live indentation, autocomplete
   suggestions, page-break rules, find-as-you-type — belongs in
   `wwwroot/js/passage.js` and CodeMirror, client-side. Porting the Avalonia
   per-keystroke logic into Razor event handlers will be visibly laggy over a
   LAN. Server round-trips are for parse, analysis, save and export.

5. **Saving is last-write-wins and there is no auth.** Do not add features that
   assume otherwise without raising it first (`docs/web-app.md`).

## Workflow

One feature per session. Read `CLAUDE.md`, `PROJECT_RULES.md` and the single
relevant row of `docs/WEB-PARITY.md` — nothing else — then implement, build,
update that row's status, and commit. `docs/WEB-SESSIONS.md` holds the
prompt for each feature.

`docs/WEB-PARITY.md` is the source of truth for remaining work. Keep it
accurate in the same session as the change; a stale row costs the next session
more than it saved this one.
