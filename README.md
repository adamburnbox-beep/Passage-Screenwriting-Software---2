# Passage

A distraction-light screenwriting app for Fountain-format scripts, with a
Windows (WPF) frontend and a Linux (Avalonia) frontend sharing the same
core libraries.

## Repository layout

```
Passage/
  Passage.Core/       Shared domain logic: goals, formatting, text analysis,
                      and platform-neutral services (session, recovery,
                      session-goal configuration)
  Passage.Parser/     Fountain parser, screenplay model, page estimation
  Passage.Export/     PDF / text export
  Passage.App/        Windows frontend (WPF, net9.0-windows) — builds on
                      Windows only
  Passage.App.Linux/  Linux frontend (Avalonia, net10.0)
  Passage.Tests/      Test runner (console project)
docs/                 Design notes, port handover, future improvements
```

Both frontends are thin UI layers over the same `Passage.Core`,
`Passage.Parser`, and `Passage.Export` libraries. That shared core is why
the two ports live in one repository: a change to parsing, exporting, or
goal logic lands in both apps at once.

## Building

### Linux

The full solution includes the WPF project, which cannot build on Linux.
Use the solution filter, which contains everything except `Passage.App`:

```bash
dotnet build Passage.Linux.slnf
dotnet run --project Passage/Passage.App.Linux/Passage.App.Linux.csproj
```

### Windows

```powershell
dotnet build "Passage Screenwriting Software.sln"
dotnet run --project Passage/Passage.App/Passage.App.csproj
```

## Tests

```bash
dotnet run --project Passage/Passage.Tests/Passage.Tests.csproj
```

(If only the .NET 10 runtime is installed, prefix with
`DOTNET_ROLL_FORWARD=Major`.)

## Documentation

- `docs/Design Principles.md` — visual and interaction design rules
- `docs/linux-port-handover.md` — state of play for the Linux port
- `docs/FUTURE_IMPROVEMENTS.md` — known gaps and suggested approaches
- `PROJECT_RULES.md` — coding behavior contract for contributors/agents
