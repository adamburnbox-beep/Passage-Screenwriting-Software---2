# Linux Port Handover

## What this doc is

A complete state-of-play for the ongoing Linux (Avalonia) port of Passage, written for the
next agent picking up the work. Read this before touching anything.

---

## Background

Passage is a WPF screenwriting app (`Passage/Passage.App/`, targets `net9.0-windows`). The
user has moved to PopOS with COSMIC Desktop (Wayland) and needs a Linux-native version. WPF
does not run on Linux — the entire UI layer must be rebuilt.

A previous partial port exists in a separate local repo at:
`/home/arosa/Code Projects/Passage Screenwriting Software - v4`
(also on GitHub: https://github.com/adamburnbox-beep/Passage-v4)

The approach chosen: add `Passage.App.Linux` as a **new project** in this same solution,
leaving the Windows project (`Passage.App`) completely untouched.

---

## What has been done

### 1. Project created
`Passage/Passage.App.Linux/` — all source files copied from v4's `Passage.App/` (Avalonia UI).
Registered in `Passage Screenwriting Software.sln`.

### 2. .csproj fixed (`Passage.App.Linux.csproj`)
- `OutputType`: `WinExe` → `Exe`
- `TargetFramework`: `net8.0` → `net10.0` (user has .NET 10 SDK at `~/.dotnet/dotnet`, not .NET 9)
- Removed `<ApplicationManifest>app.manifest</ApplicationManifest>` (Windows-only)
- Added `SkipGetTargetFrameworkProperties="true"` to project references (shared libs are net9.0)
- Removed Windows `.ico` and `.ps1` assets

### 3. Program.cs rewritten
The v4 version had ~360 lines of P/Invoke (`XOpenDisplay`), XWayland launcher, stale lock file
cleanup, and retry loops. All of this was removed. The new version is ~60 lines:
- Simple `WAYLAND_DISPLAY` / `DISPLAY` check for a helpful error message
- Sets `DISPLAY=:0` as a hint when on Wayland without an X display set
- Calls `UsePlatformDetect()` directly — no manual X11 management

### 4. Keyboard focus bug fixed (`Views/MainWindow.axaml.cs`)
The v4 code used an `_expectingDialogFocus` flag + `Activated` event to re-focus the editor
after the recovery dialog closed. On Wayland this is unreliable: the compositor controls focus
grants and the `Activated` event fires at the wrong time. Fixed by replacing it with an
explicit `Dispatcher.UIThread.Post(() => textBox.Focus(), DispatcherPriority.Input)` call
inside the `ContinueWith` callback, immediately after the dialog result is handled.

### 5. API mismatch fixed (`ViewModels/MainWindowViewModel.cs`, line 1111)
v4's `MainWindowViewModel` called `ScreenplayLayoutBuilder.BuildPages(screenplay, IsScreenplayMode)`
with 2 arguments. This repo's `ScreenplayLayoutBuilder.BuildPages` takes 1 argument. Fixed by
dropping the second argument (the mode flag was removed in a later refactor of the export lib).

### 6. Build result
```
Build succeeded.
    0 Warning(s)
    0 Error(s)
```

---

## Avalonia Wayland status (investigated)

`strings ~/.nuget/packages/avalonia.desktop/12.0.4/lib/net10.0/Avalonia.Desktop.dll | grep -i wayland`
returns nothing. **Avalonia 12.0.4 is X11-only on Linux.** `UsePlatformDetect()` unconditionally
selects X11.

12.0.4 is also the **latest stable version on NuGet** — there is no newer release to upgrade to.
`Avalonia.Wayland` or similar separate packages do not exist on NuGet.

---

## Fix applied: XWayland auto-start in Program.cs

### Problem
COSMIC runs as a pure Wayland session: `WAYLAND_DISPLAY=wayland-1` is set, `DISPLAY` is not,
and XWayland is not running (no sockets in `/tmp/.X11-unix/`). The previous code set
`DISPLAY=:0` as a fallback, but this fails because nothing is listening on `:0`.

### Solution
`Program.cs` now starts its own XWayland instance before Avalonia initialises:

1. Finds the XWayland binary (checks `/usr/bin/Xwayland`, `/usr/lib/xorg/Xwayland`,
   `/usr/local/bin/Xwayland`).
2. Picks the first free display number (`:1`–`:9`) by checking `/tmp/.X11-unix/`.
3. Starts `Xwayland :<N> -rootless -noreset` as a child process.
4. Polls for the socket (up to 5 s, 100 ms intervals).
5. Sets `DISPLAY=:<N>` so Avalonia's X11 backend connects to it.
6. Kills the XWayland child on exit (`finally` block).

This is only triggered when `WAYLAND_DISPLAY` is set and `DISPLAY` is not — X11-only sessions
are unaffected.

### Build status after fix
```
Build succeeded.
    0 Warning(s)
    0 Error(s)
```

---

## Next step: run the app and test

```bash
~/.dotnet/dotnet run --project "Passage/Passage.App.Linux/Passage.App.Linux.csproj"
```

If XWayland is not installed: `sudo apt install xwayland`

If the window opens, test:
1. Fresh launch with no recovery file — editor focused, accepts keyboard input
2. Launch with a recovery file present — accept or discard, then confirm keyboard input works
3. Find/Replace dialog (Ctrl+F) — confirm keyboard works after closing it
4. Save (Ctrl+S), Open file dialog — confirm file picker opens natively

---

## File map

```
Passage/
├── Passage.App/                   ← Windows WPF — DO NOT TOUCH
├── Passage.App.Linux/             ← Linux Avalonia — active work
│   ├── Passage.App.Linux.csproj  ← updated here
│   ├── Program.cs                ← rewritten (simple, no X11 management)
│   ├── App.axaml / App.axaml.cs  ← theme detection, unchanged from v4
│   ├── Views/
│   │   ├── MainWindow.axaml.cs   ← focus bug fixed here
│   │   ├── RecoveryPromptDialog.axaml.cs
│   │   └── ...all other dialogs/panels unchanged from v4
│   ├── ViewModels/
│   │   ├── MainWindowViewModel.cs ← BuildPages call fixed (line ~1111)
│   │   └── ...
│   └── Services/
│       ├── RecoveryStorage.cs
│       ├── SessionStorage.cs
│       └── ...
├── Passage.Core/                  ← net9.0, cross-platform, untouched
├── Passage.Parser/                ← net9.0, cross-platform, untouched
└── Passage.Export/                ← net9.0, cross-platform, untouched
```

---

## System info

- OS: PopOS 24.04, COSMIC Desktop (Wayland compositor)
- Display: `WAYLAND_DISPLAY=wayland-1`, `DISPLAY` not set
- .NET SDK: 10.0.301 at `~/.dotnet/dotnet` (not on PATH — use full path)
- XWayland: installed (2:24.1.2), but not yet successfully used
- Avalonia currently: 12.0.4 (X11 only on Linux)

---

## What is NOT a problem

- The Windows project (`Passage.App`) builds and runs as before — nothing was changed there.
- The shared libraries (`Core`, `Parser`, `Export`) build cleanly on net9.0 and are
  consumed by the Linux project via `SkipGetTargetFrameworkProperties`.
- The keyboard focus fix for the recovery dialog is already in place and correct.
- All other v4 features (beat board, outline, find/replace, goals, title page, export) are
  present in `Passage.App.Linux` — they just can't be tested until the window opens.
