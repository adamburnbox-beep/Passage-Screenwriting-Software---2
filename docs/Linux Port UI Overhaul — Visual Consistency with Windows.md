Branch: **Linux Visual Fixes**

## Context

The Linux (Avalonia) port of Passage looks "halfway designed" compared to the Windows
(WPF) version the user is happy with. Investigation found the gap is **structural, not
cosmetic**:

1. **Different palette.** Windows uses a monochrome **e-reader/e-ink** scheme (near-black
   dark theme with white/grayscale accents; the intended light theme is a cream/paper tone).
   The Linux port instead uses a VSCode-style palette (`#1E1E1E` / `#F6F6F6` with **blue
   `#3584E4` accents**). The blue accents on tabs, outline labels and card badges are the
   main thing making it read as a different app.
2. **No styling layer.** The Linux app has **no Styles/Themes folder and almost no
   `<Style>` definitions** — toolbar buttons, the workspace `TabControl`, and the outline
   `TreeView` items all fall back to stock Avalonia `FluentTheme`. That is why tabs and
   cards look unfinished and why buttons lack the Windows hover/press treatment.

**Goal:** Re-skin the Linux port to the e-reader aesthetic with two themes that are direct
inverses of one another — **cream/paper light**, **near-black dark** — and build the
missing control-styling layer so the toolbar/chrome, the (tabs + outline
cards), and the status bar match the Windows visual design. The editor surface and Fountain
syntax colors are already acceptable and stay as-is. Beat Board cards already have decent
styling and only need to inherit the new palette.

**Approach:** Visual *consistency*, not pixel-faithful WPF replication — use idiomatic
Avalonia styling, matching Windows in palette, spacing, corner radii, and feel.

## Key files

- `Passage/Passage.App.Linux/App.axaml` — currently holds ~25 inline dark-only brushes.
- `Passage/Passage.App.Linux/App.axaml.cs` — `LoadThemeResources(bool isLight)`
  (lines 173–227) sets the brushes in code for light/dark; system detection lines 31–171.
- `Passage/Passage.App.Linux/Views/MainWindow.axaml` (646 lines) — toolbar/chrome
  (~37–131), workspace `TabControl` (~217–342), outline/notes item templates (~222–299),
  beat board cards (~465–576), status bar (bottom).
- Reference spec (read-only): `Passage/Passage.App/Themes/DarkTheme.xaml`,
  `LightTheme.xaml`, and `Design Principles.md`.

All Linux brush bindings already use `DynamicResource`, so swapping palette values
propagates automatically — no per-control rebinding needed for color changes.

## Plan

### 1. Replace the palette with an e-reader monochrome scheme (two inverse themes)

Convert the two color sets in `App.axaml.cs` `LoadThemeResources` from the blue/VSCode
values to monochrome e-reader values. Keep the existing keys (the app binds to them already)
and the existing system-detection + switching mechanism. **Eliminate every blue value
(`#3584E4`)** — `ControlAccent`/`CardAccent` become ink-on-paper (near-black in light,
paper-white in dark).

Representative values (tunable during implementation; warm-neutral so the two themes mirror):

| Key | Light (cream/paper) | Dark (inverse near-black) |
|---|---|---|
| ThemeBackground / WindowBackground | `#EFE7D6` | `#15140F` |
| SurfaceBackground | `#F2ECDD` | `#1B1A15` |
| SurfaceMutedBackground | `#E6DDC8` | `#232118` |
| SurfaceRaisedBackground (cards/editor page) | `#FBF6EA` | `#100F0B` |
| SurfaceBorder / EditorPageBorder / CardBorder | `#D8CDB5` | `#2E2B22` |
| ControlBackground | `#ECE3D0` | `#232118` |
| ControlForeground / WindowForeground / EditorForeground | `#2B2620` | `#ECE3D0` |
| ControlBorder | `#CFC1A8` | `#3A352A` |
| ControlAccent / CardAccent | `#2B2620` | `#F2ECDD` |
| ControlPressedBackground | `#2B2620` | `#F2ECDD` |
| ControlPressedForeground | `#FBF6EA` | `#15140F` |
| HeaderText | `#1E1A14` | `#F7F2E6` |
| SecondaryText | `#6B6353` | `#A89E89` |
| MutedText | `#938A76` | `#756C5A` |
| CardBackground | `#FBF6EA` | `#100F0B` |

Add a few keys the new styles need (both themes): `HierarchyIndicatorBrush`
(= `ControlAccent` per theme) for the act-card left bar, and a translucent
`DragOverBackground` to replace the hardcoded `#253584E4` in the outline templates.

> Optional cleanup (recommended, low risk): move the two palettes out of code-behind into
> `App.axaml` using `<ResourceDictionary.ThemeDictionaries>` (a `Light` and a `Dark`
> dictionary). This makes the "direct inverse" relationship explicit and maintainable, and
> lets the code-behind just set `RequestedThemeVariant`. If this proves fiddly, keep the
> existing code-behind approach and only swap values — functionally equivalent.

### 2. Add a control-styling layer

Create `Passage/Passage.App.Linux/Styles/Controls.axaml` (a `Styles` resource) and include
it from `App.axaml` `<Application.Styles>` after `FluentTheme`. Define:

- **Chrome buttons** (`Button.chrome` class, applied to toolbar buttons in MainWindow):
  `CornerRadius` 8, `Padding` 10,5, `MinWidth` 86, `ControlBackground` bg, `ControlBorder`
  1px; `:pointerover` and `:pressed` invert to `ControlPressedBackground` /
  `ControlPressedForeground`. Matches the Windows toolbar treatment.
- **Menu** (File/Edit/View/Navigate): background/foreground from theme, rounded flyout,
  `SurfaceBackground` popup with `ControlBorder`.
- **Workspace tabs** (`TabControl#LeftDockTabs` / `TabItem`): strip Fluent chrome; inactive
  = `MutedText`; selected = `HeaderText` + SemiBold + bottom-border accent
  (`ControlAccent`, 2–3px, `CornerRadius` 4,4,0,0); `:pointerover` = `SurfaceMutedBackground`.
- **Status bar**: `Border` with top border (`SurfaceBorder`), consistent padding, section
  text in `SecondaryText`/`MutedText`.

### 3. Rebuild the outline cards (the most visible "halfway" element)

In `MainWindow.axaml`, replace the minimal outline/notes item `Border` templates
(~222–299) with proper cards matching Windows:

- Rounded `Border` (`CornerRadius` 8), `CardBackground` bg, `SurfaceBorder` 1px,
  vertical margin ~`0,0,0,8`, padding ~`8,10`.
- A small **type label** (Act / Sequence / Scene) in `MutedText`, ~10.5px, SemiBold,
  letter-spaced — replacing the current blue-accent label.
- **Heading** in `HeaderText`, ~14px SemiBold, wrapping.
- **Act level**: thick left border bar (`8,0,0,1`) using `HierarchyIndicatorBrush` to signal
  hierarchy, mirroring the Windows act card.
- Replace the hardcoded `#253584E4` drag-over fill with the themed `DragOverBackground`.

Beat board cards (~465–576) only need the hardcoded green save button (`#2E7D32`) reviewed
for theme consistency and otherwise inherit the new palette automatically.

### 4. Sweep for remaining hardcoded colors

Grep the Linux project for literal hex (`#`) in `.axaml`/`.cs` outside the palette
definitions and the Fountain syntax colorizer (which is intentionally left alone), and route
them through theme resources.

## Verification

1. `dotnet build Passage/Passage.App.Linux` — must compile clean.
2. Run the Linux app (via the `/run` skill or `dotnet run --project
   Passage/Passage.App.Linux`). Confirm:
   - Toolbar buttons have rounded backgrounds + hover/press inversion (no stock Fluent look).
   - Workspace tabs show an active-tab underline + bold, no blue.
   - Outline shows real cards (rounded, bordered, type label + heading; act has left bar).
   - Status bar reads as styled chrome.
   - **No blue anywhere** in the chrome/workspace.
3. Toggle **View → Light / Dark**: confirm cream/paper light and near-black dark, and that
   they look like direct inverses; confirm theme switching still works at runtime.
4. Screenshot both themes and compare side-by-side against the Windows screenshots for
   palette/spacing/feel parity.
5. Confirm the editor surface and Fountain syntax colors are unchanged.
