# Architecture Decisions

> ADR (Architecture Decision Record) Log

## Mapped Decisions

### ADR-001: Direct PDF Stream Generation
- **Context:** Exporting to standard industry screenplay format requires precise Courier layout controls without bloated runtime dependencies.
- **Decision:** Write direct PDF-1.4 stream structures (objects, page dictionary references, binary content streams) inside `ScreenplayPdfExporter`.
- **Status:** Approved
- **Consequences:** Highly efficient and dependency-free, but requires low-level stream offset calculations.

### ADR-002: WPF Presentation Layer for Windows Integration
- **Context:** The application target environment is Windows native desktops with dynamic editor syntax helper requirements.
- **Decision:** Build client GUI using WPF target framework `net9.0-windows` for rich window adorners.
- **Status:** Approved

## Phase 5: UI Redesign Decisions

**Date:** 2026-05-28

### Scope
- **Theme Conservation:** Retain and polish the high-contrast monochromatic "E-Reader/E-Ink" theme style. No third theme.
- **Redesign Targets:**
  - Redesign the Workspace panel and its tabs for a modern Windows 11 / 2020s look.
  - Polish the Goal tab's timer.
  - Redesign all buttons and banners.
  - Editor Workspace: Make the pages themselves pure white (Light) / pure black (Dark) while the background container/gaps are a distinct contrasting shade (e.g., light gray / charcoal) to visually separate pages from the frame.
- **OS Title Bar:** Leave standard OS title bar as is.

### Approach
- **Option B chosen:** Revamp both theme resources (`DarkTheme.xaml`/`LightTheme.xaml`) and UI control templates (`MainWindow.xaml`) to use rounded corners, custom scrollbars, clean hover states, and premium typography while maintaining the core monochromatic e-ink design values.

