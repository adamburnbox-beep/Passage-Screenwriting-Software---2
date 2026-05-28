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
