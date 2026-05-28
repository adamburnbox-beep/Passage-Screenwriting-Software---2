# SPEC.md — Project Specification

> **Status**: `FINALIZED`
>
> ⚠️ **Planning Lock**: No code may be written until this spec is marked `FINALIZED`.

## Vision
Passage is a Windows-native screenwriting desktop application designed for writing and formatting screenplays using the plain-text Fountain markdown language. It aims to provide screenwriters with a lightweight, distraction-free environment that combines the simplicity of markdown text editing with powerful industry-standard layout builders, page estimation tools, writing goals (word count, page count, timer targets), outline navigation, and automated session recovery.

## Goals
1. **Fountain Markup Parsing** — Accurately parse the Fountain screenplay format including scene headings, action, character cues, dialogue, parentheticals, transitions, sections, synopses, lyrics, centered text, and comments/notes.
2. **Page & Metric Estimation** — Provide real-time estimates of screenplay page lengths, word counts, and formatting dimensions based on industry standard formatting metrics (e.g., Courier 12pt, specific margins).
3. **Writing Goals & Session Timers** — Let writers set session goals for target word counts, page counts, or focused writing durations (with system timers).
4. **Outline & Beat Board Navigation** — Enable navigation through screenplay sections/scenes and visually organize ideas/cards using a beat board presentation.
5. **Lossless PDF Exporting** — Build and stream standard PDF-1.4 documents matching standard screenwriting layout specs without relying on external libraries.
6. **Session & Recovery Storage** — Cache working screenplay states periodically to prevent data loss in the event of unexpected application closures or crashes.

## Non-Goals (Out of Scope)
- Collaboration or multi-user editing features.
- Support for formats other than Fountain (e.g., Final Draft `.fdx` native file format imports/exports), except for PDF output.
- Rich-text editor formatting (the editor is plain-text, formatting is applied dynamically via the parser and layout builder).

## Constraints
- **Platform** — Must run natively on Windows using .NET 9.0 and WPF (Windows Presentation Foundation).
- **No External Dependencies** — Core parsing, page estimation, and PDF export must run on pure .NET APIs without third-party libraries.
- **Strict Fountain Spec** — The parser must comply with the canonical Fountain syntax specification.

## Success Criteria
- [x] Correctly parse Fountain screenplays and build internal models of elements.
- [x] Calculate word counts and estimate page counts dynamically.
- [x] Export screenplay pages to clean, standard PDF-1.4 files.
- [x] Maintain UI state and recover document state via local session storage.
- [x] Run and pass all unit tests successfully.

---

*Last updated: 2026-05-28*
