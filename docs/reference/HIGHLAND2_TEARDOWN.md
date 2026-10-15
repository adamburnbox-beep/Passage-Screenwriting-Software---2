# Highland 2 Teardown → Passage Feature Spec

Source: `Highland 2.app` (Quote-Unquote Apps, Mac App Store build, universal x86_64/arm64 Swift/AppKit binary, module `Highland`).
Method: static analysis only — rea Interface Builder decoder (menus, outlets, actions), Mach-O string/symbol extraction (404 Swift types, 60 source-file names, preference keys), and unpacking the bundled sample documents, themes and cheat sheets. No decompilation (rea's Ghidra/Hopper/IDA backends aren't configured on the Legion).

**Use this as a behaviour spec, not a code source.** Do not copy Highland code, assets, sample documents, fonts it licenses, or branding into Passage. Open formats (Fountain, TextBundle, FDX) and industry-standard layout are fair game. Everything below describes *what* Highland does and *what formats it reads/writes*. Implement it fresh in Passage.

---

## 1. Headline takeaways for Passage

1. **Plain text + sidecar JSON is the whole model.** The document is a Fountain/Markdown string; every "smart" feature (characters, sprints, navigator filters, revisions, milestones, cursor position) lives in small JSON files beside it. Copy this architecture — it keeps Passage files diffable and future-proof.
2. **`.highland` = a zipped TextBundle.** Supporting import/export of it costs little and lets Highland users move to Passage without friction.
3. **Directives (`{{HEADER: …}}`, `{{TOC}}`, `{{INCLUDE: …}}`) give power-user features with zero UI.** Cheap to build in a web app, high perceived value.
4. **Writer-motivation tools are first-class:** sprints, goals, milestones, reading time, gender/dialogue analysis, character highlighting.
5. **Gap Highland admits it has:** no locked pages, no production revisions (A-pages, revision sets per colour). That's Passage's clearest differentiation opportunity.

---

## 2. File format: `.highland`

A `.highland` file is a **zip** containing one TextBundle directory (`org.textbundle.package`):

```
MyScript.highland (zip)
└── MyScript.textbundle/
    ├── text.fountain        # or text.md for Markdown documents
    ├── info.json            # TextBundle metadata + app state
    ├── characters.json      # per-character metadata (gender, highlight colour)
    ├── navigatorFilters.json
    ├── sprints.json
    └── assets/              # images / included files referenced from the text
```

Other sidecars referenced in the binary: `characterHighlighting.json`, plus Codable types for **Revisions, Milestones, Bin, Scratchpad, Includes** (stored the same way).

**info.json** (real example):
```json
{
  "com.quoteunquoteapps.highland2": {
    "lastScrollPosition": 190.5,
    "printParagraphNumbers": false,
    "cursorPosition": 0,
    "templateName": "Screenplay",
    "currentTabIndex": -1,
    "showSynopses": false,
    "sidebarWidth": 264
  },
  "transient": false,
  "type": "com.quoteunquoteapps.fountain",
  "version": 2
}
```
Note the pattern: standard TextBundle keys at top level, app-specific state namespaced under a reverse-DNS key. Passage should do the same (`"app.passage": {...}`) so it can round-trip a Highland bundle without clobbering Highland's keys.

**navigatorFilters.json**:
```json
{"images":true,"includedFiles":true,"referenceLinks":true,"markers":false,
 "sectionHeaders":true,"notes":true,"links":true,"sceneHeaders":true,"synopses":true}
```

Highland also opens bare `.fountain`, `.md`, `.txt`, and has "Convert to .highland…" to upgrade a plain file into a bundle.

**Passage action:** implement `.highland` import (unzip → read `text.*` → map `characters.json`) and export (write the same structure). JSZip in the browser is enough.

---

## 3. Syntax Highland recognises

Fountain core plus Highland extensions. Element types from the source-file list: Action, SceneHeader, CharacterCue, Dialogue/Parenthetical, Transition, Lyrics, Synopsis, SectionHeader, Note, PageBreak, HorizontalRule, Blockquote, ForceCenter, ForceLeft, ForceRight, Highlight, Strikethrough, OrderedList, UnorderedList, CheckList, LinkInline, LinkReference, EndnoteInline, EndnoteReference, Directive.

| Element | Rule (as observed) |
|---|---|
| Scene header | Line starts `INT`, `EXT`, `EST`, `I/E`, `E/I` followed by `.` or space (case-insensitive); force with leading `.` (but not `..`) |
| Scene number | `#…#` at end of scene header, e.g. `EXT. MEADOW - DAY #3#` |
| Character cue | Line with no lowercase letters (may start with digits), optional `(V.O.)`/`(CONT'D)`-style extensions, **followed immediately by a non-blank line**; force with `@` |
| Dual dialogue | `^` after the second character's name |
| Parenthetical | Line wrapped in `( )` inside a dialogue block |
| Transition | Uppercase line ending `TO:`, or exactly `FADE IN:`, `FADE OUT.`, `FADE TO BLACK.`, `CUT TO BLACK.`; force with `>` |
| Centred | `> text <` |
| Right-aligned | `text >` |
| Lyrics | Line starts `~` |
| Forced action | Line starts `!` |
| Section | `#` … `######` (hidden in screenplay output, rendered as headlines in Markdown templates) |
| Synopsis | Line starts single `=` (not `==`) |
| Note | `[[ … ]]` — inline or multi-line, never printed unless option enabled. Also used for **alternate lines** of dialogue |
| Boneyard | `/* … */` |
| Page break | `===` (three or more) alone on a line |
| Horizontal rule | `---`, `___`, `***` alone on a line |
| Emphasis | `*italic*`, `**bold**`, `***bold italic***`, `_underline_`, `+highlight+`, `~~strike~~` — all backslash-escapable |
| Lists | `-`, `*`, `+`, `1.`; checklists `[ ]` / `[x]` |
| Term list | `Key :: Value` (two-column; leave key empty to continue) |
| Tables | Pipe tables; alignment markers (`> c <`, `r >`) work inside cells |
| Links | Inline `[t](url)`, reference `[t][ref]` + `[ref]: url`, images `![alt](path)`, naked URLs auto-detected |
| Endnotes | `[^id]` inline + `[^id]: text` |
| Graphic novel | Panels = forced scene headers (`.PANEL 1`); pages = `#` sections |

## 4. Directives

| Directive | Behaviour |
|---|---|
| `{{HEADER: … }}` / `{{FOOTER: … }}` | Takes effect from that page onward. `|` splits into left/centre/right slots (`{{HEADER: Title | DRAFT | %p}}`). Headers default right-aligned, footers centred. Option to alternate alignment on odd/even pages |
| `%p` / `%pp` | Current page / total pages |
| `%#`, `%##` … | Text of nearest section header at that level |
| `%none` | Turn header/footer off |
| `{{PAGENUMBER: n}}` | Reset page numbering |
| `{{SERIES: n}}` | Auto-incrementing counter (chapter numbers, scene numbers) |
| `{{TOC}}` / `{{TOC: # ###}}` | Table of contents from section headers, optionally filtered by level |
| `{{INCLUDE: path}}` | Pull another text file in at preview/export time ("Assemble…" flattens includes into one document) |
| `{{WATERMARK: …}}` | Page watermark |
| `{{%m}}` / `{{%m: name}}` | Markers (unnamed / named) used by Tools → Insert Marker / Jump to… |

---

## 5. Full feature inventory (from the main menu)

**File:** New, New Plain Text Document, Convert to .highland, **Convert to Narrated Script**, Save As PDF, Export to `.fountain` `.markdown` `.fdx` `.txt` `.rtf` `.html` `.docx` `.textbundle`, **Include…**, **Assemble…**, **Compare…** (diff two drafts), Page Setup, Print.
**Import (from type names):** FDX, FDR (legacy Final Draft), PDF → text (own `QUPDFParser` framework).
**Edit:** Cut Sentence, Cut Paragraph, Copy as Rich Text, Copy as HTML, Paste as Markdown, Cut to Bin, Copy to Bin, Quick Find, Find & Replace.
**Format:** Bold, Italic, Underline, Highlight, Strikethrough, Center, **Omit** (boneyard), Title Page…, Insert Note / Endnote / Reference Note, Page Break, Header, Footer, Watermark, Numbered/Bulleted List, Force Scene Header, Force Transition, Bigger/Smaller, Insert Link / Image Link, Convert Links Inline ↔ References, Remove Formatting.
**View:** Go to Preview (⌘E), **Live Margins**, Auto-Hide Toolbar, **Typewriter Scrolling**, Show Invisibles, Show Page Breaks, Show Paragraph Numbers, Status Bar, Sidebar panes (⌘1–⌘5: Navigator, Bin, Statistics, Assets, Scratchpad), Sessions.
**Tools:** Insert Marker / Named Marker, Next Marker, Remove Markers, Jump to…, Add / Remove Scene Numbers, **Gender Analysis**, **Word Analysis** (unique/overused words), **Character Highlighting**, **Milestones & Sprints**, Revisions submenu.
**Revisions:** Revision Mode toggle, colour (Red / Green / Blue / None), Mark Revision, Clear Revision, Clear All, Go to Next Revision. Print options: revision colour, revision stars (margin asterisks), **print revisions only**.
**Help:** Fountain & Markdown cheat sheets, Getting Started (interactive tour with short videos), What's New.

## 6. Writer tools — behaviour notes

- **Sprints:** timed sessions picked from a duration list; progress bar in toolbar; completion sound; history log (`sprints.json`, "View All History"); optional share message. Warns before quitting mid-sprint.
- **Goals:** word or page goal per document with progress ring; reset; past goals kept.
- **Milestones:** log of word-count checkpoints per file with date (table: filename / count / date).
- **Statistics pane:** words, pages, **reading time** (`wordsPerMinute` setting), advanced word count.
- **Gender Analysis:** each character tagged M/F/other in `characters.json`; report counts lines and words per gender and per character; exportable as `gender_report.pdf`.
- **Character Highlighting:** pick a colour per character; all their dialogue gets tinted. Palette of 40 pastel swatches (e.g. `#c8dcf2`, `#fee485`, `#cbe2a1`, `#e1cde4`, `#f5baba`…).
- **Character panel:** auto-extracted character list, **rename character everywhere**, colour picker.
- **Word Analysis:** tokenises the script, lists unique / repeated words, click to highlight occurrences.
- **Bin:** drag-and-drop shelf for text snippets and scenes, survives with the document.
- **Scratchpad:** free-text notes stored with the document but outside the script.
- **Assets:** images/files referenced by the document; drag back in to insert correct syntax.
- **Navigator:** outline of scene headers, sections, synopses, notes, markers, links, images, includes — each type toggleable (the filter JSON above).
- **Autocomplete:** character names, scene-header terms, locations, time of day (`DAY`, `NIGHT`, `MORNING`, `DUSK`, `THE NEXT DAY`…), transitions; custom terms list. Auto-close parentheses.
- **Automatic (MORE)/(CONT'D)** on dialogue split across pages; option to break pages only at sentence boundaries.
- **International:** US ↔ UK spelling conversion from two parallel word lists (~1,500 pairs each).
- **Backups:** timed backups to a chosen folder, optional plain-text copies.

## 7. Templates, print & layout options

**Templates:** Screenplay, Modern Standard, Stage Play, Multi-cam, Graphic Novel, Treatment, Novel, Manuscript, MLA, Article, Bulletin, Guidebook, Journal, Speech, Wiki, Plain Text. Each is a named style set (fonts, element margins/widths, line spacing). Free tier: Screenplay + Modern Standard only.

**Print/layout preference keys found** (good checklist for Passage's export dialog):
`boldSceneHeaders`, `doubleSpaceSceneHeaders`, `printSceneHeaders`, `printSectionHeaders`, `printSynopses`, `italicizeSynopses`, `printInlineNotes`, `includeNotes`, `includeSecondaryText` (alt lines), `printTitlePage`, `printPageNumberOnFirstPage`, `printParagraphNumbers`, `printLineNumbers`, `printRevisionColor`, `printRevisionStars`, `printRevisionsOnly`, `automaticMoreAndContinued`, `breakPagesAtSentenceBoundary`, `dualDialogueWidthFactor`, `alternateHeaderAlignmentBasedOnPageNumber`, `addSpaceBeforeSectionHeaders`, `breakBeforeSectionHeaders`, `indentFirstLineOfParagraph`, `capitalizeFirstNewParagraphLine`, `bodyFont`, `headingFont`, `lineSpacing`, `paperSize`.

**Bundled fonts:** Courier Prime (+ a "ManuscriptCourier" variant), plus Lato, Merriweather, Open Sans, PT Serif, Rubik, Source Sans Pro, Spectral, Cormorant Garamond, TeX Gyre Schola, Amiri (Arabic), New York, SF Pro. All but SF/New York are open-licence and usable in Passage.

**Rendering pipeline (from type names):** `Tokenizer → FountainParser/MarkdownParser → Element[] → ElementFormatter → ElementHeightCalculator → Paginator (+EndnotePaginating) → PDFBuilder / PDFPresenter`, with writers `HLFDXWriter`, `DOCXWriter`, `RTFWriter`, `HTMLWriter`. Long jobs (word counts, scene numbering, PDF build, unique-word tokenising) run as background Operations. Worth mirroring with Web Workers in Passage.

## 8. Themes

Theme = 12-key JSON. Passage can adopt the identical schema and import users' Highland themes directly:

```json
{
  "name": "Classic", "mode": "light",
  "background_color": "#ffffff", "text_color": "#2d333b",
  "sidebar_background_color": "#3f4752", "sidebar_text_color": "#c4cace",
  "theme_primary": "#b36795", "theme_secondary": "#2faad1",
  "inactive": "#9bacba", "scene_header_color": "#3f4752",
  "notes_color": "#499C91", "insertion_point_color": "#f85e66"
}
```
Built-ins: Classic, Clean, Dark Mode, Dracula, Kermit, Material, Monokai, Solarized Dark, Solarized Light, Transit, Winter. Users can add custom themes in Preferences.

---

## 9. Recommended Passage roadmap

**Tier 1 — foundation (do first)**
- [ ] Sidecar-JSON document model; namespaced app state in `info.json`
- [ ] `.highland` + `.fountain` + `.textbundle` import/export
- [ ] Full syntax table in §3 (esp. `+highlight+`, `~~strike~~`, `[[notes]]` incl. alt lines, `::` term lists, checklists, endnotes)
- [ ] Navigator with per-element-type filters
- [ ] Theme JSON schema (§8) + 4–5 built-ins incl. a dark theme
- [ ] Export options dialog driven by the §7 key list

**Tier 2 — writer delight**
- [ ] Sprints + goals + reading time in a status bar
- [ ] Character panel: auto-extract, rename-everywhere, highlight colour, gender tag
- [ ] Dialogue/gender analysis report
- [ ] Markers (`{{%m: name}}`) + Jump to…
- [ ] Bin and Scratchpad panes
- [ ] Typewriter scrolling, Live Margins toggle, autocomplete (names, locations, times, transitions)
- [ ] Automatic (MORE)/(CONT'D)

**Tier 3 — power features**
- [ ] Directives: HEADER/FOOTER with slots, PAGENUMBER, SERIES, TOC, INCLUDE + Assemble
- [ ] Revision mode with colours, stars, revisions-only print
- [ ] Compare two drafts (side-by-side diff)
- [ ] FDX import/export; PDF → Fountain import
- [ ] Word analysis (repeated/unique words)

**Differentiators Highland lacks** (Passage can own these): locked pages & A-pages, proper production revision sets (WGA colour sequence), real-time collaboration (web-native), scene cards/corkboard view of the same plain text.

---

# Part 2 — Deep dive (Ghidra 12.1.3 decompilation + Swift reflection metadata)

Method: x86_64 slice decompiled headless with Ghidra; Swift class ivars resolved by name; enum cases and struct fields read from the binary's `__swift5_fieldmd` reflection metadata. Everything below is behaviour and data, restated for a clean-room implementation.

## 10. Rendering pipeline (confirmed)

```
text ─► Tokenizer ─► FountainParser ─► [Element] ─► ElementGrouper ─► [Group] ─► Paginator ─► [Page] ─► PDFPresenter
```

- **FountainParser** state: `sceneHeaderTerms`, `transitionTerms`, `counters` (for `{{SERIES}}`), `boneyards`, `linkReferences`, `includes`/`includedAssets`, `titlePage`/`titlePageRange`, `wordCount`, `runningCount`, `previousTokenType`.
- **Page** holds: `elements`, `directives`, `footnotes`, `pdfPage` (imported PDF pages can be interleaved), `spaceLeftOnPage`.
- **Paginator** state: `pageMetrics` (the template), `printOptions`, `grouper`, `pages`, `currentPage`, `currentPageNumber`, `y`, `pageBreakLocations`, `footnotes`, `globalDirectives`, `revisionDeletionIndexes`, `charactersToHighlight`, `suppressPageBreak`, `suppressSplitting`, `skipPagination`, `previousSectionHeaderBottomPadding`.

### Element types (exact enum, 32 cases)
`action, sceneHeader, transition, centered, sectionHeader, synopsis, pageBreak, paragraphBreak, lyric, characterCue, dialogue, parenthetical, dualDialogueCharacterCue, boneyard, lineBreak, html, blockquote, orderedListItem, unorderedListItem, horizontalRule, inlineImage, codeBlock, tableRow, directive, termList, forcedRight, forcedLeft, toc, note, endNote, footnote, pdf`

Element fields: `type, text, attributedString, range, level, sceneNumber, keyword, title, term/definition` (term lists), `rowValues/isHeaderRow/suppressRowDivider` (tables), `itemNumber` (lists), `noteId/noteKeyword`, `url/imageSize`, `pageIndexOfFirstUse/referencedPageNumber` (endnotes), `childSectionHeader`.

Passage action: adopt this as the canonical AST. Keeping blank lines (`paragraphBreak`, `lineBreak`) as real elements is what makes "compress blank lines" and page-break swallowing straightforward.

## 11. Grouping (keep-together rules)

The grouper walks elements and builds unbreakable **groups** before pagination:

- `compressBlankLines`: consecutive `paragraphBreak`s collapse to one.
- A `pageBreak` swallows the paragraph break that follows it (`swallowNextParagraphBreak`); `lineBreak` can be swallowed the same way.
- `sectionHeader`: a header at a shallower-or-equal level closes the current group; headers stay with what follows.
- Dialogue: `characterCue` opens a dialogue block (`openDialogueBlock`); `parenthetical` and `dialogue` join it; `keepElementWithNext` glues scene headers and cues to the next element. Dual dialogue counts cue pairs (`dualDialogueCount`) so both sides land in one group.
- Notes and synopses only enter groups when `includeNotes` / `includeSynopses` are set; otherwise they're dropped before layout.
- Imported `pdf` pages always form their own group.
- Template-specific branch for **MLA** (different paragraph grouping).

## 12. Pagination algorithm

For each group:
1. Measure its height (`ElementHeightCalculator` → text layout at the element's width).
2. If `spaceLeftOnPage − height + tolerance ≤ 0` → **split or push to next page**; otherwise place it and subtract.
3. `tolerance` comes from the template field **`cheatPageBreakByAmount`**, which lets a block overflow the bottom margin slightly rather than break. It's a nice trick for keeping one-line orphans off the next page.

Splitting dialogue (when `automaticMoreAndContinued` is on and `suppressSplitting` is off):
- Bottom of page: cue line gets **`(MORE)`**.
- Top of next page: repeat the character cue + **` (CONT’D)`**: with a *curly* apostrophe (U+2019) and a leading space.
- Both strings are user-overridable (`useCustomTerms`, `customMoreString`, `customContinuedString`).
- `breakPagesAtSentenceBoundary` (template field) restricts splits to sentence ends.

Page numbering: pages with no elements, no footnotes and no PDF page are skipped. A `{{PAGENUMBER}}` directive on a page resets the counter. Header/footer directives are stored per page and inherited forward.

## 13. Editor-side CONT'D automation

- When a character cue is typed, the editor strips `(V.O.)`, `(O.S.)` and `(CONT’D)` from both the new and the previous cue, compares bare names, and offers/inserts `(CONT’D)` if the same character is speaking again after action.
- Typing `(CONT` triggers completion to `(CONT’D)`.
- Character highlighting is applied in the syntax highlighter (`characterNamesToHighlight`, `characterHighlightingTextColor`), not in the model.

## 14. Template system

Templates are **14 Swift structs** sharing one 55-field shape (+ `BaseTemplate` defaults): `Screenplay, MultiCam, StagePlay, GraphicNovel, Manuscript, Novel, Treatment, ModernStandard, Article, Bulletin, Speech, MLA, Wiki, PlainText`.

Template picker groups: **Scripts** (Screenplay, Multi-Cam, Stage Play, Graphic Novel) · **Literary** (Manuscript, Novel, Treatment) · **Professional** (Modern Standard, Article, Bulletin, Speech) · **Academic** (MLA Report, Wiki) · **Unformatted** (Plain Text).

### Template fields (port this as a JSON schema)
```
name, summary, format (fountain|markdown|raw), font, userSelectableOptions,
pageSize, margins, lineHeightMultiple, cheatPageBreakByAmount,
enableTextStyling, convertLinks, breakBeforeSectionHeaders, breakPagesAtSentenceBoundary,
justifyBodyText, mergeParentheticalsIntoDialogue, indentFirstLineOfParagraph, firstLineIndentWidth,
printParagraphBreaks, alternateHeaderAlignmentBasedOnPageNumber, printPageNumberOnFirstPage,
printSectionHeaders, printSceneHeaders, capitalizeFirstNewParagraphLine,
defaultElementWidth, defaultElementLeftMargin,
elementWidth{ElementType:fraction}, elementLeftMargin{ElementType:fraction}, elementStyles{ElementType:traits},
dialogueLineHeightMultiple, dualDialogueWidthFactor,
blockquoteFont, blockquoteColor,
defaultSectionHeaderAlignment, sectionHeaderAlignments[], sectionHeaderStyles[], sectionHeaderLineHeightMultiples[],
sectionHeaderPadding[], sectionHeaderFonts[], sectionHeaderColors[],
codeBackgroundColor, highlightBackgroundColor,
defaultHeader, defaultFooter, headerFooterFont, headerStyles, footerStyles, headerColor, footerColor,
headerAlignment, footerAlignment,
tableBorderColor, shouldDrawTableBorder, tableCellPadding,
titlePageTitleFontSize, titlePageTitleStyles
```
Key design point: **element geometry is stored as fractions of page width**, so one template works on Letter and A4 with no extra tables.

### Recovered element geometry (default dialogue-style template)
Fractions × 612 pt (US Letter width). Left margin is the offset from the body text's left edge.

| Element | Left offset | Width |
|---|---|---|
| Character cue | 0.2304 → 141 pt (1.96 in) | 0.6536 → 400 pt |
| Dual-dialogue cue | 0.2304 → 141 pt | 0.4069 → 249 pt |
| Dialogue | 0.1160 → 71 pt (≈1.0 in) | 0.4069 → 249 pt (3.46 in) |
| Lyric | 0.1160 → 71 pt | 0.4069 → 249 pt |
| Parenthetical | 0.1618 → 99 pt (1.375 in) | 0.3448 → 211 pt (2.93 in) |
| Scene header | (default) | 0.6536 → 400 pt |
| Term list | (default) | 0.3660 → 224 pt |
| Blockquote | 0.0588 → 36 pt (0.5 in) | (default) |
| TOC | 0.0294 → 18 pt | (default) |

Section header sizes in that template: 24 / 18 pt; body 13 pt; highlight background `#bcebc0`. Screenplay itself uses **Courier Prime 12 pt**.

### Print options (exact set)
`printSynopses, printInlineNotes, printParagraphNumbers, printRevisionStars, printRevisionColor, printRevisionsOnly, printTitlePage`

## 15. Other recovered structures

- **Navigator outline item types:** root, blank, misc, sectionHeader, synopsis, sceneHeader, note, includedFile, link, referenceLink, image, marker. Filters: sectionHeaders, synopses, sceneHeaders, notes, includedFiles, links, referenceLinks, images, markers.
- **Theme model:** name, mode, textColor, backgroundColor, themePrimaryColor, themeSecondaryColor, sidebarBackgroundColor, sidebarTextColor, inactiveColor, sceneHeaderColor, notesColor, insertionPointColor.
- **Editor view settings:** columnWidth/useColumnWidth, lineSpacing, useTypewriterScrolling, highlightCurrentLine + lineHighlightColor, suppressBlinking, autoCloseParentheses, insertionPointWidth, canResizeMargins, liveMargins, markerGutter, showInvisibles, showPageBreaks.
- **Gender report** (PDF): (1) Characters: count by Male / Female / Unspecified; (2) Dialogue: lines spoken by gender; (3) words spoken by gender; (4) Interactions: gender-pair interaction counts; then a per-character table Name / Lines / Words. Chart palette `#356878`, `#2faad1`, `#846a7d`.
- **Title page keys** recognised: title, author/authors, draft date, date (+ standard Fountain keys).
- **Revisions** are tracked as index sets on the text (`revisionIndexes`, `deletions`); printing can mark them with colour, margin stars, or output revised pages only.

## 16. Updated Passage priorities from the deep dive

1. **Model:** 32-type Element AST + Group + Page objects as above.
2. **Template JSON** using the §14 schema, geometry as page-width fractions; ship Screenplay first with the §14 numbers.
3. **Paginator:** keep-together grouping, tolerance (`cheatPageBreakByAmount`), MORE/CONT'D splitting with overridable strings, sentence-boundary option.
4. **Editor:** auto-(CONT'D) on repeated speaker; `(CONT` completion.
5. **Reports:** gender/dialogue report with the four sections listed.
