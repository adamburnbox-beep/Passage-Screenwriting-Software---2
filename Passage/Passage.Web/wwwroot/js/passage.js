// Client-side half of the Passage web editor.
//
// The editor is CodeMirror 5 (vendored in lib/codemirror), so typed text
// renders instantly in the browser. Content changes are debounced and pushed
// to the Blazor circuit, which re-parses with the shared Fountain pipeline and
// pushes back per-line element classes. Those classes drive both syntax
// colour and screenplay indentation (character/dialogue/parenthetical
// margins, right-aligned transitions, centred text) via CSS on the line —
// the same "style the line, leave the text alone" approach the desktop
// editor's FountainIndentationGenerator takes. A version counter guards
// against stale class batches arriving after further edits.
window.passage = (function () {
    let dotnetRef = null;
    let cm = null;
    let version = 0;
    let inputTimer = null;
    let caretTimer = null;
    let lastCaretLine = -1;
    let dirty = false;
    let appliedClasses = [];
    let sessionReady = false;
    let sessionTimer = null;
    let session = { fileName: "", caretLine: 1, editorFontPx: 15, previewZoom: 1.25, recentFiles: [] };

    const INPUT_DEBOUNCE_MS = 200;
    const SESSION_KEY = "passage.session.v1";
    const SESSION_DEBOUNCE_MS = 400;
    const RECOVERY_KEY = "passage.recovery.v1";
    const RECOVERY_INTERVAL_MS = 3000;
    const THEME_KEY = "passage.theme.v1";
    const LINE_CLASSES = [
        "sx-scene", "sx-character", "sx-dialogue", "sx-paren", "sx-transition",
        "sx-section", "sx-synopsis", "sx-note", "sx-boneyard", "sx-centered",
        "sx-lyrics", "sx-titlepage", "md-heading"
    ];

    function pushContent() {
        if (!dotnetRef || !cm) return;
        const caretLine = cm.getCursor().line + 1;
        dotnetRef.invokeMethodAsync("OnEditorChanged", cm.getValue("\n"), version, caretLine);
    }

    function scheduleInput() {
        version++;
        dirty = true;
        clearTimeout(inputTimer);
        inputTimer = setTimeout(pushContent, INPUT_DEBOUNCE_MS);
    }

    function reportCaret() {
        if (!dotnetRef || !cm) return;
        clearTimeout(caretTimer);
        caretTimer = setTimeout(() => {
            const line = cm.getCursor().line + 1;
            if (line !== lastCaretLine) {
                lastCaretLine = line;
                session.caretLine = line;
                scheduleSessionSave();
                dotnetRef.invokeMethodAsync("OnCaretMoved", line);
            }
        }, 120);
    }

    // Session restore lives in localStorage, not on the server: the data volume
    // is shared by every browser that opens the app, so a server-side "last
    // document" would leave two clients fighting over one value. The caret is
    // tracked here rather than pushed from Blazor so a moving cursor costs no
    // round-trips.
    function scheduleSessionSave() {
        if (!sessionReady) return;
        clearTimeout(sessionTimer);
        sessionTimer = setTimeout(() => {
            try {
                window.localStorage.setItem(SESSION_KEY, JSON.stringify(session));
            } catch (e) {
                // Private browsing, or the quota is full. Restore is a
                // convenience; losing it must not break the editor.
            }
        }, SESSION_DEBOUNCE_MS);
    }

    // Called once at startup. Saving stays disabled until this has run, so an
    // early cursor event cannot overwrite the stored state before it is read.
    function loadSession() {
        sessionReady = true;
        try {
            const raw = window.localStorage.getItem(SESSION_KEY);
            if (!raw) return null;
            const stored = JSON.parse(raw);
            if (!stored || typeof stored !== "object") return null;
            session = Object.assign(session, stored);
            return session;
        } catch (e) {
            return null;
        }
    }

    function setSessionDocument(fileName, editorFontPx, previewZoom, recentFiles) {
        session.fileName = fileName || "";
        session.editorFontPx = editorFontPx;
        session.previewZoom = previewZoom;
        session.recentFiles = Array.isArray(recentFiles) ? recentFiles : [];
        scheduleSessionSave();
    }

    // Crash recovery. Distinct from the file autosave in Editor.razor, which
    // writes the real file on the server; this is an unsaved-work safety net.
    // It is written client-side on a timer because it has to survive the server
    // dying, and a server-side snapshot needs the live circuit that a crash
    // takes away. localStorage also scopes it to one browser, so a recovered
    // draft is never offered to a different client.
    //
    // This deliberately never clears. A snapshot may be sitting in front of the
    // user in the recovery prompt, and a timer tick that wiped it would destroy
    // the very work being offered back. Clearing is explicit: a real save
    // (setDirty(false)), or Discard.
    function saveRecoverySnapshot() {
        if (!cm || !dirty) return;
        try {
            window.localStorage.setItem(RECOVERY_KEY, JSON.stringify({
                text: cm.getValue("\n"),
                fileName: session.fileName,
                savedAtUtc: new Date().toISOString()
            }));
        } catch (e) {
            // Private browsing or a full quota; the editor keeps working.
        }
    }

    function readRecoverySnapshot() {
        try {
            const raw = window.localStorage.getItem(RECOVERY_KEY);
            if (!raw) return null;
            const stored = JSON.parse(raw);
            if (!stored || typeof stored !== "object") return null;
            if (typeof stored.text !== "string" || stored.text.length === 0) return null;
            return stored;
        } catch (e) {
            return null;
        }
    }

    // Theme. Kept in its own key because App.razor reads it inline before first
    // paint, well before this file has loaded.
    function getTheme() {
        return document.documentElement.dataset.theme === "light" ? "light" : "dark";
    }

    function setTheme(theme) {
        const next = theme === "light" ? "light" : "dark";
        document.documentElement.dataset.theme = next;
        try {
            window.localStorage.setItem(THEME_KEY, next);
        } catch (e) {
        }
        // CodeMirror caches measurements against the old colours.
        if (cm) cm.refresh();
        return next;
    }

    function clearRecoverySnapshot() {
        try {
            window.localStorage.removeItem(RECOVERY_KEY);
        } catch (e) {
        }
    }

    // Editor keymap handlers. Zoom is server state (_editorFontPx drives a CSS
    // variable on the editor stack), so it round-trips — that is fine at one
    // call per keypress, unlike anything per-keystroke.
    function onSave() {
        if (dotnetRef) dotnetRef.invokeMethodAsync("OnSaveShortcut");
    }

    function zoomIn() {
        if (dotnetRef) dotnetRef.invokeMethodAsync("OnZoomShortcut", 1);
    }

    function zoomOut() {
        if (dotnetRef) dotnetRef.invokeMethodAsync("OnZoomShortcut", -1);
    }

    function goToLine() {
        if (dotnetRef) dotnetRef.invokeMethodAsync("OnGoToLineShortcut");
    }

    function goToScene() {
        if (dotnetRef) dotnetRef.invokeMethodAsync("OnGoToSceneShortcut");
    }

    function scrollIntoView(selector) {
        const el = document.querySelector(selector);
        if (el) el.scrollIntoView({ block: "nearest" });
    }

    function openFind() {
        if (dotnetRef) dotnetRef.invokeMethodAsync("OnFindShortcut", false);
    }

    function openReplace() {
        if (dotnetRef) dotnetRef.invokeMethodAsync("OnFindShortcut", true);
    }

    function selectedText() {
        return cm && cm.somethingSelected() ? cm.getSelection() : "";
    }

    function toggleSyntaxPanel() {
        if (dotnetRef) dotnetRef.invokeMethodAsync("OnSyntaxPanelShortcut");
    }

    // navigator.clipboard needs a secure context. This app is self-hosted on a
    // LAN over plain HTTP (docs/web-app.md), where that API is simply absent,
    // so fall back to the old selection-based copy rather than failing.
    async function copyText(text) {
        try {
            if (window.isSecureContext && navigator.clipboard) {
                await navigator.clipboard.writeText(text);
                return true;
            }
        } catch (e) {
            // Denied or unavailable; try the fallback below.
        }

        try {
            const field = document.createElement("textarea");
            field.value = text;
            field.setAttribute("readonly", "");
            field.style.position = "fixed";
            field.style.top = "-1000px";
            document.body.appendChild(field);
            field.select();
            const copied = document.execCommand("copy");
            field.remove();
            return copied;
        } catch (e) {
            return false;
        }
    }

    function init(reference) {
        dotnetRef = reference;
        const host = document.getElementById("editor-host");
        if (!host || typeof CodeMirror === "undefined") return;

        cm = CodeMirror(host, {
            value: "",
            mode: null,
            lineWrapping: true,
            placeholder: "INT. OPENING SCENE - DAY",
            viewportMargin: 50,
            // Ctrl-Z / Ctrl-Y / Ctrl-Shift-Z already come from CodeMirror's
            // default PC keymap, so they are not repeated here. Ctrl-N/O/W/Q
            // from the Linux set are reserved by the browser and cannot be
            // intercepted by a page — see docs/WEB-PARITY.md row 1.4.
            extraKeys: {
                "Ctrl-S": onSave,
                "Cmd-S": onSave,
                "Ctrl-=": zoomIn,
                "Cmd-=": zoomIn,
                "Shift-Ctrl-=": zoomIn,
                "Shift-Cmd-=": zoomIn,
                "Ctrl--": zoomOut,
                "Cmd--": zoomOut,
                "Ctrl-G": goToLine,
                "Cmd-G": goToLine,
                "Shift-Ctrl-G": goToScene,
                "Shift-Cmd-G": goToScene,
                "Ctrl-F": openFind,
                "Cmd-F": openFind,
                "Ctrl-H": openReplace,
                "Cmd-H": openReplace,
                "F1": toggleSyntaxPanel
            }
        });

        cm.on("change", (_, changeObj) => {
            if (changeObj.origin !== "setValue") {
                scheduleInput();
            }
            reportCaret();
        });
        cm.on("cursorActivity", reportCaret);
        cm.on("changes", updatePageRules);
        cm.on("refresh", updatePageRules);

        // Close any open dropdown menu after a choice or an outside click.
        document.addEventListener("click", (event) => {
            document.querySelectorAll("details.menu[open]").forEach((menu) => {
                if (!menu.contains(event.target) || event.target.closest(".menu-items")) {
                    menu.removeAttribute("open");
                }
            });
        });

        document.addEventListener("dragover", trackDropSide, true);

        setInterval(saveRecoverySnapshot, RECOVERY_INTERVAL_MS);

        window.addEventListener("beforeunload", (event) => {
            if (dirty) {
                event.preventDefault();
                event.returnValue = "";
            }
        });
    }

    function setLineClass(lineIndex, cssClass) {
        const previous = appliedClasses[lineIndex];
        if (previous === cssClass) return;
        if (previous) {
            cm.removeLineClass(lineIndex, "text", previous);
        }
        if (cssClass) {
            cm.addLineClass(lineIndex, "text", cssClass);
        }
        appliedClasses[lineIndex] = cssClass;
    }

    function applyHighlights(forVersion, classes) {
        if (!cm || forVersion !== version) return;
        const lineCount = cm.lineCount();
        cm.operation(() => {
            for (let i = 0; i < lineCount; i++) {
                setLineClass(i, (classes && classes[i]) || "");
            }
            appliedClasses.length = lineCount;
        });
    }

    // Re-apply line classes against the document already in the editor. Use this
    // wherever only the classification changed — Save As can flip screenplay and
    // markdown mode via the file extension without touching a character of text.
    // Routing that through setContent would clear the undo stack and throw the
    // caret back to line 1 for what is only a rename.
    function refreshHighlights(classes) {
        if (!cm) return;
        applyHighlights(version, classes);
    }

    function undo() {
        if (!cm) return;
        cm.undo();
        cm.focus();
    }

    function redo() {
        if (!cm) return;
        cm.redo();
        cm.focus();
    }

    // ---- Find / Replace ----
    //
    // Built on the vendored searchcursor addon. CodeMirror's own search.js is
    // deliberately not used: it has no whole-word option at all, and decides
    // case sensitivity from a smart-case heuristic rather than an explicit
    // checkbox, so it cannot reproduce the Avalonia dialog's option set.
    //
    // Whole word uses lookarounds rather than \b so it behaves the same when
    // the term itself starts or ends with a non-word character ("INT."), which
    // is what the Linux IsWholeWord does — it inspects the characters either
    // side of the match, never the term.
    const WORD_CHAR = "A-Za-z0-9_";

    function escapeRegExp(text) {
        return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    function buildQuery(term, matchCase, wholeWord) {
        if (!wholeWord) {
            return { query: term, options: { caseFold: !matchCase } };
        }
        const pattern = "(?<![" + WORD_CHAR + "])" + escapeRegExp(term) + "(?![" + WORD_CHAR + "])";
        return { query: new RegExp(pattern, matchCase ? "" : "i"), options: {} };
    }

    function searchFrom(term, matchCase, wholeWord, pos, forward) {
        const built = buildQuery(term, matchCase, wholeWord);
        const cursor = cm.getSearchCursor(built.query, pos, built.options);
        return cursor.find(!forward) ? cursor : null;
    }

    // Forward searches start at the end of the selection and wrap to the top;
    // backward starts at its head and wraps to the bottom. Same as FindText.
    function findInEditor(term, matchCase, wholeWord, forward) {
        if (!cm || !term) return false;
        const from = forward ? cm.getCursor("to") : cm.getCursor("from");
        let cursor = searchFrom(term, matchCase, wholeWord, from, forward);
        if (!cursor) {
            const wrapPos = forward
                ? { line: cm.firstLine(), ch: 0 }
                : { line: cm.lastLine(), ch: cm.getLine(cm.lastLine()).length };
            cursor = searchFrom(term, matchCase, wholeWord, wrapPos, forward);
        }
        if (!cursor) return false;

        cm.setSelection(cursor.from(), cursor.to());
        cm.scrollIntoView({ from: cursor.from(), to: cursor.to() }, 80);
        cm.focus();
        reportCaret();
        return true;
    }

    function findNext(term, matchCase, wholeWord) {
        return findInEditor(term, matchCase, wholeWord, true);
    }

    function findPrevious(term, matchCase, wholeWord) {
        return findInEditor(term, matchCase, wholeWord, false);
    }

    function selectionMatches(term, matchCase, wholeWord) {
        if (!cm || !cm.somethingSelected()) return false;
        const selected = cm.getSelection();
        const built = buildQuery(term, matchCase, wholeWord);
        if (built.query instanceof RegExp) {
            const anchored = new RegExp("^(?:" + built.query.source + ")$", built.query.flags);
            return anchored.test(selected);
        }
        return matchCase
            ? selected === term
            : selected.toLowerCase() === term.toLowerCase();
    }

    // Replace the current match if the selection is one, then advance —
    // matching ReplaceCurrent. Otherwise just advance to the first match.
    function replaceCurrent(term, replacement, matchCase, wholeWord) {
        if (!cm || !term) return { replaced: 0, found: false };
        if (selectionMatches(term, matchCase, wholeWord)) {
            cm.replaceSelection(replacement, "around");
            scheduleInput();
            const found = findNext(term, matchCase, wholeWord);
            return { replaced: 1, found: found };
        }
        return { replaced: 0, found: findNext(term, matchCase, wholeWord) };
    }

    function replaceAll(term, replacement, matchCase, wholeWord) {
        if (!cm || !term) return { replaced: 0, found: false };
        const built = buildQuery(term, matchCase, wholeWord);
        let replaced = 0;
        cm.operation(function () {
            const cursor = cm.getSearchCursor(built.query, { line: cm.firstLine(), ch: 0 }, built.options);
            while (cursor.findNext()) {
                cursor.replace(replacement);
                replaced++;
            }
        });
        if (replaced > 0) scheduleInput();
        return { replaced: replaced, found: false };
    }

    // Replace an inclusive range of lines in place. This is the whole point of
    // the ranged write-back: cm.replaceRange is a normal edit, so it joins the
    // undo history and leaves the caret alone, where setContent would clear
    // both (see WEB-PARITY 1.6).
    function replaceLineRange(startLine, endLine, text) {
        if (!cm) return false;
        const last = cm.lineCount() - 1;
        if (startLine < 0 || startLine > last) return false;

        const to = Math.min(endLine, last);
        cm.replaceRange(
            text,
            { line: startLine, ch: 0 },
            { line: to, ch: cm.getLine(to).length });
        scheduleInput();
        return true;
    }

    // Insert lines before lineIndex, or at the end of the document when the
    // index is past the last line. Ranged like replaceLineRange, so it keeps
    // the undo history.
    function insertLinesAt(lineIndex, text) {
        if (!cm) return false;
        const last = cm.lineCount() - 1;
        const at = Math.max(0, Math.min(lineIndex, last + 1));

        if (at > last) {
            const end = { line: last, ch: cm.getLine(last).length };
            cm.replaceRange("\n" + text, end, end);
        } else {
            const pos = { line: at, ch: 0 };
            cm.replaceRange(text + "\n", pos, pos);
        }

        scheduleInput();
        return true;
    }

    // Remove an inclusive line range, taking the line break with it so no blank
    // line is left behind.
    function deleteLineRange(startLine, endLine) {
        if (!cm) return false;
        const last = cm.lineCount() - 1;
        if (startLine < 0 || startLine > last) return false;

        const to = Math.min(endLine, last);
        let from = { line: startLine, ch: 0 };
        let until;

        if (to < last) {
            until = { line: to + 1, ch: 0 };
        } else {
            until = { line: to, ch: cm.getLine(to).length };
            if (startLine > 0) {
                from = { line: startLine - 1, ch: cm.getLine(startLine - 1).length };
            }
        }

        cm.replaceRange("", from, until);
        scheduleInput();
        return true;
    }

    // The desktop decides drop-before/drop-after from the target's horizontal
    // midpoint. The web board stacks cards in a column, so the vertical midpoint
    // is the meaningful equivalent. Tracked here because only the browser knows
    // the card's real geometry; Blazor asks once, on drop.
    let dropAfter = false;

    function trackDropSide(event) {
        const card = event.target && event.target.closest
            ? event.target.closest(".board-card")
            : null;
        if (!card) return;
        const rect = card.getBoundingClientRect();
        if (rect.height > 0) {
            dropAfter = (event.clientY - rect.top) > rect.height / 2;
        }
    }

    function dropIsAfter() {
        return dropAfter;
    }

    // ---- Page-break rules ----
    //
    // Ports ScreenplayPageRuler: a dashed rule every 55 lines with a "PAGE N"
    // pill, snapped to a line boundary so it sits between lines instead of
    // through them. Purely visual — the document is untouched, and this is the
    // same line-count approximation the page estimator uses; the Preview tab
    // stays the exact layout.
    //
    // Like the desktop editor this one wraps, so the boundary is every 55 line
    // *heights* of vertical space, not every 55 document lines — with wrapping
    // those diverge, and the desktop measures the visual one.
    const LINES_PER_PAGE = 55;
    let pageRuleLayer = null;
    let pageRulesEnabled = false;

    function ensurePageRuleLayer() {
        if (!cm) return null;
        // Re-created if CodeMirror ever rebuilds the sizer out from under it.
        if (pageRuleLayer && pageRuleLayer.isConnected) return pageRuleLayer;

        const sizer = cm.getWrapperElement().querySelector(".CodeMirror-sizer");
        if (!sizer) return null;

        pageRuleLayer = document.createElement("div");
        pageRuleLayer.className = "page-rules";
        sizer.insertBefore(pageRuleLayer, sizer.firstChild);
        return pageRuleLayer;
    }

    function updatePageRules() {
        if (!cm) return;
        const layer = ensurePageRuleLayer();
        if (!layer) return;

        layer.textContent = "";
        if (!pageRulesEnabled) return;

        const lineHeight = cm.defaultTextHeight();
        if (!(lineHeight > 0)) return;

        const pageHeight = LINES_PER_PAGE * lineHeight;
        const docHeight = cm.heightAtLine(cm.lastLine(), "local") + lineHeight;
        if (docHeight <= pageHeight) return;

        const totalPages = Math.ceil(docHeight / pageHeight);
        for (let page = 1; page < totalPages; page++) {
            const line = cm.lineAtHeight(page * pageHeight, "local");
            const rule = document.createElement("div");
            rule.className = "page-rule";
            rule.style.top = cm.heightAtLine(line, "local") + "px";

            const pill = document.createElement("span");
            pill.className = "page-rule-pill";
            pill.textContent = "PAGE " + (page + 1);
            rule.appendChild(pill);
            layer.appendChild(rule);
        }
    }

    // Screenplay mode only, matching ApplyEditorWriteMode, which adds the ruler
    // for screenplays and removes it for markdown.
    function setPageRules(enabled) {
        pageRulesEnabled = !!enabled;
        updatePageRules();
    }

    function setContent(text) {
        if (!cm) return version;
        version++;
        appliedClasses = [];
        cm.setValue(text);
        cm.clearHistory();
        cm.setCursor({ line: 0, ch: 0 });
        cm.scrollTo(0, 0);
        lastCaretLine = -1;
        dirty = false;
        return version;
    }

    function setDirty(value) {
        dirty = value;
        if (!value) {
            // Saved, so there is no unsaved work left to recover.
            clearRecoverySnapshot();
        }
    }

    function scrollToLine(line) {
        if (!cm) return;
        const target = Math.max(0, Math.min(line - 1, cm.lineCount() - 1));
        cm.setCursor({ line: target, ch: 0 });
        const coords = cm.charCoords({ line: target, ch: 0 }, "local");
        const scroller = cm.getScrollInfo();
        cm.scrollTo(null, Math.max(0, coords.top - scroller.clientHeight / 3));
        cm.focus();
        reportCaret();
    }

    async function exportDocument(format, name) {
        if (!cm) return;
        const baseName = (name && name.trim().length > 0 ? name : "Untitled")
            .replace(/\.(fountain|md|txt)$/i, "");

        if (format === "fountain") {
            const blob = new Blob([cm.getValue("\n")], { type: "text/plain" });
            triggerDownload(blob, baseName + ".fountain");
            return;
        }

        const form = new FormData();
        form.append("content", cm.getValue("\n"));
        form.append("format", format);
        form.append("name", baseName);
        const response = await fetch("api/export", { method: "POST", body: form });
        if (!response.ok) {
            throw new Error("Export failed: " + response.status);
        }
        const blob = await response.blob();
        triggerDownload(blob, baseName + "." + format);
    }

    function triggerDownload(blob, fileName) {
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = fileName;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        setTimeout(() => URL.revokeObjectURL(url), 5000);
    }

    function focusEditor() {
        if (cm) cm.focus();
    }

    // The CodeMirror instance is exposed for end-to-end tests.
    return {
        init, applyHighlights, setContent, setDirty, scrollToLine,
        exportDocument, focusEditor,
        loadSession, setSessionDocument,
        readRecoverySnapshot, clearRecoverySnapshot,
        refreshHighlights, undo, redo, copyText, scrollIntoView, replaceLineRange, insertLinesAt, deleteLineRange, dropIsAfter, setPageRules,
        findNext, findPrevious, replaceCurrent, replaceAll, selectedText,
        getTheme, setTheme,
        get editor() { return cm; }
    };
})();
