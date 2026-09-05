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
    let session = { fileName: "", caretLine: 1, editorFontPx: 15, previewZoom: 1.25 };

    const INPUT_DEBOUNCE_MS = 200;
    const SESSION_KEY = "passage.session.v1";
    const SESSION_DEBOUNCE_MS = 400;
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

    function setSessionDocument(fileName, editorFontPx, previewZoom) {
        session.fileName = fileName || "";
        session.editorFontPx = editorFontPx;
        session.previewZoom = previewZoom;
        scheduleSessionSave();
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
            extraKeys: {
                "Ctrl-S": () => { dotnetRef.invokeMethodAsync("OnSaveShortcut"); },
                "Cmd-S": () => { dotnetRef.invokeMethodAsync("OnSaveShortcut"); }
            }
        });

        cm.on("change", (_, changeObj) => {
            if (changeObj.origin !== "setValue") {
                scheduleInput();
            }
            reportCaret();
        });
        cm.on("cursorActivity", reportCaret);

        // Close any open dropdown menu after a choice or an outside click.
        document.addEventListener("click", (event) => {
            document.querySelectorAll("details.menu[open]").forEach((menu) => {
                if (!menu.contains(event.target) || event.target.closest(".menu-items")) {
                    menu.removeAttribute("open");
                }
            });
        });

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
        get editor() { return cm; }
    };
})();
