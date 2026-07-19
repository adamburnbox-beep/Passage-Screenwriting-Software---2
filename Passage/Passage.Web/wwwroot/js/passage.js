// Client-side half of the Passage web editor.
//
// The <textarea> is owned by this script (not data-bound) so typing stays
// native and instant. Content changes are debounced and pushed to the Blazor
// circuit, which re-parses with the shared Fountain pipeline and pushes back
// per-line syntax classes for the highlight overlay. A version counter guards
// against stale highlight batches arriving after further edits.
window.passage = (function () {
    let dotnetRef = null;
    let textarea = null;
    let overlay = null;
    let version = 0;
    let inputTimer = null;
    let caretTimer = null;
    let lastCaretLine = -1;
    let dirty = false;

    const INPUT_DEBOUNCE_MS = 200;

    function lineAtOffset(text, offset) {
        let line = 1;
        for (let i = 0; i < offset && i < text.length; i++) {
            if (text.charCodeAt(i) === 10) line++;
        }
        return line;
    }

    function escapeHtml(value) {
        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
    }

    function pushContent() {
        if (!dotnetRef || !textarea) return;
        const caretLine = lineAtOffset(textarea.value, textarea.selectionStart);
        dotnetRef.invokeMethodAsync("OnEditorChanged", textarea.value, version, caretLine);
    }

    function scheduleInput() {
        version++;
        dirty = true;
        clearTimeout(inputTimer);
        inputTimer = setTimeout(pushContent, INPUT_DEBOUNCE_MS);
    }

    function reportCaret() {
        if (!dotnetRef || !textarea) return;
        clearTimeout(caretTimer);
        caretTimer = setTimeout(() => {
            const line = lineAtOffset(textarea.value, textarea.selectionStart);
            if (line !== lastCaretLine) {
                lastCaretLine = line;
                dotnetRef.invokeMethodAsync("OnCaretMoved", line);
            }
        }, 120);
    }

    function syncScroll() {
        if (!textarea || !overlay) return;
        overlay.scrollTop = textarea.scrollTop;
        overlay.scrollLeft = textarea.scrollLeft;
    }

    function init(reference) {
        dotnetRef = reference;
        textarea = document.getElementById("editor-input");
        overlay = document.getElementById("editor-overlay");
        if (!textarea || !overlay) return;

        textarea.addEventListener("input", () => { scheduleInput(); reportCaret(); });
        textarea.addEventListener("scroll", syncScroll);
        textarea.addEventListener("keyup", reportCaret);
        textarea.addEventListener("click", reportCaret);

        textarea.addEventListener("keydown", (event) => {
            if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
                event.preventDefault();
                dotnetRef.invokeMethodAsync("OnSaveShortcut");
            }
        });

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

    function applyHighlights(forVersion, classes) {
        if (!textarea || !overlay || forVersion !== version) return;
        const lines = textarea.value.replace(/\r\n?/g, "\n").split("\n");
        const parts = [];
        for (let i = 0; i < lines.length; i++) {
            const cls = (classes && classes[i]) ? " " + classes[i] : "";
            parts.push('<span class="ln' + cls + '" data-line="' + (i + 1) + '">'
                + escapeHtml(lines[i]) + "\n</span>");
        }
        overlay.innerHTML = parts.join("");
        syncScroll();
    }

    function setContent(text) {
        if (!textarea) return version;
        version++;
        textarea.value = text;
        textarea.scrollTop = 0;
        textarea.setSelectionRange(0, 0);
        lastCaretLine = -1;
        dirty = false;
        return version;
    }

    function setDirty(value) {
        dirty = value;
    }

    function scrollToLine(line) {
        if (!textarea || !overlay) return;
        const span = overlay.querySelector('.ln[data-line="' + line + '"]');
        if (span) {
            const target = span.offsetTop - textarea.clientHeight / 3;
            textarea.scrollTop = Math.max(0, target);
            syncScroll();
        }

        const text = textarea.value;
        let offset = 0;
        let current = 1;
        while (current < line && offset < text.length) {
            const next = text.indexOf("\n", offset);
            if (next < 0) break;
            offset = next + 1;
            current++;
        }
        textarea.focus();
        textarea.setSelectionRange(offset, offset);
        reportCaret();
    }

    async function exportDocument(format, name) {
        if (!textarea) return;
        const baseName = (name && name.trim().length > 0 ? name : "Untitled")
            .replace(/\.(fountain|md|txt)$/i, "");

        if (format === "fountain") {
            const blob = new Blob([textarea.value], { type: "text/plain" });
            triggerDownload(blob, baseName + ".fountain");
            return;
        }

        const form = new FormData();
        form.append("content", textarea.value);
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
        if (textarea) textarea.focus();
    }

    return { init, applyHighlights, setContent, setDirty, scrollToLine, exportDocument, focusEditor };
})();
