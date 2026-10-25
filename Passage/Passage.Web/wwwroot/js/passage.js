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
    let session = { fileName: "", caretLine: 1, editorFontPx: 15, previewZoom: 1.25, recentFiles: [], lineOverrides: {}, workshopOpen: false, workshopWidth: 400, besideOpen: false, besideBeats: false, besideWidth: 300, boardDescriptions: true, arcMode: "board", typewriterScrolling: false, focusLine: false, wordsPerMinute: 230 };

    const INPUT_DEBOUNCE_MS = 200;
    const SESSION_KEY = "passage.session.v1";
    const SESSION_DEBOUNCE_MS = 400;
    const RECOVERY_KEY = "passage.recovery.v1";
    const RECOVERY_INTERVAL_MS = 3000;
    const THEME_KEY = "passage.theme.v1";
    const SYNTAX_SCHEME_KEY = "passage.syntaxScheme.v1";
    const AUTO_CONTD_KEY = "passage.autoContd.v1";
    const WORKSHOP_MIN_WIDTH = 260;
    const WORKSHOP_MAX_WIDTH = 720;
    const BESIDE_MIN_WIDTH = 200;
    const BESIDE_MAX_WIDTH = 560;
    const LINE_CLASSES = [
        "sx-scene", "sx-character", "sx-dialogue", "sx-paren", "sx-transition",
        "sx-section", "sx-synopsis", "sx-note", "sx-boneyard", "sx-centered",
        "sx-lyrics", "sx-titlepage", "md-heading"
    ];

    function pushContent() {
        if (!dotnetRef || !cm) return;
        const caretLine = cm.getCursor().line + 1;
        pruneDeletedOverrides();
        const overrides = overridesByLineNumber();

        // Refresh the stored copy as well as the server's: the handles have
        // moved with the text, and a reload restores from these numbers.
        if (JSON.stringify(session.lineOverrides) !== JSON.stringify(overrides)) {
            session.lineOverrides = overrides;
            scheduleSessionSave();
        }

        dotnetRef.invokeMethodAsync(
            "OnEditorChanged", cm.getValue("\n"), version, caretLine, overrides);
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
            applyWorkshopWidth();
            applyBesideWidth();
            applyViewToggles();
            return session;
        } catch (e) {
            return null;
        }
    }

    function setSessionDocument(fileName, editorFontPx, previewZoom, recentFiles, workshopOpen) {
        session.fileName = fileName || "";
        session.editorFontPx = editorFontPx;
        session.previewZoom = previewZoom;
        session.recentFiles = Array.isArray(recentFiles) ? recentFiles : [];
        session.workshopOpen = !!workshopOpen;
        scheduleSessionSave();
    }

    // The Slate dock's width is a per-browser preference like the font size,
    // but it never needs a server round-trip: the drag handle writes the CSS
    // variable directly and the session key remembers it. Whether the dock is
    // open is Blazor's, since it renders the dock; only the width lives here.
    function applyWorkshopWidth() {
        const width = Math.min(WORKSHOP_MAX_WIDTH, Math.max(WORKSHOP_MIN_WIDTH, Number(session.workshopWidth) || 400));
        session.workshopWidth = width;
        document.documentElement.style.setProperty("--workshop-width", width + "px");
    }

    function initWorkshopResize() {
        const handle = document.getElementById("workshop-resize");
        if (!handle) return;
        applyWorkshopWidth();
        handle.addEventListener("pointerdown", (down) => {
            down.preventDefault();
            const startX = down.clientX;
            const startWidth = session.workshopWidth;
            handle.setPointerCapture(down.pointerId);
            const move = (e) => {
                // The handle sits on the dock's left edge, so dragging left widens it.
                session.workshopWidth = startWidth + (startX - e.clientX);
                applyWorkshopWidth();
            };
            const up = () => {
                handle.removeEventListener("pointermove", move);
                handle.removeEventListener("pointerup", up);
                handle.removeEventListener("pointercancel", up);
                scheduleSessionSave();
            };
            handle.addEventListener("pointermove", move);
            handle.addEventListener("pointerup", up);
            handle.addEventListener("pointercancel", up);
        });
    }

    // SCRIPT's Beside pane (STORY-PLAN S2.3): width is dragged here and kept in
    // the session key, like the dock's; open/mode are Blazor's to render.
    function applyBesideWidth() {
        const width = Math.min(BESIDE_MAX_WIDTH, Math.max(BESIDE_MIN_WIDTH, Number(session.besideWidth) || 300));
        session.besideWidth = width;
        document.documentElement.style.setProperty("--beside-width", width + "px");
    }

    function initBesideResize() {
        const handle = document.getElementById("beside-resize");
        if (!handle) return;
        applyBesideWidth();
        handle.addEventListener("pointerdown", (down) => {
            down.preventDefault();
            const startX = down.clientX;
            const startWidth = session.besideWidth;
            handle.setPointerCapture(down.pointerId);
            const move = (e) => {
                session.besideWidth = startWidth + (startX - e.clientX);
                applyBesideWidth();
            };
            const up = () => {
                handle.removeEventListener("pointermove", move);
                handle.removeEventListener("pointerup", up);
                handle.removeEventListener("pointercancel", up);
                scheduleSessionSave();
            };
            handle.addEventListener("pointermove", move);
            handle.addEventListener("pointerup", up);
            handle.addEventListener("pointercancel", up);
        });
    }

    function setBeside(open, beats) {
        session.besideOpen = !!open;
        session.besideBeats = !!beats;
        scheduleSessionSave();
    }

    function setArcMode(mode) {
        session.arcMode = mode === "pages" ? "pages" : "board";
        scheduleSessionSave();
    }

    function setBoardDescriptions(show) {
        session.boardDescriptions = !!show;
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
        applySyntaxColours(syntaxColours);
        // CodeMirror caches measurements against the old colours.
        if (cm) cm.refresh();
        return next;
    }

    // Syntax colour scheme. The server owns the presets and resolves the state
    // to one palette per theme; this side only remembers the state blob and
    // paints whichever palette matches the theme that is showing. Inline
    // --syntax-* properties on <html> override both theme rules in app.css,
    // which is why setTheme re-applies them.
    let syntaxColours = null;

    function getSyntaxScheme() {
        try {
            return window.localStorage.getItem(SYNTAX_SCHEME_KEY);
        } catch (e) {
            return null;
        }
    }

    function setSyntaxScheme(json) {
        try {
            window.localStorage.setItem(SYNTAX_SCHEME_KEY, json);
        } catch (e) {
        }
    }

    function applySyntaxColours(colours) {
        syntaxColours = colours;
        if (!colours) return;
        const palette = colours[getTheme()];
        if (!palette) return;
        const style = document.documentElement.style;
        for (const token of Object.keys(palette)) {
            style.setProperty("--syntax-" + token, palette[token]);
        }
        if (cm) cm.refresh();
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

    // Selecting and tagging Board cards and beats (STORY-PLAN S1.4, decisions D
    // and G). Selection, arrow keys and the popover's filtering all stay in the
    // browser (trap 4); only Enter in the popover, and Enter / double-click on
    // the selection, reach the server.
    let tagTargets = [];
    let itemTags = {};
    let selectedItem = null; // { kind, id, line }
    let tagPopover = null;

    function setTagTargets(targets, items) {
        tagTargets = Array.isArray(targets) ? targets : [];
        itemTags = items || {};
        if (tagPopover) tagPopover.refresh();
    }

    function itemElements() {
        return Array.from(document.querySelectorAll("[data-item-kind]")).filter((el) => el.offsetParent !== null);
    }

    function selectedElement() {
        if (!selectedItem) return null;
        const all = itemElements().filter((el) => el.dataset.itemKind === selectedItem.kind);
        let found = selectedItem.id ? all.find((el) => el.dataset.itemId === selectedItem.id) : null;
        // A card not yet stamped is found by line until it has an id.
        if (!found && selectedItem.kind === "card" && selectedItem.line) {
            found = all.find((el) => el.dataset.line === String(selectedItem.line));
        }
        return found || null;
    }

    function applySelection() {
        const el = selectedElement();
        document.querySelectorAll(".item-selected").forEach((other) => {
            if (other !== el) other.classList.remove("item-selected");
        });
        if (el && !el.classList.contains("item-selected")) el.classList.add("item-selected");
        if (el && selectedItem.kind === "card") {
            selectedItem.line = Number(el.dataset.line) || selectedItem.line;
            if (el.dataset.itemId) selectedItem.id = el.dataset.itemId;
        }
    }

    function selectItem(el) {
        selectedItem = el
            ? { kind: el.dataset.itemKind, id: el.dataset.itemId || "", line: Number(el.dataset.line) || 0 }
            : null;
        applySelection();
        if (el) el.scrollIntoView({ block: "nearest", inline: "nearest" });
    }

    function typingInField() {
        const active = document.activeElement;
        if (!active || active === document.body) return false;
        return /^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName) || active.isContentEditable
            || !!active.closest(".cm-editor");
    }

    function activateSelection() {
        const el = selectedElement();
        if (!el || !dotnetRef) return;
        if (selectedItem.kind === "beat") dotnetRef.invokeMethodAsync("OnBeatActivated", el.dataset.itemId);
        else dotnetRef.invokeMethodAsync("OnJumpTarget", Number(el.dataset.line));
    }

    function openTagPopover() {
        if (tagPopover) return;
        const el = selectedElement();
        if (!el || !dotnetRef) return;
        const kind = el.dataset.itemKind;
        let targetId = el.dataset.itemId || ("line:" + el.dataset.line);
        let on = new Set(itemTags[targetId] || []);
        let matches = [];
        let selected = 0;

        const box = document.createElement("div");
        box.className = "tag-popover";
        box.setAttribute("role", "dialog");
        box.setAttribute("aria-label", "Tag");
        box.innerHTML = '<input type="text" placeholder="Tag with a storyline, character or location..." />' +
            '<div class="tag-list"></div>';
        const input = box.querySelector("input");
        const list = box.querySelector(".tag-list");

        function render() {
            const query = input.value.trim().toLowerCase();
            matches = query.length === 0
                ? tagTargets
                : tagTargets.filter((target) => target.label.toLowerCase().includes(query));
            selected = Math.max(0, Math.min(selected, matches.length - 1));
            list.replaceChildren();
            if (matches.length === 0) {
                const empty = document.createElement("div");
                empty.className = "sidebar-empty";
                empty.textContent = tagTargets.length === 0
                    ? "Nothing to tag with yet. Add a storyline in ELEMENTS."
                    : "Nothing matches that filter.";
                list.appendChild(empty);
                return;
            }
            matches.forEach((target, index) => {
                const row = document.createElement("div");
                row.className = "scene-row tag-row" + (index === selected ? " active" : "");
                const dot = document.createElement("span");
                dot.className = "tag-swatch" + (target.colour ? "" : " neutral");
                if (target.colour) dot.style.background = target.colour;
                const kindLabel = document.createElement("span");
                kindLabel.className = "jump-kind";
                kindLabel.textContent = target.kind;
                const tick = document.createElement("span");
                tick.className = "tag-tick";
                tick.textContent = on.has(target.key) ? "✓" : "";
                row.append(dot, document.createTextNode(target.label), kindLabel, tick);
                row.addEventListener("mousedown", (event) => {
                    event.preventDefault();
                    selected = index;
                    toggle();
                });
                list.appendChild(row);
            });
            const active = list.children[selected];
            if (active) active.scrollIntoView({ block: "nearest" });
        }

        function toggle() {
            const target = matches[selected];
            if (!target) return;
            const before = new Set(on);
            if (!on.delete(target.key)) on.add(target.key);
            render();
            dotnetRef.invokeMethodAsync("OnTagToggled", targetId, kind, target.key).then((id) => {
                if (!id) {
                    on = before;
                } else {
                    targetId = id;
                    if (selectedItem) selectedItem.id = id;
                    on = new Set(itemTags[targetId] || []);
                }
                if (tagPopover) render();
            });
        }

        function close() {
            document.removeEventListener("mousedown", outside, true);
            box.remove();
            tagPopover = null;
        }

        function outside(event) {
            if (!box.contains(event.target)) close();
        }

        input.addEventListener("input", () => {
            selected = 0;
            render();
        });
        input.addEventListener("keydown", (event) => {
            if (event.key === "ArrowDown") selected = Math.min(selected + 1, matches.length - 1);
            else if (event.key === "ArrowUp") selected = Math.max(selected - 1, 0);
            else if (event.key === "Enter") toggle();
            else if (event.key === "Escape") close();
            else return;
            event.preventDefault();
            event.stopPropagation();
            if (tagPopover) render();
        });

        document.body.appendChild(box);
        const rect = el.getBoundingClientRect();
        const left = Math.max(8, Math.min(rect.left, window.innerWidth - box.offsetWidth - 8));
        const below = rect.bottom + 6;
        const top = below + box.offsetHeight > window.innerHeight - 8
            ? Math.max(8, rect.top - box.offsetHeight - 6)
            : below;
        box.style.left = left + "px";
        box.style.top = top + "px";
        document.addEventListener("mousedown", outside, true);
        tagPopover = {
            refresh() {
                on = new Set(itemTags[targetId] || []);
                render();
            }
        };
        render();
        input.focus();
    }

    document.addEventListener("click", (event) => {
        if (!event.target.closest || event.target.closest("button, input, textarea, select, a, .tag-popover")) return;
        const el = event.target.closest("[data-item-kind]");
        if (el) selectItem(el);
    });

    document.addEventListener("dblclick", (event) => {
        if (!event.target.closest || event.target.closest("button, input, textarea, select, a, .tag-popover")) return;
        const el = event.target.closest("[data-item-kind]");
        if (!el) return;
        selectItem(el);
        activateSelection();
    });

    document.addEventListener("keydown", (event) => {
        if (!selectedItem || tagPopover || typingInField() || event.ctrlKey || event.metaKey || event.altKey) return;
        const el = selectedElement();
        if (!el) return;
        if (event.key === "#") {
            openTagPopover();
        } else if (event.key === "Enter" && document.activeElement?.tagName !== "BUTTON") {
            activateSelection();
        } else if (event.key === "Escape") {
            selectedItem = null;
            applySelection();
            return;
        } else if (/^Arrow(Left|Right|Up|Down)$/.test(event.key)) {
            const all = itemElements();
            const index = all.indexOf(el);
            const next = all[Math.max(0, Math.min(all.length - 1, index + (/Left|Up/.test(event.key) ? -1 : 1)))];
            if (next) selectItem(next);
        } else {
            return;
        }
        event.preventDefault();
    });

    // Blazor re-renders replace or rewrite the cards; put the ring back.
    let selectionFrame = 0;
    new MutationObserver(() => {
        if (!selectedItem || selectionFrame) return;
        selectionFrame = requestAnimationFrame(() => {
            selectionFrame = 0;
            applySelection();
        });
    }).observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["class"] });

    // Jump to (HIGHLAND-PLAN H4.2): scenes, sections and named markers from the
    // last parse, filtered here so typing never round-trips (trap 4). Only the
    // chosen jump goes to the server, which also switches back to the script.
    let jumpTargets = [];
    let jumpPalette = null;

    function setJumpTargets(targets) {
        jumpTargets = Array.isArray(targets) ? targets : [];
    }

    function openJumpPalette() {
        if (jumpPalette) return;
        const backdrop = document.createElement("div");
        backdrop.className = "modal-backdrop";
        backdrop.innerHTML =
            '<div class="modal modal-tall jump-palette" role="dialog" aria-label="Jump to">' +
            "<h3>Jump To</h3>" +
            '<input type="text" placeholder="Filter scenes, sections, markers..." />' +
            '<div class="scene-list"></div>' +
            "</div>";
        const input = backdrop.querySelector("input");
        const list = backdrop.querySelector(".scene-list");
        let matches = [];
        let selected = 0;

        function render() {
            const query = input.value.trim().toLowerCase();
            matches = query.length === 0
                ? jumpTargets
                : jumpTargets.filter((target) => target.label.toLowerCase().includes(query));
            selected = Math.max(0, Math.min(selected, matches.length - 1));
            list.replaceChildren();
            if (matches.length === 0) {
                const empty = document.createElement("div");
                empty.className = "sidebar-empty";
                empty.textContent = jumpTargets.length === 0
                    ? "No scenes, sections or named markers in this script."
                    : "Nothing matches that filter.";
                list.appendChild(empty);
                return;
            }
            matches.forEach((target, index) => {
                const row = document.createElement("div");
                row.className = "scene-row" + (index === selected ? " active" : "");
                const kind = document.createElement("span");
                kind.className = "jump-kind";
                kind.textContent = target.kind;
                row.append(kind, document.createTextNode(target.label));
                row.addEventListener("mousedown", (event) => {
                    event.preventDefault();
                    choose(index);
                });
                list.appendChild(row);
            });
            const active = list.children[selected];
            if (active) active.scrollIntoView({ block: "nearest" });
        }

        function close() {
            backdrop.remove();
            jumpPalette = null;
        }

        function choose(index) {
            const target = matches[index];
            close();
            if (!target) {
                if (cm) cm.focus();
                return;
            }
            if (dotnetRef) dotnetRef.invokeMethodAsync("OnJumpTarget", target.line);
            else scrollToLine(target.line);
        }

        input.addEventListener("input", () => {
            selected = 0;
            render();
        });
        input.addEventListener("keydown", (event) => {
            if (event.key === "ArrowDown") selected = Math.min(selected + 1, matches.length - 1);
            else if (event.key === "ArrowUp") selected = Math.max(selected - 1, 0);
            else if (event.key === "Enter") choose(selected);
            else if (event.key === "Escape") {
                close();
                if (cm) cm.focus();
            }
            else return;
            event.preventDefault();
            if (jumpPalette) render();
        });
        backdrop.addEventListener("mousedown", (event) => {
            if (event.target === backdrop) {
                close();
                if (cm) cm.focus();
            }
        });

        document.body.appendChild(backdrop);
        jumpPalette = backdrop;
        render();
        input.focus();
    }

    function scrollIntoView(selector) {
        const el = document.querySelector(selector);
        if (el) el.scrollIntoView({ block: "nearest" });
    }

    function scrollToTop(selector) {
        const el = document.querySelector(selector);
        if (el) el.scrollTop = 0;
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

    // Scene/section ids ([[id:GUID]]) live in the text but are never shown:
    // each is folded into an atomic, zero-width mark. The text itself is
    // untouched, so saving, parsing and the board still see it.
    const ID_MARKER = /\s*\[\[id:[a-f\d\-]+\]\]/gi;

    function hideIdMarkers(from, to) {
        cm.operation(() => {
            cm.eachLine(from, to, (handle) => {
                const lineNo = cm.getLineNumber(handle);
                if (lineNo === null) return;
                const text = handle.text;
                if (text.indexOf("[[id:") < 0) return;
                ID_MARKER.lastIndex = 0;
                let match;
                while ((match = ID_MARKER.exec(text)) !== null) {
                    const start = { line: lineNo, ch: match.index };
                    const end = { line: lineNo, ch: match.index + match[0].length };
                    const exists = cm.findMarksAt(start).some((m) => m.__idMarker) ||
                        cm.findMarks(start, end).some((m) => m.__idMarker);
                    if (exists) continue;
                    const mark = cm.markText(start, end, { collapsed: true, atomic: true, inclusiveLeft: false, inclusiveRight: false });
                    mark.__idMarker = true;
                }
            });
        });
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
                "Ctrl-J": openJumpPalette,
                "Cmd-J": openJumpPalette,
                "Ctrl-F": openFind,
                "Cmd-F": openFind,
                "Ctrl-H": openReplace,
                "Cmd-H": openReplace,
                "F1": toggleSyntaxPanel
            }
        });

        initWorkshopResize();
        initBesideResize();

        cm.on("change", (_, changeObj) => {
            if (changeObj.origin !== "setValue") {
                scheduleInput();
                if (changeObj.origin === "+input") {
                    maybeHint();
                }
            }
            reportCaret();
        });
        cm.on("cursorActivity", reportCaret);
        cm.on("cursorActivity", onViewCursorActivity);
        cm.getWrapperElement().addEventListener("mousedown", () => { lastInputWasPointer = true; }, true);
        cm.getWrapperElement().addEventListener("keydown", () => { lastInputWasPointer = false; }, true);
        cm.on("changes", updatePageRules);
        cm.on("refresh", updatePageRules);
        cm.on("changes", (_, changes) => {
            changes.forEach((c) => {
                const first = Math.max(0, c.from.line);
                const last = Math.min(cm.lastLine(), c.from.line + c.text.length - 1);
                hideIdMarkers(first, last + 1);
            });
        });

        // Close any open dropdown menu after a choice or an outside click.
        document.addEventListener("click", (event) => {
            document.querySelectorAll("details.menu[open]").forEach((menu) => {
                if (!menu.contains(event.target) || event.target.closest(".menu-items")) {
                    menu.removeAttribute("open");
                }
            });
        });

        document.addEventListener("dragover", trackDropSide, true);
        document.addEventListener("pointerdown", startBeatGesture);
        document.addEventListener("pointerdown", startArcDrag);

        cm.getWrapperElement().addEventListener("contextmenu", (event) => {
            event.preventDefault();
            openClassifyMenu(event);
        });
        document.addEventListener("click", closeClassifyMenu);
        cm.on("scroll", closeClassifyMenu);

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

    // Replace one span inside a line, matched by its text rather than trusting
    // a column: the server's copy of the document can be a debounce behind the
    // editor, and a placeholder that has moved or gone must not clobber other
    // text. Returns false when the span is not on that line any more. Ranged,
    // so undo keeps working (docs/WEB-PARITY.md 1.6).
    function replaceInLine(lineIndex, expected, replacement) {
        if (!cm) return false;
        if (lineIndex < 0 || lineIndex >= cm.lineCount()) return false;
        const line = cm.getLine(lineIndex);
        const ch = line.indexOf(expected);
        if (ch < 0) return false;

        cm.replaceRange(replacement, { line: lineIndex, ch }, { line: lineIndex, ch: ch + expected.length });
        cm.setCursor({ line: lineIndex, ch: ch + replacement.length });
        scheduleInput();
        return true;
    }

    // Rewrite whole lines as one undo step (one cm.operation is one history
    // event). Each edit names the text it expects; a line that no longer reads
    // that (the server copy lags the editor by a debounce) is skipped and
    // counted, never clobbered.
    function replaceLines(edits) {
        if (!cm) return { applied: 0, skipped: edits.length };
        let applied = 0;
        cm.operation(function () {
            for (const edit of edits) {
                if (edit.lineIndex >= cm.lineCount() || cm.getLine(edit.lineIndex) !== edit.expected) continue;
                cm.replaceRange(edit.replacement, { line: edit.lineIndex, ch: 0 }, { line: edit.lineIndex, ch: edit.expected.length });
                applied++;
            }
        });
        if (applied > 0) scheduleInput();
        return { applied: applied, skipped: edits.length - applied };
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

    // ---- Beat Sheet: drag-move and resize ----
    //
    // Both gestures run entirely here and send one result to the server when
    // the pointer is released: OnBeatMoved(id, page) with page 0 meaning the
    // tray, or OnBeatResized(id, span). Nothing round-trips while dragging.
    // The clamps mirror BeatSheet.ClampMove / BeatSheet.Clamp; the server's
    // are the truth.
    const BEAT_DRAG_THRESHOLD = 4;
    const BEAT_EDGE_SCROLL = 48;

    function beatColumnAt(grid, clientX) {
        const first = grid.querySelector(".beat-column");
        const width = first ? first.getBoundingClientRect().width : 0;
        if (width <= 0) return 1;
        return Math.floor((clientX - grid.getBoundingClientRect().left) / width) + 1;
    }

    function pointIn(element, event) {
        if (!element) return false;
        const rect = element.getBoundingClientRect();
        return event.clientX >= rect.left && event.clientX <= rect.right
            && event.clientY >= rect.top && event.clientY <= rect.bottom;
    }

    // A click follows the pointerup that ends a drag; swallow it so the
    // dropped beat does not also open its editor.
    function swallowNextClick() {
        const swallow = (event) => { event.stopPropagation(); event.preventDefault(); };
        window.addEventListener("click", swallow, { capture: true, once: true });
        setTimeout(() => window.removeEventListener("click", swallow, { capture: true }), 0);
    }

    // Scrolls the grid while the pointer is held near its left or right
    // edge, re-running the gesture's move so the target follows the scroll.
    // Returns a function that stops it.
    function startEdgeScroll(surface, onScroll) {
        let last = null;
        let frame = 0;
        const tick = () => {
            frame = requestAnimationFrame(tick);
            if (!surface || !last) return;
            const rect = surface.getBoundingClientRect();
            const before = surface.scrollLeft;
            if (last.clientX < rect.left + BEAT_EDGE_SCROLL) surface.scrollLeft -= 16;
            else if (last.clientX > rect.right - BEAT_EDGE_SCROLL) surface.scrollLeft += 16;
            if (surface.scrollLeft !== before) onScroll(last);
        };
        frame = requestAnimationFrame(tick);
        return {
            track(event) { last = event; },
            stop() { cancelAnimationFrame(frame); }
        };
    }

    // Capture can fail if the pointer is already gone; the gesture still
    // works from the card's own events, just less tightly.
    function capturePointer(card, down) {
        try { card.setPointerCapture(down.pointerId); } catch { }
    }

    function startBeatGesture(down) {
        if (down.button !== 0 || !dotnetRef || !down.target.closest) return;
        const card = down.target.closest(".beat-card");
        if (card && card.closest(".board-inbox")) {
            moveInboxBeat(down, card, card.dataset.beatId);
            return;
        }
        const grid = document.querySelector(".beats-grid");
        if (!card || !grid) return;

        const id = card.dataset.beatId;
        const columns = Number(grid.dataset.columns) || 1;
        const page = Number(card.dataset.page) || 0;
        const span = Math.max(1, Number(card.dataset.span) || 1);

        if (down.target.closest(".beat-resize") && page > 0) {
            down.preventDefault();
            resizeBeat(down, card, grid, id, columns, page, span);
            return;
        }

        moveBeat(down, card, grid, id, columns, page, span);
    }

    function resizeBeat(down, card, grid, id, columns, page, span) {
        const original = card.style.gridColumn;
        const surface = document.querySelector(".beats-surface");
        let next = span;
        capturePointer(card, down);
        card.classList.add("resizing");

        const update = (event) => {
            next = Math.min(Math.max(1, beatColumnAt(grid, event.clientX) - page + 1), columns - page + 1);
            card.style.gridColumn = page + " / span " + next;
        };
        const scroller = startEdgeScroll(surface, update);
        const move = (event) => { scroller.track(event); update(event); };
        const up = (event) => {
            scroller.stop();
            card.removeEventListener("pointermove", move);
            card.removeEventListener("pointerup", up);
            card.removeEventListener("pointercancel", up);
            card.classList.remove("resizing");
            swallowNextClick();
            if (event.type === "pointercancel" || next === span) {
                card.style.gridColumn = original;
                return;
            }
            dotnetRef.invokeMethodAsync("OnBeatResized", id, next);
        };
        card.addEventListener("pointermove", move);
        card.addEventListener("pointerup", up);
        card.addEventListener("pointercancel", up);
    }

    function moveBeat(down, card, grid, id, columns, page, span) {
        const tray = document.querySelector(".beats-tray");
        const surface = document.querySelector(".beats-surface");
        // Grabbing a long beat by its third page keeps that page under the pointer.
        const grabOffset = page > 0 ? beatColumnAt(grid, down.clientX) - page : 0;
        const fit = Math.min(span, columns);
        let ghost = null;
        let hint = null;
        let target = null;

        let scroller = null;

        const clear = () => {
            if (scroller) scroller.stop();
            if (ghost) ghost.remove();
            if (hint) hint.remove();
            card.classList.remove("dragging");
            if (tray) tray.classList.remove("drop-target");
        };

        const move = (event) => {
            if (!ghost) {
                if (Math.abs(event.clientX - down.clientX) < BEAT_DRAG_THRESHOLD
                    && Math.abs(event.clientY - down.clientY) < BEAT_DRAG_THRESHOLD) return;
                const rect = card.getBoundingClientRect();
                ghost = card.cloneNode(true);
                ghost.classList.add("beat-ghost");
                ghost.style.width = rect.width + "px";
                ghost.dataset.dx = String(down.clientX - rect.left);
                ghost.dataset.dy = String(down.clientY - rect.top);
                document.body.appendChild(ghost);
                hint = document.createElement("div");
                hint.className = "beat-drop-hint";
                card.classList.add("dragging");
                scroller = startEdgeScroll(surface, place);
            }

            ghost.style.left = (event.clientX - Number(ghost.dataset.dx)) + "px";
            ghost.style.top = (event.clientY - Number(ghost.dataset.dy)) + "px";
            scroller.track(event);
            place(event);
        };

        const place = (event) => {
            if (pointIn(tray, event)) {
                target = 0;
                tray.classList.add("drop-target");
                hint.remove();
            } else if (pointIn(surface, event)) {
                const start = Math.min(Math.max(1, beatColumnAt(grid, event.clientX) - grabOffset), columns - fit + 1);
                target = start;
                if (tray) tray.classList.remove("drop-target");
                hint.style.gridColumn = start + " / span " + fit;
                if (!hint.parentNode) grid.appendChild(hint);
            } else {
                target = null;
                if (tray) tray.classList.remove("drop-target");
                hint.remove();
            }
        };

        const up = (event) => {
            card.removeEventListener("pointermove", move);
            card.removeEventListener("pointerup", up);
            card.removeEventListener("pointercancel", up);
            const dragged = !!ghost;
            clear();
            if (!dragged) return;
            swallowNextClick();
            if (event.type === "pointercancel" || target === null) return;
            if (target === page || (target === 0 && page === 0)) return;
            dotnetRef.invokeMethodAsync("OnBeatMoved", id, target);
        };

        capturePointer(card, down);
        card.addEventListener("pointermove", move);
        card.addEventListener("pointerup", up);
        card.addEventListener("pointercancel", up);
    }

    // ---- Arc view: drag a point vertically (S3.1) ----
    //
    // Runs entirely here. The point and its storyline's polyline follow the
    // pointer; the server hears once, on release: OnArcPointMoved(storylineId,
    // targetId, value). The clamp mirrors ArcLayout.Clamp; the server's is the
    // truth. The svg carries data-top / data-height, the plot's vertical extent
    // in viewBox units.
    function startArcDrag(down) {
        if (down.button !== 0 || !dotnetRef || !down.target.closest) return;
        const point = down.target.closest(".arc-point");
        const svg = point && point.ownerSVGElement;
        if (!svg) return;
        down.preventDefault();
        const top = parseFloat(svg.dataset.top);
        const height = parseFloat(svg.dataset.height);
        const startY = point.getAttribute("cy");
        const startFill = point.classList.contains("untouched");
        let value = null;

        const toValue = (event) => {
            const ctm = svg.getScreenCTM();
            if (!ctm) return null;
            const local = new DOMPoint(event.clientX, event.clientY).matrixTransform(ctm.inverse());
            return Math.max(0, Math.min(100, Math.round(100 - ((local.y - top) / height) * 100)));
        };

        const redraw = () => {
            const circles = Array.from(svg.querySelectorAll(".arc-point")).filter((c) => c.dataset.line === point.dataset.line);
            const line = svg.querySelector('.arc-line[data-line="' + point.dataset.line + '"]');
            if (line) line.setAttribute("points", circles.map((c) => c.getAttribute("cx") + "," + c.getAttribute("cy")).join(" "));
        };

        const move = (event) => {
            const next = toValue(event);
            if (next === null) return;
            value = next;
            point.setAttribute("cy", String(top + ((100 - value) / 100) * height));
            point.classList.remove("untouched");
            point.classList.add("dragging");
            redraw();
        };

        const up = (event) => {
            point.removeEventListener("pointermove", move);
            point.removeEventListener("pointerup", up);
            point.removeEventListener("pointercancel", up);
            point.classList.remove("dragging");
            if (value === null) return;
            if (event.type === "pointercancel") {
                point.setAttribute("cy", startY);
                if (startFill) point.classList.add("untouched");
                redraw();
                return;
            }
            dotnetRef.invokeMethodAsync("OnArcPointMoved", point.dataset.line, point.dataset.target, value);
        };

        capturePointer(point, down);
        point.addEventListener("pointermove", move);
        point.addEventListener("pointerup", up);
        point.addEventListener("pointercancel", up);
    }

    // ---- Board inbox: drop a beat onto a card ----
    //
    // The same pointer gesture as moveBeat, but the target is a Board card or
    // an act / sequence header (anything carrying data-drop-id). Before or after
    // is the vertical midpoint, as dropIsAfter. The highlight is a class toggle;
    // the server hears once, on release: OnInboxBeatPlaced(beatId, targetId, after).
    function moveInboxBeat(down, card, id) {
        let ghost = null;
        let target = null;
        let after = false;

        const mark = (element, isAfter) => {
            if (target) target.classList.remove("inbox-drop-before", "inbox-drop-after");
            target = element;
            if (target) target.classList.add(isAfter ? "inbox-drop-after" : "inbox-drop-before");
        };

        const move = (event) => {
            if (!ghost) {
                if (Math.abs(event.clientX - down.clientX) < BEAT_DRAG_THRESHOLD
                    && Math.abs(event.clientY - down.clientY) < BEAT_DRAG_THRESHOLD) return;
                const rect = card.getBoundingClientRect();
                ghost = card.cloneNode(true);
                ghost.classList.add("beat-ghost");
                ghost.style.width = rect.width + "px";
                ghost.dataset.dx = String(down.clientX - rect.left);
                ghost.dataset.dy = String(down.clientY - rect.top);
                document.body.appendChild(ghost);
                card.classList.add("dragging");
            }

            ghost.style.left = (event.clientX - Number(ghost.dataset.dx)) + "px";
            ghost.style.top = (event.clientY - Number(ghost.dataset.dy)) + "px";

            const under = document.elementFromPoint(event.clientX, event.clientY);
            const drop = under && under.closest ? under.closest("[data-drop-id]") : null;
            if (!drop || !drop.dataset.dropId) {
                mark(null, false);
                return;
            }
            const rect = drop.getBoundingClientRect();
            after = rect.height > 0 && (event.clientY - rect.top) > rect.height / 2;
            mark(drop, after);
        };

        const up = (event) => {
            document.removeEventListener("pointermove", move);
            document.removeEventListener("pointerup", up);
            document.removeEventListener("pointercancel", up);
            const dragged = !!ghost;
            const dropId = target ? target.dataset.dropId : "";
            mark(null, false);
            if (ghost) ghost.remove();
            card.classList.remove("dragging");
            if (!dragged) return;
            swallowNextClick();
            if (event.type === "pointercancel" || !dropId) return;
            dotnetRef.invokeMethodAsync("OnInboxBeatPlaced", id, dropId, after);
        };

        // Listening on the document, with no pointer capture, so a click on the
        // beat's star still reaches the star.
        document.addEventListener("pointermove", move);
        document.addEventListener("pointerup", up);
        document.addEventListener("pointercancel", up);
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

    // ---- Autocomplete ----
    //
    // Ports UpdateUniqueScreenplayElements / UpdateSuggestions. The lists are
    // pushed from the server after each parse; matching and display happen
    // entirely here, so typing never round-trips (trap 4).
    //
    // The current line's element type comes from appliedClasses, the per-line
    // classes the parse already sends for colouring — so no extra call is
    // needed to know whether the caret is on a scene heading or a character.
    let sceneHeadingSuggestions = [];
    let characterSuggestions = [];
    let suggestionKinds = [];

    const MAX_SUGGESTIONS = 10;

    // Re-attach stored overrides to the handles of the document just loaded.
    function restoreLineOverrides(map) {
        if (!cm || !map) return;
        lineOverrides = new Map();
        for (const key of Object.keys(map)) {
            const lineIndex = parseInt(key, 10) - 1;
            if (!(lineIndex >= 0) || lineIndex > cm.lastLine()) continue;
            const handle = cm.getLineHandle(lineIndex);
            if (handle) lineOverrides.set(handle, map[key]);
        }
        session.lineOverrides = overridesByLineNumber();
        pushContent();
    }

    function setSuggestions(sceneHeadings, characters, kinds) {
        sceneHeadingSuggestions = Array.isArray(sceneHeadings) ? sceneHeadings : [];
        characterSuggestions = Array.isArray(characters) ? characters : [];
        suggestionKinds = Array.isArray(kinds) ? kinds : [];
        offerContdAfterParse();
    }

    // Character highlighting (HIGHLAND-PLAN H3.1). speakers is per line, from
    // the last parse; colours maps a speaker to a palette hex. Each run of
    // lines under one coloured speaker becomes one mark, so typing inside a
    // speech keeps its tint until the next parse re-lays them.
    let characterMarks = [];

    function setCharacterHighlights(speakers, colours) {
        if (!cm) return;
        const lineSpeakers = Array.isArray(speakers) ? speakers : [];
        const lineColours = colours || {};
        cm.operation(() => {
            characterMarks.forEach((mark) => mark.clear());
            characterMarks = [];
            const last = Math.min(lineSpeakers.length, cm.lineCount()) - 1;
            for (let start = 0; start <= last; start++) {
                const colour = lineColours[lineSpeakers[start]];
                if (!colour) continue;
                let end = start;
                while (end < last && lineSpeakers[end + 1] === lineSpeakers[start]) end++;
                characterMarks.push(cm.markText(
                    { line: start, ch: 0 },
                    { line: end, ch: cm.getLine(end).length },
                    { className: "character-highlight", css: "--character-colour: " + colour }));
                start = end;
            }
        });
    }

    function looksLikeSceneHeading(prefix) {
        const upper = prefix.toUpperCase();
        return upper.startsWith(".") || upper.startsWith("INT.") ||
            upper.startsWith("EXT.") || upper.startsWith("I/E.");
    }

    // The server decides which list a line wants, using the shared TextAnalysis
    // helpers; this only picks the list. The INT./EXT. override is applied here
    // too because it is instant and the desktop applies it before consulting the
    // parse as well.
    function suggestionSource(lineIndex, prefix) {
        if (looksLikeSceneHeading(prefix)) return sceneHeadingSuggestions;

        const kind = suggestionKinds[lineIndex] || "";
        if (kind === "character") return characterSuggestions;
        if (kind === "scene") return sceneHeadingSuggestions;
        return null;
    }

    // Mirrors UpdateSuggestions: prefix match, drop an exact hit, sort, cap at 10.
    function screenplayHints(editor) {
        const cursor = editor.getCursor();
        const prefix = editor.getLine(cursor.line).slice(0, cursor.ch).trim();
        if (prefix.length === 0) return null;

        const source = suggestionSource(cursor.line, prefix);
        if (!source || source.length === 0) return null;

        const normalized = prefix.toUpperCase();
        const matches = source
            .filter(entry => entry.startsWith(normalized) && entry !== normalized)
            .sort()
            .slice(0, MAX_SUGGESTIONS);

        if (source === characterSuggestions) {
            const contd = contdOffer(editor, cursor.line, normalized);
            if (contd) {
                matches.unshift(contd);
                if (matches.length > MAX_SUGGESTIONS) matches.pop();
            }
        }

        if (matches.length === 0) return null;

        // Replace the whole line, as ApplyAutoCompleteSuggestion does.
        return {
            list: matches,
            from: { line: cursor.line, ch: 0 },
            to: { line: cursor.line, ch: editor.getLine(cursor.line).length }
        };
    }

    // ---- Typewriter scrolling and focus line (teardown section 15) ----
    //
    // Client-side only (trap 4): both run on cursorActivity, which CodeMirror
    // fires once per operation, so a keystroke costs one scroll and one line
    // class move however long the script is. The toggles live in the session
    // key; Blazor reads them back from the loadSession snapshot.
    let lastInputWasPointer = false;
    let focusLineHandle = null;

    function setTypewriterScrolling(enabled) {
        session.typewriterScrolling = !!enabled;
        applyViewToggles();
        scheduleSessionSave();
    }

    function setFocusLine(enabled) {
        session.focusLine = !!enabled;
        applyViewToggles();
        scheduleSessionSave();
    }

    // OUTLINE panel filters (H4.3); Blazor owns the shape and the defaults.
    function setOutlineFilters(filters) {
        session.outlineFilters = filters;
        scheduleSessionSave();
    }

    // BOARD and BEATS filter bars (STORY-PLAN S1.5), one set per view.
    function setStoryFilter(view, filter) {
        session[view === 'beats' ? 'beatFilter' : 'boardFilter'] = filter;
        scheduleSessionSave();
    }

    // Markdown reading speed for the status bar's reading time; Blazor clamps it.
    function setWordsPerMinute(wordsPerMinute) {
        session.wordsPerMinute = wordsPerMinute;
        scheduleSessionSave();
    }

    function applyViewToggles() {
        if (!cm) return;
        const wrapper = cm.getWrapperElement();
        const typewriter = !!session.typewriterScrolling;
        if (wrapper.classList.contains("typewriter") !== typewriter) {
            // The bottom padding changes the document height, which
            // CodeMirror only measures on refresh.
            wrapper.classList.toggle("typewriter", typewriter);
            cm.refresh();
        }
        lastInputWasPointer = false;
        onViewCursorActivity();
    }

    function onViewCursorActivity() {
        updateFocusLine();
        // A click places the caret where the writer is looking; re-centring
        // would yank that text out from under the pointer. Typing or
        // keyboard movement after the click centres again.
        if (session.typewriterScrolling && !lastInputWasPointer) centreCaret();
    }

    function centreCaret() {
        const caret = cm.cursorCoords(null, "window");
        const scroller = cm.getScrollerElement().getBoundingClientRect();
        const delta = (caret.top + caret.bottom) / 2 - (scroller.top + scroller.height / 2);
        if (Math.abs(delta) < 1) return;
        cm.scrollTo(null, cm.getScrollInfo().top + delta);
    }

    // Hidden while text is selected: the line background would paint over the
    // selection on that line, and a selection already marks where you are.
    function updateFocusLine() {
        const handle = session.focusLine && !cm.somethingSelected()
            ? cm.getLineHandle(cm.getCursor().line)
            : null;
        if (handle === focusLineHandle) return;
        if (focusLineHandle) cm.removeLineClass(focusLineHandle, "background", "focus-line");
        if (handle) cm.addLineClass(handle, "background", "focus-line");
        focusLineHandle = handle;
    }

    // ---- Auto-(CONT'D) (teardown section 13) ----
    //
    // Client-side only (trap 4). Reads the element types the last parse already
    // painted onto the lines above, so no call is made per keystroke.
    const CONTD = "(CONT'D)";
    let autoContdEnabled = readAutoContd();

    function readAutoContd() {
        try {
            return window.localStorage.getItem(AUTO_CONTD_KEY) !== "off";
        } catch (e) {
            return true;
        }
    }

    function getAutoContd() {
        return autoContdEnabled;
    }

    function setAutoContd(enabled) {
        autoContdEnabled = !!enabled;
        try {
            window.localStorage.setItem(AUTO_CONTD_KEY, autoContdEnabled ? "on" : "off");
        } catch (e) {
        }
    }

    // The bare speaker: forced-character @, dual-dialogue ^ and the
    // (V.O.)/(O.S.)/(CONT'D) extensions removed, as Highland compares them.
    function bareCueName(text) {
        return text.toUpperCase()
            .replace(/\((V\.O\.|O\.S\.|CONT['\u2019]D)\)/g, "")
            .replace(/^\s*@/, "")
            .replace(/\^\s*$/, "")
            .trim();
    }

    function sxClass(lineIndex) {
        const textClass = cm.getLineHandle(lineIndex).textClass || "";
        return textClass.split(/\s+/).find(name => name.startsWith("sx-")) || "";
    }

    // The previous speaker in this scene, if only action lies between their
    // speech and the line being typed. Anything else between — a scene
    // heading, a transition, another speech with no action — returns null.
    function previousSpeakerAfterAction(editor, lineIndex) {
        let sawAction = false;
        for (let i = lineIndex - 1; i >= 0; i--) {
            const text = editor.getLine(i);
            if (text.trim().length === 0) continue;
            const kind = sxClass(i);
            if (kind === "") {
                sawAction = true;
            } else if (kind === "sx-character") {
                return sawAction ? bareCueName(text) : null;
            } else if (kind === "sx-dialogue" || kind === "sx-paren") {
                if (!sawAction) return null;
            } else if (kind !== "sx-note" && kind !== "sx-synopsis" && kind !== "sx-boneyard") {
                return null;
            }
        }
        return null;
    }

    // A completed cue repeating the previous speaker is offered " (CONT'D)";
    // a trailing "(CONT" on any cue completes to "(CONT'D)".
    function contdOffer(editor, lineIndex, typed) {
        if (!autoContdEnabled) return null;

        const partial = typed.match(/\(CONT(['\u2019]D?)?$/);
        if (partial) return typed.slice(0, partial.index) + CONTD;

        // A dual-dialogue cue (trailing ^) is a simultaneous speech, not a continuation.
        if (/\(CONT['\u2019]D\)/.test(typed) || /\^\s*$/.test(typed)) return null;
        const name = bareCueName(typed);
        if (name.length === 0 || name !== previousSpeakerAfterAction(editor, lineIndex)) return null;
        return typed + " " + CONTD;
    }

    // A new line only becomes a cue once the parse classifies it, which is
    // usually after the name has been typed in one burst. So when the parse
    // lands, offer (CONT'D) on the caret line without waiting for a keystroke.
    function offerContdAfterParse() {
        if (!cm || !cm.hasFocus() || cm.somethingSelected() || cm.state.completionActive) return;
        const cursor = cm.getCursor();
        const line = cm.getLine(cursor.line);
        if (cursor.ch !== line.length || suggestionKinds[cursor.line] !== "character") return;
        if (!contdOffer(cm, cursor.line, line.trim().toUpperCase())) return;
        maybeHint();
    }

    function maybeHint() {
        if (!cm || cm.state.completionActive) return;
        const cursor = cm.getCursor();
        if (cm.somethingSelected()) return;
        if (!screenplayHints(cm)) return;

        cm.showHint({
            hint: screenplayHints,
            completeSingle: false
        });
    }

    // ---- "Classify As" line-type overrides ----
    //
    // Held against CodeMirror line handles rather than line numbers, so an
    // override follows its line when text is inserted or removed above it. The
    // desktop keys a plain dictionary by line number and never renumbers it, so
    // there the override silently lands on the wrong line after such an edit.
    // Converted back to numbers only when handing them to the parser.
    const CLASSIFY_TYPES = [
        ["Action", "Classify As Action"],
        ["SceneHeading", "Classify As Scene Heading"],
        ["Character", "Classify As Character"],
        ["Dialogue", "Classify As Dialogue"],
        ["Parenthetical", "Classify As Parenthetical"],
        ["Transition", "Classify As Transition"],
        ["Note", "Classify As Note"]
    ];

    let lineOverrides = new Map();
    let classifyMenu = null;

    function overridesByLineNumber() {
        const map = {};
        for (const [handle, type] of lineOverrides) {
            const line = cm.getLineNumber(handle);
            if (line === null) continue;      // the line was deleted
            map[String(line + 1)] = type;     // the parser counts from 1
        }
        return map;
    }

    function pruneDeletedOverrides() {
        for (const handle of [...lineOverrides.keys()]) {
            if (cm.getLineNumber(handle) === null) {
                lineOverrides.delete(handle);
            }
        }
    }

    function setLineOverride(lineIndex, type) {
        if (!cm) return;
        const handle = cm.getLineHandle(lineIndex);
        if (!handle) return;

        if (type) {
            lineOverrides.set(handle, type);
        } else {
            lineOverrides.delete(handle);
        }

        session.lineOverrides = overridesByLineNumber();
        scheduleSessionSave();
        pushContent();
    }

    function closeClassifyMenu() {
        if (classifyMenu) {
            classifyMenu.remove();
            classifyMenu = null;
        }
    }

    function openClassifyMenu(event) {
        if (!cm) return;
        closeClassifyMenu();

        // The line comes from the caret, not the click point — SetLineTypeOverride
        // works off CaretOffset too. It also avoids hit-testing the mouse against
        // wrapped lines, which is fiddly and easy to get subtly wrong.
        const lineIndex = cm.getCursor().line;

        classifyMenu = document.createElement("div");
        classifyMenu.className = "classify-menu";

        // Say which line will change, since it is the caret's, not the click's.
        const heading = document.createElement("div");
        heading.className = "classify-menu-head";
        heading.textContent = "Line " + (lineIndex + 1);
        classifyMenu.appendChild(heading);
        classifyMenu.style.left = event.clientX + "px";
        classifyMenu.style.top = event.clientY + "px";

        const current = lineOverrides.get(cm.getLineHandle(lineIndex));

        for (const [type, label] of CLASSIFY_TYPES) {
            const item = document.createElement("button");
            item.textContent = label;
            if (current === type) item.classList.add("active");
            item.addEventListener("click", () => {
                closeClassifyMenu();
                setLineOverride(lineIndex, type);
            });
            classifyMenu.appendChild(item);
        }

        if (current) {
            // The desktop offers no way back once a line is classified; this does.
            const clear = document.createElement("button");
            clear.textContent = "Clear Classification";
            clear.className = "clear";
            clear.addEventListener("click", () => {
                closeClassifyMenu();
                setLineOverride(lineIndex, null);
            });
            classifyMenu.appendChild(clear);
        }

        document.body.appendChild(classifyMenu);
    }

    function setContent(text) {
        if (!cm) return version;
        version++;
        appliedClasses = [];
        // A different document: its line numbers mean nothing here, and the
        // stored copy has to go too or it would be re-saved and then restored
        // onto whatever happens to sit at those lines next.
        lineOverrides = new Map();
        session.lineOverrides = {};
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

    function scrollToLine(line, focus = true) {
        if (!cm) return;
        const target = Math.max(0, Math.min(line - 1, cm.lineCount() - 1));
        cm.setCursor({ line: target, ch: 0 });
        const coords = cm.charCoords({ line: target, ch: 0 }, "local");
        const scroller = cm.getScrollInfo();
        cm.scrollTo(null, Math.max(0, coords.top - scroller.clientHeight / 3));
        if (focus) cm.focus();
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

    function downloadText(text, fileName) {
        triggerDownload(new Blob([text], { type: "text/plain" }), fileName);
    }

    function downloadBase64(base64, fileName, mimeType) {
        const bytes = Uint8Array.from(atob(base64), ch => ch.charCodeAt(0));
        triggerDownload(new Blob([bytes], { type: mimeType }), fileName);
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
        if (!cm) return;
        cm.focus();
        // Lines written while the editor was hidden (a template or inbox
        // insert from BOARD) are unmeasured until it is visible again.
        requestAnimationFrame(() => cm.refresh());
    }

    // ---- Burst timer for the Writer's Tools runners (SLATE-PLAN decision E) ----
    //
    // Ready → read-in → write → advance, entirely client-side so a circuit
    // blip cannot stall it. The slots are the runner's textarea[data-slot]
    // elements in DOM order. Enter commits an answer and moves on; so does the
    // clock running out, and that is all expiry does — the answer stays
    // editable, nothing is marked, nothing is counted. Blazor hears each
    // answer through the ordinary change event the focus move fires, and is
    // told once when the run ends.
    let burst = null;

    function startBurst(rootId, readInSeconds, writeSeconds) {
        endBurst(false);
        const root = document.getElementById(rootId);
        if (!root) return false;

        burst = {
            root,
            display: root.querySelector("[data-burst-display]") || document.querySelector("[data-burst-display]"),
            readIn: readInSeconds,
            write: writeSeconds,
            slot: null,
            phase: "",
            timer: null,
            waitTimer: null,
            observer: null,
            advancing: false
        };
        root.addEventListener("keydown", onBurstKeyDown);
        root.addEventListener("input", onBurstInput);
        root.addEventListener("focusout", onBurstFocusOut);

        const slots = burstSlots();
        const first = slots.find(slot => slot.value.trim() === "") || slots[0];
        if (!first) {
            endBurst(false);
            return false;
        }
        beginBurstSlot(first);
        return true;
    }

    function stopBurst() {
        endBurst(true);
    }

    // notify: tell Blazor the run is over. Not when a new run replaces it,
    // or the "ended" would land after the "started" and flip the button back.
    function endBurst(notify) {
        if (!burst) return;
        clearInterval(burst.timer);
        clearTimeout(burst.waitTimer);
        if (burst.observer) burst.observer.disconnect();
        burst.root.removeEventListener("keydown", onBurstKeyDown);
        burst.root.removeEventListener("input", onBurstInput);
        burst.root.removeEventListener("focusout", onBurstFocusOut);
        showBurst("", 0);
        burst = null;
        if (notify && dotnetRef) dotnetRef.invokeMethodAsync("OnBurstEnded");
    }

    function burstSlots() {
        return Array.from(burst.root.querySelectorAll("textarea[data-slot]"));
    }

    function burstSlotAfter(slot) {
        const slots = burstSlots();
        return slots[slots.indexOf(slot) + 1] || null;
    }

    // Ready: the caret is in the slot before anything counts. Then the
    // read-in, which typing cuts short.
    function beginBurstSlot(slot) {
        burst.slot = slot;
        burst.advancing = true;
        slot.focus();
        burst.advancing = false;
        runBurstCountdown("read", burst.readIn, () => runBurstCountdown("write", burst.write, advanceBurst));
    }

    function runBurstCountdown(phase, seconds, done) {
        clearInterval(burst.timer);
        burst.phase = phase;
        let left = seconds;
        showBurst(phase, left);
        burst.timer = setInterval(() => {
            left--;
            if (left <= 0) {
                clearInterval(burst.timer);
                burst.timer = null;
                done();
                return;
            }
            showBurst(phase, left);
        }, 1000);
    }

    function showBurst(phase, left) {
        if (!burst || !burst.display) return;
        burst.display.textContent = phase === "read" ? "read\u2026 " + left : phase === "write" ? String(left) : "";
    }

    function onBurstInput(event) {
        if (burst && burst.phase === "read" && event.target === burst.slot) {
            runBurstCountdown("write", burst.write, advanceBurst);
        }
    }

    // Enter commits a one-line slot. A slot marked multiline (a Lens rewrite,
    // minutes long) keeps Enter as a line break; only the clock moves it on.
    function onBurstKeyDown(event) {
        if (!burst || event.target !== burst.slot || event.key !== "Enter" || event.shiftKey) return;
        if ("slotMultiline" in event.target.dataset) return;
        event.preventDefault();
        advanceBurst();
    }

    // Focus leaving the slot other than by our own move: a tap on another
    // slot re-targets the run there; anything else means the writer has
    // stepped out, and the run ends quietly.
    function onBurstFocusOut(event) {
        if (!burst || burst.advancing) return;
        const next = event.relatedTarget;
        if (next && burst.root.contains(next) && next.matches("textarea[data-slot]")) {
            beginBurstSlot(next);
        } else {
            stopBurst();
        }
    }

    function advanceBurst() {
        if (!burst) return;
        clearInterval(burst.timer);
        burst.timer = null;
        burst.phase = "";
        const current = burst.slot;
        const next = burstSlotAfter(current);
        if (next) {
            beginBurstSlot(next);
            return;
        }

        // No slot after this one yet. Committing the answer may make Blazor
        // add a round (WOAC), so wait briefly for it to render; otherwise the
        // chain is complete and the run is over.
        burst.advancing = true;
        current.blur();
        burst.advancing = false;
        showBurst("", 0);
        burst.observer = new MutationObserver(() => {
            const added = burstSlotAfter(current);
            if (!added) return;
            burst.observer.disconnect();
            burst.observer = null;
            clearTimeout(burst.waitTimer);
            beginBurstSlot(added);
        });
        burst.observer.observe(burst.root, { childList: true, subtree: true });
        burst.waitTimer = setTimeout(stopBurst, 2000);
    }

    // ---- The tools dock's runner (docs/SLATE-REDO-PLAN.md §4.2) ----
    //
    // Per-browser preferences that never need a server round-trip live in the
    // session key: the launcher's chip, and whether "Why?" is left open.
    // Enter moves a runner on without a per-keystroke event (CLAUDE.md trap 4):
    // it blurs the answer, which fires the change event that saves it, and
    // taps Next only where nothing else will move the run on.
    // Read from storage when the session hasn't been loaded yet: the dock can
    // render before Blazor asks for the session.
    function getToolsPref(key) {
        let tools = session.tools;
        if (!sessionReady) {
            try {
                tools = JSON.parse(window.localStorage.getItem(SESSION_KEY) || "{}").tools;
            } catch (e) {
                tools = null;
            }
        }
        return tools && typeof tools[key] === "string" ? tools[key] : null;
    }

    function setToolsPref(key, value) {
        session.tools = Object.assign(session.tools || {}, { [key]: value });
        scheduleSessionSave();
    }

    function restoreWhy() {
        const open = !!(session.tools && session.tools.whyOpen === "1");
        document.querySelectorAll("details[data-why]").forEach(el => {
            if (el.open !== open && !el.dataset.whyTouched) el.open = open;
        });
    }

    document.addEventListener("toggle", event => {
        const el = event.target;
        if (!(el instanceof HTMLElement) || !el.matches("details[data-why]")) return;
        el.dataset.whyTouched = "1";
        if (session.tools && (session.tools.whyOpen === "1") === el.open) return;
        setToolsPref("whyOpen", el.open ? "1" : "0");
    }, true);

    function focusStep(rootId) {
        const root = document.getElementById(rootId);
        const slots = root ? Array.from(root.querySelectorAll("textarea[data-slot]")) : [];
        const target = slots.find(slot => slot.value.trim() === "") || slots[0];
        if (target) target.focus({ preventScroll: false });
    }

    document.addEventListener("keydown", event => {
        if (event.key !== "Enter" || event.isComposing) return;
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        if (target.matches("input[data-anchor-input]")) {
            event.preventDefault();
            focusStep("tools-step");
            return;
        }
        if (burst || !target.matches("#tools-step textarea[data-slot]")) return;
        const multiline = "slotMultiline" in target.dataset;
        if (multiline ? !(event.ctrlKey || event.metaKey) : event.shiftKey) return;
        const root = document.getElementById("tools-step");
        const slots = Array.from(root.querySelectorAll("textarea[data-slot]"));
        const at = slots.indexOf(target);
        event.preventDefault();
        if (at >= 0 && at < slots.length - 1) {
            slots[at + 1].focus();
            return;
        }
        const next = root.querySelector("[data-next]");
        target.blur();
        // An answered step holds the card (so Send to stays reachable) until
        // Next, so Enter always taps it: the change event saves, then Next moves on.
        if (next) next.click();
    });

    // The CodeMirror instance is exposed for end-to-end tests.
    return {
        init, applyHighlights, setContent, setDirty, scrollToLine, setBeside, setBoardDescriptions, setArcMode,
        exportDocument, downloadText, downloadBase64, focusEditor, startBurst, stopBurst, getToolsPref, setToolsPref, restoreWhy, focusStep,
        loadSession, setSessionDocument,
        readRecoverySnapshot, clearRecoverySnapshot,
        refreshHighlights, undo, redo, copyText, scrollIntoView, scrollToTop, replaceLineRange, replaceInLine, replaceLines, insertLinesAt, deleteLineRange, dropIsAfter, setPageRules, setSuggestions, setJumpTargets, setTagTargets, openJumpPalette, setCharacterHighlights, restoreLineOverrides,
        findNext, findPrevious, replaceCurrent, replaceAll, selectedText,
        getTheme, setTheme,
        getSyntaxScheme, setSyntaxScheme, applySyntaxColours,
        getAutoContd, setAutoContd,
        setTypewriterScrolling, setFocusLine, setWordsPerMinute, setOutlineFilters, setStoryFilter,
        get editor() { return cm; }
    };
})();
