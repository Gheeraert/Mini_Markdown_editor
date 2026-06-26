/**
 * Mini WYSIWYG Web -- logique principale
 *
 * Logique cle (symetrique a main.py) :
 *   - Quand le WYSIWYG gauche change, le Markdown droit se regenere (debounce 300 ms).
 *   - Quand le focus entre dans la zone Markdown : regeneration suspendue.
 *   - Quand le focus quitte la zone Markdown : regeneration reprise si changement en attente.
 */

"use strict";

// ---------------------------------------------------------------------------
// Etat global
// ---------------------------------------------------------------------------

let tinyEditor = null;          // instance TinyMCE
let suspended = false;          // generation suspendue (focus dans le panneau MD)
let pendingRender = false;      // changement en attente pendant la suspension
let renderTimer = null;         // debounce timer pour la generation
let autosaveTimer = null;       // debounce timer pour l'autosave

const RENDER_DELAY_MS   = 300;
const AUTOSAVE_DELAY_MS = 5000;

let metadata = {
    title: "", subtitle: "", author: "", date: "", lang: "", abstract: ""
};

// ---------------------------------------------------------------------------
// Generation Markdown (coeur du mecanisme)
// ---------------------------------------------------------------------------

function onWysiwygChange() {
    if (suspended) {
        pendingRender = true;
        return;
    }
    clearTimeout(renderTimer);
    renderTimer = setTimeout(generateMarkdown, RENDER_DELAY_MS);
    scheduleAutosave();
}

function generateMarkdown(force) {
    if (suspended && !force) {
        pendingRender = true;
        return;
    }
    if (!tinyEditor) return;

    const html = tinyEditor.getContent();

    fetch("/api/to-markdown", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ html, metadata }),
    })
    .then(function(r) { return r.json(); })
    .then(function(data) {
        document.getElementById("markdown-view").value = data.markdown || "";
    })
    .catch(function(err) {
        setStatus("Erreur de conversion : " + err.message);
    });
}

// ---------------------------------------------------------------------------
// Focus : panneau Markdown droit
// ---------------------------------------------------------------------------

function onMarkdownFocusIn() {
    suspended = true;
    document.getElementById("pane-right-status").textContent = "edition tampon — generation suspendue";
    document.getElementById("pane-right-status").style.display = "inline";
}

function onMarkdownFocusOut() {
    suspended = false;
    document.getElementById("pane-right-status").style.display = "none";
    if (pendingRender) {
        pendingRender = false;
        generateMarkdown(true);
    }
}

// ---------------------------------------------------------------------------
// Autosave
// ---------------------------------------------------------------------------

function scheduleAutosave() {
    clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(doAutosave, AUTOSAVE_DELAY_MS);
}

function doAutosave() {
    if (!tinyEditor) return;
    const html     = tinyEditor.getContent();
    const markdown = document.getElementById("markdown-view").value;

    fetch("/api/autosave", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ html, markdown }),
    })
    .then(function(r) { return r.json(); })
    .then(function(data) {
        if (data.status === "ok") {
            setAutosaveStatus("Autosauve " + new Date().toLocaleTimeString());
        }
    })
    .catch(function() { /* silencieux */ });
}

// ---------------------------------------------------------------------------
// Telechargement client (Blob)
// ---------------------------------------------------------------------------

function downloadBlob(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href     = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// Actions de la barre d'outils
// ---------------------------------------------------------------------------

function saveHtml() {
    if (!tinyEditor) return;
    const html = tinyEditor.getContent();
    fetch("/api/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ format: "html", content: html, filename: "document" }),
    })
    .then(function(r) { return r.json(); })
    .then(function(data) {
        setStatus(data.status === "ok"
            ? "Document HTML sauvegarde : documents/" + data.saved_as
            : "Erreur : " + data.message);
    });
}

function exportMarkdown() {
    // Regenere depuis le WYSIWYG pour avoir un MD a jour
    if (!tinyEditor) return;
    const html = tinyEditor.getContent();
    fetch("/api/to-markdown", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ html, metadata }),
    })
    .then(function(r) { return r.json(); })
    .then(function(data) {
        const md = data.markdown || "";
        document.getElementById("markdown-view").value = md;
        downloadBlob(md, "document.md", "text/markdown;charset=utf-8");
        setStatus("Export Markdown telecharge.");
    });
}

function exportHtml() {
    if (!tinyEditor) return;
    const html = tinyEditor.getContent();
    downloadBlob(html, "document.html", "text/html;charset=utf-8");
    setStatus("Export HTML telecharge.");
}

function exportPdf() {
    // Prepare un contenu imprimable et declenche window.print()
    if (!tinyEditor) return;
    const html = tinyEditor.getContent();
    const printDiv = document.getElementById("print-content");
    printDiv.innerHTML = html;
    window.print();
    printDiv.innerHTML = "";
}

// ---------------------------------------------------------------------------
// Modal Metadonnees
// ---------------------------------------------------------------------------

function openMetadata() {
    document.getElementById("meta-title").value    = metadata.title;
    document.getElementById("meta-subtitle").value = metadata.subtitle;
    document.getElementById("meta-author").value   = metadata.author;
    document.getElementById("meta-date").value     = metadata.date;
    document.getElementById("meta-lang").value     = metadata.lang;
    document.getElementById("meta-abstract").value = metadata.abstract;
    document.getElementById("metadata-overlay").classList.add("open");
}

function closeMetadata() {
    document.getElementById("metadata-overlay").classList.remove("open");
}

function saveMetadata() {
    metadata.title    = document.getElementById("meta-title").value.trim();
    metadata.subtitle = document.getElementById("meta-subtitle").value.trim();
    metadata.author   = document.getElementById("meta-author").value.trim();
    metadata.date     = document.getElementById("meta-date").value.trim();
    metadata.lang     = document.getElementById("meta-lang").value.trim();
    metadata.abstract = document.getElementById("meta-abstract").value.trim();
    closeMetadata();
    setStatus("Metadonnees mises a jour.");
    // Regenere le Markdown pour inclure le YAML
    clearTimeout(renderTimer);
    generateMarkdown(true);
}

// Fermer en cliquant hors du dialogue
document.addEventListener("DOMContentLoaded", function() {
    const overlay = document.getElementById("metadata-overlay");
    overlay.addEventListener("click", function(e) {
        if (e.target === overlay) closeMetadata();
    });
    document.addEventListener("keydown", function(e) {
        if (e.key === "Escape") closeMetadata();
    });
});

// ---------------------------------------------------------------------------
// Barre d'etat
// ---------------------------------------------------------------------------

let statusTimer = null;

function setStatus(msg, duration) {
    document.getElementById("status-msg").textContent = msg;
    clearTimeout(statusTimer);
    if (duration !== 0) {
        statusTimer = setTimeout(function() {
            document.getElementById("status-msg").textContent = "Pret";
        }, duration || 3000);
    }
}

function setAutosaveStatus(msg) {
    document.getElementById("autosave-msg").textContent = msg;
}

// ---------------------------------------------------------------------------
// Splitter redimensionnable
// ---------------------------------------------------------------------------

(function initSplitter() {
    const handle    = document.getElementById("split-handle");
    const leftPane  = document.getElementById("left-pane");
    const rightPane = document.getElementById("right-pane");
    let dragging = false;

    handle.addEventListener("mousedown", function(e) {
        dragging = true;
        handle.classList.add("dragging");
        e.preventDefault();
    });

    document.addEventListener("mousemove", function(e) {
        if (!dragging) return;
        const container = document.getElementById("split-container");
        const rect = container.getBoundingClientRect();
        const offset = e.clientX - rect.left;
        const total  = rect.width - handle.offsetWidth;
        const pct    = Math.max(15, Math.min(85, (offset / total) * 100));
        leftPane.style.flex  = "none";
        leftPane.style.width = pct + "%";
        rightPane.style.flex = "1";
    });

    document.addEventListener("mouseup", function() {
        if (dragging) {
            dragging = false;
            handle.classList.remove("dragging");
        }
    });
})();

// ---------------------------------------------------------------------------
// Initialisation TinyMCE
// ---------------------------------------------------------------------------

tinymce.init({
    selector: "#editor",
    license_key: "gpl",   // TinyMCE Community (autohéberge) — supprime la notification
    plugins: "lists link table",
    toolbar: [
        "blocks | bold italic underline strikethrough | smallcaps",
        "bullist numlist | blockquote | link unlink | table | inlinenote deflist | removeformat"
    ].join(" | "),
    menubar: false,
    statusbar: false,
    height: "100%",
    resize: false,
    block_formats: "Paragraphe=p; Titre 1=h1; Titre 2=h2; Titre 3=h3; Titre 4=h4; Titre 5=h5; Titre 6=h6",
    content_style: [
        "body { font-family: Georgia, serif; font-size: 14px; line-height: 1.6; margin: 1em 1.5em; }",
        ".smallcaps { font-variant: small-caps; }",
        ".pandoc-note { color: #888; font-style: italic; }",
        "blockquote { border-left: 3px solid #ccc; margin-left: 0; padding-left: 1em; color: #555; }",
        "table { border-collapse: collapse; }",
        "td, th { border: 1px solid #ccc; padding: 4px 8px; }",
        "th { background: #f5f5f5; font-weight: bold; }",
    ].join(" "),

    setup: function(editor) {

        // -- Format petites capitales --
        editor.on("init", function() {
            editor.formatter.register("smallcaps", {
                inline: "span",
                classes: "smallcaps",
                styles: { "font-variant": "small-caps" },
            });
        });

        editor.ui.registry.addButton("smallcaps", {
            text: "sc",
            tooltip: "Petites capitales",
            onAction: function() { editor.formatter.toggle("smallcaps"); },
        });

        // -- Note inline Pandoc ^[...] --
        editor.ui.registry.addButton("inlinenote", {
            text: "Note",
            tooltip: "Note inline Pandoc ^[...]",
            onAction: function() {
                const selected = editor.selection.getContent({ format: "text" }).trim();
                if (selected) {
                    editor.selection.setContent(
                        '<span class="pandoc-note">^[' + escapeHtml(selected) + ']</span>'
                    );
                } else {
                    editor.windowManager.open({
                        title: "Note inline",
                        body: {
                            type: "panel",
                            items: [{ type: "input", name: "note", label: "Texte de la note" }],
                        },
                        buttons: [
                            { type: "submit", text: "Inserer", primary: true },
                            { type: "cancel",  text: "Annuler" },
                        ],
                        onSubmit: function(api) {
                            const note = (api.getData().note || "").trim();
                            if (note) {
                                editor.insertContent(
                                    '<span class="pandoc-note">^[' + escapeHtml(note) + ']</span>'
                                );
                            }
                            api.close();
                        },
                    });
                }
            },
        });

        // -- Liste de definition Pandoc --
        editor.ui.registry.addButton("deflist", {
            text: "Def",
            tooltip: "Liste de definition",
            onAction: function() {
                editor.windowManager.open({
                    title: "Liste de definition",
                    body: {
                        type: "panel",
                        items: [
                            { type: "input",    name: "term", label: "Terme" },
                            { type: "textarea", name: "def",  label: "Definition" },
                        ],
                    },
                    buttons: [
                        { type: "submit", text: "Inserer", primary: true },
                        { type: "cancel",  text: "Annuler" },
                    ],
                    onSubmit: function(api) {
                        const d = api.getData();
                        const term = (d.term || "").trim();
                        const def  = (d.def  || "").trim();
                        if (term) {
                            editor.insertContent(
                                "<p><strong>" + escapeHtml(term) + "</strong></p>" +
                                "<p style=\"margin-left:2em\">:&nbsp;&nbsp;&nbsp;" +
                                escapeHtml(def) + "</p><p>&nbsp;</p>"
                            );
                        }
                        api.close();
                    },
                });
            },
        });

        // -- Signaux de changement -> regeneration Markdown --
        editor.on("input change NodeChange", function() {
            onWysiwygChange();
        });

        // -- Reference globale une fois init --
        editor.on("init", function() {
            tinyEditor = editor;
            generateMarkdown(true);
        });
    },
});

// ---------------------------------------------------------------------------
// Zone Markdown : evenements focus
// ---------------------------------------------------------------------------

document.addEventListener("DOMContentLoaded", function() {
    const mdView = document.getElementById("markdown-view");
    mdView.addEventListener("focus", onMarkdownFocusIn);
    mdView.addEventListener("blur",  onMarkdownFocusOut);
});

// ---------------------------------------------------------------------------
// Utilitaire
// ---------------------------------------------------------------------------

function escapeHtml(str) {
    return str
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}
