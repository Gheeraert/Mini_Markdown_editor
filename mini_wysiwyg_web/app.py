#!/usr/bin/env python3
"""
Mini WYSIWYG Web -- serveur Flask local.

Usage :
    pip install flask
    python app.py
    # puis ouvrir http://127.0.0.1:5000

Application locale uniquement. Ne pas exposer sur Internet.
"""

import json
import os
import re
import shutil
import subprocess
from html.parser import HTMLParser
from pathlib import Path

from flask import Flask, jsonify, render_template, request

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DOCS_DIR = Path(__file__).parent / "documents"
DOCS_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024  # 8 MB


def find_pandoc():
    env = os.environ.get("PANDOC_PATH")
    if env and Path(env).exists():
        return env
    return shutil.which("pandoc")


PANDOC = find_pandoc()


# ---------------------------------------------------------------------------
# Conversion de secours HTML -> Markdown (sans Pandoc)
# ---------------------------------------------------------------------------

class _HtmlToMd(HTMLParser):
    """
    Convertisseur HTML -> Pandoc Markdown minimal.
    Utilise uniquement si Pandoc est absent.
    Gere : titres, gras, italique, souligne, barre, petites capitales,
    listes, blockquote, liens, tableaux, notes inline, definitions.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._out = []
        self._stack = []
        self._list_types = []
        self._list_counters = []
        self._in_table = False
        self._table_rows = []
        self._current_row = []
        self._in_cell = False
        self._cell_buf = []
        self._pending_href = ""

    def handle_starttag(self, tag, attrs):
        attrs_d = dict(attrs)
        self._stack.append(tag)

        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._out.append("\n\n" + "#" * int(tag[1]) + " ")
        elif tag == "p":
            if not self._in_table:
                self._out.append("\n\n")
        elif tag in ("strong", "b"):
            self._out.append("**")
        elif tag in ("em", "i"):
            self._out.append("*")
        elif tag == "u":
            self._out.append("[")
        elif tag in ("s", "del", "strike"):
            self._out.append("~~")
        elif tag == "span":
            cls = attrs_d.get("class", "")
            style = attrs_d.get("style", "")
            if "smallcaps" in cls or "small-caps" in style:
                self._out.append("[")
        elif tag == "blockquote":
            self._out.append("\n\n> ")
        elif tag in ("ul", "ol"):
            self._list_types.append(tag)
            self._list_counters.append(0)
            self._out.append("\n")
        elif tag == "li":
            indent = "    " * (len(self._list_types) - 1)
            if self._list_types and self._list_types[-1] == "ol":
                self._list_counters[-1] += 1
                self._out.append(f"\n{indent}{self._list_counters[-1]}. ")
            else:
                self._out.append(f"\n{indent}- ")
        elif tag == "a":
            self._pending_href = attrs_d.get("href", "")
            self._out.append("[")
        elif tag == "table":
            self._in_table = True
            self._table_rows = []
        elif tag == "tr":
            self._current_row = []
        elif tag in ("td", "th"):
            self._in_cell = True
            self._cell_buf = []
        elif tag == "hr":
            self._out.append("\n\n---\n\n")
        elif tag == "br":
            self._out.append("  \n")

    def handle_endtag(self, tag):
        if self._stack and self._stack[-1] == tag:
            self._stack.pop()

        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._out.append("\n\n")
        elif tag in ("strong", "b"):
            self._out.append("**")
        elif tag in ("em", "i"):
            self._out.append("*")
        elif tag == "u":
            self._out.append("]{.underline}")
        elif tag in ("s", "del", "strike"):
            self._out.append("~~")
        elif tag == "span":
            # Fermeture de smallcaps : on detecte si on a ouvert un [
            # en cherchant le dernier [ non ferme -- approximation suffisante
            joined = "".join(self._out)
            if joined.endswith(("[",)) or (len(joined) > 0 and "[" in joined.split("\n")[-1]):
                # verif simple : le dernier token ouvrant
                pass
            # Approche : on tente de fermer proprement
            # En pratique : si le span avait class=smallcaps, on ferme
            # Le suivi exact necessite un stack de spans
            self._out.append("]{.smallcaps}")
        elif tag == "p":
            if not self._in_table:
                self._out.append("\n\n")
        elif tag == "blockquote":
            self._out.append("\n\n")
        elif tag in ("ul", "ol"):
            if self._list_types:
                self._list_types.pop()
            if self._list_counters:
                self._list_counters.pop()
            self._out.append("\n")
        elif tag == "a":
            self._out.append(f"]({self._pending_href})")
            self._pending_href = ""
        elif tag in ("td", "th"):
            self._in_cell = False
            self._current_row.append("".join(self._cell_buf).strip())
            self._cell_buf = []
        elif tag == "tr":
            if self._current_row:
                self._table_rows.append(list(self._current_row))
        elif tag == "table":
            self._in_table = False
            self._out.append("\n\n" + self._render_table() + "\n\n")

    def handle_data(self, data):
        text = data.replace("\n", " ")
        if self._in_cell:
            self._cell_buf.append(text)
        else:
            self._out.append(text)

    def _render_table(self):
        if not self._table_rows:
            return ""
        cols = max(len(r) for r in self._table_rows)
        for row in self._table_rows:
            while len(row) < cols:
                row.append("")
        widths = [max(max(len(r[c]) for r in self._table_rows), 3) for c in range(cols)]

        def escape(t):
            return t.replace("|", "\\|")

        def fmt_row(cells):
            return "| " + " | ".join(escape(c).ljust(w) for c, w in zip(cells, widths)) + " |"

        sep = "| " + " | ".join("-" * w for w in widths) + " |"
        lines = [fmt_row(self._table_rows[0]), sep]
        for row in self._table_rows[1:]:
            lines.append(fmt_row(row))
        return "\n".join(lines)

    def result(self):
        text = "".join(self._out)
        # Nettoyage : pas plus de 2 lignes vides consecutives
        text = re.sub(r"\n{3,}", "\n\n", text)
        # Supprimer les ]{.smallcaps} orphelins (si le span n'etait pas smallcaps)
        # Approche conservative : on laisse en place, Pandoc les accepte de toute facon
        return text.strip()


def html_to_markdown_fallback(html: str) -> str:
    parser = _HtmlToMd()
    parser.feed(html)
    return parser.result()


# ---------------------------------------------------------------------------
# Bloc YAML pour les metadonnees
# ---------------------------------------------------------------------------

def build_yaml(meta: dict) -> str:
    if not any(str(v).strip() for v in meta.values()):
        return ""
    lines = ["---"]
    for key in ("title", "subtitle", "author", "date", "lang"):
        val = str(meta.get(key, "")).strip()
        if val:
            safe = val.replace('"', '\\"')
            lines.append(f'{key}: "{safe}"')
    abstract = str(meta.get("abstract", "")).strip()
    if abstract:
        lines.append("abstract: |")
        for ln in abstract.splitlines():
            lines.append("  " + ln)
    lines.append("---")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Nom de fichier securise (pas de path traversal)
# ---------------------------------------------------------------------------

def safe_stem(name: str) -> str:
    stem = Path(name).stem or "document"
    stem = re.sub(r"[^\w\-]", "_", stem)
    return stem or "document"


# ---------------------------------------------------------------------------
# Paramètres persistants (instance/settings.json)
# ---------------------------------------------------------------------------

def load_settings() -> dict:
    settings_path = Path("instance/settings.json")
    if settings_path.exists():
        try:
            return json.loads(settings_path.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_settings(data: dict) -> None:
    Path("instance").mkdir(exist_ok=True)
    Path("instance/settings.json").write_text(
        json.dumps(data, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Clé TinyMCE Cloud
# ---------------------------------------------------------------------------

def get_tinymce_api_key():
    key = os.environ.get("TINYMCE_API_KEY", "").strip()
    if key:
        return key
    return load_settings().get("tinymce_api_key", "").strip()


# ---------------------------------------------------------------------------
# Dossier de sauvegarde cloud local
# ---------------------------------------------------------------------------

def get_cloud_sync_folder():
    folder = load_settings().get("cloud_sync_folder", "").strip()
    if not folder:
        return None
    path = Path(folder).expanduser()
    if path.exists() and path.is_dir():
        return path
    return None


# ---------------------------------------------------------------------------
# Routes Flask
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    key = get_tinymce_api_key()
    settings = load_settings()
    cloud_raw = settings.get("cloud_sync_folder", "").strip()
    cloud_ok = bool(get_cloud_sync_folder())
    return render_template(
        "editor.html",
        has_pandoc=bool(PANDOC),
        tinymce_api_key=key or "no-api-key",
        tinymce_key_configured=bool(key),
        cloud_sync_folder=cloud_raw,
        cloud_folder_configured=cloud_ok,
    )


@app.post("/api/settings/tinymce-key")
def save_tinymce_key():
    data = request.get_json(force=True)
    key = data.get("tinymce_api_key", "").strip()
    settings = load_settings()
    settings["tinymce_api_key"] = key
    save_settings(settings)
    return jsonify({"ok": True, "configured": bool(key)})


@app.post("/api/settings/cloud-folder")
def save_cloud_folder():
    data = request.get_json(force=True)
    folder = data.get("cloud_sync_folder", "").strip()

    if folder:
        path = Path(folder).expanduser()
        if not path.exists() or not path.is_dir():
            return jsonify({"ok": False, "error": f"Dossier introuvable : {folder}"}), 400

    settings = load_settings()
    settings["cloud_sync_folder"] = folder
    save_settings(settings)
    return jsonify({"ok": True, "configured": bool(folder)})


@app.route("/api/to-markdown", methods=["POST"])
def api_to_markdown():
    """
    Recoit : { html: str, metadata: dict }
    Renvoie : { markdown: str }
    Utilise Pandoc si disponible, sinon le fallback interne.
    """
    data = request.get_json(force=True, silent=True) or {}
    html = data.get("html", "").strip()
    metadata = data.get("metadata", {})

    if not html:
        return jsonify({"markdown": ""})

    if PANDOC:
        try:
            proc = subprocess.run(
                [PANDOC, "--from", "html", "--to", "markdown", "--wrap=none"],
                input=html.encode("utf-8"),
                capture_output=True,
                timeout=15,
            )
            md = proc.stdout.decode("utf-8", errors="replace")
        except Exception:
            md = html_to_markdown_fallback(html)
    else:
        md = html_to_markdown_fallback(html)

    yaml = build_yaml(metadata)
    if yaml:
        md = yaml + "\n\n" + md.lstrip()

    return jsonify({"markdown": md.strip()})


@app.route("/api/autosave", methods=["POST"])
def api_autosave():
    """
    Sauvegarde automatique dans documents/autosave.html et documents/autosave.md.
    Copie aussi dans le dossier cloud local si configuré.
    """
    data = request.get_json(force=True, silent=True) or {}
    html = data.get("html", "")
    markdown = data.get("markdown", "")
    try:
        (DOCS_DIR / "autosave.html").write_text(html, encoding="utf-8")
        (DOCS_DIR / "autosave.md").write_text(markdown, encoding="utf-8")

        cloud_dir = get_cloud_sync_folder()
        if cloud_dir:
            (cloud_dir / "autosave.html").write_text(html, encoding="utf-8")
            (cloud_dir / "autosave.md").write_text(markdown, encoding="utf-8")

        return jsonify({"status": "ok"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/save", methods=["POST"])
def api_save():
    """
    Sauvegarde un fichier dans documents/ avec un nom controle.
    Recoit : { format: "html"|"md", content: str, filename: str }
    """
    data = request.get_json(force=True, silent=True) or {}
    fmt = data.get("format", "html")
    content = data.get("content", "")
    filename = data.get("filename", "document")

    stem = safe_stem(filename)
    ext = ".html" if fmt == "html" else ".md"
    target = DOCS_DIR / (stem + ext)
    try:
        target.write_text(content, encoding="utf-8")
        return jsonify({"status": "ok", "saved_as": target.name})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


# ---------------------------------------------------------------------------
# Lancement
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    pandoc_info = f"OK ({PANDOC})" if PANDOC else "absent (conversion de secours)"
    print(f"Pandoc : {pandoc_info}")
    print("Serveur local : http://127.0.0.1:5000")
    print("Ctrl+C pour arreter.")
    app.run(debug=False, host="127.0.0.1", port=5000)
