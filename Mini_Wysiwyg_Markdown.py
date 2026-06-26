#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mini_Wysiwyg_Markdown.py -- Editeur WYSIWYG split avec apercu Pandoc Markdown.

Panneau gauche  : QTextEdit WYSIWYG (source de verite)
Panneau droit   : QPlainTextEdit affichant le Pandoc Markdown genere en temps reel
                  (editable comme tampon de copier/coller, non reinjecte dans le WYSIWYG)

Logique de rendu (symetrique a main.py) :
  - Le Markdown droit se met a jour automatiquement quand le WYSIWYG gauche change.
  - Quand le focus passe a droite : le rendu automatique est suspendu.
  - Quand le focus revient a gauche : le Markdown est regenere.

Formats de fichier :
  - Sauvegarde du document de travail : HTML (.html)
  - Export : Pandoc Markdown (.md)

Limites :
  - Pas de bibliographie, CSL, citeproc, Zotero.
  - Pas de tableaux complexes (pas de fusion de cellules).
  - La conversion WYSIWYG->Markdown peut perdre certains raffinements visuels.
  - Les notes inline ^[...] sont stockees comme texte literal dans le document.
"""

import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import (
    QAction, QColor, QFont, QKeySequence,
    QTextBlockFormat, QTextCharFormat, QTextCursor,
    QTextDocument, QTextListFormat, QTextTableFormat,
)
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QFormLayout, QInputDialog, QLabel,
    QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit,
    QSpinBox, QSplitter, QTextEdit, QVBoxLayout, QWidget,
)


# ---------------------------------------------------------------------------
# Utilitaires
# ---------------------------------------------------------------------------

BLOCKQUOTE_LEFT_MARGIN = 40.0


def find_pandoc():
    env = os.environ.get("PANDOC_PATH")
    if env and Path(env).exists():
        return env
    return shutil.which("pandoc")


def _escape_pipe(text):
    return text.replace("|", "\\|").replace("\n", " ")


# ---------------------------------------------------------------------------
# Convertisseur QTextDocument -> Pandoc Markdown
# ---------------------------------------------------------------------------

class PandocMarkdownExporter:

    def export(self, doc, metadata=None):
        parts = []
        if metadata:
            y = self._yaml_block(metadata)
            if y:
                parts.append(y)
        self._walk_frame(doc.rootFrame(), parts)
        text = "\n".join(parts)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip() + "\n"

    def _walk_frame(self, frame, parts):
        from PySide6.QtGui import QTextTable
        it = frame.begin()
        while not it.atEnd():
            child = it.currentFrame()
            if child is not None:
                if isinstance(child, QTextTable):
                    parts.append(self._table_to_md(child))
                else:
                    self._walk_frame(child, parts)
            else:
                block = it.currentBlock()
                if block.isValid():
                    line = self._block_to_md(block)
                    if line is not None:
                        parts.append(line)
            it += 1

    def _block_to_md(self, block):
        bfmt = block.blockFormat()
        heading = bfmt.headingLevel()
        text_list = block.textList()
        inline = self._inline_to_md(block)

        if not inline.strip():
            return ""

        if heading > 0:
            return "#" * heading + " " + inline

        if not text_list and bfmt.leftMargin() >= BLOCKQUOTE_LEFT_MARGIN:
            return "> " + inline

        if text_list:
            lfmt = text_list.format()
            indent = "    " * max(0, lfmt.indent() - 1)
            if lfmt.style() in (
                QTextListFormat.ListDisc,
                QTextListFormat.ListCircle,
                QTextListFormat.ListSquare,
            ):
                return indent + "- " + inline
            else:
                return indent + "1. " + inline

        return inline

    def _inline_to_md(self, block):
        raw = []
        it = block.begin()
        while not it.atEnd():
            frag = it.fragment()
            if frag.isValid():
                text = frag.text()
                # U+2028 line sep, U+2029 para sep, U+FFFC object replacement
                text = text.replace(" ", "\n").replace(" ", "\n")
                if text and text != "￼":
                    raw.append((text, frag.charFormat()))
            it += 1

        if not raw:
            return ""

        merged = [[raw[0][0], raw[0][1]]]
        for text, fmt in raw[1:]:
            if self._same_fmt(fmt, merged[-1][1]):
                merged[-1][0] += text
            else:
                merged.append([text, fmt])

        return "".join(self._apply_char_fmt(t, f) for t, f in merged)

    @staticmethod
    def _same_fmt(a, b):
        return (
            a.fontWeight() == b.fontWeight()
            and a.fontItalic() == b.fontItalic()
            and a.fontUnderline() == b.fontUnderline()
            and a.fontStrikeOut() == b.fontStrikeOut()
            and a.fontCapitalization() == b.fontCapitalization()
            and a.anchorHref() == b.anchorHref()
        )

    @staticmethod
    def _apply_char_fmt(text, fmt):
        href = fmt.anchorHref()
        is_bold = fmt.fontWeight() >= QFont.Bold
        is_italic = fmt.fontItalic()
        is_underline = fmt.fontUnderline()
        is_strike = fmt.fontStrikeOut()
        is_smallcaps = fmt.fontCapitalization() == QFont.SmallCaps

        if is_smallcaps:
            text = "[" + text + "]{.smallcaps}"
        if is_strike:
            text = "~~" + text + "~~"
        if is_underline:
            text = "[" + text + "]{.underline}"
        if is_italic:
            text = "*" + text + "*"
        if is_bold:
            text = "**" + text + "**"
        if href:
            text = "[" + text + "](" + href + ")"
        return text

    def _table_to_md(self, table):
        rows, cols = table.rows(), table.columns()
        if rows == 0 or cols == 0:
            return ""

        grid = []
        for r in range(rows):
            row = []
            for c in range(cols):
                row.append(_escape_pipe(self._cell_text(table.cellAt(r, c))))
            grid.append(row)

        widths = [max(max(len(grid[r][c]) for r in range(rows)), 3) for c in range(cols)]

        def fmt_row(cells):
            return "| " + " | ".join(c.ljust(w) for c, w in zip(cells, widths)) + " |"

        sep = "| " + " | ".join("-" * w for w in widths) + " |"
        lines = [fmt_row(grid[0]), sep] + [fmt_row(row) for row in grid[1:]]
        return "\n".join(lines)

    def _cell_text(self, cell):
        parts = []
        frame = cell.frame()
        if frame:
            it = frame.begin()
            while not it.atEnd():
                if it.currentFrame() is None:
                    block = it.currentBlock()
                    if block.isValid():
                        parts.append(self._inline_to_md(block))
                it += 1
        return " ".join(p for p in parts if p).strip()

    @staticmethod
    def _yaml_block(meta):
        if not any(v.strip() for v in meta.values()):
            return ""
        lines = ["---"]
        for key in ("title", "subtitle", "author", "date", "lang"):
            val = meta.get(key, "").strip()
            if val:
                safe = val.replace('"', '\\"')
                lines.append(key + ': "' + safe + '"')
        abstract = meta.get("abstract", "").strip()
        if abstract:
            lines.append("abstract: |")
            for ln in abstract.splitlines():
                lines.append("  " + ln)
        lines.append("---")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Panneau Markdown de droite
# ---------------------------------------------------------------------------

class MarkdownBufferEdit(QPlainTextEdit):
    """
    Affiche le Pandoc Markdown genere en temps reel.
    Editable comme tampon de copier/coller.
    Le rendu automatique est suspendu quand ce widget a le focus.
    Les modifications ne sont pas reinjectees dans le WYSIWYG.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.on_focus_in = None
        self.on_focus_out = None
        mono = QFont("Consolas")
        mono.setStyleHint(QFont.Monospace)
        mono.setPointSize(11)
        self.setFont(mono)
        self.setReadOnly(False)

    def focusInEvent(self, event):
        if callable(self.on_focus_in):
            self.on_focus_in()
        super().focusInEvent(event)

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        if callable(self.on_focus_out):
            self.on_focus_out()


# ---------------------------------------------------------------------------
# Dialogues
# ---------------------------------------------------------------------------

class MetadataDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Metadonnees du document")
        self.setMinimumWidth(450)
        layout = QFormLayout(self)
        self.fields = {}
        for key, label in [
            ("title",    "Titre"),
            ("subtitle", "Sous-titre"),
            ("author",   "Auteur"),
            ("date",     "Date"),
            ("lang",     "Langue (ex. fr-FR)"),
        ]:
            w = QLineEdit(data.get(key, ""))
            self.fields[key] = w
            layout.addRow(label, w)
        self.abstract = QPlainTextEdit(data.get("abstract", ""))
        self.abstract.setMaximumHeight(90)
        layout.addRow("Resume (abstract)", self.abstract)
        btn = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn.accepted.connect(self.accept)
        btn.rejected.connect(self.reject)
        layout.addRow(btn)

    def get_data(self):
        d = {k: w.text() for k, w in self.fields.items()}
        d["abstract"] = self.abstract.toPlainText()
        return d


class TableDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Inserer un tableau")
        layout = QFormLayout(self)
        self.rows_spin = QSpinBox()
        self.rows_spin.setRange(1, 50)
        self.rows_spin.setValue(3)
        self.cols_spin = QSpinBox()
        self.cols_spin.setRange(1, 20)
        self.cols_spin.setValue(2)
        self.header_chk = QCheckBox("Premiere ligne comme en-tete")
        self.header_chk.setChecked(True)
        self.caption_edit = QLineEdit()
        self.caption_edit.setPlaceholderText("Legende (optionnelle)")
        layout.addRow("Lignes (corps)", self.rows_spin)
        layout.addRow("Colonnes", self.cols_spin)
        layout.addRow(self.header_chk)
        layout.addRow("Legende", self.caption_edit)
        btn = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn.accepted.connect(self.accept)
        btn.rejected.connect(self.reject)
        layout.addRow(btn)

    def get_params(self):
        return {
            "rows": self.rows_spin.value(),
            "cols": self.cols_spin.value(),
            "header": self.header_chk.isChecked(),
            "caption": self.caption_edit.text().strip(),
        }


class DefinitionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Inserer une liste de definition")
        layout = QFormLayout(self)
        self.term_edit = QLineEdit()
        self.term_edit.setPlaceholderText("Terme a definir")
        self.def_edit = QPlainTextEdit()
        self.def_edit.setMaximumHeight(80)
        self.def_edit.setPlaceholderText("Definition du terme")
        layout.addRow("Terme", self.term_edit)
        layout.addRow("Definition", self.def_edit)
        btn = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn.accepted.connect(self.accept)
        btn.rejected.connect(self.reject)
        layout.addRow(btn)

    def get_data(self):
        return {
            "term": self.term_edit.text().strip(),
            "definition": self.def_edit.toPlainText().strip(),
        }


class LinkDialog(QDialog):
    def __init__(self, selected_text="", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Inserer un lien")
        layout = QFormLayout(self)
        self.text_edit = QLineEdit(selected_text)
        self.text_edit.setPlaceholderText("Texte affiche")
        if selected_text:
            self.text_edit.setEnabled(False)
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("https://example.com")
        layout.addRow("Texte", self.text_edit)
        layout.addRow("URL", self.url_edit)
        btn = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn.accepted.connect(self.accept)
        btn.rejected.connect(self.reject)
        layout.addRow(btn)

    def get_data(self):
        return {"text": self.text_edit.text().strip(), "url": self.url_edit.text().strip()}


# ---------------------------------------------------------------------------
# Configuration autosave
# ---------------------------------------------------------------------------

@dataclass
class AutosaveConfig:
    enabled: bool = True
    idle_ms: int = 3000
    companion_md: bool = True   # genere un .md a cote du .html a chaque sauvegarde


# ---------------------------------------------------------------------------
# Fenetre principale
# ---------------------------------------------------------------------------

PARA_STYLES = [
    ("Paragraphe normal", 0),
    ("Titre 1", 1), ("Titre 2", 2), ("Titre 3", 3),
    ("Titre 4", 4), ("Titre 5", 5), ("Titre 6", 6),
]


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mini WYSIWYG Markdown -- editeur split")

        self.current_path = None
        self.metadata = {
            "title": "", "subtitle": "", "author": "",
            "date": "", "lang": "", "abstract": "",
        }
        self._dirty = False
        self._last_autosave_hash = None
        self._suspend_render = False
        self._pending_render = False

        self.cfg = AutosaveConfig()
        self.exporter = PandocMarkdownExporter()
        self.pandoc_path = find_pandoc()
        self.has_pandoc = self.pandoc_path is not None

        # -- Widgets ---------------------------------------------------------
        self.editor = QTextEdit()
        self.editor.setAcceptRichText(True)
        self.editor.document().setDefaultFont(QFont("Georgia", 12))

        self.markdown_view = MarkdownBufferEdit()

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.editor)
        splitter.addWidget(self.markdown_view)
        splitter.setSizes([650, 650])
        self.setCentralWidget(splitter)

        # -- Focus callbacks (meme mecanique que main.py) --------------------
        self.markdown_view.on_focus_in = self._markdown_focus_in
        self.markdown_view.on_focus_out = self._markdown_focus_out

        # -- Timers ----------------------------------------------------------
        self._render_timer = QTimer(self)
        self._render_timer.setSingleShot(True)
        self._render_timer.setInterval(150)
        self._render_timer.timeout.connect(self._render_markdown_now)

        self._autosave_timer = QTimer(self)
        self._autosave_timer.setSingleShot(True)
        self._autosave_timer.setInterval(self.cfg.idle_ms)
        self._autosave_timer.timeout.connect(self._autosave_now)

        # -- Connexions ------------------------------------------------------
        self.editor.document().contentsChanged.connect(self._on_wysiwyg_changed)
        self.editor.cursorPositionChanged.connect(self._refresh_style_combo)

        # -- Menus / barre d outils / barre d etat ---------------------------
        self._build_menus()
        self._build_toolbar()

        sb = self.statusBar()
        sb.showMessage("Pret")
        pandoc_lbl = QLabel("Pandoc OK" if self.has_pandoc else "Pandoc absent")
        pandoc_lbl.setToolTip(self.pandoc_path or "Pandoc introuvable")
        sb.addPermanentWidget(pandoc_lbl)

        # -- Contenu initial -------------------------------------------------
        self.editor.setHtml(
            "<h1>Mini WYSIWYG Markdown</h1>"
            "<p>Ecrivez ici a <b>gauche</b>."
            " Le <b>Pandoc Markdown</b> s'affiche en temps reel a droite.</p>"
            "<p>La fenetre droite est un <i>tampon</i> :"
            " vous pouvez y copier/coller le Markdown,"
            " mais les modifications n'affectent pas le document de gauche.</p>"
        )
        self._render_markdown_now(force=True)
        self._dirty = False
        self._update_title()

    # -----------------------------------------------------------------------
    # Coeur du mecanisme : rendu Markdown en temps reel
    # -----------------------------------------------------------------------

    def _markdown_focus_in(self):
        self._suspend_render = True
        self.statusBar().showMessage(
            "Markdown : edition tampon active, generation suspendue", 1200
        )

    def _markdown_focus_out(self):
        self._suspend_render = False
        if self._pending_render:
            self._pending_render = False
            self._render_markdown_now(force=True)

    def _on_wysiwyg_changed(self):
        self._dirty = True
        self._update_title()
        if self._suspend_render:
            self._pending_render = True
        else:
            self._render_timer.start()
        if self.cfg.enabled:
            self._autosave_timer.start(self.cfg.idle_ms)

    def _render_markdown_now(self, force=False):
        if self._suspend_render and not force:
            self._pending_render = True
            return
        md = self.exporter.export(self.editor.document(), self.metadata)
        self.markdown_view.blockSignals(True)
        try:
            self.markdown_view.setPlainText(md)
        finally:
            self.markdown_view.blockSignals(False)

    # -----------------------------------------------------------------------
    # Titre de fenetre
    # -----------------------------------------------------------------------

    def _update_title(self):
        base = self.current_path.name if self.current_path else "Sans titre"
        marker = " *" if self._dirty else ""
        self.setWindowTitle("Mini WYSIWYG Markdown -- " + base + marker)

    # -----------------------------------------------------------------------
    # Style combo
    # -----------------------------------------------------------------------

    def _refresh_style_combo(self):
        self.style_combo.blockSignals(True)
        level = self.editor.textCursor().blockFormat().headingLevel()
        self.style_combo.setCurrentIndex(min(level, self.style_combo.count() - 1))
        self.style_combo.blockSignals(False)

    def _on_style_combo(self, idx):
        self.set_heading(idx)

    # -----------------------------------------------------------------------
    # Menus
    # -----------------------------------------------------------------------

    def _build_menus(self):
        mb = self.menuBar()

        m = mb.addMenu("Fichier")
        self._act(m, "Nouveau",             QKeySequence.New,    self.new_document)
        self._act(m, "Ouvrir...",            QKeySequence.Open,   self.open_file)
        m.addSeparator()
        self._act(m, "Enregistrer",          QKeySequence.Save,   self.save_file)
        self._act(m, "Enregistrer sous...",   QKeySequence.SaveAs, self.save_file_as)
        m.addSeparator()
        self.act_autosave = QAction("Autosave", self)
        self.act_autosave.setCheckable(True)
        self.act_autosave.setChecked(self.cfg.enabled)
        self.act_autosave.triggered.connect(self._toggle_autosave)
        m.addAction(self.act_autosave)
        m.addSeparator()
        self._act(m, "Quitter", QKeySequence.Quit, self.close)

        m = mb.addMenu("Edition")
        self._act(m, "Annuler",          QKeySequence.Undo,      self.editor.undo)
        self._act(m, "Retablir",         QKeySequence.Redo,      self.editor.redo)
        m.addSeparator()
        self._act(m, "Couper",           QKeySequence.Cut,       self._smart_cut)
        self._act(m, "Copier",           QKeySequence.Copy,      self._smart_copy)
        self._act(m, "Coller",           QKeySequence.Paste,     self._smart_paste)
        self._act(m, "Selectionner tout", QKeySequence.SelectAll, self._smart_select_all)
        m.addSeparator()
        self._act(m, "Supprimer le formatage", None, self.clear_formatting)
        sub = m.addMenu("Changer la casse")
        self._act(sub, "minuscules",               None, lambda: self._change_case("lower"))
        self._act(sub, "MAJUSCULES",               None, lambda: self._change_case("upper"))
        self._act(sub, "Capitalisation des mots",  None, lambda: self._change_case("title"))

        m = mb.addMenu("Format")
        self._act(m, "Gras",              QKeySequence.Bold,      self.toggle_bold)
        self._act(m, "Italique",          QKeySequence.Italic,    self.toggle_italic)
        self._act(m, "Souligne",          QKeySequence.Underline, self.toggle_underline)
        self._act(m, "Barre",             None,                   self.toggle_strikethrough)
        self._act(m, "Petites capitales", None,                   self.toggle_smallcaps)
        m.addSeparator()
        for label, level in PARA_STYLES:
            lv = level
            self._act(m, label, None, lambda checked=False, l=lv: self.set_heading(l))
        m.addSeparator()
        self._act(m, "Liste a puces",        None, self.insert_bullet_list)
        self._act(m, "Liste numerotee",       None, self.insert_numbered_list)
        self._act(m, "Citation (blockquote)", None, self.insert_blockquote)

        m = mb.addMenu("Insertion")
        self._act(m, "Lien hypertexte...",     None, self.insert_link)
        self._act(m, "Note inline...",          None, self.insert_inline_note)
        self._act(m, "Tableau simple...",       None, self.insert_table)
        self._act(m, "Liste de definition...", None, self.insert_definition_list)
        self._act(m, "Ligne horizontale",       None, self.insert_horizontal_rule)

        m = mb.addMenu("Document")
        self._act(m, "Metadonnees...", None, self.edit_metadata)

        m = mb.addMenu("Export")
        self._act(m, "Exporter en Markdown Pandoc...", None, self.export_markdown)
        self._act(m, "Exporter en HTML...",            None, self.export_html)
        self._act(m, "Exporter en PDF (Qt)...",        None, self.export_pdf)
        act = QAction("Exporter en DOCX (Pandoc)...", self)
        act.setEnabled(self.has_pandoc)
        act.triggered.connect(self.export_docx_pandoc)
        m.addAction(act)

    def _act(self, menu, label, shortcut, slot):
        act = QAction(label, self)
        if shortcut:
            act.setShortcut(shortcut)
        act.triggered.connect(slot)
        menu.addAction(act)
        return act

    # -----------------------------------------------------------------------
    # Barre d outils
    # -----------------------------------------------------------------------

    def _build_toolbar(self):
        tb = self.addToolBar("Format")
        tb.setMovable(False)

        self.style_combo = QComboBox()
        for label, _ in PARA_STYLES:
            self.style_combo.addItem(label)
        self.style_combo.currentIndexChanged.connect(self._on_style_combo)
        tb.addWidget(self.style_combo)
        tb.addSeparator()

        def ta(label, tip, slot, shortcut=None):
            act = QAction(label, self)
            act.setToolTip(tip)
            if shortcut:
                act.setShortcut(shortcut)
            act.triggered.connect(slot)
            self.addAction(act)
            tb.addAction(act)

        ta("G",    "Gras (Ctrl+B)",      self.toggle_bold,         QKeySequence.Bold)
        ta("I",    "Italique (Ctrl+I)",  self.toggle_italic,       QKeySequence.Italic)
        ta("S",    "Souligne (Ctrl+U)",  self.toggle_underline,    QKeySequence.Underline)
        ta("~~",   "Barre",              self.toggle_strikethrough)
        ta("sc",   "Petites capitales",  self.toggle_smallcaps)
        tb.addSeparator()
        ta("<<",   "Citation/blockquote", self.insert_blockquote)
        ta("*",    "Liste a puces",       self.insert_bullet_list)
        ta("1.",   "Liste numerotee",     self.insert_numbered_list)
        tb.addSeparator()
        ta("Lien", "Lien hypertexte",         self.insert_link)
        ta("Note", "Note inline ^[...]",       self.insert_inline_note)
        ta("Tbl",  "Tableau simple",           self.insert_table)
        ta("Def",  "Liste de definition",      self.insert_definition_list)
        ta("Meta", "Metadonnees du document",  self.edit_metadata)

    # -----------------------------------------------------------------------
    # Mise en forme de caractere
    # -----------------------------------------------------------------------

    def _merge_char_fmt(self, fmt):
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)
        self.editor.setFocus()

    def toggle_bold(self):
        fmt = QTextCharFormat()
        w = self.editor.currentCharFormat().fontWeight()
        fmt.setFontWeight(QFont.Normal if w >= QFont.Bold else QFont.Bold)
        self._merge_char_fmt(fmt)

    def toggle_italic(self):
        fmt = QTextCharFormat()
        fmt.setFontItalic(not self.editor.currentCharFormat().fontItalic())
        self._merge_char_fmt(fmt)

    def toggle_underline(self):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not self.editor.currentCharFormat().fontUnderline())
        self._merge_char_fmt(fmt)

    def toggle_strikethrough(self):
        fmt = QTextCharFormat()
        fmt.setFontStrikeOut(not self.editor.currentCharFormat().fontStrikeOut())
        self._merge_char_fmt(fmt)

    def toggle_smallcaps(self):
        fmt = QTextCharFormat()
        cur = self.editor.currentCharFormat().fontCapitalization()
        new_cap = QFont.MixedCase if cur == QFont.SmallCaps else QFont.SmallCaps
        fmt.setFontCapitalization(new_cap)
        self._merge_char_fmt(fmt)

    def clear_formatting(self):
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.BlockUnderCursor)
        cursor.setCharFormat(QTextCharFormat())
        self.editor.setTextCursor(cursor)
        self.editor.setFocus()

    # -----------------------------------------------------------------------
    # Mise en forme de bloc
    # -----------------------------------------------------------------------

    def set_heading(self, level):
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        bfmt = QTextBlockFormat()
        bfmt.setHeadingLevel(level)
        cursor.mergeBlockFormat(bfmt)
        cfmt = QTextCharFormat()
        sizes = {0: 12, 1: 22, 2: 18, 3: 15, 4: 13, 5: 12, 6: 11}
        cfmt.setFontPointSize(sizes.get(level, 12))
        cfmt.setFontWeight(QFont.Bold if level > 0 else QFont.Normal)
        cursor.mergeBlockCharFormat(cfmt)
        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)
        self.editor.setFocus()

    def _change_case(self, mode):
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            return
        text = cursor.selectedText()
        if mode == "lower":
            text = text.lower()
        elif mode == "upper":
            text = text.upper()
        else:
            text = text.title()
        cursor.insertText(text)
        self.editor.setFocus()

    def insert_blockquote(self):
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        bfmt = QTextBlockFormat()
        bfmt.setLeftMargin(BLOCKQUOTE_LEFT_MARGIN)
        bfmt.setRightMargin(BLOCKQUOTE_LEFT_MARGIN)
        cursor.mergeBlockFormat(bfmt)
        cursor.endEditBlock()
        self.editor.setFocus()

    def insert_bullet_list(self):
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.ListDisc)
        self.editor.textCursor().createList(fmt)
        self.editor.setFocus()

    def insert_numbered_list(self):
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.ListDecimal)
        self.editor.textCursor().createList(fmt)
        self.editor.setFocus()

    def insert_horizontal_rule(self):
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.movePosition(QTextCursor.EndOfBlock)
        cursor.insertBlock()
        cursor.insertHtml("<hr/>")
        cursor.insertBlock()
        cursor.endEditBlock()
        self.editor.setFocus()

    # -----------------------------------------------------------------------
    # Insertions speciales
    # -----------------------------------------------------------------------

    def insert_link(self):
        cursor = self.editor.textCursor()
        selected = cursor.selectedText().strip()
        dlg = LinkDialog(selected, self)
        if dlg.exec() != QDialog.Accepted:
            return
        data = dlg.get_data()
        text = data["text"] or selected or "lien"
        url = data["url"]
        if not url:
            return
        fmt = QTextCharFormat()
        fmt.setAnchor(True)
        fmt.setAnchorHref(url)
        fmt.setFontUnderline(True)
        fmt.setForeground(QColor("#0000CC"))
        cursor.beginEditBlock()
        if cursor.hasSelection():
            cursor.mergeCharFormat(fmt)
        else:
            cursor.insertText(text, fmt)
            cursor.setCharFormat(QTextCharFormat())
        cursor.endEditBlock()
        self.editor.setFocus()

    def insert_inline_note(self):
        cursor = self.editor.textCursor()
        selected = cursor.selectedText().strip()
        if not selected:
            text, ok = QInputDialog.getText(self, "Note inline", "Texte de la note :")
            if not ok or not text.strip():
                return
            selected = text.strip()
        fmt = QTextCharFormat()
        fmt.setForeground(QColor("#888888"))
        fmt.setFontItalic(True)
        cursor.beginEditBlock()
        if cursor.hasSelection():
            cursor.removeSelectedText()
        cursor.insertText("^[" + selected + "]", fmt)
        cursor.setCharFormat(QTextCharFormat())
        cursor.endEditBlock()
        self.editor.setFocus()

    def insert_table(self):
        dlg = TableDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        p = dlg.get_params()
        rows = p["rows"] + (1 if p["header"] else 0)
        cols = p["cols"]
        fmt = QTextTableFormat()
        fmt.setBorder(1)
        fmt.setBorderStyle(QTextTableFormat.BorderStyle_Solid)
        fmt.setCellPadding(4)
        fmt.setCellSpacing(0)
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        table = cursor.insertTable(rows, cols, fmt)
        if p["header"]:
            hfmt = QTextCharFormat()
            hfmt.setFontWeight(QFont.Bold)
            for c in range(cols):
                cell_cursor = table.cellAt(0, c).firstCursorPosition()
                cell_cursor.mergeCharFormat(hfmt)
                cell_cursor.insertText("Col " + str(c + 1))
        if p["caption"]:
            end_cursor = table.lastCursorPosition()
            end_cursor.movePosition(QTextCursor.NextBlock)
            end_cursor.insertText(": " + p["caption"])
        cursor.endEditBlock()
        self.editor.setFocus()

    def insert_definition_list(self):
        dlg = DefinitionDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        data = dlg.get_data()
        if not data["term"]:
            return
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.movePosition(QTextCursor.EndOfBlock)
        cursor.insertBlock()
        tfmt = QTextCharFormat()
        tfmt.setFontWeight(QFont.Bold)
        cursor.insertText(data["term"], tfmt)
        cursor.insertBlock()
        bfmt = QTextBlockFormat()
        bfmt.setLeftMargin(20)
        cursor.mergeBlockFormat(bfmt)
        cursor.insertText(":   " + data["definition"], QTextCharFormat())
        cursor.insertBlock()
        cursor.mergeBlockFormat(QTextBlockFormat())
        cursor.endEditBlock()
        self.editor.setFocus()

    # -----------------------------------------------------------------------
    # Metadonnees
    # -----------------------------------------------------------------------

    def edit_metadata(self):
        dlg = MetadataDialog(self.metadata, self)
        if dlg.exec() == QDialog.Accepted:
            self.metadata = dlg.get_data()
            self.statusBar().showMessage("Metadonnees mises a jour", 1500)
            self._render_markdown_now(force=True)

    # -----------------------------------------------------------------------
    # Couper / Copier / Coller intelligents (agissent sur le widget qui a le focus)
    # -----------------------------------------------------------------------

    def _focused_editor(self):
        w = QApplication.focusWidget()
        return w if hasattr(w, "copy") else self.editor

    def _smart_cut(self):
        self._focused_editor().cut()

    def _smart_copy(self):
        self._focused_editor().copy()

    def _smart_paste(self):
        self._focused_editor().paste()

    def _smart_select_all(self):
        self._focused_editor().selectAll()

    # -----------------------------------------------------------------------
    # Fichiers
    # -----------------------------------------------------------------------

    def new_document(self):
        if self._dirty and not self._confirm_discard():
            return
        self.editor.clear()
        self.current_path = None
        self.metadata = {k: "" for k in self.metadata}
        self._dirty = False
        self._update_title()
        self._render_markdown_now(force=True)

    def open_file(self):
        if self._dirty and not self._confirm_discard():
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Ouvrir un document de travail",
            "", "HTML (*.html);;Tous les fichiers (*)"
        )
        if not path:
            return
        p = Path(path)
        try:
            content = p.read_text(encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", "Impossible d'ouvrir :\n" + str(e))
            return
        self.editor.setHtml(content)
        self.current_path = p
        self._dirty = False
        self._update_title()
        self._render_markdown_now(force=True)
        self.statusBar().showMessage("Ouvert : " + p.name, 1500)

    def save_file(self):
        if self.current_path is None:
            self.save_file_as()
        else:
            self._write_html(self.current_path)

    def save_file_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Enregistrer le document de travail",
            "", "HTML (*.html);;Tous les fichiers (*)"
        )
        if not path:
            return
        p = Path(path)
        if p.suffix.lower() not in (".html", ".htm"):
            p = p.with_suffix(".html")
        self.current_path = p
        self._write_html(p)

    def _write_html(self, path):
        try:
            path.write_text(self.editor.toHtml(), encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", "Impossible d'enregistrer :\n" + str(e))
            return
        if self.cfg.companion_md:
            try:
                path.with_suffix(".md").write_text(
                    self.exporter.export(self.editor.document(), self.metadata),
                    encoding="utf-8",
                )
            except Exception:
                pass
        self._dirty = False
        self._update_title()
        self.statusBar().showMessage("Enregistre : " + path.name, 1500)

    # -----------------------------------------------------------------------
    # Autosave
    # -----------------------------------------------------------------------

    def _toggle_autosave(self, checked):
        self.cfg.enabled = checked
        self.statusBar().showMessage(
            "Autosave active" if checked else "Autosave desactive", 1200
        )

    def _autosave_now(self):
        if not self.cfg.enabled:
            return
        html = self.editor.toHtml()
        h = hash(html)
        if h == self._last_autosave_hash:
            return
        target = self.current_path or (Path.home() / "MiniWysiwyg_autosave.html")
        try:
            target.write_text(html, encoding="utf-8")
            if self.cfg.companion_md:
                target.with_suffix(".md").write_text(
                    self.exporter.export(self.editor.document(), self.metadata),
                    encoding="utf-8",
                )
            self._last_autosave_hash = h
            self.statusBar().showMessage("Autosave : " + target.name, 900)
        except Exception as e:
            self.statusBar().showMessage("Autosave echoue : " + str(e), 2000)

    # -----------------------------------------------------------------------
    # Exports
    # -----------------------------------------------------------------------

    def _current_markdown(self):
        return self.exporter.export(self.editor.document(), self.metadata)

    def export_markdown(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter en Pandoc Markdown",
            "", "Markdown (*.md);;Tous les fichiers (*)"
        )
        if not path:
            return
        out = Path(path)
        if out.suffix.lower() not in (".md", ".markdown"):
            out = out.with_suffix(".md")
        try:
            out.write_text(self._current_markdown(), encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", "Impossible d'exporter :\n" + str(e))
            return
        self.statusBar().showMessage("Export Markdown : " + out.name, 1500)

    def export_html(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter en HTML", "", "HTML (*.html);;Tous les fichiers (*)"
        )
        if not path:
            return
        out = Path(path)
        if out.suffix.lower() not in (".html", ".htm"):
            out = out.with_suffix(".html")
        try:
            out.write_text(self.editor.toHtml(), encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", "Impossible d'exporter HTML :\n" + str(e))
            return
        self.statusBar().showMessage("Export HTML : " + out.name, 1500)

    def export_pdf(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter en PDF", "", "PDF (*.pdf);;Tous les fichiers (*)"
        )
        if not path:
            return
        out = Path(path)
        if out.suffix.lower() != ".pdf":
            out = out.with_suffix(".pdf")
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(str(out))
        try:
            self.editor.document().print_(printer)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", "Impossible d'exporter PDF :\n" + str(e))
            return
        self.statusBar().showMessage("Export PDF : " + out.name, 1500)

    def export_docx_pandoc(self):
        if not self.has_pandoc:
            QMessageBox.information(self, "Pandoc", "Pandoc n'est pas disponible.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter en DOCX (Pandoc)", "", "Word (*.docx);;Tous les fichiers (*)"
        )
        if not path:
            return
        out = Path(path)
        if out.suffix.lower() != ".docx":
            out = out.with_suffix(".docx")
        try:
            cmd = [
                self.pandoc_path,
                "--from", "markdown",
                "--to", "docx",
                "--standalone",
                "--output", str(out),
            ]
            subprocess.run(cmd, input=self._current_markdown().encode("utf-8"), check=True)
        except Exception as e:
            QMessageBox.critical(self, "Erreur DOCX", str(e))
            return
        self.statusBar().showMessage("Export DOCX : " + out.name, 1500)

    # -----------------------------------------------------------------------
    # Utilitaires
    # -----------------------------------------------------------------------

    def _confirm_discard(self):
        return QMessageBox.question(
            self, "Document modifie",
            "Des modifications ne sont pas enregistrees. Continuer quand meme ?",
            QMessageBox.Yes | QMessageBox.No,
        ) == QMessageBox.Yes

    def closeEvent(self, event):
        if self._dirty and not self._confirm_discard():
            event.ignore()
        else:
            event.accept()


# ---------------------------------------------------------------------------
# Point d entree
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.resize(1300, 780)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
