#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QAction, QKeySequence, QFont, QTextDocument
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox,
    QPlainTextEdit, QTextEdit, QSplitter
)
from PySide6.QtPrintSupport import QPrinter


# -----------------------------
# DOCX export (Markdown -> docx)
# -----------------------------

INLINE_RE = re.compile(r"(\*\*.+?\*\*|\*.+?\*|`.+?`)", re.DOTALL)

def _add_inlines_docx(paragraph, text: str):
    """
    Minimal inline markdown -> docx runs:
    **bold**, *italic*, `code`
    (Pas de gestion d’imbrication complexe : volontairement simple.)
    """
    from docx.shared import Pt  # lazy import

    parts = INLINE_RE.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith("*") and part.endswith("*") and len(part) >= 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        elif part.startswith("`") and part.endswith("`") and len(part) >= 2:
            run = paragraph.add_run(part[1:-1])
            run.font.name = "Consolas"
            run.font.size = Pt(10)
        else:
            paragraph.add_run(part)

def export_docx_from_markdown(md: str, out_path: Path):
    """
    Convertisseur volontairement “sobre” :
    - Titres #..######
    - Listes -/* et 1.
    - Paragraphes
    - Blocs de code ``` ... ```
    - Inline ** * ``
    """
    from docx import Document
    from docx.shared import Pt

    doc = Document()

    in_code = False
    code_lines = []

    lines = md.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    def flush_code():
        nonlocal code_lines
        if not code_lines:
            return
        # Un bloc de code simple (une série de paragraphes monospace)
        for cl in code_lines:
            p = doc.add_paragraph()
            run = p.add_run(cl)
            run.font.name = "Consolas"
            run.font.size = Pt(10)
        code_lines = []

    for line in lines:
        # Code fence
        if line.strip().startswith("```"):
            if not in_code:
                in_code = True
                code_lines = []
            else:
                in_code = False
                flush_code()
            continue

        if in_code:
            code_lines.append(line)
            continue

        # Titres
        m = re.match(r"^(#{1,6})\s+(.*)$", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            doc.add_heading(text, level=level)
            continue

        # Listes
        m_bullet = re.match(r"^\s*[-*]\s+(.*)$", line)
        if m_bullet:
            p = doc.add_paragraph(style="List Bullet")
            _add_inlines_docx(p, m_bullet.group(1).strip())
            continue

        m_num = re.match(r"^\s*\d+\.\s+(.*)$", line)
        if m_num:
            p = doc.add_paragraph(style="List Number")
            _add_inlines_docx(p, m_num.group(1).strip())
            continue

        # Ligne vide => séparation (Word gère assez bien sans forcer)
        if not line.strip():
            doc.add_paragraph("")
            continue

        # Paragraphe normal
        p = doc.add_paragraph()
        _add_inlines_docx(p, line)

    # Si code non fermé
    if in_code:
        flush_code()

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


# -----------------------------
# Preview (éditable) + autosave
# -----------------------------

class PreviewEdit(QTextEdit):
    """
    Aperçu rendu MAIS éditable (pour couper/copier/coller facilement).
    La mise à jour depuis le Markdown est suspendue quand le focus est à droite,
    pour ne pas écraser tes collages.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.on_focus_in = None
        self.on_focus_out = None

        self.setTextInteractionFlags(
            Qt.TextEditorInteraction | Qt.TextSelectableByMouse | Qt.LinksAccessibleByMouse
        )

    def focusInEvent(self, event):
        if callable(self.on_focus_in):
            self.on_focus_in()
        super().focusInEvent(event)

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        if callable(self.on_focus_out):
            self.on_focus_out()


@dataclass
class AutosaveConfig:
    enabled: bool = True
    idle_ms: int = 1000            # autosave 1s après la dernière frappe
    use_main_file_if_possible: bool = True  # si un fichier est ouvert/sauvé, autosave dessus
    fallback_filename: str = "MiniMarkdown_autosave.md"  # si aucun fichier courant


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mini Markdown — split editor")

        self.current_path: Path | None = None
        self.cfg = AutosaveConfig()

        # Widgets
        self.editor = QPlainTextEdit()
        self.preview = PreviewEdit()

        # Typo
        mono = QFont("Consolas")
        mono.setStyleHint(QFont.Monospace)
        mono.setPointSize(11)
        self.editor.setFont(mono)
        self.preview.document().setDefaultFont(QFont("Arial", 11))

        # Split
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.editor)
        splitter.addWidget(self.preview)
        splitter.setSizes([650, 650])
        self.setCentralWidget(splitter)

        # Render suspend pendant édition à droite
        self._suspend_render = False
        self._pending_render = False
        self.preview.on_focus_in = self._preview_focus_in
        self.preview.on_focus_out = self._preview_focus_out

        # Timers : rendu + autosave
        self._render_timer = QTimer(self)
        self._render_timer.setSingleShot(True)
        self._render_timer.setInterval(120)
        self._render_timer.timeout.connect(self._render_preview_now)

        self._autosave_timer = QTimer(self)
        self._autosave_timer.setSingleShot(True)
        self._autosave_timer.setInterval(self.cfg.idle_ms)
        self._autosave_timer.timeout.connect(self._autosave_now)

        self._dirty = False
        self._last_autosaved_hash = None

        self.editor.textChanged.connect(self._on_text_changed)

        # Menus
        self._build_actions()
        self.statusBar().showMessage("Prêt")

        # Contenu initial
        self.editor.setPlainText(
            "# Mini Markdown\n\n"
            "Éditeur à gauche, aperçu à droite.\n\n"
            "- **Gras**, *italique*, `code`\n"
            "- Listes\n"
            "- [lien](https://example.org)\n\n"
            "À droite : tu peux aussi couper/copier/coller (tampon),\n"
            "mais ce que tu y modifies n’est pas réinjecté dans le Markdown.\n"
        )
        self._render_preview_now(force=True)

    # ---------- UI actions ----------

    def _build_actions(self):
        # Fichier
        act_open = QAction("Ouvrir…", self)
        act_open.setShortcut(QKeySequence.Open)
        act_open.triggered.connect(self.open_file)

        act_save = QAction("Enregistrer", self)
        act_save.setShortcut(QKeySequence.Save)
        act_save.triggered.connect(self.save_file)

        act_save_as = QAction("Enregistrer sous…", self)
        act_save_as.setShortcut(QKeySequence.SaveAs)
        act_save_as.triggered.connect(self.save_file_as)

        act_quit = QAction("Quitter", self)
        act_quit.setShortcut(QKeySequence.Quit)
        act_quit.triggered.connect(self.close)

        # Autosave toggle
        self.act_autosave = QAction("Autosave", self)
        self.act_autosave.setCheckable(True)
        self.act_autosave.setChecked(self.cfg.enabled)
        self.act_autosave.triggered.connect(self.toggle_autosave)

        m_file = self.menuBar().addMenu("Fichier")
        m_file.addAction(act_open)
        m_file.addAction(act_save)
        m_file.addAction(act_save_as)
        m_file.addSeparator()
        m_file.addAction(self.act_autosave)
        m_file.addSeparator()
        m_file.addAction(act_quit)

        # Export
        act_export_html = QAction("Exporter en HTML…", self)
        act_export_html.triggered.connect(self.export_html)

        act_export_pdf = QAction("Exporter en PDF…", self)
        act_export_pdf.triggered.connect(self.export_pdf)

        act_export_docx = QAction("Exporter en DOCX…", self)
        act_export_docx.triggered.connect(self.export_docx)

        m_export = self.menuBar().addMenu("Export")
        m_export.addAction(act_export_html)
        m_export.addAction(act_export_pdf)
        m_export.addAction(act_export_docx)

        # Édition (agit sur widget focus : gauche ou droite)
        m_edit = self.menuBar().addMenu("Édition")

        act_cut = QAction("Couper", self)
        act_cut.setShortcut(QKeySequence.Cut)
        act_cut.triggered.connect(self._smart_cut)

        act_copy = QAction("Copier", self)
        act_copy.setShortcut(QKeySequence.Copy)
        act_copy.triggered.connect(self._smart_copy)

        act_paste = QAction("Coller", self)
        act_paste.setShortcut(QKeySequence.Paste)
        act_paste.triggered.connect(self._smart_paste)

        m_edit.addAction(act_cut)
        m_edit.addAction(act_copy)
        m_edit.addAction(act_paste)

    def toggle_autosave(self, checked: bool):
        self.cfg.enabled = checked
        if not checked:
            self._autosave_timer.stop()
            self.statusBar().showMessage("Autosave désactivé", 1200)
        else:
            self.statusBar().showMessage("Autosave activé", 1200)
            if self._dirty:
                self._autosave_timer.start(self.cfg.idle_ms)

    def _smart_cut(self):
        w = QApplication.focusWidget()
        if hasattr(w, "cut"):
            w.cut()

    def _smart_copy(self):
        w = QApplication.focusWidget()
        if hasattr(w, "copy"):
            w.copy()

    def _smart_paste(self):
        w = QApplication.focusWidget()
        if hasattr(w, "paste"):
            w.paste()

    # ---------- Rendering + focus ----------

    def _preview_focus_in(self):
        self._suspend_render = True
        self.statusBar().showMessage("Aperçu : édition active (rendu suspendu)", 1200)

    def _preview_focus_out(self):
        self._suspend_render = False
        if self._pending_render:
            self._pending_render = False
            self._render_preview_now(force=True)

    def _on_text_changed(self):
        self._dirty = True

        # rendu
        if self._suspend_render:
            self._pending_render = True
        else:
            self._render_timer.start()

        # autosave
        if self.cfg.enabled:
            self._autosave_timer.start(self.cfg.idle_ms)

    def _render_preview_now(self, force: bool = False):
        if self._suspend_render and not force:
            self._pending_render = True
            return
        md = self.editor.toPlainText()
        # On met à jour la preview (écrase le “tampon” si focus à gauche)
        self.preview.blockSignals(True)
        self.preview.setMarkdown(md)
        self.preview.blockSignals(False)

    # ---------- File ops ----------

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Ouvrir un fichier Markdown", "", "Markdown (*.md *.markdown);;Tous les fichiers (*)"
        )
        if not path:
            return
        p = Path(path)
        try:
            text = p.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            text = p.read_text(encoding="utf-8", errors="replace")

        self.editor.setPlainText(text)
        self.current_path = p
        self._dirty = False
        self._last_autosaved_hash = None

        self.statusBar().showMessage(f"Ouvert : {p.name}", 1500)
        self._render_preview_now(force=True)

    def save_file(self):
        if self.current_path is None:
            self.save_file_as()
            return
        try:
            self.current_path.write_text(self.editor.toPlainText(), encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible d’enregistrer :\n{e}")
            return
        self._dirty = False
        self.statusBar().showMessage(f"Enregistré : {self.current_path.name}", 1500)

    def save_file_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Enregistrer sous", "", "Markdown (*.md);;Tous les fichiers (*)"
        )
        if not path:
            return
        p = Path(path)
        if p.suffix.lower() not in (".md", ".markdown"):
            p = p.with_suffix(".md")
        self.current_path = p
        self.save_file()

    # ---------- Autosave ----------

    def _autosave_target(self) -> Path:
        """
        Si on a un fichier courant : autosave dessus (comme demandé “en temps réel”).
        Sinon : fallback dans le HOME.
        """
        if self.current_path and self.cfg.use_main_file_if_possible:
            return self.current_path
        return Path.home() / self.cfg.fallback_filename

    def _autosave_now(self):
        if not self.cfg.enabled:
            return

        md = self.editor.toPlainText()
        h = hash(md)
        if h == self._last_autosaved_hash:
            return

        target = self._autosave_target()
        try:
            target.write_text(md, encoding="utf-8")
        except Exception as e:
            # On ne spam pas de popups : juste un message bref
            self.statusBar().showMessage(f"Autosave échoué : {e}", 2500)
            return

        self._last_autosaved_hash = h
        self.statusBar().showMessage(f"Autosave : {target.name}", 900)

    # ---------- Exports ----------

    def export_html(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter en HTML", "", "HTML (*.html);;Tous les fichiers (*)"
        )
        if not path:
            return
        out = Path(path)
        if out.suffix.lower() != ".html":
            out = out.with_suffix(".html")

        md = self.editor.toPlainText()
        doc = QTextDocument()
        doc.setMarkdown(md)
        html = doc.toHtml()

        try:
            out.write_text(html, encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible d’exporter HTML :\n{e}")
            return
        self.statusBar().showMessage(f"Export HTML : {out.name}", 1500)

    def export_pdf(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter en PDF", "", "PDF (*.pdf);;Tous les fichiers (*)"
        )
        if not path:
            return
        out = Path(path)
        if out.suffix.lower() != ".pdf":
            out = out.with_suffix(".pdf")

        md = self.editor.toPlainText()
        doc = QTextDocument()
        doc.setMarkdown(md)

        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(str(out))

        try:
            doc.print_(printer)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible d’exporter PDF :\n{e}")
            return

        self.statusBar().showMessage(f"Export PDF : {out.name}", 1500)

    def export_docx(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter en DOCX", "", "Word (*.docx);;Tous les fichiers (*)"
        )
        if not path:
            return
        out = Path(path)
        if out.suffix.lower() != ".docx":
            out = out.with_suffix(".docx")

        md = self.editor.toPlainText()
        try:
            export_docx_from_markdown(md, out)
        except ModuleNotFoundError:
            QMessageBox.critical(
                self, "DOCX",
                "Il manque la dépendance python-docx.\n\nInstalle : pip install python-docx"
            )
            return
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible d’exporter DOCX :\n{e}")
            return

        self.statusBar().showMessage(f"Export DOCX : {out.name}", 1500)


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.resize(1300, 780)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
