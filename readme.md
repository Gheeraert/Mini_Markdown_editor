# Mini Markdown

Petit éditeur Markdown minimaliste

- **fenêtre gauche** : source Markdown
- **fenêtre droite** : aperçu rendu (éditable comme “tampon” pour couper/copier/coller)
- **mise à jour quasi temps réel**
- **autosave** (après un court temps d’inactivité)
- **exports** : HTML, PDF, DOCX (via Pandoc recommandé)

> Objectif : un outil léger, lisible, modifiable, et suffisamment “pro” sans usine à gaz.

---

## Fonctionnalités

### Édition
- Raccourcis natifs : **Ctrl+X / Ctrl+C / Ctrl+V**, sélection à la souris, menu contextuel.
- Split view redimensionnable.
- Aperçu à droite **éditable** (pratique pour copier/coller depuis du rendu), sans casser l’édition Markdown à gauche.

> Note : ce qui est modifié à droite **n’est pas réinjecté** automatiquement dans le Markdown.  
> Pour éviter d’écraser les collages, la mise à jour de l’aperçu est **suspendue tant que le focus est à droite**, puis se rafraîchit quand le focus revient à gauche.

### Autosave
- Sauvegarde automatique **après X ms sans frappe** (par défaut ~1 seconde).
- Si un fichier est ouvert/enregistré : autosave sur ce fichier.
- Sinon : autosave dans un fichier de secours (ex. `~/MiniMarkdown_autosave.md`).

### Exports
- **HTML** : export du document rendu.
- **PDF** : export “imprimé” via Qt (pas de dépendance supplémentaire).
- **DOCX** : export recommandé via **Pandoc** (meilleure fidélité : listes, tableaux, notes, etc.).
- **TEX** : export via **Pandoc**
- **ODT** : export via **Pandoc**
- **EPUB** : export via **Pandoc**
---

## Prérequis

- Python 3.10+ (3.11/3.12 conseillés)
- Dépendance UI : **PySide6**
- Option DOCX via Pandoc : **pandoc** (exécutable) + wrapper Python au choix

### Installation (minimum)
```bash
pip install pyside6
