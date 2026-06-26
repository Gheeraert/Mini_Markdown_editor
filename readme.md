# Mini Markdown — deux editeurs complementaires en split

Ce depot contient deux programmes independants et complementaires.
Ils coexistent sans interférence.

---

## 1. `main.py` — Editeur Markdown source avec apercu

Interface en deux panneaux :

| Gauche            | Droite              |
|-------------------|---------------------|
| Source Markdown   | Apercu rendu        |

- L'utilisateur ecrit le Markdown brut a gauche (syntaxe visible).
- L'apercu rendu se met a jour en quasi temps reel a droite.
- Le panneau droit est editable comme **tampon** (couper/copier/coller),
  mais les modifications n'y sont pas reinjectees dans la source.
- Quand le focus est a droite, le rendu est suspendu pour ne pas ecraser
  les manipulations ; il reprend quand le focus revient a gauche.
- Exports : HTML, PDF (Qt), DOCX (python-docx ou Pandoc), LaTeX, ODT, EPUB.
- Gestion de bibliographie via Pandoc + fichiers `.bib` / `.csl`.
- Autosave sur le fichier courant ou fichier de secours.

Lancer :

```
python main.py
```

---

## 2. `Mini_Wysiwyg_Markdown.py` — Editeur WYSIWYG avec apercu Markdown

Interface en deux panneaux **symetrique** a `main.py` :

| Gauche                    | Droite                             |
|---------------------------|------------------------------------|
| Editeur WYSIWYG riche     | Pandoc Markdown genere en temps reel |

- L'utilisateur ecrit dans une interface visuelle (gras, titres, listes, etc.).
- Le Pandoc Markdown correspondant s'affiche en temps reel a droite.
- Le panneau droit est editable comme **tampon** (couper/copier/coller le Markdown),
  mais les modifications n'y sont pas reinjectees dans le WYSIWYG.
- Quand le focus est a droite, la generation automatique est suspendue ;
  elle reprend quand le focus revient a gauche.
- **Source de verite** : le document WYSIWYG de gauche.
- **Format de sauvegarde** : HTML (`.html`) — format le plus fidele.
- **Format d'export** : Pandoc Markdown (`.md`).
- Un fichier `.md` compagnon est genere automatiquement a cote du `.html`
  a chaque sauvegarde.

### Fonctionnalites

- Mise en forme : gras (`**`), italique (`*`), souligne (`{.underline}`),
  barre (`~~`), petites capitales (`{.smallcaps}`)
- Titres H1 a H6, paragraphe normal
- Listes a puces et numerotees
- Citations / blockquotes (`>`)
- Liens hypertexte (`[texte](url)`)
- Notes inline Pandoc (`^[texte de la note]`)
- Tables simples (pipe-tables Pandoc)
- Listes de definition (syntaxe Pandoc)
- Metadonnees YAML : title, subtitle, author, date, lang, abstract
- Exports : Markdown Pandoc, HTML, PDF (Qt), DOCX (via Pandoc si disponible)
- Autosave HTML + compagnon Markdown

Lancer :

```
python Mini_Wysiwyg_Markdown.py
```

---

## Logique commune aux deux programmes

Les deux editeurs partagent la meme philosophie de panneau droit :

- Editable comme tampon pour couper/copier/coller.
- Les modifications dans le panneau droit ne modifient pas le panneau gauche.
- Le rendu/generation est suspendu quand le focus est a droite.
- Le rendu/generation reprend quand le focus revient a gauche.

La difference est la direction :

- `main.py` : Markdown -> rendu HTML (apercu visuel)
- `Mini_Wysiwyg_Markdown.py` : WYSIWYG -> Markdown Pandoc (apercu technique)

---

## Limites connues

### `Mini_Wysiwyg_Markdown.py`

- Ce n'est pas Microsoft Word.
- Pas de bibliographie, pas de citeproc, pas de CSL, pas de Zotero,
  pas de gestion de references bibliographiques.
- Pas de tableaux complexes : pas de fusion de cellules,
  pas de styles avancees, pas de tableaux imbriques.
- La conversion WYSIWYG -> Markdown peut perdre certains raffinements
  visuels fins (couleurs, marges personnalisees, polices speciales).
- Le format de travail le plus fidele est **HTML**.
  Le Markdown est une projection/export.
- Les notes inline `^[...]` sont stockees comme texte literal dans le document
  (syntaxe Pandoc directe, affichee en gris italique dans le WYSIWYG).
- Les listes de definition utilisent la syntaxe Pandoc directe dans le document
  (`:   definition` affiche avec indentation).

### `main.py`

- La gestion de bibliographie necessite Pandoc + un fichier `.bib` valide.
- L'export PDF via Pandoc necessite un moteur LaTeX (TeX Live ou MiKTeX).
- L'export DOCX simple (sans Pandoc) necessite `python-docx` (`pip install python-docx`).
