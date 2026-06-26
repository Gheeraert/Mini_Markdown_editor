# Mini WYSIWYG Web

Application Flask locale — editeur WYSIWYG en deux panneaux produisant du Pandoc Markdown.

**Application locale uniquement. Ne pas exposer sur Internet.**

---

## Principe

```
Panneau gauche          Panneau droit
────────────────────    ────────────────────────────────
Editeur WYSIWYG    →    Pandoc Markdown (temps reel)
(source de verite)      (tampon copiable)
```

- Ce que vous tapez a gauche est converti en Pandoc Markdown et affiche a droite.
- Le panneau droit est editable comme tampon (copier/coller le Markdown).
- Quand le focus est a droite, la regeneration automatique est **suspendue**.
- Quand le focus revient a gauche, le Markdown est **regenere** si necessaire.

C'est la logique symetrique de `main.py` (qui fait Markdown -> apercu rendu).

---

## Installation et lancement

```bash
cd mini_wysiwyg_web
pip install flask
python app.py
```

Ouvrir dans le navigateur : **http://127.0.0.1:5000**

---

## Dependances

| Paquet  | Obligatoire | Role                    |
|---------|-------------|-------------------------|
| flask   | oui         | Serveur web local       |
| pandoc  | non         | Conversion HTML->Markdown (meilleur resultat) |

Sans Pandoc, une conversion de secours interne est utilisee (resultat approximatif).

### Installer Pandoc

- Telecharger : https://pandoc.org/installing.html
- Ou pointer vers un pandoc existant avec la variable d'environnement :

```bash
set PANDOC_PATH=C:\chemin\vers\pandoc.exe   # Windows
export PANDOC_PATH=/usr/local/bin/pandoc    # Linux/Mac
```

---

## TinyMCE et la notification CDN

L'editeur WYSIWYG utilise **TinyMCE Community Edition** charge depuis le CDN.
Sans cle API, TinyMCE affiche une petite notification.

La cle TinyMCE Cloud peut etre entree depuis l'interface avec le bouton **"Cle TinyMCE Cloud…"**.
Elle est enregistree localement dans `instance/settings.json`, qui ne doit pas etre versionne.
On peut aussi la fournir avec la variable d'environnement `TINYMCE_API_KEY`.

L'editeur **fonctionne meme sans cle** sur localhost.

---

## Fonctionnalites WYSIWYG

- Titres H1 a H6, paragraphe normal
- Gras, italique, souligne, barre
- Petites capitales (bouton `sc`)
- Listes a puces et numerotees
- Citations (blockquote)
- Liens hypertexte
- Tableaux simples
- Notes inline Pandoc : `^[texte de la note]` (bouton Note)
- Listes de definition Pandoc : terme + `:   definition` (bouton Def)
- Metadonnees YAML : title, subtitle, author, date, lang, abstract

---

## Markdown produit

Le Markdown vise la syntaxe **Pandoc Markdown** :

```markdown
---
title: "Titre"
author: "Auteur"
date: "2026"
lang: fr-FR
---

# Titre 1

**gras**, *italique*, [souligne]{.underline}, ~~barre~~

[petites capitales]{.smallcaps}

[texte du lien](https://example.com)

> Citation en bloc

^[Note inline.]

Terme
:   Definition du terme.

| Col 1 | Col 2 |
|-------|-------|
| a     | b     |
```

---

## Sauvegarde et export

| Action            | Format  | Destination                       |
|-------------------|---------|-----------------------------------|
| Sauvegarder HTML  | `.html` | `documents/document.html`         |
| Export .md        | `.md`   | Telechargement navigateur         |
| Export HTML       | `.html` | Telechargement navigateur         |
| PDF               | PDF     | Impression navigateur (Ctrl+P)    |
| Autosave          | les deux| `documents/autosave.html` + `.md` |

Le format de travail le plus fidele est **HTML**.
Le Markdown est une projection/export.

---

## Sauvegarde dans un dossier cloud local (optionnel)

L'application ne se connecte pas directement a Google Drive, Dropbox, OneDrive ou Nextcloud.
Elle peut simplement ecrire automatiquement une copie des fichiers dans un dossier local
deja synchronise par l'un de ces services.

C'est la methode recommandee pour retrouver facilement ses fichiers sur un autre ordinateur
sans gerer d'identifiants cloud dans l'application.

**Configuration :** bouton **"Dossier cloud…"** dans la barre d'outils.
Collez le chemin du dossier local synchronise, par exemple :

```
C:\Users\Tony\Google Drive\MiniWysiwyg
C:\Users\Tony\Dropbox\MiniWysiwyg
C:\Users\Tony\OneDrive\MiniWysiwyg
```

Le chemin est enregistre dans `instance/settings.json`.
A chaque autosave, `autosave.html` et `autosave.md` sont copies dans ce dossier.

---

## Structure des fichiers

```
mini_wysiwyg_web/
├── app.py              -- Serveur Flask, routes API
├── templates/
│   └── editor.html     -- Interface utilisateur
├── static/
│   ├── editor.js       -- Logique JS (TinyMCE, rendu, autosave)
│   └── editor.css      -- Mise en page
├── documents/          -- Fichiers sauvegardes (gitignore recommande)
│   ├── autosave.html
│   ├── autosave.md
│   └── ...
└── README.md
```

---

## Limites connues

- Ce n'est pas Microsoft Word.
- **Pas** de bibliographie, citeproc, CSL, Zotero, gestion de references.
- **Pas** de tableaux complexes (pas de fusion de cellules).
- La conversion WYSIWYG -> Markdown peut perdre des raffinements visuels.
- Le Markdown est une projection ; le document HTML est la source fidele.
- Necessite une connexion Internet pour charger TinyMCE depuis le CDN
  (sauf si vous hebergez TinyMCE localement).
- Application mono-utilisateur, locale uniquement.

---

## Securite

- Le serveur ecoute uniquement sur `127.0.0.1` (localhost).
- Les sauvegardes se font uniquement dans `documents/` (pas de path traversal).
- Ne pas exposer ce serveur sur un reseau public.
