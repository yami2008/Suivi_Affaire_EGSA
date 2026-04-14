# Setup — Environnement de travail

## Statut : Skills installés — Prêt à utiliser

---

## Étape 1 — Logiciels à installer

### 1. Python ✅
- **Version installée :** Python 3.13.3
- **Important :** Cocher **"Add Python to PATH"** pendant l'installation

---

### 2. Packages Python ✅
Tous installés avec succès le 14/04/2026.

Commande utilisée :

```
pip install openpyxl pandas python-docx pdfplumber pypdf pytesseract pillow reportlab xlrd odfpy matplotlib python-pptx jinja2
```

#### Lecture et analyse de fichiers

| Package | Usage |
|---------|-------|
| `openpyxl` | Lire/créer/modifier Excel (.xlsx, .xlsm) |
| `pandas` | Analyser et traiter données Excel/CSV |
| `xlrd` | Lire vieux Excel (.xls) |
| `odfpy` | Lire fichiers .ods (OpenDocument) |
| `pdfplumber` | Lire PDF avec texte (extraction fiable) |
| `pypdf` | Manipuler PDF (pages, métadonnées) |
| `pytesseract` | OCR — PDFs scannés / images |
| `pillow` | Traitement d'images |

#### Création de documents

| Package | Usage |
|---------|-------|
| `python-docx` | Créer/modifier Word (.docx) — courriers, rapports, avis |
| `python-pptx` | Créer/modifier PowerPoint (.pptx) — présentations, bilans |
| `reportlab` | Générer PDF depuis zéro (mise en page précise) |
| `matplotlib` | Graphiques et visualisations pour rapports et Excel |
| `jinja2` | Templates de documents — courriers types, génération en série |

---

### 3. Tesseract OCR ✅
- **Version :** 5.5.0.20241111 (UB-Mannheim)
- **Installé dans :** `C:\Program Files\Tesseract-OCR`
- **Langue French (fra) :** installée
- **PATH :** ✅ configuré et vérifié — `tesseract --version` répond correctement

---

### 4. Pandoc ✅
- **Version :** 3.9.0.2
- **Installé via :** `winget install --id JohnMacFarlane.Pandoc`
- Utilisé pour conversions de formats (Word ↔ Markdown, etc.)
- **Note :** Redémarrer le terminal après installation pour que le PATH soit pris en compte

---

## Étape 2 — Skills Copilot (FAIT)

Trois skills adaptés Windows + VS Code Copilot, au format `dossier/SKILL.md` :

| Skill | Emplacement | Description |
|-------|-------------|-------------|
| `file-reading` | `.github/skills/file-reading/SKILL.md` | Routeur : quel outil pour quel type de fichier |
| `xlsx` | `.github/skills/xlsx/SKILL.md` | Créer/modifier/lire Excel (.xlsx, .xls, .ods, .csv) |
| `docx` | `.github/skills/docx/SKILL.md` | Créer/modifier/lire Word (.docx) |

### Ce qui a été fait :
1. ~~Coller chaque skill dans son fichier dédié~~ ✅
2. ~~Adapter les chemins Linux (`/mnt/...`) en chemins Windows~~ ✅
3. ~~Supprimer les dépendances non disponibles (`scripts/office/soffice.py`, LibreOffice CLI, etc.)~~ ✅
4. ~~Valider que les outils Python couvrent tous les cas d'usage~~ ✅

---

## Contexte du projet

- **Rôle :** Chef d'équipe développement et département
- **Objectif du workspace :** Gestion des tâches quotidiennes non-dev (rapports, courriers, avis écrits, suivi d'affaires)
- **Environnement :** Windows — VS Code avec GitHub Copilot
