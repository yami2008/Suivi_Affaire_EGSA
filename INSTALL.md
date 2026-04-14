# Manuel d'installation — Environnement Suivi Affaires EGSA

> Reproduire cet environnement sur un nouveau PC Windows en moins de 15 minutes.

---

## Ce que ça installe

Un environnement complet pour gérer des tâches bureautiques avec VS Code + GitHub Copilot :
- Lire/créer/modifier : Excel, Word, PDF, PowerPoint, CSV, images
- OCR sur documents scannés (français)
- Génération de courriers, rapports, avis écrits en série

---

## Prérequis

- Windows 10 ou 11
- VS Code installé
- Extension **GitHub Copilot** activée dans VS Code
- Connexion internet

---

## Étape 1 — Python

### 1.1 Téléchargement

Aller sur : https://www.python.org/downloads/

Télécharger la dernière version stable (ex : Python 3.13.x — Windows installer 64-bit).

### 1.2 Installation

Lancer le `.exe` et **OBLIGATOIRE** : cocher **"Add Python to PATH"** avant de cliquer "Install Now".

```
[ ] Install launcher for all users
[x] Add Python 3.x to PATH   <-- cocher absolument
```

### 1.3 Vérification

Ouvrir PowerShell et taper :

```powershell
python --version
pip --version
```

Les deux doivent répondre avec un numéro de version.

---

## Étape 2 — Packages Python

### 2.1 Installation en une commande

```powershell
pip install openpyxl pandas python-docx pdfplumber pypdf pytesseract pillow reportlab xlrd odfpy matplotlib python-pptx jinja2
```

### 2.2 Vérification

```powershell
pip list
```

#### Packages installés et leur rôle

| Package | Version testée | Usage |
|---------|---------------|-------|
| `openpyxl` | 3.1.5 | Lire/créer/modifier Excel (.xlsx, .xlsm) |
| `pandas` | 3.0.2 | Analyser et traiter données Excel/CSV |
| `xlrd` | 2.0.2 | Lire vieux Excel (.xls) |
| `odfpy` | 1.4.1 | Lire fichiers .ods (OpenDocument) |
| `pdfplumber` | 0.11.9 | Lire PDF avec texte (extraction fiable) |
| `pypdf` | 6.10.0 | Manipuler PDF (pages, métadonnées) |
| `pytesseract` | 0.3.13 | OCR — PDFs scannés / images |
| `pillow` | 12.2.0 | Traitement d'images |
| `python-docx` | 1.2.0 | Créer/modifier Word (.docx) |
| `python-pptx` | 1.0.2 | Créer/modifier PowerPoint (.pptx) |
| `reportlab` | 4.4.10 | Générer PDF depuis zéro |
| `matplotlib` | 3.10.8 | Graphiques pour rapports et Excel |
| `jinja2` | 3.1.6 | Templates — courriers en série |

---

## Étape 3 — Tesseract OCR (reconnaissance texte sur images/scans)

### 3.1 Téléchargement

Aller sur : https://github.com/UB-Mannheim/tesseract/wiki

Télécharger : `tesseract-ocr-w64-setup-5.x.x.exe` (version 64-bit)

### 3.2 Installation

Lancer l'installateur et à l'écran **"Choose Components"** :
- Cocher **"Additional language data"** → cocher **"French"** (fra)

Conserve le chemin par défaut : `C:\Program Files\Tesseract-OCR`

### 3.3 Ajout au PATH

Ouvrir **PowerShell en tant qu'administrateur** (clic droit → "Exécuter en tant qu'administrateur") et lancer :

```powershell
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files\Tesseract-OCR", "Machine")
```

Fermer et rouvrir PowerShell, puis vérifier :

```powershell
tesseract --version
```

Résultat attendu : `tesseract v5.x.x`

---

## Étape 4 — Pandoc (conversions de formats)

### 4.1 Installation via winget

```powershell
winget install --id JohnMacFarlane.Pandoc
```

Accepter les conditions avec `Y` si demandé.

### 4.2 Vérification

**Redémarre PowerShell** (winget modifie le PATH), puis :

```powershell
pandoc --version
```

Résultat attendu : `pandoc 3.x.x`

---

## Étape 5 — Skills Copilot

### 5.1 Structure des fichiers

Les skills sont dans `.github/skills/` du workspace. Ils sont déjà configurés dans ce projet.
Si tu clones ou copies ce workspace sur un autre PC, ils sont automatiquement disponibles.

```
.github/
  skills/
    file-reading/
      SKILL.md    ← Routeur : quel outil Python pour quel type de fichier
    xlsx/
      SKILL.md    ← Créer/modifier/analyser Excel
    docx/
      SKILL.md    ← Créer/modifier/lire Word
```

### 5.2 Fonctionnement

VS Code + GitHub Copilot détecte automatiquement les fichiers dans `.github/skills/`.
Il applique les instructions contenues dans les `SKILL.md` quand tu travailles sur des fichiers correspondants.

---

## Récapitulatif — Checklist rapide

```
[ ] Python installé + ajouté au PATH
[ ] pip install (13 packages) — commande Étape 2
[ ] Tesseract installé + langue fra cochée
[ ] Tesseract ajouté au PATH (PowerShell admin)
[ ] Pandoc installé via winget
[ ] Terminal redémarré
[ ] tesseract --version  → répond v5.x.x
[ ] pandoc --version     → répond 3.x.x
[ ] python --version     → répond 3.x.x
[ ] Skills .github/skills/ présents dans le workspace
```

---

## Versions installées sur le PC de référence (14/04/2026)

| Logiciel | Version |
|----------|---------|
| Python | 3.13.3 |
| Tesseract OCR | 5.5.0.20241111 |
| Pandoc | 3.9.0.2 |
| openpyxl | 3.1.5 |
| pandas | 3.0.2 |
| python-docx | 1.2.0 |
| pdfplumber | 0.11.9 |
| pypdf | 6.10.0 |
| pytesseract | 0.3.13 |
| pillow | 12.2.0 |
| reportlab | 4.4.10 |
| xlrd | 2.0.2 |
| odfpy | 1.4.1 |
| matplotlib | 3.10.8 |
| python-pptx | 1.0.2 |
| jinja2 | 3.1.6 |

---

## Ce PC — Statut

Tout est installé ✅. Rien de supplémentaire requis pour ce projet.
