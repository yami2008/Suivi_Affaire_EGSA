---
name: file-reading
description: Use this skill when the user mentions a file in the workspace (Excel, Word, PDF, CSV, image, archive, etc.) and you need to read or analyze it. This skill is a router that tells you which tool and library to use for each file type. Triggers include any mention of a file path, a file name with a known extension, or a user asking about the contents of a file. Do NOT use this skill if the file content is already visible in your context (attachment, active selection, etc.).
---

# Lecture de fichiers

## Overview

Ce skill est un **routeur** : il indique quel outil utiliser selon le type de fichier.

**Ne pas utiliser** si le contenu du fichier est déjà visible dans le contexte (pièce jointe, sélection active, etc.).

## Environnement

- **OS :** Windows
- **Shell :** PowerShell
- **IDE :** VS Code avec GitHub Copilot
- **Outils natifs :** `read_file`, `view_image`, `run_in_terminal` (PowerShell/Python)
- **Librairies Python disponibles :** openpyxl, pandas, python-docx, pdfplumber, pypdf, pytesseract, pillow, reportlab, xlrd, odfpy

## Protocole général

1. **Regarde l'extension** — c'est ta clé de dispatch.
2. **Vérifie la taille avant de lire** — les gros fichiers nécessitent un échantillonnage.
   ```powershell
   (Get-Item "chemin\fichier.xlsx").Length
   ```
3. **Lis juste ce qu'il faut** pour répondre à la question de l'utilisateur.
4. **Si un skill dédié existe**, va le lire (voir table ci-dessous).

---

## Table de dispatch

| Extension | Première action | Skill dédié |
|-----------|----------------|-------------|
| `.pdf` | `pdfplumber` ou `pypdf` en Python | — |
| `.docx` | `python-docx` en Python | `docx` skill |
| `.doc` (legacy) | Convertir en `.docx` d'abord | `docx` skill |
| `.xlsx`, `.xlsm` | `openpyxl` sheet names + head | `xlsx` skill |
| `.xls` (legacy) | `pd.read_excel(engine="xlrd")` | `xlsx` skill |
| `.ods` | `pd.read_excel(engine="odf")` | `xlsx` skill |
| `.csv`, `.tsv` | `pandas` avec `nrows` | — (voir section CSV) |
| `.json`, `.jsonl` | Python `json` module | — (voir section JSON) |
| `.jpg`, `.png`, `.gif`, `.webp` | `view_image` (outil natif Copilot) | — (voir section Images) |
| `.zip`, `.tar`, `.tar.gz` | Lister le contenu, ne PAS extraire | — (voir section Archives) |
| `.txt`, `.md`, `.log`, code | `read_file` (outil natif Copilot) | — |
| Inconnu | `file` en Python ou inspecter les octets | — |

---

## PDF

**Ne jamais** lire un PDF avec `read_file` — c'est du binaire.

Aperçu rapide — nombre de pages et texte extractible :

```python
from pypdf import PdfReader
r = PdfReader(r"C:\chemin\vers\fichier.pdf")
print(f"{len(r.pages)} pages")
print(r.pages[0].extract_text()[:2000])
```

Pour une extraction plus fiable (tableaux, mise en page) :

```python
import pdfplumber
with pdfplumber.open(r"C:\chemin\vers\fichier.pdf") as pdf:
    print(f"{len(pdf.pages)} pages")
    page = pdf.pages[0]
    print(page.extract_text()[:2000])
```

Pour les **PDFs scannés** (images sans texte extractible) — OCR avec Tesseract :

```python
from PIL import Image
import pytesseract
import pdfplumber
# Tesseract installé dans C:\Program Files\Tesseract-OCR
# Si non ajouté au PATH, décommenter la ligne suivante :
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
with pdfplumber.open(r"C:\chemin\vers\scan.pdf") as pdf:
    img = pdf.pages[0].to_image(resolution=300).original
    text = pytesseract.image_to_string(img, lang="fra")
    print(text)
```

---

## DOCX / DOC

Pour un aperçu rapide :

```python
from docx import Document
doc = Document(r"C:\chemin\vers\memo.docx")
for i, para in enumerate(doc.paragraphs[:20]):
    print(para.text)
```

Pour plus d'opérations (création, modification, styles) → voir le skill `docx`.

**Legacy `.doc`** : `python-docx` ne lit pas le format `.doc`. Options :
- Demander à l'utilisateur de convertir en `.docx`
- Utiliser `antiword` si installé (rare sur Windows)

---

## XLSX / XLS / Tableurs

Aperçu rapide d'un `.xlsx` / `.xlsm` :

```python
from openpyxl import load_workbook
wb = load_workbook(r"C:\chemin\vers\data.xlsx", read_only=True)
print("Feuilles:", wb.sheetnames)
ws = wb.active
for row in ws.iter_rows(max_row=5, values_only=True):
    print(row)
wb.close()
```

**Important :** `read_only=True` est obligatoire pour les gros fichiers. Ne pas faire confiance à `ws.max_row` en mode read-only (souvent `None` ou faux). Si besoin du nombre de lignes, utiliser pandas.

**Legacy `.xls`** — openpyxl lève `InvalidFileException`. Utiliser :

```python
import pandas as pd
df = pd.read_excel(r"C:\chemin\vers\old.xls", engine="xlrd", nrows=5)
print(df)
```

**`.ods` (OpenDocument)** — openpyxl le refuse aussi. Utiliser :

```python
import pandas as pd
df = pd.read_excel(r"C:\chemin\vers\data.ods", engine="odf", nrows=5)
print(df)
```

Pour plus d'opérations → voir le skill `xlsx`.

---

## CSV / TSV

**Ne pas** lire un CSV brut avec `read_file` si le fichier est gros ou contient des cellules avec retours à la ligne. Utiliser pandas avec `nrows` :

```python
import pandas as pd
df = pd.read_csv(r"C:\chemin\vers\data.csv", nrows=5)
print(df)
print(df.dtypes)
```

Nombre approximatif de lignes (rapide) :

```powershell
(Get-Content "chemin\data.csv" | Measure-Object -Line).Lines
```

TSV : même chose avec `sep="\t"`.

---

## JSON / JSONL

Structure d'abord, contenu ensuite :

```python
import json
with open(r"C:\chemin\vers\data.json", encoding="utf-8") as f:
    data = json.load(f)
print(type(data).__name__)
if isinstance(data, list):
    print(f"{len(data)} éléments")
    print(json.dumps(data[:3], indent=2, ensure_ascii=False))
elif isinstance(data, dict):
    print(f"Clés: {list(data.keys())}")
```

JSONL (un objet par ligne) — ne pas tout charger :

```python
with open(r"C:\chemin\vers\data.jsonl", encoding="utf-8") as f:
    for i, line in enumerate(f):
        if i >= 3:
            break
        print(json.loads(line))
```

---

## Images (JPG / PNG / GIF / WEBP)

Utiliser l'outil natif `view_image` de Copilot — il affiche l'image directement.

Pour un traitement programmatique :

```python
from PIL import Image
img = Image.open(r"C:\chemin\vers\photo.jpg")
print(img.size, img.mode, img.format)
```

Pour OCR sur une image :

```python
import pytesseract
from PIL import Image
img = Image.open(r"C:\chemin\vers\photo.jpg")
text = pytesseract.image_to_string(img, lang="fra")
print(text)
```

---

## Archives (ZIP / TAR)

**Lister d'abord. Ne jamais extraire automatiquement** sauf demande explicite.

```python
import zipfile
with zipfile.ZipFile(r"C:\chemin\vers\bundle.zip") as z:
    for info in z.infolist():
        print(f"{info.filename}  ({info.file_size} bytes)")
```

Pour extraire un seul fichier :

```python
with zipfile.ZipFile(r"C:\chemin\vers\bundle.zip") as z:
    content = z.read("chemin/dans/archive/fichier.txt")
    print(content.decode("utf-8"))
```

TAR :

```python
import tarfile
with tarfile.open(r"C:\chemin\vers\bundle.tar.gz") as t:
    t.list()
```

---

## Texte / Code / Logs

Utiliser `read_file` (outil natif Copilot) directement. Pour les gros fichiers (> 500 lignes), lire par tranches :

- **Fichiers courts (< 500 lignes)** : `read_file` du début à la fin.
- **Fichiers longs** : lire les 100 premières et 100 dernières lignes pour orienter, puis cibler.
- **Logs** : l'utilisateur s'intéresse presque toujours à la fin du fichier.

---

## Extension inconnue

```python
with open(r"C:\chemin\vers\mystery.bin", "rb") as f:
    header = f.read(32)
    print(header.hex())
    print(header)
```

Si les magic bytes ne correspondent à rien de connu, demander à l'utilisateur.
