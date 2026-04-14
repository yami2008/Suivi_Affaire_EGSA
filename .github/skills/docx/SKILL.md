---
name: docx
description: Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). Triggers include any mention of 'Word doc', 'document Word', '.docx', or requests to produce professional documents with formatting like headings, tables of contents, page numbers, headers/footers, or letterheads. Also use when extracting or reorganizing content from .docx files, performing find-and-replace, or converting content into a polished Word document. If the user asks for a report, memo, letter, courrier, avis, or template as a .docx file, use this skill. Do NOT use for PDFs, spreadsheets, or general coding tasks.
---

# DOCX — Création, édition et lecture de documents Word

## Overview

Utilise ce skill quand l'utilisateur veut créer, lire, éditer ou manipuler un fichier Word (.docx). Cela inclut :
- Créer un nouveau document (rapport, courrier, mémo, avis, lettre, template)
- Lire/extraire du contenu d'un .docx existant
- Modifier un document existant (ajouter/supprimer du contenu, reformater)
- Insérer ou remplacer des images
- Travailler avec des en-têtes, pieds de page, numéros de page
- Produire un document professionnel avec mise en forme

**Ne PAS utiliser** pour les PDF, tableurs, Google Docs ou tâches de code sans rapport avec la génération de documents.

## Environnement

- **OS :** Windows
- **Librairie principale :** python-docx
- **Librairies complémentaires :** pillow (images), reportlab (si conversion PDF nécessaire)
- **Pandoc :** optionnel, pour extraction texte rapide ou conversions de format

---

# Référence rapide

| Tâche | Approche |
|-------|----------|
| Lire/analyser le contenu | python-docx ou pandoc (si installé) |
| Créer un nouveau document | python-docx |
| Éditer un document existant | python-docx (charger → modifier → sauvegarder) |
| Convertir .doc legacy en .docx | Demander à l'utilisateur (pas de LibreOffice CLI) |

---

# Lecture de documents

## Lecture rapide

```python
from docx import Document

doc = Document(r"C:\chemin\vers\document.docx")

# Texte de tous les paragraphes
for para in doc.paragraphs:
    print(para.text)

# Texte des tableaux
for table in doc.tables:
    for row in table.rows:
        print([cell.text for cell in row.cells])
```

## Lecture avec styles et métadonnées

```python
from docx import Document

doc = Document(r"C:\chemin\vers\document.docx")

# Métadonnées
props = doc.core_properties
print(f"Titre: {props.title}")
print(f"Auteur: {props.author}")
print(f"Modifié: {props.modified}")

# Paragraphes avec leur style
for para in doc.paragraphs:
    if para.text.strip():
        print(f"[{para.style.name}] {para.text}")
```

## Lecture avec pandoc (si installé)

```powershell
pandoc document.docx -t markdown -o output.md
```

---

# Création de documents

## Structure de base

```python
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# Titre
doc.add_heading('Titre du document', level=0)

# Paragraphe
doc.add_paragraph('Texte du paragraphe.')

# Paragraphe avec formatage
para = doc.add_paragraph()
run = para.add_run('Texte en gras')
run.bold = True
run = para.add_run(' et ')
run = para.add_run('texte en italique')
run.italic = True

doc.save(r"C:\chemin\vers\output.docx")
```

## Police et taille par défaut

```python
from docx import Document
from docx.shared import Pt

doc = Document()
style = doc.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(12)

# Les paragraphes hériteront de cette police
doc.add_paragraph('Ce texte est en Arial 12pt.')
doc.save(r"C:\chemin\vers\output.docx")
```

## Titres et sous-titres

```python
doc.add_heading('Heading 1', level=1)
doc.add_heading('Heading 2', level=2)
doc.add_heading('Heading 3', level=3)
# level=0 → Titre principal (Title)
# level=1-9 → Headings
```

Pour personnaliser le style des headings :

```python
from docx.shared import Pt, RGBColor

style = doc.styles['Heading 1']
style.font.name = 'Arial'
style.font.size = Pt(16)
style.font.bold = True
style.font.color.rgb = RGBColor(0, 0, 0)
```

---

## Listes

```python
# Liste à puces
doc.add_paragraph('Premier élément', style='List Bullet')
doc.add_paragraph('Deuxième élément', style='List Bullet')

# Liste numérotée
doc.add_paragraph('Étape 1', style='List Number')
doc.add_paragraph('Étape 2', style='List Number')

# Sous-listes (niveaux imbriqués)
doc.add_paragraph('Élément principal', style='List Bullet')
doc.add_paragraph('Sous-élément', style='List Bullet 2')
doc.add_paragraph('Sous-sous-élément', style='List Bullet 3')
```

**Ne jamais** insérer manuellement des caractères bullet (•, \u2022). Toujours utiliser les styles de liste.

---

## Tableaux

```python
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

doc = Document()

# Créer un tableau
table = doc.add_table(rows=3, cols=3, style='Table Grid')
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# En-têtes
headers = ['Colonne A', 'Colonne B', 'Colonne C']
for i, header in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.text = header
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.bold = True

# Données
data = [
    ['Valeur 1', 'Valeur 2', 'Valeur 3'],
    ['Valeur 4', 'Valeur 5', 'Valeur 6'],
]
for row_idx, row_data in enumerate(data, start=1):
    for col_idx, value in enumerate(row_data):
        table.rows[row_idx].cells[col_idx].text = value

# Largeur des colonnes
for row in table.rows:
    row.cells[0].width = Cm(5)
    row.cells[1].width = Cm(5)
    row.cells[2].width = Cm(5)
```

### Ombrage des cellules (couleur de fond)

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_cell_shading(cell, color):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)

# Exemple : en-tête bleu clair
for cell in table.rows[0].cells:
    set_cell_shading(cell, 'D5E8F0')
```

---

## Images

```python
from docx.shared import Inches, Cm

# Image pleine largeur
doc.add_picture(r"C:\chemin\vers\image.png", width=Inches(6))

# Image avec dimensions spécifiques
doc.add_picture(r"C:\chemin\vers\logo.png", width=Cm(4), height=Cm(2))

# Image centrée
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
```

---

## Sauts de page

```python
from docx.enum.text import WD_BREAK

# Saut de page
doc.add_page_break()

# Ou dans un paragraphe existant
para = doc.add_paragraph()
run = para.add_run()
run.add_break(WD_BREAK.PAGE)
```

---

## En-têtes et pieds de page

```python
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# Accéder à la section
section = doc.sections[0]

# En-tête
header = section.header
header_para = header.paragraphs[0]
header_para.text = "Nom de l'entreprise — Document confidentiel"
header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in header_para.runs:
    run.font.size = Pt(9)
    run.font.italic = True

# Pied de page
footer = section.footer
footer_para = footer.paragraphs[0]
footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer_para.text = "Page "
```

### Numéros de page automatiques

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_page_number(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' PAGE '
    run._r.append(instrText)

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)

# Utilisation dans le footer
footer = doc.sections[0].footer
footer_para = footer.paragraphs[0]
footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer_para.add_run('Page ')
add_page_number(footer_para)
```

---

## Taille de page et marges

```python
from docx.shared import Cm, Inches
from docx.enum.section import WD_ORIENT

section = doc.sections[0]

# A4 (défaut)
section.page_width = Cm(21)
section.page_height = Cm(29.7)

# US Letter
section.page_width = Inches(8.5)
section.page_height = Inches(11)

# Marges
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)

# Orientation paysage
section.orientation = WD_ORIENT.LANDSCAPE
# IMPORTANT : en paysage, il faut aussi inverser width/height
section.page_width, section.page_height = section.page_height, section.page_width
```

---

## Liens hypertexte

python-docx ne supporte pas nativement les hyperliens. Il faut manipuler le XML :

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import docx.opc.constants

def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0563C1')
    rPr.append(c)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    new_run.append(rPr)

    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

# Utilisation
para = doc.add_paragraph('Visitez ')
add_hyperlink(para, 'notre site', 'https://example.com')
```

---

# Édition de documents existants

## Charger, modifier, sauvegarder

```python
from docx import Document

doc = Document(r"C:\chemin\vers\original.docx")

# Modifier un paragraphe existant
for para in doc.paragraphs:
    if 'ancien texte' in para.text:
        for run in para.runs:
            run.text = run.text.replace('ancien texte', 'nouveau texte')

# Ajouter du contenu à la fin
doc.add_paragraph('Paragraphe ajouté.')

# Sauvegarder (nouveau fichier pour préserver l'original)
doc.save(r"C:\chemin\vers\modifie.docx")
```

## Rechercher et remplacer

```python
def find_and_replace(doc, old_text, new_text):
    for para in doc.paragraphs:
        if old_text in para.text:
            for run in para.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if old_text in para.text:
                        for run in para.runs:
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_text)
```

**Attention :** python-docx découpe parfois le texte en plusieurs runs de manière imprévisible (ex : "Bonjour" peut devenir ["Bon", "jour"]). Si `find_and_replace` ne trouve pas le texte dans un seul run, il faut reconstituer le texte complet du paragraphe :

```python
def find_and_replace_across_runs(paragraph, old_text, new_text):
    full_text = paragraph.text
    if old_text not in full_text:
        return False
    new_full = full_text.replace(old_text, new_text)
    if paragraph.runs:
        first_run = paragraph.runs[0]
        first_run.text = new_full
        for run in paragraph.runs[1:]:
            run.text = ''
    return True
```

## Supprimer un paragraphe

```python
def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)

# Exemple : supprimer tous les paragraphes vides
for para in doc.paragraphs:
    if not para.text.strip():
        delete_paragraph(para)
```

## Insérer un paragraphe à un endroit précis

```python
from docx.oxml.ns import qn

def insert_paragraph_after(paragraph, text, style=None):
    new_p = paragraph._element.addnext(
        paragraph._element.makeelement(qn('w:p'), {})
    )
    new_para = type(paragraph)(new_p, paragraph._element.getparent())
    new_para.text = text
    if style:
        new_para.style = style
    return new_para
```

---

# Limitations de python-docx

| Fonctionnalité | Support |
|----------------|---------|
| Créer/éditer paragraphes, runs, styles | Complet |
| Tableaux (création, formatage) | Complet |
| Images | Complet |
| Headers/footers | Complet |
| Numéros de page | Via XML (voir ci-dessus) |
| Table des matières | Création du champ seulement — mise à jour par Word |
| Hyperliens | Via XML (voir ci-dessus) |
| Tracked changes (suivi des modifications) | Lecture partielle, pas de création |
| Commentaires | Non supporté nativement |
| Formulaires | Non supporté |
| Conversion .doc → .docx | Non supporté (demander à l'utilisateur) |

Pour les tracked changes et commentaires, il faut manipuler le XML sous-jacent via `doc.element` et `lxml`. À n'utiliser qu'en dernier recours.

---

# Conventions professionnelles

## Documents français (courriers, rapports, avis)

- **Police :** Arial ou Times New Roman, 12pt
- **Marges :** 2.5 cm sur tous les côtés
- **Interligne :** 1.15 ou 1.5
- **Alignement :** Justifié pour le corps du texte
- **En-tête :** Logo + nom de l'entreprise si applicable
- **Pied de page :** Numéro de page centré

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

doc = Document()

# Police par défaut
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(12)
style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
style.paragraph_format.line_spacing = 1.15

# Marges
section = doc.sections[0]
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)
```

## Template de courrier

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()
section = doc.sections[0]
section.top_margin = Cm(2)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)

# Expéditeur (aligné à gauche)
exp = doc.add_paragraph()
exp.add_run('Nom Entreprise\n').bold = True
exp.add_run('Adresse\nVille, Code postal')
exp.paragraph_format.space_after = Pt(24)

# Destinataire (aligné à droite)
dest = doc.add_paragraph()
dest.alignment = WD_ALIGN_PARAGRAPH.RIGHT
dest.add_run('Destinataire\nAdresse\nVille')
dest.paragraph_format.space_after = Pt(12)

# Date et lieu
date_para = doc.add_paragraph()
date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
date_para.add_run('Alger, le 14 avril 2026')
date_para.paragraph_format.space_after = Pt(24)

# Objet
objet = doc.add_paragraph()
objet.add_run('Objet : ').bold = True
objet.add_run("Description de l'objet du courrier")
objet.paragraph_format.space_after = Pt(24)

# Corps
doc.add_paragraph('Madame, Monsieur,')
doc.add_paragraph('Corps du courrier...')
doc.add_paragraph("Veuillez agréer, Madame, Monsieur, l'expression de nos salutations distinguées.")

# Signature
sig = doc.add_paragraph()
sig.paragraph_format.space_before = Pt(48)
sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
sig.add_run('Nom et Prénom\n').bold = True
sig.add_run('Fonction')

doc.save(r"C:\chemin\vers\courrier.docx")
```

---

# Style de code

- Code Python minimal et concis
- Pas de commentaires superflus ni de variables inutilement verbeux
- Toujours sauvegarder sous un nouveau nom pour préserver l'original
- Tester la lecture du fichier après écriture pour vérifier l'intégrité
