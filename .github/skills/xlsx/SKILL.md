---
name: xlsx
description: Use this skill any time a spreadsheet file is the primary input or output. This means any task where the user wants to open, read, edit, or fix an existing .xlsx, .xlsm, .xls, .ods, .csv, or .tsv file; create a new spreadsheet from scratch or from other data sources; convert between tabular file formats; or clean and restructure messy tabular data. Also trigger when the user references a spreadsheet file by name or path. Do NOT trigger when the primary deliverable is a Word document, PDF, HTML report, or script unrelated to spreadsheet generation.
---

# XLSX — Création, édition et analyse de tableurs

## Overview

Utilise ce skill quand un fichier tableur (.xlsx, .xlsm, .csv, .tsv, .xls, .ods) est l'entrée ou la sortie principale. Cela inclut :
- Ouvrir, lire, éditer ou corriger un fichier existant
- Créer un nouveau tableur depuis zéro ou depuis d'autres données
- Convertir entre formats tabulaires
- Nettoyer ou restructurer des données tabulaires mal formatées

**Ne PAS déclencher** quand le livrable principal est un document Word, un rapport HTML, un script Python ou une intégration API.

## Environnement

- **OS :** Windows
- **Librairies :** openpyxl, pandas, xlrd (pour .xls), odfpy (pour .ods)
- **Recalcul des formules :** Pas de LibreOffice CLI — l'utilisateur ouvre le fichier dans Excel pour recalculer, ou on vérifie les formules en Python

---

# Exigences pour les livrables

## Tous les fichiers Excel

### Police professionnelle
- Utiliser une police cohérente et professionnelle (ex : Arial, Times New Roman) sauf instruction contraire

### Zéro erreur de formule
- Tout fichier Excel DOIT être livré avec ZÉRO erreur de formule (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?)

### Préserver les templates existants
- Étudier et reproduire EXACTEMENT le format, style et conventions d'un fichier existant lors de modifications
- Ne jamais imposer un formatage standardisé sur un fichier avec des conventions établies
- Les conventions du template existant ont TOUJOURS priorité sur ces guidelines

---

## Modèles financiers

### Conventions de couleur (standard industrie)

Sauf indication contraire de l'utilisateur ou du template existant :

- **Texte bleu (RGB: 0,0,255)** : Inputs codés en dur, valeurs modifiables pour scénarios
- **Texte noir (RGB: 0,0,0)** : TOUTES les formules et calculs
- **Texte vert (RGB: 0,128,0)** : Liens vers d'autres feuilles du même classeur
- **Texte rouge (RGB: 255,0,0)** : Liens externes vers d'autres fichiers
- **Fond jaune (RGB: 255,255,0)** : Hypothèses clés nécessitant attention ou cellules à mettre à jour

### Formats de nombres

- **Années** : Format texte ("2024" et non "2,024")
- **Monnaie** : Format $#,##0 ; TOUJOURS préciser les unités dans les en-têtes ("Revenue ($mm)")
- **Zéros** : Afficher comme "-" via le format nombre (ex : `$#,##0;($#,##0);"-"`)
- **Pourcentages** : Par défaut 0.0% (une décimale)
- **Multiples** : Format 0.0x pour les multiples de valorisation (EV/EBITDA, P/E)
- **Nombres négatifs** : Utiliser les parenthèses (123) et non le moins -123

---

# CRITIQUE : Utiliser des formules, pas des valeurs codées en dur

**Toujours utiliser des formules Excel au lieu de calculer en Python et coder en dur.**

### INCORRECT — Valeurs codées en dur
```python
# Mauvais : calculer en Python et coder le résultat
total = df['Sales'].sum()
sheet['B10'] = total  # Code en dur 5000

# Mauvais : taux de croissance calculé en Python
growth = (df.iloc[-1]['Revenue'] - df.iloc[0]['Revenue']) / df.iloc[0]['Revenue']
sheet['C5'] = growth  # Code en dur 0.15
```

### CORRECT — Formules Excel
```python
# Bon : laisser Excel calculer
sheet['B10'] = '=SUM(B2:B9)'

# Bon : taux de croissance comme formule Excel
sheet['C5'] = '=(C4-C2)/C2'

# Bon : moyenne via fonction Excel
sheet['D20'] = '=AVERAGE(D2:D19)'
```

Cela s'applique à TOUS les calculs — totaux, pourcentages, ratios, différences, etc.

### Règles de construction des formules

#### Placement des hypothèses
- Placer TOUTES les hypothèses (taux de croissance, marges, multiples, etc.) dans des cellules dédiées
- Utiliser des références de cellules au lieu de valeurs codées en dur
- Exemple : `=B5*(1+$B$6)` au lieu de `=B5*1.05`

#### Prévention des erreurs
- Vérifier que toutes les références de cellules sont correctes
- Chercher les erreurs off-by-one dans les plages
- Assurer la cohérence des formules sur toutes les périodes de projection
- Tester avec des cas limites (zéros, négatifs)
- Vérifier l'absence de références circulaires non voulues

---

# Lecture et analyse de données

## Avec pandas

```python
import pandas as pd

# Lire Excel
df = pd.read_excel(r"C:\chemin\vers\fichier.xlsx")               # Première feuille par défaut
all_sheets = pd.read_excel(r"C:\chemin\vers\fichier.xlsx", sheet_name=None)  # Toutes les feuilles (dict)

# Analyser
df.head()       # Aperçu
df.info()       # Info colonnes
df.describe()   # Statistiques

# Écrire
df.to_excel(r"C:\chemin\vers\output.xlsx", index=False)
```

## Avec openpyxl (aperçu rapide)

```python
from openpyxl import load_workbook

wb = load_workbook(r"C:\chemin\vers\fichier.xlsx", read_only=True)
print("Feuilles:", wb.sheetnames)
ws = wb.active
for row in ws.iter_rows(max_row=5, values_only=True):
    print(row)
wb.close()
```

---

# Création de fichiers Excel

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

wb = Workbook()
sheet = wb.active

# Données
sheet['A1'] = 'Colonne A'
sheet['B1'] = 'Colonne B'
sheet.append(['Ligne', 'de', 'données'])

# Formule
sheet['B2'] = '=SUM(A1:A10)'

# Formatage
sheet['A1'].font = Font(bold=True, color='FF0000')
sheet['A1'].fill = PatternFill('solid', fgColor='FFFF00')
sheet['A1'].alignment = Alignment(horizontal='center')

# Largeur de colonne
sheet.column_dimensions['A'].width = 20

wb.save(r"C:\chemin\vers\output.xlsx")
```

---

# Édition de fichiers existants

```python
from openpyxl import load_workbook

wb = load_workbook(r"C:\chemin\vers\existant.xlsx")
sheet = wb.active  # ou wb['NomFeuille']

# Parcourir les feuilles
for sheet_name in wb.sheetnames:
    sheet = wb[sheet_name]
    print(f"Feuille: {sheet_name}")

# Modifier
sheet['A1'] = 'Nouvelle valeur'
sheet.insert_rows(2)
sheet.delete_cols(3)

# Ajouter une feuille
new_sheet = wb.create_sheet('NouvelleFeuille')
new_sheet['A1'] = 'Données'

wb.save(r"C:\chemin\vers\modifie.xlsx")
```

---

# Recalcul et vérification des formules

**Pas de LibreOffice CLI disponible.** Pour recalculer les formules :
- L'utilisateur ouvre le fichier dans Microsoft Excel → les formules sont recalculées automatiquement
- Sauvegarder depuis Excel pour persister les valeurs calculées

Pour **vérifier les formules en Python** avant livraison (détection d'erreurs potentielles) :

```python
from openpyxl import load_workbook

wb = load_workbook(r"C:\chemin\vers\output.xlsx")
errors_found = []
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                formula = cell.value
                if '#REF!' in formula:
                    errors_found.append(f"{sheet_name}!{cell.coordinate}: #REF! dans la formule")

if errors_found:
    print("Erreurs trouvées:")
    for e in errors_found:
        print(f"  - {e}")
else:
    print("Aucune erreur détectée dans les formules")
```

Pour lire les **valeurs calculées** (sans recalcul) :

```python
wb = load_workbook(r"C:\chemin\vers\fichier.xlsx", data_only=True)
# ATTENTION : si on sauvegarde après data_only=True, les formules sont PERDUES définitivement
```

---

# Formats legacy et alternatifs

## .xls (ancien format Excel)

openpyxl lève `InvalidFileException`. Utiliser xlrd via pandas :

```python
import pandas as pd
df = pd.read_excel(r"C:\chemin\vers\ancien.xls", engine="xlrd", nrows=5)
```

## .ods (OpenDocument)

openpyxl le refuse. Utiliser odfpy via pandas :

```python
import pandas as pd
df = pd.read_excel(r"C:\chemin\vers\data.ods", engine="odf", nrows=5)
```

---

# Checklist de vérification

### Vérifications essentielles
- [ ] **Tester 2-3 références** : vérifier qu'elles pointent vers les bonnes valeurs
- [ ] **Mapping des colonnes** : confirmer que les colonnes Excel correspondent (ex : colonne 64 = BL, pas BK)
- [ ] **Offset des lignes** : les lignes Excel sont 1-indexées (ligne DataFrame 5 = ligne Excel 6 avec en-tête)

### Pièges courants
- [ ] **Gestion des NaN** : vérifier les valeurs nulles avec `pd.notna()`
- [ ] **Colonnes éloignées** : les données FY sont souvent dans les colonnes 50+
- [ ] **Correspondances multiples** : chercher toutes les occurrences, pas juste la première
- [ ] **Division par zéro** : vérifier les dénominateurs avant `/` dans les formules
- [ ] **Références cassées** : vérifier que toutes les cellules référencées existent
- [ ] **Références inter-feuilles** : format correct `NomFeuille!A1`

### Stratégie de test
- [ ] **Commencer petit** : tester les formules sur 2-3 cellules avant d'appliquer largement
- [ ] **Vérifier les dépendances** : toutes les cellules référencées doivent exister
- [ ] **Cas limites** : inclure zéro, négatif et très grandes valeurs

---

# Choix de librairie

| Besoin | Outil |
|--------|-------|
| Analyse de données, opérations en masse, export simple | **pandas** |
| Formatage complexe, formules, fonctionnalités Excel spécifiques | **openpyxl** |
| Lecture de .xls legacy | **xlrd** (via pandas) |
| Lecture de .ods | **odfpy** (via pandas) |

### Rappels openpyxl
- Indices 1-based (row=1, column=1 = cellule A1)
- `data_only=True` pour lire les valeurs calculées — **ne jamais sauvegarder après** (perte des formules)
- `read_only=True` pour les gros fichiers en lecture
- `write_only=True` pour les gros fichiers en écriture
- Les formules sont préservées mais non évaluées par openpyxl

### Rappels pandas
- Spécifier les types pour éviter l'inférence : `dtype={'id': str}`
- Pour les gros fichiers, lire des colonnes spécifiques : `usecols=['A', 'C', 'E']`
- Gérer les dates : `parse_dates=['date_column']`

---

# Style de code

- Code Python minimal et concis, sans commentaires superflus
- Pas de noms de variables verbeux ni d'opérations redondantes
- Pas de `print()` inutiles
- **Dans les fichiers Excel** : ajouter des commentaires aux cellules avec formules complexes et documenter les sources des valeurs codées en dur
