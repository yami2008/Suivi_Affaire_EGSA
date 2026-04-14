# PRD — Workspace Suivi Affaires EGSA

**Rôle utilisateur :** Chef d'équipe développement et département  
**Problème résolu :** Trop de tâches simultanées, oublis de suivi, pas de traçabilité des tâches imprévues  
**Solution :** Un workflow VS Code + GitHub Copilot centré sur un fichier Excel unique comme source de vérité

---

## Comment on travaille

| Commande | Ce que je fais |
|----------|----------------|
| `Briefing` | Je lis le fichier Excel, je sors les urgences du jour, les alertes, les prochaines actions |
| `Clôture` | Résumé de la journée : fait / pas fait / reporté / pourquoi |
| `Tâche imprévue : [description]` | J'ajoute une ligne dans Excel avec `Imprévu = Oui`, historique daté |
| `Où on en est sur AFF-XXX ?` | Je lis la ligne et te résume l'historique + bloquants + prochaine action |
| `J'ai fait X sur AFF-XXX` | Je mets à jour historique, statut, date MAJ dans Excel |
| `Récap semaine / mois` | Je filtre par `Semaine N°` ou `Date Ouverture` et je résume |
| `[image/PDF]` | Je lis le document (OCR si scanné) et j'analyse |

---

## Fichiers du workspace

| Fichier | Rôle |
|---------|------|
| `Suivi_Affaires_EGSA.xlsx` | **Source de vérité unique** — toutes les affaires, historiques, alertes |
| `read_excel.py` | Script Python de mise à jour du fichier Excel (styles, hyperliens, nouvelles colonnes) |
| `SETUP.md` | État de l'installation (versions, statuts) |
| `INSTALL.md` | Manuel de déploiement pour reproduire l'environnement sur un autre PC |
| `.gitignore` | Exclut les dossiers `AFF-*/` du push GitHub (confidentiels + lourds) |
| `PRD.md` | Ce fichier — description du workflow |

---

## Structure des dossiers d'affaires

```
AFF-001_BIG_Informatique/     ← fichiers liés à l'affaire (contrats, courriers, PDFs...)
AFF-00X_Titre/                ← créé à la demande, lié par hyperlien dans Excel
```

Les dossiers `AFF-*/` ne sont **jamais poussés sur GitHub** (`.gitignore`).

---

## Colonnes du fichier Excel (23 colonnes)

| Colonne | Utilité |
|---------|---------|
| `ID` | Identifiant unique (AFF-001, AFF-002...) |
| `Titre` | Nom court de l'affaire |
| `Type` | Courrier, Réparation, Appel d'offres, Analyse, À la volée... |
| `Description` | Contexte de l'affaire |
| `Statut` | En cours / À traiter / En attente / Clôturée |
| `Priorité` | Haute / Moyenne / Basse |
| `Prochaine Action` | Ce qu'il faut faire ensuite |
| `Responsable` | Qui doit agir |
| `Date Ouverture` | Quand l'affaire a été ouverte |
| `Date Limite` | Deadline si applicable |
| `Date Dernière MAJ` | Dernière modification |
| `Historique` | Journal daté de tout ce qui s'est passé |
| `Observations` | Notes libres |
| `Date Clôture` | Quand l'affaire a été réglée |
| `Temps passé` | Ex : "2h", "demi-journée", "3j" |
| `Dossier` | Hyperlien vers le dossier de fichiers (cliquable dans Excel) |
| `Date Alerte` | Date à partir de laquelle une relance est nécessaire |
| `Origine` | Qui a initié (Direction, Client, Collègue, Imprévu...) |
| `Imprévu` | Oui / Non — permet d'expliquer les journées perturbées |
| `Date Prochaine Action` | Date cible pour agir — base du briefing quotidien |
| `Bloquant` | Ce qui empêche d'avancer |
| `Tags` | Mots-clés pour regrouper (CDC2026, BIG, ERP...) |
| `Semaine N°` | Semaine d'ouverture — pour les récaps hebdo/mensuel |

---

## Code couleur Excel

| Couleur | Signification |
|---------|--------------|
| Rouge | Alerte dépassée — action immédiate requise |
| Orange | Alerte dans moins de 3 jours |
| Gris | Affaire clôturée |
| Blanc | En cours, pas encore en alerte |

---

## Skills Copilot (`.github/skills/`)

Les skills sont des fichiers d'instructions que GitHub Copilot lit automatiquement avant d'agir sur certains types de fichiers.

| Skill | Fichier | Quand utilisé |
|-------|---------|---------------|
| `xlsx` | `xlsx/SKILL.md` | Dès qu'un fichier Excel est concerné — conventions couleurs, formules vs valeurs codées, structure |
| `docx` | `docx/SKILL.md` | Création ou modification d'un document Word (courriers, rapports, avis) |
| `file-reading` | `file-reading/SKILL.md` | Routeur — détermine quel outil Python utiliser selon le type de fichier (PDF, image, archive...) |

---

## Outils installés

| Outil | Usage |
|-------|-------|
| Python 3.13.3 | Moteur de tous les scripts |
| openpyxl, pandas | Lire/écrire Excel |
| python-docx | Créer/modifier Word |
| pdfplumber, pypdf | Lire PDF |
| pytesseract + Tesseract OCR | Lire texte sur images et PDF scannés (langue FR) |
| pillow | Traitement d'images |
| reportlab | Générer PDF |
| matplotlib | Graphiques |
| python-pptx | PowerPoint |
| jinja2 | Templates de documents en série |
| Pandoc | Conversions de formats |
