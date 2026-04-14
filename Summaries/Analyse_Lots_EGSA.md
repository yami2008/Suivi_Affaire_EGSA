# Analyse de l'Appel d'Offres Informatique — EGSA
> Document de contre-vérification — Chef d'équipe développement

---

## Vue d'ensemble

Le document soumis est un **cahier des charges techniques** pour un appel d'offres informatique, découpé en **10 lots**. Il a été préparé par une collègue et couvre l'acquisition d'équipements informatiques et réseau pour l'entreprise EGSA.

---

## Tableau de synthèse

| Lot | Intitulé | État | Priorité |
|-----|----------|------|----------|
| 01 | PC de bureau (bureautique) | ✅ Acceptable | 🟢 Faible |
| 02 | PC de bureau (architectes) | 🔴 VIDE | 🔴 Urgent |
| 03 | Ordinateurs portables | 🟡 Incomplet | 🟡 Moyen |
| 04 | Mini PC (client léger) | 🔴 Contradictoire | 🔴 Urgent |
| 05 | Imprimantes multifonctions | ✅ Bon | 🟢 Faible |
| 06 | Scanners de documents | 🟡 Incomplet | 🟡 Moyen |
| 07 | Imprimantes badgeuse | 🟡 Acceptable | 🟡 Moyen |
| 08 | Pare-feux (Firewall) | 🟡 Incomplet | 🔴 Urgent |
| 09 | Switches réseau | 🟡 Incomplet | 🟡 Moyen |
| 10 | Serveurs informatiques | 🔴 VIDE | 🔴 Urgent |

---

## Analyse par lot

### Lot 01 — PC de bureau (bureautique) ✅
**Verdict : Acceptable**

Specs cohérentes pour un usage bureautique pur (Office, email, navigation web).

Points validés :
- Intel Core i5 12e gen ✅
- SSD + HDD : bonne combinaison ✅
- Windows 11 Pro avec licence ✅

Points à clarifier :
- Le **DVD+RW** est anachronique — à conserver uniquement si justifié
- Le port **VGA** est obsolète en 2025 — vérifier si encore utilisé dans les locaux

---

### Lot 02 — PC de bureau pour architectes 🔴
**Verdict : CRITIQUE — Lot entièrement vide**

C'est le lot le plus important pour l'équipe technique. Les logiciels de simulation 3D et AutoCAD exigent des machines puissantes.

Specs minimales recommandées :
- **CPU** : Intel Core i7 13e/14e gen minimum (ou AMD Ryzen 7/9)
- **RAM** : 32 Go DDR5 minimum
- **Stockage** : SSD NVMe 1 To système + HDD 2 To données
- **GPU** : Carte graphique dédiée obligatoire — NVIDIA RTX 3060 minimum ou NVIDIA Quadro
- **Écran** : 27" minimum, résolution QHD (2560x1440) ou 4K, dalle IPS

Questions à poser :
- Combien de postes architectes sont prévus ?
- Quels logiciels exactement ? (AutoCAD, Revit, 3ds Max, Lumion, V-Ray ?)
- Rendu 3D intensif ou dessin 2D/3D léger ?
- Pourquoi ce lot est-il vide ?

---

### Lot 03 — Ordinateurs portables 🟡
**Verdict : Acceptable mais incomplet**

Base correcte, mais plusieurs manques importants.

Points à améliorer :
- **2 ports USB minimum** : beaucoup trop peu — recommander 4 ports dont 1 USB-C
- **Pas de mention de l'autonomie** : exiger 6-8 heures minimum
- **Pas de type de dalle** : préciser IPS pour un usage professionnel
- **Pas de USB-C / Thunderbolt** : devenu standard en 2025

Questions à poser :
- Quel profil d'utilisateur ? (dev, commercial, manager)
- Usage principalement au bureau ou en mobilité ?

---

### Lot 04 — Mini PC (client léger) 🔴
**Verdict : Contradictoire — À clarifier en urgence**

C'est le lot avec le plus de contradictions internes.

Problèmes identifiés :
- **Contradiction majeure** : la remarque n°1 dit que la carte Wi-Fi et l'OS ne sont pas nécessaires pour le télé-affichage, mais ils figurent quand même dans les specs
- **10e génération Intel** : tous les autres lots exigent la 12e gen — incohérence
- **256 Go SSD** : insuffisant si les machines sont réaffectées à d'autres usages

Questions à poser :
- Wi-Fi USB et licence Windows : inclus ou pas ? Réponse ferme requise avant l'AO
- Pourquoi la 10e génération alors que les autres lots exigent la 12e ?
- Ces mini PC seront-ils exclusivement pour le télé-affichage ?

---

### Lot 05 — Imprimantes multifonctions ✅
**Verdict : Un des lots les mieux rédigés**

Points forts :
- 2 toners inclus (démarrage + supplémentaire) : excellente pratique
- Vitesse 40 ppm et résolution 1200x1200 dpi : très correct
- Triple connectivité (Ethernet + Wi-Fi + USB) : complète
- ADF + scanner à plat : bien pensé

Points à clarifier :
- **Monochrome uniquement** : suffisant ? Les architectes n'ont pas besoin de couleur ?
- **Bac de 150 feuilles** : un peu juste pour un usage partagé intensif — recommander 250-500 feuilles
- **Le fax en 2025** : encore utilisé chez EGSA ?
- **Volume mensuel d'impression** non précisé

---

### Lot 06 — Scanners de documents 🟡
**Verdict : Acceptable avec des lacunes**

Points à améliorer :
- **25 ppm** : trop lent par rapport aux 40 ppm de l'imprimante du lot 05
- **Format A4 uniquement** : les architectes ont probablement besoin du A3
- **Erreur de copier-coller** : "Format d'impression A4" écrit dans les specs d'un scanner — à corriger
- **Aucun consommable mentionné** : les rouleaux ADF s'usent, prévoir un kit de maintenance

Questions à poser :
- Besoin de scanner des plans ou documents grand format (A3+) ?
- 600 dpi suffisent pour les documents techniques ?

---

### Lot 07 — Imprimantes badgeuse 🟡
**Verdict : Acceptable — À déléguer à un expert**

Base correcte. Points à vérifier par un expert :
- **Consommables (rubans)** non mentionnés — peuvent être très coûteux selon la marque
- **Pas de volume horaire précisé** : vérifier que la capacité correspond aux besoins d'EGSA

---

### Lot 08 — Pare-feux (Firewall) 🔴
**Verdict : Incomplet — À déléguer à l'expert réseau/sécurité en urgence**

Contexte bien expliqué (remplacement des FortiGate en fin de vie), mais :
- **Débits fibre et ligne spécialisée toujours en blanc** : bloquant pour le dimensionnement
- **Licence d'un an seulement** : coût récurrent à anticiper — un contrat 3 ans serait plus économique
- **Aucun critère de performance précisé** : débit VPN garanti, connexions simultanées

---

### Lot 09 — Switches réseau 🟡
**Verdict : Correct mais incomplet — À déléguer à l'expert réseau**

Points à vérifier par un expert :
- **Gestion des VLANs** non mentionnée — essentielle en entreprise
- **Pas de redondance d'alimentation** : risque de coupure de service
- **Pas de mention de garantie matérielle**

---

### Lot 10 — Serveurs informatiques 🔴
**Verdict : CRITIQUE — Lot entièrement vide**

Même problème que le lot 02. Les serveurs sont le cœur de l'infrastructure EGSA.

Questions à poser (à l'expert infrastructure) :
- Peut-on lancer l'AO sans les specs serveurs, ou faut-il un AO séparé ?
- Quelles applications seront hébergées sur ces serveurs ?

---

## Problèmes transversaux

1. **Aucune quantité précisée** dans aucun lot — impossible de budgétiser
2. **Deux lots complètement vides** (02 et 10) — ne peuvent pas figurer dans un AO officiel tel quel
3. **Erreur de copier-coller** dans le lot 06 — manque de relecture
4. **Incohérence de génération CPU** entre le lot 04 (10e gen) et les autres (12e gen)
5. **Informations manquantes** dans le lot 08 (débits réseau) — à compléter avant publication

---

## Recommandations finales

| Action | Responsable | Urgence |
|--------|-------------|---------|
| Compléter les specs du lot 02 (architectes) | Chef d'équipe dev + collègue | 🔴 Immédiat |
| Compléter les specs du lot 10 (serveurs) | Expert infrastructure | 🔴 Immédiat |
| Clarifier la contradiction du lot 04 | Collègue | 🔴 Immédiat |
| Renseigner les débits manquants du lot 08 | Expert réseau | 🔴 Immédiat |
| Revoir les lots 07, 08, 09 en détail | Expert réseau/sécurité | 🟡 Avant publication |
| Corriger l'erreur du lot 06 | Collègue | 🟡 Avant publication |
| Préciser les quantités pour tous les lots | Collègue | 🟡 Avant publication |

---

*Document établi suite à une contre-vérification technique — EGSA, Chef d'équipe développement*
