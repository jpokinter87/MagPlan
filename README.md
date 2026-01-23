# MagPlan - Générateur de Planning pour Magistrats

Application Python pour générer automatiquement des plannings mensuels d'audiences et de permanences pour les magistrats, au format Excel.

## Fonctionnalités

- **Génération de plannings mensuels** avec coloration automatique selon les règles de répartition
- **Génération annuelle** des 12 mois en une seule opération
- **Interface graphique** intuitive avec onglets (Planning Mensuel / Planning Annuel)
- **Interface en ligne de commande** pour automatisation
- **Gestion des jours fériés français** (fixes et mobiles : Pâques, Ascension, Pentecôte)
- **Formatage professionnel** : bordures, couleurs par zone, colonnes figées

## Installation

### Prérequis

- Python 3.8 ou supérieur
- pip (gestionnaire de paquets Python)

### Installation des dépendances

```bash
pip install openpyxl
```

### Configuration

1. Placez le fichier `tableau_repartition_audiences.xlsx` dans le dossier `data/`
2. Les plannings générés seront créés dans le dossier `sorties/`

```
MagPlan/
├── data/
│   └── tableau_repartition_audiences.xlsx  # Fichier de répartition (requis)
├── sorties/                                 # Plannings générés
├── src/                                     # Code source
└── generer_planning_gui.py                  # Interface graphique
```

## Utilisation

### Interface graphique (recommandé)

```bash
python generer_planning_gui.py
```

L'interface propose deux onglets :
- **Planning Mensuel** : sélection du mois et de l'année
- **Planning Annuel** : génération des 12 mois avec barre de progression

### Ligne de commande

```bash
# Générer un mois
python generer_planning_mensuel.py <mois> <année>
python generer_planning_mensuel.py 2 2026    # Février 2026

# Générer une année complète
python generer_annee_complete.py <année>
python generer_annee_complete.py 2026        # Les 12 mois de 2026
```

### Fichiers générés

Les plannings sont nommés au format : `MM - MOIS AAAA.xlsx`

Exemples :
- `01 - JANVIER 2026.xlsx`
- `02 - FEVRIER 2026.xlsx`

## Structure du planning généré

| Lignes | Contenu | Couleur légende |
|--------|---------|-----------------|
| 1 | Dates du mois | - |
| 2-11 | Permanences (hiérarchique, PAP, PMF, etc.) | Orange clair |
| 12 | Débats JLD | Bleu foncé |
| 13 | Permanence de Nuit | Bleu |
| 14-21 | Audiences du matin (9h) | Bleu clair |
| 22-33 | Audiences de l'après-midi (13h30) | Orange clair |
| 34-35 | Audiences civiles | Magenta |
| 36-40 | Exécution des peines | Orange |
| 41-43 | ECOFI | Bleu clair |
| 44-45 | Assises et CCD | Bordeaux |

## Tableau de répartition

Le fichier `tableau_repartition_audiences.xlsx` définit les règles d'attribution des audiences par :
- Jour de la semaine (lundi à vendredi)
- Numéro de semaine dans le mois (1ère à 5ème occurrence)
- Section (code couleur : A, M, V, E, S, P, N, C, etc.)

### Codes des sections

| Code | Section | Couleur |
|------|---------|---------|
| A | PAP | Violet |
| M | PMF | Vert |
| V | PRA PMF | Rose |
| E | ECOFI | Cyan |
| S | STUPS | Jaune |
| P | EP | Orange |
| N | PCC | Cyan foncé |
| C | CIVIL | Magenta |

## Compilation Windows

Le projet inclut un workflow GitHub Actions pour générer automatiquement un exécutable Windows 64 bits.

### Télécharger l'exécutable

1. Aller sur [GitHub Actions](../../actions)
2. Lancer le workflow "Build Windows Executable"
3. Télécharger l'artifact `MagPlan-Windows-x64`

### Créer une release

```bash
git tag v1.0.0
git push origin v1.0.0
```

Une release GitHub sera créée automatiquement avec les exécutables.

## Architecture du code

```
src/
├── config.py       # Constantes, couleurs, définitions des audiences
├── planning.py     # Orchestration de la génération
├── repartition.py  # Chargement des règles de répartition
├── audiences.py    # Génération des lignes d'audiences
├── permanences.py  # Génération des lignes de permanences
├── dates.py        # Jours fériés et calculs de dates
└── styles.py       # Formatage Excel (couleurs, bordures, polices)
```

## Dépannage

### "ModuleNotFoundError: No module named 'openpyxl'"

```bash
pip install openpyxl
```

### "Fichier de répartition non trouvé"

Vérifiez que `tableau_repartition_audiences.xlsx` est bien dans le dossier `data/`.

### Planning sans couleurs

- Vérifiez que le tableau de répartition contient les règles pour le mois demandé
- Vérifiez que les noms des audiences correspondent entre le script et le tableau

## Licence

Projet interne - Tous droits réservés.

## Version

**Version 3.1** - Janvier 2026
