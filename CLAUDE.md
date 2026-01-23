# CLAUDE.md

Ce fichier fournit des instructions à Claude Code (claude.ai/code) pour travailler avec ce dépôt.

## Présentation du projet

MagPlan est une application Python qui génère des fichiers Excel de planning mensuel pour les audiences de magistrats et les permanences. Elle lit les règles de répartition des audiences depuis un fichier Excel d'entrée et produit des plannings mensuels formatés avec cellules colorées, formules et gestion des jours fériés français.

## Commandes

### Exécuter l'application

```bash
# Générer le planning d'un mois (CLI)
python scripts/generer_planning_mensuel.py <mois> <annee>
# Exemple : python scripts/generer_planning_mensuel.py 2 2025

# Générer les 12 mois d'une année
python scripts/generer_annee_complete.py <annee>

# Lancer l'interface graphique (tkinter)
python scripts/generer_planning_gui.py
```

### Dépendances

```bash
pip install openpyxl
```

Le projet utilise `uv` pour la gestion des paquets (voir `.python-version` et `uv.lock`).

### Construire l'exécutable (PyInstaller)

```bash
pyinstaller MagPlan.spec
```

## Architecture

```
scripts/
├── generer_planning_gui.py      # Point d'entrée GUI (tkinter)
├── generer_planning_mensuel.py  # Point d'entrée CLI - génération d'un mois
└── generer_annee_complete.py    # Génération par lot pour une année complète

src/
├── config.py        # Constantes : couleurs (CODES_COULEURS), définitions des audiences, chemins
├── planning.py      # Orchestration principale : creer_planning_mensuel() coordonne tous les générateurs
├── repartition.py   # Charge les règles de répartition depuis Excel (charger_tableau_repartition)
├── audiences.py     # Génère les lignes d'audiences par catégorie (matin, après-midi, civil, EP, ECOFI, assises)
├── permanences.py   # Génère les lignes de permanences (permanences, débats JLD, nuit)
├── dates.py         # Calendrier des jours fériés français et utilitaires de dates (est_jour_ferie, calculer_numero_semaine)
└── styles.py        # Utilitaires de style Excel (couleurs, bordures, polices via openpyxl)

data/                # Entrée : tableau_repartition_audiences.xlsx
sorties/             # Sortie : fichiers planning générés (MM - MOIS AAAA.xlsx)
```

## Concepts clés

- **Répartition** : Règles de distribution chargées depuis `data/tableau_repartition_audiences.xlsx` définissant quelle section gère chaque type d'audience par jour de la semaine et numéro de semaine
- **Codes section** : Lettres simples (A, M, V, E, S, P, N, C, etc.) associées aux couleurs dans `CODES_COULEURS`
- **Numéro de semaine** : 1ère à 5ème occurrence d'un jour de la semaine dans le mois (pas la semaine ISO), calculé par `calculer_numero_semaine()`
- **Structure du planning** : Ligne 1 = dates, lignes 2-11 = permanences, ligne 12 = débats JLD, ligne 13 = permanence de nuit, puis blocs d'audiences par horaire/catégorie

## Conventions de code

- Tous les textes utilisateur en français
- Couleurs stockées en chaînes hexadécimales de 8 caractères avec préfixe alpha (ex: 'FFFF0000' pour rouge)
- La fonction `appliquer_style_cellule()` détecte automatiquement les fonds sombres et applique du texte blanc
- Les formules utilisent les noms de fonctions Excel en français dans les fichiers générés
