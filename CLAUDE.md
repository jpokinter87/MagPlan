# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

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

Python 3.13+ requis. Le projet utilise `uv` pour la gestion des paquets.

### Construire l'exécutable Windows (PyInstaller)

Le workflow GitHub Actions `.github/workflows/build-windows.yml` compile automatiquement les exécutables. Pour déclencher manuellement :

```bash
git tag v1.x.x && git push origin v1.x.x
```

## Architecture

Le point d'entrée principal est `src/planning.py:creer_planning_mensuel()` qui orchestre la génération :
1. Charge les règles depuis `repartition.py:charger_tableau_repartition()`
2. Génère les lignes de permanences via `permanences.py`
3. Génère les audiences via `audiences.py`
4. Applique les styles via `styles.py`

Flux de données : `data/tableau_repartition_audiences.xlsx` → `src/` → `sorties/MM - MOIS AAAA.xlsx`

## Concepts clés

- **Tableau de répartition** : Excel d'entrée avec colonnes par jour (5 cols/jour pour semaines 1-5), lignes par type d'audience. Chargé par `charger_tableau_repartition()` qui extrait les règles lignes 4-33
- **Codes section** : Lettres (A, M, V, E, S, P, N, C...) mappées vers couleurs dans `config.py:CODES_COULEURS`
- **Numéro de semaine** : Nième occurrence du jour dans le mois (pas ISO), via `dates.py:calculer_numero_semaine()`
- **Jours fériés français** : Fixes + mobiles (Pâques, Ascension, Pentecôte) calculés dans `dates.py:est_jour_ferie()`
- **Structure planning** : Ligne 1=dates, 2-11=permanences, 12=débats JLD, 13=nuit, puis audiences par catégorie

## Conventions de code

- Tous les textes utilisateur en français
- Couleurs : chaînes hex 8 caractères avec préfixe alpha (ex: `'FFFF0000'` pour rouge)
- `appliquer_style_cellule()` applique automatiquement du texte blanc sur fonds sombres
- Les formules Excel générées utilisent les noms de fonctions en français (IF→SI, etc.)
