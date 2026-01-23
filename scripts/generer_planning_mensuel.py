#!/usr/bin/env python3
"""
Script de génération automatique du planning mensuel des audiences de magistrats
Version 3.1 - Architecture refactorisée

Usage: python generer_planning_mensuel.py <mois> <année>
"""

import sys
import os

# Ajouter la racine du projet au path (parent du dossier scripts)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

from src.config import DOSSIER_DATA, DOSSIER_SORTIES, MOIS_FR
from src.planning import (
    creer_dossiers,
    trouver_fichier_repartition,
    generer_nom_fichier,
    creer_planning_mensuel
)


def main():
    creer_dossiers()

    if len(sys.argv) < 3:
        print("Usage: python generer_planning_mensuel.py <mois> <année>")
        print("\nExemples:")
        print("  python generer_planning_mensuel.py 2 2025")
        print("  python generer_planning_mensuel.py 11 2025")
        print("\nLe planning sera généré au format: MM - MOIS ANNEE.xlsx")
        print("Exemple: 11 - NOVEMBRE 2025.xlsx")
        print(f"\nLe tableau de répartition doit être dans le dossier '{DOSSIER_DATA}/'")
        print(f"Les plannings seront générés dans le dossier '{DOSSIER_SORTIES}/'")
        sys.exit(1)

    mois = int(sys.argv[1])
    annee = int(sys.argv[2])

    if mois < 1 or mois > 12:
        print(f"Erreur: Le mois doit être entre 1 et 12 (vous avez saisi: {mois})")
        sys.exit(1)

    fichier_repartition = trouver_fichier_repartition()

    if not fichier_repartition:
        print("Erreur: Fichier de répartition non trouvé !")
        print(f"\nVeuillez placer le fichier 'tableau_repartition_audiences.xlsx'")
        print(f"dans le dossier '{DOSSIER_DATA}/'")
        sys.exit(1)

    print(f"Tableau de répartition: {fichier_repartition}")

    nom_fichier = generer_nom_fichier(mois, annee)
    fichier_sortie = os.path.join(DOSSIER_SORTIES, nom_fichier)

    creer_planning_mensuel(mois, annee, fichier_repartition, fichier_sortie)

    print(f"\nFichier disponible dans: {fichier_sortie}")


if __name__ == "__main__":
    main()
