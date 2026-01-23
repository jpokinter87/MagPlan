#!/usr/bin/env python3
"""
Script pour générer tous les plannings mensuels d'une année
Usage: python generer_annee_complete.py <annee>
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


def generer_annee_complete(annee):
    """Génère les 12 plannings mensuels pour l'année donnée"""

    print(f"{'=' * 60}")
    print(f"  Génération des plannings pour l'année {annee}")
    print(f"{'=' * 60}\n")

    creer_dossiers()

    fichier_repartition = trouver_fichier_repartition()
    if not fichier_repartition:
        print("Erreur: Fichier de répartition non trouvé !")
        print(f"\nVeuillez placer le fichier 'tableau_repartition_audiences.xlsx'")
        print(f"dans le dossier '{DOSSIER_DATA}/'")
        sys.exit(1)

    print(f"Tableau de répartition: {fichier_repartition}\n")

    succes = 0
    echecs = 0
    fichiers_generes = []

    for mois in range(1, 13):
        print(f"[{mois:2d}/12] Génération de {MOIS_FR[mois-1]} {annee}...", end=" ")

        try:
            nom_fichier = generer_nom_fichier(mois, annee)
            fichier_sortie = os.path.join(DOSSIER_SORTIES, nom_fichier)
            creer_planning_mensuel(mois, annee, fichier_repartition, fichier_sortie)
            print("OK")
            fichiers_generes.append(nom_fichier)
            succes += 1
        except Exception as e:
            print("ERREUR")
            print(f"  Exception: {e}")
            echecs += 1

    print("\n" + "=" * 60)
    print(f"Résumé: {succes} plannings générés avec succès")
    if echecs > 0:
        print(f"{echecs} échecs")
    print("=" * 60)

    if succes > 0:
        print(f"\nTous les plannings sont dans le dossier: {DOSSIER_SORTIES}/")
        print(f"\nFichiers générés:")
        for fichier in fichiers_generes:
            print(f"  - {fichier}")


def main():
    if len(sys.argv) < 2:
        print("=" * 60)
        print("  Générateur de Plannings Annuels")
        print("=" * 60)
        print("\nUsage: python generer_annee_complete.py <annee>")
        print("\nExemple:")
        print("  python generer_annee_complete.py 2025")
        print("\nCe script génère automatiquement les 12 plannings mensuels")
        print("au format: MM - MOIS ANNEE.xlsx")
        print("\nPrérequis:")
        print(f"  - Le tableau de répartition doit être dans {DOSSIER_DATA}/")
        sys.exit(1)

    annee = int(sys.argv[1])
    generer_annee_complete(annee)
    print("\nGénération terminée !")


if __name__ == "__main__":
    main()
