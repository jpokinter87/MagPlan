#!/usr/bin/env python3
"""
Script pour générer tous les plannings mensuels d'une année
Version 2.1 - Simplifié
Usage: python generer_annee_complete.py <annee>
"""

import sys
import subprocess
import os

MOIS_FR = [
    "JANVIER", "FÉVRIER", "MARS", "AVRIL", "MAI", "JUIN",
    "JUILLET", "AOÛT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DÉCEMBRE"
]

def generer_annee_complete(annee):
    """Génère les 12 plannings mensuels pour l'année donnée"""
    
    print(f"╔════════════════════════════════════════════════════════╗")
    print(f"║  Génération des plannings pour l'année {annee}          ║")
    print(f"╚════════════════════════════════════════════════════════╝\n")
    
    succes = 0
    echecs = 0
    fichiers_generes = []
    
    for mois in range(1, 13):
        print(f"[{mois:2d}/12] Génération de {MOIS_FR[mois-1]} {annee}...", end=" ")
        
        try:
            # Appeler le script principal
            cmd = [
                sys.executable,
                "generer_planning_mensuel.py",
                str(mois),
                str(annee)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✅")
                fichier = f"{mois:02d} - {MOIS_FR[mois-1]} {annee}.xlsx"
                fichiers_generes.append(fichier)
                succes += 1
            else:
                print("❌")
                print(f"  Erreur: {result.stderr}")
                echecs += 1
        except Exception as e:
            print("❌")
            print(f"  Exception: {e}")
            echecs += 1
    
    print("\n" + "=" * 60)
    print(f"✅ Résumé: {succes} plannings générés avec succès")
    if echecs > 0:
        print(f"❌ {echecs} échecs")
    print("=" * 60)
    
    if succes > 0:
        print(f"\n📂 Tous les plannings sont dans le dossier: sorties/")
        print(f"\nFichiers générés:")
        for fichier in fichiers_generes:
            print(f"  • {fichier}")


def main():
    if len(sys.argv) < 2:
        print("╔════════════════════════════════════════════════════════╗")
        print("║  Générateur de Plannings Annuels - Version 2.1        ║")
        print("╚════════════════════════════════════════════════════════╝")
        print("\nUsage: python generer_annee_complete.py <annee>")
        print("\nExemple:")
        print("  python generer_annee_complete.py 2025")
        print("\nCe script génère automatiquement les 12 plannings mensuels")
        print("au format: MM - MOIS ANNEE.xlsx")
        print("\nPrérequis:")
        print("  • Le fichier generer_planning_mensuel.py doit être présent")
        print("  • Le tableau de répartition doit être dans data/")
        sys.exit(1)
    
    annee = int(sys.argv[1])
    
    if not os.path.exists("generer_planning_mensuel_old.py"):
        print("❌ Erreur: Le fichier 'generer_planning_mensuel.py' est introuvable.")
        print("   Assurez-vous d'être dans le bon dossier.")
        sys.exit(1)
    
    generer_annee_complete(annee)
    
    print("\n🎉 Génération terminée !")


if __name__ == "__main__":
    main()
