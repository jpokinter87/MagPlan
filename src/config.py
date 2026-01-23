"""
Configuration et constantes du générateur de planning
"""

import os
import sys

# Racine du projet
# - En mode compilé (PyInstaller) : dossier de l'exécutable
# - En mode script : parent du dossier src
if getattr(sys, 'frozen', False):
    # Exécutable PyInstaller
    PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    # Script Python
    PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Chemins absolus basés sur la racine du projet
DOSSIER_DATA = os.path.join(PROJECT_ROOT, "data")
DOSSIER_SORTIES = os.path.join(PROJECT_ROOT, "sorties")
FICHIER_REPARTITION_DEFAUT = "tableau_repartition_audiences.xlsx"

# Couleurs de base
GRIS = 'FFBFBFBF'
BLANC = 'FFFFFFFF'
ROUGE = 'FFFF0000'
JAUNE_CLAIR = 'FFFFFF99'
BLEU_FONCE = 'FF2E5090'  # Bleu foncé pour Débats JLD (moins soutenu que la nuit)

# Couleurs des légendes (colonnes A-B)
COULEUR_LEGENDE_PERMANENCES = 'FFF4B084'
COULEUR_LEGENDE_PERM_NUIT = 'FF4472C4'
COULEUR_LEGENDE_DEBATS_JLD = 'FF2E5090'  # Nouvelle couleur pour Débats JLD
COULEUR_LEGENDE_MATIN = 'FFBDD7EE'
COULEUR_LEGENDE_APRES_MIDI = 'FFFCE4D6'
COULEUR_LEGENDE_CIVIL = 'FFFF00FF'
COULEUR_LEGENDE_EP = 'FFED7D31'
COULEUR_LEGENDE_ECOFI = 'FF00B0F0'
COULEUR_LEGENDE_CRIMINEL = 'FFB80047'  # Bordeaux pour assises/CCD

# Mapping des codes sections vers couleurs
CODES_COULEURS = {
    'A': 'FF7030A0',      # PAP
    'M': 'FF00FF00',      # PMF
    'V': 'FFFF9999',      # PRA PMF
    'E': 'FF66FFFF',      # ECOFI
    'S': 'FFFFFF00',      # STUPS
    'J': 'FFFFFF00',      # PRA STUPS
    'P': 'FFFF9933',      # EP
    'N': 'FF009999',      # PCC
    'C': 'FFFF00FF',      # CIVIL
    'SG': 'FF0070C0',     # MAG SG
    'PR': 'FFFF0000',     # PR
    'PRA': 'FFFF0000'     # PRA
}

# Noms des mois en français
MOIS_FR = [
    "JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN",
    "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"
]

# Structure des permanences
PERMANENCES = [
    {'libelle': 'Permanence hiérarchique', 'couleur': 'FFFF0000'},
    {'libelle': 'Permanence PAP secteur A', 'couleur': 'FF9999FF'},
    {'libelle': 'Permanence PAP secteur B', 'couleur': 'FF9999FF'},
    {'libelle': 'Permanence préliminaire PAP', 'couleur': 'FF9966CC'},
    {'libelle': 'Permanence PMF', 'couleur': 'FF00FF00'},
    {'libelle': 'Permanence préliminaire PMF', 'couleur': 'FF00AE00'},
    {'libelle': 'Permanence Stups-crim.org.', 'couleur': 'FFFFFF00'},
    {'libelle': 'Permanence pôle Eco-Fi.', 'couleur': 'FF00FFFF'},
    {'libelle': 'Permanence pôle éxecution des peines', 'couleur': 'FFFFC000'},
    {'libelle': 'Permanence pôle civil', 'couleur': 'FFFF00FF'}
]

# Définition des audiences par zone
AUDIENCES_MATIN = [
    {'horaire': '9h', 'type': '11 JU', 'nom_planning': '11ème JU - Route'},
    {'horaire': '9h', 'type': '11 Coll', 'nom_planning': '11ème Coll - Route'},
    {'horaire': '9h', 'type': '12 JU', 'nom_planning': '12ème JU'},
    {'horaire': '9h', 'type': '18 JU', 'nom_planning': '18ème JU'},
    {'horaire': '9h', 'type': '20 JU', 'nom_planning': '20ème JU'},
    {'horaire': '9h', 'type': '21 Coll', 'nom_planning': '21ème Coll'},
    {'horaire': '9h', 'type': 'CRPC', 'nom_planning': 'CRPC'},
    {'horaire': '9h', 'type': 'TPE', 'nom_planning': 'TPE'}
]

AUDIENCES_APRES_MIDI = [
    {'horaire': '13h30', 'type': '12 Coll', 'nom_planning': '12ème Coll'},
    {'horaire': '13h30', 'type': '14 JU', 'nom_planning': '14ème JU'},
    {'horaire': '13h30', 'type': '14 Coll', 'nom_planning': '14ème Coll'},
    {'horaire': '13h30', 'type': '15 JU', 'nom_planning': '15ème JU'},
    {'horaire': '13h30', 'type': '15 Coll', 'nom_planning': '15ème Coll'},
    {'horaire': '13h30', 'type': '16 – CI', 'nom_planning': '16ème – CI'},
    {'horaire': '13h30', 'type': '17 Coll', 'nom_planning': '17ème Coll'},
    {'horaire': '13h30', 'type': '18 Coll', 'nom_planning': '18ème Coll'},
    {'horaire': '13h30', 'type': '20 Coll', 'nom_planning': '20ème Coll'},
    {'horaire': '13h30', 'type': '21 Coll', 'nom_planning': '21ème Coll'},
    {'horaire': '13h30', 'type': 'T.Police', 'nom_planning': 'T.Police'},
    {'horaire': '13h30', 'type': 'TPE', 'nom_planning': 'TPE'}
]

AUDIENCES_CIVILES = [
    {'horaire': '9h30', 'type': 'Aud. Civ.', 'nom_planning': 'Aud. Civ.'},
    {'horaire': '13h30', 'type': 'Aud. Civ.', 'nom_planning': 'Aud. Civ.'}
]

AUDIENCES_EP = [
    {'horaire': '9h', 'type': 'CAP', 'nom_planning': 'CAP'},
    {'horaire': '13h30', 'type': 'CAP', 'nom_planning': 'CAP'},
    {'horaire': '13h30', 'type': 'DCMO', 'nom_planning': 'DCMO'},
    {'horaire': '14h', 'type': 'DCMF', 'nom_planning': 'DCMF'},
    {'horaire': '13h30', 'type': 'CJ', 'nom_planning': 'CJ'}
]

AUDIENCES_ECOFI = [
    {'horaire': '9h', 'type': 'TJ', 'nom_planning': 'TJ'},
    {'horaire': '9h', 'type': 'T Com', 'nom_planning': 'T Com'},
    {'horaire': '13h30', 'type': 'T Com', 'nom_planning': 'T Com'}
]