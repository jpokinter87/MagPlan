#!/bin/bash
# Script pour générer facilement un planning mensuel
# Usage: ./generer_planning.sh <mois> <annee>

if [ $# -lt 2 ]; then
    echo "Usage: ./generer_planning.sh <mois> <annee>"
    echo "Exemple: ./generer_planning.sh 2 2025"
    echo ""
    echo "Ce script génère un planning pour le mois et l'année spécifiés."
    exit 1
fi

MOIS=$1
ANNEE=$2

# Nom de fichier avec zero padding
MOIS_PADDED=$(printf "%02d" $MOIS)
OUTPUT="planning_${MOIS_PADDED}_${ANNEE}.xlsx"

echo "Génération du planning pour ${MOIS}/${ANNEE}..."
python3 generer_planning_mensuel.py $MOIS $ANNEE tableau_répartition_audiences.xlsx $OUTPUT

if [ $? -eq 0 ]; then
    echo ""
    echo "Planning généré avec succès: $OUTPUT"
else
    echo ""
    echo "Erreur lors de la génération du planning."
fi
