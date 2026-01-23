"""
Génération des lignes d'audiences
"""

from datetime import datetime

from .config import (
    AUDIENCES_MATIN, AUDIENCES_APRES_MIDI, AUDIENCES_CIVILES,
    AUDIENCES_EP, AUDIENCES_ECOFI,
    COULEUR_LEGENDE_MATIN, COULEUR_LEGENDE_APRES_MIDI,
    COULEUR_LEGENDE_CIVIL, COULEUR_LEGENDE_EP, COULEUR_LEGENDE_ECOFI,
    GRIS
)
from .dates import est_jour_non_ouvre, obtenir_nom_jour, calculer_numero_semaine
from .styles import appliquer_bordures, coloriser_legende, appliquer_style_cellule


def ajouter_audiences(ws, ligne_actuelle, mois, annee, nb_jours,
                     audiences, couleur_legende, repartition,
                     bordure_debut=True, bordure_fin=False):
    """
    Ajoute un bloc d'audiences (matin, après-midi, civil, EP ou ECOFI)

    Returns:
        int: La ligne actuelle après ajout
    """
    for idx, audience in enumerate(audiences):
        # Colonne A : horaire
        ws.cell(ligne_actuelle, 1, audience['horaire'])
        # Colonne B : nom de l'audience
        ws.cell(ligne_actuelle, 2, audience['nom_planning'])
        coloriser_legende(ws, ligne_actuelle, couleur_legende, nb_jours)

        # Bordure épaisse en haut pour la première ligne
        if idx == 0 and bordure_debut:
            for col in range(1, 3 + nb_jours):
                cell = ws.cell(ligne_actuelle, col)
                appliquer_bordures(cell, top='medium')

        # Chercher la clé dans la répartition
        cle = f"{audience['horaire']}_{audience['type']}"

        for jour in range(1, nb_jours + 1):
            date = datetime(annee, mois, jour)
            nom_jour = obtenir_nom_jour(date)
            col = 2 + jour
            cell = ws.cell(ligne_actuelle, col)

            if est_jour_non_ouvre(date):
                appliquer_style_cellule(cell, GRIS)
            elif cle in repartition:
                regles = repartition[cle]
                numero_semaine = calculer_numero_semaine(date, mois, annee)

                if numero_semaine and nom_jour in regles and numero_semaine in regles[nom_jour]:
                    info = regles[nom_jour][numero_semaine]
                    # Appliquer le style avec le code si nécessaire
                    if '/' in info['code']:
                        appliquer_style_cellule(cell, info['couleur'], texte=info['code'])
                    else:
                        appliquer_style_cellule(cell, info['couleur'])
                else:
                    appliquer_style_cellule(cell, GRIS)
            else:
                appliquer_style_cellule(cell, GRIS)

            # Bordures
            appliquer_bordures(cell, top='medium' if (idx == 0 and bordure_debut) else 'thin')

        # Bordure épaisse en bas pour la dernière ligne
        if idx == len(audiences) - 1 and bordure_fin:
            for col in range(1, 3 + nb_jours):
                cell = ws.cell(ligne_actuelle, col)
                appliquer_bordures(cell, bottom='medium')

        ligne_actuelle += 1

    return ligne_actuelle


def ajouter_toutes_audiences(ws, ligne_actuelle, mois, annee, nb_jours, repartition):
    """
    Ajoute tous les blocs d'audiences

    Returns:
        int: La ligne actuelle après ajout
    """
    # Audiences du matin
    ligne_actuelle = ajouter_audiences(
        ws, ligne_actuelle, mois, annee, nb_jours,
        AUDIENCES_MATIN, COULEUR_LEGENDE_MATIN, repartition,
        bordure_debut=True, bordure_fin=True
    )

    # Audiences de l'après-midi
    ligne_actuelle = ajouter_audiences(
        ws, ligne_actuelle, mois, annee, nb_jours,
        AUDIENCES_APRES_MIDI, COULEUR_LEGENDE_APRES_MIDI, repartition,
        bordure_debut=True, bordure_fin=True
    )

    # Audiences civiles
    ligne_actuelle = ajouter_audiences(
        ws, ligne_actuelle, mois, annee, nb_jours,
        AUDIENCES_CIVILES, COULEUR_LEGENDE_CIVIL, repartition,
        bordure_debut=True, bordure_fin=True
    )

    # Audiences EP
    ligne_actuelle = ajouter_audiences(
        ws, ligne_actuelle, mois, annee, nb_jours,
        AUDIENCES_EP, COULEUR_LEGENDE_EP, repartition,
        bordure_debut=True, bordure_fin=True
    )

    # Audiences ECOFI
    ligne_actuelle = ajouter_audiences(
        ws, ligne_actuelle, mois, annee, nb_jours,
        AUDIENCES_ECOFI, COULEUR_LEGENDE_ECOFI, repartition,
        bordure_debut=True, bordure_fin=True
    )

    return ligne_actuelle


def ajouter_assises_et_ccd(ws, ligne_actuelle, nb_jours):
    """
    Ajoute les lignes Assises et CCD avec légendes bordeaux

    Returns:
        int: La ligne actuelle après ajout
    """
    from .config import COULEUR_LEGENDE_CRIMINEL, GRIS

    # ASSISES
    ws.merge_cells(start_row=ligne_actuelle, start_column=1,
                  end_row=ligne_actuelle, end_column=2)
    ws.cell(ligne_actuelle, 1, 'Assises')

    # Légende en bordeaux avec police blanche
    cell_a = ws.cell(ligne_actuelle, 1)
    appliquer_style_cellule(cell_a, COULEUR_LEGENDE_CRIMINEL, gras=False)
    appliquer_bordures(cell_a, top='medium', bottom='medium')

    cell_b = ws.cell(ligne_actuelle, 2)
    appliquer_style_cellule(cell_b, COULEUR_LEGENDE_CRIMINEL, gras=False)
    appliquer_bordures(cell_b, top='medium', bottom='medium')

    # Cellules de dates en gris
    for jour in range(1, nb_jours + 1):
        col = 2 + jour
        cell = ws.cell(ligne_actuelle, col)
        appliquer_style_cellule(cell, GRIS)
        appliquer_bordures(cell, top='medium', bottom='medium')

    ligne_actuelle += 1

    # CCD
    ws.merge_cells(start_row=ligne_actuelle, start_column=1,
                  end_row=ligne_actuelle, end_column=2)
    ws.cell(ligne_actuelle, 1, 'CCD')

    # Légende en bordeaux avec police blanche
    cell_a = ws.cell(ligne_actuelle, 1)
    appliquer_style_cellule(cell_a, COULEUR_LEGENDE_CRIMINEL, gras=False)
    appliquer_bordures(cell_a)

    cell_b = ws.cell(ligne_actuelle, 2)
    appliquer_style_cellule(cell_b, COULEUR_LEGENDE_CRIMINEL, gras=False)
    appliquer_bordures(cell_b)

    # Cellules de dates en gris
    for jour in range(1, nb_jours + 1):
        col = 2 + jour
        cell = ws.cell(ligne_actuelle, col)
        appliquer_style_cellule(cell, GRIS)
        appliquer_bordures(cell)

    ligne_actuelle += 1
    return ligne_actuelle