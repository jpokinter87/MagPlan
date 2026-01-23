"""
Utilitaires pour la gestion des dates et jours fériés
"""

from datetime import datetime, timedelta
from calendar import monthrange


def est_weekend(date):
    """Vérifie si la date est un weekend"""
    return date.weekday() >= 5


def est_jour_ferie(date):
    """Vérifie si la date est un jour férié en France"""
    annee = date.year
    mois = date.month
    jour = date.day

    # Jours fériés fixes
    jours_feries_fixes = [
        (1, 1),  # Jour de l'an
        (5, 1),  # Fête du travail
        (5, 8),  # Victoire 1945
        (7, 14),  # Fête nationale
        (8, 15),  # Assomption
        (11, 1),  # Toussaint
        (11, 11),  # Armistice 1918
        (12, 25)  # Noël
    ]

    if (mois, jour) in jours_feries_fixes:
        return True

    # Pâques et jours fériés mobiles
    a = annee % 19
    b = annee // 100
    c = annee % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mois_paques = (h + l - 7 * m + 114) // 31
    jour_paques = ((h + l - 7 * m + 114) % 31) + 1

    date_paques = datetime(annee, mois_paques, jour_paques)
    lundi_paques = date_paques + timedelta(days=1)
    ascension = date_paques + timedelta(days=39)
    lundi_pentecote = date_paques + timedelta(days=50)

    jours_feries_mobiles = [lundi_paques, ascension, lundi_pentecote]

    for jour_ferie in jours_feries_mobiles:
        if date.date() == jour_ferie.date():
            return True

    return False


def est_jour_non_ouvre(date):
    """Vérifie si la date est un weekend ou un jour férié"""
    return est_weekend(date) or est_jour_ferie(date)


def obtenir_nom_jour(date):
    """Retourne le nom du jour en français"""
    jours = ['LUNDI', 'MARDI', 'MERCREDI', 'JEUDI', 'VENDREDI', 'SAMEDI', 'DIMANCHE']
    return jours[date.weekday()]


def calculer_numero_semaine(date, mois, annee):
    """Calcule le numéro d'occurrence du jour de la semaine dans le mois"""
    jour_semaine = date.weekday()
    nb_jours = monthrange(annee, mois)[1]

    occurrences = []
    for jour in range(1, nb_jours + 1):
        date_test = datetime(annee, mois, jour)
        if date_test.weekday() == jour_semaine:
            occurrences.append(jour)

    if date.day in occurrences:
        return occurrences.index(date.day) + 1
    return None