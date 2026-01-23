# Générateur de Planning Mensuel - Récapitulatif

## ✅ Fichiers livrés

### Scripts principaux
1. **generer_planning_mensuel.py** - Script principal de génération des plannings
2. **generer_annee_complete.py** - Script pour générer les 12 mois d'une année en une fois

### Scripts d'aide
3. **generer_planning.bat** - Script Windows pour faciliter l'utilisation
4. **generer_planning.sh** - Script Linux/Mac pour faciliter l'utilisation

### Documentation
5. **README.md** - Documentation complète du script
6. **INSTALLATION.md** - Guide d'installation rapide
7. **RECAPITULATIF.md** - Ce fichier

### Exemples
8. **planning_02_2025.xlsx** - Exemple de planning généré pour février 2025
9. **planning_03_2025.xlsx** - Exemple de planning généré pour mars 2025

## 🎯 Fonctionnalités

Le script génère automatiquement un planning mensuel avec :

### ✓ Structure complète
- Ligne 1 : Dates du mois (une colonne par jour)
- Lignes 2-13 : 12 permanences avec coloration automatique
- Lignes 14+ : Toutes les audiences du tableau de répartition

### ✓ Coloration intelligente
- Calcul automatique du numéro de semaine (1er, 2e, 3e, 4e, 5e lundi/mardi/etc.)
- Application des règles du tableau de répartition
- Gestion des cas multiples (ex: M/S)
- Respect des couleurs par section :
  - A (PAP) : Violet #7030A0
  - M (PMF) : Vert #00FF00
  - E (ECOFI) : Cyan clair #66FFFF
  - S (STUPS) : Jaune #FFFF00
  - P (EP) : Orange #FF9933
  - C (CIVIL) : Magenta #FF00FF
  - V (PRA PMF) : Rose #FF9999
  - J (PRA STUPS) : Jaune #FFFF00
  - N (PCC) : Cyan foncé #009999
  - SG (MAG SG) : Bleu #0070C0

### ✓ Respect de la nomenclature
- Conversion automatique de la nomenclature du tableau vers celle du planning
- Ex: "11 JU" → "11ème JU - Route"

## 📋 Utilisation rapide

### Cas simple : Un mois
```bash
python generer_planning_mensuel.py 2 2025 tableau_répartition_audiences.xlsx
```

### Cas avancé : Toute l'année
```bash
python generer_annee_complete.py 2025 tableau_répartition_audiences.xlsx
```

### Avec les scripts d'aide (Windows)
```bash
generer_planning.bat 2 2025
```

### Avec les scripts d'aide (Linux/Mac)
```bash
./generer_planning.sh 2 2025
```

## 🔧 Personnalisation possible

Si vous souhaitez modifier le comportement du script, les sections clés sont :

### Dans generer_planning_mensuel.py

**Lignes 14-26** : `CODES_COULEURS`
- Modifiez les couleurs RGB des sections

**Lignes 29-41** : `PERMANENCES`
- Modifiez les libellés ou couleurs des permanences

**Lignes 44-73** : `AUDIENCES`
- Ajoutez, supprimez ou modifiez des audiences
- Format : `{'horaire': '9h', 'type': '11 JU', 'nom_planning': '11ème JU - Route'}`

**Fonction `calculer_numero_semaine`** (lignes 116-132)
- Modifiez la logique de calcul des semaines si nécessaire

## 📊 Exemple de résultat

Pour février 2025, le script a généré :
- 28 colonnes de dates
- 12 lignes de permanences (toutes colorisées)
- 30 lignes d'audiences
- 171 cellules colorisées sur 840 (20.4%)

## ⚠️ Limitations connues

1. **Vacances judiciaires** : Le script ne gère pas automatiquement les périodes de vacances (comme la première semaine de janvier). Ces périodes doivent être traitées manuellement.

2. **Week-ends** : Les samedis et dimanches ne sont pas colorisés.

3. **Dépendance au tableau de répartition** : Le script se base entièrement sur le fichier `tableau_répartition_audiences.xlsx`. Toute modification de ce tableau nécessite une nouvelle génération des plannings.

4. **Noms des magistrats** : Le script ne remplit pas les noms des magistrats dans les cellules. Seule la coloration par section est effectuée.

## 🆘 Dépannage

### Problème : "No module named 'openpyxl'"
**Solution** : Installez la bibliothèque avec `pip install openpyxl`

### Problème : "File not found: tableau_répartition_audiences.xlsx"
**Solution** : Assurez-vous que le fichier est dans le même dossier que le script

### Problème : Planning vide ou sans couleurs
**Solution** : Vérifiez que :
- Le mois et l'année sont corrects
- Le tableau de répartition contient bien les règles pour ce mois
- Les noms des audiences correspondent entre le script et le tableau

### Problème : Couleurs incorrectes
**Solution** : Vérifiez le mapping dans `CODES_COULEURS` (lignes 14-26)

## 📞 Support

Pour toute question ou amélioration, consultez :
1. **README.md** pour la documentation détaillée
2. **INSTALLATION.md** pour les problèmes d'installation
3. Les commentaires dans le code source

## 🎓 Évolutions possibles

Le script peut être étendu pour :
- Gérer automatiquement les vacances judiciaires
- Importer les noms des magistrats depuis un autre fichier
- Générer des statistiques de charge de travail
- Créer des plannings multi-juridictions
- Exporter en PDF
- Envoyer les plannings par email automatiquement

## ✨ Version

Version 1.0 - Octobre 2025
Créé pour l'automatisation des plannings de magistrats
