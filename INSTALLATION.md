# Guide d'Installation Rapide

Ce guide vous permet de commencer à utiliser le générateur de planning en quelques minutes.

## Étape 1 : Vérifier Python

Ouvrez un terminal (ou invite de commandes) et tapez :

```bash
python --version
```

ou

```bash
python3 --version
```

Vous devriez voir quelque chose comme "Python 3.8.10" ou une version supérieure. Si Python n'est pas installé, téléchargez-le depuis [python.org](https://www.python.org/downloads/).

## Étape 2 : Installer openpyxl

Dans le terminal, tapez :

```bash
pip install openpyxl
```

Ou si cette commande ne fonctionne pas :

```bash
pip3 install openpyxl
```

Sur certains systèmes Linux, vous devrez peut-être utiliser :

```bash
pip install openpyxl --break-system-packages
```

## Étape 3 : Organiser vos fichiers

Placez tous les fichiers dans le même dossier :

```
mon_dossier_planning/
├── generer_planning_mensuel.py
├── tableau_répartition_audiences.xlsx
├── generer_planning.bat (Windows)
├── generer_planning.sh (Linux/Mac)
└── README.md
```

## Étape 4 : Générer votre premier planning

### Sur Windows

Double-cliquez sur `generer_planning.bat` et suivez les instructions, ou ouvrez une invite de commandes dans le dossier et tapez :

```bash
generer_planning.bat 2 2025
```

### Sur Linux/Mac

Ouvrez un terminal dans le dossier et tapez :

```bash
./generer_planning.sh 2 2025
```

Ou utilisez directement Python :

```bash
python3 generer_planning_mensuel.py 2 2025 tableau_répartition_audiences.xlsx
```

## Étape 5 : Ouvrir le planning généré

Le fichier créé s'appelle `planning_02_2025.xlsx` (pour février 2025). Ouvrez-le avec Excel, LibreOffice Calc, ou Google Sheets.

## Problèmes fréquents

### "Python n'est pas reconnu..."
- Installez Python depuis python.org
- Redémarrez votre terminal/invite de commandes

### "pip n'est pas reconnu..."
- Utilisez `python -m pip install openpyxl` au lieu de `pip install openpyxl`

### "ModuleNotFoundError: No module named 'openpyxl'"
- Vous n'avez pas installé openpyxl. Retournez à l'étape 2.

### Le script ne trouve pas le tableau de répartition
- Vérifiez que `tableau_répartition_audiences.xlsx` est bien dans le même dossier que le script
- Vérifiez l'orthographe du nom de fichier

## Besoin d'aide ?

Consultez le fichier `README.md` pour plus de détails sur l'utilisation du script.

## Génération de plannings pour toute l'année

Pour générer rapidement tous les plannings de l'année, vous pouvez créer un script :

### Windows (creer_tous_plannings.bat)
```batch
@echo off
for %%m in (1 2 3 4 5 6 7 8 9 10 11 12) do (
    echo Generation du planning pour le mois %%m...
    python generer_planning_mensuel.py %%m 2025 tableau_répartition_audiences.xlsx
)
echo Tous les plannings ont ete generes !
```

### Linux/Mac (creer_tous_plannings.sh)
```bash
#!/bin/bash
for mois in {1..12}; do
    echo "Génération du planning pour le mois $mois..."
    python3 generer_planning_mensuel.py $mois 2025 tableau_répartition_audiences.xlsx
done
echo "Tous les plannings ont été générés !"
```

Rendez le script exécutable sur Linux/Mac :
```bash
chmod +x creer_tous_plannings.sh
./creer_tous_plannings.sh
```
