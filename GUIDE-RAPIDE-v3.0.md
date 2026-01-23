# ⚡ VERSION 3.0 - FORMATAGE PROFESSIONNEL

## 🎯 Commande Unique

```bash
python generer_planning_mensuel.py 11 2025
```

→ Génère `11 - NOVEMBRE 2025.xlsx` **formaté professionnellement**

---

## ✨ Nouveautés v3.0

✅ **Bordures complètes** sur toutes les cellules  
✅ **Couleurs des légendes** (colonnes A-C) selon les zones  
✅ **Colonnes figées** (A, B, C toujours visibles)  
✅ **Lignes Assises & CCD** ajoutées (grisées)  
✅ **Séparateurs visuels** (traits gras) entre zones

---

## 🎨 Zones de Couleur

- **Permanences** : Orange clair
- **Perm. Nuit** : Bleu
- **Audiences 9h** : Bleu clair
- **Audiences 13h30** : Orange très clair
- **Civiles** : Magenta
- **Exécution Peines** : Orange
- **ECOFI** : Bleu clair
- **Criminel** : Bordeaux

---

## 📐 Séparations Visuelles

```
Permanences (2-11)
═══════════════════
Débats JLD + Perm Nuit (12-13)
═══════════════════
Audiences 9h (14-21)
═══════════════════
Audiences 13h30 (22-33)
═══════════════════
Civiles (34-35)
═══════════════════
Exécution Peines (36-40)
═══════════════════
ECOFI (41-43)
═══════════════════
Assises + CCD (44-45)
```

---

## 🔒 Colonnes Figées

Les 3 premières colonnes restent visibles quand vous défilez → **Légendes toujours accessibles**

---

## 📋 Installation (30 secondes)

```bash
pip install openpyxl
mkdir data && cp tableau_repartition_audiences.xlsx data/
python generer_planning_mensuel.py 11 2025
```

---

**C'est LA version parfaite !** 🎊  
**100% conforme au modèle** ✨
