#!/usr/bin/env python3
"""
Interface graphique pour le générateur de planning
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
import os
from datetime import datetime

# Ajouter le répertoire src au path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.config import DOSSIER_DATA, DOSSIER_SORTIES, MOIS_FR
from src.planning import (
    creer_dossiers,
    trouver_fichier_repartition,
    generer_nom_fichier,
    creer_planning_mensuel
)


class PlanningGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MagPlan - Générateur de Planning")
        self.root.geometry("500x400")
        self.root.resizable(False, False)

        # Style
        style = ttk.Style()
        style.theme_use('clam')

        self.create_widgets()

    def create_widgets(self):
        # Titre
        title_frame = tk.Frame(self.root, bg="#2E5090", height=80)
        title_frame.pack(fill=tk.X)

        title_label = tk.Label(
            title_frame,
            text="📅 Générateur de Planning Mensuel",
            font=("Calibri", 18, "bold"),
            bg="#2E5090",
            fg="white"
        )
        title_label.pack(pady=20)

        # Contenu principal
        main_frame = tk.Frame(self.root, padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Sélection du mois
        tk.Label(
            main_frame,
            text="Mois :",
            font=("Calibri", 11)
        ).grid(row=0, column=0, sticky="w", pady=10)

        self.mois_var = tk.StringVar()
        mois_combo = ttk.Combobox(
            main_frame,
            textvariable=self.mois_var,
            values=[f"{i:02d} - {MOIS_FR[i - 1]}" for i in range(1, 13)],
            state="readonly",
            width=25,
            font=("Calibri", 10)
        )
        mois_combo.grid(row=0, column=1, padx=10, pady=10)
        mois_combo.current(datetime.now().month - 1)

        # Sélection de l'année
        tk.Label(
            main_frame,
            text="Année :",
            font=("Calibri", 11)
        ).grid(row=1, column=0, sticky="w", pady=10)

        self.annee_var = tk.StringVar(value=str(datetime.now().year))
        annee_spin = tk.Spinbox(
            main_frame,
            from_=2024,
            to=2030,
            textvariable=self.annee_var,
            width=24,
            font=("Calibri", 10)
        )
        annee_spin.grid(row=1, column=1, padx=10, pady=10)

        # Fichier de répartition
        tk.Label(
            main_frame,
            text="Fichier de répartition :",
            font=("Calibri", 11)
        ).grid(row=2, column=0, sticky="w", pady=10)

        self.fichier_var = tk.StringVar()
        fichier_label = tk.Label(
            main_frame,
            textvariable=self.fichier_var,
            font=("Calibri", 9),
            fg="gray"
        )
        fichier_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=5)

        # Boutons
        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=30)

        generate_btn = tk.Button(
            button_frame,
            text="🚀 Générer le Planning",
            command=self.generer,
            bg="#4CAF50",
            fg="white",
            font=("Calibri", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        generate_btn.pack(side=tk.LEFT, padx=5)

        quit_btn = tk.Button(
            button_frame,
            text="❌ Quitter",
            command=self.root.quit,
            bg="#f44336",
            fg="white",
            font=("Calibri", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        quit_btn.pack(side=tk.LEFT, padx=5)

        # Status bar
        self.status_var = tk.StringVar(value="Prêt")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=("Calibri", 9)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Vérifier le fichier de répartition
        self.verifier_fichier_repartition()

    def verifier_fichier_repartition(self):
        """Vérifie si le fichier de répartition existe"""
        creer_dossiers()
        fichier = trouver_fichier_repartition()

        if fichier:
            self.fichier_var.set(f"✓ Trouvé : {os.path.basename(fichier)}")
        else:
            self.fichier_var.set("❌ Non trouvé dans data/")
            messagebox.showwarning(
                "Fichier manquant",
                "Le fichier 'tableau_repartition_audiences.xlsx' \n"
                "n'a pas été trouvé dans le dossier 'data/'.\n\n"
                "Veuillez l'ajouter avant de générer un planning."
            )

    def generer(self):
        """Génère le planning"""
        try:
            # Récupérer les valeurs
            mois_str = self.mois_var.get().split(" - ")[0]
            mois = int(mois_str)
            annee = int(self.annee_var.get())

            # Vérifier le fichier de répartition
            fichier_repartition = trouver_fichier_repartition()
            if not fichier_repartition:
                messagebox.showerror(
                    "Erreur",
                    "Fichier de répartition non trouvé !\n\n"
                    "Veuillez placer 'tableau_repartition_audiences.xlsx'\n"
                    "dans le dossier 'data/'"
                )
                return

            # Générer le nom de fichier
            nom_fichier = generer_nom_fichier(mois, annee)
            fichier_sortie = os.path.join(DOSSIER_SORTIES, nom_fichier)

            # Mise à jour du status
            self.status_var.set(f"Génération en cours...")
            self.root.update()

            # Générer le planning
            creer_planning_mensuel(mois, annee, fichier_repartition, fichier_sortie)

            # Succès
            self.status_var.set(f"✓ Planning généré avec succès")

            # Demander si on veut ouvrir le fichier
            reponse = messagebox.askyesno(
                "Succès",
                f"Planning généré avec succès !\n\n"
                f"Fichier : {nom_fichier}\n"
                f"Dossier : {DOSSIER_SORTIES}/\n\n"
                f"Voulez-vous ouvrir le fichier ?"
            )

            if reponse:
                os.startfile(fichier_sortie)  # Windows seulement

        except Exception as e:
            self.status_var.set("❌ Erreur")
            messagebox.showerror(
                "Erreur",
                f"Une erreur s'est produite :\n\n{str(e)}"
            )


def main():
    root = tk.Tk()
    app = PlanningGeneratorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()