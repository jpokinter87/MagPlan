#!/usr/bin/env python3
"""
Interface graphique pour le générateur de planning
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import subprocess
import platform
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


def ouvrir_fichier(chemin):
    """Ouvre un fichier avec l'application par défaut (cross-platform)"""
    if platform.system() == 'Windows':
        os.startfile(chemin)
    elif platform.system() == 'Darwin':  # macOS
        subprocess.run(['open', chemin])
    else:  # Linux
        subprocess.run(['xdg-open', chemin])


def ouvrir_dossier(chemin):
    """Ouvre un dossier dans l'explorateur de fichiers (cross-platform)"""
    if platform.system() == 'Windows':
        os.startfile(chemin)
    elif platform.system() == 'Darwin':  # macOS
        subprocess.run(['open', chemin])
    else:  # Linux
        subprocess.run(['xdg-open', chemin])


class GestionPlanningGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MagPlan - Gestion de Planning")
        self.root.geometry("520x450")
        self.root.resizable(False, False)

        # Style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TNotebook.Tab', font=('Calibri', 11, 'bold'), padding=[20, 8])

        self.create_widgets()

    def create_widgets(self):
        # Titre
        title_frame = tk.Frame(self.root, bg="#2E5090", height=70)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        title_label = tk.Label(
            title_frame,
            text="Gestion de Planning",
            font=("Calibri", 20, "bold"),
            bg="#2E5090",
            fg="white"
        )
        title_label.pack(pady=18)

        # Notebook (onglets)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Onglet Planning Mensuel
        self.tab_mensuel = tk.Frame(self.notebook, padx=20, pady=15)
        self.notebook.add(self.tab_mensuel, text="Planning Mensuel")
        self.create_tab_mensuel()

        # Onglet Planning Annuel
        self.tab_annuel = tk.Frame(self.notebook, padx=20, pady=15)
        self.notebook.add(self.tab_annuel, text="Planning Annuel")
        self.create_tab_annuel()

        # Status bar
        self.status_var = tk.StringVar(value="Prêt")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=("Calibri", 9),
            padx=10
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Vérifier le fichier de répartition
        self.verifier_fichier_repartition()

    def create_tab_mensuel(self):
        """Crée le contenu de l'onglet Planning Mensuel"""
        # Sélection du mois
        tk.Label(
            self.tab_mensuel,
            text="Mois :",
            font=("Calibri", 11)
        ).grid(row=0, column=0, sticky="w", pady=10)

        self.mois_var = tk.StringVar()
        mois_combo = ttk.Combobox(
            self.tab_mensuel,
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
            self.tab_mensuel,
            text="Année :",
            font=("Calibri", 11)
        ).grid(row=1, column=0, sticky="w", pady=10)

        self.annee_mensuel_var = tk.StringVar(value=str(datetime.now().year))
        annee_spin = tk.Spinbox(
            self.tab_mensuel,
            from_=2024,
            to=2035,
            textvariable=self.annee_mensuel_var,
            width=24,
            font=("Calibri", 10)
        )
        annee_spin.grid(row=1, column=1, padx=10, pady=10)

        # Fichier de répartition
        tk.Label(
            self.tab_mensuel,
            text="Fichier de répartition :",
            font=("Calibri", 11)
        ).grid(row=2, column=0, sticky="w", pady=10)

        self.fichier_mensuel_var = tk.StringVar()
        fichier_label = tk.Label(
            self.tab_mensuel,
            textvariable=self.fichier_mensuel_var,
            font=("Calibri", 9),
            fg="gray"
        )
        fichier_label.grid(row=2, column=1, sticky="w", padx=10)

        # Boutons
        button_frame = tk.Frame(self.tab_mensuel)
        button_frame.grid(row=3, column=0, columnspan=2, pady=30)

        generate_btn = tk.Button(
            button_frame,
            text="Générer le Planning",
            command=self.generer_mensuel,
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
            text="Quitter",
            command=self.root.quit,
            bg="#f44336",
            fg="white",
            font=("Calibri", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        quit_btn.pack(side=tk.LEFT, padx=5)

    def create_tab_annuel(self):
        """Crée le contenu de l'onglet Planning Annuel"""
        # Description
        tk.Label(
            self.tab_annuel,
            text="Génère les 12 plannings mensuels pour l'année sélectionnée.",
            font=("Calibri", 10),
            fg="gray"
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))

        # Sélection de l'année
        tk.Label(
            self.tab_annuel,
            text="Année :",
            font=("Calibri", 11)
        ).grid(row=1, column=0, sticky="w", pady=10)

        self.annee_annuel_var = tk.StringVar(value=str(datetime.now().year))
        annee_spin = tk.Spinbox(
            self.tab_annuel,
            from_=2024,
            to=2035,
            textvariable=self.annee_annuel_var,
            width=24,
            font=("Calibri", 10)
        )
        annee_spin.grid(row=1, column=1, padx=10, pady=10)

        # Fichier de répartition
        tk.Label(
            self.tab_annuel,
            text="Fichier de répartition :",
            font=("Calibri", 11)
        ).grid(row=2, column=0, sticky="w", pady=10)

        self.fichier_annuel_var = tk.StringVar()
        fichier_label = tk.Label(
            self.tab_annuel,
            textvariable=self.fichier_annuel_var,
            font=("Calibri", 9),
            fg="gray"
        )
        fichier_label.grid(row=2, column=1, sticky="w", padx=10)

        # Barre de progression
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.tab_annuel,
            variable=self.progress_var,
            maximum=12,
            length=300
        )
        self.progress_bar.grid(row=3, column=0, columnspan=2, pady=20)

        # Boutons
        button_frame = tk.Frame(self.tab_annuel)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        generate_btn = tk.Button(
            button_frame,
            text="Générer l'Année Complète",
            command=self.generer_annuel,
            bg="#2196F3",
            fg="white",
            font=("Calibri", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        generate_btn.pack(side=tk.LEFT, padx=5)

        quit_btn = tk.Button(
            button_frame,
            text="Quitter",
            command=self.root.quit,
            bg="#f44336",
            fg="white",
            font=("Calibri", 12, "bold"),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        quit_btn.pack(side=tk.LEFT, padx=5)

    def verifier_fichier_repartition(self):
        """Vérifie si le fichier de répartition existe"""
        creer_dossiers()
        fichier = trouver_fichier_repartition()

        if fichier:
            texte = f"Trouvé : {os.path.basename(fichier)}"
            self.fichier_mensuel_var.set(texte)
            self.fichier_annuel_var.set(texte)
        else:
            texte = "Non trouvé dans data/"
            self.fichier_mensuel_var.set(texte)
            self.fichier_annuel_var.set(texte)
            messagebox.showwarning(
                "Fichier manquant",
                "Le fichier 'tableau_repartition_audiences.xlsx' \n"
                "n'a pas été trouvé dans le dossier 'data/'.\n\n"
                "Veuillez l'ajouter avant de générer un planning."
            )

    def generer_mensuel(self):
        """Génère un planning mensuel"""
        try:
            # Récupérer les valeurs
            mois_str = self.mois_var.get().split(" - ")[0]
            mois = int(mois_str)
            annee = int(self.annee_mensuel_var.get())

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
            self.status_var.set("Génération en cours...")
            self.root.update()

            # Générer le planning
            creer_planning_mensuel(mois, annee, fichier_repartition, fichier_sortie)

            # Succès
            self.status_var.set("Planning généré avec succès")

            # Demander si on veut ouvrir le fichier
            reponse = messagebox.askyesno(
                "Succès",
                f"Planning généré avec succès !\n\n"
                f"Fichier : {nom_fichier}\n"
                f"Dossier : {DOSSIER_SORTIES}/\n\n"
                f"Voulez-vous ouvrir le fichier ?"
            )

            if reponse:
                ouvrir_fichier(fichier_sortie)

        except Exception as e:
            self.status_var.set("Erreur")
            messagebox.showerror(
                "Erreur",
                f"Une erreur s'est produite :\n\n{str(e)}"
            )

    def generer_annuel(self):
        """Génère les 12 plannings mensuels pour une année"""
        try:
            annee = int(self.annee_annuel_var.get())

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

            # Réinitialiser la barre de progression
            self.progress_var.set(0)
            fichiers_generes = []

            for mois in range(1, 13):
                # Mise à jour du status
                self.status_var.set(f"Génération de {MOIS_FR[mois - 1]} {annee}...")
                self.root.update()

                # Générer le planning
                nom_fichier = generer_nom_fichier(mois, annee)
                fichier_sortie = os.path.join(DOSSIER_SORTIES, nom_fichier)
                creer_planning_mensuel(mois, annee, fichier_repartition, fichier_sortie)
                fichiers_generes.append(nom_fichier)

                # Mettre à jour la progression
                self.progress_var.set(mois)
                self.root.update()

            # Succès
            self.status_var.set(f"12 plannings générés pour {annee}")

            # Demander si on veut ouvrir le dossier
            reponse = messagebox.askyesno(
                "Succès",
                f"12 plannings générés avec succès pour {annee} !\n\n"
                f"Dossier : {DOSSIER_SORTIES}/\n\n"
                f"Voulez-vous ouvrir le dossier ?"
            )

            if reponse:
                ouvrir_dossier(os.path.abspath(DOSSIER_SORTIES))

        except Exception as e:
            self.status_var.set("Erreur")
            messagebox.showerror(
                "Erreur",
                f"Une erreur s'est produite :\n\n{str(e)}"
            )


def main():
    root = tk.Tk()
    app = GestionPlanningGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
