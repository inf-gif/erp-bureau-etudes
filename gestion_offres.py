import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime

# ============ BASE DE DONNÉES ============
class Database:
    def __init__(self, db_name="offres.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.creer_table()

    def creer_table(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS offres (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom TEXT NOT NULL,
                description TEXT,
                prix REAL NOT NULL,
                date_debut TEXT,
                date_fin TEXT,
                statut TEXT DEFAULT 'Active'
            )
        """)
        self.conn.commit()

    def ajouter(self, nom, description, prix, date_debut, date_fin, statut):
        self.cursor.execute("""
            INSERT INTO offres (nom, description, prix, date_debut, date_fin, statut)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nom, description, prix, date_debut, date_fin, statut))
        self.conn.commit()

    def modifier(self, id_offre, nom, description, prix, date_debut, date_fin, statut):
        self.cursor.execute("""
            UPDATE offres SET nom=?, description=?, prix=?, date_debut=?, date_fin=?, statut=?
            WHERE id=?
        """, (nom, description, prix, date_debut, date_fin, statut, id_offre))
        self.conn.commit()

    def supprimer(self, id_offre):
        self.cursor.execute("DELETE FROM offres WHERE id=?", (id_offre,))
        self.conn.commit()

    def recuperer_tout(self):
        self.cursor.execute("SELECT * FROM offres ORDER BY id DESC")
        return self.cursor.fetchall()

    def rechercher(self, terme):
        self.cursor.execute("""
            SELECT * FROM offres
            WHERE nom LIKE ? OR description LIKE ? OR statut LIKE ?
            ORDER BY id DESC
        """, (f"%{terme}%", f"%{terme}%", f"%{terme}%"))
        return self.cursor.fetchall()


# ============ INTERFACE GRAPHIQUE ============
class AppGestionOffres:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestion des Offres")
        self.root.geometry("950x650")
        self.root.configure(bg="#f0f0f0")

        self.db = Database()
        self.id_selectionne = None

        self.creer_widgets()
        self.charger_offres()

    def creer_widgets(self):
        # ---- Titre ----
        titre = tk.Label(self.root, text="📦 Gestion des Offres",
                         font=("Arial", 20, "bold"), bg="#f0f0f0", fg="#2c3e50")
        titre.pack(pady=10)

        # ---- Formulaire ----
        frame_form = tk.LabelFrame(self.root, text="Informations de l'offre",
                                    font=("Arial", 11, "bold"), bg="#f0f0f0", padx=10, pady=10)
        frame_form.pack(fill="x", padx=20, pady=5)

        # Ligne 1
        tk.Label(frame_form, text="Nom :", bg="#f0f0f0").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_nom = tk.Entry(frame_form, width=30)
        self.entry_nom.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_form, text="Prix (€) :", bg="#f0f0f0").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.entry_prix = tk.Entry(frame_form, width=15)
        self.entry_prix.grid(row=0, column=3, padx=5, pady=5)

        # Ligne 2
        tk.Label(frame_form, text="Description :", bg="#f0f0f0").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_desc = tk.Entry(frame_form, width=30)
        self.entry_desc.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame_form, text="Statut :", bg="#f0f0f0").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.combo_statut = ttk.Combobox(frame_form, values=["Active", "Inactive", "Expirée"], width=13)
        self.combo_statut.set("Active")
        self.combo_statut.grid(row=1, column=3, padx=5, pady=5)

        # Ligne 3
        tk.Label(frame_form, text="Date début (AAAA-MM-JJ) :", bg="#f0f0f0").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.entry_debut = tk.Entry(frame_form, width=30)
        self.entry_debut.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(frame_form, text="Date fin (AAAA-MM-JJ) :", bg="#f0f0f0").grid(row=2, column=2, sticky="w", padx=5, pady=5)
        self.entry_fin = tk.Entry(frame_form, width=15)
        self.entry_fin.grid(row=2, column=3, padx=5, pady=5)

        # ---- Boutons ----
        frame_btn = tk.Frame(self.root, bg="#f0f0f0")
        frame_btn.pack(pady=10)

        tk.Button(frame_btn, text="➕ Ajouter", bg="#27ae60", fg="white",
                  font=("Arial", 10, "bold"), width=12, command=self.ajouter_offre).pack(side="left", padx=5)
        tk.Button(frame_btn, text="✏️ Modifier", bg="#f39c12", fg="white",
                  font=("Arial", 10, "bold"), width=12, command=self.modifier_offre).pack(side="left", padx=5)
        tk.Button(frame_btn, text="🗑️ Supprimer", bg="#e74c3c", fg="white",
                  font=("Arial", 10, "bold"), width=12, command=self.supprimer_offre).pack(side="left", padx=5)
        tk.Button(frame_btn, text="🔄 Vider", bg="#7f8c8d", fg="white",
                  font=("Arial", 10, "bold"), width=12, command=self.vider_champs).pack(side="left", padx=5)

        # ---- Recherche ----
        frame_search = tk.Frame(self.root, bg="#f0f0f0")
        frame_search.pack(pady=5)
        tk.Label(frame_search, text="🔍 Rechercher :", bg="#f0f0f0",
                 font=("Arial", 10)).pack(side="left", padx=5)
        self.entry_recherche = tk.Entry(frame_search, width=30)
        self.entry_recherche.pack(side="left", padx=5)
        self.entry_recherche.bind("<KeyRelease>", lambda e: self.rechercher_offres())
        tk.Button(frame_search, text="Tout afficher", command=self.charger_offres).pack(side="left", padx=5)

        # ---- Tableau ----
        frame_table = tk.Frame(self.root)
        frame_table.pack(fill="both", expand=True, padx=20, pady=10)

        colonnes = ("ID", "Nom", "Description", "Prix", "Début", "Fin", "Statut")
        self.tree = ttk.Treeview(frame_table, columns=colonnes, show="headings", height=12)

        for col in colonnes:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")
        self.tree.column("Description", width=200)

        scrollbar = ttk.Scrollbar(frame_table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self.selectionner_offre)

    # ---- Actions ----
    def charger_offres(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for offre in self.db.recuperer_tout():
            self.tree.insert("", "end", values=offre)

    def rechercher_offres(self):
        terme = self.entry_recherche.get().strip()
        for row in self.tree.get_children():
            self.tree.delete(row)
        resultats = self.db.recuperer_tout() if not terme else self.db.rechercher(terme)
        for offre in resultats:
            self.tree.insert("", "end", values=offre)

    def valider_champs(self):
        nom = self.entry_nom.get().strip()
        prix = self.entry_prix.get().strip()
        if not nom:
            messagebox.showwarning("Champ vide", "Le nom est obligatoire.")
            return False
        try:
            float(prix)
        except ValueError:
            messagebox.showwarning("Prix invalide", "Le prix doit être un nombre.")
            return False
        return True

    def ajouter_offre(self):
        if not self.valider_champs():
            return
        self.db.ajouter(
            self.entry_nom.get().strip(),
            self.entry_desc.get().strip(),
            float(self.entry_prix.get()),
            self.entry_debut.get().strip(),
            self.entry_fin.get().strip(),
            self.combo_statut.get()
        )
        messagebox.showinfo("Succès", "Offre ajoutée avec succès !")
        self.vider_champs()
        self.charger_offres()

    def modifier_offre(self):
        if self.id_selectionne is None:
            messagebox.showwarning("Aucune sélection", "Sélectionnez une offre à modifier.")
            return
        if not self.valider_champs():
            return
        self.db.modifier(
            self.id_selectionne,
            self.entry_nom.get().strip(),
            self.entry_desc.get().strip(),
            float(self.entry_prix.get()),
            self.entry_debut.get().strip(),
            self.entry_fin.get().strip(),
            self.combo_statut.get()
        )
        messagebox.showinfo("Succès", "Offre modifiée avec succès !")
        self.vider_champs()
        self.charger_offres()

    def supprimer_offre(self):
        if self.id_selectionne is None:
            messagebox.showwarning("Aucune sélection", "Sélectionnez une offre à supprimer.")
            return
        if messagebox.askyesno("Confirmation", "Voulez-vous vraiment supprimer cette offre ?"):
            self.db.supprimer(self.id_selectionne)
            messagebox.showinfo("Succès", "Offre supprimée.")
            self.vider_champs()
            self.charger_offres()

    def selectionner_offre(self, event):
        selection = self.tree.selection()
        if not selection:
            return
        valeurs = self.tree.item(selection[0], "values")
        self.id_selectionne = int(valeurs[0])
        self.entry_nom.delete(0, "end"); self.entry_nom.insert(0, valeurs[1])
        self.entry_desc.delete(0, "end"); self.entry_desc.insert(0, valeurs[2])
        self.entry_prix.delete(0, "end"); self.entry_prix.insert(0, valeurs[3])
        self.entry_debut.delete(0, "end"); self.entry_debut.insert(0, valeurs[4])
        self.entry_fin.delete(0, "end"); self.entry_fin.insert(0, valeurs[5])
        self.combo_statut.set(valeurs[6])

    def vider_champs(self):
        self.id_selectionne = None
        self.entry_nom.delete(0, "end")
        self.entry_desc.delete(0, "end")
        self.entry_prix.delete(0, "end")
        self.entry_debut.delete(0, "end")
        self.entry_fin.delete(0, "end")
        self.combo_statut.set("Active")
        self.tree.selection_remove(self.tree.selection())


# ============ LANCEMENT ============
if __name__ == "__main__":
    root = tk.Tk()
    app = AppGestionOffres(root)
    root.mainloop()