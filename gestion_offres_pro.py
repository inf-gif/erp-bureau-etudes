import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, date
from tkcalendar import DateEntry
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer


# ==================== BASE DE DONNÉES ====================
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

    def statistiques(self):
        self.cursor.execute("SELECT COUNT(*) FROM offres")
        total = self.cursor.fetchone()[0]
        self.cursor.execute("SELECT COUNT(*) FROM offres WHERE statut='Active'")
        actives = self.cursor.fetchone()[0]
        self.cursor.execute("SELECT COALESCE(SUM(prix),0) FROM offres WHERE statut='Active'")
        ca = self.cursor.fetchone()[0]
        aujourd_hui = date.today().isoformat()
        self.cursor.execute("""
            SELECT COUNT(*) FROM offres
            WHERE date_fin != '' AND date_fin < ? AND statut='Active'
        """, (aujourd_hui,))
        expirees = self.cursor.fetchone()[0]
        return {"total": total, "actives": actives, "ca": ca, "expirees": expirees}


# ==================== EXPORTS ====================
def exporter_excel(offres, chemin):
    wb = Workbook()
    ws = wb.active
    ws.title = "Offres"

    entetes = ["ID", "Nom", "Description", "Prix (€)", "Date début", "Date fin", "Statut"]
    ws.append(entetes)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2C3E50")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for offre in offres:
        ws.append(list(offre))

    largeurs = [6, 25, 40, 12, 14, 14, 12]
    for i, w in enumerate(largeurs, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    wb.save(chemin)


def exporter_pdf(offres, chemin):
    doc = SimpleDocTemplate(chemin, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    titre = Paragraph("<b>Rapport des Offres</b>", styles["Title"])
    elements.append(titre)
    elements.append(Paragraph(f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
                              styles["Normal"]))
    elements.append(Spacer(1, 15))

    data = [["ID", "Nom", "Description", "Prix", "Début", "Fin", "Statut"]]
    for o in offres:
        data.append([
            str(o[0]), str(o[1]), str(o[2])[:30], f"{o[3]:.2f} €",
            str(o[4]), str(o[5]), str(o[6])
        ])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#ECF0F1")]),
    ]))
    elements.append(table)

    doc.build(elements)


# ==================== INTERFACE ====================
class AppGestionOffres:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestion des Offres - Édition Pro")
        self.root.geometry("1050x780")
        self.root.configure(bg="#f0f0f0")

        self.db = Database()
        self.id_selectionne = None

        self.creer_widgets()
        self.charger_offres()
        self.actualiser_stats()

    def creer_widgets(self):
        # ---- Titre + Stats ----
        header = tk.Frame(self.root, bg="#2c3e50", height=70)
        header.pack(fill="x")
        tk.Label(header, text="📦 Gestion des Offres Pro",
                 font=("Arial", 18, "bold"), bg="#2c3e50", fg="white").pack(side="left", padx=20, pady=15)

        self.label_stats = tk.Label(header, text="", font=("Arial", 10),
                                     bg="#2c3e50", fg="#ecf0f1")
        self.label_stats.pack(side="right", padx=20)

        # ---- Formulaire ----
        frame_form = tk.LabelFrame(self.root, text="Informations de l'offre",
                                    font=("Arial", 11, "bold"), bg="#f0f0f0", padx=10, pady=10)
        frame_form.pack(fill="x", padx=20, pady=8)

        tk.Label(frame_form, text="Nom :", bg="#f0f0f0").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_nom = tk.Entry(frame_form, width=30)
        self.entry_nom.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_form, text="Prix (€) :", bg="#f0f0f0").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.entry_prix = tk.Entry(frame_form, width=15)
        self.entry_prix.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(frame_form, text="Description :", bg="#f0f0f0").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_desc = tk.Entry(frame_form, width=30)
        self.entry_desc.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame_form, text="Statut :", bg="#f0f0f0").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.combo_statut = ttk.Combobox(frame_form, values=["Active", "Inactive", "Expirée"], width=13)
        self.combo_statut.set("Active")
        self.combo_statut.grid(row=1, column=3, padx=5, pady=5)

        tk.Label(frame_form, text="Date début :", bg="#f0f0f0").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.date_debut = DateEntry(frame_form, width=27, date_pattern="yyyy-mm-dd")
        self.date_debut.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(frame_form, text="Date fin :", bg="#f0f0f0").grid(row=2, column=2, sticky="w", padx=5, pady=5)
        self.date_fin = DateEntry(frame_form, width=13, date_pattern="yyyy-mm-dd")
        self.date_fin.grid(row=2, column=3, padx=5, pady=5)

        # ---- Boutons d'action ----
        frame_btn = tk.Frame(self.root, bg="#f0f0f0")
        frame_btn.pack(pady=8)

        def bouton(parent, txt, bg, cmd, w=13):
            return tk.Button(parent, text=txt, bg=bg, fg="white",
                             font=("Arial", 10, "bold"), width=w, command=cmd)

        bouton(frame_btn, "➕ Ajouter", "#27ae60", self.ajouter_offre).pack(side="left", padx=4)
        bouton(frame_btn, "✏️ Modifier", "#f39c12", self.modifier_offre).pack(side="left", padx=4)
        bouton(frame_btn, "🗑️ Supprimer", "#e74c3c", self.supprimer_offre).pack(side="left", padx=4)
        bouton(frame_btn, "🔄 Vider", "#7f8c8d", self.vider_champs).pack(side="left", padx=4)
        bouton(frame_btn, "📊 Export Excel", "#2980b9", self.exporter_xls).pack(side="left", padx=4)
        bouton(frame_btn, "📄 Export PDF", "#8e44ad", self.exporter_pdf_click).pack(side="left", padx=4)

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
            self.tree.column(col, width=110, anchor="center")
        self.tree.column("Description", width=220)

        scrollbar = ttk.Scrollbar(frame_table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.selectionner_offre)

        # Tags couleur
        self.tree.tag_configure("expiree", background="#ffe6e6")
        self.tree.tag_configure("active", background="#e6ffe6")

    # ---- Helpers ----
    def actualiser_stats(self):
        s = self.db.statistiques()
        self.label_stats.config(
            text=f"📊 Total: {s['total']}  |  ✅ Actives: {s['actives']}  "
                 f"|  💰 CA: {s['ca']:.2f} €  |  ⚠️ Expirées: {s['expirees']}"
        )

    def charger_offres(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        aujourd_hui = date.today().isoformat()
        for offre in self.db.recuperer_tout():
            tag = ""
            if offre[6] == "Active":
                tag = "expiree" if offre[5] and offre[5] < aujourd_hui else "active"
            self.tree.insert("", "end", values=offre, tags=(tag,))
        self.actualiser_stats()

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

    # ---- Actions CRUD ----
    def ajouter_offre(self):
        if not self.valider_champs(): return
        self.db.ajouter(
            self.entry_nom.get().strip(), self.entry_desc.get().strip(),
            float(self.entry_prix.get()),
            self.date_debut.get(), self.date_fin.get(),
            self.combo_statut.get()
        )
        messagebox.showinfo("Succès", "Offre ajoutée !")
        self.vider_champs(); self.charger_offres()

    def modifier_offre(self):
        if self.id_selectionne is None:
            messagebox.showwarning("Aucune sélection", "Sélectionnez une offre.")
            return
        if not self.valider_champs(): return
        self.db.modifier(
            self.id_selectionne,
            self.entry_nom.get().strip(), self.entry_desc.get().strip(),
            float(self.entry_prix.get()),
            self.date_debut.get(), self.date_fin.get(),
            self.combo_statut.get()
        )
        messagebox.showinfo("Succès", "Offre modifiée !")
        self.vider_champs(); self.charger_offres()

    def supprimer_offre(self):
        if self.id_selectionne is None:
            messagebox.showwarning("Aucune sélection", "Sélectionnez une offre.")
            return
        if messagebox.askyesno("Confirmation", "Supprimer cette offre ?"):
            self.db.supprimer(self.id_selectionne)
            messagebox.showinfo("Succès", "Offre supprimée.")
            self.vider_champs(); self.charger_offres()

    def selectionner_offre(self, event):
        selection = self.tree.selection()
        if not selection: return
        v = self.tree.item(selection[0], "values")
        self.id_selectionne = int(v[0])
        self.entry_nom.delete(0, "end"); self.entry_nom.insert(0, v[1])
        self.entry_desc.delete(0, "end"); self.entry_desc.insert(0, v[2])
        self.entry_prix.delete(0, "end"); self.entry_prix.insert(0, v[3])
        self.date_debut.set_date(datetime.strptime(v[4], "%Y-%m-%d").date() if v[4] else date.today())
        self.date_fin.set_date(datetime.strptime(v[5], "%Y-%m-%d").date() if v[5] else date.today())
        self.combo_statut.set(v[6])

    def vider_champs(self):
        self.id_selectionne = None
        for e in (self.entry_nom, self.entry_desc, self.entry_prix):
            e.delete(0, "end")
        self.date_debut.set_date(date.today())
        self.date_fin.set_date(date.today())
        self.combo_statut.set("Active")
        self.tree.selection_remove(self.tree.selection())

    # ---- Exports ----
    def exporter_xls(self):
        chemin = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=f"offres_{date.today()}.xlsx")
        if not chemin: return
        try:
            exporter_excel(self.db.recuperer_tout(), chemin)
            messagebox.showinfo("Export réussi", f"Fichier créé :\n{chemin}")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def exporter_pdf_click(self):
        chemin = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialfile=f"rapport_offres_{date.today()}.pdf")
        if not chemin: return
        try:
            exporter_pdf(self.db.recuperer_tout(), chemin)
            messagebox.showinfo("Export réussi", f"PDF créé :\n{chemin}")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = AppGestionOffres(root)
    root.mainloop()