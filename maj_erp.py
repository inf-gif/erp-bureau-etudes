import sqlite3

conn = sqlite3.connect('offres.db')
cursor = conn.cursor()

# 1. Table des Soumissions (Appels d'offres et Devis)
cursor.execute('''
    CREATE TABLE IF NOT EXISTS soumissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        reference_ao TEXT,
        client TEXT NOT NULL,
        objet TEXT NOT NULL,
        montant_estime REAL,
        devise TEXT DEFAULT 'DZD',
        date_limite TEXT,
        statut TEXT DEFAULT 'En préparation',
        notes TEXT
    )
''')

# 2. Ajout du lien entre Projet et Soumission (Clé étrangère)
try:
    cursor.execute("ALTER TABLE offres ADD COLUMN soumission_id INTEGER REFERENCES soumissions(id)")
    print("Lien 'soumission_id' ajouté aux projets.")
except sqlite3.OperationalError:
    print("Le lien 'soumission_id' existe déjà.")

conn.commit()
conn.close()
print("Base de données mise à jour pour le module ERP !")