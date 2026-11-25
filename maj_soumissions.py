import sqlite3

conn = sqlite3.connect('offres.db')
cursor = conn.cursor()

# Création de la table des soumissions
cursor.execute('''
    CREATE TABLE IF NOT EXISTS soumissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        reference_ao TEXT,
        client TEXT NOT NULL,
        objet TEXT NOT NULL,
        montant_estime REAL,
        devise TEXT DEFAULT 'DZD',
        date_limite TEXT,
        statut TEXT DEFAULT 'En preparation',
        notes TEXT
    )
''')
print("Table 'soumissions' creee avec succes !")

# Ajout de la colonne de liaison (si elle n'existe pas)
try:
    cursor.execute("ALTER TABLE offres ADD COLUMN soumission_id INTEGER REFERENCES soumissions(id)")
    print("Lien 'soumission_id' ajoute aux projets.")
except sqlite3.OperationalError:
    print("Le lien 'soumission_id' existe deja.")

conn.commit()
conn.close()
print("Base de donnees prete pour le module ERP !")