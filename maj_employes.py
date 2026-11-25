import sqlite3
conn = sqlite3.connect('offres.db')
conn.execute('''CREATE TABLE IF NOT EXISTS employes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    role TEXT,
    email TEXT,
    telephone TEXT
)''')
conn.execute('''CREATE TABLE IF NOT EXISTS taches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    projet_id INTEGER,
    employe_id INTEGER,
    description TEXT NOT NULL,
    statut TEXT DEFAULT 'A faire',
    date_echeance TEXT,
    FOREIGN KEY(projet_id) REFERENCES offres(id),
    FOREIGN KEY(employe_id) REFERENCES employes(id)
)''')
conn.commit()
conn.close()
print("Tables 'employes' et 'taches' creees avec succes !")