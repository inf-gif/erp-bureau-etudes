import sqlite3

conn = sqlite3.connect('offres.db')
cursor = conn.cursor()

# Créer la table des devises
cursor.execute('''
    CREATE TABLE IF NOT EXISTS devises (
        code TEXT PRIMARY KEY,
        nom TEXT,
        taux REAL
    )
''')

# Insérer les devises par défaut
devises = [
    ('DZD', 'Dinar Algérien', 1.0),
    ('EUR', 'Euro', 145.0),
    ('USD', 'Dollar US', 135.0),
    ('GBP', 'Livre Sterling', 170.0),
    ('TND', 'Dinar Tunisien', 43.0),
    ('MAD', 'Dirham Marocain', 13.5)
]

for d in devises:
    cursor.execute("INSERT OR IGNORE INTO devises VALUES (?, ?, ?)", d)

conn.commit()
conn.close()
print("Table des devises créée avec succès !")