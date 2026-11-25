import sqlite3

conn = sqlite3.connect('offres.db')
cursor = conn.cursor()

# Ajout des colonnes
try:
    cursor.execute("ALTER TABLE offres ADD COLUMN client TEXT DEFAULT 'Inconnu'")
    print("Colonne 'client' ajoutée.")
except sqlite3.OperationalError:
    print("La colonne 'client' existe déjà.")

try:
    cursor.execute("ALTER TABLE offres ADD COLUMN chef_projet TEXT DEFAULT 'Non assigne'")
    print("Colonne 'chef_projet' ajoutée.")
except sqlite3.OperationalError:
    print("La colonne 'chef_projet' existe déjà.")

try:
    cursor.execute("ALTER TABLE offres ADD COLUMN avancement INTEGER DEFAULT 0")
    print("Colonne 'avancement' ajoutée.")
except sqlite3.OperationalError:
    print("La colonne 'avancement' existe déjà.")

conn.commit()
conn.close()
print("Mise à jour terminée avec succès !")