from flask import Flask, render_template_string, request, redirect, url_for, flash, send_file
import sqlite3, os, io
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.secret_key = "cle_secrete_change_me"
DB = "offres.db"


# ---------- Initialisation ----------
def init_db():
    conn = sqlite3.connect(DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS offres (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT NOT NULL, description TEXT, prix REAL NOT NULL,
            date_debut TEXT, date_fin TEXT, statut TEXT DEFAULT 'Active'
        )
    """)
    conn.commit(); conn.close()


def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


# ---------- Template HTML ----------
TEMPLATE = """
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Gestion des Offres</title>
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
  body { margin: 0; background: #ecf0f1; color: #2c3e50; }
  header { background: #2c3e50; color: white; padding: 15px 30px; }
  header h1 { margin: 0; }
  .container { max-width: 1200px; margin: 20px auto; padding: 0 20px; }
  .stats { display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap; }
  .card { flex: 1; min-width: 180px; background: white; padding: 15px;
          border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  .card h3 { margin: 0 0 5px; color: #7f8c8d; font-size: 14px; }
  .card .value { font-size: 26px; font-weight: bold; }
  form { background: white; padding: 20px; border-radius: 8px;
         box-shadow: 0 2px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
  form .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px,1fr)); gap: 12px; }
  input, select { padding: 8px; border: 1px solid #bdc3c7; border-radius: 5px; width: 100%; }
  button, .btn { padding: 9px 16px; border: none; border-radius: 5px;
                 cursor: pointer; color: white; font-weight: bold; text-decoration: none;
                 display: inline-block; font-size: 14px; }
  .btn-green { background: #27ae60; } .btn-orange { background: #f39c12; }
  .btn-red { background: #e74c3c; } .btn-blue { background: #2980b9; }
  .btn-gray { background: #7f8c8d; }
  table { width: 100%; border-collapse: collapse; background: white;
          border-radius: 8px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  th { background: #2c3e50; color: white; padding: 10px; text-align: left; }
  td { padding: 10px; border-bottom: 1px solid #ecf0f1; }
  tr.active { background: #eafaf1; } tr.expiree { background: #fdecea; }
  .actions a { margin-right: 5px; font-size: 13px; padding: 5px 10px; }
  .search { display: flex; gap: 10px; margin-bottom: 15px; }
  .flash { background: #2ecc71; color: white; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
</style>
</head>
<body>
<header><h1>📦 Gestion des Offres — Web</h1></header>
<div class="container">

  {% with messages = get_flashed_messages() %}
    {% for m in messages %}<div class="flash">{{ m }}</div>{% endfor %}
  {% endwith %}

  <div class="stats">
    <div class="card"><h3>Total offres</h3><div class="value">{{ stats.total }}</div></div>
    <div class="card"><h3>Actives</h3><div class="value" style="color:#27ae60">{{ stats.actives }}</div></div>
    <div class="card"><h3>CA potentiel</h3><div class="value">{{ "%.2f"|format(stats.ca) }} €</div></div>
    <div class="card"><h3>Expirées</h3><div class="value" style="color:#e74c3c">{{ stats.expirees }}</div></div>
  </div>

  <form method="POST" action="{{ url_for('ajouter') }}">
    <h3 style="margin-top:0">{{ 'Modifier' if edit else 'Ajouter' }} une offre</h3>
    <div class="grid">
      <input name="nom" placeholder="Nom *" required value="{{ edit.nom if edit else '' }}">
      <input name="description" placeholder="Description" value="{{ edit.description if edit else '' }}">
      <input name="prix" type="number" step="0.01" placeholder="Prix € *" required value="{{ edit.prix if edit else '' }}">
      <input name="date_debut" type="date" value="{{ edit.date_debut if edit else '' }}">
      <input name="date_fin" type="date" value="{{ edit.date_fin if edit else '' }}">
      <select name="statut">
        {% for s in ['Active','Inactive','Expirée'] %}
        <option {% if edit and edit.statut==s %}selected{% endif %}>{{ s }}</option>
        {% endfor %}
      </select>
    </div>
    <br>
    <button type="submit" class="btn-green">{{ '💾 Enregistrer' if edit else '➕ Ajouter' }}</button>
    {% if edit %}<a href="{{ url_for('index') }}" class="btn btn-gray">Annuler</a>{% endif %}
  </form>

  <div class="search">
    <form method="GET" style="flex:1; display:flex; gap:10px; background:none; padding:0; box-shadow:none;">
      <input name="q" placeholder="🔍 Rechercher..." value="{{ q }}">
      <button class="btn-blue" type="submit">Chercher</button>
      <a href="{{ url_for('index') }}" class="btn btn-gray">Tout</a>
    </form>
    <a href="{{ url_for('export_excel') }}" class="btn btn-blue">📊 Excel</a>
  </div>

  <table>
    <tr><th>ID</th><th>Nom</th><th>Description</th><th>Prix</th>
        <th>Début</th><th>Fin</th><th>Statut</th><th>Actions</th></tr>
    {% for o in offres %}
    <tr class="{{ 'expiree' if o.date_fin and o.date_fin < today and o.statut=='Active' else ('active' if o.statut=='Active' else '') }}">
      <td>{{ o.id }}</td>
      <td><b>{{ o.nom }}</b></td>
      <td>{{ o.description }}</td>
      <td>{{ "%.2f"|format(o.prix) }} €</td>
      <td>{{ o.date_debut }}</td>
      <td>{{ o.date_fin }}</td>
      <td>{{ o.statut }}</td>
      <td class="actions">
        <a class="btn btn-orange" href="{{ url_for('modifier', id=o.id) }}">✏️</a>
        <a class="btn btn-red" href="{{ url_for('supprimer', id=o.id) }}"
           onclick="return confirm('Supprimer ?')">🗑️</a>
      </td>
    </tr>
    {% endfor %}
  </table>
  <br>
</div>
</body>
</html>
"""


# ---------- Routes ----------
@app.route("/")
def index():
    q = request.args.get("q", "").strip()
    conn = get_db()
    if q:
        offres = conn.execute("""
            SELECT * FROM offres WHERE nom LIKE ? OR description LIKE ? OR statut LIKE ?
            ORDER BY id DESC""", (f"%{q}%",)*3).fetchall()
    else:
        offres = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()

    stats_row = conn.execute("""
        SELECT COUNT(*) as total,
               SUM(CASE WHEN statut='Active' THEN 1 ELSE 0 END) as actives,
               COALESCE(SUM(CASE WHEN statut='Active' THEN prix ELSE 0 END),0) as ca
        FROM offres""").fetchone()
    expirees = conn.execute("""
        SELECT COUNT(*) FROM offres
        WHERE date_fin != '' AND date_fin < ? AND statut='Active'
    """, (date.today().isoformat(),)).fetchone()[0]
    conn.close()

    stats = {"total": stats_row[0] or 0, "actives": stats_row[1] or 0,
             "ca": stats_row[2] or 0, "expirees": expirees}

    return render_template_string(TEMPLATE, offres=offres, stats=stats,
                                   q=q, edit=None, today=date.today().isoformat())


@app.route("/ajouter", methods=["POST"])
def ajouter():
    conn = get_db()
    conn.execute("INSERT INTO offres (nom,description,prix,date_debut,date_fin,statut) VALUES (?,?,?,?,?,?)",
                 (request.form["nom"], request.form["description"],
                  float(request.form["prix"]), request.form["date_debut"],
                  request.form["date_fin"], request.form["statut"]))
    conn.commit(); conn.close()
    flash("✅ Offre ajoutée avec succès.")
    return redirect(url_for("index"))


@app.route("/modifier/<int:id>", methods=["GET", "POST"])
def modifier(id):
    conn = get_db()
    if request.method == "POST":
        conn.execute("""UPDATE offres SET nom=?,description=?,prix=?,date_debut=?,date_fin=?,statut=?
                        WHERE id=?""",
                     (request.form["nom"], request.form["description"],
                      float(request.form["prix"]), request.form["date_debut"],
                      request.form["date_fin"], request.form["statut"], id))
        conn.commit(); conn.close()
        flash("✏️ Offre modifiée.")
        return redirect(url_for("index"))

    offre = conn.execute("SELECT * FROM offres WHERE id=?", (id,)).fetchone()
    all_offres = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    stats_row = conn.execute("""
        SELECT COUNT(*), SUM(CASE WHEN statut='Active' THEN 1 ELSE 0 END),
               COALESCE(SUM(CASE WHEN statut='Active' THEN prix ELSE 0 END),0) FROM offres""").fetchone()
    stats = {"total": stats_row[0] or 0, "actives": stats_row[1] or 0,
             "ca": stats_row[2] or 0, "expirees": 0}
    conn.close()
    return render_template_string(TEMPLATE, offres=all_offres, stats=stats,
                                   q="", edit=offre, today=date.today().isoformat())


@app.route("/supprimer/<int:id>")
def supprimer(id):
    conn = get_db()
    conn.execute("DELETE FROM offres WHERE id=?", (id,))
    conn.commit(); conn.close()
    flash("🗑️ Offre supprimée.")
    return redirect(url_for("index"))


@app.route("/export/excel")
def export_excel():
    conn = get_db()
    offres = conn.execute("SELECT * FROM offres ORDER BY id").fetchall()
    conn.close()

    wb = Workbook(); ws = wb.active; ws.title = "Offres"
    ws.append(["ID", "Nom", "Description", "Prix (€)", "Date début", "Date fin", "Statut"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2C3E50")
        c.alignment = Alignment(horizontal="center")
    for o in offres:
        ws.append(list(o))
    for i, w in enumerate([6, 25, 40, 12, 14, 14, 12], start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"offres_{date.today()}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    init_db()
    app.run(debug=True, port=5000)