from flask import (Flask, render_template_string, request, redirect, url_for,
                   flash, send_file, jsonify, session)
from flask_login import (LoginManager, UserMixin, login_user, login_required,
                         logout_user, current_user)
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
import sqlite3, os, io
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from notifications import envoyer_alerte_expiration

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-change-me")
DB = os.getenv("DB_PATH", "offres.db")

# ---------- Login Manager ----------
login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Veuillez vous connecter pour accéder à cette page."


class User(UserMixin):
    def __init__(self, username):
        self.id = username


@login_manager.user_loader
def load_user(username):
    if username == os.getenv("ADMIN_USER", "admin"):
        return User(username)
    return None


# ---------- Base de données ----------
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


def get_stats():
    conn = get_db()
    row = conn.execute("""
        SELECT COUNT(*), SUM(CASE WHEN statut='Active' THEN 1 ELSE 0 END),
               COALESCE(SUM(CASE WHEN statut='Active' THEN prix ELSE 0 END),0)
        FROM offres""").fetchone()
    expirees = conn.execute("""
        SELECT COUNT(*) FROM offres
        WHERE date_fin != '' AND date_fin < ? AND statut='Active'
    """, (date.today().isoformat(),)).fetchone()[0]
    conn.close()
    return {"total": row[0] or 0, "actives": row[1] or 0,
            "ca": row[2] or 0, "expirees": expirees}


# ==================== TEMPLATES ====================
BASE_CSS = """
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
  body { margin: 0; background: #ecf0f1; color: #2c3e50; }
  header { background: #2c3e50; color: white; padding: 15px 30px;
           display: flex; justify-content: space-between; align-items: center; }
  header h1 { margin: 0; font-size: 22px; }
  header .user { font-size: 14px; }
  header .user a { color: #f39c12; text-decoration: none; font-weight: bold; margin-left: 10px; }
  .container { max-width: 1200px; margin: 20px auto; padding: 0 20px; }
  .stats { display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap; }
  .card { flex: 1; min-width: 180px; background: white; padding: 15px;
          border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  .card h3 { margin: 0 0 5px; color: #7f8c8d; font-size: 14px; }
  .card .value { font-size: 26px; font-weight: bold; }
  form.box { background: white; padding: 20px; border-radius: 8px;
             box-shadow: 0 2px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
  form.box .grid { display: grid; grid-template-columns: repeat(auto-fit,minmax(200px,1fr)); gap: 12px; }
  input, select { padding: 8px; border: 1px solid #bdc3c7; border-radius: 5px; width: 100%; }
  button, .btn { padding: 9px 16px; border: none; border-radius: 5px;
                 cursor: pointer; color: white; font-weight: bold; text-decoration: none;
                 display: inline-block; font-size: 14px; }
  .btn-green { background: #27ae60; } .btn-orange { background: #f39c12; }
  .btn-red { background: #e74c3c; } .btn-blue { background: #2980b9; }
  .btn-gray { background: #7f8c8d; } .btn-purple { background: #8e44ad; }
  table { width: 100%; border-collapse: collapse; background: white;
          border-radius: 8px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  th { background: #2c3e50; color: white; padding: 10px; text-align: left; }
  td { padding: 10px; border-bottom: 1px solid #ecf0f1; }
  tr.active { background: #eafaf1; } tr.expiree { background: #fdecea; }
  .actions a { margin-right: 5px; font-size: 13px; padding: 5px 10px; }
  .flash { background: #2ecc71; color: white; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
  .flash-error { background: #e74c3c; }
  .search { display: flex; gap: 10px; margin-bottom: 15px; }
  .login-box { max-width: 400px; margin: 80px auto; background: white;
               padding: 30px; border-radius: 10px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); }
  .login-box h2 { text-align: center; margin-top: 0; }
  .login-box label { display: block; margin: 12px 0 5px; font-weight: bold; }
</style>
"""

LOGIN_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Connexion</title>""" + BASE_CSS + """</head><body>
<div class="login-box">
  <h2>🔐 Connexion Admin</h2>
  {% for m in get_flashed_messages() %}<div class="flash flash-error">{{ m }}</div>{% endfor %}
  <form method="POST">
    <label>Nom d'utilisateur</label>
    <input name="username" required autofocus>
    <label>Mot de passe</label>
    <input name="password" type="password" required>
    <br><br>
    <button class="btn-green" style="width:100%">Se connecter</button>
  </form>
</div></body></html>"""


INDEX_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Gestion des Offres</title>""" + BASE_CSS + """</head><body>
<header>
  <h1>📦 Gestion des Offres</h1>
  <div class="user">👤 {{ current_user.id }} <a href="{{ url_for('logout') }}">Déconnexion</a></div>
</header>
<div class="container">

  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}

  <div class="stats">
    <div class="card"><h3>Total</h3><div class="value">{{ stats.total }}</div></div>
    <div class="card"><h3>Actives</h3><div class="value" style="color:#27ae60">{{ stats.actives }}</div></div>
    <div class="card"><h3>CA potentiel</h3><div class="value">{{ "%.2f"|format(stats.ca) }} €</div></div>
    <div class="card"><h3>Expirées</h3><div class="value" style="color:#e74c3c">{{ stats.expirees }}</div></div>
  </div>

  <form class="box" method="POST" action="{{ url_for('ajouter') }}">
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
    <button class="btn-green">{{ '💾 Enregistrer' if edit else '➕ Ajouter' }}</button>
    {% if edit %}<a href="{{ url_for('index') }}" class="btn btn-gray">Annuler</a>{% endif %}
  </form>

  <div class="search">
    <form method="GET" style="flex:1;display:flex;gap:10px;background:none;padding:0;box-shadow:none;">
      <input name="q" placeholder="🔍 Rechercher..." value="{{ q }}">
      <button class="btn-blue">Chercher</button>
      <a href="{{ url_for('index') }}" class="btn btn-gray">Tout</a>
    </form>
    <a href="{{ url_for('export_excel') }}" class="btn btn-blue">📊 Excel</a>
    <a href="{{ url_for('notifier_expirations') }}" class="btn btn-purple"
       onclick="return confirm('Envoyer les alertes par email ?')">📧 Notifier</a>
  </div>

  <table>
    <tr><th>ID</th><th>Nom</th><th>Description</th><th>Prix</th>
        <th>Début</th><th>Fin</th><th>Statut</th><th>Actions</th></tr>
    {% for o in offres %}
    <tr class="{{ 'expiree' if o.date_fin and o.date_fin < today and o.statut=='Active' else ('active' if o.statut=='Active' else '') }}">
      <td>{{ o.id }}</td><td><b>{{ o.nom }}</b></td><td>{{ o.description }}</td>
      <td>{{ "%.2f"|format(o.prix) }} €</td>
      <td>{{ o.date_debut }}</td><td>{{ o.date_fin }}</td><td>{{ o.statut }}</td>
      <td class="actions">
        <a class="btn btn-orange" href="{{ url_for('modifier', id=o.id) }}">✏️</a>
        <a class="btn btn-red" href="{{ url_for('supprimer', id=o.id) }}"
           onclick="return confirm('Supprimer ?')">🗑️</a>
      </td>
    </tr>
    {% endfor %}
  </table>
  <br>
</div></body></html>"""


# ==================== ROUTES AUTH ====================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        admin_user = os.getenv("ADMIN_USER", "admin")
        admin_pass = os.getenv("ADMIN_PASSWORD", "admin123")
        if username == admin_user and password == admin_pass:
            login_user(User(username))
            flash(f"Bienvenue, {username} !")
            return redirect(url_for("index"))
        flash("Identifiants incorrects.")
    return render_template_string(LOGIN_TEMPLATE)


@app.route("/logout")
@login_required
def logout():
    logout_user()
    flash("Déconnecté.")
    return redirect(url_for("login"))


# ==================== ROUTES CRUD ====================
@app.route("/")
@login_required
def index():
    q = request.args.get("q", "").strip()
    conn = get_db()
    if q:
        offres = conn.execute("""
            SELECT * FROM offres WHERE nom LIKE ? OR description LIKE ? OR statut LIKE ?
            ORDER BY id DESC""", (f"%{q}%",)*3).fetchall()
    else:
        offres = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    conn.close()
    return render_template_string(INDEX_TEMPLATE, offres=offres, stats=get_stats(),
                                   q=q, edit=None, today=date.today().isoformat())


@app.route("/ajouter", methods=["POST"])
@login_required
def ajouter():
    conn = get_db()
    conn.execute("""INSERT INTO offres (nom,description,prix,date_debut,date_fin,statut)
                    VALUES (?,?,?,?,?,?)""",
                 (request.form["nom"], request.form["description"],
                  float(request.form["prix"]), request.form["date_debut"],
                  request.form["date_fin"], request.form["statut"]))
    conn.commit(); conn.close()
    flash("✅ Offre ajoutée.")
    return redirect(url_for("index"))


@app.route("/modifier/<int:id>", methods=["GET", "POST"])
@login_required
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
    conn.close()
    return render_template_string(INDEX_TEMPLATE, offres=all_offres, stats=get_stats(),
                                   q="", edit=offre, today=date.today().isoformat())


@app.route("/supprimer/<int:id>")
@login_required
def supprimer(id):
    conn = get_db(); conn.execute("DELETE FROM offres WHERE id=?", (id,))
    conn.commit(); conn.close()
    flash("🗑️ Offre supprimée.")
    return redirect(url_for("index"))


# ==================== EXPORT EXCEL ====================
@app.route("/export/excel")
@login_required
def export_excel():
    conn = get_db()
    offres = conn.execute("SELECT * FROM offres ORDER BY id").fetchall()
    conn.close()
    wb = Workbook(); ws = wb.active; ws.title = "Offres"
    ws.append(["ID", "Nom", "Description", "Prix (€)", "Début", "Fin", "Statut"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2C3E50")
        c.alignment = Alignment(horizontal="center")
    for o in offres:
        ws.append(list(o))
    for i, w in enumerate([6,25,40,12,14,14,12], start=1):
        ws.column_dimensions[chr(64+i)].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"offres_{date.today()}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ==================== NOTIFICATIONS ====================
@app.route("/notifier")
@login_required
def notifier_expirations():
    nb = envoyer_alerte_expiration(DB)
    if nb > 0:
        flash(f"📧 {nb} alerte(s) email envoyée(s).")
    else:
        flash("ℹ️ Aucune offre proche de l'expiration.")
    return redirect(url_for("index"))


# ==================== API REST ====================
@app.route("/api/offres", methods=["GET"])
def api_list_offres():
    conn = get_db()
    offres = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    conn.close()
    return jsonify([dict(o) for o in offres])


@app.route("/api/offres/<int:id>", methods=["GET"])
def api_get_offre(id):
    conn = get_db()
    o = conn.execute("SELECT * FROM offres WHERE id=?", (id,)).fetchone()
    conn.close()
    if not o:
        return jsonify({"error": "Offre non trouvée"}), 404
    return jsonify(dict(o))


@app.route("/api/offres", methods=["POST"])
def api_create_offre():
    data = request.get_json()
    if not data or not data.get("nom") or data.get("prix") is None:
        return jsonify({"error": "Champs 'nom' et 'prix' requis"}), 400
    conn = get_db()
    cur = conn.execute("""INSERT INTO offres (nom,description,prix,date_debut,date_fin,statut)
                          VALUES (?,?,?,?,?,?)""",
                       (data["nom"], data.get("description", ""), float(data["prix"]),
                        data.get("date_debut", ""), data.get("date_fin", ""),
                        data.get("statut", "Active")))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return jsonify({"id": new_id, "message": "Offre créée"}), 201


@app.route("/api/offres/<int:id>", methods=["PUT"])
def api_update_offre(id):
    data = request.get_json()
    conn = get_db()
    o = conn.execute("SELECT * FROM offres WHERE id=?", (id,)).fetchone()
    if not o:
        conn.close()
        return jsonify({"error": "Offre non trouvée"}), 404
    conn.execute("""UPDATE offres SET nom=?,description=?,prix=?,date_debut=?,date_fin=?,statut=?
                    WHERE id=?""",
                 (data.get("nom", o["nom"]), data.get("description", o["description"]),
                  float(data.get("prix", o["prix"])), data.get("date_debut", o["date_debut"]),
                  data.get("date_fin", o["date_fin"]), data.get("statut", o["statut"]), id))
    conn.commit(); conn.close()
    return jsonify({"message": "Offre mise à jour"})


@app.route("/api/offres/<int:id>", methods=["DELETE"])
def api_delete_offre(id):
    conn = get_db()
    cur = conn.execute("DELETE FROM offres WHERE id=?", (id,))
    conn.commit(); conn.close()
    if cur.rowcount == 0:
        return jsonify({"error": "Offre non trouvée"}), 404
    return jsonify({"message": "Offre supprimée"})


@app.route("/api/stats", methods=["GET"])
def api_stats():
    return jsonify(get_stats())


if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5000)