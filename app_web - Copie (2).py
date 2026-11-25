from flask import (Flask, render_template_string, request, redirect, url_for,
                   flash, send_file, jsonify)
from flask_login import (LoginManager, UserMixin, login_user, login_required,
                         logout_user, current_user)
from dotenv import load_dotenv
import sqlite3, os, io, math
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from notifications import envoyer_alerte_expiration

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-change-me")
DB = os.getenv("DB_PATH", "offres.db")

login_manager = LoginManager(app)
login_manager.login_view = "login"

class User(UserMixin):
    def __init__(self, username): self.id = username

@login_manager.user_loader
def load_user(username):
    if username == os.getenv("ADMIN_USER", "admin"):
        return User(username)
    return None

# ---------- DB ----------
def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def get_devises():
    conn = get_db()
    rows = conn.execute("SELECT * FROM devises ORDER BY code").fetchall()
    conn.close()
    return rows

def get_stats():
    conn = get_db()
    rows = conn.execute("SELECT prix, devise FROM offres WHERE statut='En cours'").fetchall()
    total_row = conn.execute("SELECT COUNT(*) FROM offres").fetchone()
    actives_row = conn.execute("SELECT COUNT(*) FROM offres WHERE statut='En cours'").fetchone()
    
    expirees_row = conn.execute("""
        SELECT COUNT(*) FROM offres WHERE date_fin != '' AND date_fin < ? AND statut='En cours'
    """, (date.today().isoformat(),)).fetchone()
    
    taux_map = {d["code"]: d["taux"] for d in get_devises()}
    ca_dzd = sum(r["prix"] * taux_map.get(r["devise"], 1.0) for r in rows)
    
    conn.close()
    return {"total": total_row[0] or 0, "actives": actives_row[0] or 0,
            "ca": ca_dzd, "expirees": expirees_row[0] or 0}

# ==================== CSS ====================
BASE_CSS = """
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
  body { margin: 0; background: #ecf0f1; color: #2c3e50; }
  header { background: #2c3e50; color: white; padding: 15px 30px; display: flex; justify-content: space-between; align-items: center; }
  header h1 { margin: 0; font-size: 20px; }
  header nav a { color: #f39c12; text-decoration: none; margin-left: 15px; font-weight: bold; }
  .container { max-width: 1400px; margin: 20px auto; padding: 0 20px; }
  .stats { display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap; }
  .card { flex: 1; min-width: 180px; background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  .card h3 { margin: 0 0 5px; color: #7f8c8d; font-size: 13px; }
  .card .value { font-size: 22px; font-weight: bold; }
  form.box { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
  form.box .grid { display: grid; grid-template-columns: repeat(auto-fit,minmax(200px,1fr)); gap: 12px; }
  input, select { padding: 8px; border: 1px solid #bdc3c7; border-radius: 5px; width: 100%; }
  button, .btn { padding: 9px 16px; border: none; border-radius: 5px; cursor: pointer; color: white; font-weight: bold; text-decoration: none; display: inline-block; font-size: 14px; }
  .btn-green { background: #27ae60; } .btn-orange { background: #f39c12; }
  .btn-red { background: #e74c3c; } .btn-blue { background: #2980b9; }
  .btn-gray { background: #7f8c8d; } .btn-purple { background: #8e44ad; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  th { background: #2c3e50; color: white; padding: 10px; text-align: left; font-size: 13px; }
  td { padding: 10px; border-bottom: 1px solid #ecf0f1; font-size: 14px; }
  tr.en-cours { background: #eafaf1; } 
  tr.en-attente { background: #fef9e7; }
  tr.termine { background: #ebf5fb; }
  tr.annule { background: #fdecea; }
  .actions a { margin-right: 5px; font-size: 13px; padding: 5px 10px; }
  .flash { background: #2ecc71; color: white; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
  .badge { padding: 3px 8px; border-radius: 12px; font-size: 12px; background: #ecf0f1; }
  .progress-bar { width: 100%; background: #ecf0f1; border-radius: 5px; height: 8px; }
  .progress-fill { height: 8px; border-radius: 5px; background: #27ae60; }
  .login-box { max-width: 400px; margin: 80px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); }
  .login-box h2 { text-align: center; margin-top: 0; }
  .login-box label { display: block; margin: 12px 0 5px; font-weight: bold; }
</style>
"""

LOGIN_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Connexion - Bureau d'Études</title>""" + BASE_CSS + """</head><body>
<div class="login-box">
  <h2>🏗️ Bureau d'Études</h2>
  {% for m in get_flashed_messages() %}<div class="flash" style="background:#e74c3c;">{{ m }}</div>{% endfor %}
  <form method="POST">
    <label>Utilisateur</label><input name="username" required autofocus>
    <label>Mot de passe</label><input name="password" type="password" required>
    <br><br>
    <button class="btn-green" style="width:100%">Se connecter</button>
  </form>
</div></body></html>"""

INDEX_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Gestion des Projets - Bureau d'Études</title>
""" + BASE_CSS + """</head><body>
<header>
  <h1>🏗️ Bureau d'Études - Suivi des Projets</h1>
  <nav>
    <a href="{{ url_for('index') }}">📁 Projets</a>
    <a href="{{ url_for('devises') }}">💱 Devises</a>
    <a href="{{ url_for('logout') }}">🚪 Déconnexion ({{ current_user.id }})</a>
  </nav>
</header>
<div class="container">

  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}

  <div class="stats">
    <div class="card"><h3>Total Projets</h3><div class="value">{{ stats.total }}</div></div>
    <div class="card"><h3>En cours</h3><div class="value" style="color:#27ae60">{{ stats.actives }}</div></div>
    <div class="card"><h3>Budget Total (DZD)</h3><div class="value">{{ "{:,.2f}".format(stats.ca).replace(',', ' ') }} DZD</div></div>
    <div class="card"><h3>Deadlines dépassées</h3><div class="value" style="color:#e74c3c">{{ stats.expirees }}</div></div>
  </div>

  <form class="box" method="POST" action="{{ url_for('ajouter') }}">
    <h3 style="margin-top:0">{{ 'Modifier' if edit else 'Ajouter' }} un projet</h3>
    <div class="grid">
      <input name="nom" placeholder="Nom du projet *" required value="{{ edit.nom if edit else '' }}">
      <input name="client" placeholder="Client" value="{{ edit.client if edit else '' }}">
      <input name="chef_projet" placeholder="Chef de projet" value="{{ edit.chef_projet if edit else '' }}">
      <input name="prix" type="number" step="0.01" placeholder="Budget (DZD) *" required value="{{ edit.prix if edit else '' }}">
      <select name="devise">
        {% for d in devises_list %}
        <option value="{{ d.code }}" {% if (edit and edit.devise==d.code) or (not edit and d.code=='DZD') %}selected{% endif %}>{{ d.code }} - {{ d.nom }}</option>
        {% endfor %}
      </select>
      <input name="categorie" placeholder="Type d'étude" value="{{ edit.categorie if edit else 'Général' }}">
      <input name="date_debut" type="date" value="{{ edit.date_debut if edit else '' }}">
      <input name="date_fin" type="date" value="{{ edit.date_fin if edit else '' }}" title="Date de livraison">
      <select name="statut">
        {% for s in ['En cours','En attente','Terminé','Annulé'] %}
        <option {% if edit and edit.statut==s %}selected{% endif %}>{{ s }}</option>
        {% endfor %}
      </select>
      <input name="avancement" type="number" min="0" max="100" placeholder="Avancement %" value="{{ edit.avancement if edit else '0' }}">
    </div>
    <br>
    <button class="btn-green">{{ '💾 Enregistrer' if edit else '➕ Ajouter le projet' }}</button>
    {% if edit %}<a href="{{ url_for('index') }}" class="btn btn-gray">Annuler</a>{% endif %}
  </form>

  <table>
    <tr>
      <th>ID</th><th>Projet / Client</th><th>Type d'étude</th><th>Budget</th>
      <th>Chef de projet</th><th>Avancement</th><th>Livraison</th><th>Statut</th><th>Actions</th>
    </tr>
    {% for o in offres %}
    <tr class="{{ 'en-cours' if o.statut=='En cours' else ('en-attente' if o.statut=='En attente' else ('termine' if o.statut=='Terminé' else 'annule')) }}">
      <td>{{ o.id }}</td>
      <td><b>{{ o.nom }}</b><br><small>Client: {{ o.client }}</small></td>
      <td><span class="badge">{{ o.categorie }}</span></td>
      <td><b>{{ "{:,.2f}".format(o.prix).replace(',', ' ') }} {{ o.devise }}</b></td>
      <td>{{ o.chef_projet }}</td>
      <td style="width:100px;">
        <div class="progress-bar"><div class="progress-fill" style="width:{{ o.avancement }}%;"></div></div>
        <small>{{ o.avancement }}%</small>
      </td>
      <td>{{ o.date_fin }}</td>
      <td>{{ o.statut }}</td>
      <td class="actions">
        <a class="btn btn-orange" href="{{ url_for('modifier', id=o.id) }}">✏️</a>
        <a class="btn btn-red" href="{{ url_for('supprimer', id=o.id) }}" onclick="return confirm('Supprimer ?')">🗑️</a>
      </td>
    </tr>
    {% endfor %}
  </table>
</div>
</body></html>"""

DEVISES_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Devises - Bureau d'Études</title>""" + BASE_CSS + """</head><body>
<header>
  <h1>💱 Gestion des Devises</h1>
  <nav>
    <a href="{{ url_for('index') }}">📁 Projets</a>
    <a href="{{ url_for('devises') }}">💱 Devises</a>
    <a href="{{ url_for('logout') }}">🚪 Déconnexion ({{ current_user.id }})</a>
  </nav>
</header>
<div class="container">
  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}
  <form class="box" method="POST" action="{{ url_for('ajouter_devise') }}">
    <h3 style="margin-top:0">➕ Ajouter une devise</h3>
    <div class="grid">
      <input name="code" placeholder="Code (ex: CAD)" required maxlength="5">
      <input name="nom" placeholder="Nom (ex: Dollar Canadien)" required>
      <input name="taux" type="number" step="0.01" placeholder="Taux en DZD" required>
    </div>
    <br><button class="btn-green">➕ Ajouter</button>
  </form>
  <table>
    <tr><th>Code</th><th>Nom</th><th>Taux (1 unité = X DZD)</th><th>Actions</th></tr>
    {% for d in devises %}
    <tr>
      <td><span class="badge" style="background:#3498db;color:white;">{{ d.code }}</span></td>
      <td>{{ d.nom }}</td>
      <td>
        {% if d.code == 'DZD' %}<b>1.00</b> <small>(référence)</small>
        {% else %}
          <form method="POST" action="{{ url_for('modifier_devise', code=d.code) }}" style="display:flex;gap:5px;margin:0;background:none;padding:0;box-shadow:none;">
            <input name="taux" type="number" step="0.01" value="{{ d.taux }}" style="width:120px;padding:5px;">
            <button class="btn-green" style="padding:5px 10px;">💾</button>
          </form>
        {% endif %}
      </td>
      <td>
        {% if d.code != 'DZD' %}
          <a class="btn btn-red" href="{{ url_for('supprimer_devise', code=d.code) }}" onclick="return confirm('Supprimer ?')">🗑️</a>
        {% else %}<span class="badge">🔒</span>{% endif %}
      </td>
    </tr>
    {% endfor %}
  </table>
</div></body></html>"""

# ==================== ROUTES ====================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        p = request.form.get("password", "")
        if u == os.getenv("ADMIN_USER", "admin") and p == os.getenv("ADMIN_PASSWORD", "admin123"):
            login_user(User(u))
            return redirect(url_for("index"))
        flash("Identifiants incorrects.")
    return render_template_string(LOGIN_TEMPLATE)

@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    q = request.args.get("q", "").strip()
    conn = get_db()
    query = "SELECT * FROM offres WHERE 1=1"
    params = []
    if q:
        query += " AND (nom LIKE ? OR client LIKE ? OR statut LIKE ? OR categorie LIKE ?)"
        params.extend([f"%{q}%"] * 4)
    query += " ORDER BY id DESC"
    all_rows = conn.execute(query, params).fetchall()
    conn.close()

    return render_template_string(INDEX_TEMPLATE, offres=all_rows, stats=get_stats(),
                                   q=q, edit=None, today=date.today().isoformat(),
                                   devises_list=get_devises())

@app.route("/ajouter", methods=["POST"])
@login_required
def ajouter():
    conn = get_db()
    conn.execute("""INSERT INTO offres (nom,description,prix,date_debut,date_fin,statut,categorie,devise,client,chef_projet,avancement)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                 (request.form["nom"], "", float(request.form["prix"]), 
                  request.form["date_debut"], request.form["date_fin"], 
                  request.form["statut"], request.form.get("categorie", "Général"),
                  request.form.get("devise", "DZD"), request.form.get("client", ""),
                  request.form.get("chef_projet", ""), int(request.form.get("avancement", 0))))
    conn.commit(); conn.close()
    flash("✅ Projet ajouté.")
    return redirect(url_for("index"))

@app.route("/modifier/<int:id>", methods=["GET", "POST"])
@login_required
def modifier(id):
    conn = get_db()
    if request.method == "POST":
        conn.execute("""UPDATE offres SET nom=?,prix=?,date_debut=?,date_fin=?,statut=?,categorie=?,devise=?,client=?,chef_projet=?,avancement=?
                        WHERE id=?""",
                     (request.form["nom"], float(request.form["prix"]), 
                      request.form["date_debut"], request.form["date_fin"], 
                      request.form["statut"], request.form.get("categorie", "Général"),
                      request.form.get("devise", "DZD"), request.form.get("client", ""),
                      request.form.get("chef_projet", ""), int(request.form.get("avancement", 0)), id))
        conn.commit(); conn.close()
        flash("✏️ Projet modifié.")
        return redirect(url_for("index"))
    
    offre = conn.execute("SELECT * FROM offres WHERE id=?", (id,)).fetchone()
    all_rows = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    conn.close()
    return render_template_string(INDEX_TEMPLATE, offres=all_rows, stats=get_stats(),
                                   q="", edit=offre, today=date.today().isoformat(),
                                   devises_list=get_devises())

@app.route("/supprimer/<int:id>")
@login_required
def supprimer(id):
    conn = get_db(); conn.execute("DELETE FROM offres WHERE id=?", (id,))
    conn.commit(); conn.close()
    flash("🗑️ Projet supprimé.")
    return redirect(url_for("index"))

@app.route("/devises")
@login_required
def devises():
    return render_template_string(DEVISES_TEMPLATE, devises=get_devises())

@app.route("/devises/ajouter", methods=["POST"])
@login_required
def ajouter_devise():
    code = request.form["code"].strip().upper()
    nom = request.form["nom"].strip()
    taux = float(request.form["taux"])
    conn = get_db()
    try:
        conn.execute("INSERT INTO devises (code, nom, taux) VALUES (?,?,?)", (code, nom, taux))
        conn.commit(); flash(f"✅ Devise {code} ajoutée.")
    except sqlite3.IntegrityError: flash(f"❌ La devise existe déjà.")
    conn.close()
    return redirect(url_for("devises"))

@app.route("/devises/modifier/<code>", methods=["POST"])
@login_required
def modifier_devise(code):
    conn = get_db()
    conn.execute("UPDATE devises SET taux=? WHERE code=?", (float(request.form["taux"]), code))
    conn.commit(); conn.close()
    return redirect(url_for("devises"))

@app.route("/devises/supprimer/<code>")
@login_required
def supprimer_devise(code):
    if code == "DZD": return redirect(url_for("devises"))
    conn = get_db(); conn.execute("DELETE FROM devises WHERE code=?", (code,))
    conn.commit(); conn.close()
    return redirect(url_for("devises"))

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)