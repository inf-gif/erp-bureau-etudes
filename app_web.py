from flask import (Flask, render_template_string, request, redirect, url_for,
                   flash, send_file, jsonify)
from flask_login import (LoginManager, UserMixin, login_user, login_required,
                         logout_user, current_user)
from dotenv import load_dotenv
import sqlite3, os, io
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

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
    ao_total = conn.execute("SELECT COUNT(*) FROM soumissions").fetchone()[0]
    ao_gagnees = conn.execute("SELECT COUNT(*) FROM soumissions WHERE statut='Gagnée'").fetchone()[0]
    projets_actifs = conn.execute("SELECT COUNT(*) FROM offres WHERE statut='En cours'").fetchone()[0]
    taches_en_cours = conn.execute("SELECT COUNT(*) FROM taches WHERE statut='En cours'").fetchone()[0]
    rows = conn.execute("SELECT prix, devise FROM offres WHERE statut='En cours'").fetchall()
    taux_map = {d["code"]: d["taux"] for d in get_devises()}
    budget_dzd = sum(r["prix"] * taux_map.get(r["devise"], 1.0) for r in rows)
    conn.close()
    return {"ao_total": ao_total, "ao_gagnees": ao_gagnees, "projets_actifs": projets_actifs, "budget": budget_dzd, "taches_en_cours": taches_en_cours}

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
  input, select, textarea { padding: 8px; border: 1px solid #bdc3c7; border-radius: 5px; width: 100%; }
  button, .btn { padding: 9px 16px; border: none; border-radius: 5px; cursor: pointer; color: white; font-weight: bold; text-decoration: none; display: inline-block; font-size: 14px; }
  .btn-green { background: #27ae60; } .btn-orange { background: #f39c12; }
  .btn-red { background: #e74c3c; } .btn-blue { background: #2980b9; }
  .btn-gray { background: #7f8c8d; } .btn-purple { background: #8e44ad; }
  .btn-dark { background: #2c3e50; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  th { background: #2c3e50; color: white; padding: 10px; text-align: left; font-size: 13px; }
  td { padding: 10px; border-bottom: 1px solid #ecf0f1; font-size: 14px; }
  .badge { padding: 3px 8px; border-radius: 12px; font-size: 12px; background: #ecf0f1; }
  .progress-bar { width: 100%; background: #ecf0f1; border-radius: 5px; height: 8px; }
  .progress-fill { height: 8px; border-radius: 5px; background: #27ae60; }
  .flash { background: #2ecc71; color: white; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
  .nav-tabs { display: flex; margin-bottom: 20px; border-bottom: 2px solid #bdc3c7; }
  .nav-tabs a { padding: 10px 20px; text-decoration: none; color: #7f8c8d; font-weight: bold; border-bottom: 3px solid transparent; }
  .nav-tabs a.active { color: #2c3e50; border-bottom-color: #2980b9; }
  .login-box { max-width: 400px; margin: 80px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); }
  .login-box h2 { text-align: center; margin-top: 0; }
  .login-box label { display: block; margin: 12px 0 5px; font-weight: bold; }
</style>
"""

# ==================== TEMPLATES ====================
NAV_MENU = """<nav>
  <a href="{{ url_for('index') }}">📁 Projets</a>
  <a href="{{ url_for('soumissions') }}">📝 Soumissions</a>
  <a href="{{ url_for('employes') }}">👥 Employés</a>
  <a href="{{ url_for('taches') }}">📋 Tâches</a>
  <a href="{{ url_for('devises') }}">💱 Devises</a>
  <a href="{{ url_for('logout') }}">🚪 Déconnexion</a>
</nav>"""

LOGIN_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Connexion - ERP Bureau d'Études</title>""" + BASE_CSS + """</head><body>
<div class="login-box">
  <h2>🏗️ ERP Bureau d'Études</h2>
  {% for m in get_flashed_messages() %}<div class="flash" style="background:#e74c3c;">{{ m }}</div>{% endfor %}
  <form method="POST">
    <label>Utilisateur</label><input name="username" required autofocus>
    <label>Mot de passe</label><input name="password" type="password" required>
    <br><br>
    <button class="btn-green" style="width:100%">Se connecter</button>
  </form>
</div></body></html>"""

SOUMISSIONS_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Soumissions - ERP</title>""" + BASE_CSS + """</head><body>
<header><h1>🏗️ ERP Bureau d'Études</h1>""" + NAV_MENU + """</header>
<div class="container">
  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}
  <form class="box" method="POST" action="{{ url_for('ajouter_soumission') }}">
    <h3 style="margin-top:0">➕ Nouvelle soumission / Appel d'offres</h3>
    <div class="grid">
      <input name="reference_ao" placeholder="Réf. Appel d'offres (ex: AO-2026-01)">
      <input name="client" placeholder="Client *" required>
      <input name="objet" placeholder="Objet de l'étude *" required>
      <input name="montant_estime" type="number" step="0.01" placeholder="Montant estimé *" required>
      <select name="devise">
        {% for d in devises_list %}<option value="{{ d.code }}" {% if d.code=='DZD' %}selected{% endif %}>{{ d.code }}</option>{% endfor %}
      </select>
      <input name="date_limite" type="date" title="Date limite de remise">
      <select name="statut">
        <option>En préparation</option><option>Remise</option><option>Gagnée</option><option>Perdue</option><option>Abandonnée</option>
      </select>
    </div>
    <br><button class="btn-blue">📝 Enregistrer la soumission</button>
  </form>
  <table>
    <tr><th>ID</th><th>Réf. AO</th><th>Client</th><th>Objet</th><th>Montant</th><th>Date limite</th><th>Statut</th><th>Actions</th></tr>
    {% for s in soumissions %}
    <tr>
      <td>{{ s.id }}</td><td><b>{{ s.reference_ao }}</b></td><td>{{ s.client }}</td><td>{{ s.objet }}</td>
      <td>{{ "{:,.2f}".format(s.montant_estime).replace(',', ' ') }} {{ s.devise }}</td>
      <td>{{ s.date_limite }}</td>
      <td><span class="badge" style="background:#f39c12;color:white;">{{ s.statut }}</span></td>
      <td>
        <a href="{{ url_for('generer_devis_pdf', id=s.id) }}" class="btn btn-purple" target="_blank">📄 Devis PDF</a>
        {% if s.statut == 'Gagnée' %}
          <a href="{{ url_for('creer_projet_depuis_soumission', id=s.id) }}" class="btn btn-green">Créer Projet ➔</a>
        {% endif %}
      </td>
    </tr>
    {% endfor %}
  </table>
</div></body></html>"""

PROJETS_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Projets - ERP</title>""" + BASE_CSS + """</head><body>
<header><h1>🏗️ ERP Bureau d'Études</h1>""" + NAV_MENU + """</header>
<div class="container">
  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}

  {% if edit %}
  <form class="box" method="POST" action="{{ url_for('modifier', id=edit.id) }}">
    <h3 style="margin-top:0; color: #e67e22;">✏️ Modifier le projet : {{ edit.nom }}</h3>
    <div class="grid">
      <input name="nom" placeholder="Nom du projet *" required value="{{ edit.nom }}">
      <input name="client" placeholder="Client" value="{{ edit.client }}">
      <input name="chef_projet" placeholder="Chef de projet" value="{{ edit.chef_projet }}">
      <input name="prix" type="number" step="0.01" placeholder="Budget *" required value="{{ edit.prix }}">
      <select name="devise">
        {% for d in devises_list %}
        <option value="{{ d.code }}" {% if edit.devise==d.code %}selected{% endif %}>{{ d.code }}</option>
        {% endfor %}
      </select>
      <input name="categorie" placeholder="Type d'étude" value="{{ edit.categorie }}">
      <input name="date_debut" type="date" value="{{ edit.date_debut }}">
      <input name="date_fin" type="date" value="{{ edit.date_fin }}">
      <select name="statut">
        {% for s in ['En cours','En attente','Terminé','Annulé'] %}
        <option {% if edit.statut==s %}selected{% endif %}>{{ s }}</option>
        {% endfor %}
      </select>
      <input name="avancement" type="number" min="0" max="100" placeholder="Avancement %" value="{{ edit.avancement }}">
    </div>
    <br>
    <button class="btn btn-green">💾 Enregistrer les modifications</button>
    <a href="{{ url_for('index') }}" class="btn btn-gray">Annuler</a>
  </form>
  {% endif %}

  <div class="stats">
    <div class="card"><h3>Projets Actifs</h3><div class="value" style="color:#27ae60">{{ stats.projets_actifs }}</div></div>
    <div class="card"><h3>Soumissions Gagnées</h3><div class="value" style="color:#2980b9">{{ stats.ao_gagnees }}</div></div>
    <div class="card"><h3>Tâches en cours</h3><div class="value" style="color:#f39c12">{{ stats.taches_en_cours }}</div></div>
    <div class="card"><h3>Budget Total Projets (DZD)</h3><div class="value">{{ "{:,.2f}".format(stats.budget).replace(',', ' ') }} DZD</div></div>
  </div>

  <table>
    <tr><th>ID</th><th>Projet / Client</th><th>Type</th><th>Budget</th><th>Chef</th><th>Avancement</th><th>Livraison</th><th>Statut</th><th>Actions</th></tr>
    {% for o in offres %}
    <tr>
      <td>{{ o.id }}</td>
      <td><b>{{ o.nom }}</b><br><small>Client: {{ o.client }}</small></td>
      <td><span class="badge">{{ o.categorie }}</span></td>
      <td><b>{{ "{:,.2f}".format(o.prix).replace(',', ' ') }} {{ o.devise }}</b></td>
      <td>{{ o.chef_projet }}</td>
      <td style="width:100px;"><div class="progress-bar"><div class="progress-fill" style="width:{{ o.avancement }}%;"></div></div><small>{{ o.avancement }}%</small></td>
      <td>{{ o.date_fin }}</td>
      <td>{{ o.statut }}</td>
      <td class="actions"><a class="btn btn-orange" href="{{ url_for('modifier', id=o.id) }}">✏️</a></td>
    </tr>
    {% endfor %}
  </table>
</div></body></html>"""

DEVISES_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Devises - ERP</title>""" + BASE_CSS + """</head><body>
<header><h1>🏗️ ERP Bureau d'Études</h1>""" + NAV_MENU + """</header>
<div class="container">
  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}
  <form class="box" method="POST" action="{{ url_for('ajouter_devise') }}">
    <h3 style="margin-top:0">➕ Ajouter une devise</h3>
    <div class="grid"><input name="code" placeholder="Code (ex: CAD)" required maxlength="5"><input name="nom" placeholder="Nom" required><input name="taux" type="number" step="0.01" placeholder="Taux en DZD" required></div>
    <br><button class="btn-green">➕ Ajouter</button>
  </form>
  <table>
    <tr><th>Code</th><th>Nom</th><th>Taux (1 = X DZD)</th><th>Actions</th></tr>
    {% for d in devises %}
    <tr><td><span class="badge" style="background:#3498db;color:white;">{{ d.code }}</span></td><td>{{ d.nom }}</td>
      <td>{% if d.code == 'DZD' %}<b>1.00</b>{% else %}<form method="POST" action="{{ url_for('modifier_devise', code=d.code) }}" style="display:flex;gap:5px;"><input name="taux" type="number" step="0.01" value="{{ d.taux }}" style="width:120px;"><button class="btn-green">💾</button></form>{% endif %}</td>
      <td>{% if d.code != 'DZD' %}<a class="btn btn-red" href="{{ url_for('supprimer_devise', code=d.code) }}">🗑️</a>{% else %}🔒{% endif %}</td>
    </tr>
    {% endfor %}
  </table>
</div></body></html>"""

EMPLOYES_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Employés - ERP</title>""" + BASE_CSS + """</head><body>
<header><h1>🏗️ ERP Bureau d'Études</h1>""" + NAV_MENU + """</header>
<div class="container">
  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}
  <form class="box" method="POST" action="{{ url_for('ajouter_employe') }}">
    <h3 style="margin-top:0">➕ Ajouter un employé</h3>
    <div class="grid">
        <input name="nom" placeholder="Nom complet *" required>
        <input name="role" placeholder="Rôle (ex: Ingénieur)">
        <input name="email" type="email" placeholder="Email">
        <input name="telephone" placeholder="Téléphone">
    </div>
    <br><button class="btn-green">➕ Ajouter</button>
  </form>
  <table>
    <tr><th>ID</th><th>Nom</th><th>Rôle</th><th>Email</th><th>Téléphone</th></tr>
    {% for e in employes %}
    <tr><td>{{ e.id }}</td><td><b>{{ e.nom }}</b></td><td>{{ e.role }}</td><td>{{ e.email }}</td><td>{{ e.telephone }}</td></tr>
    {% endfor %}
  </table>
</div></body></html>"""

TACHES_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Tâches - ERP</title>""" + BASE_CSS + """</head><body>
<header><h1>🏗️ ERP Bureau d'Études</h1>""" + NAV_MENU + """</header>
<div class="container">
  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}
  <form class="box" method="POST" action="{{ url_for('ajouter_tache') }}">
    <h3 style="margin-top:0">➕ Assigner une tâche</h3>
    <div class="grid">
        <select name="projet_id" required>
            <option value="">-- Projet --</option>
            {% for p in projets %}<option value="{{ p.id }}">{{ p.nom }}</option>{% endfor %}
        </select>
        <select name="employe_id" required>
            <option value="">-- Employé --</option>
            {% for e in employes %}<option value="{{ e.id }}">{{ e.nom }}</option>{% endfor %}
        </select>
        <input name="description" placeholder="Description de la tâche *" required>
        <input name="date_echeance" type="date" required>
        <select name="statut"><option>À faire</option><option>En cours</option><option>Terminé</option></select>
    </div>
    <br><button class="btn-blue">➕ Assigner</button>
  </form>
  <table>
    <tr><th>ID</th><th>Projet</th><th>Employé</th><th>Tâche</th><th>Échéance</th><th>Statut</th></tr>
    {% for t in taches %}
    <tr>
        <td>{{ t.id }}</td><td>{{ t.projet_nom }}</td><td>{{ t.employe_nom }}</td>
        <td>{{ t.description }}</td><td>{{ t.date_echeance }}</td>
        <td><span class="badge" style="background:{% if t.statut == 'Terminé' %}#27ae60{% elif t.statut == 'En cours' %}#f39c12{% else %}#bdc3c7{% endif %};color:white;">{{ t.statut }}</span></td>
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
    conn = get_db()
    offres = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    conn.close()
    return render_template_string(PROJETS_TEMPLATE, offres=offres, stats=get_stats(), devises_list=get_devises(), edit=None)

@app.route("/soumissions")
@login_required
def soumissions():
    conn = get_db()
    soum = conn.execute("SELECT * FROM soumissions ORDER BY id DESC").fetchall()
    conn.close()
    return render_template_string(SOUMISSIONS_TEMPLATE, soumissions=soum, stats=get_stats(), devises_list=get_devises())

@app.route("/soumissions/ajouter", methods=["POST"])
@login_required
def ajouter_soumission():
    conn = get_db()
    conn.execute("""INSERT INTO soumissions (reference_ao, client, objet, montant_estime, devise, date_limite, statut)
                    VALUES (?,?,?,?,?,?,?)""",
                 (request.form["reference_ao"], request.form["client"], request.form["objet"],
                  float(request.form["montant_estime"]), request.form["devise"], request.form["date_limite"], request.form["statut"]))
    conn.commit(); conn.close()
    flash("✅ Soumission enregistrée.")
    return redirect(url_for("soumissions"))

@app.route("/soumissions/creer_projet/<int:id>")
@login_required
def creer_projet_depuis_soumission(id):
    conn = get_db()
    s = conn.execute("SELECT * FROM soumissions WHERE id=?", (id,)).fetchone()
    if s:
        conn.execute("""INSERT INTO offres (nom, description, prix, date_debut, date_fin, statut, categorie, devise, client, chef_projet, avancement, soumission_id)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (s["objet"], f"Issu de l'AO: {s['reference_ao']}", s["montant_estime"], date.today().isoformat(), "", "En cours", "Étude", s["devise"], s["client"], "À assigner", 0, s["id"]))
        conn.commit()
        flash("✅ Projet créé avec succès à partir de la soumission !")
    conn.close()
    return redirect(url_for("index"))

@app.route("/ajouter", methods=["POST"])
@login_required
def ajouter():
    conn = get_db()
    conn.execute("""INSERT INTO offres (nom,description,prix,date_debut,date_fin,statut,categorie,devise,client,chef_projet,avancement)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                 (request.form["nom"], "", float(request.form["prix"]), request.form["date_debut"], request.form["date_fin"], request.form["statut"], request.form.get("categorie", "Général"), request.form.get("devise", "DZD"), request.form.get("client", ""), request.form.get("chef_projet", ""), int(request.form.get("avancement", 0))))
    conn.commit(); conn.close()
    return redirect(url_for("index"))

@app.route("/modifier/<int:id>", methods=["GET", "POST"])
@login_required
def modifier(id):
    conn = get_db()
    if request.method == "POST":
        conn.execute("""UPDATE offres SET nom=?,prix=?,date_debut=?,date_fin=?,statut=?,categorie=?,devise=?,client=?,chef_projet=?,avancement=? WHERE id=?""",
                     (request.form["nom"], float(request.form["prix"]), request.form["date_debut"], 
                      request.form["date_fin"], request.form["statut"], request.form.get("categorie", "Général"), 
                      request.form.get("devise", "DZD"), request.form.get("client", ""), 
                      request.form.get("chef_projet", ""), int(request.form.get("avancement", 0)), id))
        conn.commit(); conn.close()
        flash("✏️ Projet modifié avec succès.")
        return redirect(url_for("index"))
    
    offre = conn.execute("SELECT * FROM offres WHERE id=?", (id,)).fetchone()
    all_rows = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    conn.close()
    return render_template_string(PROJETS_TEMPLATE, offres=all_rows, stats=get_stats(), 
                                   edit=offre, devises_list=get_devises())

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

@app.route("/employes")
@login_required
def employes():
    conn = get_db()
    emps = conn.execute("SELECT * FROM employes ORDER BY nom").fetchall()
    conn.close()
    return render_template_string(EMPLOYES_TEMPLATE, employes=emps)

@app.route("/employes/ajouter", methods=["POST"])
@login_required
def ajouter_employe():
    conn = get_db()
    conn.execute("INSERT INTO employes (nom, role, email, telephone) VALUES (?,?,?,?)",
                 (request.form["nom"], request.form["role"], request.form["email"], request.form["telephone"]))
    conn.commit(); conn.close()
    flash("✅ Employé ajouté.")
    return redirect(url_for("employes"))

@app.route("/taches")
@login_required
def taches():
    conn = get_db()
    t = conn.execute("""SELECT t.*, o.nom as projet_nom, e.nom as employe_nom 
                        FROM taches t 
                        LEFT JOIN offres o ON t.projet_id = o.id 
                        LEFT JOIN employes e ON t.employe_id = e.id 
                        ORDER BY t.id DESC""").fetchall()
    projets = conn.execute("SELECT id, nom FROM offres").fetchall()
    emps = conn.execute("SELECT id, nom FROM employes").fetchall()
    conn.close()
    return render_template_string(TACHES_TEMPLATE, taches=t, projets=projets, employes=emps)

@app.route("/taches/ajouter", methods=["POST"])
@login_required
def ajouter_tache():
    conn = get_db()
    conn.execute("INSERT INTO taches (projet_id, employe_id, description, statut, date_echeance) VALUES (?,?,?,?,?)",
                 (request.form["projet_id"], request.form["employe_id"], request.form["description"], request.form["statut"], request.form["date_echeance"]))
    conn.commit(); conn.close()
    flash("✅ Tâche ajoutée.")
    return redirect(url_for("taches"))

@app.route("/export/excel")
@login_required
def export_excel():
    conn = get_db()
    offres = conn.execute("SELECT * FROM offres ORDER BY id").fetchall()
    conn.close()
    wb = Workbook(); ws = wb.active; ws.title = "Projets"
    ws.append(["ID", "Projet", "Client", "Type", "Budget", "Devise", "Chef", "Avancement", "Livraison", "Statut"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2C3E50")
        c.alignment = Alignment(horizontal="center")
    for o in offres:
        ws.append([o["id"], o["nom"], o["client"], o["categorie"], o["prix"], o["devise"], o["chef_projet"], f"{o['avancement']}%", o["date_fin"], o["statut"]])
    for i, w in enumerate([6, 25, 20, 15, 15, 8, 20, 12, 14, 12], start=1):
        ws.column_dimensions[chr(64+i)].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"projets_{date.today()}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/notifier")
@login_required
def notifier_expirations():
    nb = envoyer_alerte_expiration(DB)
    if nb > 0: flash(f"📧 {nb} alerte(s) email envoyée(s).")
    else: flash("ℹ️ Aucune alerte.")
    return redirect(url_for("index"))

@app.route("/api/offres")
def api_offres():
    conn = get_db()
    offres = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    conn.close()
    return jsonify([dict(o) for o in offres])

@app.route("/api/stats")
def api_stats():
    return jsonify(get_stats())

@app.route("/devis/pdf/<int:id>")
@login_required
def generer_devis_pdf(id):
    conn = get_db()
    s = conn.execute("SELECT * FROM soumissions WHERE id=?", (id,)).fetchone()
    conn.close()
    
    if not s:
        flash("Soumission introuvable.")
        return redirect(url_for("soumissions"))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("<b>BUREAU D'ÉTUDES - LOGO</b>", styles["Title"]))
    elements.append(Paragraph("12 Rue des Ingénieurs, Alger, Algérie", styles["Normal"]))
    elements.append(Paragraph("Tél : +213 XX XX XX XX | Email : contact@be.dz", styles["Normal"]))
    elements.append(Spacer(1, 30))

    ref = s['reference_ao'] if s['reference_ao'] else f"REF-{s['id']}"
    elements.append(Paragraph(f"<b>DEVIS N° : {ref}</b>", styles["Heading2"]))
    elements.append(Paragraph(f"Date : {date.today().strftime('%d/%m/%Y')}", styles["Normal"]))
    elements.append(Spacer(1, 20))

    elements.append(Paragraph(f"<b>Client :</b> {s['client']}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Objet :</b> {s['objet']}", styles["Normal"]))
    elements.append(Spacer(1, 30))

    montant = s['montant_estime']
    tva = montant * 0.19
    ttc = montant + tva
    data = [
        ["Désignation", "Montant HT", "TVA (19%)", "Montant TTC"],
        [s['objet'], f"{montant:,.2f} {s['devise']}", f"{tva:,.2f} {s['devise']}", f"{ttc:,.2f} {s['devise']}"]
    ]
    table = Table(data, colWidths=[200, 100, 100, 100])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 50))

    elements.append(Paragraph("Le Directeur du Bureau d'Études", styles["Normal"]))
    elements.append(Spacer(1, 40))
    elements.append(Paragraph("_________________________", styles["Normal"]))

    doc.build(elements)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"Devis_{ref}.pdf", mimetype="application/pdf")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)