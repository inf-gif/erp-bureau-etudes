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
    rows = conn.execute("SELECT prix, devise FROM offres WHERE statut='Active'").fetchall()
    total_row = conn.execute("SELECT COUNT(*) FROM offres").fetchone()
    actives_row = conn.execute("SELECT COUNT(*) FROM offres WHERE statut='Active'").fetchone()
    expirees_row = conn.execute("""
        SELECT COUNT(*) FROM offres WHERE date_fin != '' AND date_fin < ? AND statut='Active'
    """, (date.today().isoformat(),)).fetchone()
    
    taux_map = {d["code"]: d["taux"] for d in get_devises()}
    ca_dzd = sum(r["prix"] * taux_map.get(r["devise"], 1.0) for r in rows)
    
    conn.close()
    return {"total": total_row[0] or 0, "actives": actives_row[0] or 0,
            "ca": ca_dzd, "expirees": expirees_row[0] or 0}

BASE_CSS = """
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
  body { margin: 0; background: #ecf0f1; color: #2c3e50; }
  header { background: #2c3e50; color: white; padding: 15px 30px; display: flex; justify-content: space-between; align-items: center; }
  header h1 { margin: 0; font-size: 20px; }
  header nav a { color: #f39c12; text-decoration: none; margin-left: 15px; font-weight: bold; }
  .container { max-width: 1200px; margin: 20px auto; padding: 0 20px; }
  .stats { display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap; }
  .card { flex: 1; min-width: 180px; background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  .card h3 { margin: 0 0 5px; color: #7f8c8d; font-size: 13px; }
  .card .value { font-size: 22px; font-weight: bold; }
  form.box { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
  form.box .grid { display: grid; grid-template-columns: repeat(auto-fit,minmax(180px,1fr)); gap: 12px; }
  input, select { padding: 8px; border: 1px solid #bdc3c7; border-radius: 5px; width: 100%; }
  button, .btn { padding: 9px 16px; border: none; border-radius: 5px; cursor: pointer; color: white; font-weight: bold; text-decoration: none; display: inline-block; font-size: 14px; }
  .btn-green { background: #27ae60; } .btn-orange { background: #f39c12; }
  .btn-red { background: #e74c3c; } .btn-blue { background: #2980b9; }
  .btn-gray { background: #7f8c8d; } .btn-purple { background: #8e44ad; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  th { background: #2c3e50; color: white; padding: 10px; text-align: left; }
  td { padding: 10px; border-bottom: 1px solid #ecf0f1; }
  tr.active { background: #eafaf1; } tr.expiree { background: #fdecea; }
  .actions a { margin-right: 5px; font-size: 13px; padding: 5px 10px; }
  .flash { background: #2ecc71; color: white; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
  .flash-error { background: #e74c3c; }
  .search { display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap; }
  .pagination { display: flex; justify-content: center; gap: 5px; margin-top: 15px; }
  .pagination a, .pagination span { padding: 6px 12px; background: white; border: 1px solid #bdc3c7; border-radius: 5px; text-decoration: none; color: #2c3e50; }
  .pagination .active { background: #2c3e50; color: white; font-weight: bold; }
  .login-box { max-width: 400px; margin: 80px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); }
  .login-box h2 { text-align: center; margin-top: 0; }
  .login-box label { display: block; margin: 12px 0 5px; font-weight: bold; }
  .chart-container { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  .badge { padding: 3px 8px; border-radius: 12px; font-size: 12px; background: #ecf0f1; }
  .badge-devise { background: #3498db; color: white; }
  .info { background: #d1ecf1; color: #0c5460; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
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
<title>Gestion des Offres Multi-Devises</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
""" + BASE_CSS + """</head><body>
<header>
  <h1>📦 Gestion des Offres</h1>
  <nav>
    <a href="{{ url_for('index') }}">🏠 Offres</a>
    <a href="{{ url_for('devises') }}">💱 Devises</a>
    <a href="{{ url_for('logout') }}">🚪 Déconnexion ({{ current_user.id }})</a>
  </nav>
</header>
<div class="container">
  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}
  <div class="stats">
    <div class="card"><h3>Total offres</h3><div class="value">{{ stats.total }}</div></div>
    <div class="card"><h3>Actives</h3><div class="value" style="color:#27ae60">{{ stats.actives }}</div></div>
    <div class="card"><h3>CA potentiel (converti)</h3><div class="value">{{ "{:,.2f}".format(stats.ca).replace(',', ' ') }} DZD</div></div>
    <div class="card"><h3>Expirées</h3><div class="value" style="color:#e74c3c">{{ stats.expirees }}</div></div>
  </div>
  <div class="info">
    💡 <b>Info :</b> Le CA total est converti automatiquement en <b>DZD</b> selon les taux configurés dans l'onglet <a href="{{ url_for('devises') }}">Devises</a>.
  </div>
  <div class="chart-container">
    <canvas id="caChart" height="70"></canvas>
  </div>
  <form class="box" method="POST" action="{{ url_for('ajouter') }}">
    <h3 style="margin-top:0">{{ 'Modifier' if edit else 'Ajouter' }} une offre</h3>
    <div class="grid">
      <input name="nom" placeholder="Nom *" required value="{{ edit.nom if edit else '' }}">
      <input name="description" placeholder="Description" value="{{ edit.description if edit else '' }}">
      <input name="prix" type="number" step="0.01" placeholder="Prix *" required value="{{ edit.prix if edit else '' }}">
      <select name="devise">
        {% for d in devises_list %}
        <option value="{{ d.code }}" {% if (edit and edit.devise==d.code) or (not edit and d.code=='DZD') %}selected{% endif %}>{{ d.code }} - {{ d.nom }}</option>
        {% endfor %}
      </select>
      <input name="categorie" placeholder="Catégorie" value="{{ edit.categorie if edit else 'Général' }}">
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
    <a href="{{ url_for('notifier_expirations') }}" class="btn btn-purple" onclick="return confirm('Envoyer les alertes par email ?')">📧 Notifier</a>
  </div>
  <table>
    <tr>
      <th>ID</th>
      <th><a href="?sort=nom&order={{ 'desc' if sort=='nom' and order=='asc' else 'asc' }}">Nom ↕</a></th>
      <th>Catégorie</th>
      <th><a href="?sort=prix&order={{ 'desc' if sort=='prix' and order=='asc' else 'asc' }}">Prix ↕</a></th>
      <th>Devise</th>
      <th>Équivalent DZD</th>
      <th><a href="?sort=date_fin&order={{ 'desc' if sort=='date_fin' and order=='asc' else 'asc' }}">Fin ↕</a></th>
      <th>Statut</th>
      <th>Actions</th>
    </tr>
    {% for o in offres %}
    <tr class="{{ 'expiree' if o.date_fin and o.date_fin < today and o.statut=='Active' else ('active' if o.statut=='Active' else '') }}">
      <td>{{ o.id }}</td>
      <td><b>{{ o.nom }}</b><br><small style="color:#7f8c8d">{{ o.description }}</small></td>
      <td><span class="badge">{{ o.categorie }}</span></td>
      <td><b>{{ "{:,.2f}".format(o.prix).replace(',', ' ') }}</b></td>
      <td><span class="badge badge-devise">{{ o.devise }}</span></td>
      <td>{{ "{:,.2f}".format(o.prix * taux_map[o.devise]).replace(',', ' ') }} DZD</td>
      <td>{{ o.date_fin }}</td>
      <td>{{ o.statut }}</td>
      <td class="actions">
        <a class="btn btn-orange" href="{{ url_for('modifier', id=o.id) }}">✏️</a>
        <a class="btn btn-red" href="{{ url_for('supprimer', id=o.id) }}" onclick="return confirm('Supprimer ?')">🗑️</a>
      </td>
    </tr>
    {% endfor %}
  </table>
  {% if total_pages > 1 %}
  <div class="pagination">
    {% for p in range(1, total_pages + 1) %}
      {% if p == page %}<span class="active">{{ p }}</span>
      {% else %}<a href="?page={{ p }}&q={{ q }}&sort={{ sort }}&order={{ order }}">{{ p }}</a>{% endif %}
    {% endfor %}
  </div>
  {% endif %}
</div>
<script>
  new Chart(document.getElementById('caChart'), {
    type: 'bar',
    data: {
      labels: {{ chart_labels|safe }},
      datasets: [{ label: 'CA par catégorie (DZD)', data: {{ chart_data|safe }}, backgroundColor: 'rgba(52, 152, 219, 0.7)', borderColor: 'rgba(52, 152, 219, 1)', borderWidth: 1, borderRadius: 5 }]
    },
    options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, title: {display:true, text:'DZD'} } } }
  });
</script>
</body></html>"""

DEVISES_TEMPLATE = """<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Gestion des Devises</title>""" + BASE_CSS + """</head><body>
<header>
  <h1>💱 Gestion des Devises</h1>
  <nav>
    <a href="{{ url_for('index') }}">🏠 Offres</a>
    <a href="{{ url_for('devises') }}">💱 Devises</a>
    <a href="{{ url_for('logout') }}">🚪 Déconnexion ({{ current_user.id }})</a>
  </nav>
</header>
<div class="container">
  {% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}
  <div class="info">
    ⚙️ <b>Comment ça marche :</b> Le taux indique combien vaut <b>1 unité</b> de la devise en <b>DZD</b>. La devise de référence <b>DZD</b> a toujours un taux de <code>1.0</code>.
  </div>
  <form class="box" method="POST" action="{{ url_for('ajouter_devise') }}">
    <h3 style="margin-top:0">➕ Ajouter une devise</h3>
    <div class="grid">
      <input name="code" placeholder="Code (ex: CAD)" required maxlength="5" style="text-transform:uppercase">
      <input name="nom" placeholder="Nom (ex: Dollar Canadien)" required>
      <input name="taux" type="number" step="0.01" placeholder="Taux en DZD" required>
    </div>
    <br>
    <button class="btn-green">➕ Ajouter</button>
  </form>
  <table>
    <tr><th>Code</th><th>Nom</th><th>Taux (1 unité = X DZD)</th><th>Actions</th></tr>
    {% for d in devises %}
    <tr>
      <td><span class="badge badge-devise">{{ d.code }}</span></td>
      <td>{{ d.nom }}</td>
      <td>
        {% if d.code == 'DZD' %}
          <b>1.00</b> <small>(référence)</small>
        {% else %}
          <form method="POST" action="{{ url_for('modifier_devise', code=d.code) }}" style="display:flex;gap:5px;margin:0;background:none;padding:0;box-shadow:none;">
            <input name="taux" type="number" step="0.01" value="{{ d.taux }}" style="width:120px;padding:5px;">
            <button class="btn-green" style="padding:5px 10px;">💾</button>
          </form>
        {% endif %}
      </td>
      <td>
        {% if d.code != 'DZD' %}
          <a class="btn btn-red" href="{{ url_for('supprimer_devise', code=d.code) }}" onclick="return confirm('Supprimer la devise {{ d.code }} ?')">🗑️</a>
        {% else %}
          <span class="badge">🔒</span>
        {% endif %}
      </td>
    </tr>
    {% endfor %}
  </table>
</div></body></html>"""

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        p = request.form.get("password", "")
        if u == os.getenv("ADMIN_USER", "admin") and p == os.getenv("ADMIN_PASSWORD", "admin123"):
            login_user(User(u))
            flash(f"Bienvenue, {u} !")
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
    sort = request.args.get("sort", "id")
    order = request.args.get("order", "desc")
    page = int(request.args.get("page", 1))
    
    valid_cols = ["id", "nom", "prix", "date_debut", "date_fin"]
    if sort not in valid_cols: sort = "id"
    if order not in ["asc", "desc"]: order = "desc"

    conn = get_db()
    query = "SELECT * FROM offres WHERE 1=1"
    params = []
    if q:
        query += " AND (nom LIKE ? OR description LIKE ? OR statut LIKE ? OR categorie LIKE ?)"
        params.extend([f"%{q}%"] * 4)
    query += f" ORDER BY {sort} {order}"
    all_rows = conn.execute(query, params).fetchall()
    
    per_page = 5
    total_items = len(all_rows)
    total_pages = math.ceil(total_items / per_page) if total_items > 0 else 1
    start = (page - 1) * per_page
    offres = all_rows[start:start + per_page]

    taux_map = {d["code"]: d["taux"] for d in get_devises()}
    chart_data_dict = {}
    for r in conn.execute("SELECT categorie, prix, devise FROM offres WHERE statut='Active'").fetchall():
        cat = r["categorie"] or "Général"
        chart_data_dict[cat] = chart_data_dict.get(cat, 0) + (r["prix"] * taux_map.get(r["devise"], 1.0))
    conn.close()

    chart_labels = [f'"{k}"' for k in chart_data_dict.keys()]
    chart_data = list(chart_data_dict.values())

    return render_template_string(INDEX_TEMPLATE, offres=offres, stats=get_stats(),
                                   q=q, edit=None, today=date.today().isoformat(),
                                   sort=sort, order=order, page=page, total_pages=total_pages,
                                   devises_list=get_devises(), taux_map=taux_map,
                                   chart_labels=f"[{', '.join(chart_labels)}]",
                                   chart_data=f"[{', '.join(map(str, chart_data))}]")

@app.route("/ajouter", methods=["POST"])
@login_required
def ajouter():
    conn = get_db()
    conn.execute("""INSERT INTO offres (nom,description,prix,date_debut,date_fin,statut,categorie,devise)
                    VALUES (?,?,?,?,?,?,?,?)""",
                 (request.form["nom"], request.form["description"],
                  float(request.form["prix"]), request.form["date_debut"],
                  request.form["date_fin"], request.form["statut"],
                  request.form.get("categorie", "Général"), request.form.get("devise", "DZD")))
    conn.commit(); conn.close()
    flash("✅ Offre ajoutée.")
    return redirect(url_for("index"))

@app.route("/modifier/<int:id>", methods=["GET", "POST"])
@login_required
def modifier(id):
    conn = get_db()
    if request.method == "POST":
        conn.execute("""UPDATE offres SET nom=?,description=?,prix=?,date_debut=?,date_fin=?,statut=?,categorie=?,devise=?
                        WHERE id=?""",
                     (request.form["nom"], request.form["description"],
                      float(request.form["prix"]), request.form["date_debut"],
                      request.form["date_fin"], request.form["statut"],
                      request.form.get("categorie", "Général"), request.form.get("devise", "DZD"), id))
        conn.commit(); conn.close()
        flash("✏️ Offre modifiée.")
        return redirect(url_for("index"))
    offre = conn.execute("SELECT * FROM offres WHERE id=?", (id,)).fetchone()
    all_offres = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    taux_map = {d["code"]: d["taux"] for d in get_devises()}
    conn.close()
    return render_template_string(INDEX_TEMPLATE, offres=all_offres, stats=get_stats(),
                                   q="", edit=offre, today=date.today().isoformat(),
                                   sort="id", order="desc", page=1, total_pages=1,
                                   devises_list=get_devises(), taux_map=taux_map,
                                   chart_labels="[]", chart_data="[]")

@app.route("/supprimer/<int:id>")
@login_required
def supprimer(id):
    conn = get_db(); conn.execute("DELETE FROM offres WHERE id=?", (id,))
    conn.commit(); conn.close()
    flash("🗑️ Offre supprimée.")
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
        conn.commit()
        flash(f"✅ Devise {code} ajoutée.")
    except sqlite3.IntegrityError:
        flash(f"❌ La devise {code} existe déjà.")
    conn.close()
    return redirect(url_for("devises"))

@app.route("/devises/modifier/<code>", methods=["POST"])
@login_required
def modifier_devise(code):
    taux = float(request.form["taux"])
    conn = get_db()
    conn.execute("UPDATE devises SET taux=? WHERE code=?", (taux, code))
    conn.commit(); conn.close()
    flash(f"💾 Taux {code} mis à jour.")
    return redirect(url_for("devises"))

@app.route("/devises/supprimer/<code>")
@login_required
def supprimer_devise(code):
    if code == "DZD":
        flash("❌ Impossible de supprimer la devise de référence.")
        return redirect(url_for("devises"))
    conn = get_db()
    conn.execute("DELETE FROM devises WHERE code=?", (code,))
    conn.commit(); conn.close()
    flash(f"🗑️ Devise {code} supprimée.")
    return redirect(url_for("devises"))

@app.route("/export/excel")
@login_required
def export_excel():
    conn = get_db()
    offres = conn.execute("SELECT * FROM offres ORDER BY id").fetchall()
    taux_map = {d["code"]: d["taux"] for d in get_devises()}
    conn.close()
    wb = Workbook(); ws = wb.active; ws.title = "Offres"
    ws.append(["ID", "Nom", "Description", "Catégorie", "Prix", "Devise", "Équivalent DZD", "Début", "Fin", "Statut"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2C3E50")
        c.alignment = Alignment(horizontal="center")
    for o in offres:
        equiv = o["prix"] * taux_map.get(o["devise"], 1.0)
        ws.append([o["id"], o["nom"], o["description"], o["categorie"], o["prix"], o["devise"], equiv, o["date_debut"], o["date_fin"], o["statut"]])
    for i, w in enumerate([6, 25, 35, 15, 12, 8, 15, 14, 14, 12], start=1):
        ws.column_dimensions[chr(64+i)].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"offres_{date.today()}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/notifier")
@login_required
def notifier_expirations():
    nb = envoyer_alerte_expiration(DB)
    if nb > 0: flash(f"📧 {nb} alerte(s) email envoyée(s).")
    else: flash("ℹ️ Aucune offre proche de l'expiration.")
    return redirect(url_for("index"))

@app.route("/api/offres")
def api_offres():
    conn = get_db()
    offres = conn.execute("SELECT * FROM offres ORDER BY id DESC").fetchall()
    taux_map = {d["code"]: d["taux"] for d in get_devises()}
    conn.close()
    result = []
    for o in offres:
        d = dict(o)
        d["equivalent_dzd"] = round(o["prix"] * taux_map.get(o["devise"], 1.0), 2)
        result.append(d)
    return jsonify(result)

@app.route("/api/devises")
def api_devises():
    return jsonify([dict(d) for d in get_devises()])

@app.route("/api/stats")
def api_stats():
    return jsonify(get_stats())

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
