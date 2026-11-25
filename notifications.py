import smtplib
import sqlite3
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import date, timedelta
from dotenv import load_dotenv

load_dotenv()


def offres_a_notifier(db_path, jours_avant=7):
    """Retourne les offres qui expirent dans les N prochains jours."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    aujourd_hui = date.today()
    limite = aujourd_hui + timedelta(days=jours_avant)
    rows = conn.execute("""
        SELECT * FROM offres
        WHERE statut='Active'
          AND date_fin != '' AND date_fin >= ? AND date_fin <= ?
        ORDER BY date_fin
    """, (aujourd_hui.isoformat(), limite.isoformat())).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def construire_email(offres):
    html = """
    <html><body style="font-family:Arial,sans-serif;color:#2c3e50;">
      <h2 style="color:#e74c3c;">⚠️ Offres proches de l'expiration</h2>
      <p>Les offres suivantes expirent dans les prochains jours :</p>
      <table style="border-collapse:collapse;width:100%;max-width:700px;">
        <tr style="background:#2c3e50;color:white;">
          <th style="padding:8px;text-align:left;">Nom</th>
          <th style="padding:8px;">Prix</th>
          <th style="padding:8px;">Date fin</th>
        </tr>
    """
    for o in offres:
        html += f"""
        <tr style="border-bottom:1px solid #ecf0f1;">
          <td style="padding:8px;"><b>{o['nom']}</b></td>
          <td style="padding:8px;text-align:center;">{o['prix']:.2f} €</td>
          <td style="padding:8px;text-align:center;color:#e74c3c;">
            {o['date_fin']}</td>
        </tr>"""
    html += """
      </table>
      <p style="font-size:12px;color:#7f8c8d;margin-top:20px;">
        Message automatique — Système de gestion des offres.
      </p>
    </body></html>"""
    return html


def envoyer_alerte_expiration(db_path, jours_avant=None):
    """Envoie un email avec les offres expirant bientôt. Retourne le nombre d'offres."""
    jours_avant = int(jours_avant or os.getenv("DAYS_BEFORE_EXPIRY", 7))
    offres = offres_a_notifier(db_path, jours_avant)
    if not offres:
        return 0

    # Mode simulation si SMTP non configuré
    smtp_host = os.getenv("SMTP_HOST")
    smtp_user = os.getenv("SMTP_USER")
    smtp_pass = os.getenv("SMTP_PASSWORD")
    destinataire = os.getenv("EMAIL_DESTINATAIRE", smtp_user)

    if not all([smtp_host, smtp_user, smtp_pass, destinataire]):
        print("⚠️ SMTP non configuré — simulation :")
        for o in offres:
            print(f"  - {o['nom']} expire le {o['date_fin']}")
        return len(offres)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"⚠️ {len(offres)} offre(s) expire(nt) bientôt"
    msg["From"] = smtp_user
    msg["To"] = destinataire

    msg.attach(MIMEText(f"{len(offres)} offre(s) proche(s) de l'expiration.", "plain"))
    msg.attach(MIMEText(construire_email(offres), "html"))

    try:
        with smtplib.SMTP(smtp_host, int(os.getenv("SMTP_PORT", 587))) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, [destinataire], msg.as_string())
        print(f"✅ Email envoyé à {destinataire} ({len(offres)} offres).")
    except Exception as e:
        print(f"❌ Erreur envoi email : {e}")
        return 0

    return len(offres)


if __name__ == "__main__":
    nb = envoyer_alerte_expiration(os.getenv("DB_PATH", "offres.db"))
    print(f"Total offres signalées : {nb}")