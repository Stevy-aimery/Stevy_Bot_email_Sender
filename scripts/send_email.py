import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import time

# Chemins des fichiers
excel_path = "\\Kaba_Bot_email-sender\\data\\email.xlsx"
body_template_path = "\\Kaba_Bot_email-sender\\templates\\mail_CandidaturePFE.html"
resume_path = "\\Kaba_Bot_email-sender\\docs\\Mon_CV.pdf"
log_path = "\\Kaba_Bot_email-sender\\logs\\email_log.txt"

# Configuration de l'email :
mail_from = "stevy <test@gmail.com>"
message_subject = "Candidature Stage PFE - Développeur d'applications Web et Mobiles"
smtp_server = "smtp.gmail.com"
smtp_port = 587

# Saisir les informations d'identification
username = input("Entrez votre adresse e-mail Gmail : ")
password = input("Entrez votre mot de passe Gmail : ")

# Valider les chemins
if not os.path.exists(excel_path):
    print(f"Fichier Excel introuvable : {excel_path}")
    exit()
if not os.path.exists(body_template_path):
    print(f"Fichier de modèle HTML introuvable : {body_template_path}")
    exit()

# Charger les données Excel
def log_message(message):
    with open(log_path, "a") as log_file:
        log_file.write(f"{message}\n")

log_message(f"Email Log - {pd.Timestamp.now()}")
df = pd.read_excel(excel_path)

# Configurer le client SMTP
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
try:
    server.login(username, password)
except smtplib.SMTPAuthenticationError:
    print("Erreur d'authentification SMTP. Vérifiez vos identifiants.")
    exit()

# Envoyer les emails
for index, row in df.iterrows():
    mail_to = row.get("EMAIL")
    if not pd.notna(mail_to):
        print("Adresse email absente, passage à la ligne suivante...")
        log_message("Adresse email absente pour une ligne, sautée.")
        continue

    print(f"Envoi de l'email à {mail_to}")

    # Construire le message
    msg = MIMEMultipart()
    msg["From"] = mail_from
    msg["To"] = mail_to
    msg["Subject"] = message_subject

    # Lire le contenu HTML
    with open(body_template_path, "r", encoding="utf-8") as file:
        html_body = file.read()
    msg.attach(MIMEText(html_body, "html"))

    # Ajouter une pièce jointe si disponible
    if os.path.exists(resume_path):
        with open(resume_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(resume_path)}"
        )
        msg.attach(part)
    else:
        print(f"Pièce jointe introuvable : {resume_path}")

    # Envoyer l'email
    try:
        server.send_message(msg)
        print(f"Email envoyé avec succès à {mail_to}")
        log_message(f"Succès : Email envoyé à {mail_to} à {pd.Timestamp.now()}")
    except Exception as e:
        print(f"Erreur lors de l'envoi de l'email à {mail_to} : {e}")
        log_message(f"Erreur : Échec de l'envoi de l'email à {mail_to} : {e}")
    finally:
        del msg

    # Pause entre les envois
    time.sleep(5)

# Fermeture de la connexion SMTP
server.quit()
print("Processus d'envoi des emails terminé.")
log_message(f"Processus terminé à {pd.Timestamp.now()}")
