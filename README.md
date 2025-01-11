# author : Stevy-aimery

# Email Sender Automation

Ce projet coder en PowerShell permet d'envoyer des e-mails personnalisés en masse à partir d'une liste d'e-mails dans un fichier Excel ou csv.

# Mot de pass
Mot de passe d’application à 16 Chars (necessite une activaion de l'authentification en deux étapes sur le count Google, Puis choisir "Mail" comme App cible) 
 ou 
Mot de pass count GMail


## Structure du projet

- **scripts/** : Contient le script PowerShell principal (`send_email.ps1`).
- **templates/** : Contient les modèles HTML pour les e-mails.
- **data/** : Contient les fichiers Excel ou csv avec les adresses e-mail des destinataires.
- **attachments/** : Contient les fichiers à joindre (par exemple: CV, Lettre de motivation, certifications, etc).
- **logs/** : Contient les fichiers de log pour suivre les envois.

## Configuration

1. Installez le module PowerShell `ImportExcel` :
   ```powershell
   Install-Module -Name ImportExcel -Scope CurrentUser
