import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

class MonEmail:
    def __init__(self, serveur_smtp, port_smtp, utilisateur, mot_de_passe):
        self.serveur_smtp = serveur_smtp
        self.port_smtp = port_smtp
        self.utilisateur = utilisateur
        self.mot_de_passe = mot_de_passe

    def envoyer_email(self, destinataire, sujet, message, pieces_jointes=None):
        email = MIMEMultipart()
        email['From'] = self.utilisateur
        email['To'] = destinataire
        email['Subject'] = sujet

        email.attach(MIMEText(message, 'plain'))

        if pieces_jointes:
            for piece_jointe in pieces_jointes:
                with open(piece_jointe, 'rb') as pj:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(pj.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename= {piece_jointe}')
                email.attach(part)

        with smtplib.SMTP(self.serveur_smtp, self.port_smtp) as serveur_smtp:
            status_code, response = serveur_smtp.ehlo()
            print(f"[*] Echoing the server: {status_code} {response}")
            serveur_smtp.starttls()
            print(f"[*] Starting TLS connection: {status_code} {response}")
            serveur_smtp.login(self.utilisateur, self.mot_de_passe)
            print(f"[*] Logging in: {status_code} {response}")
            serveur_smtp.send_message(email)
            serveur_smtp.quit()
