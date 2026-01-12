import os
import base64
import time
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pathlib

# Caminho da planilha e dos anexos
caminho_certificados = pathlib.Path(r"aprovados")
df = pd.read_excel("ruivos.xlsx", sheet_name="APROVADOS")

# Escopo necessário para enviar e-mails
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def get_gmail_service():
    """Autentica e retorna um serviço da API do Gmail"""
    creds = None
    if os.path.exists("token.json"):
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file("client_secret_909845293687-85b52o5kjiaefi1bv9m8kf4u6nohhn4a.apps.googleusercontent.com.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)

def send_email(service, nome, email, attachment_path):
    """Envia um e-mail com anexo pelo Gmail API"""
    msg = MIMEMultipart()
    msg["to"] = email
    msg["subject"] = f"Resultado da Prova - Semana da Conquista com Ruivos - {nome}"
    
    # Corpo do e-mail em HTML
    html_content = f"""
    <p>Olá, {nome}</p>
    <br>
    <p>Você foi aprovado(a)! 🥳</p>
    <p>Segue o gabarito e seu desconto especial.</p>
    <br>
    <p><a href='https://wa.me/5561995859112'>Fale com nosso suporte</a></p>
    """
    msg.attach(MIMEText(html_content, "html"))

    # Anexar certificado
    with open(attachment_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment_path)}"'
        msg.attach(part)

    # Enviar e-mail via API
    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    message = {"raw": raw_message}
    service.users().messages().send(userId="me", body=message).execute()

# Criar serviço do Gmail
service = get_gmail_service()

# Loop para envio de e-mails
# for index, row in df.iterrows():
nome = "PEDRO LEON"
email = "pedro.leon23@gmail.com"
attachment = caminho_certificados / f"{nome}.png"
# nome = row["NOME"]
    # email = row["E-MAIL"]
    # attachment = caminho_certificados / f"{nome}.png"

if attachment.exists():
    send_email(service, nome, email, str(attachment))
    print(f"E-mail enviado para {nome} ({email})")
else:
    print(f"Anexo não encontrado para {nome}")

time.sleep(2)  # Pequeno delay para evitar limite do Google
