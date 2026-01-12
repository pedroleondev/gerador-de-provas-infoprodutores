import os
import base64
import time
import pandas as pd
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pathlib
import datetime

# Variável com data e hora atual no formato desejado
def get_timestamp():
    return datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

# Função para exibir mensagens com timestamp
def log_print(message):
    print(f"[{get_timestamp()}] {message}")



# Caminho da planilha e dos anexos
caminho_certificados = pathlib.Path(r"aprovados")
df = pd.read_excel("WORKSHOP RUIVO EM 2H.xlsx", sheet_name="APROVADAS")

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
    msg["subject"] = f"Resultado da Prova - Imersão Correção de Ruivos - {nome}"
    
    # Corpo do e-mail em HTML
    html_content = f"""
        <p>Olá, {nome},</p>
        <p>É com muito orgulho que comunico que você foi aprovada(o) ! 🥳</p>
        
        <p><b>Segue o gabarito:</b></p>
        <p>1-C | 2-B | 3-A | 4-C | 5-B | 6-C | 7-B | 8-C | 9-A | 10-B</p>

        <p>Agora eu quero te ajudar a entrar na formação completa <b>Eu, Especialista em Ruivos</b>.</p>

        <p>Me chame no chat, pois sua matrícula pode ser feita por:</p>
        <ul>
            <li>💵 <b>Cartão de Crédito ou Pix</b></li>
            <li>🧾 <b>Boleto parcelado</b></li>
            <li>💳 <b>Crédito recorrente (que não ocupa limite)</b></li>
        </ul>

        <p><b>⚠ Mas as vagas são limitadas!</b></p>

        <p>💬 <b>WhatsApp:</b> <a href='https://wa.me/556195859112'>Clique aqui para falar comigo</a></p>

        <p>Com carinho,</p>
        <p><b>Gleici</b></p>
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
for index, row in df.iterrows():
    nome = row["NOME"]
    email = row["E-MAIL"]
    attachment = caminho_certificados / f"{nome}.png"

    if attachment.exists():
        send_email(service, nome, email, str(attachment))
        log_print(f"E-mail enviado para {nome} ({email})")
    else:
        log_print(f"Anexo não encontrado para {nome}")

    # Pausa aleatória entre 5 e 30 segundos
    time.sleep(random.randint(5, 30))
