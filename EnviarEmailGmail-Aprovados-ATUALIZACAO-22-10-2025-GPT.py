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

# === Configurações ===
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
PLANILHA = "arquivos\\WORKSHOP RUIVO EM 2H.xlsx"
ABA = "APROVADAS"
PASTA_CERTIFICADOS = pathlib.Path("alunos")  # Ajustado para seu diretório real
PAUSA_MIN, PAUSA_MAX = 5, 15  # segundos (reduzido para agilidade)

# === Utilidades ===
def get_timestamp():
    return datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

def log_print(msg):
    print(f"[{get_timestamp()}] {msg}")

# === Autenticação Gmail ===
def get_gmail_service():
    from google.oauth2.credentials import Credentials
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            "client_secret_909845293687-85b52o5kjiaefi1bv9m8kf4u6nohhn4a.apps.googleusercontent.com.json",
            SCOPES
        )
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)

# === Função de envio ===
def send_email(service, nome, email, attachment_path):
    msg = MIMEMultipart()
    msg["to"] = email
    msg["subject"] = f"Resultado da Prova - Imersão Correção de Ruivos - {nome}"

    html_content = f"""
        <p>Olá, {nome},</p>
        <p>É com muito orgulho que comunico que você foi aprovada(o)! 🥳</p>
        <p><b>Segue o gabarito:</b></p>
        <p>1-C | 2-B | 3-A | 4-C | 5-B | 6-C | 7-B | 8-C | 9-A | 10-B</p>
        <p>Agora eu quero te ajudar a entrar na formação completa <b>Eu, Especialista em Ruivos</b>.</p>
        <p>💬 <b>WhatsApp:</b> <a href='https://wa.me/556195859112'>Clique aqui para falar comigo</a></p>
        <p>Com carinho,<br><b>Gleici</b></p>
    """
    msg.attach(MIMEText(html_content, "html"))

    with open(attachment_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment_path)}"'
        msg.attach(part)

    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    message = {"raw": raw_message}
    service.users().messages().send(userId="me", body=message).execute()

# === Execução principal ===
service = get_gmail_service()
df = pd.read_excel(PLANILHA, sheet_name=ABA)

enviados, erros = 0, 0
log_path = "log_envio_emails.txt"

with open(log_path, "a", encoding="utf-8") as log:
    for _, row in df.iterrows():
        nome = str(row.get("NOME", "")).strip()
        email = str(row.get("E-MAIL", "")).strip()

        if not nome or email.lower() in ["nan", "none", ""]:
            log_print(f"[IGNORADO] Linha sem nome/email válido → {nome} / {email}")
            continue

        attachment = PASTA_CERTIFICADOS / f"{nome}.png"

        try:
            if attachment.exists():
                send_email(service, nome, email, str(attachment))
                log_print(f"[OK] E-mail enviado → {nome} ({email})")
                log.write(f"[{get_timestamp()}] OK - {nome} ({email})\n")
                enviados += 1
            else:
                log_print(f"[ERRO] Certificado não encontrado → {nome}")
                log.write(f"[{get_timestamp()}] ERRO - Certificado ausente: {nome}\n")
                erros += 1

        except Exception as e:
            log_print(f"[FALHA] {nome} ({email}) → {e}")
            log.write(f"[{get_timestamp()}] FALHA - {nome} ({email}) → {e}\n")
            erros += 1

        time.sleep(random.randint(PAUSA_MIN, PAUSA_MAX))

# === Resumo final ===
print("\n" + "="*60)
print(f"✅ Enviados com sucesso: {enviados}")
print(f"⚠️ Falhas/ausentes: {erros}")
print(f"📄 Log salvo em: {log_path}")
print("="*60)
