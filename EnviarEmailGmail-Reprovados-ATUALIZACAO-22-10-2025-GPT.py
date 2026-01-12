import os
import base64
import time
import pandas as pd
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import datetime

# === CONFIGURAÇÕES ===
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
PLANILHA = "arquivos\\WORKSHOP RUIVO EM 2H.xlsx"
ABA = "REPROVADAS"
PAUSA_MIN, PAUSA_MAX = 5, 15  # segundos
LOG_PATH = "log_envio_reprovadas.txt"

# === UTILIDADES ===
def get_timestamp():
    return datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

def log_print(message):
    print(f"[{get_timestamp()}] {message}")

# === AUTENTICAÇÃO GMAIL ===
def get_gmail_service():
    from google.oauth2.credentials import Credentials
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)

# === FUNÇÃO DE ENVIO ===
def send_email(service, nome, email):
    msg = MIMEMultipart()
    msg["to"] = email
    msg["subject"] = f"Resultado da Prova - Semana da Conquista com Ruivos - {nome}"

    html_content = f"""
        <p>Olá, {nome},</p>
        <p>Infelizmente, você não atingiu o número de acertos necessários… 😢</p>

        <p><b>Segue o gabarito:</b></p>
        <p>1-C | 2-B | 3-A | 4-C | 5-B | 6-C | 7-B | 8-C | 9-A | 10-B</p>

        <p>Mas eu quero te ajudar a <b>MELHORAR</b> e conquistar sua certificação com a formação completa <b>Eu, Especialista em Ruivos</b>.</p>

        <p>Me chame no chat, pois sua matrícula pode ser feita por:</p>
        <ul>
            <li>💵 <b>Cartão de Crédito ou Pix</b></li>
            <li>🧾 <b>Boleto parcelado</b></li>
            <li>💳 <b>Crédito recorrente (que não ocupa limite)</b></li>
        </ul>

        <p><b>⚠ Mas as vagas são limitadas!</b></p>
        <p>💬 <b>WhatsApp:</b> <a href='https://wa.me/556195859112'>Clique aqui para falar comigo</a></p>

        <p>Aguardo você na próxima turma!</p>
        <p>Com carinho,<br><b>Gleici</b></p>
    """

    msg.attach(MIMEText(html_content, "html"))

    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    message = {"raw": raw_message}
    service.users().messages().send(userId="me", body=message).execute()

# === EXECUÇÃO PRINCIPAL ===
service = get_gmail_service()
df = pd.read_excel(PLANILHA, sheet_name=ABA)

enviados, erros = 0, 0

with open(LOG_PATH, "a", encoding="utf-8") as log:
    for _, row in df.iterrows():
        nome = str(row.get("NOME", "")).strip()
        email = str(row.get("E-MAIL", "")).strip()

        # Ignora linhas vazias
        if not nome or email.lower() in ["nan", "none", ""]:
            log_print(f"[IGNORADO] Linha sem nome/email válido → {nome} / {email}")
            continue

        try:
            send_email(service, nome, email)
            log_print(f"[OK] E-mail enviado → {nome} ({email})")
            log.write(f"[{get_timestamp()}] OK - {nome} ({email})\n")
            enviados += 1
        except Exception as e:
            log_print(f"[FALHA] {nome} ({email}) → {e}")
            log.write(f"[{get_timestamp()}] FALHA - {nome} ({email}) → {e}\n")
            erros += 1

        # Pausa aleatória entre envios
        time.sleep(random.randint(PAUSA_MIN, PAUSA_MAX))

# === RESUMO FINAL ===
print("\n" + "="*60)
print(f"✅ Enviados com sucesso: {enviados}")
print(f"⚠️ Falhas: {erros}")
print(f"📄 Log salvo em: {LOG_PATH}")
print("="*60)
