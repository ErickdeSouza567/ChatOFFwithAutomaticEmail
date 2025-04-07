from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import base64
import os
import time
import pandas as pd
from email.mime.text import MIMEText

# CONFIGS
URL_BASE = "https://example.com/chat/list/(chat_status_ids)/5/(sortby)/id_desc/(page)/"
LOGIN_URL = "https://example.com/user/login"
USUARIO = "SEU_USUARIO_AQUI"
SENHA = "SUA_SENHA_AQUI"

LIMITE_DIAS_ANALISE = 1
LIMITE_DATA = (datetime.now() - timedelta(days=LIMITE_DIAS_ANALISE)).replace(hour=17, minute=0, second=0, microsecond=0)

# Gmail API
CLIENT_ID = 'SEU_CLIENT_ID_AQUI'
CLIENT_SECRET = 'SEU_CLIENT_SECRET_AQUI'
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
TOKEN_FILE = 'token.json'

def autenticar_gmail():
    creds = None
    original_window = driver.current_window_handle
    
    if os.path.exists(TOKEN_FILE):
        print(f"↳ Carregando credenciais existentes de {TOKEN_FILE}")
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    if not creds or not creds.valid:
        print("↳ Credenciais inválidas ou não encontradas. Iniciando fluxo de autenticação...")
        flow = InstalledAppFlow.from_client_config(
            {
                "installed": {
                    "client_id": CLIENT_ID,
                    "client_secret": CLIENT_SECRET,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token"
                }
            },
            SCOPES
        )
        creds = flow.run_local_server(port=0)

        for window_handle in driver.window_handles:
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                driver.close()
        driver.switch_to.window(original_window)

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        print(f"↳ Novas credenciais salvas em {TOKEN_FILE}")
    
    return build('gmail', 'v1', credentials=creds)

def criar_email(destinatario, assunto, corpo):
    print(f"↳ Criando e-mail para: {destinatario}")
    mensagem = MIMEText(corpo)
    mensagem['to'] = destinatario
    mensagem['subject'] = assunto
    raw = base64.urlsafe_b64encode(mensagem.as_bytes()).decode()
    return {'raw': raw}

def enviar_email(service, destinatario, assunto, corpo):
    print(f"↳ Tentando enviar e-mail para: {destinatario}")
    mensagem = criar_email(destinatario, assunto, corpo)
    try:
        response = service.users().messages().send(userId="me", body=mensagem).execute()
        print(f"📤 E-mail enviado com sucesso para: {destinatario}, ID da mensagem: {response.get('id')}")
    except Exception as e:
        print(f"❌ Falha ao enviar e-mail para {destinatario}: {e}")

# PASTA DE DESTINO
PASTA_DESTINO = r"C:\CAMINHO\PARA\PASTA\DESTINO"
DATA_HOJE = datetime.now().strftime("%d_%m_%Y")
NOME_ARQUIVO = f"CHAT_{DATA_HOJE}.xlsx"
CAMINHO_COMPLETO = os.path.join(PASTA_DESTINO, NOME_ARQUIVO)

# SETUP DO NAVEGADOR
options = Options()
options.add_argument('--start-maximized')
options.add_experimental_option("detach", True)
driver = webdriver.Edge(options=options)
wait = WebDriverWait(driver, 10)

# LOGIN
print("↳ Iniciando login no sistema...")
driver.get(LOGIN_URL)
time.sleep(2)
driver.find_element(By.NAME, "Username").send_keys(USUARIO)
driver.find_element(By.NAME, "Password").send_keys(SENHA)
driver.find_element(By.NAME, "Login").click()
time.sleep(3)

if "login" in driver.current_url.lower():
    print("⚠️ Falha no login.")
    driver.quit()
    exit()

print("✅ Login realizado com sucesso!")

resultados = []
pagina = 1
navegador_ativo = True
gmail_service = autenticar_gmail()

while navegador_ativo:
    try:
        print(f"↳ Carregando página {pagina}...")
        driver.get(URL_BASE + str(pagina))
    except Exception as e:
        print(f"🚫 Erro ao carregar página {pagina}: {e}")
        break

    time.sleep(2)
    linhas = driver.find_elements(By.CSS_SELECTOR, "tr.chat-row-tr")

    if not linhas:
        print("⏹️ Nenhum chat encontrado na página atual. Encerrando.")
        break

    for linha in linhas:
        try:
            tds = linha.find_elements(By.TAG_NAME, "td")
            if len(tds) < 2:
                print(f"↳ Linha sem dados suficientes: {linha.text}")
                continue

            texto = tds[1].text.strip()
            if 'visitor' in texto.lower():
                continue

            partes = texto.split(", ")
            if len(partes) < 3:
                print(f"↳ Formato inválido na linha: {texto}")
                continue

            nome = partes[0]
            data_msg = datetime.strptime(partes[1], "%Y-%m-%d %H:%M:%S")
            origem = partes[2]

            print(f"↳ Processando chat: Nome={nome}, Data={data_msg}, Origem={origem}")

            if data_msg < LIMITE_DATA:
                print("⏹️ Limite de data atingido. Encerrando.")
                navegador_ativo = False
                break

            chat_id = linha.get_attribute("data-chat-id")
            chat_url = f"{URL_BASE}{pagina}#/" + f"chat-id-{chat_id}"
            print(f"↳ Abrindo chat ID {chat_id} em nova aba: {chat_url}")
            driver.execute_script(f"window.open('{chat_url}', '_blank');")
            driver.switch_to.window(driver.window_handles[1])

            email = None
            try:
                wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.col-6.pb-1")))
                time.sleep(1)
                divs_email = driver.find_elements(By.CSS_SELECTOR, 'div.col-6.pb-1')
                print(f"↳ Procurando e-mails no chat {chat_id}...")
                for div in divs_email:
                    if "email" in div.text.lower() and "@" in div.text:
                        link = div.find_element(By.TAG_NAME, "a")
                        email = link.text.strip()
                        break

                if email and "@" in email and "." in email.split("@")[1]:
                    print(f"↳ ✅ {nome} | {email}")
                    resultados.append({
                        "ID": chat_id,
                        "Data": data_msg.strftime("%Y-%m-%d %H:%M:%S"),
                        "Origem": origem,
                        "Email": email
                    })

                    if "EJA SEED" in origem.upper() and "SESI" not in origem.upper():
                        assunto = "Retorno sobre a solicitação no CHAT EJA - Suporte"
                        corpo = (
                            f"Olá!\n\n"
                            f"Verificamos que você acessou o CHAT EJA na data e horário de {data_msg.strftime('%d/%m/%Y %H:%M')}.\n"
                            f"Como estávamos sem atendentes no momento, estamos entrando em contato para saber se ainda precisa de ajuda.\n\n"
                            f"Você pode simplesmente responder este e-mail com sua dúvida, ou acessar novamente o chat pelo site.\n\n"
                            f"Atenciosamente,\nEquipe de Suporte EJA"
                        )
                        enviar_email(gmail_service, email, assunto, corpo)

                else:
                    print(f"↳ ❌ E-mail inválido: {email}")

            except Exception as e:
                print(f"Erro ao buscar e-mail no chat {chat_id}: {e}")

            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(1)

        except Exception as e:
            print(f"Erro ao processar um chat: {e}")
            continue

    pagina += 1

# SALVAR RESULTADOS
if resultados:
    df = pd.DataFrame(resultados)
    os.makedirs(PASTA_DESTINO, exist_ok=True)

    with pd.ExcelWriter(CAMINHO_COMPLETO := os.path.join(PASTA_DESTINO, "CHAT.xlsx"), engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        (max_row, max_col) = df.shape

        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'name': 'TabelaChats',
            'columns': [{'header': col} for col in df.columns]
        })

    print(f"\n✅ Dados salvos em: {CAMINHO_COMPLETO}")
else:
    print("\n⚠️ Nenhum dado com e-mail foi encontrado.")

# DELETAR TOKEN.JSON
try:
    if os.path.exists(TOKEN_FILE):
        os.remove(TOKEN_FILE)
        print(f"↳ Token {TOKEN_FILE} removido com sucesso.")
except Exception as e:
    print(f"⚠️ Erro ao remover token: {e}")

# Fechar navegador
try:
    print("↳ Fechando navegador...")
    driver.quit()
except Exception as e:
    print(f"⚠️ Erro ao fechar navegador: {e}")
