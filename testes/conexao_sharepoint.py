from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import os
from dotenv import load_dotenv

#CAMINHO UNICO DO ARQUIVO .env
from pathlib import Path
env_path = Path(__file__).resolve().parent.parent / 'config' / '.env'
load_dotenv(dotenv_path=env_path)

# Carrega variáveis do .env
load_dotenv()
username = os.getenv("LDPA")
password = os.getenv("PWD")

#teste = f'user {username}; senha {password}'
#print(teste)

# URL do site SharePoint
sharepoint_url = 'C:\Users\joao_loliveira\OneDrive - Sicredi\Power BI 4501 - Geral\Relatorio Diario'


# Autenticação
ctx_auth = AuthenticationContext(sharepoint_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(sharepoint_url, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"Conectado com sucesso ao site: {web.properties['Title']}")
else:
    print("Erro na autenticação com o SharePoint.")