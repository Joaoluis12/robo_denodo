from dotenv import load_dotenv
import os
import pyodbc
# Caminho absoluto ou relativo para o .env

from pathlib import Path
env_path = Path(__file__).resolve().parent.parent / 'config' / '.env'
load_dotenv(dotenv_path=env_path)


load_dotenv()

uid = os.getenv("UID")
pwd = os.getenv("PWD")


conn_str = (
    "DRIVER={DenodoODBC Unicode(x64)};"
    "DATABASE=ldw;"
    "SERVER=virtualizador.sicredi.net;"
    "PORT=9996;"
    f"UID={uid};"
    f"PWD={pwd}"
)

print(pwd)

try:
    conn = pyodbc.connect(conn_str)
    print("Conexão bem-sucedida com autenticação integrada!")
except Exception as e:
    print("Erro na conexão:", e)


