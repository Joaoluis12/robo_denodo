from dotenv import load_dotenv
import os
import pyodbc

# Caminho absoluto ou relativo para o .env
load_dotenv(dotenv_path='config/.env')

uid = os.getenv("LDAP_USER")
pwd = os.getenv("PWD")


conn_str = (
    "DRIVER={DenodoODBC Unicode(x64)};"
    "DATABASE=ldw;"
    "SERVER=virtualizador.sicredi.net;"
    "PORT=9996;"
    f"UID={uid};"
    f"PWD={pwd}"
)


try:
    conn = pyodbc.connect(conn_str)
    print("Conexão bem-sucedida com autenticação integrada!")
except Exception as e:
    print("Erro na conexão:", e)


