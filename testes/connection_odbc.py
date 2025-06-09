import pyodbc

conn_str = (
    "DRIVER={DenodoODBC Unicode(x64)};"
    "DATABASE=ldw;"
    "SERVER=virtualizador.sicredi.net;"
    "PORT=9996;"
    "Trusted_Connection=yes;"
)

try:
    conn = pyodbc.connect(conn_str)
    print("Conexão bem-sucedida com autenticação integrada!")
except Exception as e:
    print("Erro na conexão:", e)
