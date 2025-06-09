from dotenv import load_dotenv
import os
import pyodbc
import pandas as pd
from datetime import datetime

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

consultas = {
    "relatorio_clientes": "SELECT * FROM ldw.tabela_clientes",
    "relatorio_produtos": "SELECT * FROM ldw.tabela_produtos"
}

pasta_destino = r"C:\\Users\\SEU_USUARIO\\SharePoint\\Pasta\\Relatorios"

def exportar_dados():
    with pyodbc.connect(conn_str) as conn:
        for nome, sql in consultas.items():
            df = pd.read_sql(sql, conn)
            nome_arquivo = f"{nome}_{datetime.today().strftime('%Y-%m-%d')}.csv"
            caminho_completo = os.path.join(pasta_destino, nome_arquivo)
            df.to_csv(caminho_completo, index=False, sep=';', encoding='utf-8-sig')
            print(f"Relatório salvo: {caminho_completo}")

if __name__ == "__main__":
    exportar_dados()
