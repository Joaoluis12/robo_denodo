import os
import pyodbc
import pandas as pd
from dotenv import load_dotenv

# Carrega variáveis de ambiente
load_dotenv(dotenv_path="config/.env")
usuario = os.getenv("LDAP_USER")
senha = os.getenv("PWD")

print("Usuário:", usuario)
print("Senha:", "****" if senha else "Não carregada")

# Conexão com Denodo
conn_str = (
    "DRIVER={DenodoODBC Unicode(x64)};"
    "DATABASE=ldw;"
    "SERVER=virtualizador.sicredi.net;"
    "PORT=9996;"
    f"UID={usuario};"
    f"PWD={senha}"
)

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    query = """
    SELECT 
        carteira_por_gestor.coop AS coop,
        carteira_por_gestor.cod_ua AS cod_ua,
        carteira_por_gestor.cod_carteira AS cod_carteira,
        carteira_por_gestor.des_carteira AS des_carteira,
        carteira_por_gestor.des_gestor AS des_gestor
    FROM carteira_por_gestor
    """

    cursor.execute(query)
    columns = [col[0] for col in cursor.description]
    rows = [tuple(row) for row in cursor.fetchall()]

    print("Número de colunas detectadas:", len(columns))
    print("Colunas:", columns)
    print("Primeira linha:", rows[0] if rows else "Nenhum dado retornado")

    df = pd.DataFrame(rows, columns=columns)

    # Caminho de destino
    path_sharepoint = r"C:\Users\joao_loliveira\OneDrive - Sicredi\Power BI 4501 - Geral\Relatorio Diario"
    nome_arquivo = "carteira_por_gestor"

    if not os.path.exists(path_sharepoint):
        os.makedirs(path_sharepoint)

    # Salvar CSV
    csv_path = os.path.join(path_sharepoint, f"{nome_arquivo}.csv")
    df.to_csv(csv_path, index=False)
    print(f"✅ Arquivo CSV salvo em: {csv_path}")

    # Salvar XLSX
    try:
        excel_path = os.path.join(path_sharepoint, f"{nome_arquivo}.xlsx")
        df.to_excel(excel_path, index=False, engine="openpyxl")
        print(f"✅ Arquivo Excel salvo em: {excel_path}")
    except ModuleNotFoundError:
        print("❌ Erro: módulo 'openpyxl' não encontrado. Instale com: pip install openpyxl")

    cursor.close()
    conn.close()

except Exception as e:
    print(f"❌ Erro na execução: {e}")