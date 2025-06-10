import os
import pyodbc
import pandas as pd
from dotenv import load_dotenv

def carregar_credenciais(env_path="config/.env"):
    """Carrega as credenciais do arquivo .env"""
    load_dotenv(dotenv_path=env_path)
    usuario = os.getenv("LDAP_USER")
    senha = os.getenv("PWD")

    if not usuario or not senha:
        raise ValueError("Credenciais não carregadas corretamente.")
    
    return usuario, senha

def conectar_denodo(usuario, senha):
    """Estabelece conexão com o Denodo"""
    conn_str = (
        "DRIVER={DenodoODBC Unicode(x64)};"
        "DATABASE=ldw;"
        "SERVER=virtualizador.sicredi.net;"
        "PORT=9996;"
        f"UID={usuario};"
        f"PWD={senha}"
    )
    return pyodbc.connect(conn_str)

def executar_consulta(cursor, query):
    """Executa a consulta SQL e retorna os dados como DataFrame"""
    cursor.execute(query)
    colunas = [col[0] for col in cursor.description]
    dados = [tuple(row) for row in cursor.fetchall()]
    return pd.DataFrame(dados, columns=colunas)

def salvar_excel(df, caminho, nome_arquivo):
    """Salva o DataFrame em formato Excel"""
    if not os.path.exists(caminho):
        os.makedirs(caminho)

    caminho_arquivo = os.path.join(caminho, f"{nome_arquivo}.xlsx")
    try:
        df.to_excel(caminho_arquivo, index=False, engine="openpyxl")
        print(f"✅ Arquivo Excel salvo em: {caminho_arquivo}")
    except ModuleNotFoundError:
        print("❌ Erro: módulo 'openpyxl' não encontrado. Instale com: pip install openpyxl")

def main():
    try:
        usuario, senha = carregar_credenciais()
        print("🔐 Credenciais carregadas com sucesso.")

        with conectar_denodo(usuario, senha) as conn:
            cursor = conn.cursor()
            query = """
                SELECT
                    ag_cart.ano_mes_data AS ano_mes_data,
                    ag_cart.competencia AS competencia,
                    ag_cart.coop AS coop,
                    ag_cart.agencia AS agencia,
                    ag_cart.carteira AS carteira,
                    ag_cart.des_pessoa AS des_pessoa,
                    ag_cart.des_produto AS des_produto,
                    ag_cart.des_produto2 AS des_produto2,
                    ag_cart.vlr_planejado AS vlr_planejado,
                    ag_cart.vlr_realizado AS vlr_realizado,
                    ag_cart.vlr_dif AS vlr_dif,
                    ag_cart.vlr_realizado_saldo AS vlr_realizado_saldo,
                    ag_cart.qt_produto AS qt_produto,
                    ag_cart.central AS central
                FROM ag_cart
                WHERE ano_mes > 202412
            """
            df = executar_consulta(cursor, query)
            print(f"📊 Dados carregados: {df.shape[0]} linhas, {df.shape[1]} colunas.")

            salvar_excel(
                df,
                caminho=r"C:\Users\joao_loliveira\OneDrive - Sicredi\Power BI 4501 - Geral\Bases Denodo RPA",
                nome_arquivo="ag_cart_2025"
            )

    except Exception as e:
        print(f"❌ Erro na execução: {e}")

if __name__ == "__main__":
    main()
