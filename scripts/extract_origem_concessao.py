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
                origem_concessao.ano_mes AS ano_mes,
                origem_concessao.cod_cooperativa AS cod_cooperativa,
                origem_concessao.cpf_cnpj AS cpf_cnpj,
                origem_concessao.cod_agencia AS cod_agencia,
                origem_concessao.num_conta AS num_conta,
                origem_concessao.cod_carteira AS cod_carteira,
                origem_concessao.valor_concessao AS valor_concessao,
                origem_concessao.dat_concessao AS dat_concessao,
                origem_concessao.dat_contratacao AS dat_contratacao,
                origem_concessao.des_produto AS des_produto,
                origem_concessao.grupo_produto AS grupo_produto,
                origem_concessao.renda_mensal AS renda_mensal,
                origem_concessao.num_parcelas AS num_parcelas,
                origem_concessao.taxa_juros_normal AS taxa_juros_normal,
                (origem_concessao.cod_agencia || origem_concessao.cod_carteira) AS codigo_carteira
            FROM origem_concessao
            """
            df = executar_consulta(cursor, query)
            print(f"📊 Dados carregados: {df.shape[0]} linhas, {df.shape[1]} colunas.")

            salvar_excel(
                df,
                caminho=r"C:\Users\joao_loliveira\OneDrive - Sicredi\Power BI 4501 - Geral\Bases Denodo RPA",
                nome_arquivo="origem_concessao_2025"
            )

    except Exception as e:
        print(f"❌ Erro na execução: {e}")

if __name__ == "__main__":
    main()
