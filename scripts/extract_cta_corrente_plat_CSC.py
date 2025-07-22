import os
import pyodbc
import pandas as pd
import logging
from datetime import date, timedelta
from dotenv import load_dotenv

def log_info(msg):
    print(msg)
    logging.info(msg)

def log_error(msg):
    print(msg)
    logging.error(msg)

def configurar_logging(pasta_logs, nome_base):
    if not os.path.exists(pasta_logs):
        os.makedirs(pasta_logs)
    data_str = date.today().strftime("%d_%m_%Y")
    log_path = os.path.join(pasta_logs, f"{nome_base}_{data_str}.log")
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        encoding="utf-8"
    )

def carregar_credenciais(env_path="config/.env"):
    load_dotenv(dotenv_path=env_path)
    usuario = os.getenv("LDAP_USER")
    senha = os.getenv("PWD")
    if not usuario or not senha:
        raise ValueError("Credenciais não carregadas corretamente.")
    return usuario, senha

def conectar_denodo(usuario, senha):
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
    cursor.execute(query)
    colunas = [col[0] for col in cursor.description]
    dados = [tuple(row) for row in cursor.fetchall()]
    return pd.DataFrame(dados, columns=colunas)

def salvar_excel(df, caminho, nome_arquivo):
    if not os.path.exists(caminho):
        os.makedirs(caminho)
    caminho_arquivo = os.path.join(caminho, f"{nome_arquivo}.xlsx")
    try:
        df.to_excel(caminho_arquivo, index=False, engine="openpyxl")
        log_info(f"✅ Arquivo Excel salvo em: {caminho_arquivo}")
    except ModuleNotFoundError:
        log_error("❌ Erro: módulo 'openpyxl' não encontrado. Instale com: pip install openpyxl")

def main():
    pasta_saida = r"C:\Users\joao_loliveira\OneDrive - Sicredi\Marina Rocha da Silva - CSC - Centro de Serviços Compartilhados 4501\CONFERÊNCIA CADASTRO - ABERTURA PLATAFORMA\2025\Conferência Aberturas Plataforma - RPA"
    pasta_logs = os.path.join(pasta_saida, "logs")
    nome_base = "contas_correntes_abertas_plataforma"
    configurar_logging(pasta_logs, nome_base)

    try:
        usuario, senha = carregar_credenciais()
        log_info("🔐 Credenciais carregadas com sucesso.")

        with conectar_denodo(usuario, senha) as conn:
            cursor = conn.cursor()

            ontem = date.today() - timedelta(days=1)
            query = f"""
            SELECT
              conta_corrente_plataforma.cod_agencia AS cod_agencia,
              conta_corrente_plataforma.nome_associado AS nome_associado,
              conta_corrente_plataforma.num_conta AS num_conta,
              conta_corrente_plataforma.data_abertura AS data_abertura,
              conta_corrente_plataforma.data_associacao AS data_associacao,
              conta_corrente_plataforma.status_conta AS status_conta,
              conta_corrente_plataforma.nome_assistente_negocio AS nome_assistente_negocio,
              conta_corrente_plataforma.nome_marca AS nome_marca,
              conta_corrente_plataforma.nome_carteira AS nome_carteira
            FROM conta_corrente_plataforma
            WHERE nome_assistente_negocio <> 'Proprio Associado'
              AND data_abertura = DATE '{ontem}'
            """

            df = executar_consulta(cursor, query)
            log_info(f"📊 Dados carregados: {df.shape[0]} linhas, {df.shape[1]} colunas.")

            salvar_excel(df, pasta_saida, f"{nome_base}_{ontem.strftime('%d_%m_%Y')}")

    except Exception as e:
        log_error(f"❌ Erro na execução: {e}")

if __name__ == "__main__":
    main()

