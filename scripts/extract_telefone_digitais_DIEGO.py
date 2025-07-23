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

def obter_data_referencia():
    hoje = date.today()
    if hoje.weekday() == 0:  # Segunda-feira
        return hoje - timedelta(days=3)  # Sexta-feira anterior
    else:
        return hoje - timedelta(days=2)

def main():
    pasta_saida = r"C:\Users\joao_loliveira\OneDrive - Sicredi\Marina Rocha da Silva - Boas vindas Botmaker - Sicredi X"
    pasta_logs = os.path.join(pasta_saida, "logs")
    nome_base = "telefones_e_cpfs_associacao_digital"
    configurar_logging(pasta_logs, nome_base)

    data_referencia = obter_data_referencia()
    data_formatada = data_referencia.strftime('%Y-%m-%d')

    query_telefone = """
    SELECT 
        pessoa_fisica_telefone_digital.cod_ddd AS cod_ddd, 
        pessoa_fisica_telefone_digital.num_telefone AS num_telefone, 
        pessoa_fisica_telefone_digital.num_cpf AS num_cpf, 
        (pessoa_fisica_telefone_digital.cod_ddd||pessoa_fisica_telefone_digital.num_telefone) AS telefone_com_ddd 
    FROM pessoa_fisica_telefone_digital
    """

    query_contas = f"""
    SELECT 
        conta_corrente_plataforma.cpf_cnpj AS cpf_cnpj, 
        conta_corrente_plataforma.data_associacao AS data_associacao, 
        conta_corrente_plataforma.nome_assistente_negocio AS nome_assistente_negocio
    FROM conta_corrente_plataforma
    WHERE nome_assistente_negocio = 'Proprio Associado'
      AND  data_associacao = DATE '{data_formatada}' 
    """

    try:
        usuario, senha = carregar_credenciais()
        log_info("🔐 Credenciais carregadas com sucesso.")
        with conectar_denodo(usuario, senha) as conn:
            cursor = conn.cursor()
            df_telefone = executar_consulta(cursor, query_telefone)
            df_contas = executar_consulta(cursor, query_contas)
            log_info(f"📞 Telefones carregados: {df_telefone.shape[0]} linhas.")
            log_info(f"📁 Contas carregadas: {df_contas.shape[0]} linhas.")

        cpfs_comuns = set(df_contas['cpf_cnpj']).intersection(set(df_telefone['num_cpf']))
        df_telefone_filtrado = df_telefone[df_telefone['num_cpf'].isin(cpfs_comuns)].drop_duplicates(subset=['num_cpf'])
        df_contas_filtrada = df_contas[df_contas['cpf_cnpj'].isin(cpfs_comuns)].drop_duplicates(subset=['cpf_cnpj'])

        df_resultado = pd.merge(
            df_telefone_filtrado,
            df_contas_filtrada,
            left_on="num_cpf",
            right_on="cpf_cnpj",
            how="inner"
        )[[ "num_cpf", "telefone_com_ddd", "data_associacao", "nome_assistente_negocio" ]]

        salvar_excel(df_resultado, pasta_saida, f"{nome_base}_{data_referencia.strftime('%d_%m_%Y')}")

    except Exception as e:
        log_error(f"❌ Erro na execução: {e}")

if __name__ == "__main__":
    main()
