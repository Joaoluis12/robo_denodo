import os
import pyodbc
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime, timedelta

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

def salvar_por_agencia(df):
    data_hoje = datetime.today().strftime('%d-%m-%Y')
    data_ontem = (datetime.today() - timedelta(days=1)).strftime('%d-%m-%Y')
    base_path = r"C:\Users\joao_loliveira\OneDrive - Sicredi\INAD - Relatórios Denodo - Base 4501"

    for agencia in df['cod_agencia'].unique():
        df_agencia = df[df['cod_agencia'] == agencia]
        pasta_agencia = os.path.join(base_path, f"{agencia}_RPA")

        arquivo_ontem = os.path.join(pasta_agencia, f"BASE_INADIMPLENTES_{data_ontem}.xlsx")
        if os.path.exists(arquivo_ontem):
            os.remove(arquivo_ontem)
            print(f"🗑️ Arquivo antigo removido: {arquivo_ontem}")

        arquivo_hoje = os.path.join(pasta_agencia, f"BASE_INADIMPLENTES_{data_hoje}.xlsx")
        df_agencia.to_excel(arquivo_hoje, index=False, engine="openpyxl")
        print(f"✅ Novo arquivo salvo: {arquivo_hoje}")

def main():
    try:
        usuario, senha = carregar_credenciais()
        print("🔐 Credenciais carregadas com sucesso.")

        with conectar_denodo(usuario, senha) as conn:
            cursor = conn.cursor()
            query = """
            SELECT 
                base_inadimplentes.dt_base AS dt_base,
                base_inadimplentes.dt_liberacao AS dt_liberacao,
                base_inadimplentes.dt_vencimento AS dt_vencimento,
                base_inadimplentes.cod_agencia AS cod_agencia,
                base_inadimplentes.cod_carteira AS cod_carteira,
                base_inadimplentes.num_conta AS num_conta,
                base_inadimplentes.contrato AS contrato,
                base_inadimplentes.nome_associado AS nome_associado,
                base_inadimplentes.ds_produto AS ds_produto,
                base_inadimplentes.faixa_atraso AS faixa_atraso,
                base_inadimplentes.prejuizo AS prejuizo,
                base_inadimplentes.risco_atual AS risco_atual,
                base_inadimplentes.assessoria AS assessoria,
                base_inadimplentes.fone_1 AS fone_1,
                base_inadimplentes.fone_2 AS fone_2,
                base_inadimplentes.fone_3 AS fone_3,
                base_inadimplentes.flag_falecido AS flag_falecido,
                base_inadimplentes.ajuizado AS ajuizado,
                base_inadimplentes.flag_origem AS flag_origem,
                base_inadimplentes.atraso AS atraso,
                base_inadimplentes.saldo_online AS saldo_online,
                base_inadimplentes.qtd_parcelas AS qtd_parcelas,
                base_inadimplentes.qtd_vencidas AS qtd_vencidas,
                base_inadimplentes.qtd_pagas AS qtd_pagas,
                base_inadimplentes.valor_liberado_credito AS valor_liberado_credito,
                base_inadimplentes.saldo_cart_contrato AS saldo_cart_contrato,
                base_inadimplentes.saldo_cart_cpf AS saldo_cart_cpf,
                base_inadimplentes.saldo_prej_contrato AS saldo_prej_contrato,
                base_inadimplentes.saldo_prej_cpf AS saldo_prej_cpf
            FROM base_inadimplentes
            """
            df = executar_consulta(cursor, query)
            print(f"📊 Dados carregados: {df.shape[0]} linhas, {df.shape[1]} colunas.")

            salvar_por_agencia(df)

    except Exception as e:
        print(f"❌ Erro na execução: {e}")

if __name__ == "__main__":
    main()
