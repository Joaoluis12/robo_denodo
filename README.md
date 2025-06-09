# Projeto de Extração de Dados do Denodo

Este projeto contém um robô em Python que:

- Conecta ao banco Denodo via ODBC.
- Executa consultas SQL.
- Salva os dados em arquivos CSV.
- Armazena os arquivos em uma pasta do SharePoint sincronizada.

## Estrutura

- `scripts/`: scripts Python.
- `config/`: arquivos de configuração, como `.env`.
- `data/`: dados extraídos.

## Como usar

1. Crie um arquivo `.env` em `config/` com suas credenciais.
2. Instale as dependências com `pip install python-dotenv pyodbc pandas`.
3. Execute o script com `python scripts/extrair_dados.py`.
