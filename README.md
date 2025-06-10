# Projeto de Extração de Dados do Denodo

Este projeto contém um robô em Python que:

- Conecta ao banco Denodo via ODBC.
- Executa consultas SQL.
- Converte os dados em um DataFrame.
- Salva os dados em arquivos **CSV** e **Excel (.xlsx)**.
- Armazena os arquivos em uma pasta do SharePoint sincronizada.

## Estrutura

- `scripts/`: scripts Python.
- `config/`: arquivos de configuração, como `.env`.
- `data/`: dados extraídos (opcional).
- `README.md`: documentação do projeto.

## Pré-requisitos

- Python 3.10 ou superior
- Driver ODBC do Denodo instalado
- Conta com acesso ao Denodo
- Pasta do SharePoint sincronizada localmente

## Instalação

1. Clone o repositório.
2. Crie um arquivo `.env` dentro da pasta `config/` com as seguintes variáveis:
```bash
LDAP_USER=seu_usuario 
PWD=sua_senha
```
3. Instale as dependências:

```bash
pip install python-dotenv pyodbc pandas openpyxl
```

## Como Usar
Execute o script principal:
```python scripts/extrair_dados.py```

O script irá:
- Conectar ao Denodo usando as credenciais do .env.
- Executar a consulta SQL definida no código.
- Salvar os dados em dois formatos:
    - ``carteira_por_gestor.csv``
    - ``carteira_por_gestor.xlsx``
- Armazenar os arquivos na pasta do SharePoint definida no script.

### Observações:
- Se o módulo ``iopenpyxl`` não estiver instalado, o script emitirá uma mensagem de aviso e continuará normalmente após salvar o ``.csv``.
- O caminho da pasta SharePoint pode ser ajustado diretamente no script.
---

