import os
import subprocess

# Caminho da pasta onde estão os scripts
pasta_scripts = r"C:\Users\joao_loliveira\OneDrive - Sicredi\Robo Denodo Power BI\scripts"

# Lista todos os arquivos .py na pasta
scripts = [f for f in os.listdir(pasta_scripts) if f.endswith(".py") and f != "execute_all.py"]

# Executa cada script
for script in scripts:
    caminho_completo = os.path.join(pasta_scripts, script)
    print(f"🚀 Executando: {script}")
    subprocess.run(["python", caminho_completo], check=True)
