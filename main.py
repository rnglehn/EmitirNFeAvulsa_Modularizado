#!/usr/bin/env python3

import os, sys, logging
from datetime import datetime
import tkinter as tk

from pegar_dados_GTA import extrair_dados_gta_via_interface
from pauta          import download_and_load_pauta
from report         import generate_report
from login          import perform_login_with_selenium, escolher_operacao_gui
from utils          import get_latest_file

# logs
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)
ts = datetime.now().strftime("%Y%m%d_%H%M%S")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    handlers=[
        logging.FileHandler(f"{LOG_DIR}/run_{ts}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

def selecionar_classe_gui(df_pauta):
    classes = sorted(df_pauta['Classe'].dropna().unique())
    root = tk.Tk()
    root.title("Seleção de Classe")
    root.geometry("300x150")
    tk.Label(root, text="📋 Escolha a classe de gado:").pack(pady=(10,5))
    var = tk.StringVar(value=classes[0])
    tk.OptionMenu(root, var, *classes).pack(pady=5)
    tk.Button(root, text="OK", command=root.quit).pack(pady=(5,10))
    root.mainloop()
    escolha = var.get()
    root.destroy()
    return escolha

def main():
    logging.info("▶️ Iniciando fluxo")
    dados = extrair_dados_gta_via_interface()
    if not dados.get('categorias'):
        logging.error("❌ GTA falhou."); sys.exit(1)
    logging.info("✅ GTA extraída")

    df_pauta = download_and_load_pauta()
    logging.info("✅ Pauta carregada")

    classe = selecionar_classe_gui(df_pauta)
    logging.info(f"📋 Classe: {classe}")

    excel_rel = generate_report(dados, classe, df_pauta)
    logging.info(f"✅ Relatório: {excel_rel}")

    try:
        cred = get_latest_file("Arquivos", ".xlsx")
        logging.info(f"🔐 Credenciais: {cred}")
    except FileNotFoundError as e:
        logging.error(f"❌ {e}"); sys.exit(1)

    cpf_p = dados.get('cpf_procedencia','')
    cpf_d = dados.get('cpf_destino','')
    if cpf_p and cpf_p == cpf_d:
        op = "REMESSA INTERNA DE TRANSFERÊNCIA DE BOVINO"
    else:
        op = escolher_operacao_gui()
    logging.info(f"🔄 Operação: {op}")

    driver, ok = perform_login_with_selenium(
        excel_path=cred,
        farm_name=dados.get('estabelecimento_procedencia',''),
        operacao=op
    )
    if not ok:
        logging.error("❌ Login/menu falhou."); sys.exit(1)

    logging.info("🏁 Concluído! Navegador aberto.")

if __name__ == "__main__":
    main()
