import re
import time
import pandas as pd
import tkinter as tk

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from utils import normalize_text

URL_SEFAZ = "https://nfewebprodutor.sefaz.to.gov.br/nfeacontribuinte/servlet/logincontribuinte"
TIMEOUT   = 20

def get_credentials(farm_name: str, excel_path: str) -> tuple[str,str]:
    df = pd.read_excel(excel_path, sheet_name="Planilha1", engine="openpyxl")
    raw  = normalize_text(farm_name).replace("FAZENDA", "").strip()
    clean = re.sub(r"[^\w\s]", "", raw).upper()
    simple_key = clean.replace(" ", "")

    df["FARM_NORM"]   = df["FAZENDA"].astype(str).apply(normalize_text)\
                        .str.replace(r"[^\w\s]", "", regex=True).str.upper()
    df["FARM_SIMPLE"] = df["FARM_NORM"].str.replace(r"\s+", "", regex=True)

    match = df[df["FARM_SIMPLE"].str.contains(simple_key, na=False)]
    if match.empty:
        match = df[df["FARM_NORM"].str.contains(clean, na=False)]
    if match.empty:
        raise ValueError(f"Fazenda '{farm_name}' não encontrada.")
    row = match.iloc[0]

    ie_raw = str(row["INSCRICAO ESTADUAL"])
    ie = re.sub(r"[.\-/]", "", ie_raw)
    pwd = str(row["SENHA SEFAZ"])
    return ie, pwd

def escolher_operacao_gui():
    ops = [
        "REMESSA INTERNA DE TRANSFERÊNCIA DE BOVINO",
        "VENDA INTERNA DE BOVINO PARA ABATE",
        "VENDA INTERNA DE BOVINO PARA RECRIA, MONTARIA, TRAÇÃO E ENGORDA"
    ]
    root = tk.Tk()
    root.title("Selecione a Operação")
    root.geometry("450x160")
    tk.Label(root, text="📋 Escolha a operação:", wraplength=430).pack(pady=(10,5))
    var = tk.StringVar(value=ops[0])
    tk.OptionMenu(root, var, *ops).pack(pady=5)
    tk.Button(root, text="OK", command=root.quit).pack(pady=(5,10))
    root.mainloop()
    escolha = var.get()
    root.destroy()
    return escolha

def perform_login_with_selenium(excel_path: str, farm_name: str, operacao: str):
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    service = ChromeService(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=chrome_options)
    wait    = WebDriverWait(driver, TIMEOUT)

    try:
        driver.get(URL_SEFAZ)
        ie, pwd = get_credentials(farm_name, excel_path)

        # Preenche IE
        ie_field = wait.until(EC.element_to_be_clickable((By.ID, "vCONINSEST")))
        driver.execute_script("arguments[0].scrollIntoView(true);", ie_field)
        ie_field.clear(); ie_field.send_keys(ie)
        time.sleep(0.2)
        if re.sub(r"\D+", "", ie_field.get_attribute("value")) != ie:
            ie_field.send_keys(Keys.CONTROL, "a", Keys.DELETE, ie)
            time.sleep(0.2)
        if re.sub(r"\D+", "", ie_field.get_attribute("value")) != ie:
            driver.execute_script(
                "arguments[0].value=arguments[1];arguments[0].dispatchEvent(new Event('input'));",
                ie_field, ie
            )
            time.sleep(0.1)
        atual = ie_field.get_attribute("value")
        if re.sub(r"\D+", "", atual) != ie:
            raise RuntimeError(f"IE esperado '{ie}', mas site mostrou '{atual}'")

        # Preenche senha
        pwd_field = wait.until(EC.element_to_be_clickable((By.ID, "vSENHA")))
        driver.execute_script("arguments[0].scrollIntoView(true);", pwd_field)
        pwd_field.clear(); pwd_field.send_keys(pwd)

        # Submete
        pwd_field.send_keys(Keys.TAB); time.sleep(0.2); pwd_field.send_keys(Keys.ENTER)
        wait.until(EC.url_changes(URL_SEFAZ))

        # NF-E Avulsa
        menu = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#Smoothnavmenu1 > ul > li:nth-child(2) > a")
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", menu)
        menu.click()

        # Seleciona operação
        sel_el = wait.until(EC.element_to_be_clickable((By.ID, "vNFEAOPERACAOID")))
        Select(sel_el).select_by_visible_text(operacao)

        # Emissão (linha 6)
        btn = wait.until(EC.element_to_be_clickable((
            By.CSS_SELECTOR,
            "#TBNFE > tbody > tr:nth-child(6) > td:nth-child(2) > input:nth-child(2)"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn)
        btn.click()

    except Exception:
        import traceback; traceback.print_exc()
        print("❌ Erro no fluxo de login/menu. Navegador permanece aberto.")
        return driver, False

    return driver, True
