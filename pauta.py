# pauta.py

import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from utils import normalize_text, get_latest_file

PASTA_DESTINO = os.path.join(os.getcwd(), "Pautas Fiscais")

def baixar_pauta():
    """
    Baixa a pauta fiscal via Selenium e fecha o navegador ao final.
    """
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    chrome_options = Options()
    prefs = {
        "download.default_directory": PASTA_DESTINO,
        "download.prompt_for_download": False,
        "directory_upgrade": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    try:
        driver.get("https://pautafiscal.sefaz.to.gov.br/secao/1/3")
        wait = WebDriverWait(driver, 20)
        botao = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[.//span[contains(text(),'Exportar Excel')]]")
        ))
        botao.click()
        print("⏳ Aguardando download do Excel...")
        time.sleep(10)

        # Renomeia o arquivo baixado
        arquivos = [f for f in os.listdir(PASTA_DESTINO) if f.lower().endswith(".xlsx")]
        arquivos.sort(key=lambda f: os.path.getmtime(os.path.join(PASTA_DESTINO, f)), reverse=True)
        if not arquivos:
            raise FileNotFoundError("❌ Nenhum .xlsx encontrado em Pautas Fiscais.")
        antigo = os.path.join(PASTA_DESTINO, arquivos[0])
        data = time.strftime("%d.%m.%Y")
        novo = os.path.join(PASTA_DESTINO, f"PAUTA FISCAL - {data}.xlsx")
        os.replace(antigo, novo)
        print(f"✅ Download concluído e salvo como: {novo}")
    finally:
        driver.quit()

def download_and_load_pauta() -> pd.DataFrame:
    """
    Baixa a pauta (se ainda não baixou hoje), carrega o Excel,
    adiciona colunas de normalização e retorna o DataFrame.
    """
    baixar_pauta()
    # pega o arquivo mais recente .xlsx na pasta
    arquivo = get_latest_file(PASTA_DESTINO, ext=".xlsx")
    print(f"📁 Pauta fiscal carregada: {arquivo}\n")

    # lê a aba 'Dados'
    df = pd.read_excel(arquivo, sheet_name="Dados")

    # normaliza descrições e classes
    df['Descricao_Normalizada'] = df['Descrição'].apply(normalize_text)
    df['Classe_Normalizada']   = df['Classe'].apply(normalize_text)

    return df
