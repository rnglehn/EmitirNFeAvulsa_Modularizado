import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Pasta para salvar os xlsx
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

        # Renomeia o arquivo mais recente
        arquivos = [f for f in os.listdir(PASTA_DESTINO) if f.lower().endswith(".xlsx")]
        arquivos.sort(key=lambda f: os.path.getmtime(os.path.join(PASTA_DESTINO, f)), reverse=True)
        if not arquivos:
            raise FileNotFoundError("❌ Nenhum .xlsx encontrado em Pautas Fiscais.")
        antigo = os.path.join(PASTA_DESTINO, arquivos[0])
        data = datetime.now().strftime("%d.%m.%Y")
        novo = os.path.join(PASTA_DESTINO, f"PAUTA FISCAL - {data}.xlsx")
        os.replace(antigo, novo)
        print(f"✅ Download concluído e salvo como: {novo}")

    finally:
        # **Aqui fechamos o navegador do Selenium**
        driver.quit()
