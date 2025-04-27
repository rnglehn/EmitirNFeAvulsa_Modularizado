# pegar_dados_PAUTA_FISCAL.py

import os
import time
import pandas as pd
from datetime import date
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from utils import normalize_text, get_latest_file

# Pasta onde as pautas fiscais são salvas
PASTA_DESTINO = os.path.join(os.getcwd(), "Pautas Fiscais")
# Diretório de download (mesmo que PASTA_DESTINO)
DOWNLOAD_DIR = PASTA_DESTINO
# Tempo máximo para aguardar o download (em segundos)
DOWNLOAD_TIMEOUT = 10


def baixar_pauta():
    """
    Abre o site de consulta de pauta via Selenium (URL_DA_PAGINA_DE_PAUTAS),
    faz login se necessário, clica em "Exportar Excel" e aguarda o .xlsx ser baixado.
    Retorna o caminho completo do arquivo baixado.
    """
    os.makedirs(PASTA_DESTINO, exist_ok=True)

    # Configura o Chrome para download automático
    options = Options()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)

    # Acessa a página de pautas
    driver.get("URL_DA_PAGINA_DE_PAUTAS")  # Ajuste para a URL correta
    # TODO: implemente login via Selenium, se necessário

    # Clica no botão Exportar Excel
    btn_export = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[.//span[text()='Exportar Excel']]")
    ))
    btn_expaort.click()

    # Aguarda até DOWNLOAD_TIMEOUT segundos pelo download do arquivo .xlsx
    deadline = time.time() + DOWNLOAD_TIMEOUT
    arquivo_baixado = None
    while time.time() < deadline:
        xls = [f for f in os.listdir(PASTA_DESTINO) if f.lower().endswith('.xlsx')]
        if xls:
            xls.sort(key=lambda f: os.path.getmtime(os.path.join(PASTA_DESTINO, f)), reverse=True)
            latest = xls[0]
            # Desconsidera arquivos temporários do Excel
            if not latest.startswith('~$'):
                arquivo_baixado = os.path.join(PASTA_DESTINO, latest)
                break
        time.sleep(1)

    driver.quit()
    if not arquivo_baixado:
        raise FileNotFoundError(f"❌ Timeout de {DOWNLOAD_TIMEOUT}s ao baixar pauta fiscal")

    print(f"✅ Download concluído e salvo como: {arquivo_baixado}")
    return arquivo_baixado


def download_and_load_pauta() -> pd.DataFrame:
    """
    Verifica se já existe arquivo .xlsx da pauta fiscal de hoje em PASTA_DESTINO.
    - Se sim, retorna esse arquivo;
    - Caso contrário, chama baixar_pauta().

    Em seguida, lê a aba 'Dados', normaliza colunas e retorna o DataFrame.
    """
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    hoje = date.today()
    arquivo_local = None

    # Procura por arquivo local modificado hoje
    for f in os.listdir(PASTA_DESTINO):
        if f.lower().endswith('.xlsx') and not f.startswith('~$'):
            caminho = os.path.join(PASTA_DESTINO, f)
            try:
                if date.fromtimestamp(os.path.getmtime(caminho)) == hoje:
                    arquivo_local = caminho
                    break
            except OSError:
                continue

    if arquivo_local:
        print(f"📁 Pauta fiscal de hoje já existe: {arquivo_local}")
        arquivo = arquivo_local
    else:
        arquivo = baixar_pauta()

    print(f"📁 Pauta fiscal carregada: {arquivo}\n")

    # Lê e normaliza
    df = pd.read_excel(arquivo, sheet_name="Dados")
    df['Descricao_Normalizada'] = df['Descrição'].apply(normalize_text)
    df['Classe_Normalizada']   = df['Classe'].apply(normalize_text)
    return df