import os
import unicodedata
import re

def normalize_text(text: str) -> str:
    """
    Converte para maiúsculas, remove acentos e normaliza espaços.
    Exemplo: "Olho D'Água" → "OLHO D AGUA"
    """
    s = str(text).upper()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return ' '.join(s.split())

def limpar_nome(text: str) -> str:
    """
    Prepara uma string para nomes de arquivo:
    remove acentos, pontuação e troca espaços por underscore.
    """
    s = normalize_text(text)
    s = re.sub(r'[\\/.,:;()\[\]\-]', '', s)
    return '_'.join(s.split())

def get_latest_file(folder: str, ext: str = ".xlsx") -> str:
    """
    Retorna o caminho do arquivo mais recente em `folder` com extensão `ext`,
    ignorando arquivos que comecem com '~$' (temp do Excel).
    """
    arquivos = [
        f for f in os.listdir(folder)
        if f.lower().endswith(ext) and not f.startswith("~$")
    ]
    if not arquivos:
        raise FileNotFoundError(f"Não há arquivos {ext} em {folder}")
    arquivos.sort(
        key=lambda x: os.path.getmtime(os.path.join(folder, x)),
        reverse=True
    )
    return os.path.join(folder, arquivos[0])
