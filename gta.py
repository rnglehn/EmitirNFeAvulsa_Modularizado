# gta.py

from pegar_dados_GTA import extrair_dados_gta_via_interface

def extract_gta_data() -> dict:
    """
    Extrai categorias e metadados do PDF da GTA via pegar_dados_GTA.
    Se não encontrar nenhuma categoria, apenas avisa e continua com lista vazia.
    """
    data = extrair_dados_gta_via_interface()

    # Se não houver categorias, em vez de lançar exceção, apenas avisa
    if not data.get('categorias'):
        print("⚠️ Atenção: nenhuma categoria extraída da GTA — continuando sem categorizar.")
        data['categorias'] = []

    return data
