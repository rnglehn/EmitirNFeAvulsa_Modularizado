# report.py

import os
import pandas as pd
from datetime import datetime
from utils import normalize_text, limpar_nome

def generate_report(dados: dict, classe_escolhida: str, df_pauta: pd.DataFrame) -> str:
    """
    1) Converte dados['categorias'] em DataFrame
    2) Normaliza descrição + classe
    3) Merge com df_pauta para puxar o Valor
    4) Remove linhas com quantidade zero
    5) Adiciona data e formata colunas finais
    """

    # Data para preencher na coluna Data
    hoje = datetime.now().strftime('%d/%m/%Y')

    # Normalização da classe
    classe_norm = normalize_text(classe_escolhida)

    # 1) DataFrame das categorias extraídas do JSON
    df_cat = pd.DataFrame(dados.get('categorias', []))
    if df_cat.empty:
        print("❌ Sem categorias na GTA.")
        return

    # Coluna Classe (texto original) + Normalizada
    df_cat['Classe'] = classe_escolhida
    df_cat['Classe_Normalizada'] = df_cat['Classe'].apply(normalize_text)

    # 2) Constrói a Descricao_Normalizada exatamente igual à pauta
    df_cat['Descricao_Normalizada'] = (
        df_cat['especie'].astype(str) + " " +
        df_cat['sexo'].astype(str)    + " " +
        df_cat['faixa'].astype(str)
    ).apply(normalize_text)

    # ** Novo passo: filtra só as categorias com quantidade > 0 **
    df_cat = df_cat[df_cat['quantidade'] > 0]
    if df_cat.empty:
        print("❌ Todas as categorias tinham quantidade zero.")
        return

    # 3) Merge com a pauta fiscal
    df_merge = pd.merge(
        df_cat,
        df_pauta[['Descricao_Normalizada', 'Classe_Normalizada', 'Valor']],
        on=['Descricao_Normalizada', 'Classe_Normalizada'],
        how='left'
    )

    # 4) Monta as colunas finais
    df_merge['Preço Pauta Fiscal'] = df_merge['Valor']
    df_merge['Data'] = hoje

    # 5) Seleciona e renomeia colunas para a aba Produtos – Animais
    df_prod = df_merge[[
        'especie', 'sexo', 'faixa', 'quantidade',
        'Classe', 'Preço Pauta Fiscal', 'Data'
    ]].rename(columns={
        'especie':    'Espécie',
        'sexo':       'Sexo',
        'faixa':      'Faixa',
        'quantidade': 'Quantidade'
    })

    # 6) Grava o Excel com duas abas
    os.makedirs("Relatórios", exist_ok=True)
    num = dados.get('numero_gta', 'SEM_NUMERO')
    ts  = datetime.now().strftime('%d.%m.%Y_%H-%M-%S')
    nome_arquivo = (
        f"Relatórios/RELATORIO_NFE_GTA_{num}_"
        f"{limpar_nome(dados.get('nome_procedencia', ''))}_"
        f"{ts}.xlsx"
    )

    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        #
