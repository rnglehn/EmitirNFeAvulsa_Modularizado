# pegar_dados_GTA.py

import os
import json
import tkinter as tk
from tkinter import filedialog
import fitz  # PyMuPDF
import re

def selecionar_pdf():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    try:
        root.eval('tk::PlaceWindow . center')
    except tk.TclError:
        root.update_idletasks()
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        root.geometry(f"+{sw//2}+{sh//2}")
    caminho = filedialog.askopenfilename(
        parent=root,
        filetypes=[("PDF","*.pdf")],
        title="Selecione o PDF da GTA"
    )
    root.destroy()
    return caminho

def ler_pdf(caminho):
    texto = ""
    with fitz.open(caminho) as doc:
        for pag in doc:
            texto += pag.get_text()
    return texto

def extrair_dados_gta_via_interface():
    caminho = selecionar_pdf()
    if not caminho:
        print("❌ Nenhum arquivo selecionado.")
        return {}

    texto = ler_pdf(caminho)
    linhas = texto.splitlines()
    dados = {}

    # Número da GTA
    m_num = re.search(r'\bNumero\b.*?(\d{5,})', texto, re.IGNORECASE | re.DOTALL)
    dados["numero_gta"] = m_num.group(1).strip() if m_num else None

    # UF
    dados["uf"] = None
    for i, ln in enumerate(linhas):
        if ln.strip().upper() == "UF":
            for nx in linhas[i+1:i+6]:
                v = nx.strip()
                if re.fullmatch(r'[A-Z]{2}', v):
                    dados["uf"] = v
                    break
            break

    # Série
    dados["serie"] = None
    for i, ln in enumerate(linhas):
        if ln.strip().upper() in ("SÉRIE", "SERIE"):
            for nx in linhas[i+1:i+6]:
                v = nx.strip()
                if len(v)==1 and v.isalnum():
                    dados["serie"] = v
                    break
            break

    # Validade
    m_val = re.search(r'Validade\s*[:\-]\s*(\d{2}/\d{2}/\d{4})', texto)
    dados["validade"] = m_val.group(1) if m_val else None

    # CPF/CNPJ
    cpfs = re.findall(r'CPF/CNPJ[:]\s*(\d+)', texto)
    dados["cpf_procedencia"] = cpfs[0] if len(cpfs)>0 else None
    dados["cpf_destino"]     = cpfs[1] if len(cpfs)>1 else None

    # Nomes
    nomes = re.findall(r'Nome[:]\s*(.+)', texto)
    dados["nome_procedencia"] = nomes[0].strip() if len(nomes)>0 else None
    dados["nome_destino"]     = nomes[1].strip() if len(nomes)>1 else None

    # Estabelecimentos
    ests = re.findall(r'Estabelecimento[:]\s*(.+)', texto)
    dados["estabelecimento_procedencia"] = ests[0].strip() if len(ests)>0 else None
    dados["estabelecimento_destino"]     = ests[1].strip() if len(ests)>1 else None

    # Municípios
    muns = re.findall(r'Município - UF[:]\s*(.+)', texto)
    dados["municipio_procedencia"] = muns[0].strip() if len(muns)>0 else None
    dados["municipio_destino"]     = muns[1].strip() if len(muns)>1 else None

    # Finalidade
    m_fin = re.search(r'Finalidade[:]\s*(.+?)\s+Meio de Transporte', texto, re.DOTALL)
    dados["finalidade"] = m_fin.group(1).strip() if m_fin else None

    # Extrai categorias: layout vertical
    categorias = []
    # localiza índice do header "Grupo"
    for idx, ln in enumerate(linhas):
        if ln.strip().upper() == "GRUPO":
            # data começa 6 linhas adiante
            start = idx + 6
            # enquanto houver blocos completos
            while start + 5 < len(linhas):
                grp = linhas[start].strip()
                esp = linhas[start+1].strip()    # ESPÉCIE correta
                cat = linhas[start+2].strip() if linhas[start+2].strip()!='-' else None
                fx  = linhas[start+3].strip()
                sx  = linhas[start+4].strip()
                qt  = re.sub(r'\D+', '', linhas[start+5].strip())
                if not grp or not esp or not fx or not qt:
                    break
                try:
                    qtd = int(qt)
                except:
                    break
                categorias.append({
                    "grupo": grp,
                    "especie": esp,
                    "categoria": cat,
                    "faixa": fx,
                    "sexo": sx,
                    "quantidade": qtd
                })
                start += 6
            break

    # opções horizontais (fallback)
    padrao_h = re.findall(
        r'(Bovideos|Bovídeos)\s+(Bovinos)\s+\-\s+(.+?)\s+(Macho|Fêmea|Femea)\s+(\d+)',
        texto, re.IGNORECASE
    )
    for grp, esp, fx, sx, qt in padrao_h:
        categorias.append({
            "grupo": grp,
            "especie": esp,
            "categoria": None,
            "faixa": fx,
            "sexo": sx,
            "quantidade": int(qt)
        })

    # remove duplicados
    unicos = []
    seen = set()
    for c in categorias:
        key = (c['especie'], c.get('categoria') or '', c['faixa'], c['sexo'], c['quantidade'])
        if key not in seen:
            seen.add(key)
            unicos.append(c)
    dados["categorias"] = unicos

    # salva JSON
    pasta = "JSON"
    os.makedirs(pasta, exist_ok=True)
    base = os.path.splitext(os.path.basename(caminho))[0]
    out = os.path.join(pasta, f"{base}_dados.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    print(f"✅ Dados da GTA extraídos e salvos em: {out}")
    return dados
