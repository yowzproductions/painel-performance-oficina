import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import unicodedata

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Central de Relatórios WLM", layout="wide")
st.title("🏭 Central de Processamento de Relatórios")

# ID da sua planilha
ID_PLANILHA_MESTRA = "1XibBlm2x46Dk5bf4JvfrMepD4gITdaOtTALSgaFcwV0"

# --- FUNÇÕES AUXILIARES ---
def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def conectar_sheets():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

def processar_unificacao():
    """
    Lê as abas com os nomes exatos fornecidos pelo Ronaldo e cruza os dados.
    """
    client = conectar_sheets()
    sh = client.open_by_key(ID_PLANILHA_MESTRA)

    # 1. Ler as abas de origem
    try:
        ws_com = sh.worksheet("Comissoes")
        ws_aprov = sh.worksheet("Aproveitamento")
    except:
        return False, "Erro: As abas 'Comissoes' ou 'Aproveitamento' não foram encontradas."

    # 2. Ler os dados
    dados_com = ws_com.get_all_records()
    dados_aprov = ws_aprov.get_all_records()

    if not dados_com or not dados_aprov:
        return False, "Uma das abas está vazia. Faça upload dos arquivos primeiro."

    df_com = pd.DataFrame(dados_com)
    df_aprov = pd.DataFrame(dados_aprov)

    # 3. Limpeza de Nomes de Colunas (Remove espaços acidentais)
    df_com.columns = [c.strip() for c in df_com.columns]
    df_aprov.columns = [c.strip() for c in df_aprov.columns]

    # --- AJUSTE DE NOMES (O SEGREDO DO SUCESSO) ---
    # Mapear as colunas da aba Comissões para um padrão comum
    # "Data Processamento" vira "Data"
    # "Sigla Técnico" vira "Técnico"
    
    renomear_comissao = {
        "Data Processamento": "Data",
        "Sigla Técnico": "Técnico"
    }
    df_com.rename(columns=renomear_comissao, inplace=True)

    # Verifica se a troca funcionou (ou seja, se os nomes estavam certos na planilha)
    if "Data" not in df_com.columns or "Técnico" not in df_com.columns:
        return False, f"Erro: Não achei as colunas 'Data Processamento' ou 'Sigla Técnico' na aba Comissões. Colunas lidas: {df_com.columns.tolist()}"

    # Na aba Aproveitamento, os nomes já devem ser "Data" e "Técnico".
    # Se não forem, o merge vai falhar, então vamos garantir.
    if "Data" not in df_aprov.columns or "Técnico" not in df_aprov.columns:
        return False, f"Erro: Não achei as colunas 'Data' ou 'Técnico' na aba Aproveitamento. Colunas lidas: {df_aprov.columns.tolist()}"

    # 4. Padronização de Tipos (Texto)
    df_com['Data'] = df_com['Data'].astype(str)
    df_com['Técnico'] = df_com['Técnico'].astype(str)
    df_aprov['Data'] = df_aprov['Data'].astype(str)
    df_aprov['Técnico'] = df_aprov['Técnico'].astype(str)

    # 5. O Merge (Cruzamento)
    # Une as duas tabelas usando Data e Técnico como chave
    df_final = pd.merge(
        df_com, 
        df_aprov, 
        on=['Data', 'Técnico'], 
        how='outer', # Mantém tudo
        suffixes=('_Comissao', '_Aprov')
    )
    
    df_final.fillna("", inplace=True)

    # 6. Salvar na aba 'Consolidado'
    try:
        ws_final = sh.worksheet("Consolidado")
        ws_final.clear()
    except:
        ws_final = sh.add_worksheet(title="Consolidado", rows=1000, cols=20)
    
    ws_final.update([df_final.columns.values.tolist()] + df_final.values.tolist())
    
    return True, f"Sucesso! {len(df_final)} linhas consolidadas."

# --- INTERFACE (TABS) ---
aba_comissoes, aba_aproveitamento, aba_unificacao = st.tabs([
    "💰 Pagamento de Comissões", 
    "⚙️ Aproveitamento Técnico",
    "📊 Relatório Unificado"
])

# --- TAB 1: COMISSÕES ---
with aba_comissoes:
    st.header("Processador de Comissões")
    arquivos_comissao = st.file_uploader("Upload Comissões HTML", type=["html", "htm"], accept_multiple_files=True, key="uploader_comissao")
    if arquivos_comissao:
        dados_comissao = []
        st.write(f"📂 Processando {len(arquivos_comissao)} arquivos...")
        for arquivo in arquivos_comissao:
            try:
                try: conteudo = arquivo.read().decode("utf-8")
                except: 
                    arquivo.seek(0)
                    conteudo = arquivo.read().decode("latin-1")
                soup = BeautifulSoup(conteudo, "html.parser")
                texto_completo = soup.get_text(separator=" ", strip=True)
                match_data = re.search(r"até\s+(\d{2}/\d{2}/\d{4})", texto_completo, re.IGNORECASE)
                data_relatorio = match_data.group(1) if match_data else datetime.now().strftime("%d/%m/%Y")
                tecnico_atual = None
                for linha in soup.find_all("tr"):
                    texto_linha = linha.get_text(separator=" ", strip=True).upper()
                    if "TOTAL DA FILIAL" in texto_linha or "TOTAL DA EMPRESA" in texto_linha: break
                    if "TOTAL DO FUNCIONARIO" in texto_linha:
                        try: tecnico_atual = texto_linha.split("TOTAL DO FUNCIONARIO")[1].replace(":", "").strip().split()[0]
                        except: continue 
                    if tecnico_atual and "HORAS VENDIDAS:" in texto_linha:
                        celulas = linha.find_all("td")
                        for celula in celulas:
                            txt = celula.get_text(strip=True).upper()
                            if "HORAS" in txt and any(c.isdigit() for c in txt) and "VENDIDAS" not in txt:
                                dados_comissao.append([data_relatorio, arquivo.name, tecnico_atual, txt.replace("HORAS", "").strip()])
                                break 
            except Exception as e: st.error(f"Erro: {e}")

        if len(dados_comissao) > 0:
            # AJUSTE DE COLUNAS AQUI TAMBÉM
