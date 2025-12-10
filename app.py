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
    Lê as abas com os nomes exatos fornecidos e cruza os dados.
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

    # --- AJUSTE DE NOMES ---
    # Mapear as colunas da aba Comissões para um padrão comum
    renomear_comissao = {
        "Data Processamento": "Data",
        "Sigla Técnico": "Técnico"
    }
    df_com.rename(columns=renomear_comissao, inplace=True)

    # Verifica se a troca funcionou
    if "Data" not in df_com.columns or "Técnico" not in df_com.columns:
        return False, f"Erro: Não achei as colunas 'Data Processamento' ou 'Sigla Técnico' na aba Comissões. Colunas lidas: {df_com.columns.tolist()}"

    # Na aba Aproveitamento, os nomes já devem ser "Data" e "Técnico".
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
        how='outer', 
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
            colunas_comissao = ["Data Processamento", "Nome do Arquivo", "Sigla Técnico", "Horas Vendidas"]
            df_comissao = pd.DataFrame(dados_comissao, columns=colunas_comissao)
            st.dataframe(df_comissao)
            
            if st.button("Gravar Comissões", key="btn_comissao"):
                with st.spinner("Enviando..."):
                    client = conectar_sheets(); aba = client.open_by_key(ID_PLANILHA_MESTRA).worksheet("Comissoes")
                    if not aba.get_all_values():
                        aba.append_row(colunas_comissao)
                    aba.append_rows(dados_comissao)
                    st.success("✅ Sucesso!")

# --- TAB 2: APROVEITAMENTO ---
with aba_aproveitamento:
    st.header("Extrator de Aproveitamento")
    arquivos_aprov = st.file_uploader("Upload Aproveitamento HTML", type=["html", "htm"], accept_multiple_files=True, key="uploader_aprov")
    
    if arquivos_aprov:
        dados_aprov = []
        for arquivo in arquivos_aprov:
            try:
                raw_data = arquivo.read()
                try: conteudo = raw_data.decode("utf-8")
                except:
                    try: conteudo = raw_data.decode("latin-1")
                    except: conteudo = raw_data.decode("utf-16")
                
                soup = BeautifulSoup(conteudo, "html.parser")
                tecnico_atual_aprov = None
                linhas = soup.find_all("tr")
                
                for linha in linhas:
                    texto_original = linha.get_text(separator=" ", strip=True).upper()
                    texto_limpo = remover_acentos(texto_original)
                    if "TOTAL FILIAL:" in texto_original: break
                    if "MECANICO" in texto_limpo and "TOT.MEC" not in texto_limpo:
                        try:
                            parte_direita = texto_limpo.split("MECANICO")[1]
                            parte_direita = parte_direita.replace(":", "").strip()
                            if "-" in parte_direita: tecnico_atual_aprov = parte_direita.split("-")[0].strip()
                            else: tecnico_atual_aprov = parte_direita.split()[0]
                        except: continue
                    if "TOT.MEC.:" in texto_original:
                        tecnico_atual_aprov = None; continue
                    if tecnico_atual_aprov:
                        celulas = linha.find_all("td")
                        if not celulas: continue
                        txt_cel0 = celulas[0].get_text(strip=True)
                        if re.match(r"\d{2}/\d{2}/\d{2}", txt_cel0):
                            try:
                                if len(celulas) >= 4:
                                    dados_aprov.append([txt_cel0.split()[0], arquivo.name, tecnico_atual_aprov, 
                                                      celulas[1].get_text(strip=True), 
                                                      celulas[2].get_text(strip=True), 
                                                      celulas[3].get_text(strip=True)])
                            except: continue
            except Exception as e: st.error(f"Erro leitura: {e}")

        if len(dados_aprov) > 0:
            # AJUSTE DE COLUNAS AQUI TAMBÉM
            colunas_aprov = ["Data", "Arquivo", "Técnico", "Disp", "TP", "TG"]
            df_aprov = pd.DataFrame(dados_aprov, columns=colunas_aprov)
            
            st.success(f"✅ Sucesso! {len(dados_aprov)} registros.")
            st.dataframe(df_aprov)
            
            if st.button("Gravar Aproveitamento", key="btn_aprov"):
                with st.spinner("Enviando..."):
                    client = conectar_sheets(); aba = client.open_by_key(ID_PLANILHA_MESTRA).worksheet("Aproveitamento")
                    if not aba.get_all_values():
                        aba.append_row(colunas_aprov)
                    aba.append_rows(dados_aprov)
                    st.success("✅ Gravado!")

# --- TAB 3: RELATÓRIO UNIFICADO ---
with aba_unificacao:
    st.header("🔗 Unificação de Dados (Comissões + Aproveitamento)")
    st.info("Este módulo lê 'Data Processamento' e 'Sigla Técnico' da aba Comissões e cruza com 'Data' e 'Técnico' da aba Aproveitamento.")
    
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("🚀 Gerar Relatório Unificado"):
            with st.spinner("Lendo planilhas e cruzando dados..."):
                sucesso, mensagem = processar_unificacao()
                if sucesso:
                    st.success(mensagem)
                    st.balloons()
                else:
                    st.error(mensagem)
                
