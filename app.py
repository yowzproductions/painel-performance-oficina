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
    """Lê Comissões e Aproveitamento, une os dados e salva na aba Consolidado."""
    client = conectar_sheets()
    sh = client.open_by_key(ID_PLANILHA_MESTRA)

    # 1. Ler as abas de origem
    try:
        ws_com = sh.worksheet("Comissoes")
        ws_aprov = sh.worksheet("Aproveitamento")
    except:
        return False, "Erro: Certifique-se que as abas 'Comissoes' e 'Aproveitamento' existem e têm dados."

    # 2. Converter para DataFrame
    df_com = pd.DataFrame(ws_com.get_all_records())
    df_aprov = pd.DataFrame(ws_aprov.get_all_records())

    if df_com.empty or df_aprov.empty:
        return False, "Uma das abas (Comissoes ou Aproveitamento) está vazia."

    # 3. Padronização para o Cruzamento (ETL)
    # Renomeia 'Data Ref.' para 'Data' para bater com a outra tabela
    if 'Data Ref.' in df_com.columns:
        df_com.rename(columns={'Data Ref.': 'Data'}, inplace=True)
    
    # Converte chaves para texto (String) para evitar erro de tipo
    df_com['Data'] = df_com['Data'].astype(str)
    df_com['Técnico'] = df_com['Técnico'].astype(str)
    df_aprov['Data'] = df_aprov['Data'].astype(str)
    df_aprov['Técnico'] = df_aprov['Técnico'].astype(str)

    # 4. O Merge (Cruzamento)
    # Usa 'Data' e 'Técnico' como âncoras. 'outer' garante que ninguém suma.
    df_final = pd.merge(
        df_com, 
        df_aprov, 
        on=['Data', 'Técnico'], 
        how='outer', 
        suffixes=('_Comissao', '_Aprov')
    )
    
    # Preenche vazios com string vazia
    df_final.fillna("", inplace=True)

    # 5. Salvar na aba 'Consolidado'
    try:
        ws_final = sh.worksheet("Consolidado")
        ws_final.clear()
    except:
        # Se não existir, cria a aba
        ws_final = sh.add_worksheet(title="Consolidado", rows=1000, cols=20)
    
    # Update final
    ws_final.update([df_final.columns.values.tolist()] + df_final.values.tolist())
    
    return True, f"Sucesso! {len(df_final)} linhas consolidadas."

# --- INTERFACE (TABS) ---
# Agora temos 3 abas
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
            df_comissao = pd.DataFrame(dados_comissao, columns=["Data Ref.", "Arquivo", "Técnico", "Horas"])
            st.dataframe(df_comissao)
            if st.button("Gravar Comissões", key="btn_comissao"):
                with st.spinner("Enviando..."):
                    client = conectar_sheets(); aba = client.open_by_key(ID_PLANILHA_MESTRA).worksheet("Comissoes")
                    aba.append_rows(dados_comissao); st.success("✅ Sucesso!")

# --- TAB 2: APROVEITAMENTO ---
with aba_aproveitamento:
    st.header("Extrator de Aproveitamento")
    arquivos_aprov = st.file_uploader("Upload Aproveitamento HTML", type=["html", "htm"], accept_multiple_files=True, key="uploader_aprov")
    
    if arquivos_aprov:
        dados_aprov = []
        amostra_linhas = [] # Debug
        
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
                
                for i, l in enumerate(linhas[:10]):
                    amostra_linhas.append(l.get_text(separator=" ", strip=True))

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
            df_aprov = pd.DataFrame(dados_aprov, columns=["Data", "Arquivo", "Técnico", "T. Disp", "TP", "TG"])
            st.success(f"✅ Sucesso! {len(dados_aprov)} registros.")
            st.dataframe(df_aprov)
            if st.button("Gravar Aproveitamento", key="btn_aprov"):
                with st.spinner("Enviando..."):
                    client = conectar_sheets(); aba = client.open_by_key(ID_PLANILHA_MESTRA).worksheet("Aproveitamento")
                    aba.append_rows(dados_aprov); st.success("✅ Gravado!")
        else:
            st.warning("⚠️ Nenhum dado encontrado. Veja abaixo o que o robô enxergou:")
            with st.expander("🕵️‍♂️ RAIO-X (O que o robô leu no arquivo?)"):
                if amostra_linhas:
                    for l in amostra_linhas:
                        st.text(l)
                else:
                    st.error("O robô não encontrou nenhuma linha de tabela (<tr>). O arquivo pode não ser um HTML padrão.")

# --- TAB 3: RELATÓRIO UNIFICADO (NOVO) ---
with aba_unificacao:
    st.header("🔗 Unificação de Dados (Comissões + Aproveitamento)")
    st.info("Este módulo lê os dados que já estão no Google Sheets, cruza as informações por 'Data' e 'Técnico' e gera uma tabela consolidada.")
    
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
