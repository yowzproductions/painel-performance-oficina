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

# --- O MOTOR DE UNIFICAÇÃO (AUTOMÁTICO E LIMPO) ---
def processar_unificacao():
    """
    Lê as abas, remove colunas inúteis (arquivos), cruza os dados e atualiza.
    """
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)

        # 1. Ler as abas de origem
        try:
            ws_com = sh.worksheet("Comissoes")
            ws_aprov = sh.worksheet("Aproveitamento")
        except:
            return False

        # 2. Ler os dados
        dados_com = ws_com.get_all_records()
        dados_aprov = ws_aprov.get_all_records()

        if not dados_com or not dados_aprov:
            return False

        df_com = pd.DataFrame(dados_com)
        df_aprov = pd.DataFrame(dados_aprov)

        # 3. Limpeza de Colunas (strip)
        df_com.columns = [c.strip() for c in df_com.columns]
        df_aprov.columns = [c.strip() for c in df_aprov.columns]

        # 4. Ajuste de Nomes (Padronização)
        renomear_comissao = {"Data Processamento": "Data", "Sigla Técnico": "Técnico"}
        df_com.rename(columns=renomear_comissao, inplace=True)

        # Validação básica
        if "Data" not in df_com.columns or "Técnico" not in df_com.columns:
            return False
        if "Data" not in df_aprov.columns or "Técnico" not in df_aprov.columns:
            return False

        # --- LIMPEZA DE DADOS (NOVA ETAPA) ---
        # Aqui selecionamos APENAS as colunas que interessam para o relatório final
        # Jogamos fora "Nome do Arquivo" e "Arquivo"
        
        colunas_uteis_comissao = ['Data', 'Técnico', 'Horas Vendidas']
        # Verifica se as colunas existem antes de filtrar para não dar erro
        df_com = df_com[[c for c in colunas_uteis_comissao if c in df_com.columns]]

        colunas_uteis_aprov = ['Data', 'Técnico', 'Disp', 'TP', 'TG']
        df_aprov = df_aprov[[c for c in colunas_uteis_aprov if c in df_aprov.columns]]

        # 5. Padronização de Tipos
        df_com['Data'] = df_com['Data'].astype(str)
        df_com['Técnico'] = df_com['Técnico'].astype(str)
        df_aprov['Data'] = df_aprov['Data'].astype(str)
        df_aprov['Técnico'] = df_aprov['Técnico'].astype(str)

        # 6. Merge (Cruzamento Limpo)
        df_final = pd.merge(
            df_com, 
            df_aprov, 
            on=['Data', 'Técnico'], 
            how='outer', 
            suffixes=('_Com', '_Aprov')
        )
        df_final.fillna("", inplace=True)

        # 7. Salvar
        try:
            ws_final = sh.worksheet("Consolidado")
            ws_final.clear()
        except:
            ws_final = sh.add_worksheet(title="Consolidado", rows=1000, cols=20)
        
        ws_final.update([df_final.columns.values.tolist()] + df_final.values.tolist())
        return True

    except Exception as e:
        print(f"Erro silencioso na unificação: {e}")
        return False

# --- INTERFACE ---
aba_comissoes, aba_aproveitamento = st.tabs([
    "💰 Pagamento de Comissões", 
    "⚙️ Aproveitamento Técnico"
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
            colunas_comissao = ["Data Processamento", "Nome do Arquivo", "Sigla Técnico", "Horas Vendidas"]
            df_comissao = pd.DataFrame(dados_comissao, columns=colunas_comissao)
            st.dataframe(df_comissao)
            
            if st.button("💾 Gravar Comissões e Atualizar Base", key="btn_comissao"):
                progresso = st.progress(0, text="Iniciando gravação...")
                try:
                    progresso.progress(30, text="Enviando dados para a nuvem...")
                    client = conectar_sheets()
                    aba = client.open_by_key(ID_PLANILHA_MESTRA).worksheet("Comissoes")
                    if not aba.get_all_values():
                        aba.append_row(colunas_comissao)
                    aba.append_rows(dados_comissao)
                    
                    progresso.progress(70, text="Recalculando Relatório Unificado...")
                    processar_unificacao()
                    
                    progresso.progress(100, text="Concluído!")
                    st.success("✅ Sucesso! Dados gravados e Relatório Limpo gerado.")
                    st.balloons()
                except Exception as e:
                    st.error(f"Erro ao gravar: {e}")

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
            colunas_aprov = ["Data", "Arquivo", "Técnico", "Disp", "TP", "TG"]
            df_aprov = pd.DataFrame(dados_aprov, columns=colunas_aprov)
            
            st.success(f"✅ Sucesso! {len(dados_aprov)} registros.")
            st.dataframe(df_aprov)
            
            if st.button("💾 Gravar Aproveitamento e Atualizar Base", key="btn_aprov"):
                progresso = st.progress(0, text="Iniciando gravação...")
                try:
                    progresso.progress(30, text="Enviando dados para a nuvem...")
                    client = conectar_sheets()
                    aba = client.open_by_key(ID_PLANILHA_MESTRA).worksheet("Aproveitamento")
                    if not aba.get_all_values():
                        aba.append_row(colunas_aprov)
                    aba.append_rows(dados_aprov)
                    
                    progresso.progress(70, text="Recalculando Relatório Unificado...")
                    processar_unificacao()
                    
                    progresso.progress(100, text="Concluído!")
                    st.success("✅ Sucesso! Dados gravados e Relatório Limpo gerado.")
                    st.balloons()
                except Exception as e:
                    st.error(f"Erro ao gravar: {e}")
