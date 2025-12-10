import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Central de Relatórios WLM", layout="wide")
st.title("🏭 Central de Processamento de Relatórios")

# --- 2. CONEXÃO SEGURA ---
def conectar_sheets():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 
             'https://www.googleapis.com/auth/drive']
    credentials_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

# ID DA SUA PLANILHA
ID_PLANILHA_MESTRA = "1XibBlm2x46Dk5bf4JvfrMepD4gITdaOtTALSgaFcwV0"

# --- ABAS ---
aba_comissoes, aba_aproveitamento = st.tabs(["💰 Pagamento de Comissões", "⚙️ Aproveitamento Técnico"])

# ==============================================================================
# SISTEMA 1: PAGAMENTO DE COMISSÕES (MANTIDO IGUAL)
# ==============================================================================
with aba_comissoes:
    st.header("Processador de Comissões")
    st.write("Arraste os relatórios de 'Pagamento de Comissões' (HTML).")
    
    arquivos_comissao = st.file_uploader("Upload Comissões HTML", type=["html", "htm"], accept_multiple_files=True, key="uploader_comissao")

    if arquivos_comissao:
        dados_comissao = []
        st.write(f"📂 Processando {len(arquivos_comissao)} arquivos...")
        
        for arquivo in arquivos_comissao:
            try:
                conteudo = arquivo.read().decode("utf-8", errors='ignore')
                soup = BeautifulSoup(conteudo, "html.parser")
                
                texto_completo = soup.get_text(separator=" ", strip=True)
                match_data = re.search(r"até\s+(\d{2}/\d{2}/\d{4})", texto_completo, re.IGNORECASE)
                data_relatorio = match_data.group(1) if match_data else datetime.now().strftime("%d/%m/%Y")

                tecnico_atual = None
                linhas = soup.find_all("tr")
                
                for linha in linhas:
                    texto_linha = linha.get_text(separator=" ", strip=True).upper()
                    
                    if "TOTAL DA FILIAL" in texto_linha or "TOTAL DA EMPRESA" in texto_linha:
                        break
                    
                    if "TOTAL DO FUNCIONARIO" in texto_linha:
                        try:
                            # Limpeza da Sigla (Comissões)
                            tecnico_atual = texto_linha.split("TOTAL DO FUNCIONARIO")[1].replace(":", "").strip().split()[0]
                        except:
                            continue 
                            
                    if tecnico_atual and "HORAS VENDIDAS:" in texto_linha:
                        celulas = linha.find_all("td")
                        for celula in celulas:
                            texto_celula = celula.get_text(strip=True).upper()
                            if "HORAS" in texto_celula and any(c.isdigit() for c in texto_celula) and "VENDIDAS" not in texto_celula:
                                valor_limpo = texto_celula.replace("HORAS", "").strip()
                                dados_comissao.append([data_relatorio, arquivo.name, tecnico_atual, valor_limpo])
                                break 
            except Exception as e:
                st.error(f"Erro no arquivo {arquivo.name}: {e}")

        if len(dados_comissao) > 0:
            df_comissao = pd.DataFrame(dados_comissao, columns=["Data Ref.", "Arquivo", "Técnico", "Horas"])
            st.dataframe(df_comissao)
            
            if st.button("Gravar Comissões no Sheets", key="btn_comissao"):
                with st.spinner("Enviando..."):
                    try:
                        client = conectar_sheets()
                        sheet = client.open_by_key(ID_PLANILHA_MESTRA)
                        aba = sheet.worksheet("Comissoes")
                        aba.append_rows(dados_comissao)
                        st.success(f"✅ Sucesso! {len(dados_comissao)} linhas gravadas.")
                    except Exception as e:
                        if "200" in str(e): st.success("✅ Sucesso (200).")
                        else: st.error(f"Erro: {e}")

# ==============================================================================
# SISTEMA 2: APROVEITAMENTO TÉCNICO (DATA E SIGLA LIMPAS)
# ==============================================================================
with aba_aproveitamento:
    st.header("Extrator de Aproveitamento (T.Disp / TP / TG)")
    st.write("Arraste os relatórios de 'Aproveitamento Tempo Mecânico' (HTML).")
    
    arquivos_aprov = st.file_uploader("Upload Aproveitamento HTML", type=["html", "htm"], accept_multiple_files=True, key="uploader_aprov")
    
    if arquivos_aprov:
        dados_aprov = []
        st.write(f"📂 Processando {len(arquivos_aprov)} arquivos...")
        
        for arquivo in arquivos_aprov:
            try:
                conteudo = arquivo.read().decode("utf-8", errors='ignore')
                soup = BeautifulSoup(conteudo, "html.parser")
                tecnico_atual_aprov = None
                linhas = soup.find_all("tr")
                
                for linha in linhas:
                    texto_linha = linha.get_text(separator=" ", strip=True).upper()
                    
                    if "TOTAL FILIAL:" in texto_linha:
                        break

                    # 1. Identifica e LIMPA o Técnico
                    if "MECÂNICO:" in texto_linha or "MECANICO:" in texto_linha:
                        try:
                            # Pega o que vem depois de MECANICO:
                            parte_direita = texto_linha.split("MECANICO:")[1] if "MECANICO:" in texto_linha else texto_linha.split("MECÂNICO:")[1]
                            
                            # Lógica de Limpeza Pesada:
                            # Se tiver traço ("AAD - ALLAN"), pega só o que vem antes do traço
                            if "-" in parte_direita:
                                tecnico_limpo = parte_direita.split("-")[0].strip()
                            else:
                                # Se não tiver traço ("AAD ALLAN"), pega só a primeira palavra
                                tecnico_limpo = parte_direita.strip().split()[0]
                            
                            tecnico_atual_aprov = tecnico_limpo
                        except:
                            continue

                    if "TOT.MEC.:" in texto_linha:
                        tecnico_atual_aprov = None
                        continue

                    # 2. Identifica e LIMPA a Data
                    if tecnico_atual_aprov:
                        celulas = linha.find_all("td")
                        if not celulas: continue
                        
                        texto_primeira_celula = celulas[0].get_text(strip=True)
                        
                        # Verifica se começa com formato de data DD/MM/YY
                        if re.match(r"\d{2}/\d{2}/\d{2}", texto_primeira_celula):
                            try:
                                # LIMPEZA DA DATA:
                                # Pega "01/12/25 SEG", divide por espaço e pega só o índice [0]
                                data_limpa = texto_primeira_celula.split()[0] 
                                
                                t_disp = celulas[1].get_text(strip=True)
                                tp = celulas[2].get_text(strip=True)
                                tg = celulas[3].get_text(strip=True)
                                
                                # Adiciona os dados já limpos
                                dados_aprov.append([data_limpa, arquivo.name, tecnico_atual_aprov, t_disp, tp, tg])
                            except IndexError:
                                continue

            except Exception as e:
                st.error(f"Erro ao ler arquivo {arquivo.name}: {e}")
                
        if len(dados_aprov) > 0:
            df_aprov = pd.DataFrame(dados_aprov, columns=["Data", "Arquivo", "Técnico", "T. Disp", "TP", "TG"])
            st.success(f"Encontrei {len(dados_aprov)} registros limpos!")
            st.dataframe(df_aprov)
            
            if st.button("Gravar Aproveitamento no Sheets", key="btn_aprov"):
                with st.spinner("Enviando..."):
                    try:
                        client = conectar_sheets()
                        sheet = client.open_by_key(ID_PLANILHA_MESTRA)
                        
                        try:
                            aba = sheet.worksheet("Aproveitamento")
                        except:
                            st.error("❌ Erro: Crie a aba 'Aproveitamento'!")
                            st.stop()
                            
                        aba.append_rows(dados_aprov)
                        st.success(f"✅ Sucesso! Dados gravados na aba 'Aproveitamento'.")
                    except Exception as e:
                        if "200" in str(e): st.success("✅ Sucesso (200).")
                        else: st.error(f"Erro: {e}")
