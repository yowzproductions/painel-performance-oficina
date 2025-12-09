import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Comissões", layout="wide")

st.title("📊 Processador de Comissões em Lote (Multi-Arquivos)")
st.write("Arraste VÁRIOS relatórios de dias diferentes. O sistema organizará tudo automaticamente.")

# --- 2. CONEXÃO SEGURA ---
def conectar_sheets():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 
             'https://www.googleapis.com/auth/drive']
    credentials_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

# --- 3. UPLOAD DO ARQUIVO (Agora aceita múltiplos!) ---
# Mudança chave: accept_multiple_files=True
arquivos = st.file_uploader("Solte seus relatórios HTML aqui", type=["html", "htm"], accept_multiple_files=True)

# Só começa se tiver pelo menos 1 arquivo
if arquivos:
    dados_para_enviar = [] # Lista única para acumular dados de TODOS os arquivos
    
    st.write(f"📂 Iniciando processamento de {len(arquivos)} arquivos...")
    
    # --- LOOP PARA LER CADA ARQUIVO DA LISTA ---
    for arquivo_atual in arquivos:
        try:
            # Lê o arquivo atual
            conteudo = arquivo_atual.read().decode("utf-8", errors='ignore')
            soup = BeautifulSoup(conteudo, "html.parser")
            
            # --- CAPTURA A DATA DESTE ARQUIVO ESPECÍFICO ---
            texto_completo = soup.get_text(separator=" ", strip=True)
            match_data = re.search(r"até\s+(\d{2}/\d{2}/\d{4})", texto_completo, re.IGNORECASE)
            
            if match_data:
                data_relatorio = match_data.group(1)
            else:
                match_generico = re.search(r"(\d{2}/\d{2}/\d{4})", texto_completo)
                if match_generico:
                    data_relatorio = match_generico.group(1)
                else:
                    data_relatorio = datetime.now().strftime("%d/%m/%Y")

            # --- PROCESSAMENTO DOS TÉCNICOS ---
            tecnico_atual = None
            linhas = soup.find_all("tr")
            
            for linha in linhas:
                texto_linha = linha.get_text(separator=" ", strip=True).upper()
                
                # Trava de fim de arquivo
                if "TOTAL DA FILIAL" in texto_linha or "TOTAL DA EMPRESA" in texto_linha:
                    break
                
                # Identifica Técnico
                if "TOTAL DO FUNCIONARIO" in texto_linha:
                    try:
                        parte_nome = texto_linha.split("TOTAL DO FUNCIONARIO")[1]
                        texto_sujo = parte_nome.replace(":", "").strip()
                        tecnico_atual = texto_sujo.split()[0] # Pega só a sigla
                    except:
                        continue 
                        
                # Pega Horas
                if tecnico_atual and "HORAS VENDIDAS:" in texto_linha:
                    celulas = linha.find_all("td")
                    for celula in celulas:
                        texto_celula = celula.get_text(strip=True).upper()
                        if "HORAS" in texto_celula and any(c.isdigit() for c in texto_celula) and "VENDIDAS" not in texto_celula:
                            valor_limpo = texto_celula.replace("HORAS", "").strip()
                            
                            # Adiciona à lista geral
                            # Note que 'arquivo_atual.name' muda a cada loop
                            dados_para_enviar.append([data_relatorio, arquivo_atual.name, tecnico_atual, valor_limpo])
                            break 
                            
        except Exception as e:
            st.error(f"Erro ao ler o arquivo {arquivo_atual.name}: {e}")

    # --- 4. EXIBIÇÃO E ENVIO (Tudo de uma vez) ---
    if len(dados_para_enviar) > 0:
        df = pd.DataFrame(dados_para_enviar, columns=["Data Ref.", "Arquivo Original", "Técnico", "Horas"])
        st.success(f"✅ Processamento concluído! Total de {len(dados_para_enviar)} registros extraídos de {len(arquivos)} arquivos.")
        st.dataframe(df)
        
        if st.button("Confirmar e Gravar TUDO no Sheets"):
            with st.spinner("Enviando lote gigante para o Google..."):
                try:
                    client = conectar_sheets()
                    ID_PLANILHA = "1XibBlm2x46Dk5bf4JvfrMepD4gITdaOtTALSgaFcwV0" # Seu ID
                    arquivo_sheet = client.open_by_key(ID_PLANILHA)
                    
                    try:
                        aba = arquivo_sheet.worksheet("Comissoes")
                    except:
                        st.error("❌ Erro: Aba 'Comissoes' não encontrada.")
                        st.stop()
                    
                    aba.append_rows(dados_para_enviar)
                    
                    st.balloons()
                    st.success(f"✅ Sucesso Absoluto! {len(dados_para_enviar)} linhas gravadas.")
                    
                except Exception as e:
                    if "200" in str(e):
                        st.balloons()
                        st.success("✅ Sucesso confirmado (Protocolo 200).")
                    else:
                        st.error(f"Erro no envio: {e}")
    else:
        st.warning("Nenhum dado válido encontrado nos arquivos enviados.")
