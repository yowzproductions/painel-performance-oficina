import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Comissões", layout="wide")

st.title("📊 Processador de Comissões em Lote")
st.write("Identifica cada técnico e suas respectivas horas vendidas automaticamente.")

# --- 2. CONEXÃO SEGURA ---
def conectar_sheets():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 
             'https://www.googleapis.com/auth/drive']
    credentials_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

# --- 3. UPLOAD DO ARQUIVO ---
arquivo = st.file_uploader("Solte o relatório HTML aqui", type=["html", "htm"])

if arquivo:
    # Lê o arquivo
    conteudo = arquivo.read().decode("utf-8", errors='ignore')
    soup = BeautifulSoup(conteudo, "html.parser")
    
    # Lista para guardar todos os dados encontrados
    dados_para_enviar = []
    
    # Memória do técnico atual
    tecnico_atual = None
    
    # Pega todas as linhas da tabela
    linhas = soup.find_all("tr")
    
    st.write(f"🔍 Analisando {len(linhas)} linhas do arquivo...")
    
    for linha in linhas:
        texto_linha = linha.get_text(separator=" ", strip=True).upper()
        
        # Acha o técnico
        if "TOTAL DO FUNCIONARIO" in texto_linha:
            try:
                parte_nome = texto_linha.split("TOTAL DO FUNCIONARIO")[1]
                tecnico_atual = parte_nome.replace(":", "").strip()
            except:
                continue 
                
        # Se tem técnico, busca horas
        if tecnico_atual and "HORAS VENDIDAS:" in texto_linha:
            celulas = linha.find_all("td")
            
            for celula in celulas:
                texto_celula = celula.get_text(strip=True).upper()
                
                if "HORAS" in texto_celula and any(c.isdigit() for c in texto_celula) and "VENDIDAS" not in texto_celula:
                    valor_limpo = texto_celula.replace("HORAS", "").strip()
                    
                    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    # Adiciona à lista
                    dados_para_enviar.append([timestamp, arquivo.name, tecnico_atual, valor_limpo])
                    break 

    # --- 4. EXIBIÇÃO E ENVIO (Alinhado corretamente dentro do if arquivo) ---
    if len(dados_para_enviar) > 0:
        df = pd.DataFrame(dados_para_enviar, columns=["Data", "Arquivo", "Técnico", "Horas"])
        st.success(f"Encontrei {len(dados_para_enviar)} registros!")
        st.dataframe(df)
        
        if st.button("Confirmar e Gravar na Aba 'Comissoes'"):
            with st.spinner("Enviando dados..."):
                try:
                    client = conectar_sheets()
                    
                    # Abre o arquivo "Dados_HTML"
                    sheet_file = client.open("Dados_HTML")
                    
                    # Tenta acessar a aba específica "Comissoes"
                    try:
                        aba = sheet_file.worksheet("Comissoes")
                    except:
                        st.error("❌ Não encontrei a aba 'Comissoes'. Verifique o nome na planilha.")
                        st.stop() # Para o código aqui se não achar a aba
                    
                    # Envia os dados
                    aba.append_rows(dados_para_enviar)
                    
                    st.balloons()
                    st.success("✅ Sucesso! Dados gravados na aba 'Comissoes'.")
                    
                except Exception as e:
                    # Tratamento para o erro falso-positivo "200"
                    if "200" in str(e):
                        st.balloons()
                        st.success("✅ Sucesso confirmado pelo Google (Código 200).")
                    else:
                        st.error(f"Erro técnico: {e}")
    else:
        st.warning("Nenhum dado encontrado. Verifique se o HTML contém 'TOTAL DO FUNCIONARIO' e 'HORAS VENDIDAS'.")
