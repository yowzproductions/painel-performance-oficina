import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import unicodedata

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Central de Relatórios WLM", layout="wide", page_icon="🔒")
ID_PLANILHA_MESTRA = "1XibBlm2x46Dk5bf4JvfrMepD4gITdaOtTALSgaFcwV0"

# --- AUXILIARES ---
def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def conectar_sheets():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

def converter_br_para_float(valor):
    """
    Limpa o valor para garantir que seja processável como número.
    Nota: A divisão por 100 ocorrerá APENAS na exportação final.
    """
    if pd.isna(valor) or valor == "": 
        return 0.0
    
    if isinstance(valor, (int, float)): 
        return float(valor)
    
    valor_str = str(valor).strip()
    valor_str = valor_str.replace('\xa0', '').replace('R$', '').strip()

    if not valor_str:
        return 0.0

    # Remove ponto de milhar se existir
    if '.' in valor_str and ',' in valor_str: 
        valor_str = valor_str.replace('.', '')
    
    # Troca vírgula por ponto para o Python entender
    valor_str = valor_str.replace(',', '.')

    try: 
        return float(valor_str)
    except: 
        return 0.0

def padronizar_data_quatro_digitos(data_str):
    """
    Transforma '08/12/25' em '08/12/2025'.
    Garante que as chaves de data sejam idênticas para o merge.
    """
    if pd.isna(data_str) or data_str == "":
        return ""
    
    data_str = str(data_str).strip()
    
    # Verifica se tem barras
    if '/' in data_str:
        partes = data_str.split('/')
        # Se tiver 3 partes (dia, mes, ano)
        if len(partes) == 3:
            dia, mes, ano = partes
            # Se o ano tiver apenas 2 dígitos, adiciona '20' na frente
            if len(ano) == 2:
                ano = '20' + ano
            
            # Reconstrói a data padronizada com zeros à esquerda se precisar
            return f"{dia.zfill(2)}/{mes.zfill(2)}/{ano}"
            
    return data_str

def verificar_acesso():
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        try: return sh.worksheet("Config").acell('B1').value
        except: return 'admin'
    except: return None

# --- PARSERS (LEITURA) ---
def parse_comissoes(arquivos):
    dados = []
    for arquivo in arquivos:
        try:
            arquivo.seek(0)
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
                            valor_limpo = txt.replace("HORAS", "").strip()
                            dados.append([data_relatorio, arquivo.name, tecnico_atual, valor_limpo])
                            break 
        except Exception as e: st.error(f"Erro no arquivo {arquivo.name}: {e}")
    return dados

def parse_aproveitamento(arquivos):
    dados = []
    for arquivo in arquivos:
        try:
            arquivo.seek(0)
            try: conteudo = arquivo.read().decode("utf-8")
            except:
                try: conteudo = arquivo.read().decode("latin-1")
                except: conteudo = arquivo.read().decode("utf-16")
            
            soup = BeautifulSoup(conteudo, "html.parser")
            tecnico_atual_aprov = None
            linhas = soup.find_all("tr")
            for linha in linhas:
                texto_original = linha.get_text(separator=" ", strip=True).upper()
                texto_limpo = remover_acentos(texto_original)
                
                if "TOTAL FILIAL:" in texto_original: break
                if "MECANICO" in texto_limpo and "TOT.MEC" not in texto_limpo:
                    try:
                        parte_direita = texto_limpo.split("MECANICO")[1].replace(":", "").strip()
                        if "-" in parte_direita: tecnico_atual_aprov = parte_direita.split("-")[0].strip()
                        else: tecnico_atual_aprov = parte_direita.split()[0]
                    except: continue
                
                if "TOT.MEC.:" in texto_original: tecnico_atual_aprov = None; continue
                
                if tecnico_atual_aprov:
                    celulas = linha.find_all("td")
                    if not celulas: continue
                    txt_cel0 = celulas[0].get_text(strip=True)
                    if re.match(r"\d{2}/\d{2}/\d{2}", txt_cel0):
                        try:
                            if len(celulas) >= 4:
                                dados.append([
                                    txt_cel0.split()[0], 
                                    arquivo.name, 
                                    tecnico_atual_aprov, 
                                    celulas[1].get_text(strip=True), 
                                    celulas[2].get_text(strip=True), 
                                    celulas[3].get_text(strip=True)
                                ])
                        except: continue
        except Exception as e: st.error(f"Erro no arquivo {arquivo.name}: {e}")
    return dados

# --- GRAVAÇÃO ---
def atualizar_planilha_preservando_formato(sh, nome_aba, df_final):
    try:
        ws = sh.worksheet(nome_aba)
    except:
        ws = sh.add_worksheet(title=nome_aba, rows=2000, cols=20)

    if not ws.get_all_values():
        ws.update('A1', [df_final.columns.values.tolist()])
        try: ws.format('A1:Z1', {'textFormat': {'bold': True}})
        except: pass

    ws.batch_clear(["A2:Z10000"])
    
    # Preenche vazios com 0.0
    df_final = df_final.fillna(0.0)
    
    dados_para_enviar = df_final.values.tolist()
    if dados_para_enviar:
        ws.update('A2', dados_para_enviar)
        
    return True

# --- UPSERT ---
def salvar_com_upsert(nome_aba, novos_dados_df, colunas_chaves):
    client = conectar_sheets()
    sh = client.open_by_key(ID_PLANILHA_MESTRA)
    
    try:
        ws = sh.worksheet(nome_aba)
        dados_antigos = ws.get_all_records()
        df_antigo = pd.DataFrame(dados_antigos)
    except:
        df_antigo = pd.DataFrame()

    if not df_antigo.empty:
        for col in df_antigo.columns: df_antigo[col] = df_antigo[col].astype(str)
    for col in novos_dados_df.columns: novos_dados_df[col] = novos_dados_df[col].astype(str)

    df_total = pd.concat([df_antigo, novos_dados_df])
    df_final = df_total.drop_duplicates(subset=colunas_chaves, keep='last')
    
    atualizar_planilha_preservando_formato(sh, nome_aba, df_final)
    return len(df_final)

# --- NOVA FUNÇÃO: SALVAR AJUSTE MANUAL ---
def salvar_ajuste_manual(data, tecnico, metrica, valor, motivo):
    client = conectar_sheets()
    sh = client.open_by_key(ID_PLANILHA_MESTRA)
    try:
        ws = sh.worksheet("Ajustes")
    except:
        ws = sh.add_worksheet(title="Ajustes", rows=1000, cols=10)
        ws.append_row(["Data", "Técnico", "Métrica", "Valor", "Motivo", "Data do Registro"])
    
    # Salva o novo ajuste
    ws.append_row([
        str(data.strftime('%d/%m/%Y')), 
        tecnico, 
        metrica, 
        float(valor), 
        motivo, 
        datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    ])

# --- NOVA FUNÇÃO: APLICAR LÓGICA DE AJUSTES ---
def aplicar_logica_ajustes(df_base):
    """
    Lê a aba 'Ajustes' e aplica matematicamente ao DataFrame antes de salvar no Consolidado.
    """
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        ws_ajustes = sh.worksheet("Ajustes")
        dados_ajustes = ws_ajustes.get_all_records()
        
        if not dados_ajustes:
            return df_base

        df_ajustes = pd.DataFrame(dados_ajustes)
        
        # Mapeamento de métricas (Nome no Dropdown -> Nome na Coluna do DF)
        mapa = {
            "Horas Vendidas (HV)": "Horas Vendidas",
            "Tempo Padrão (TP)": "TP",
            "Tempo Disponível (Disp)": "Disp",
            "Tempo Garantia (TG)": "TG"
        }

        # Garante datas comparáveis
        df_base['Key_D_Comp'] = pd.to_datetime(df_base['Data'], dayfirst=True, errors='coerce')
        
        for _, row in df_ajustes.iterrows():
            try:
                # Dados do Ajuste
                dt_ajuste = pd.to_datetime(row['Data'], dayfirst=True, errors='coerce')
                tec_ajuste = str(row['Técnico']).strip()
                metrica_ajuste = mapa.get(row['Métrica'])
                valor_ajuste = float(str(row['Valor']).replace(',', '.'))

                if metrica_ajuste and metrica_ajuste in df_base.columns:
                    # Filtra a linha correta no DataFrame Base
                    mask = (df_base['Key_D_Comp'] == dt_ajuste) & (df_base['Técnico'] == tec_ajuste)
                    
                    if mask.any():
                        # Aplica a soma/subtração
                        df_base.loc[mask, metrica_ajuste] += valor_ajuste
            except Exception as e:
                print(f"Erro ao processar linha de ajuste: {e}")
                continue
        
        # Remove coluna auxiliar
        if 'Key_D_Comp' in df_base.columns:
            df_base.drop(columns=['Key_D_Comp'], inplace=True)
            
        return df_base

    except Exception as e:
        print(f"Erro geral ao ler ajustes: {e}")
        return df_base # Retorna o original se der erro nos ajustes

# --- UNIFICAÇÃO (ATUALIZADA COM AJUSTES) ---
def processar_unificacao():
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        ws_com = sh.worksheet("Comissoes")
        ws_aprov = sh.worksheet("Aproveitamento")

        dados_com = ws_com.get_all_records()
        dados_aprov = ws_aprov.get_all_records()

        if not dados_com or not dados_aprov: return False

        df_com = pd.DataFrame(dados_com)
        df_aprov = pd.DataFrame(dados_aprov)

        # Limpeza e Padronização
        df_com.columns = [c.strip() for c in df_com.columns]
        df_aprov.columns = [c.strip() for c in df_aprov.columns]
        renomear_comissao = {"Data Processamento": "Data", "Sigla Técnico": "Técnico"}
        df_com.rename(columns=renomear_comissao, inplace=True)

        cols_com = ['Data', 'Técnico', 'Horas Vendidas']
        df_com = df_com[[c for c in cols_com if c in df_com.columns]]
        cols_aprov = ['Data', 'Técnico', 'Disp', 'TP', 'TG']
        df_aprov = df_aprov[[c for c in cols_aprov if c in df_aprov.columns]]

        # Padronização de Data
        if 'Data' in df_com.columns:
            df_com['Data'] = df_com['Data'].apply(padronizar_data_quatro_digitos)
        
        if 'Data' in df_aprov.columns:
            df_aprov['Data'] = df_aprov['Data'].apply(padronizar_data_quatro_digitos)

        # Conversão Numérica Inicial
        cols_numericas = ['Horas Vendidas', 'Disp', 'TP', 'TG']
        for col in cols_numericas:
            if col in df_com.columns: df_com[col] = df_com[col].apply(converter_br_para_float)
            if col in df_aprov.columns: df_aprov[col] = df_aprov[col].apply(converter_br_para_float)

        # Merge
        df_com['Key_D'] = df_com['Data'].astype(str)
        df_com['Key_T'] = df_com['Técnico'].astype(str)
        df_aprov['Key_D'] = df_aprov['Data'].astype(str)
        df_aprov['Key_T'] = df_aprov['Técnico'].astype(str)

        df_final = pd.merge(
            df_com, df_aprov, 
            left_on=['Key_D', 'Key_T'],
