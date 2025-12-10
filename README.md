# 🏭 Central de Processamento de Relatórios (ETL)

> **Status:** ✅ Em Produção  
> **Tecnologia:** Python + Streamlit + Google Sheets API

Este projeto é uma ferramenta de automação **ETL (Extract, Transform, Load)** desenvolvida para processar relatórios operacionais em HTML, extrair indicadores chaves (KPIs) e alimentar automaticamente uma planilha de gestão na nuvem.

O objetivo é eliminar a digitação manual, reduzir erros humanos e acelerar o fechamento de comissões e análise de produtividade técnica.

---

## 🚀 Funcionalidades

O sistema opera com uma interface web amigável (Drag & Drop) dividida em dois módulos:

### 1. 💰 Módulo de Comissões
* **Entrada:** Relatórios de Pagamento de Comissões (HTML).
* **Processamento:**
    * Lê múltiplos arquivos simultaneamente.
    * Identifica a **Data de Competência** real do relatório (ignora data de upload).
    * Isola a **Sigla do Técnico** (ex: "AAD").
    * Extrai as **Horas Vendidas**.
    * Ignora totais gerais (Filial/Empresa) para evitar sujeira nos dados.
* **Saída:** Grava na aba `Comissoes` do Google Sheets.

### 2. ⚙️ Módulo de Aproveitamento Técnico
* **Entrada:** Relatórios de Aproveitamento de Tempo Mecânico (HTML/SLK).
* **Processamento:**
    * Suporta codificações antigas (Latin-1) e modernas (UTF-8).
    * Limpa nomes complexos de técnicos e datas com dias da semana.
    * Extrai indicadores: **T. Disp** (Tempo Disponível), **TP** (Tempo Padrão) e **TG** (Tempo Gasto).
* **Saída:** Grava na aba `Aproveitamento` do Google Sheets.

---

## 🛠️ Arquitetura e Tecnologias

* **Frontend:** [Streamlit](https://streamlit.io/) (Interface Web Interativa).
* **Backend:** Python 3.9+.
* **Processamento de Dados:**
    * `BeautifulSoup4`: Para raspagem (scraping) e leitura dos arquivos HTML.
    * `Pandas`: Para estruturação e manipulação tabular dos dados.
    * `Regex`: Para captura inteligente de padrões de texto (datas e siglas).
* **Banco de Dados:** Google Sheets (via API `gspread`).

---

## 📋 Pré-requisitos de Configuração

Para rodar este projeto, é necessário configurar o acesso ao Google Cloud Platform (GCP).

### 1. Planilha Google
Crie uma planilha e garanta que ela tenha as seguintes abas e cabeçalhos na **Linha 1**:

* **Aba `Comissoes`:**
    `Data Ref. | Arquivo | Técnico | Horas`
* **Aba `Aproveitamento`:**
    `Data | Arquivo | Técnico | T. Disp | TP | TG`

### 2. Credenciais (Google Service Account)
1.  Crie um projeto no Google Cloud Console.
2.  Ative as APIs: **Google Sheets API** e **Google Drive API**.
3.  Crie uma Service Account e baixe a chave JSON.
4.  **Importante:** Compartilhe a sua planilha (botão Share) com o e-mail da Service Account (ex: `bot-sheets@...iam.gserviceaccount.com`) como **Editor**.

---

## ☁️ Como Rodar no Streamlit Cloud

Este projeto foi desenhado para rodar na nuvem sem instalação local.

1.  Faça o Fork/Clone deste repositório.
2.  Acesse [share.streamlit.io](https://share.streamlit.io/).
3.  Crie um novo app apontando para este repositório.
4.  Nas configurações do App (**Settings > Secrets**), adicione suas credenciais no formato TOML:

```toml
[gcp_service_account]
type = "service_account"
project_id = "seu-project-id"
private_key_id = "sua-key-id"
private_key = "-----BEGIN PRIVATE KEY-----\n..."
client_email = "seu-bot@..."
client_id = "..."
auth_uri = "[https://accounts.google.com/o/oauth2/auth](https://accounts.google.com/o/oauth2/auth)"
token_uri = "[https://oauth2.googleapis.com/token](https://oauth2.googleapis.com/token)"
auth_provider_x509_cert_url = "..."
client_x509_cert_url = "..."
