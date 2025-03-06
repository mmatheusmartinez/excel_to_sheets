import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Nome do arquivo de credenciais JSON baixado do Google Cloud
FILENAME = "xxx.json"

# Caminho do arquivo Excel
EXCEL_FILE = "C:/Users/caminho/testeintegracao.xlsx"
SHEET_NAME = "Planilha1"  # Nome da aba dentro do Excel

# Escopos necessários para acessar o Google Sheets
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# Autenticação com as credenciais
creds = ServiceAccountCredentials.from_json_keyfile_name(
    filename=FILENAME,
    scopes=SCOPES
)

client = gspread.authorize(creds)

# Acessar a planilha pelo título
planilha_completa = client.open("Integração")

# Pegar a primeira aba da planilha
planilha = planilha_completa.get_worksheet(0)

# Obter os dados do Google Sheets e transformar em DataFrame
dados = planilha.get_all_records()  # Obtém os dados existentes
df_google = pd.DataFrame(dados)

# ✅ **Ler os dados do Excel**
df_excel = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

# ✅ **Juntar os dados (Anexando novas observações)**
df_atualizado = pd.concat([df_google, df_excel], ignore_index=True)

# ✅ **Corrigir valores NaN antes de enviar para o Google Sheets**
df_atualizado = df_atualizado.fillna("")  # Substitui NaN por string vazia

# ✅ **Converter o DataFrame atualizado em lista de listas para enviar ao Google Sheets**
dados_atualizados = [df_atualizado.columns.tolist()] + df_atualizado.values.tolist()

# ✅ **Atualizar o Google Sheets**
planilha.clear()  # Limpa a planilha antes de enviar os novos dados
planilha.update(dados_atualizados)  # Envia os novos dados

print("✅ Dados do Excel adicionados ao Google Sheets com sucesso!")