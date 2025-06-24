import pandas as pd
import os
import pickle
import re
import unicodedata
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Configurações
CSV_FILE = 'dados.csv'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1nSw2PbhMJOzu2b8R6Zzg0LL2-0UXRRjLYTc73i966QU'
CLEAN_RANGE = 'Limpo!A1'
ANALYSIS_RANGE = 'Analysis!A1'

# Função para autenticar com Google Sheets
def authenticate_google():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('sheets', 'v4', credentials=creds)

def limpar_texto(texto):
    if not isinstance(texto, str):
        return texto
    texto = texto.strip()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')  # remove acentos
    texto = re.sub(r'[^\w\s@._-]', '', texto)  # remove caracteres especiais indesejados
    return texto

def corrigir_nome(nome):
    # Remove espaços internos para juntar nomes quebrados (ex: "Mat eus" → "Mateus")
    if not isinstance(nome, str):
        return nome
    nome_corrigido = nome.replace(' ', '')
    return nome_corrigido

def validar_nome(nome):
    if not isinstance(nome, str) or len(nome) < 3:
        return False
    return True

def corrigir_email(email):
    if not isinstance(email, str):
        return email
    # Remove espaços internos, incluindo espaços antes e depois do '.'
    email_corrigido = email.replace(' ', '').replace('..', '.')
    return email_corrigido

def validar_email(email):
    if not isinstance(email, str):
        return False
    padrao_email = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
    return bool(padrao_email.match(email))

def corrigir_telefone(telefone):
    if not isinstance(telefone, str):
        return telefone
    # Remove tudo que não for número
    telefone_corrigido = re.sub(r'\D', '', telefone)
    return telefone_corrigido

def validar_telefone(telefone):
    if not isinstance(telefone, str):
        return False
    if not telefone.isdigit():
        return False
    return 10 <= len(telefone) <= 13

def main():
    print("🔄 Lendo dados do CSV...")
    df = pd.read_csv(CSV_FILE, encoding='utf-8-sig')

    print("🧹 Limpando dados básicos (acentos e caracteres especiais)...")
    df = df.astype(str).applymap(limpar_texto)  # Pode trocar por map em versões novas

    # Corrigir e validar NOME
    if 'Nome' in df.columns:
        df['Nome_corrigido'] = df['Nome'].map(corrigir_nome)
        df['Nome_valido'] = df['Nome_corrigido'].map(validar_nome)
        df['Nome_sugestao'] = df.apply(
            lambda row: f"Nome muito curto ({row['Nome_corrigido']})" if not row['Nome_valido'] else '',
            axis=1
        )
    else:
        print("⚠️ Coluna 'Nome' não encontrada")

    # Corrigir e validar E-MAIL
    if 'E-mail' in df.columns:
        df['E-mail_corrigido'] = df['E-mail'].map(corrigir_email)
        df['Email_valido'] = df['E-mail_corrigido'].map(validar_email)
        df['Email_sugestao'] = df.apply(
            lambda row: "Formato inválido" if not row['Email_valido'] else '',
            axis=1
        )
    else:
        print("⚠️ Coluna 'E-mail' não encontrada")

    # Corrigir e validar TELEFONE
    if 'Telefone' in df.columns:
        df['Telefone_corrigido'] = df['Telefone'].map(corrigir_telefone)
        df['Telefone_valido'] = df['Telefone_corrigido'].map(validar_telefone)
        df['Telefone_sugestao'] = df.apply(
            lambda row: "Formato inválido" if not row['Telefone_valido'] else '',
            axis=1
        )
    else:
        print("⚠️ Coluna 'Telefone' não encontrada")

    print("📊 Gerando estatísticas descritivas...")
    desc = df.describe(include='all')
    desc_reset = desc.reset_index()

    print("📈 Gerando KPI de contagem por Nome corrigido...")
    if 'Nome_corrigido' in df.columns:
        kpi = df['Nome_corrigido'].value_counts().reset_index()
        kpi.columns = ['Nome_corrigido', 'Quantidade']
    else:
        kpi = pd.DataFrame([['Coluna Nome_corrigido não encontrada', 0]], columns=['Nome_corrigido', 'Quantidade'])

    print("🔐 Autenticando no Google Sheets...")
    service = authenticate_google()

    print("✏️ Escrevendo dados limpos e validados na planilha...")
    # Seleciona colunas para enviar, ordenando por interesse (original + corrigidos + flags)
    colunas_envio = list(df.columns)
    clean_values = [colunas_envio] + df.astype(str).values.tolist()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=CLEAN_RANGE,
        valueInputOption='RAW',
        body={'values': clean_values}
    ).execute()
    print(f'✅ Dados limpos e validados enviados em: {CLEAN_RANGE}')

    print("✏️ Escrevendo estatísticas na planilha...")
    stats_values = [desc_reset.columns.tolist()] + desc_reset.astype(str).values.tolist()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=ANALYSIS_RANGE,
        valueInputOption='RAW',
        body={'values': stats_values}
    ).execute()

    start_row = len(stats_values) + 2
    print("📈 Escrevendo KPIs na planilha...")
    kpi_values = [kpi.columns.tolist()] + kpi.astype(str).values.tolist()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f'Analysis!A{start_row}',
        valueInputOption='RAW',
        body={'values': kpi_values}
    ).execute()
    print('✅ Estatísticas e KPIs enviados em: Analysis!A1')

if __name__ == '__main__':
    main()