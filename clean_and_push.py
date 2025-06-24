from __future__ import print_function
import os
import pickle
import pandas as pd
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Configurações
CSV_FILE = 'dados.csv'
SPREADSHEET_ID = '1nSw2PbhMJOzu2b8R6Zzg0LL2-0UXRRjLYTc73i966QU'
OUTPUT_RANGE = 'Sheet1!A1'  # altere se quiser outra aba/célula inicial
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_service():
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

def main():
    print('🔄 Lendo dados do CSV...')
    df = pd.read_csv(CSV_FILE)

    print('🧹 Limpando dados...')
    df.columns = df.columns.str.strip()
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].str.strip()

    print('🔐 Autenticando no Google Sheets...')
    service = get_service()

    values = [df.columns.tolist()] + df.values.tolist()
    body = {'values': values}

    print('✏️ Escrevendo dados limpos na planilha...')
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=OUTPUT_RANGE,
        valueInputOption='RAW',
        body=body
    ).execute()
    print('✅ Dados atualizados em:', OUTPUT_RANGE)

if __name__ == '__main__':
    main()