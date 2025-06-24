from __future__ import print_function
import os.path
import pickle
import csv

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file']

def main():
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

    service = build('sheets', 'v4', credentials=creds)

    SPREADSHEET_ID = '1nSw2PbhMJOzu2b8R6Zzg0LL2-0UXRRjLYTc73i966QU'
    RANGE_NAME = 'Sheet1!A1:C10'

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('Nenhum dado encontrado.')
    else:
        print('Dados da planilha:')
        for row in values:
            print(row)

        # Gravar os dados no CSV
        with open('dados.csv', mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(values)
        print('\n✅ Arquivo CSV criado: dados.csv')

if __name__ == '__main__':
    main()