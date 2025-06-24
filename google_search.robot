*** Settings ***
Library    RPA.Google.Sheets

*** Variables ***
${SPREADSHEET_ID}    sua_planilha_id_aqui

*** Test Cases ***
Ler Dados Do Google Sheets
    Authenticate Google Sheets
    ${rows}=    Read Worksheet    ${SPREADSHEET_ID}    Sheet1
    Log    ${rows}
