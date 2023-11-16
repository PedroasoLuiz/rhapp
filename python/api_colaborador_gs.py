from __future__ import print_function

import os.path
import sys
import json

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1fPESmAJ_L3YjEu8KPR-7QoZh4U9L1cMqaixoGJoCIxM'
SAMPLE_RANGE_NAME = 'db_Colaborador!A:I'

credenciais = "C:\\Users\\pedro.luiz\\OneDrive - NUTRIALFA\\Área de Trabalho\\Nova pasta\\arq\\json\\credentials.json"
oToken = "C:\\Users\\pedro.luiz\\OneDrive - NUTRIALFA\\Área de Trabalho\\Nova pasta\\arq\\json\\token.json"

operacao = sys.argv[1]
list_values = sys.argv[2:]
dictionary_values = \
    [
        list_values
    ]


def main():
    match operacao:
        case "C":
            # Chama a função para cadastrar um novo valor na planilha
            create_record(dictionary_values)
        case "R":
            # Chama a função que atualiza a fonte de dados no arquivo json
            read_record()
        case "U":
            # Chama a função para atualizar a planilha com os valores
            update_record(dictionary_values)
        case "D":
            # Chama a função para deletar o registro
            delete_record(list_values[1])


# Cadastro
def create_record(valores):
    print("Cadastrado")


# Consulta
def read_record():
    global creds
    if os.path.exists(oToken):
        creds = Credentials.from_authorized_user_file(oToken, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credenciais, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(oToken, 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Ler informacoes do Google Sheets
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        valores = result['values']
        with open("C:\\Users\\pedro.luiz\\OneDrive - NUTRIALFA\\Área de Trabalho\\Nova pasta\\arq\\json\\api_colaborador_gs.json",
                  'w') as arquivo:
            json.dump(valores, arquivo)
    except HttpError as err:
        print(err)


# Atualização dos dados
def update_record(valor_atualizado):
    global creds
    if os.path.exists(oToken):
        creds = Credentials.from_authorized_user_file(oToken, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credenciais, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(oToken, 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Ler informações do Google Sheets
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()

        values = result.get('values', [])
        linha_encontrada = -1

        for linha, row in enumerate(values, start=1):  # Comece com 1 para corresponder às linhas da planilha
            if row and row[0] == str(dictionary_values[0][0]):  # Procurando na terceira coluna (coluna C)
                linha_encontrada = linha
                break

        # Se encontrou a linha
        if linha_encontrada != -1:
            # Certifique-se de que os valores a serem adicionados têm a mesma quantidade de colunas que os cabeçalhos
            sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="A" + str(linha_encontrada),
                                  valueInputOption="USER_ENTERED",
                                  body={'values': valor_atualizado}).execute()
            #ctypes.windll.user32.MessageBoxW(0, "Valores atualizados com sucesso!" , "Sucesso", 0)
        else:
            print("")
            #ctypes.windll.user32.MessageBoxW(0, "O ID não foi encontrado na base de dados", "Atenção",1)

    except HttpError as err:
        print(err)


def delete_record(id):
    print("")


if __name__ == '__main__':
    main()
