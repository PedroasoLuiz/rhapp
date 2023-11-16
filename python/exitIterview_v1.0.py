from __future__ import print_function

import json
import os.path
import os
import sys
import time
import ctypes

# Documentation gspread
# https://docs.gspread.org/en/v5.10.0/user-guide.html#updating-cells
import gspread

from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file key.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Representam o id da planilha no Google drive e seu intervalo
Sheet_ID = '1VaB144DhBdK3BIewuV_CaBsxYlqMj_PxCBSNMyCAaGg'
Sheet_range = 'Respostas!A:AF'

# Obtém o nome do usuário ativo
username = os.getenv('USERNAME')

# # Define o local do download como a pasta raiz do usuário
path = "C:\\Users\\" + username + "\\RH APP\\scope\\json\\"

credentials = path + "credentials.json"
key = path + "key.json"

operation = sys.argv[1]
list_values = sys.argv[2:]
dictionary_values = [list_values]

def main():
    match operation:
        case "R":
            # Chama a função que atualiza a fonte de dados no arquivo json
            read_record()
        case "U":
            # Chama a função para atualizar a planilha com os valores
            update_record()
        case "D":
            # Chama a função para deletar o registro
            delete_record()

# Consulta
def read_record():
    try:
        # Faz o acesso à API pela chave de usuário
        gc = gspread.service_account(filename=key)
        sh = gc.open_by_key(Sheet_ID)
        ws = sh.worksheet('Respostas')
        values = ws.get_all_values()

        with open(path + "exitInterview_v1.0.json", 'w') as arquivo:
            json.dump(values, arquivo)

        print("Create file success.")

        # Ele aguarda 1,5 segundos para apagar o arquivo
        time.sleep(1.5)

        if os.path.exists(path + "exitInterview_v1.0.json"):
            os.remove(path + "exitInterview_v1.0.json")

        print("Remove file success.")
    except HttpError as err:
        print(err)


# Atualização dos dados
def update_record():
    try:
        # Ler informações do Google Sheets
        gc = gspread.service_account(filename=key)
        sh = gc.open_by_key(Sheet_ID)
        ws = sh.worksheet('Respostas APE')
        values = ws.get_all_values()

        linha_encontrada = -1

        for linha, row in enumerate(values, start=1):  # Comece com 1 para corresponder às linhas da planilha
            if row and row[2] == str(dictionary_values[0][1]):  # Procurando na terceira coluna (coluna C)
                linha_encontrada = linha
                break

        rangeUpdate = 'B' + str(linha_encontrada) + ':F' + str(linha_encontrada)

        # Se encontrou a linha
        if linha_encontrada != -1:
            # Certifique-se de que os valores a serem adicionados têm a mesma quantidade de colunas que os cabeçalhos
            ws.update(rangeUpdate, dictionary_values)
            ctypes.windll.user32.MessageBoxW(0, "Valores atualizados com sucesso!", "Sucesso", 0)
        else:
            ctypes.windll.user32.MessageBoxW(0, "O ID não foi encontrado na base de dados", "Atenção", 1)

    except HttpError as err:
        print(err)


def delete_record():
    try:
        # Ler informações do Google Sheets
        gc = gspread.service_account(filename=key)
        sh = gc.open_by_key(Sheet_ID)
        ws = sh.worksheet('Respostas APE')
        values = ws.get_all_values()

        linha_encontrada = -1

        for linha, row in enumerate(values, start=1):  # Comece com 1 para corresponder às linhas da planilha
            if row and row[2] == str(dictionary_values[0][0]):  # Procurando na terceira coluna (coluna C)
                linha_encontrada = linha
                break

        rangeUpdate = 'AH' + str(linha_encontrada)

        # Se encontrou a linha
        if linha_encontrada != -1:
            # Certifique-se de que os valores a serem adicionados têm a mesma quantidade de colunas que os cabeçalhos
            ws.update(rangeUpdate, dictionary_values[0][1])
            ctypes.windll.user32.MessageBoxW(0, "Registro inativado com sucesso!", "Sucesso", 0)
        else:
            ctypes.windll.user32.MessageBoxW(0, "O ID não foi encontrado na base de dados", "Atenção", 1)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
