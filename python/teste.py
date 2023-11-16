from __future__ import print_function

import os
import time
import requests

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Representam o id da planilha no Google drive e seu intervalo
SAMPLE_SPREADSHEET_ID = '1dSS2mm5IeFhpjVGee7_m6b0AESaKa99FeiUsgEE7xJw'
SAMPLE_RANGE_NAME = 'Respostas APE!A:AG'

# Obtém o nome do usuário ativo
username = os.getenv('USERNAME')

# Link de download do arquivo 'credentials.json' e 'token.json' diretamente do google drive
credenciais_url = "https://drive.google.com/uc?export=download&id=16YPZDeai06OJ5n_CGhhW6kW6FwaqMfaz"
token_url = "https://drive.google.com/uc?export=download&id=1KN267WBH1B9KQRjm3V3La2sLD40OcCqe"

# Define o local do download como a pasta raiz do usuário
path = "C:\\Users\\" + username + "\\"

# Seta o caminho para o arquivo
credenciais = path + "credentials.json"
oToken = path + "token.json"

# Realiza o download para a pasta raiz
response = requests.get(credenciais_url)
if response.status_code == 200:
    with open(credenciais, 'wb') as file:
        file.write(response.content)

response = requests.get(token_url)
if response.status_code == 200:
    with open(oToken, 'wb') as file:
        file.write(response.content)


def main():
    time.sleep(9.5)
