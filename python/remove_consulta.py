import os
import sys

arquivo=sys.argv[1]
caminho_arquivo ="C:\\Users\\pedro.luiz\\OneDrive - NUTRIALFA\\Área de Trabalho\\Nova pasta\\arq\\json\\"

# Verificar se o arquivo existe antes de tentar excluí-lo
if os.path.exists(caminho_arquivo):
    try:
        os.remove(caminho_arquivo + arquivo)
        print(f'O arquivo {caminho_arquivo} foi excluído com sucesso.')
    except Exception as e:
        print(f"Ocorreu um erro ao excluir o arquivo: {str(e)}")
else:
    print(f'O arquivo {caminho_arquivo} não existe.')
