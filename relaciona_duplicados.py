import os
import pandas as pd

pecas = pd.read_excel('pecas_em_elaboracao.xlsx') # Lê a planilha gerada pelo SharePoint
lista_pecas = set(pecas['Name'].to_list()) # Cria o conjunto de peças existentes no SharePoint

pastas_pecas_concluidas = [
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2021', 
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2020',
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2019',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2018',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2017',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2016'
    ]

pecas_concluidas = set()

def relaciona_arquivos(pasta):
    relacao = []
    for root, dirs, files in os.walk(pasta):
        for file in files:
            relacao.append(file)
    return set(relacao)

for pasta in pastas_pecas_concluidas:
    pecas_concluidas = pecas_concluidas | relaciona_arquivos(pasta)

repetidas = sorted(lista_pecas.intersection(pecas_concluidas))

arquivo = open ('repetidas.txt', 'w')
for peca in repetidas:
    arquivo.write (peca +'\n')
arquivo.close()