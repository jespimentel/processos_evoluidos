import os
import pandas as pd

pecas = pd.read_excel('pecas_em_elaboracao.xlsx') # Lê a planilha gerada pelo SharePoint
pecas['Name'] = pecas['Name'].apply(lambda x: x[:25]) # Pega o número do processo
lista_pecas = set(pecas['Name'].to_list()) # Cria o conjunto de peças existentes no SharePoint

print('O SharePoint tem', len(lista_pecas), 'peças em elaboração')

pastas_pecas_concluidas = [
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2021\\alegacoes_finais', 
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2021\\contrarrazoes',
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2021\\razoes'
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2020\\Alegações finais',
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2020\\Contrarrazões',
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2020\\Razões',
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2019\\Alegações finais',
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2019\\Contrarrazões',
    'D:\\Du\\OneDrive - Ministério Público SP\\_Promotoria\\2019\\Razões',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2018\\AF',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2018\\Contrarrazões',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2018\\Razões',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2017\\Alegações finais',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2017\\Contrarrazões',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2017\\Razões',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2016\\Alegacoes finais',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2016\\Contrarrazoes',
    'D:\\Du\\OneDrive - Ministério Público SP\\Transferidos - Personal\\2016\\Razoes',
    ]

pecas_concluidas = set()

def relaciona_arquivos(pasta):
    relacao = []
    for root, dirs, files in os.walk(pasta):
        for file in files:
            relacao.append(file[:25]) # Somente o número do processo
    return set(relacao)

for pasta in pastas_pecas_concluidas:
    pecas_concluidas = pecas_concluidas | relaciona_arquivos(pasta)

repetidas = sorted(lista_pecas.intersection(pecas_concluidas))

arquivo = open ('processos_evoluidos.txt', 'w')
for peca in repetidas:
    arquivo.write (peca +'\n')
arquivo.close()

print ('Concluído!')