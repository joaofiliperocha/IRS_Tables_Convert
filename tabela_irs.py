# import xlrd
import pandas as pd
from collections import OrderedDict
import simplejson as json
import codecs


def setData(tx, ds, writeMode):
    taxas_retencao = dict()
    tabela = list()
    tab = OrderedDict()
    localRowNum = 0
    firstColVals = ''
    print("Extrair dados de ", tx)
    for index, col in ds.iterrows():
        #print(index, col[0])
        #print(isinstance(col[0], str))

        if isinstance(col[0], str):
            firstColVals = col[0].split(' ')
            #print(len(firstColVals), firstColVals[0])
            if firstColVals[0] == 'TABELAS':
                localRowNum = index

        #print(index - (localRowNum + 4))
        if index - (localRowNum + 4) == 0:
            escaloes = list()
            categoria = col[0]
        if col[0] == 'Até' or col[0] == 'Superior a':
            dependentes = OrderedDict()
            dependentes['max'] = col[1]
            dependentes['0'] = round(col[2], 3)
            dependentes['1'] = round(col[3], 3)
            dependentes['2'] = round(col[4], 3)
            dependentes['3'] = round(col[5], 3)
            dependentes['4'] = round(col[6], 3)
            dependentes['5'] = round(col[7], 3)
            escaloes.append(dependentes)
            tab[categoria] = escaloes

    tabela.append(tab)
    #print(tabela)
    taxas_retencao[tx] = tabela

    j = json.dumps(taxas_retencao)
    #print(j)
    print("Gravar dados de ", tx)
    if writeMode == 'a':
        j = "," + j
    with open('data.json', writeMode) as f:
        f.write(j)


CNT = pd.read_excel(
    'https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_Ret_IRS_2020_Continente.xlsx',
    sheet_name=0)
print("Tabela Continente carregada para CNT")
setData("CNT", CNT, 'w')

AZO = pd.read_excel(
    'https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_Ret_IRS_2020_Acores.xlsx',
    sheet_name=0)
print("Tabela Açores carregada para AZO")
setData("AZO", AZO, 'a')

MAD = pd.read_excel(
    'https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_Ret_IRS_2020_RAMadeira.xlsx',
    sheet_name=0)
print("Tabela Madeira carregada para MAD")
setData("MAD", MAD, 'a')
