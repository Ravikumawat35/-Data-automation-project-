import pandas as pd
import numpy as np
import sys, os
import os.path
import time
from datetime import date


def month_converter(month):
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    return months.index(month) + 1


def extrairpdf(path):

    print(path)
    df = pd.read_excel(path, sheet_name='Form')
    baixarelatorio = []
    baixarelatorio.append(df['Unnamed: 6'][20])  # ocorrencia
    baixarelatorio.append(df['Unnamed: 6'][22])  # unidade
    baixarelatorio.append(df['Unnamed: 6'][25])  # sintoma
    baixarelatorio.append(df['Unnamed: 6'][15])  # relatorio
    baixarelatorio.append(df['Unnamed: 6'][12])  # responsavel

    dataa = df['Unnamed: 10'][12]
    dataa = str(dataa)
    dia = dataa[8:10]
    mes = dataa[5:7]
    ano = dataa[0:4]
    dataa = f'{dia}/{mes}/{ano}'
    baixarelatorio.append(dataa)  # data

    databaixa = time.ctime(os.path.getctime(path))
    dia = databaixa[8:10]
    mes = month_converter(databaixa[4:7])
    ano = databaixa[-4:]
    straux = '{}/{}/{}'.format(dia, mes, ano)

    baixarelatorio.append(straux)  # data baixa

    baixarelatorio.append(df['Unnamed: 6'][14])  # tipo

    return baixarelatorio


def extrair6m(path):

    print(path)
    df = pd.read_excel(path, header=2)

    df = df.loc[:, 'N° Ocorrência':'Tipo da Análise']
    df = df.dropna()
    # df = df.iloc[1:]
    df.drop(index=df.index[-1:], axis=0, inplace=True)

    df.insert(2, 1, "", allow_duplicates=True)
    df.insert(4, 2, "", allow_duplicates=True)
    df.insert(7, "Data Baixa", "", allow_duplicates=True)

    dataa = df['Data de Realização']
    dataa = str(dataa)
    dia = dataa[13:15]
    mes = dataa[10:12]
    ano = dataa[5:9]
    dataa = f'{dia}/{mes}/{ano}'

    if '-' in dataa:
        dataa = df['Data de Realização']
        dataa = str(dataa)
        dia = dataa[12:14]
        mes = dataa[9:11]
        ano = dataa[5:8]
        dataa = f'{dia}/{mes}/{ano}'

    df['Data de Realização'] = dataa

    databaixa = time.ctime(os.path.getctime(path))
    dia = databaixa[8:10]
    mes = month_converter(databaixa[4:7])
    ano = databaixa[-4:]
    straux = '{}/{}/{}'.format(dia, mes, ano)

    df['Data Baixa'] = straux

    df = df.astype({"N° Ocorrência": int, "Vagão": int})

    total_rows = df['Data de Realização'].count()
    total_rows = int(total_rows)

    conv = df.to_numpy()

    return conv, total_rows


# MAIN ----------------------------------------------------------------------
if __name__ == "__main__":

    directory = r'_:\_\_\_'

    baixa = []
    cpdf = 0
    c6m = 0

    for file in os.listdir(directory):
        if file.endswith(".xlsm") or file.endswith(".xls") or file.endswith(".xlsx"):
            if '6M' in file:
                path = os.path.join(directory, file)
                aux, cont = extrair6m(path)
                flat_list = [item for sublist in aux for item in sublist]
                baixa.append(flat_list)
                c6m += cont

            else:
                path = os.path.join(directory, file)
                baixa.append(extrairpdf(path))
                cpdf += 1

    nlin = cpdf + c6m
    print(nlin)

    flat = [item for sublist in baixa for item in sublist]
    flat = np.reshape(flat, (nlin, 9))

    dfbaixa = pd.DataFrame(flat)

    pd.DataFrame(dfbaixa).to_csv(f"saida_vagoes_{date.today()}.csv")
