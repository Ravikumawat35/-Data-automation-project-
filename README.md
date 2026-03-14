import pandas as pd
import numpy as np
import sys
import os
import pathlib
import os.path
import time
from datetime import date
import datetime
import glob
import csv
import shutil
from qvd import qvd_reader
import warnings


def month_converter(month):
    """
    Converte a formatacao do mes nas datas do datetime para numeros - usada na montagem da caixa preta
    """
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    return months.index(month) + 1


def backups(kmed_c, solicitados_c, lit_c):
    today = datetime.date.today()
    shutil.copyfile(kmed_c, r'C:\_\backup_dw\Base_KMED_{}.xlsx'.format(today))
    shutil.copyfile(solicitados_c, r'C:\_\backup_dw\Base_Solicitados_{}.xlsx'.format(today))
    shutil.copyfile(lit_c, r'C:\_\backup_dw\Base_LIT_{}.xlsx'.format(today))


def leitura_dw_drive(directory, ext):
    """
    Funcao para receber o diretorio e buscar todos os arquivos que estao nesse diretorio
    e em seus subdiretorios que tem a extensao especifica que queremos.
    :param directory: diretorio que será buscado
    :param ext: extensao do arquivo
    :return: lista de arquivos [numero loco, data modif], tamanho dessa lista
    """

    CxPreta = []
    CxPretaMD = []
    CxPretaDate = []

    searchfile = glob.glob(directory + f"/**/*.{ext}", recursive=True)
    CxPreta.append(searchfile)

    tam = np.shape(CxPreta)
    CxPreta = np.reshape(CxPreta, (tam[0]*tam[1], 1))

    tam = int(len(CxPreta[:, 0]))

    i = 0
    while i < tam:
        a = CxPreta[i]
        a = str(a)
        a = a[2:-2]
        CxPreta[i] = a
        a = a.encode()
        datamodif = time.ctime(os.path.getmtime(a))
        CxPretaMD.append(datamodif)

        dia = CxPretaMD[i][8:10]
        mes = month_converter(CxPretaMD[i][4:7])
        ano = CxPretaMD[i][-4:]
        horario = CxPretaMD[i][11:19]

        straux = '{}/{}/{} {}'.format(dia, mes, ano, horario)
        CxPretaDate.append(straux)

        i += 1

    BaseCaixaPreta = []
    i = 0
    while i < len(CxPretaDate):
        BaseCaixaPreta.append([CxPreta[i], CxPretaDate[i]])
        i += 1

    return BaseCaixaPreta, tam


def gerar_cx_preta(path_RE):
    """
    recebe os parametros para enviar para a funcao 'leitura_dw_drive()' e formata os dados para
    exportar para csv a base caixa preta do drive Q.
    tambem gera uma lista para cada processador no formato [num_loco, data modif arquivo]
    :param path_RE: downloads RE estao no automaweb e nao na rede, este é o caminho do arquivo com as baixas
    :return: listas de cada processador no formato [num_loco, data modif arquivo]
    """
    CxPretaPy = []
    pulse, tam_pulse = leitura_dw_drive(r'Q:\_\PULSE (.DAT)', 'dat')
    locotrol, tam_locotrol = leitura_dw_drive(r'Q:\_\LOCOTROL (DP-EAB)', 'bin')
    locotrolLXA, tam_locotrolLXA = leitura_dw_drive(r'Q:\_\LOCOTROL LXA', 'tar')
    zeit, tam_zeit = leitura_dw_drive(r'Q:\_\PAINEL ZEIT', 'rar')
    sdis, tam_sdis = leitura_dw_drive(r'Q:\_\SDIS -  Versão PORTUGUÊS (.snp .inc .ru)', 'inc')
    cab, tam_cab = leitura_dw_drive(r'Q:\_\CAB (FLT .VAX .EGU)', 'flt')
    d7up, tam_d7up = leitura_dw_drive(r'Q:\_\D7UP (.TXT)', 'txt')
    atc, tam_atc = leitura_dw_drive(r'Q:\_\ATC_CPTM', 'txt')

    # MONTAR ARQUIVO CAIXA PRETA FINAL
    CxPretaPy.append(pulse)
    CxPretaPy.append(locotrol)
    CxPretaPy.append(locotrolLXA)
    CxPretaPy.append(zeit)
    CxPretaPy.append(sdis)
    CxPretaPy.append(cab)
    CxPretaPy.append(d7up)
    CxPretaPy.append(atc)

    tamf = tam_pulse + tam_locotrol + tam_locotrolLXA + tam_zeit + tam_sdis + tam_cab + tam_d7up + tam_atc
    print(f'Quantidade total de downloads no drive: {tamf}')

    flat = [item for sublist in CxPretaPy for item in sublist]
    flat = [item for sublist in flat for item in sublist]
    flat = np.reshape(flat, (tamf, 2))

    today = datetime.date.today().strftime("%d.%m.%y")
    pd.DataFrame(flat).to_csv(f"CxPreta_{today}.csv", index=False, encoding='utf-8')

    with open(r'CxPreta_BI.csv', 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(flat)

    pulse = np.array(pulse)
    for i in range(tam_pulse):
        pulse[i, 1] = str(pulse[i, 1])
        pulse[i, 0] = tratar_string_cxpreta(pulse[i, 0], f=False)

    locotrol = np.array(locotrol)
    for i in range(tam_locotrol):
        locotrol[i, 1] = str(locotrol[i, 1])
        locotrol[i, 0] = tratar_string_cxpreta(locotrol[i, 0], f=False)

    locotrolLXA = np.array(locotrolLXA)
    for i in range(tam_locotrolLXA):
        locotrolLXA[i, 1] = str(locotrolLXA[i, 1])
        locotrolLXA[i, 0] = tratar_string_cxpreta(locotrolLXA[i, 0], f=False)

    zeit = np.array(zeit)
    for i in range(tam_zeit):
        zeit[i, 1] = str(zeit[i, 1])
        zeit[i, 0] = tratar_string_cxpreta(zeit[i, 0], f=False)

    sdis = np.array(sdis)
    for i in range(tam_sdis):
        sdis[i, 1] = str(sdis[i, 1])
        sdis[i, 0] = tratar_string_cxpreta(sdis[i, 0], f=False)

    cab = np.array(cab)
    for i in range(tam_cab):
        cab[i, 1] = str(cab[i, 1])
        cab[i, 0] = tratar_string_cxpreta(cab[i, 0], f=True)

    d7up = np.array(d7up)
    for i in range(tam_d7up):
        d7up[i, 1] = str(d7up[i, 1])
        d7up[i, 0] = tratar_string_cxpreta(d7up[i, 0], f=True)

    atc = np.array(atc)
    for i in range(tam_atc):
        atc[i, 1] = str(atc[i, 1])
        atc[i, 0] = tratar_string_cxpreta(atc[i, 0], f=False)

    # DOWNLOADS RE --------------------------------------------------------------- \\\\\\\\\\\\\\\\\
    dw_RE = pd.read_excel(path_RE, header=1)
    dw_RE = dw_RE.dropna(how='all', axis=1)
    dw_RE = dw_RE[dw_RE["Status"] == 'FINALIZADO']
    dw_RE = dw_RE[dw_RE.columns.intersection(['Número de Locomotiva', 'Data de Importação'])]
    dw_RE = dw_RE.to_numpy()
    for i in range(len(dw_RE)):
        aux = str(dw_RE[i][0])
        aux = aux[:6]
        dw_RE[i][0] = aux
        dw_RE[i][1] = str(dw_RE[i][1])
        dw_RE[i][1] = datetime.datetime.strptime(dw_RE[i][1], '%Y-%m-%d %H:%M:%S')
        dw_RE[i][1] = dw_RE[i][1].strftime('%d/%m/%Y %H:%M:%S')

    info_dw = [pulse, locotrol, locotrolLXA, zeit, sdis, cab, d7up, dw_RE, atc]

    return info_dw


def tratar_string_cxpreta(str_cx, f):
    """
    Funcao auxilixar para transformar o caminho do arquivo num numero de locomotiva com 6 digitos
    :param str_cx: string original
    :param f: flag, pois o tratamento da string muda a depender do tipo de processador
    :return: string tratada. exemplo: 900624
    """

    if f:
        str_cx = str(str_cx)
        str_cx = str_cx.split(r'\\\\')
        str_cx = str_cx[-1]
        str_cx = str_cx[3:7]
        str_cx = '90' + str_cx
        str_cx = ''.join(ch for ch in str_cx if ch.isdigit())

    else:
        str_cx = str(str_cx)
        str_cx = str_cx.split(r'\\\\')
        str_cx = str_cx[-1]
        str_cx = str_cx[:4]
        str_cx = '90' + str_cx
        str_cx = ''.join(ch for ch in str_cx if ch.isdigit())

    return str_cx


def baixa_processadores(i, base, loco, data_oco):
    """
    Funcao auxiliar da execucao do KMED, responsavel por
    comparar as datas de download e ocorrencia para realizar a baixa

    :param i: contador do loop que a funcao é chamada, significa o indice da linha da base kmed atualizada
    :param base: lista [n loco, data baixa], como pulse_dw
    :param kmed_atualizado: baixa kmed atualizada e processada com pendencias
    :return: data da baixa mais proxima da data da ocorrencia ou pendencia, caso não tenha dw valido
    """

    dw_proc_e = []
    for k in range(len(base)):
        if int(base[k][0]) == loco:
            # dw_proc_e: lista com os horarios dos downloads do processador em loop da locomotiva que esta no loop
            dw_proc_e.append(base[k][1])

    if not len(dw_proc_e) == 0:
        dw_validos = []
        for r in range(len(dw_proc_e)):
            # DATETIME COMPARA AS MAIORES DATAS E SOBRESCREVE
            data_dw = datetime.datetime.strptime(dw_proc_e[r], '%d/%m/%Y %H:%M:%S')

            if data_dw > data_oco:
                dw_validos.append(data_dw)

        if not len(dw_validos) == 0:
            return min(dw_validos)
        else:
            return 'Pendente'
    else:
        return 'Pendente'


def processadores(path_frota):
    """
    Criar um array que diz quais processadores estao em cada locomotiva
    :param path_frota: caminho do arquivoQVD para se acessar a frota MRS atualizada pelo qlik
    :return: array [locos com pulse, locos com locotrol, ...]
    """
    frota = qvd_reader.read(path_frota)
    pd.DataFrame(frota).to_excel(r'FROTA_DW.xlsx', index=True, encoding='utf-8')
    frota = frota.drop(columns=["Data"])
    frota["Locomotiva"] = pd.to_numeric(frota["Locomotiva"])
    tamanho_frota = len(frota.index)

    locos_com_pulse = []
    locos_com_locotrol = []
    locos_com_sdis = []
    locos_com_zeit = []
    locos_com_cab = []
    locos_com_d7up = []
    locos_com_re = []

    for i in range(tamanho_frota):
        if frota["Registrador"][i] == 'Pulse':
            locos_com_pulse.append(int(frota["Locomotiva"][i]))

        if frota["Registrador"][i] == 'Atan (Accenture)':
            locos_com_re.append(int(frota["Locomotiva"][i]))

        if frota["Microproc"][i] == 'MICROPROC. CIO':
            locos_com_sdis.append(int(frota["Locomotiva"][i]))

        if frota["Microproc"][i] == 'ZEIT - SAL-03':
            locos_com_zeit.append(int(frota["Locomotiva"][i]))

        if frota["Microproc"][i] == 'MICROPROC. CAB':
            locos_com_cab.append(int(frota["Locomotiva"][i]))

        if frota["Microproc"][i] == 'BRIGHTSTAR - D7up':
            locos_com_d7up.append(int(frota["Locomotiva"][i]))

        if frota["Locotrol"][i] == 'Sim':
            locos_com_locotrol.append(int(frota["Locomotiva"][i]))

    info_frota = [locos_com_pulse, locos_com_locotrol, locos_com_sdis, locos_com_zeit, locos_com_cab,
                  locos_com_d7up, locos_com_re]

    return info_frota


def execucao_kmed(info_dw, info_frota, path_base_kmed, path_base_atu):
    """
    LEITURA DA BASE KMED, DO ARQUIVO DE ATUALIZACAO QLIK -> BASE ATUALIZADA COM PENDENCIAS, BAIXAS E CALCULO DE PRAZO

    :param info_dw: lista de listas [pulse_dw, locotrol_dw, locotrolLXA_dw, zeit_dw, sdis_dw, cab_dw, d7up_dw, re_dw]
    :param info_frota:  lista de listas [locos_com_pulse, ..., locos_com_re]
    :param path_base_kmed: caminho do arquivo de base raiz do controle
    :param path_base_atu: caminho do arquivo de base de atualizacao do controle
    :return: None
    """
    # ABRE ARQUIVO BASE E ARQUIVO DE ATUALIZACAO --------------------------------- \\\\\\\\\\\\\\\\\

    kmed_new = pd.read_excel(path_base_atu, sheet_name='Sheet1')
    kmed = pd.read_excel(path_base_kmed, sheet_name='Base')
    cabecalho = kmed.columns

    # PREPARACAO DESSAS BASES -------------------------------------------------- \\\\\\\\\\\\\\\\\

    kmed_lastrow = kmed.iloc[-1]
    ultima_ocorrencia = kmed_lastrow.iat[0]

    ult_oco_index = (kmed_new[kmed_new["num(Ocorrência)"] == ultima_ocorrencia]).index
    ult_oco_index = ult_oco_index[0] + 1

    tamanho_base_atualizacao = len(kmed_new.index)

    ocorrencias_novas = kmed_new.tail(tamanho_base_atualizacao - ult_oco_index)
    ocorrencias_novas.drop(ocorrencias_novas.tail(1).index, inplace=True)

    for i in range(15, 23):
        ocorrencias_novas[i] = "-"

    ocorrencias_novas.columns = kmed.columns

    # CRIAR LISTAS DE LOCOMOTIVAS COM CADA PROCESSADOR --------------------------- \\\\\\\\\\\\\\\\\
    locos_com_pulse, locos_com_locotrol, locos_com_sdis = info_frota[0], info_frota[1], info_frota[2]
    locos_com_zeit, locos_com_cab, locos_com_d7up = info_frota[3], info_frota[4], info_frota[5]
    locos_com_re = info_frota[6]

    pulse_dw, locotrol_dw, locotrolLXA_dw, zeit_dw = info_dw[0], info_dw[1], info_dw[2], info_dw[3]
    sdis_dw, cab_dw, d7up_dw, re_dw = info_dw[4], info_dw[5], info_dw[6], info_dw[7]

    sintomas_locotrol = ['TRAÇÃO DISTRIBUÍDA - NÃO REPÕE VÁLVULA DE FREIO', 'FREIO ELETRÔNICO AVARIADO',
                         'VÁLVULA "PCS" NÃO REARMA - (COMANDADA)', 'NÃO ALIVIA FREIO',
                         'TRAÇÃO DISTRIBUÍDA', 'NÃO APLICA FREIO']

    # IDENTIFICAR PENDENCIAS PARA CADA PROCESSADOR NA BASE DE ATUALIZACAO -------- \\\\\\\\\\\\\\\\\

    if len(ocorrencias_novas) == 0:
        print("=-=-=-=-=-=--=-= SEM ATUALIZACOES HOJE =-=-=-=-=-=--=-=")

    else:
        for j in range(ocorrencias_novas.index[0], ocorrencias_novas.index[-1]+1):
            ocorrencias_novas["Locomotiva"][j] = int(ocorrencias_novas["Locomotiva"][j])
            if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_pulse:
                ocorrencias_novas["Pulse"][j] = 'Pendente'

            if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_re:
                ocorrencias_novas["RE"][j] = 'Pendente'

            if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_zeit:
                ocorrencias_novas["Zeit ( microp)"][j] = 'Pendente'

            if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_sdis:
                ocorrencias_novas["SDIS ( microp)"][j] = 'Pendente'

            if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_cab:
                ocorrencias_novas["CAB ( micro)"][j] = 'Pendente'

            if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_d7up:
                ocorrencias_novas["D7up (microp)"][j] = 'Pendente'

            if (int(ocorrencias_novas["Locomotiva"][j]) in locos_com_locotrol) and (ocorrencias_novas["Sintoma"][j] in sintomas_locotrol):
                ocorrencias_novas["Locotrol"][j] = 'Pendente'

    # ADICIONAR OCORRENCIAS NOVAS E AS PENDENCIAS NA BASE ANTERIOR ---------------- \\\\\\\\\\\\\\\\\

    kmed = kmed.to_numpy()
    kmed = np.array(kmed)
    ocorrencias_novas = ocorrencias_novas.to_numpy()
    ocorrencias_novas = np.array(ocorrencias_novas)
    tamanho_base_atualizada = np.shape(kmed)[0] + np.shape(ocorrencias_novas)[0]
    kmed_atualizado = [kmed, ocorrencias_novas]

    flat = [item for sublist in kmed_atualizado for item in sublist]
    flat = [item for sublist in flat for item in sublist]
    kmed_atualizado = np.reshape(flat, (tamanho_base_atualizada, 22))

    # PESQUISAR PENDENCIAS DE TODAS AS OCORRENCIAS E ATUALIZAR ------------------- \\\\\\\\\\\\\\\\\

    baixa = 0
    # O loop percorre linha a linha, confere cada processador daquela ocorrencia e passa para a proxima
    for i in range(len(kmed_atualizado)):
        for c in range(14, 21):
            if not kmed_atualizado[i][c] == '-':
                if not kmed_atualizado[i][c] == 'Pendente':
                    if not isinstance(kmed_atualizado[i][c], datetime.datetime):
                        kmed_atualizado[i][c] = 'Pendente'

        # PULSE
        if kmed_atualizado[i][14] == 'Pendente':
            kmed_atualizado[i][14] = baixa_processadores(i, pulse_dw, kmed_atualizado[i][1], kmed_atualizado[i][3])
            baixa += 1
        # LOCOTROL E LOCOTROL_LXA
        if kmed_atualizado[i][15] == 'Pendente':
            # tamaux = len(locotrol_dw) + len(locotrolLXA_dw)
            # locotrol_dw = [locotrol_dw, locotrolLXA_dw]
            # flat = [item for sublist in locotrol_dw for item in sublist]
            # locotrol_dw = [item for sublist in flat for item in sublist]
            # locotrol_dw = np.reshape(locotrol_dw, (tamaux, 2))
            kmed_atualizado[i][15] = baixa_processadores(i, locotrol_dw, kmed_atualizado[i][1], kmed_atualizado[i][3])
            baixa += 1
        # ZEIT
        if kmed_atualizado[i][16] == 'Pendente':
            kmed_atualizado[i][16] = baixa_processadores(i, zeit_dw, kmed_atualizado[i][1], kmed_atualizado[i][3])
            baixa += 1
        # SDIS
        if kmed_atualizado[i][17] == 'Pendente':
            kmed_atualizado[i][17] = baixa_processadores(i, sdis_dw, kmed_atualizado[i][1], kmed_atualizado[i][3])
            baixa += 1
        # CAB
        if kmed_atualizado[i][18] == 'Pendente':
            kmed_atualizado[i][18] = baixa_processadores(i, cab_dw, kmed_atualizado[i][1], kmed_atualizado[i][3])
            baixa += 1
        # D7UP
        if kmed_atualizado[i][19] == 'Pendente':
            kmed_atualizado[i][19] = baixa_processadores(i, d7up_dw, kmed_atualizado[i][1], kmed_atualizado[i][3])
            baixa += 1
        # RE
        if kmed_atualizado[i][20] == 'Pendente':
            kmed_atualizado[i][20] = baixa_processadores(i, re_dw, kmed_atualizado[i][1], kmed_atualizado[i][3])
            baixa += 1

        # CALCULAR PRAZO E ADERENCIA ------------------------------------------------- \\\\\\\\\\\\\\\\\
        baixas_da_loco = []
        if not kmed_atualizado[i][21] == 2:
            if 'Pendente' in kmed_atualizado[i][:]:
                kmed_atualizado[i][21] = 1
            else:
                for j in range(14, 21):
                    if not kmed_atualizado[i][j] == '-':
                        if isinstance(kmed_atualizado[i][j], datetime.datetime):
                            baixas_da_loco.append(kmed_atualizado[i][j])  # criar lista de datas
                if len(baixas_da_loco) == 0:
                    kmed_atualizado[i][21] = 0
                else:
                    ultimo_dw = max(baixas_da_loco)  # pegar a maior data da lista, o ultimo download
                    data_ocorrencia = kmed_atualizado[i][3]  # pegar a data da ocorrencia e criar condicional de prazo

                    diferenca_ocorrencia_dw = ultimo_dw - data_ocorrencia

                    # comparar a maior data de download e calcular aderencia
                    data_ocorrencia = str(data_ocorrencia)
                    data_ocorrencia = datetime.datetime.strptime(data_ocorrencia, '%Y-%m-%d %H:%M:%S')
                    dia_ocorrencia = data_ocorrencia.day

                    # Prazos para realização de download KMED:
                    # dia 01 a 19 - 5 dias, dia 20 a 23 - 4 dias, dia 24 a 31 - 3 dias

                    if dia_ocorrencia < 20:
                        prazo = 5
                    elif dia_ocorrencia < 24:
                        prazo = 4
                    else:
                        prazo = 3

                    diferenca_em_segundos = diferenca_ocorrencia_dw.total_seconds()
                    dias = diferenca_em_segundos / (60*60*24)
                    # print(dias)
                    if dias <= prazo:
                        kmed_atualizado[i][21] = 2
                    else:
                        kmed_atualizado[i][21] = 1

    for i in range(len(kmed_atualizado)):
        for j in range(14, 21):
            if isinstance(kmed_atualizado[i][j], datetime.datetime):
                kmed_atualizado[i][j] = str(kmed_atualizado[i][j])

    print(f"total de ocorrências novas: {len(ocorrencias_novas)}")
    print(f"total de baixas novas: {baixa}")

    # SALVAR BASE FINAL ATUALIZADA ----------------------------------------------- \\\\\\\\\\\\\\\\\

    df = pd.DataFrame(kmed_atualizado)
    writer = pd.ExcelWriter(path_base_kmed, engine='xlsxwriter')
    df.to_excel(writer, index=False, encoding='utf-8', header=cabecalho, sheet_name='Base')
    writer.save()

    # ----------------------------------------------------------------------------------------------


def execucao_solicitados(info_dw, info_frota, path_base_solicitados, path_base_atu):
    """
    LEITURA DA BASE, DO ARQUIVO DE ATUALIZACAO FORMS -> BASE ATUALIZADA COM PENDENCIAS, BAIXAS E CALCULO DE PRAZO

    :param info_dw: lista de listas [pulse_dw, locotrol_dw, locotrolLXA_dw, zeit_dw, sdis_dw, cab_dw, d7up_dw, re_dw]
    :param info_frota:  lista de listas [locos_com_pulse, ..., locos_com_re]
    :param path_base_solicitados: caminho do arquivo de base raiz do controle
    :param path_base_atu: caminho do arquivo de base de atualizacao do controle, a planilha do formulario
    :return: None
    """

    # ABRE ARQUIVO BASE E ARQUIVO DE ATUALIZACAO --------------------------------- \\\\\\\\\\\\\\\\\

    solic_new = pd.read_excel(path_base_atu, sheet_name='Form1')
    solic = pd.read_excel(path_base_solicitados, sheet_name='Solicitados')
    cabecalho = solic.columns

    # PREPARACAO DESSAS BASES -------------------------------------------------- \\\\\\\\\\\\\\\\\
    
    solic_lastrow = solic.iloc[-1]
    ultimo_index = solic_lastrow[0]

    solic_new_lastrow = solic_new.iloc[-1]
    ultimo_new_index = solic_new_lastrow[0]

    tamanho_base_atualizacao = ultimo_new_index - ultimo_index
    if tamanho_base_atualizacao == 0:
        tamanho_base_atualizacao = 1

    ocorrencias_novas = solic_new.tail(tamanho_base_atualizacao-1)
    ocorrencias_novas = ocorrencias_novas[['ID', 'Número da Locomotiva', 'Hora de conclusão', 'Nome', 'Caixa Preta',
                                           'Locotrol', 'Microproc.', 'Microproc. Nível 3', 'ATC', 'Zeit', 'CBL']]

    for i in range(12, 20):
        ocorrencias_novas[i] = "-"
    ocorrencias_novas.columns = solic.columns

    # tratar os numeros de locomotivas inseridos pelo usuario para colocar no padrao de 6 digitos
    solic_new_locos = ocorrencias_novas.to_numpy()
    solic_new_locos = solic_new_locos[:, 1]
    for i in range(len(solic_new_locos)):
        # Tratar numero de locomotiva em caso de preenchimento errado do usuario
        if len(str(solic_new_locos[i])) == 4:
            solic_new_locos[i] = '90' + str(solic_new_locos[i])
            solic_new_locos[i] = int(solic_new_locos[i])
        if len(str(solic_new_locos[i])) == 7:
            aux = str(solic_new_locos[i])
            solic_new_locos[i] = aux[:-1]
            solic_new_locos[i] = int(solic_new_locos[i])
        if len(str(solic_new_locos[i])) == 5:
            aux = str(solic_new_locos[i])
            solic_new_locos[i] = aux[:-1]
            solic_new_locos[i] = '90' + str(solic_new_locos[i])
            solic_new_locos[i] = int(solic_new_locos[i])
    solic_new_locos = pd.DataFrame(solic_new_locos)
    solic_new_locos = solic_new_locos.astype(int)
    ocorrencias_novas['Locomotiva'] = solic_new_locos[0].values

    # CRIAR LISTAS DE LOCOMOTIVAS COM CADA PROCESSADOR --------------------------- \\\\\\\\\\\\\\\\\
    locos_com_pulse, locos_com_locotrol, locos_com_sdis = info_frota[0], info_frota[1], info_frota[2]
    locos_com_zeit, locos_com_cab, locos_com_d7up = info_frota[3], info_frota[4], info_frota[5]
    locos_com_re = info_frota[6]
    pulse_dw, locotrol_dw, locotrolLXA_dw, zeit_dw = info_dw[0], info_dw[1], info_dw[2], info_dw[3]
    sdis_dw, cab_dw, d7up_dw, re_dw, atc_dw = info_dw[4], info_dw[5], info_dw[6], info_dw[7], info_dw[8]

    # IDENTIFICAR PENDENCIAS PARA CADA PROCESSADOR NA BASE DE ATUALIZACAO -------- \\\\\\\\\\\\\\\\\
    if len(ocorrencias_novas) == 0:
        print("=-=-=-=-=-=--=-= SEM ATUALIZACOES HOJE =-=-=-=-=-=--=-=")

    else:
        for j in range(ocorrencias_novas.index[0], ocorrencias_novas.index[-1]+1):
            if ocorrencias_novas["Cx. Preta"][j] == 'Sim':
                if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_pulse:
                    ocorrencias_novas["Pulse"][j] = 'Pendente'
                elif int(ocorrencias_novas["Locomotiva"][j]) in locos_com_re:
                    ocorrencias_novas["RE"][j] = 'Pendente'
                else:
                    ocorrencias_novas["Pulse"][j] = 'ERRO'  # conferencia se realmente há o processador
                    ocorrencias_novas["RE"][j] = 'ERRO'

            if ocorrencias_novas["Locotrol"][j] == 'Sim':
                if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_locotrol:
                    ocorrencias_novas["Locotrol.1"][j] = 'Pendente'
                else:
                    ocorrencias_novas["Locotrol.1"][j] = 'ERRO'

            if ocorrencias_novas["Microproc."][j] == 'Sim' or ocorrencias_novas["Microproc. Nivel 3"][j] == 'Sim':
                if int(ocorrencias_novas["Locomotiva"][j]) in locos_com_zeit:
                    ocorrencias_novas["Zeit ( microp)"][j] = 'Pendente'
                elif int(ocorrencias_novas["Locomotiva"][j]) in locos_com_sdis:
                    ocorrencias_novas["SDIS ( microp)"][j] = 'Pendente'
                elif int(ocorrencias_novas["Locomotiva"][j]) in locos_com_cab:
                    ocorrencias_novas["CAB ( micro)"][j] = 'Pendente'
                elif int(ocorrencias_novas["Locomotiva"][j]) in locos_com_d7up:
                    ocorrencias_novas["D7up (microp)"][j] = 'Pendente'
                else:
                    ocorrencias_novas["Zeit ( microp)"][j] = 'ERRO'
                    ocorrencias_novas["SDIS ( microp)"][j] = 'ERRO'
                    ocorrencias_novas["CAB ( micro)"][j] = 'ERRO'
                    ocorrencias_novas["D7up (microp)"][j] = 'ERRO'

    # ADICIONAR OCORRENCIAS NOVAS E AS PENDENCIAS NA BASE ANTERIOR ---------------- \\\\\\\\\\\\\\\\\

    solic = solic.to_numpy()
    solic = np.array(solic)
    ocorrencias_novas = ocorrencias_novas.to_numpy()
    ocorrencias_novas = np.array(ocorrencias_novas)
    tamanho_base_atualizada = np.shape(solic)[0] + np.shape(ocorrencias_novas)[0]
    solic_atualizado = [solic, ocorrencias_novas]

    flat = [item for sublist in solic_atualizado for item in sublist]
    flat = [item for sublist in flat for item in sublist]
    solic_atualizado = np.reshape(flat, (tamanho_base_atualizada, 19))

    baixas = 0
    # PESQUISAR PENDENCIAS DE TODAS AS OCORRENCIAS E ATUALIZAR ------------------- \\\\\\\\\\\\\\\\\
    # O loop percorre linha a linha, confere cada processador daquela ocorrencia e passa para a proxima
    for i in range(len(solic_atualizado)):
        # Tratar numero de locomotiva em caso de preenchimento errado do usuario
        if len(str(solic_atualizado[i][1])) == 4:
            solic_atualizado[i][1] = '90' + str(solic_atualizado[i][1])
            solic_atualizado[i][1] = int(solic_atualizado[i][1])
        if len(str(solic_atualizado[i][1])) == 7:
            aux = str(solic_atualizado[i][1])
            solic_atualizado[i][1] = aux[:-1]
            solic_atualizado[i][1] = int(solic_atualizado[i][1])

        # PESQUISAR DATAS E SOBRESCREVER

        # Em caso de Pendencia lançada corrompida
        for c in range(11, 18):
            if not solic_atualizado[i][c] == '-':
                if not solic_atualizado[i][c] == 'Pendente':
                    if not isinstance(solic_atualizado[i][c], datetime.datetime):
                        solic_atualizado[i][c] = 'Pendente'
        # PULSE
        if solic_atualizado[i][11] == 'Pendente':
            solic_atualizado[i][11] = baixa_processadores(i, pulse_dw, solic_atualizado[i][1], solic_atualizado[i][2])
            baixas += 1
        # LOCOTROL E LOCOTROL_LXA
        if solic_atualizado[i][12] == 'Pendente':
            # tamaux = len(locotrol_dw) + len(locotrolLXA_dw)
            # locotrol_dw = [locotrol_dw, locotrolLXA_dw]
            # flat = [item for sublist in locotrol_dw for item in sublist]
            # locotrol_dw = [item for sublist in flat for item in sublist]
            # locotrol_dw = np.reshape(locotrol_dw, (tamaux, 2))
            solic_atualizado[i][12] = baixa_processadores(i, locotrol_dw, solic_atualizado[i][1], solic_atualizado[i][2])
            baixas += 1
        # ZEIT
        if solic_atualizado[i][13] == 'Pendente':
            solic_atualizado[i][13] = baixa_processadores(i, zeit_dw, solic_atualizado[i][1], solic_atualizado[i][2])
            baixas += 1
        # SDIS
        if solic_atualizado[i][14] == 'Pendente':
            solic_atualizado[i][14] = baixa_processadores(i, sdis_dw, solic_atualizado[i][1], solic_atualizado[i][2])
            baixas += 1
        # CAB
        if solic_atualizado[i][15] == 'Pendente':
            solic_atualizado[i][15] = baixa_processadores(i, cab_dw, solic_atualizado[i][1], solic_atualizado[i][2])
            baixas += 1
        # D7UP
        if solic_atualizado[i][16] == 'Pendente':
            solic_atualizado[i][16] = baixa_processadores(i, d7up_dw, solic_atualizado[i][1], solic_atualizado[i][2])
        # RE
        if solic_atualizado[i][17] == 'Pendente':
            solic_atualizado[i][17] = baixa_processadores(i, re_dw, solic_atualizado[i][1], solic_atualizado[i][2])
            baixas += 1
        # ATC
        if solic_atualizado[i][8] == 'Sim':
            solic_atualizado[i][8] = baixa_processadores(i, atc_dw, solic_atualizado[i][1], solic_atualizado[i][2])
            baixas += 1

        # CALCULAR PRAZO E ADERENCIA ------------------------------------------------- \\\\\\\\\\\\\\\\\
        baixas_da_loco = []
        if not solic_atualizado[i][18] == 2:
            if 'Pendente' in solic_atualizado[i][:]:
                solic_atualizado[i][18] = 1
            elif solic_atualizado[i][8] == 'Sim':
                solic_atualizado[i][18] = 1
            elif solic_atualizado[i][10] == 'Sim':
                solic_atualizado[i][18] = 1
            else:
                for j in range(8, 18):
                    if not solic_atualizado[i][j] == '-':
                        if isinstance(solic_atualizado[i][j], datetime.datetime):
                            baixas_da_loco.append(solic_atualizado[i][j])  # criar lista de datas
                if len(baixas_da_loco) == 0:
                    solic_atualizado[i][18] = 0
                else:
                    ultimo_dw = max(baixas_da_loco)  # pegar a maior data da lista, o ultimo download
                    data_ocorrencia = solic_atualizado[i][2]  # pegar a data da ocorrencia e criar condicional de prazo

                    diferenca_ocorrencia_dw = ultimo_dw - data_ocorrencia

                    # comparar a maior data de download e calcular aderencia
                    data_ocorrencia = str(data_ocorrencia)
                    data_ocorrencia = datetime.datetime.strptime(data_ocorrencia, '%Y-%m-%d %H:%M:%S')
                    dia_ocorrencia = data_ocorrencia.day

                    # Prazos para realização de download Solicitados:
                    # até dia 23 - 5 dias, dia 24 a 27 - 6 dias, dia 28 a 31 - 7 dias
                    if dia_ocorrencia < 24:
                        prazo = 5
                    elif (dia_ocorrencia > 23 and dia_ocorrencia < 28):
                        prazo = 6
                    else:
                        prazo = 7

                    diferenca_em_segundos = diferenca_ocorrencia_dw.total_seconds()
                    dias = diferenca_em_segundos / (60*60*24)

                    if dias <= prazo:
                        solic_atualizado[i][18] = 2
                    else:
                        solic_atualizado[i][18] = 1

    for i in range(len(solic_atualizado)):
        if isinstance(solic_atualizado[i][8], datetime.datetime):
            solic_atualizado[i][8] = str(solic_atualizado[i][8])
        for j in range(11, 18):
            if isinstance(solic_atualizado[i][j], datetime.datetime):
                solic_atualizado[i][j] = str(solic_atualizado[i][j])

    print(f"total de ocorrências novas: {len(ocorrencias_novas)}")
    print(f"total de baixas novas: {baixas}")

    # SALVAR BASE FINAL ATUALIZADA ----------------------------------------------- \\\\\\\\\\\\\\\\\

    df = pd.DataFrame(solic_atualizado)
    writer = pd.ExcelWriter(path_base_solicitados, engine='xlsxwriter')
    df.to_excel(writer, index=False, encoding='utf-8', header=cabecalho, sheet_name='Solicitados')
    writer.save()

    # ----------------------------------------------------------------------------------------------


def execucao_lit(info_dw, info_frota, path_base_lit, path_base_visitas_novas):
    """
    LEITURA DA BASE LIT, DO ARQUIVO DE ATUALIZACAO -> BASE ATUALIZADA COM PENDENCIAS, BAIXAS E CALCULO DE PRAZO
    :param info_dw: lista de listas [pulse_dw, locotrol_dw, locotrolLXA_dw, zeit_dw, sdis_dw, cab_dw, d7up_dw, re_dw]
    :param info_frota: lista de listas [locos_com_pulse, ..., locos_com_re]
    :param path_base_lit: caminho do arquivo de base raiz do controle
    :param path_base_visitas_novas: caminho do arquivo de base de atualizacao do controle
    :return:
    """
    # DECLARACAO DA ATRIBUICAO DE CADA PATIO PARA SUA RESPECTIVA GERENCIA--------- \\\\\\\\\\\\\\\\\
    MG = ['FCB', 'FAF', 'FDM', 'FOO', 'FPK', 'FEE']
    RJ = ['FFI', 'FSP', 'FOJ', 'FBE']
    SP = ['IEE', 'IJL', 'ZKL', 'FFM']
    OFICINA = ['FPJ', 'FHL', 'FBK', 'IJI', 'IOH']

    now = datetime.datetime.now()

    # ABRE ARQUIVO BASE E ARQUIVO DE ATUALIZACAO --------------------------------- \\\\\\\\\\\\\\\\\
    lit_new = pd.read_excel(path_base_visitas_novas, sheet_name='Folha 1', header=2)
    lit = pd.read_excel(path_base_lit, sheet_name='LIT')

    cabecalho = lit.columns

    # PREPARACAO DESSAS BASES ---------------------------------------------------- \\\\\\\\\\\\\\\\\

    lit_lastrow = lit.iloc[-1]
    ultima_visita = lit_lastrow.iat[2]

    ult_vis_index = (lit_new[lit_new["Visita"] == ultima_visita]).index
    ult_vis_index = ult_vis_index[0] + 1

    tamanho_base_atualizacao = len(lit_new.index)

    visitas_novas = lit_new.tail(tamanho_base_atualizacao - ult_vis_index)
    visitas_novas.drop(visitas_novas.columns[12], axis=1, inplace=True)

    for i in range(13, 21):
        visitas_novas[i] = "-"
        if i == 20:
            visitas_novas[i] = "Pendente"

    visitas_novas.columns = lit.columns

    # CRIAR LISTAS DE LOCOMOTIVAS COM CADA PROCESSADOR --------------------------- \\\\\\\\\\\\\\\\\
    locos_com_pulse, locos_com_locotrol, locos_com_sdis = info_frota[0], info_frota[1], info_frota[2]
    locos_com_zeit, locos_com_cab, locos_com_d7up = info_frota[3], info_frota[4], info_frota[5]
    locos_com_re = info_frota[6]

    pulse_dw, locotrol_dw, locotrolLXA_dw, zeit_dw = info_dw[0], info_dw[1], info_dw[2], info_dw[3]
    sdis_dw, cab_dw, d7up_dw, re_dw = info_dw[4], info_dw[5], info_dw[6], info_dw[7]

    # IDENTIFICAR PENDENCIAS PARA CADA PROCESSADOR NA BASE DE ATUALIZACAO -------- \\\\\\\\\\\\\\\\\

    if len(visitas_novas) == 0:
        print("=-=-=-=-=-=--=-= SEM ATUALIZACOES HOJE =-=-=-=-=-=--=-=")

    else:
        for j in range(visitas_novas.index[0], visitas_novas.index[-1] + 1):
            if int(visitas_novas["Locomotiva"][j]) in locos_com_pulse:
                visitas_novas["Pulse"][j] = 'Pendente'

            if int(visitas_novas["Locomotiva"][j]) in locos_com_re:
                visitas_novas["RE"][j] = 'Pendente'

            if int(visitas_novas["Locomotiva"][j]) in locos_com_zeit:
                visitas_novas["Zeit"][j] = 'Pendente'

            if int(visitas_novas["Locomotiva"][j]) in locos_com_sdis:
                visitas_novas["SDIS"][j] = 'Pendente'

            if int(visitas_novas["Locomotiva"][j]) in locos_com_cab:
                visitas_novas["CAB"][j] = 'Pendente'

            if int(visitas_novas["Locomotiva"][j]) in locos_com_d7up:
                visitas_novas["D7up"][j] = 'Pendente'

    # ADICIONAR OCORRENCIAS NOVAS E AS PENDENCIAS NA BASE ANTERIOR ---------------- \\\\\\\\\\\\\\\\\

    lit = lit.to_numpy()
    lit = np.array(lit)
    visitas_novas = visitas_novas.to_numpy()
    visitas_novas = np.array(visitas_novas)
    tamanho_base_atualizada = np.shape(lit)[0] + np.shape(visitas_novas)[0]
    lit_atualizado = [lit, visitas_novas]

    flat = [item for sublist in lit_atualizado for item in sublist]
    flat = [item for sublist in flat for item in sublist]
    lit_atualizado = np.reshape(flat, (tamanho_base_atualizada, 20))

    # PESQUISAR PENDENCIAS DE TODAS AS OCORRENCIAS E ATUALIZAR ------------------- \\\\\\\\\\\\\\\\\

    baixa = 0
    # O loop percorre linha a linha, confere cada processador daquela ocorrencia e passa para a proxima
    for i in range(len(lit_atualizado)):
        # Tratamento data da visita dessa linha
        data_vis = lit_atualizado[i][6]
        hora_vis = str(lit_atualizado[i][7])
        hora_vis = datetime.datetime.strptime(hora_vis, '%H:%M:%S')
        data_vis = datetime.datetime(data_vis.year, data_vis.month, data_vis.day, hora_vis.hour, hora_vis.minute)

        # Atribuicao do patio para uma gerencia
        if lit_atualizado[i][9] in MG:
            lit_atualizado[i][12] = 'MG'
        elif lit_atualizado[i][9] in RJ:
            lit_atualizado[i][12] = 'RJ'
        elif lit_atualizado[i][9] in SP:
            lit_atualizado[i][12] = 'SP'
        elif lit_atualizado[i][9] in OFICINA:
            lit_atualizado[i][12] = 'OFICINA'
        else:
            lit_atualizado[i][12] = 'Erro'

        # BAIXAS --------------------------------------------------------------------- \\\\\\\\\\\\\\\\\
        for c in range(13, 19):
            if not lit_atualizado[i][c] == '-':
                if not lit_atualizado[i][c] == 'Pendente':
                    if not isinstance(lit_atualizado[i][c], datetime.datetime):
                        lit_atualizado[i][c] = 'Pendente'
        # PULSE
        if lit_atualizado[i][13] == 'Pendente':
            lit_atualizado[i][13] = baixa_processadores(i, pulse_dw, lit_atualizado[i][0], data_vis)
            baixa += 1
        # ZEIT
        if lit_atualizado[i][14] == 'Pendente':
            lit_atualizado[i][14] = baixa_processadores(i, zeit_dw, lit_atualizado[i][0], data_vis)
            baixa += 1
        # SDIS
        if lit_atualizado[i][15] == 'Pendente':
            lit_atualizado[i][15] = baixa_processadores(i, sdis_dw, lit_atualizado[i][0], data_vis)
            baixa += 1
        # CAB
        if lit_atualizado[i][16] == 'Pendente':
            lit_atualizado[i][16] = baixa_processadores(i, cab_dw, lit_atualizado[i][0], data_vis)
            baixa += 1
        # D7UP
        if lit_atualizado[i][17] == 'Pendente':
            lit_atualizado[i][17] = baixa_processadores(i, d7up_dw, lit_atualizado[i][0], data_vis)
            baixa += 1
        # RE
        if lit_atualizado[i][18] == 'Pendente':
            lit_atualizado[i][18] = baixa_processadores(i, re_dw, lit_atualizado[i][0], data_vis)
            baixa += 1

        # CALCULAR PRAZO E ADERENCIA ------------------------------------------------- \\\\\\\\\\\\\\\\\
        baixas_da_loco = []
        # o prazo eh o dia seguinte da visita ate as 23:59:59
        prazo = data_vis + datetime.timedelta(days=1)
        prazo = prazo.replace(hour=23, minute=59, second=59)

        if lit_atualizado[i][19] == 'Pendente':
            # if prazo < now + datetime.timedelta(days=1):
            #     lit_atualizado[i][19] = 'Não Coletado'
            # else:
            for j in range(13, 20):
                if not lit_atualizado[i][j] == '-':
                    if isinstance(lit_atualizado[i][j], datetime.datetime):
                        baixas_da_loco.append(lit_atualizado[i][j])  # criar lista de datas
            if all(element == '-' for element in lit_atualizado[i][13:19]):
                lit_atualizado[i][19] = 'Erro Frota'
            elif 'Pendente' in lit_atualizado[i][13:19]:
                lit_atualizado[i][19] = 'Pendente'
            else:
                ultimo_dw = max(baixas_da_loco)  # pegar a maior data da lista, o ultimo download
                if ultimo_dw <= prazo:
                    lit_atualizado[i][19] = 'Coletado'
                else:
                    lit_atualizado[i][19] = 'Não coletado'

    for i in range(len(lit_atualizado)):
        for j in range(13, 19):
            if isinstance(lit_atualizado[i][j], datetime.datetime):
                lit_atualizado[i][j] = str(lit_atualizado[i][j])

    print(f"total de visitas novas: {len(visitas_novas)}")
    print(f"total de baixas novas: {baixa}")

    # SALVAR BASE FINAL ATUALIZADA ----------------------------------------------- \\\\\\\\\\\\\\\\\

    df = pd.DataFrame(lit_atualizado)
    writer = pd.ExcelWriter(path_base_lit, engine='xlsxwriter')
    df.to_excel(writer, index=False, encoding='utf-8', header=cabecalho, sheet_name='LIT')
    writer.save()

    # ----------------------------------------------------------------------------------------------


# MAIN ------------------------------------------------------------------------------------------
if __name__ == '__main__':
    # GERAR BASE CAIXA PRETA --------------------------------------------------------------------
    # Leitura dos arquivos nas pastas dos processadores PULSE, LOCOTROL, ZEIT, SDIS, CAB e D7UP
    # Montagem de arquivo resposta csv na pasta do controle no drive R no formato (Nome, Data de Modificação)

    # LEITURA DA BASE KMED, DO ARQUIVO DE ATUALIZACAO QLIK -> BASE ATUALIZADA COM PENDENCIAS, BAIXAS E CALCULO DE PRAZO

    # CAMINHO PARA CARREGAR A ATUALIZACAO DA FROTA - QVD
    path_frota = r'\\repositorioqlik\Qlik\_\Frota atualizada.qvd'

    # CAMINHOS DE ARQUIVO PARA EXECUTAR O KMED
    path_base_atu = r'\\repositorioqlik\Qlik\_\Atualização Controle de DW.xlsx'
    path_base_kmed = r'C:\_\Base_KMED.xlsx'
    path_RE = r'C:\_\Relatorio.xlsx'

    # CAMINHOS DE ARQUIVO PARA EXECUTAR O SOLICITADOS
    path_base_solitados = r'C:\_\Base_Solicitados.xlsx'
    path_base_solicitacoes_novas = r'C:\_\Solicitações de Downloads - Controle EngManut.xlsx'

    # CAMINHOS DE ARQUIVO PARA EXECUTAR O LIT
    path_base_lit = r'C:\_\Base_LIT.xlsx'
    path_base_visitas_novas = r'C:\_\Lit - dw.xls'

    info_dw = gerar_cx_preta(path_RE)

    processadores_lista = ['pulse', 'locotrol', 'locotrolLXA', 'zeit', 'sdis', 'cab', 'd7up', 're', 'atc']

    # CRIAR LISTAS DE LOCOMOTIVAS COM CADA PROCESSADOR --------------------------- \\\\\\\\\\\\\\\\\
    info_frota = processadores(path_frota)

    # FAZER BACKUPS DO DIA ANTERIOR EM CASO DE ERRO
    backups(path_base_kmed, path_base_solitados, path_base_lit)

    # DAR BAIXAS
    print("_____________________EXECUCAO_KMED__________________________")
    execucao_kmed(info_dw, info_frota, path_base_kmed, path_base_atu)
    print("__________________EXECUCAO_SOLICITADOS______________________")
    execucao_solicitados(info_dw, info_frota, path_base_solitados, path_base_solicitacoes_novas)
    print("______________________EXECUCAO_LIT__________________________")
    execucao_lit(info_dw, info_frota, path_base_lit, path_base_visitas_novas)
    print("_=_=_=_=_=_=_=_=_RODAMOS_SEM_PROBLEMAS_=_=_=_=_=_=_=_=_=_=_")

    # PONTOS DE MELHORIA ----------------------------------------------------------------------------------------------\
    #
    # O QUE AINDA É MANUAL?
    # 1 baixar a base do RE (automaweb), a base dos solicitados (onedrive), Localizacao-0057 e LIT (discoverer)
    # 2 Tracao distribuida (insercao das outras locos da composicao na base)
    #       IMPORTANTE: NAO INSERIR NOVAS LOCOMOTIVAS NO FINAL DA BASE, DEIXAR A ULTIMA SEMPRE PRESERVADA PELO
    #                   PREENCHIMENTO DO PYTHON NO DIA ANTERIOR
    # 3 SOLICITADOS: baixa de CBL -> despadronizacao na nomeacao dos arquivos e solicitacao pouquissimo frequente
    #
    # 4 Criar um clock para agendar o play e automatizar a rodagem do script
    #
    # POSSIVEIS ERROS?
    # 1 base do qlik eventualmente perde uma ocorrencia, se for a ultima do dia anterior, terá erro na atualizacao
    #       COMO RESOLVER: basta inverter a ultima e a penultima linhas manualmente na base
    # 2 nao ter a locomotiva na frota -> existe uma flag em cada um para isso: indice 0 ou ERRO em vez da aderencia
    #       COMO RESOLVER: observar se nao há pendencia alguma na linha -> inserir loco na frota e rodar novamente
    # 3 DATETIME: pela inconsistência nas bases, as vezes a coluna de data vem num formato diferente -> ajuste manual
    #
    # LOCOTROL LXA!
    # assim que o atendimento externo tiver o cabo e puder realizar os downloads desse processador
    # descomentar as linhas 396-400, 597-601
    #
    #  ----------------------------------------------------------------------------------------------------------------/
