import io
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.layout import LAParams
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage

import pandas as pd
import tkinter as tk
from tkinter import filedialog
import numpy as np
import sys, os
import pathlib
import os.path
import time
from datetime import date

def pdf_to_text(path):
    with open(path, 'rb') as fp:
        rsrcmgr = PDFResourceManager()
        outfp = io.StringIO()
        laparams = LAParams()
        device = TextConverter(rsrcmgr, outfp, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.get_pages(fp):
            interpreter.process_page(page)
    text = outfp.getvalue()
    return text

# MAIN ----------------------------------------------------------------------
if __name__ == "__main__":

    directory = r'_:\_\_\_'
    relatorios = []

    for file in os.listdir(directory):
        if file.endswith(".pdf"):
            path = os.path.join(directory, file)

            pdfconv = pdf_to_text(path)
            pdfconv = pdfconv.split("\n")

            responsavel = pdfconv[50]
            data = pdfconv[40]
            datains = data
            tipo = pdfconv[54]
            numerorelatorio = pdfconv[56]
            numeroocorrencia = pdfconv[58]
            vagao = pdfconv[60]
            localidade = numerorelatorio[5:7]
            sintoma = pdfconv[63]

            dadosrelatorio = responsavel, data, datains, tipo, numerorelatorio, numeroocorrencia, vagao, localidade, sintoma

            relatorios.append(dadosrelatorio)

    print(relatorios)

    for file in os.listdir(directory):
        if file.endswith(".xlsx" or ".xls" or ".xlsb" or ".xlsm"):
            path = os.path.join(directory, file)
            df = pd.read_excel(path, index_col=0, dtype={"nome": str, "cargo": str, "idade": int, "altura": float})

            conv1 = df.to_records(index=True)
            nline = len(conv1)

            # CONVERTER DADOS DE PANDAS PARA NUMPY - TRATAMENTO
            conv = df.to_numpy()
            print(conv)

            np.savetxt('baixas_atualizado.csv', conv, delimiter=',')
