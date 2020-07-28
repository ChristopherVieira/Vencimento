# Vencimento

'''
Imports das bibliotecas
'''

import pandas as pd

from openpyxl import *



pv = ('Planilha.xlsm')
planilha = (pd.read_excel(pv, encoding='ISO-8859-1'))
#definir nomes para as colunas pq pqp

A = ("Unnamed: 0")
B =	("ACOMPANHAMENTO                 DE VALIDADES") #(Ean)
C =	("Unnamed: 2")#(Departamentos)
D =	("Unnamed: 3")#(Descrição)
E =	("Unnamed: 4")#(Quantidade)
F =	("Unnamed: 5")#(Lote ?)
G =	("Unnamed: 6")#(Data de validade)
H =	("Unnamed: 7")#(Dias para vencimento)
I =	("Unnamed: 8")#(Status (10 ou 20 ou 50 de desconto ou vencido))
J =	("Unnamed: 9")#(Tratado ou não)
K =	("Unnamed: 10")#(Sem tratativa/tipo ???)
L =	("Unnamed: 11")#(Sem tratativa QTD ???)
M =	("Unnamed: 12")#(Tratados)
N =	("Unnamed: 13")#()



"""
c = 0
for linha in planilha[I]:
    status = (planilha[I][c])
    c = c+1

    if status == "Controle":
        print("Item abaixo vai vencer daqui ")
        print(planilha[D][c])
        print(f"Melhorar a exposição, vai vencer daqui {planilha[H][c]} dias")
        print(30*"_-_-")
        print(c)

    elif status == "Baixa de 10%":
        print("Item abaixo com Baixa de 10%")
        print(planilha[D][c])
        print(f"Melhorar a exposição, vai vencer daqui {planilha[H][c]} dias")
        print(30 * "_-_-")

    elif status == "Baixa de 20%":
        print("Item abaixo com Baixa de 20%")
        print(planilha[D][c])
        print(f"Melhorar a exposição, vai vencer daqui {planilha[H][c]} dias")
        print(30 * "_-_-")

    elif status == "Baixa de 50%":
        print("Item abaixo com Baixa de 50%")
        print(planilha[D][c])
        print(f"Melhorar a exposição, vai vencer daqui {planilha[H][c]} dias")
        print(30 * "_-_-")

    elif status == "Vencido":
        print("Item abaixo Vencido")
        print(planilha[D][c])
        print(f"item vencido a {planilha[H][c]} dias, retirar da área de venda")
        print(30 * "_-_-")

"""



arquivo_excel = load_workbook('Planilha.xlsm',read_only=False, keep_vba= True)

planilha1 = arquivo_excel.active
planilha1['B14'] = 7891000049686
print(planilha1.cell(row=14, column=2).value)

ws = arquivo_excel.active
ws.security.

arquivo_excel.save('Planilha.xlsm')
c = 0

