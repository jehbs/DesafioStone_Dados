'''
Desafio Stone

Jéssica Barbosa de Souza

Analise dos chargeback

'''
from openpyxl import load_workbook
from openpyxl.cell import Cell
from datetime import datetime
from datetime import timedelta
import datetime
import pandas as pd
import numpy as np

dadosTransacao = {}
dadosGeralaba2 =[]
dadosGeralCall=[]
dadosValor = []
dadosCartao = []
dadosNaopadronizado = []


#Carregando os dados
wb = load_workbook('Missão Stone - Dados de trx.xlsx')
ws = wb['Aba 2']
# montando listas
for i in range(2,11821):

    if isinstance(ws._get_cell(i, 2).value, datetime.datetime):
        dadosTransacao = ({'Data':datetime.datetime.combine(ws._get_cell(i,1).value,ws._get_cell(i,2).value.time()),'Valor':ws._get_cell(i,3).value,'Cartao':ws._get_cell(i,4).value,'CBK':ws._get_cell(i,5).value,'Causa':"-"})
    else:
        dadosTransacao = (
            {'Data': datetime.datetime.combine(ws._get_cell(i, 1).value, ws._get_cell(i, 2).value),
             'Valor': ws._get_cell(i, 3).value, 'Cartao': ws._get_cell(i, 4).value, 'CBK': ws._get_cell(i, 5).value,
             'Causa': "-"})

    dadosGeralaba2.append(dadosTransacao)

for i in range(0,len(dadosGeralaba2)):

    if ((dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 1]['Cartao']) and (
            dadosGeralaba2[i]['Valor'] == dadosGeralaba2[i - 1]['Valor'])) :
        dadosGeralaba2[i]['Causa'] = "ValorRepetido"
        dadosGeralaba2[i - 1]['Causa'] = "ValorRepetido"
    elif (dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 4]['Cartao']) and (
            dadosGeralaba2[i]['Valor'] == dadosGeralaba2[i - 4]['Valor']):
        dadosGeralaba2[i]['Causa'] = "ValorRepetido"
        dadosGeralaba2[i - 4]['Causa'] = "ValorRepetido"
    elif (dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 5]['Cartao']) and (
            dadosGeralaba2[i]['Valor'] == dadosGeralaba2[i - 5]['Valor']):
        dadosGeralaba2[i]['Causa'] = "ValorRepetido"
        dadosGeralaba2[i - 5]['Causa'] = "ValorRepetido"
    elif (dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 6]['Cartao']) and (
            dadosGeralaba2[i]['Valor'] == dadosGeralaba2[i - 6]['Valor']):
        dadosGeralaba2[i]['Causa'] = "ValorRepetido"
        dadosGeralaba2[i - 6]['Causa'] = "ValorRepetido"
    elif ((dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 1]['Cartao']) and (
            dadosGeralaba2[i]['Data'] < dadosGeralaba2[i - 1]['Data'] + timedelta(seconds=300))):
        dadosGeralaba2[i]['Causa'] = "TentativaSucessiva"
        dadosGeralaba2[i - 1]['Causa'] = "TentativaSucessiva"
    elif ((dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 2]['Cartao']) and (
            dadosGeralaba2[i]['Data'] < dadosGeralaba2[i - 2]['Data'] + timedelta(seconds=300))):
        dadosGeralaba2[i]['Causa'] = "TentativaSucessiva"
        dadosGeralaba2[i - 2]['Causa'] = "TentativaSucessiva"
    elif ((dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 3]['Cartao']) and (
            dadosGeralaba2[i]['Data'] < dadosGeralaba2[i - 3]['Data'] + timedelta(seconds=300))):
        dadosGeralaba2[i]['Causa'] = "TentativaSucessiva"
        dadosGeralaba2[i - 3]['Causa'] = "TentativaSucessiva"
    elif ((dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 4]['Cartao']) and (
            dadosGeralaba2[i]['Data'] < dadosGeralaba2[i - 4]['Data'] + timedelta(seconds=300))):
        dadosGeralaba2[i]['Causa'] = "TentativaSucessiva"
        dadosGeralaba2[i - 4]['Causa'] = "TentativaSucessiva"
    elif ((dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 5]['Cartao']) and (
            dadosGeralaba2[i]['Data'] < dadosGeralaba2[i - 5]['Data'] + timedelta(seconds=300))):
        dadosGeralaba2[i]['Causa'] = "TentativaSucessiva"
        dadosGeralaba2[i - 5]['Causa'] = "TentativaSucessiva"
    elif ((dadosGeralaba2[i]['Cartao'] == dadosGeralaba2[i - 6]['Cartao']) and (
            dadosGeralaba2[i]['Data'] < dadosGeralaba2[i - 6]['Data'] + timedelta(seconds=300))):
        dadosGeralaba2[i]['Causa'] = "TentativaSucessiva"
        dadosGeralaba2[i - 6]['Causa'] = "TentativaSucessiva"

for i in range(0,len(dadosGeralaba2)):
    if dadosGeralaba2[i]['Causa'] == '-':
        dadosGeralaba2[i]["CBK"] = "Não"
    else:
        dadosGeralaba2[i]["CBK"] = "Sim"
        dadosGeralCall.append(dadosGeralaba2[i])

PercentualCharge = (len(dadosGeralCall)/len(dadosGeralaba2))*100
print("sucesso")