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
dadosGeralAba1 =[]
dadosGeralCall=[]
dadosValor = []
dadosCartao = []
dadosNaopadronizado = []


#Carregando os dados
wb = load_workbook('Missão Stone - Dados de trx.xlsx')
ws = wb['Aba 1']
# montando listas
for i in range(2,11129):

    if ws._get_cell(i,4).value =="Sim" or  ws._get_cell(i,4).value =="Não":
        dadosTransacao = (
        {'Data':ws._get_cell(i,1).value,'Valor': ws._get_cell(i, 2).value,
         'Cartao': ws._get_cell(i,3).value, 'CBK': ws._get_cell(i, 4).value, 'Causa': "-"})
        dadosValor.append(ws._get_cell(i, 2).value)
        dadosCartao.append(ws._get_cell(i,3).value)
    else:
        if isinstance(ws._get_cell(i, 2).value, datetime.datetime):
            dadosTransacao = ({'Data':datetime.datetime.combine(ws._get_cell(i,1).value,ws._get_cell(i,2).value.time()),'Valor':ws._get_cell(i,3).value,'Cartao':ws._get_cell(i,4).value,'CBK':ws._get_cell(i,5).value,'Causa':"-"})
        else:
            dadosTransacao = (
            {'Data': datetime.datetime.combine(ws._get_cell(i, 1).value, ws._get_cell(i, 2).value),
             'Valor': ws._get_cell(i, 3).value, 'Cartao': ws._get_cell(i, 4).value, 'CBK': ws._get_cell(i, 5).value,
             'Causa': "-"})

        dadosValor.append(ws._get_cell(i, 3).value)
        dadosCartao.append(ws._get_cell(i, 4).value)
    dadosGeralAba1.append(dadosTransacao)
    dadosCartao.append(ws._get_cell(i,4).value)

for i in range(0,len(dadosGeralAba1)):
    if dadosGeralAba1[i]['CBK'] == "Sim":
        if ((dadosGeralAba1[i]['Cartao']  == dadosGeralAba1[i-1]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i-1]['Valor']) )or((dadosGeralAba1[i]['Cartao']  == dadosGeralAba1[i-2]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i-2]['Valor'])) or ((dadosGeralAba1[i]['Cartao']  == dadosGeralAba1[i-3]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i-3]['Valor'])):
            dadosGeralAba1[i]['Causa'] ="ValorRepetido"
            dadosGeralAba1[i - 1]['Causa'] = "ValorRepetido"
        elif (dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i+1]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i+1]['Valor']):
            dadosGeralAba1[i]['Causa'] = "ValorRepetido"
            dadosGeralAba1[i+1]['Causa'] = "ValorRepetido"
        elif (dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i+2]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i+2]['Valor']):
            dadosGeralAba1[i]['Causa'] = "ValorRepetido"
            dadosGeralAba1[i+2]['Causa'] = "ValorRepetido"
        elif (dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i+3]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i+3]['Valor']):
            dadosGeralAba1[i]['Causa'] = "ValorRepetido"
            dadosGeralAba1[i+3]['Causa'] = "ValorRepetido"
        elif (dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i-4]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i-4]['Valor']):
            dadosGeralAba1[i]['Causa'] = "ValorRepetido"
            dadosGeralAba1[i-4]['Causa'] = "ValorRepetido"
        elif (dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i-5]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i-5]['Valor']):
            dadosGeralAba1[i]['Causa'] = "ValorRepetido"
            dadosGeralAba1[i-5]['Causa'] = "ValorRepetido"
        elif (dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i-6]['Cartao']) and (dadosGeralAba1[i]['Valor']==dadosGeralAba1[i-6]['Valor']):
            dadosGeralAba1[i]['Causa'] = "ValorRepetido"
            dadosGeralAba1[i-6]['Causa'] = "ValorRepetido"
        elif((dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i-1]['Cartao']) and ( dadosGeralAba1[i]['Data'] <dadosGeralAba1[i-1]['Data']+timedelta(seconds=300))):
            dadosGeralAba1[i]['Causa'] = "TentativaSucessiva"
            dadosGeralAba1[i - 1]['Causa'] = "TentativaSucessiva"
        elif ((dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i - 2]['Cartao']) and (
                dadosGeralAba1[i]['Data'] < dadosGeralAba1[i - 2]['Data'] + timedelta(seconds=300))):
            dadosGeralAba1[i]['Causa'] = "TentativaSucessiva"
            dadosGeralAba1[i - 2]['Causa'] = "TentativaSucessiva"
        elif ((dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i - 3]['Cartao']) and (
                dadosGeralAba1[i]['Data'] < dadosGeralAba1[i - 3]['Data'] + timedelta(seconds=300))):
            dadosGeralAba1[i]['Causa'] = "TentativaSucessiva"
            dadosGeralAba1[i-3]['Causa'] = "TentativaSucessiva"
        elif((dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i - 4]['Cartao']) and (
                 dadosGeralAba1[i]['Data'] < dadosGeralAba1[i - 4]['Data'] + timedelta(seconds=300))):
            dadosGeralAba1[i]['Causa'] = "TentativaSucessiva"
            dadosGeralAba1[i - 4]['Causa'] = "TentativaSucessiva"
        elif((dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i - 5]['Cartao']) and (
                dadosGeralAba1[i]['Data'] < dadosGeralAba1[i - 5]['Data'] + timedelta(seconds=300))):
            dadosGeralAba1[i]['Causa'] = "TentativaSucessiva"
            dadosGeralAba1[i - 5]['Causa'] = "TentativaSucessiva"
        elif ((dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i - 6]['Cartao']) and (
                dadosGeralAba1[i]['Data'] < dadosGeralAba1[i - 6]['Data'] + timedelta(seconds=300))):
            dadosGeralAba1[i]['Causa'] = "TentativaSucessiva"
            dadosGeralAba1[i - 6]['Causa'] = "TentativaSucessiva"
        elif ((dadosGeralAba1[i]['Cartao'] == dadosGeralAba1[i + 1]['Cartao']) and (
                    dadosGeralAba1[i]['Data'] < dadosGeralAba1[i + 1]['Data'] + timedelta(seconds=300))):
            dadosGeralAba1[i]['Causa'] = "TentativaSucessiva"
            dadosGeralAba1[i+1]['Causa'] = "TentativaSucessiva"
        dadosGeralCall.append(dadosGeralAba1[i])


DadosValorSerie = pd.Series(dadosValor)
#Calculo de parametros
PercertualDechargeback = (len(dadosGeralCall)/len(dadosGeralAba1))*100
valorMaximo= max(dadosValor)
valorMinimo= min(dadosValor)
valorMedio= sum(dadosValor)/len(dadosValor)
mediana = DadosValorSerie.median()
desvio = DadosValorSerie.std()


listaDeClientes = set(dadosCartao)
quantidadeDeClientes = len(listaDeClientes)
#analise dos valores identificados
for j in range(0,len(dadosGeralCall)):
    if dadosGeralCall[j]['Causa']=="-":
        dadosNaopadronizado.append(dadosGeralCall[j])

PorcentagemPadronizada= 100-(len(dadosNaopadronizado)/len(dadosGeralCall))*100

print("sucesso")
