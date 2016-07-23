# -*- coding: utf-8 -*-

""" Integração Parcial do Sistema Smart Doctor para o Zenvia
    Autor: Thiago Oliveira Castro Vieira
    Licença GPL v3.0
"""

import re
import urllib.request
import xlrd
import configparser

config = configparser.ConfigParser()
config.read('config.txt')
authorization = config.get('configuration', 'authorization')


url = 'https://api-rest.zenvia360.com.br/services/send-sms'
wb = xlrd.open_workbook('exemplo.xls', encoding_override="cp1252", ragged_rows=True) # enconding_override remove o erro de ausência de condificação em XLS antigos.
worksheet = wb.sheet_by_index(0)

for linha in range(1, worksheet.nrows):
    # Ignora as linhas em que o nome do(a) paciente estiver em branco
    if worksheet.cell(linha, 0).value == xlrd.empty_cell.value:  # You can detect an empty cell by using empty_cell in xlrd.empty_cell.value
        continue

    # Captura o nome completo do(a) paciente na planilha.
    nome_paciente = worksheet.cell(linha, 0).value
    # Expressão Regular para isolar o primeiro nome. Utilizar nome.group(0) para eliminar <_sre.SRE_Match object at
    nome_paciente = re.search(r'([A-Z]*)\s', nome_paciente)
    nome_paciente = nome_paciente.group(0)

    numero_celular = worksheet.cell(linha, 2).value
    numero_celular = str(numero_celular)
    numero_celular = re.search(r'([0-9]*)', numero_celular)
    numero_celular = '55' + numero_celular.group(0)

    nome_medico = worksheet.cell(linha, 3).value
    data_hora = worksheet.cell(linha, 4).value

    msg = ('Sr(a) ' + nome_paciente + 'confirmamos sua consulta com o(a) médico(a) ' + nome_medico + ' no dia ' + data_hora + '. Responda gratis S para confirmar ou N para cancelar.')
    print (msg)

    values = '''
    {
        "sendSmsRequest": {
        "from" : "CLIMAE",
        "to":  "%s",
        "msg": "%s",
        "aggregateId": "1111"
        }
    }''' % (numero_celular, msg)

    headers = {
        "Content-Type": "application/json",
        "Authorization": 0, ##authorization,
        "Accept": "application/json"
    }

    values = values.encode('utf-8')
    request = urllib.request.Request(url, data=values, headers=headers)
    response_body = urllib.request.urlopen(request).read()
    print (response_body)


