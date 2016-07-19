# -*- coding: utf-8 -*-
''' Integração Parcial do Sistema Smart Doctor para o Zenvia
    Autor: Thiago Oliveira Castro Vieira
    Licença GPL v3.0
'''
import re
import urllib.request
import xlrd
import configparser
import requests

config = configparser.ConfigParser()
config.read('config.txt')
authorization = config.get('configuration', 'authorization')

url = 'https://api-rest.zenvia360.com.br/services/send-sms'
wb = xlrd.open_workbook('exemplo.xls', encoding_override="cp1252", ragged_rows=True) # enconding_override remove o erro de ausência de condificação em XLS antigos.
worksheet = wb.sheet_by_index(0)
n = 1
while worksheet.cell(n,0).value != xlrd.empty_cell.value: # You can detect an empty cell by using empty_cell in xlrd.empty_cell.value
    # Captura o nome completo da paciente na planilha.
    nomecompleto = worksheet.cell(n, 0).value
    # Expressão Regular para isolar o primeiro nome. Utilizar nome.group(0) para eliminar <_sre.SRE_Match object at
    nome = re.search(r'([A-Z]*)\s', nomecompleto)
    celular = worksheet.cell(n, 2).value
    celularfloat = str(celular)
    celular = re.search(r'([0-9]*)', celularfloat)
    celularbr = '55' + celular.group(0)
    msg = ('Senhora ' + nome.group(0) + 'a CLIMAE confirma sua consulta com o(a) médico(a) ' + worksheet.cell(n,3).value + ' no dia ' + worksheet.cell(n,4).value + '. Responda gratis S para confirmar ou N para cancelar.' + ';' + 'CLIMAE')
    print (celular.group(0) + ';' + 'Senhor(a) ' + nome.group(0) + 'a CLIMAE confirma sua consulta com o(a) médico(a) ' + worksheet.cell(n,3).value + ' no dia ' + worksheet.cell(n,4).value + '. Responda gratis S para confirmar ou N para cancelar. CLIMAE')
    values = """
    {
        "sendSmsRequest": {
        "from" : "CLIMAE",
        "to":  "%s",
        "schedule": "NONE",
        "msg": "%s",
        "callBackOption": "NONE",
        "id": "002",
        "aggregateId": "1111"
        }
    }""" % (celularbr, msg)
    headers = {
        "Content-Type": "application/json",
        "Authorization": authorization,
        "Accept": "application/json"
    }
    values = values.encode('utf-8')
#    request = urllib.request.Request("https://api-rest.zenvia360.com.br/services/send-sms", data=values, headers=headers)
#    response_body = urllib.request.urlopen(request).read()
    session_requests = requests.session()

    result = session_requests.post(
        url,
        data=values,
        headers=dict(referer=url)
    )
    n = n + 1
    if n >= worksheet.nrows:
        break



