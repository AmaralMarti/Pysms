# -*- coding: utf-8 -*-
''' Integração Parcial do Sistema Smart Doctor para o Zenvia
    Autor: Thiago Oliveira Castro Vieira
    Licença GPL v3.0
'''

import re
import xlrd
import configparser
from urllib.request import Request, urlopen
import datetime
import json
from  pprint import pprint
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


config = configparser.ConfigParser()
config.read('config.txt')
authorization = config.get('configuration', 'authorization')
fromaddr = config.get('configuration', 'fromaddr')
toaddr = config.get('configuration', 'toaddr')
senhaemail= config.get('configuration', 'senhaemail')


url = 'https://api-rest.zenvia360.com.br/services/send-sms'
wb = xlrd.open_workbook('exemplo.xls', encoding_override="cp1252", ragged_rows=True) # enconding_override remove o erro de ausência de condificação em XLS antigos.
worksheet = wb.sheet_by_index(0)
headers = {
  "Content-Type": "application/json",
  "Authorization": authorization,
  "Accept": "application/json"
}
def getreplies (headers):
    hoje = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d')
    urlapi = "https://api-rest.zenvia360.com.br/services/received/search/%sT00:00:00/%sT23:59:59" % (hoje, hoje)
    request = Request(urlapi, headers=headers)
    replies = urlopen(request).read()
    replies = replies.decode('utf-8')
    return replies

def cleanreplies (replies):
    data_replies = json.loads(replies)
    try:
#        Caminho do telefone = data_replies['receivedResponse']['receivedMessages'][0]['mobile']
#        Caminho da Resposta = data_replies['receivedResponse']['receivedMessages'][0]['body']
#        for dados in data_replies['receivedResponse']['receivedMessages']:
        respostas = {dados['mobile']: dados['body'] for dados in data_replies['receivedResponse']['receivedMessages']}
        # print(dados['mobile'], dados['body'])
    except TypeError:
        print ('ninguém respondeu')
    return respostas

respostas = cleanreplies(getreplies(headers))
pprint(respostas)
email = []
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
    if celularbr in respostas.keys():
      email.append(nomecompleto + ' - ' + celularbr + ': ' + respostas[celularbr])
    n = n + 1
    if n >= worksheet.nrows:
      break

def sendmail(email):
    hoje = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d')
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = "[CLIMAE - SMS] Respostas  %s" % hoje
    body = """	Olá,
    Os seguintes pacientes responderam: %s

    Atenciosamente,

    Mr. Robot
    """ % email
    print(body)
    msg.attach(MIMEText(body, 'plain'))
    server = smtplib.SMTP('smtp.zoho.com', 587)
    server.starttls()
    server.login(fromaddr, senhaemail)
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    print("E-mail enviado com sucesso!")
    server.quit()

sendmail(email)
