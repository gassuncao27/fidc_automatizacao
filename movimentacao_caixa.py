# script para movimentar caixa aplicações Santander
# caminho para o arquivo no sistema + abrir o arquivo em excel

from dotenv import load_dotenv
from datetime import datetime

import openpyxl
import os
import sys

import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def tipooperacao_str(nr_operacao):
    operacao_caixa = {
        1: 'Aplicação',
        2: 'Resgate Parcial',
        5: 'Resgate Total',
    }
    operation_type = operacao_caixa.get(nr_operacao)
    return operation_type

data_atual = datetime.now()
data_formatada = data_atual.strftime('%d/%m/%y')
tipo_operacao = int(input("\nTipo de Operação: "))
valor_operacao = float(input("Valor da Operação: "))

if tipooperacao_str(tipo_operacao) is None:
    print('Tipo de Operação não existente. Favor Verificar!\n')
    sys.exit()

load_dotenv()
caminho_arquivo = os.getenv('CAMINHO_ARQUIVOCAIXA')
caminho_template = caminho_arquivo + '/modelo_santander.xlsx'
arquivo = openpyxl.load_workbook(caminho_template)


planilha = arquivo['NOVO MODELO SANTANDER']
planilha.cell(5, 2).value = tipo_operacao
planilha.cell(5, 6).value = valor_operacao
planilha.cell(5, 8).value = data_formatada


print('')
print(planilha.cell(5, 2).value)
print(planilha.cell(5, 6).value)
print(planilha.cell(5, 8).value)
print('')

caminho_novo_arquivo = caminho_arquivo+'/VanguardII_'+str(valor_operacao)+'.xlsx'
arquivo.save(caminho_novo_arquivo)

print(tipooperacao_str(tipo_operacao))


# smtp_server = 'smtp.gmail.com'
# smtp_port = 587
# remetente_email = 'gassuncao27@gmail.com'
# remetente_senha = 'prbodrwvszovzips'
# destinatarios = ['gabriel@cartor.com.br']
# msg = MIMEMultipart()
# msg['From'] = remetente_email
# msg['To'] = ', '.join(destinatarios)
# msg['Subject'] = tipooperacao_str(tipo_operacao)+' - Vanguard II - fundo Soberano Santander'

# nome_arquivo = caminho_novo_arquivo
# parte_anexada = MIMEBase('application', 'octet-stream')
# parte_anexada.set_payload(open(nome_arquivo, 'rb').read())
# encoders.encode_base64(parte_anexada)
# parte_anexada.add_header('Content-Disposition', 'attachment', filename=nome_arquivo.split('/')[-1])
# msg.attach(parte_anexada)

# server = smtplib.SMTP(smtp_server, smtp_port)
# server.starttls()
# server.login(remetente_email, remetente_senha)
# server.sendmail(remetente_email, destinatarios, msg.as_string())


# celulas mexer = 5,2 - 5,6 - 5,8
# modelo de data = 10/10/23
# modelo de numero = 203.000,00
# tipo de operação = 1:Aplicaçãp., 2:Resgate Parcial, 5: Resgate Total
