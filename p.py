from excel import Excel
from openpyxl import load_workbook
import json

planilha = load_workbook(filename='Digitação Inovação Completo com CNAE.xlsx')
planilha_banco_de_dados = load_workbook(filename='SP.xlsx')


digitacao = planilha['Digitação']
banco_de_dados = planilha_banco_de_dados['Arquivo']



tel1_digitacao = Excel(digitacao, 'e', 'e')
tel1_banco = Excel(banco_de_dados, 'l', 'l')


tel1_digitacao.busca_duplicatas(tel1_banco.lista)







        