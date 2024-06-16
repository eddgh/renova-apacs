import openpyxl
from functions import *
import mapeaArquivo

# O OBJETIVO DESTE CÓDIGO É PREPARAR O ARQUIVO FINAL PRA RENOVAÇÃO
# PEGANDO OS NOMES REAIS DOS PACIENTES EXATAMENTE COMO ESTÃO NO TRS
# PORQUE A PARTIR DOS NOMES REAIS OUTRO CÓDIGO SOLICITARÁ AS RENOVAÇÕES (renovarApacs.py)

# MANIPULANDO OS ARQUIVOS DO EXCEL
# Carregando arquivo com a lista de apacs para renovar
book = openpyxl.load_workbook('ApacsRenovar.xlsx')
# Selecionando uma página
pacientes_renova = book['Sheet']

for rows in pacientes_renova.iter_rows(min_row=2):
    cpf = rows[1].value # cpf na planilha de apacs vencidas
    nomeTrs = acharCpf(cpf) # Pega na Planilha 'PacientesEmTratamento v2' os nomes exatamente como estão no TRS
    rows[0].value = nomeTrs # Atribui esses nomes na planilha ApacsRenovar pra solicitar renovação pelo nome
    
    # Colocar a primeira letra em maiuscula e o restante em minuscula de cada parte do nome do medico
    rows[2].value = (rows[2]).value.title()

book.save('ApacsRenovar.xlsx')