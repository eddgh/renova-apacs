import openpyxl
from functions import *

# O OBJETIVO DESTE CÓDIGO É PREPARAR O ARQUIVO QUE CONTERÁ AS APACS A SEREM SOLICITADAS
# JÁ COM OS FORMATOS NECESSÁRIOS EM CADA LINHA

# MANIPULANDO OS ARQUIVOS DO EXCEL
# Carregando arquivo com a lista de apacs para renovar
book = openpyxl.load_workbook('ApacsVencidas.xlsx')
# Selecionando uma página
pacientes_renova = book['Sheet']

for rows in pacientes_renova.iter_rows(min_row=2):
    cpf = rows[1].value # cpf na planilha de apacs vencidas
    nomeTrs = acharCpf(cpf) # Pega na Planilha 'PacientesEmTratamento v2' os nomes exatamente como estão no TRS
    rows[0].value = nomeTrs # Atribui esses nomes na planilha ApacsRenovar pra solicitar renovação pelo nome
    
    # Colocar a primeira letra em maiuscula e o restante em minuscula de cada parte do nome do medico
    rows[2].value = (rows[2]).value.title()
    
    # passando os valores de albumina(alb), hemoglobina(hb) e taxa de redução de ureia(urr), respectivamente:
    # values = [rows[4].value, rows[5].value, rows[6].value] 
    
    # formatando os valores multiplicando alb*10, hb*10 e urr*100:
    # rows[4].value = transformValues(values)[0] # alb
    # rows[5].value = transformValues(values)[1] # hb
    # rows[6].value = transformValues(values)[2] # urr      
    
book.save('ApacsRenovar.xlsx')