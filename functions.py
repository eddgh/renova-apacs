import openpyxl
import math
from decimal import Decimal

# Carregando arquivo-base para pegar os nomes dos pacientes exatamente como estao no TRS
book2 = openpyxl.load_workbook('../pacientes_tratamento/PacientesEmTratamento v2.xlsx')
# Selecionando uma página
pacientes_page = book2['Plan1']

# função para pegar o nome do paciente de acordo com o CPF
def acharCpf(cpf): 
    for rows in pacientes_page.iter_rows(min_row=2):
        if cpf == rows[3].value:
            paciente = rows[1].value
            return paciente 