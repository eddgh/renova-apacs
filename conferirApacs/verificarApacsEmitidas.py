from conectar import *

import pandas as pd
import sys
import openpyxl # Manipular planilhas

import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

link = "http://sia.datasus.gov.br/principal/index.php"
navegador.maximize_window()
navegador.get(url=link)

consulta = navegador.find_element(By.XPATH, value="/html/body/div[1]/div[2]/fieldset/div[2]/div/div[2]/a")
consulta.click()

book = openpyxl.load_workbook('ApacEmitidas.xlsx')
listaApacs = book['Plan1']

apac = []
paciente = []
msg = []
municipio = []


for rows in listaApacs.iter_rows(min_row=2):    
    navegador.get("http://sia.datasus.gov.br/remessa/ConsultaApac.php")    
    numApac = rows[0].value # número da Apac Emitida    
    campoApac = navegador.find_element(By.XPATH, value='/html/body/div[2]/div/div[2]/div/div/form/div/input')
    campoApac.send_keys(numApac)    
    okButton = navegador.find_element(By.XPATH, value="/html/body/div[2]/div/div[2]/div/div/form/p/input")
    okButton.click()    
    apacFind = navegador.find_element(By.XPATH, value='/html/body/div[2]/div/div[2]/div/table/tbody/tr[4]/td[1]')
    gestor = navegador.find_element(By.XPATH, value='/html/body/div[2]/div/div[2]/div/table/tbody/tr[4]/td[2]')
        
    if(apacFind.text.isnumeric()): 
        # print(f'Apac encontrada!!! : {apacFind.text} - {rows[1].value} - Gestor: {gestor.text} - Apac de Outra unidade')
        apac.append(apacFind.text)
        paciente.append(rows[1].value)
        msg.append('Apac de outra unidade!!!')
        municipio.append(gestor.text)      
    else:
        # print(f'Apac: {numApac} - {rows[1].value} - Gestor: {gestor.text} - Apac Exclusiva')
        apac.append(numApac)
        paciente.append(rows[1].value)
        msg.append('Apac Exclusiva')
        municipio.append(gestor.text)
        
resultado = {
    "APAC":apac,
    "PACIENTE":paciente,
    "MSG":msg,
    "GESTOR":municipio
    }

df = pd.DataFrame(resultado)
df.to_excel("ApacsVerificadas.xlsx", index=None, sheet_name="Sheet", )
    
