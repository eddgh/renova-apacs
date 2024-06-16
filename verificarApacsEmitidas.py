import openpyxl # Manipular planilhas
from conecta import *

link = "http://sia.datasus.gov.br/principal/index.php"
navegador.maximize_window()
navegador.get(url=link)

consulta = navegador.find_element(By.XPATH, value="/html/body/div[1]/div[2]/fieldset/div[2]/div/div[2]/a")
consulta.click()
apac = navegador.find_element(By.XPATH, value="/html/body/div[1]/div[3]/div/div[5]/a")
apac.click()

book = openpyxl.load_workbook('./ApacEmitidas.xlsx')
listaApacs = book['Plan1']

dados = []

for rows in listaApacs.iter_rows(min_row=2, max_row=3):
    numApac = rows[0].value # número da Apac Emitida
    
    campoApac = navegador.find_element(By.XPATH, value='/html/body/div[2]/div/div[2]/div/div/form/div/input')
    campoApac.send_keys(numApac)
    
    okButton = navegador.find_element(By.XPATH, value="/html/body/div[2]/div/div[2]/div/div/form/p/input")
    okButton.click()
    
    apacFind = navegador.find_element(By.XPATH, value='/html/body/div[2]/div/div[2]/div/table/tbody/tr[4]/td[1]')
    if(apacFind.text == ""):
        print(f'Apac: {numApac} - {rows[1].value} - Exclusiva')
    else:
        print(f'Apac: {numApac} - {rows[1].value} - Outra unidade')
        
    # apac = navegador.find_element(By.XPATH, value="/html/body/div[1]/div[3]/div/div[5]/a")
    # apac.click()
        
        
    
    
    
#     if len(data_em_texto) == 10:
#       data_em_texto = '{}{}{}'.format(inicio.day, inicio.month, inicio.year)
#     else:
#       data_em_texto = inicio.strftime('%d%m%Y')
    
