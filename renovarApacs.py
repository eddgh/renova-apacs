import openpyxl # Manipular planilhas
import dotenv #
import os

dotenv.load_dotenv(dotenv.find_dotenv()) # carrega as variáveis do arquivo .env

from selenium import webdriver

#GeckoDriver = webdriver como um serviço (atualiza o geckodriver automaticamente)
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.service import Service

from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options 
from selenium.webdriver.common.keys import Keys
# from time import sleep

options = Options()
# options.add_argument('--headless') # executar o navegador de forma oculta

# Firefox
servico = Service(GeckoDriverManager().install())
navegador = webdriver.Firefox(service=servico, options=options)

link = "https://ser.saude.rj.gov.br/trs/login/login"

navegador.maximize_window()
navegador.get(url=link)

login = navegador.find_element(By.ID, value="login")
login.send_keys(os.getenv("LOGIN"))

senha = navegador.find_element(By.NAME, value="passwd")
senha.send_keys(os.getenv("SENHA"))

entrar = navegador.find_element(By.XPATH, value="/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[5]/td[3]/input")
entrar.click()

link2 = "https://ser.saude.rj.gov.br/trs/renovacaoAPAC/list" # página onde solicitar as renovações das apacs
navegador.get(url=link2)

campoNome = navegador.find_element(By.XPATH, value='//*[@id="nome"]')
btn_buscar = navegador.find_element(By.XPATH, value="/html/body/div[4]/form/table/tbody/tr[2]/td[2]/input[2]")

book = openpyxl.load_workbook('./ApacsRenovar.xlsx')
pacientes_renova = book['Sheet']

for rows in pacientes_renova.iter_rows(min_row=43):
    nome = rows[0].value # nome do paciente  
    medico = rows[2].value # medico responsavel
    inicio = rows[3].value # inicio na clinica
    alb = rows[4].value # albumina
    hb = rows[5].value # hemoglobina
    urr = rows[6].value # taxa de redução de uréia
    acesso = rows[7].value # acesso do paciente para dialisar
    
    print(nome)

    campoNome.send_keys(nome)
    
    btn_buscar.click()
    btn_editarPaciente = navegador.find_element(By.XPATH, value='/html/body/div[4]/div/table/tbody/tr/td[1]/span/a')
    btn_editarPaciente.click()
    


