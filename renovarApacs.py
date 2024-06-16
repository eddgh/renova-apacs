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
from selenium.webdriver.support.ui import Select
from datetime import datetime
from time import sleep

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

book = openpyxl.load_workbook('./ApacsRenovar.xlsx')
pacientes_renova = book['Sheet']

for rows in pacientes_renova.iter_rows(min_row=2):
    nome = rows[0].value # nome do paciente  
    medico = rows[2].value # medico responsavel
    inicio = rows[3].value # inicio na clinica
    data_em_texto = '{}{}{}'.format(inicio.day, inicio.month, inicio.year)
    
    if len(data_em_texto) == 10:
      data_em_texto = '{}{}{}'.format(inicio.day, inicio.month, inicio.year)
    else:
      data_em_texto = inicio.strftime('%d%m%Y')
    
    alb = rows[4].value # albumina
    hb = rows[5].value # hemoglobina
    urr = rows[6].value # taxa de redução de uréia
    acessoPaciente = rows[7].value # acesso do paciente para dialisar
    
    # acessar a página de solicitação de apac por paciente
    link2 = "https://ser.saude.rj.gov.br/trs/renovacaoAPAC/list" # página onde solicitar as renovações das apacs
    navegador.get(url=link2)
    campoNome = navegador.find_element(By.XPATH, value='//*[@id="nome"]')
    btnBuscar = navegador.find_element(By.XPATH, value="/html/body/div[4]/form/table/tbody/tr[2]/td[2]/input[2]")   
    campoNome.send_keys(nome)
    btnBuscar.click()
    btnEditarPaciente = navegador.find_element(By.XPATH, value='/html/body/div[4]/div/table/tbody/tr/td[1]/span/a')
    btnEditarPaciente.click()
   
    selectMedico = navegador.find_element(By.XPATH, value='//*[@id="medico.id"]')
    selectSessoes = navegador.find_element(By.XPATH, value='/html/body/div[4]/form/table[2]/tbody/tr[2]/td[2]/input[2]')
    primeiraDialise = navegador.find_element(By.XPATH, value='//*[@id="dtPrimDialise"]')
    albumina = navegador.find_element(By.XPATH, value='//*[@id="albumina"]')
    hemoglobina = navegador.find_element(By.XPATH, value='//*[@id="hemoglobina"]')
    tru = navegador.find_element(By.XPATH, value='//*[@id="TRU"]')
    acesso = navegador.find_element(By.XPATH, value='/html/body/div[4]/form/table[2]/tbody/tr[7]/td[2]/input[1]')
    btnSalvar = navegador.find_element(By.XPATH, value='/html/body/div[4]/form/div/span/input')
        
    # Acessos
    cdl = navegador.find_element(By.XPATH, value='/html/body/div[4]/form/table[2]/tbody/tr[7]/td[2]/input[1]')
    tenckhoff = navegador.find_element(By.XPATH, value='/html/body/div[4]/form/table[2]/tbody/tr[7]/td[2]/input[2]')
    fav = navegador.find_element(By.XPATH, value='/html/body/div[4]/form/table[2]/tbody/tr[7]/td[2]/input[3]')
    permicath = navegador.find_element(By.XPATH, value='/html/body/div[4]/form/table[2]/tbody/tr[7]/td[2]/input[4]')
    
    # selecionar um médico (medico responsavel)
    drop = Select(selectMedico)
    drop.select_by_visible_text(medico)
    # selecionar quantas sessoes (3 sessoes sempre) <= sessoes de hemodialise por semana
    selectSessoes.click()
    # Primeira Diálise na Unidade
    primeiraDialise.click()
    primeiraDialise.clear()
    primeiraDialise.send_keys(data_em_texto)
    
    # 'digitando' albumina, hemoglobina e urr
    albumina.send_keys(alb)
    hemoglobina.send_keys(hb)
    tru.send_keys(urr)
    
    # desmarcar todos os chekcboxes de opções de acesso do paciente    
    if cdl.is_selected():
        cdl.click()
    if  tenckhoff.is_selected():
        tenckhoff.click()
    if  fav.is_selected():
        fav.click()
    if  permicath.is_selected():
        permicath.click()
          
    # marcar o acesso do paciente	
    if  "DUPLO" in acessoPaciente:
        cdl.click()       
    if  "TENCKHOFF" in acessoPaciente:
        tenckhoff.click()
    if  "FAV" in acessoPaciente:
        fav.click()
    if  "LONGA" in acessoPaciente:
        permicath.click()

    # rolar a tela até o botão "Salvar" e submeter a solicitação
    body = navegador.find_element(By.XPATH, value='/html/body')
    body.send_keys(Keys.ARROW_DOWN)            
    body.send_keys(Keys.ARROW_DOWN)   
    btnSalvar.click()