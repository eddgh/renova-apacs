# inicio da conexao
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
options.add_argument('--headless') # executar o navegador de forma oculta

# Firefox
servico = Service(GeckoDriverManager().install())
navegador = webdriver.Firefox(service=servico, options=options)
# fim da conexao