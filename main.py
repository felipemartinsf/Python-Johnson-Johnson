import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time


nome = 'PLANILHAS1.xlsx' #colocar o nome da tabela aqui
driver = webdriver.Chrome()
url = 'https://yellow-pond-023165e0f.4.azurestaticapps.net/equipment/' #colocar o link do site aqui
driver.get(url)
time.sleep(3) #dar um tempo pra pagina carregar
continueEmail = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/button')
continueEmail.click()
time.sleep(4)
window = driver.window_handles
driver.switch_to.window(window[1])
login = driver.find_element(By.XPATH,'//*[@id="i0116"]')
login.send_keys('FFrancisco8@its.jnj.com')
time.sleep(2)
avancar = driver.find_element(By.XPATH,'//*[@id="idSIButton9"]')
#//*[@id="__next"]/div/div/div[1]/div/button
avancar.click()
time.sleep(7)
driver.switch_to.window(window[0])
teste = driver.find_element(By.XPATH,'//*[@id="__next"]/div/div/div[1]/div/button')
teste.click()
time.sleep(4)
action_chains = ActionChains(driver)
action_chains.move_to_element_with_offset(teste, 0, 50).click().perform()
time.sleep(4)
equipment = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/div/div[2]/a[2]')
equipment.click()
time.sleep(4)

franchise = driver.find_element(By.XPATH,'//*[@id="field-:r32:"]')
newMachine = driver.find_element(By.XPATH,'//*[@id="__next"]/div/div/main/header/button')

workbook = openpyxl.load_workbook(nome)
tabela = 'Summary'

dados = pd.read_excel(nome, sheet_name=tabela, na_values=[''])
num_linhas, num_colunas = dados.shape


indice_inicio = 16 
i = 0 
if 0 <= indice_inicio < num_linhas:
  for indice_linha in range(indice_inicio, num_linhas):
       linha = dados.iloc[indice_linha].tolist()
       minha_lista = ['' if pd.isna(elemento) else elemento for elemento in linha] #aqui verifica se o elemento é nulo, se for ele coloca ''
       newMachine.click()
       franchise.send_keys(minha_lista[i])
            

#else:
 #   print(f"O índice de início {indice_inicio} está fora dos limites do DataFrame.")