import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
from selenium.webdriver.remote.webelement import WebElement

driver = webdriver.Chrome()

dataBaseID = {
 "PI":"Packaging Integrity",
 "ST":"Sealing Technology",
 "ACCP":"Automated Control for Critical Parameters",
 "VSI":"Vision System Inspection",
 "SI":"Seal Integrity",
 "CV":"Component Verification",
 "ELV":"Electronic Label Verification",
 "EVPL":"Electronic Verification of printed labels.",
 "HD":"Human Dependency",
 "AHD":"Assessment of Human Dependancy "
}

choose = {
   "Yes":"2",
   "No":"3",
   "":"4"
}

choose2 = {
   "Yes" :"3",
   "No" :"2",
    "":"4"
}

'''dataBaseNAME ={
   "CycleRangeTemp":"answer_id7249" ,
   "AlarmRangeTemp" : "answer_id6894",
   "RejectsImpacTemp":"answer_id6895",
   "CycleRangePress": "answer_id6896",
   "AlarmRangePress" : "answer_id6897",
   "RejectsImpacPress": "answer_id6898",
   "MonitoringPressure":"answer_id6899",
   "AlarmDwell" : "answer_id6900",
   "RejectsImpacDwell":"answer_id6901",
   "RejectsInterruption":"answer_id6902",
   "RedudantControls":"answer_id6903",
   "Channels/Wrinkles":"answer_id6904"
}'''
nome = 'PLANILHA1.xlsx' #colocar o nome da tabela aqui

def get_element_by_name(number):
   element = driver.find_element(By.NAME,f"answer_id{number}")
   print(element)
   return element

def get_element_by_td(element_name: str) -> WebElement:
   elem = driver.find_element(By.XPATH, f"//tr/td[text()='{element_name}']")
   return elem

def get_element_by_id(id:str) -> WebElement:
   check = driver.find_element(By.XPATH, f"//button/div/p[text()='{id}']")
   return check

def run(minha_lista,istart,iend,j):
   for i in range(istart,iend):
      if(i ==7249 or i==7252):
         j+=1
         continue
      print(i)
      select = Select(get_element_by_name(i))
      print(j)
      select.select_by_value(choose.get(minha_lista[j]))
      j+=1
   
   
url = 'https://yellow-pond-023165e0f.4.azurestaticapps.net/equipment/' #colocar o link do site aqui
driver.get(url)
time.sleep(3) #dar um tempo pra pagina carregar
driver.maximize_window()
time.sleep(3)

continueEmail = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/button')
continueEmail.click()
time.sleep(5)
window = driver.window_handles
driver.switch_to.window(window[1])
time.sleep(10)
login = driver.find_element(By.NAME,'loginfmt')
login.send_keys('FFrancisco8@its.jnj.com')
time.sleep(2)

avancar = driver.find_element(By.XPATH,'//*[@id="idSIButton9"]')
avancar.click()
time.sleep(10)

driver.switch_to.window(window[0])

teste = driver.find_element(By.XPATH,'//*[@id="__next"]/div/div/div[1]/div/button')
teste.click()
time.sleep(4)

action_chains = ActionChains(driver)
action_chains.move_to_element_with_offset(teste, 0, 50).click().perform()

time.sleep(5)
equipment = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/div/div[2]/a[2]')
equipment.click()
time.sleep(20)

tabela = 'Questions'

dados = pd.read_excel(nome, sheet_name=tabela, na_values=[''])
num_minha_listas, num_colunas = dados.shape
indice_inicio = 3


if 0 <= indice_inicio < num_colunas:
  for indice_minha_lista in range(indice_inicio,num_colunas):
       lista = dados[indice_minha_lista].tolist()
       minha_lista = ['' if pd.isna(elemento) else elemento for elemento in lista] #aqui verelifica se o elemento é nulo, se for ele coloca ''
       print(minha_lista)
       if(minha_lista[0]==''):break

       find = get_element_by_td(minha_lista[0])
       find.click()
       time.sleep(5)
       questions = driver.find_element(By.XPATH, "//button[text()='Questions']")
       questions.click()
       
       PI = get_element_by_id(dataBaseID.get("PI"))
       PI.click()
       time.sleep(2)
       ST = get_element_by_id(dataBaseID.get("ST"))
       ST.click()
       time.sleep(2)
       ACCP = get_element_by_id(dataBaseID.get("ACCP"))
       ACCP.click()
       time.sleep(5)
       number = 7249
       number2 = number + 11 #subiu 11
       number3 = number2 + 13 #subiu 13
       number4 = number3 + 6 #subiu 6
       number5 = number4 + 11 #subiu 11
       j = 1
       run(minha_lista,number,number2-1,j)
       time.sleep(5)
       ERROR1 = Select(get_element_by_name(7249))
       ERROR1.select_by_value(choose2.get(minha_lista[1]))
       time.sleep(5)
       ERROR2 = Select(get_element_by_name(7252))
       ERROR2.select_by_value(choose2.get(minha_lista[4]))
       time.sleep(5)

       VSI = get_element_by_id(dataBaseID.get("VSI"))
       VSI.click()
       time.sleep(2)
       SI = get_element_by_id(dataBaseID.get("SI"))
       SI.click()
       j = 12
       #for i in range(number2,number3-1):
       run(minha_lista,number2,number3-1,j)
       time.sleep(5)
       CV = get_element_by_id(dataBaseID.get("CV"))
       CV.click()
       time.sleep(2)
       j = 25
       #for i in range(number3,number4-1):
       run(minha_lista,number3,number4-1,j)

       ELV = get_element_by_id(dataBaseID.get("ELV"))
       ELV.click()
       time.sleep(2)
       EVPL = get_element_by_id(dataBaseID.get("EVPL"))
       EVPL.click()
       j = 31
       #for i in range(number4,number5-1):
       run(minha_lista,number4,number5-1,j)
       time.sleep(5)
       HD = get_element_by_id(dataBaseID.get("HD"))
       HD.click()
       time.sleep(2)
       AHD = get_element_by_id(dataBaseID.get("AHD"))
       AHD.click()
       j=42
       #for i in range(number5,6938):
       run(minha_lista,number5,6938,j)
       time.sleep(500)




driver.quit()
     