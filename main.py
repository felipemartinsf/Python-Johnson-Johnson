import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import time

from selenium.webdriver.remote.webelement import WebElement


nome = 'PLANILHA1.xlsx' #colocar o nome da tabela aqui
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
time.sleep(5)




#site = SEMPRE BRAZIL
#area //*[@id="field-:re:"] 
# EB //*[@id="field-:rf:"]
#EM //*[@id="field-:rg:"]
#en = driver.find_element(By.XPATH,'//*[@id="field-:rh:"]')
#PL //*[@id="field-:ri:"]
# QD //*[@id="field-:rj:"]
#YS //*[@id="field-:rk:"]
#TMACHINE //*[@id="field-:rl:"] SEMPRE CUSTOM
#TEQUI //*[@id="field-:rm:"]
#TMATERIAL//*[@id="field-:rn:"]
#loading //*[@id="field-:ro:"]
#offloading //*[@id="field-:rp:"]
#SV //*[@id="field-:rs:"]
#TSV //*[@id="field-:rt:"]
# PG //*[@id="field-:rt:"]
# SVP //*[@id="field-:rt:"] aqui eh porcentagem
# AVGL //*[@id="field-:r13:"]
#AVGEY //*[@id="field-:r13:"]
#AVCM //*[@id="field-:r16:"]
#AEVALIDATED //*[@id="field-:r14:"]
#AECURRENTLY //*[@id="field-:r15:"]
# NPC //*[@id="field-:r17:"]
#SQRM //*[@id="field-:r18:"]
#op //*[@id="field-:r131:"]
#MC //*[@id="field-:r1a:"]

'''
    


           if(minha_lista[4]==''):
           EM.select_by_value("1")
       elif(minha_lista[4]=='R245'):
           EM.select_by_value("2")
       elif(minha_lista[4]=='OM-L'):
           EM.select_by_value("3")
'''

elems = {
   "NewMachine" :"newMachine_id",
   "Franchise" : "franchise_id",
   "Area": "area_id",
   "Site": "site_id",
   "EB" : "eb_id",
   "EM" : "em_id"
}

def get_element_by_name(element_name: str) -> WebElement:
   required_element = elems.get(element_name)
   elem = driver.find_element(By.NAME, required_element)
   return elem

def searchArea(name:str,area:WebElement):
   
   match name:
      case "Endo":
       area.select_by_value("2")
      case "Mesh":
       area.select_by_value("3")
      case "Print Shop":
       area.select_by_value("4")
      case "Catgut":
       area.select_by_value("5")
      case "EO":
       area.select_by_value("6")
      case "Packaging":
       area.select_by_value("7")
      case "Other":
       area.select_by_value("9")
        

def searchEB(name:str,EB:WebElement):
   match name:
      case "Haramura":
       EB.select_by_value("1")
      case "JNJ":
       EB.select_by_value("2")
      case "VSM":
       EB.select_by_value("3")
      case "AS2":
       EB.select_by_value("4")
      case "NTEC":
       EB.select_by_value("5") 
      case "MGS+Multivac":
       EB.select_by_value("6")
      case "Bodolay":
       EB.select_by_value("7")
      case "SELOVAC":
       EB.select_by_value("8")
      case "AGV":
       EB.select_by_value("9")
      case "Feinmechanik":
       EB.select_by_value("10")
      case "GPACK":
       EB.select_by_value("11")

def searchEM(name:str,EM):
  match name:
    case "":
      EM.select_by_value("1")
    case "R245":
      EM.select_by_value("2")
    case "OM-L":
      EM.select_by_value("3")



workbook = openpyxl.load_workbook(nome)
tabela = 'Equipment'

dados = pd.read_excel(nome, sheet_name=tabela, na_values=[''])
num_minha_listas, num_colunas = dados.shape


indice_inicio = 16
i = 0 
if 0 <= indice_inicio < num_minha_listas:
  for indice_minha_lista in range(indice_inicio, num_minha_listas):
       
       lista = dados.iloc[indice_minha_lista].tolist()
       minha_lista = ['' if pd.isna(elemento) else elemento for elemento in lista] #aqui verelifica se o elemento é nulo, se for ele coloca ''
       print(minha_lista)

       newMachine = get_element_by_name("NewMachine")

       franchise = get_element_by_name("Franchise")
       franchise.send_keys(minha_lista[0])

       site = Select(get_element_by_name("Site"))
       site.select_by_value("2")

       AR = Select(get_element_by_name("Area"))
       searchArea(minha_lista[2],AR)

                    
       EB = Select(get_element_by_name("EB"))
       searchEB(minha_lista[3],EB)
       EM = Select(get_element_by_name("EM"))
       searchEM(minha_lista[4],EM)
       
       EN = driver.find_element(By.XPATH,'//*[@id="field-:r12:"]') 
       EN.send_keys(minha_lista[5])
       PL = driver.find_element(By.XPATH,'//*[@id="field-:r13:"]') 
       PL.send_keys(minha_lista[6])
       QD = driver.find_element(By.XPATH, '//*[@id="field-:r14:"]')
       QD.send_keys(minha_lista[7])
       YS = driver.find_element(By.XPATH,'//*[@id="field-:r15:"]')
       YS.send_keys(minha_lista[8])
       TM = Select(driver.find_element(By.XPATH,'//*[@id="field-:r16:"]'))
       TM.select_by_value("2")
       TE = Select(driver.find_element(By.XPATH,'//*[@id="field-:r17:"]'))
       if(minha_lista[10]=='First Sealing'):
           TE.select_by_value("1")
       elif(minha_lista[10]=='Forming'):
           TE.select_by_value("2")
       elif(minha_lista[10]=='Fill/First Sealing'):
           TE.select_by_value("3")
       elif(minha_lista[10]=='Printer'):
           TE.select_by_value("4")
       elif(minha_lista[10]=='Form/Fill/Firt Sealing'):
           TE.select_by_value("5")
       elif(minha_lista[10]=='Other'):
           TE.select_by_value("6")
       elif(minha_lista[10]=='Form/Fill/Firt Sealing/Cutting'):
           TE.select_by_value("7")
       elif(minha_lista[10]=='Form/Fill/Firts Seal/Cutter/Cartoning'):
           TE.select_by_value("8")
       elif(minha_lista[10]=='Secondary Sealing'):
           TE.select_by_value("13")
       elif(minha_lista[10]=='Blanking'):
           TE.select_by_value("13")
       elif(minha_lista[10]=='Blanking/Cartoning'):
           TE.select_by_value("13")
       elif(minha_lista[10]=='Rotary Cutting'):
           TE.select_by_value("13")
       elif(minha_lista[10]=='Cartoning'):
           TE.select_by_value("13")
       elif(minha_lista[10]=='Dust Wrapping'):
           TE.select_by_value("14")
       TMAT = Select(driver.find_element(By.XPATH,'//*[@id="field-:r18:"]'))
       if(minha_lista[11]==''):
           TMAT.select_by_value("1")
       elif(minha_lista[11]=='Overwrap'):
           TMAT.select_by_value("2")
       elif(minha_lista[11]=='Foil'):
           TMAT.select_by_value("3")
       elif(minha_lista[11]=='Tyvek / Copolymer'):
           TMAT.select_by_value("4")  
       LOAD = Select(driver.find_element(By.XPATH,'//*[@id="field-:r19:"]'))
       if(minha_lista[12]=='Manual'):
           LOAD.select_by_value("1")
       elif(minha_lista[12]=='Automatic'):
           LOAD.select_by_value("2")
       OFFLOAD = Select(driver.find_element(By.XPATH,'//*[@id="field-:r1a:"]'))
       if(minha_lista[13]=='Manual'):
           OFFLOAD.select_by_value("1")
       elif(minha_lista[13]=='Automatic'):
           OFFLOAD.select_by_value("2")
       SV = driver.find_element(By.XPATH, '//*[@id="field-:r1d:"]')
       SV.send_keys(int(minha_lista[18]))
       TVS = driver.find_element(By.XPATH, '//*[@id="field-:r1e:"]')
       TVS.send_keys('83827629')
       PC = Select(driver.find_element(By.XPATH,'//*[@id="field-:r1f:"]'))
       if(minha_lista[20]=='PI'):
           PC.select_by_value("PI")
       elif(minha_lista[20]=='MP'):
           PC.select_by_value("MP")
       elif(minha_lista[20]=='PI/MP'):
           PC.select_by_value("MP/PI")
       TOTALCY = driver.find_element(By.XPATH,'//*[@id="field-:r1h:"]')
       if(minha_lista[22]!=''):
           TOTALCY.send_keys(minha_lista[21])
      #AVGLY //*[@id="field-:r13:"]
       AVGS= driver.find_element(By.XPATH, '//*[@id="field-:r1i:"]')
       AVGS.send_keys(int(round(minha_lista[23])))
       #aey //*[@id="field-:r13:"]
       AEY = driver.find_element(By.XPATH,'//*[@id="field-:r1k:"]')
       if(minha_lista[24]!=''):
        AEY.send_keys(100*(minha_lista[24]))
       AVGEV = driver.find_element(By.XPATH, '//*[@id="field-:r1l:"]')
       if(minha_lista[25]!=''):
         AVGEV.send_keys(minha_lista[25])
       AVGEC = driver.find_element(By.XPATH, '//*[@id="field-:r1m:"]')
       if(minha_lista[26]!=''):
         AVGEC.send_keys(100*(minha_lista[26]))
       AVGCM = driver.find_element(By.XPATH, '//*[@id="field-:r1n:"]')
       if(minha_lista[27]!=''):
         AVGCM.send_keys(minha_lista[27])
       NPC = driver.find_element(By.XPATH, '//*[@id="field-:r1o:"]')
       if(minha_lista[28]!=''):
         NPC.send_keys(int(round(minha_lista[28])))
       SQRO = driver.find_element(By.XPATH,'//*[@id="field-:r1p:"]')
       SQRO.send_keys(minha_lista[32])
       SHIFTO = driver.find_element(By.XPATH,'//*[@id="field-:r1q:"]')
       SHIFTO.send_keys(minha_lista[33])
       MC = driver.find_element(By.XPATH,'//*[@id="field-:r1r:"]')
       MC.send_keys(minha_lista[34])
       OPTS = driver.find_element(By.XPATH, '//*[@id="field-:r20:"]') #round(numero*130, 2)
       OPTS.send_keys(minha_lista[16])
       OPTD = driver.find_element(By.XPATH, '//*[@id="field-:r21:"]')
       OPTD.send_keys(minha_lista[17])

       time.sleep(1000)
       FBPTEMPMIN = driver.find_element(By.XPATH, '//*[@id="field-:r24:"]')
       if(minha_lista[35]!=''):
         FBPTEMPMIN.send_keys(int(round(minha_lista[35])))
       FBPTEMPSET = driver.find_element(By.XPATH, '//*[@id="field-:r25:"]')
       if(minha_lista[36]!=''):
         FBPTEMPSET.send_keys(int(round(minha_lista[36])))
       FBPTEMPMAX = driver.find_element(By.XPATH, '//*[@id="field-:r26:"]')    
       if(minha_lista[37]!=''):
         FBPTEMPMAX.send_keys(int(round(minha_lista[37])))
       FBPPRESSUREMIN = driver.find_element(By.XPATH, '//*[@id="field-:r27:"]')
       if(minha_lista[38]!=''):
         FBPPRESSUREMIN.send_keys(int(round(minha_lista[38])))
       FBPPRESSURESET = driver.find_element(By.XPATH, '//*[@id="field-:r28:"]')
       if(minha_lista[39]!=''):
         FBPPRESSURESET.send_keys(int(round(minha_lista[39])))
       FBPPRESSUREMAX = driver.find_element(By.XPATH, '//*[@id="field-:r29:"]')    
       if(minha_lista[40]!=''):
         FBPPRESSUREMAX.send_keys(int(round(minha_lista[40])))
       FBPTIMEMIN = driver.find_element(By.XPATH, '//*[@id="field-:r2a:"]')  
       if(minha_lista[41]!=''):
         FBPTIMEMIN.send_keys(int(round(minha_lista[41])))
       FBPTIMESET = driver.find_element(By.XPATH, '//*[@id="field-:r2b:"]')
       if(minha_lista[42]!=''):
         FBPTIMESET.send_keys(int(round(minha_lista[42])))
       FBPTIMEMAX = driver.find_element(By.XPATH, '//*[@id="field-:r2c:"]')    
       if(minha_lista[43]!=''):
         FBPTIMEMAX.send_keys(int(round(minha_lista[43])))
     


       FCPTEMPMIN = driver.find_element(By.XPATH, '//*[@id="field-:r2d:"]')
       if(minha_lista[44]!=''):
         FCPTEMPMIN.send_keys(int(round(minha_lista[44])))
       FCPTEMPSET = driver.find_element(By.XPATH, '//*[@id="field-:r2e:"]')
       if(minha_lista[45]!=''):
         FCPTEMPSET.send_keys(int(round(minha_lista[45])))
       FCPTEMPMAX = driver.find_element(By.XPATH, '//*[@id="field-:r2f:"]')    
       if(minha_lista[46]!=''):
         FCPTEMPMAX.send_keys(int(round(minha_lista[46])))
       FCPPRESSUREMIN = driver.find_element(By.XPATH, '//*[@id="field-:r2g:"]')
       if(minha_lista[47]!=''):
         FCPPRESSUREMIN.send_keys(int(round(minha_lista[47])))
       FCPPRESSURESET = driver.find_element(By.XPATH, '//*[@id="field-:r2h:"]')
       if(minha_lista[48]!=''):
         FCPPRESSURESET.send_keys(int(round(minha_lista[48])))
       FCPPRESSUREMAX = driver.find_element(By.XPATH, '//*[@id="field-:r2i:"]')    
       if(minha_lista[49]!=''):
         FCPPRESSUREMAX.send_keys(int(round(minha_lista[49])))
       FCPTIMEMIN = driver.find_element(By.XPATH, '//*[@id="field-:r2j:"]')  
       if(minha_lista[50]!=''):
         FCPTIMEMIN.send_keys(int(round(minha_lista[50])))
       FCPTIMESET = driver.find_element(By.XPATH, '//*[@id="field-:r2k:"]')
       if(minha_lista[51]!=''):
         FCPTIMESET.send_keys(int(round(minha_lista[51])))
       FCPTIMEMAX = driver.find_element(By.XPATH, '//*[@id="field-:r2l:"]')    
       if(minha_lista[52]!=''):
         FCPTIMEMAX.send_keys(int(round(minha_lista[52])))


       SPTEMPMIN = driver.find_element(By.XPATH, '//*[@id="field-:r2m:"]')
       if(minha_lista[53]!=''):
         SPTEMPMIN.send_keys(int(round(minha_lista[53])))
       SPTEMPSET = driver.find_element(By.XPATH, '//*[@id="field-:r2n:"]')
       if(minha_lista[54]!=''):
         SPTEMPSET.send_keys(int(round(minha_lista[54])))
       SPTEMPMAX = driver.find_element(By.XPATH, '//*[@id="field-:r2o:"]')    
       if(minha_lista[55]!=''):
         SPTEMPMAX.send_keys(int(round(minha_lista[55])))
       SPPRESSUREMIN = driver.find_element(By.XPATH, '//*[@id="field-:r2p:"]')
       if(minha_lista[56]!=''):
         SPPRESSUREMIN.send_keys(int(round(minha_lista[56])))
       SPPRESSURESET = driver.find_element(By.XPATH, '//*[@id="field-:r2q:"]')
       if(minha_lista[57]!=''):
         SPPRESSURESET.send_keys(int(round(minha_lista[57])))
       SPPRESSUREMAX = driver.find_element(By.XPATH, '//*[@id="field-:r2r:"]')    
       if(minha_lista[58]!=''):
         SPPRESSUREMAX.send_keys(int(round(minha_lista[58])))
       SPTIMEMIN = driver.find_element(By.XPATH, '//*[@id="field-:r2s:"]')  
       if(minha_lista[59]!=''):
         SPTIMEMIN.send_keys(int(round(minha_lista[59])))
       SPTIMESET = driver.find_element(By.XPATH, '//*[@id="field-:r2t:"]')
       if(minha_lista[60]!=''):
         SPTIMESET.send_keys(int(round(minha_lista[60])))
       SPTIMEMAX = driver.find_element(By.XPATH, '//*[@id="field-:r2u:"]')    
       if(minha_lista[61]!=''):
         SPTIMEMAX.send_keys(int(round(minha_lista[61])))
    
       time.sleep(200)

       #sqrO //*[@id="field-:r18:"]
       #shift op //*[@id="field-:r19:"]
       #Machine cost //*[@id="field-:r1a:"]
       # CYCLE //*[@id="field-:r17:"]
       # AVERAGE EQUIPMENT OEE VALIDATED //*[@id="field-:r14:"]
       # average equipment oee cucrrently //*[@id="field-:r15:"]
       # average cylcles per minute //*[@id="field-:r16:"]





#15 E 16
       time.sleep(300)    
#else:
 #   print(f"O índice de início {indice_inicio} está fora dos limites do DataFrame.")