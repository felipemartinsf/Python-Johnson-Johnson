import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import time
from selenium.webdriver.remote.webelement import WebElement

elems = {
   "NewMachine" :"newMachine_id",
   "Franchise" : "franchise_id",
   "Area": "area_id",
   "Site": "site_id",
   "EB" : "eb_id",
   "EM" : "em_id",
   "EN" : "en_id",
   "PL" : "pl_id",
   "QD" : "qd_id",
   "YS" : "ys_id",
   "TM" : "tm_id",
   "TE" : "te_id",
   "TMAT" : "tmat_id",
   "Load" : "load_id",
   "OffLoad" : "offload_id",
   "SV" :  "sv_id",
   "TVS" : "tvs_id",
   "PC" : "pc_id",
   "TOTALCY" : "totalcy_id",
    "AVGS": "avgs_id",
    "AEY": "aey_id",
    "AVGEV": "avgev_id",
    "AVGEC": "avgc_id",
    "AVGCM": "avgcm_id",
    "NPC": "npc_id",
    "SQRO": "sqro_id",
    "SHIFTO": "shifto_id",
    "MC": "mc_id",
    "OPTS": "opts_id",
    "OPTD": "optd_id",
     "FBPTEMPMIN": "24",
    "FBPTEMPSET": "25",
    "FBPTEMPMAX": "26",
    "FBPPRESSUREMIN": "27",
    "FBPPRESSURESET": "28",
    "FBPPRESSUREMAX": "29",
    "FBPTIMEMIN": "2a",
    "FBPTIMESET": "2b",
    "FBPTIMEMAX": "2c",
    "FCPTEMPMIN": "2d",
    "FCPTEMPSET": "2e",
    "FCPTEMPMAX": "2f",
    "FCPPRESSUREMIN": "2g",
    "FCPPRESSURESET": "2h",
    "FCPPRESSUREMAX": "2i",
    "FCPTIMEMIN": "2j",
    "FCPTIMESET": "2k",
    "FCPTIMEMAX": "2l",
    "SPTEMPMIN": "2m",
    "SPTEMPSET": "2n",
    "SPTEMPMAX": "2o",
    "SPPRESSUREMIN": "2p",
    "SPPRESSURESET": "2q",
    "SPPRESSUREMAX": "2r",
    "SPTIMEMIN": "2s",
    "SPTIMESET": "2t",
    "SPTIMEMAX": "2u",

}

searchArea = {
        "Endo": "2",
        "Mesh": "3",
        "Print Shop": "4",
        "Catgut": "5",
        "EO": "6",
        "Packaging": "7",
        "Other": "9"
}

searchTE= {
    'First Sealing': "1",
    'Forming': "2",
    'Fill/First Sealing': "3",
    'Printer': "4",
    'Form/Fill/First Sealing': "5",
    'Other': "6",
    'Form/Fill/First Sealing/Cutting': "7",
    'Form/Fill/First Sealing/Cutting': "8",
    'Secondary Sealing': "9",
    'Blanking': "10",
    'Blanking/Cartoning': "11",
    'Rotary Cutting': "12",
    'Cartoning': "13",
    'Dust Wrapping': "14"
}
searchEB = {
        "Haramura": "1",
        "JNJ": "2",
        "VSM": "3",
        "AS2": "4",
        "NTEC": "5",
        "MGS+Multivac": "6",
        "Bodolay": "7",
        "SELOVAC": "8",
        "AGV": "9",
        "Feinmechanik": "10",
        "GPACK": "11"
}

searchEM = {
   "":"1",
   "R245":"2",
   "OM-L" : "3"
}

searchTMAT = {
        '': '1',
        'Overwrap': '2',
        'Foil': '3',
        'Tyvek / Copolymer': '4'
}
searchLoad = {
   'Manual' : '1',
   'Automatic' : '2'
}
searchPC = {
   'PI': 'PI',
   'MP': 'MP',
   'PI/MP': 'MP/PI'
} 
def get_element_by_name(element_name: str) -> WebElement:
   required_element = elems.get(element_name)
   elem = driver.find_element(By.NAME, required_element)
   return elem


nome = 'PLANILHA1.xlsx' #colocar o nome da tabela aqui

driver = webdriver.Chrome()
url = 'https://yellow-pond-023165e0f.4.azurestaticapps.net/equipment/' #colocar o link do site aqui
driver.get(url)
time.sleep(3) #dar um tempo pra pagina carregar

continueEmail = driver.find_element(By.NAME, '//*[@id="__next"]/div/div/button')
continueEmail.click()
time.sleep(4)
window = driver.window_handles
driver.switch_to.window(window[1])

login = driver.find_element(By.NAME,'//*[@id="i0116"]')
login.send_keys('FFrancisco8@its.jnj.com')
time.sleep(2)

avancar = driver.find_element(By.NAME,'//*[@id="idSIButton9"]')
avancar.click()
time.sleep(7)

driver.switch_to.window(window[0])

teste = driver.find_element(By.NAME,'//*[@id="__next"]/div/div/div[1]/div/button')
teste.click()
time.sleep(4)

action_chains = ActionChains(driver)
action_chains.move_to_element_with_offset(teste, 0, 50).click().perform()
time.sleep(4)

equipment = driver.find_element(By.NAME, '//*[@id="__next"]/div/div/div/div[2]/a[2]')
equipment.click()
time.sleep(5)

tabela = 'Equipment'

dados = pd.read_excel(nome, sheet_name=tabela, na_values=[''])
num_minha_listas, num_colunas = dados.shape
indice_inicio = 16


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

       area = Select(get_element_by_name("Area"))
       area.select_by_value(searchArea.get(minha_lista[2],''))
                    
       EB = Select(get_element_by_name("EB"))
       EB.select_by_value(searchEB.get(minha_lista[3],''))

       EM = Select(get_element_by_name("EM"))
       EM.select_by_value(searchEM.get(minha_lista[4],''))
       
       EN = driver.find_element(get_element_by_name("EN")) 
       EN.send_keys(minha_lista[5])

       PL = driver.find_element(get_element_by_name("PL")) 
       PL.send_keys(minha_lista[6])

       QD = driver.find_element((get_element_by_name("QD")))
       QD.send_keys(minha_lista[7])

       YS = driver.find_element(get_element_by_name("YS"))
       YS.send_keys(minha_lista[8])

       TM = Select(get_element_by_name("TM"))
       TM.select_by_value("2")

       TE = Select(get_element_by_name("TE"))
       TE.select_by_value(searchTE.get(minha_lista[10],''))

       TMAT = Select(get_element_by_name("TMAT"))
       TMAT.select_by_value(searchTMAT.get(minha_lista[11],''))

       LOAD = Select(get_element_by_name("Load"))
       LOAD.select_by_value(searchLoad.get(minha_lista[12],''))

       OFFLOAD = Select(get_element_by_name("OffLoad"))
       OFFLOAD.select_by_value(searchLoad.get(minha_lista[13],''))

       SV = driver.find_element(get_element_by_name("SV"))
       SV.send_keys(int(minha_lista[18]))

       TVS = driver.find_element(get_element_by_name("TVS"))
       TVS.send_keys('83827629')

       PC = Select(get_element_by_name("PC"))
       PC.select_by_value(searchPC.get(minha_lista[20]),'')

       TOTALCY = driver.find_element(get_element_by_name("TOTALCY"))
       if(minha_lista[22]!=''):
        TOTALCY.send_keys(minha_lista[21])
      
       AVGS= driver.find_element(get_element_by_name("AVGS"))
       AVGS.send_keys(int(round(minha_lista[23])))
     
       AEY = driver.find_element(get_element_by_name("AEY"))
       if(minha_lista[24]!=''):
        AEY.send_keys(100*(minha_lista[24]))

       AVGEV = driver.find_element(get_element_by_name("AVGEV"))
       if(minha_lista[25]!=''):
         AVGEV.send_keys(minha_lista[25])

       AVGEC = driver.find_element(get_element_by_name('AVGEC'))
       if(minha_lista[26]!=''):
         AVGEC.send_keys(100*(minha_lista[26]))

       AVGCM = driver.find_element(get_element_by_name("AVGCM"))
       if(minha_lista[27]!=''):
         AVGCM.send_keys(minha_lista[27])
       NPC = driver.find_element(get_element_by_name("NPC"))
       if(minha_lista[28]!=''):
         NPC.send_keys(int(round(minha_lista[28])))
       SQRO = driver.find_element(get_element_by_name("SQRO"))
       SQRO.send_keys(minha_lista[32])
       SHIFTO = driver.find_element(get_element_by_name("SHIFTO"))
       SHIFTO.send_keys(minha_lista[33])
       MC = driver.find_element(get_element_by_name("MC"))
       MC.send_keys(minha_lista[34])
       OPTS = driver.find_element(get_element_by_name("OPTS")) #round(numero*130, 2)
       OPTS.send_keys(minha_lista[16])
       OPTD = driver.find_element(get_element_by_name("OPTD"))
       OPTD.send_keys(minha_lista[17])

      
       FBPTEMPMIN = driver.find_element(get_element_by_name("FBPTEMPMIN"))
       if(minha_lista[35]!=''):
         FBPTEMPMIN.send_keys(int(round(minha_lista[35])))
       FBPTEMPSET = driver.find_element(get_element_by_name("FBPTEMPSET"))
       if(minha_lista[36]!=''):
         FBPTEMPSET.send_keys(int(round(minha_lista[36])))
       FBPTEMPMAX = driver.find_element(get_element_by_name("FBPTEMPMAX"))    
       if(minha_lista[37]!=''):
         FBPTEMPMAX.send_keys(int(round(minha_lista[37])))
       FBPPRESSUREMIN = driver.find_element(get_element_by_name("FBPPRESSUREMIN"))
       if(minha_lista[38]!=''):
         FBPPRESSUREMIN.send_keys(int(round(minha_lista[38])))
       FBPPRESSURESET = driver.find_element(get_element_by_name('FBPPRESSURESET'))
       if(minha_lista[39]!=''):
         FBPPRESSURESET.send_keys(int(round(minha_lista[39])))
       FBPPRESSUREMAX = driver.find_element(get_element_by_name('FBPPRESSUREMAX'))    
       if(minha_lista[40]!=''):
         FBPPRESSUREMAX.send_keys(int(round(minha_lista[40])))
       FBPTIMEMIN = driver.find_element(get_element_by_name('FBPTIMEMIN'))  
       if(minha_lista[41]!=''):
         FBPTIMEMIN.send_keys(int(round(minha_lista[41])))
       FBPTIMESET = driver.find_element(get_element_by_name('FBPTIMESET'))
       if(minha_lista[42]!=''):
         FBPTIMESET.send_keys(int(round(minha_lista[42])))
       FBPTIMEMAX = driver.find_element(get_element_by_name('FBPTIMEMAX'))    
       if(minha_lista[43]!=''):
         FBPTIMEMAX.send_keys(int(round(minha_lista[43])))
     


       FCPTEMPMIN = driver.find_element(get_element_by_name('FCPTEMPMIN'))
       if(minha_lista[44]!=''):
         FCPTEMPMIN.send_keys(int(round(minha_lista[44])))
       FCPTEMPSET = driver.find_element(get_element_by_name('FCPTEMPSET'))
       if(minha_lista[45]!=''):
         FCPTEMPSET.send_keys(int(round(minha_lista[45])))
       FCPTEMPMAX = driver.find_element(get_element_by_name('FCPTEMPMAX'))    
       if(minha_lista[46]!=''):
         FCPTEMPMAX.send_keys(int(round(minha_lista[46])))
       FCPPRESSUREMIN = driver.find_element(get_element_by_name('FCPPRESSUREMIN'))
       if(minha_lista[47]!=''):
         FCPPRESSUREMIN.send_keys(int(round(minha_lista[47])))
       FCPPRESSURESET = driver.find_element(get_element_by_name('FCPPRESSURESET'))
       if(minha_lista[48]!=''):
         FCPPRESSURESET.send_keys(int(round(minha_lista[48])))
       FCPPRESSUREMAX = driver.find_element(get_element_by_name('FCPPRESSUREMAX'))    
       if(minha_lista[49]!=''):
         FCPPRESSUREMAX.send_keys(int(round(minha_lista[49])))
       FCPTIMEMIN = driver.find_element(get_element_by_name('FCPTIMEMIN'))  
       if(minha_lista[50]!=''):
         FCPTIMEMIN.send_keys(int(round(minha_lista[50])))
       FCPTIMESET = driver.find_element('FCPTIMESET')
       if(minha_lista[51]!=''):
         FCPTIMESET.send_keys(int(round(minha_lista[51])))
       FCPTIMEMAX = driver.find_element(get_element_by_name('FCPTIMEMAX'))    
       if(minha_lista[52]!=''):
         FCPTIMEMAX.send_keys(int(round(minha_lista[52])))


       SPTEMPMIN = driver.find_element(get_element_by_name('SPTEMPMIN'))
       if(minha_lista[53]!=''):
         SPTEMPMIN.send_keys(int(round(minha_lista[53])))
       SPTEMPSET = driver.find_element(get_element_by_name('SPTEMPSET'))
       if(minha_lista[54]!=''):
         SPTEMPSET.send_keys(int(round(minha_lista[54])))
       SPTEMPMAX = driver.find_element(get_element_by_name('SPTEMPMAX'))    
       if(minha_lista[55]!=''):
         SPTEMPMAX.send_keys(int(round(minha_lista[55])))
       SPPRESSUREMIN = driver.find_element(get_element_by_name('SPPRESSUREMIN'))
       if(minha_lista[56]!=''):
         SPPRESSUREMIN.send_keys(int(round(minha_lista[56])))
       SPPRESSURESET = driver.find_element(get_element_by_name('SPPRESSURESET'))
       if(minha_lista[57]!=''):
         SPPRESSURESET.send_keys(int(round(minha_lista[57])))
       SPPRESSUREMAX = driver.find_element('s')    
       if(minha_lista[58]!=''):
         SPPRESSUREMAX.send_keys(int(round(minha_lista[58])))
       SPTIMEMIN = driver.find_element('SPTIMEMIN')  
       if(minha_lista[59]!=''):
         SPTIMEMIN.send_keys(int(round(minha_lista[59])))
       SPTIMESET = driver.find_element('SPTIMESET')
       if(minha_lista[60]!=''):
         SPTIMESET.send_keys(int(round(minha_lista[60])))
       SPTIMEMAX = driver.find_element('SPTIMEMAX')    
       if(minha_lista[61]!=''):
         SPTIMEMAX.send_keys(int(round(minha_lista[61])))
    
       time.sleep(200)

     






     
