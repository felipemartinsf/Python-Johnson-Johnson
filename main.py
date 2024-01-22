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

elems = {
   "Franchise" : "franchise",
   "Area": "area_id",
   "Site": "site_id",
   "EB" : "equipment_brand_id",
   "EM" : "equipment_model_id",
   "EN" : "number",
   "PL" : "name",
   "QD" : "qt_of_duplicate",
   "YS" : "years_in_service",
   "TM" : "machine_type_id",
   "TE" : "equipment_type_id",
   "TMAT" : "material_type_id",
   "Load" : "equipment_loading_id",
   "OffLoad" : "equipment_offloading_id",
   "SV" :  "site_volume_processed",
   "TVS" : "total_site_volume",
   "PC" : "program_classification",
   "TOTALCY" : "total_codes_per_year",
    "AVGS": "average_lots_shift",
    "AEY": "average_yield_full_last_year",
    "AVGEV": "average_oee_full_last_year_validated",
    "AVGEC": "average_oee_full_last_year_currently",
    "AVGCM": "average_cycle_per_minute",
    "NPC": "products_per_cycle",
    "SQRO": "square_meters_to_be_operated",
    "SHIFTO": "operator_needed_per_shift",
    "MC": "machine_cost",
    "OPTS": "operation_run_time_shift",
    "OPTD": "operation_run_time_day",
     "FBPTEMPMIN": "forming_blowing_parameters_temperature_celsius_min",
    "FBPTEMPSET": "forming_blowing_parameters_temperature_celsius_setp",
    "FBPTEMPMAX": "forming_blowing_parameters_temperature_celsius_max",
    "FBPPRESSUREMIN": "forming_blowing_parameters_pressure_psi_min",
    "FBPPRESSURESET": "forming_blowing_parameters_pressure_psi_setp",
    "FBPPRESSUREMAX": "forming_blowing_parameters_pressure_psi_max",
    "FBPTIMEMIN": "forming_blowing_parameters_time_seconds_min",
    "FBPTIMESET": "forming_blowing_parameters_time_seconds_setp",
    "FBPTIMEMAX": "forming_blowing_parameters_time_seconds_max",
    "FCPTEMPMIN": "forming_clamping_parameters_temperature_celsius_min",
    "FCPTEMPSET": "forming_clamping_parameters_temperature_celsius_setp",
    "FCPTEMPMAX": "forming_clamping_parameters_temperature_celsius_max",
    "FCPPRESSUREMIN": "forming_clamping_parameters_pressure_psi_min",
    "FCPPRESSURESET": "forming_clamping_parameters_pressure_psi_setp",
    "FCPPRESSUREMAX": "forming_clamping_parameters_pressure_psi_max",
    "FCPTIMEMIN": "forming_clamping_parameters_time_seconds_min",
    "FCPTIMESET": "forming_clamping_parameters_time_seconds_setp",
    "FCPTIMEMAX": "forming_clamping_parameters_time_seconds_max",
    "SPTEMPMIN": "sealing_parameters_temperature_celsius_min",
    "SPTEMPSET": "sealing_parameters_temperature_celsius_setp",
    "SPTEMPMAX": "sealing_parameters_temperature_celsius_max",
    "SPPRESSUREMIN": "sealing_parameters_pressure_psi_min",
    "SPPRESSURESET": "sealing_parameters_pressure_psi_setp",
    "SPPRESSUREMAX": "sealing_parameters_pressure_psi_max",
    "SPTIMEMIN": "sealing_parameters_time_seconds_min",
    "SPTIMESET": "sealing_parameters_time_seconds_setp",
    "SPTIMEMAX": "sealing_parameters_time_seconds_max",
    "TYTEMPMIN": "tyvek_sealing_parameters_temperature_celsius_min",
    "TYTEMPSET": "tyvek_sealing_parameters_temperature_celsius_setp",
    "TYTEMPMAX": "tyvek_sealing_parameters_temperature_celsius_max",
    "TYPRESSUREMIN": "tyvek_sealing_parameters_pressure_psi_min",
    "TYPRESSURESET": "tyvek_sealing_parameters_pressure_psi_setp",
    "TYPRESSUREMAX": "tyvek_sealing_parameters_pressure_psi_max",
    "TYTIMEMIN": "tyvek_sealing_parameters_time_seconds_min",
    "TYTIMESET": "tyvek_sealing_parameters_time_seconds_setp",
    "TYTIMEMAX": "forming_blowing_parameters_time_seconds_max",
    "UTIELE": "utility_electricity",
    "UTIAIR": "utility_compressed_air",
    "UTISTEAM": "utility_steam",
    "UTIWATER": "utility_water",
    "UTIVACUUM": "utility_vacuum",
}

searchArea = {
        "Endo": "2",
        "Mesh": "3",
        "Print Shop": "4",
        "Catgut": "5",
        "EO": "6",
        "Packaging": "7",
        "Over": "9"
}

searchTE= {
    'First Sealing': "1",
    'Forming': "2",
    'Fill/First Sealing': "3",
    'Printer': "4",
    'Form/Fill/Firt Sealing': "5",
    'Other': "6",
    'Form/Fill/Firt Sealing/Cutting': "7",
    'Form/Fill/Firts Sealing/Cutting': "8",
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
        'Foil and Overwrap': '3'
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
driver.maximize_window()
time.sleep(3)

continueEmail = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/button')
continueEmail.click()
time.sleep(5)
window = driver.window_handles
driver.switch_to.window(window[1])

login = driver.find_element(By.NAME,'loginfmt')
login.send_keys('FFrancisco8@its.jnj.com')
time.sleep(2)

avancar = driver.find_element(By.XPATH,'//*[@id="idSIButton9"]')
avancar.click()
time.sleep(7)

driver.switch_to.window(window[0])

teste = driver.find_element(By.XPATH,'//*[@id="__next"]/div/div/div[1]/div/button')
teste.click()
time.sleep(4)

action_chains = ActionChains(driver)
action_chains.move_to_element_with_offset(teste, 0, 50).click().perform()

time.sleep(5)
equipment = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/div/div[2]/a[2]')
equipment.click()
time.sleep(4)

tabela = 'Equipment'

dados = pd.read_excel(nome, sheet_name=tabela, na_values=[''])
num_minha_listas, num_colunas = dados.shape
indice_inicio = 55


if 0 <= indice_inicio < num_minha_listas:
  for indice_minha_lista in range(indice_inicio, num_minha_listas):
       
       lista = dados.iloc[indice_minha_lista].tolist()
       minha_lista = ['' if pd.isna(elemento) else elemento for elemento in lista] #aqui verelifica se o elemento é nulo, se for ele coloca ''
       print(minha_lista)
       if(minha_lista[0]==''):break

       newMachine = driver.find_element(By.XPATH,'//*[@id="__next"]/div/div/main/header/button')
       newMachine.click()
       time.sleep(4)
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
       
       EN = (get_element_by_name("EN")) 
       EN.send_keys(minha_lista[5])

       PL = (get_element_by_name("PL")) 
       PL.send_keys(minha_lista[6])

       QD = ((get_element_by_name("QD")))
       QD.send_keys(minha_lista[7])

       YS = (get_element_by_name("YS"))
       YS.send_keys(int(minha_lista[8]))

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

       SV = (get_element_by_name("SV"))
       SV.send_keys(int(minha_lista[18]))

       TVS = (get_element_by_name("TVS"))
       TVS.send_keys('83827629')

       PC = Select(get_element_by_name("PC"))
       PC.select_by_value(searchPC.get(minha_lista[20]))

       TOTALCY = (get_element_by_name("TOTALCY"))
       if(minha_lista[22] !='' and minha_lista[35]!='-'):
        TOTALCY.send_keys(minha_lista[21])
      
       AVGS= (get_element_by_name("AVGS"))
       AVGS.send_keys(int(round(minha_lista[23])))
     
       AEY = (get_element_by_name("AEY"))
       if(minha_lista[24] !=''):
        AEY.send_keys(100*(minha_lista[24]))

       AVGEV = (get_element_by_name("AVGEV"))
       if(minha_lista[25] !=''):
         AVGEV.send_keys(round(100*(minha_lista[25]),1))

       AVGEC = (get_element_by_name('AVGEC'))
       if(minha_lista[26] !=''):
         AVGEC.send_keys(round(100*(minha_lista[26]),1))

       AVGCM = (get_element_by_name("AVGCM"))
       if(minha_lista[27] !=''):
         AVGCM.send_keys(minha_lista[27])
       NPC = (get_element_by_name("NPC"))
       if(minha_lista[28] !=''):
         NPC.send_keys((minha_lista[28]))
       SQRO = (get_element_by_name("SQRO"))
       SQRO.send_keys(minha_lista[32])
       SHIFTO = (get_element_by_name("SHIFTO"))
       SHIFTO.send_keys(minha_lista[33])
       MC = (get_element_by_name("MC"))
       MC.send_keys(minha_lista[34])
       OPTS = (get_element_by_name("OPTS")) 
       OPTS.send_keys(minha_lista[16])
       OPTD = (get_element_by_name("OPTD"))
       OPTD.send_keys(minha_lista[17])

      
       FBPTEMPMIN = (get_element_by_name("FBPTEMPMIN"))
       if(minha_lista[35] !='' and minha_lista[35]!='-'):
         FBPTEMPMIN.send_keys(((minha_lista[35])))
       FBPTEMPSET = (get_element_by_name("FBPTEMPSET"))
       if(minha_lista[36] !='' and minha_lista[36]!='-'):
         FBPTEMPSET.send_keys(((minha_lista[36])))
       FBPTEMPMAX = (get_element_by_name("FBPTEMPMAX"))    
       if(minha_lista[37] !='' and minha_lista[37]!='-'):
         FBPTEMPMAX.send_keys(((minha_lista[37])))
       FBPPRESSUREMIN = (get_element_by_name("FBPPRESSUREMIN"))
       if(minha_lista[38] !='' and minha_lista[38]!='-'):
         FBPPRESSUREMIN.send_keys(((minha_lista[38])))
       FBPPRESSURESET = (get_element_by_name('FBPPRESSURESET'))
       if(minha_lista[39] !='' and minha_lista[39]!='-'):
         FBPPRESSURESET.send_keys(((minha_lista[39])))
       FBPPRESSUREMAX = (get_element_by_name('FBPPRESSUREMAX'))    
       if(minha_lista[40] !='' and minha_lista[40]!='-'):
         FBPPRESSUREMAX.send_keys(((minha_lista[40])))
       FBPTIMEMIN = (get_element_by_name('FBPTIMEMIN'))  
       if(minha_lista[41] !='' and minha_lista[41]!='-'):
         FBPTIMEMIN.send_keys(((minha_lista[41])))
       FBPTIMESET = (get_element_by_name('FBPTIMESET'))
       if(minha_lista[42] !='' and minha_lista[42]!='-'):
         FBPTIMESET.send_keys(((minha_lista[42])))
         FBPTIMEMAX = (get_element_by_name('FBPTIMEMAX'))    
       if(minha_lista[43] !='' and minha_lista[43]!='-'):
         FBPTIMEMAX.send_keys(((minha_lista[43])))



       FCPTEMPMIN = (get_element_by_name('FCPTEMPMIN'))
       if(minha_lista[44] !='' and minha_lista[44]!='-'):
         FCPTEMPMIN.send_keys(((minha_lista[44])))
       FCPTEMPSET = (get_element_by_name('FCPTEMPSET'))
       if(minha_lista[45] !='' and minha_lista[45]!='-'):
         FCPTEMPSET.send_keys(((minha_lista[45])))
       FCPTEMPMAX = (get_element_by_name('FCPTEMPMAX'))    
       if(minha_lista[46] !='' and minha_lista[46]!='-'):
         FCPTEMPMAX.send_keys(((minha_lista[46])))
       FCPPRESSUREMIN = (get_element_by_name('FCPPRESSUREMIN'))
       if(minha_lista[47] !='' and minha_lista[47]!='-'):
         FCPPRESSUREMIN.send_keys(((minha_lista[47])))
       FCPPRESSURESET = (get_element_by_name('FCPPRESSURESET'))
       if(minha_lista[48] !='' and minha_lista[48]!='-'):
         FCPPRESSURESET.send_keys(((minha_lista[48])))
       FCPPRESSUREMAX = (get_element_by_name('FCPPRESSUREMAX'))    
       if(minha_lista[49] !='' and minha_lista[49]!='-'):
         FCPPRESSUREMAX.send_keys(((minha_lista[49])))
       FCPTIMEMIN = (get_element_by_name('FCPTIMEMIN'))  
       if(minha_lista[50] !='' and minha_lista[50]!='-'):
         FCPTIMEMIN.send_keys(((minha_lista[50])))
       FCPTIMESET = get_element_by_name('FCPTIMESET')
       if(minha_lista[51] !='' and minha_lista[51]!='-'):
         FCPTIMESET.send_keys(((minha_lista[51])))
         FCPTIMEMAX = (get_element_by_name('FCPTIMEMAX'))    
       if(minha_lista[52] !='' and minha_lista[52]!='-'):
         FCPTIMEMAX.send_keys(((minha_lista[52])))


       SPTEMPMIN = (get_element_by_name('SPTEMPMIN'))
       if(minha_lista[53] !='' and minha_lista[53]!='-'):
         SPTEMPMIN.send_keys(((minha_lista[53])))
       SPTEMPSET = (get_element_by_name('SPTEMPSET'))
       if(minha_lista[54] !='' and minha_lista[54]!='-'):
         SPTEMPSET.send_keys(((minha_lista[54])))
       SPTEMPMAX = (get_element_by_name('SPTEMPMAX'))    
       if(minha_lista[55] !='' and minha_lista[55]!='-'):
         SPTEMPMAX.send_keys(((minha_lista[55])))
       SPPRESSUREMIN = (get_element_by_name('SPPRESSUREMIN'))
       if(minha_lista[56] !='' and minha_lista[56]!='-'):
         SPPRESSUREMIN.send_keys(((minha_lista[56])))
       SPPRESSURESET = (get_element_by_name('SPPRESSURESET'))
       if(minha_lista[57] !='' and minha_lista[57]!='-'):
         SPPRESSURESET.send_keys(((minha_lista[57])))
       SPPRESSUREMAX = get_element_by_name('SPPRESSUREMAX')    
       if(minha_lista[58] !='' and minha_lista[58]!='-'):
         SPPRESSUREMAX.send_keys(((minha_lista[58])))
       SPTIMEMIN = get_element_by_name('SPTIMEMIN')    
       if(minha_lista[59] !='' and minha_lista[59]!='-'):
         SPTIMEMIN.send_keys(((minha_lista[59])))
       SPTIMESET = get_element_by_name('SPTIMESET')   
       if(minha_lista[60] !='' and minha_lista[60]!='-'):
         SPTIMESET.send_keys(((minha_lista[60])))
       SPTIMEMAX = get_element_by_name('SPTIMEMAX')     
       if(minha_lista[61] !='' and minha_lista[61]!='-'):
        SPTIMEMAX.send_keys(((minha_lista[61])))

        TYTEMPMIN = (get_element_by_name('TYTEMPMIN'))
       if(minha_lista[62] !='' and minha_lista[62]!='-'):
         TYTEMPMIN.send_keys(((minha_lista[62])))
       TYTEMPSET = (get_element_by_name('TYTEMPSET'))
       if(minha_lista[63] !='' and minha_lista[63]!='-'):
         TYTEMPSET.send_keys(((minha_lista[63])))
       TYTEMPMAX = (get_element_by_name('TYTEMPMAX'))    
       if(minha_lista[64] !='' and minha_lista[64]!='-'):
         TYTEMPMAX.send_keys(((minha_lista[64])))
       TYPRESSUREMIN = (get_element_by_name('TYPRESSUREMIN'))
       if(minha_lista[65] !='' and minha_lista[65]!='-'):
         TYPRESSUREMIN.send_keys(((minha_lista[65])))
       TYPRESSURESET = (get_element_by_name('TYPRESSURESET'))
       if(minha_lista[66] !='' and minha_lista[66]!='-'):
         TYPRESSURESET.send_keys(((minha_lista[66])))
       TYPRESSUREMAX = get_element_by_name('TYPRESSUREMAX')    
       if(minha_lista[67] !='' and minha_lista[67]!='-'):
         TYPRESSUREMAX.send_keys(((minha_lista[67])))
       TYTIMEMIN = get_element_by_name('TYTIMEMIN')    
       if(minha_lista[68] !='' and minha_lista[68]!='-'):
         TYTIMEMIN.send_keys(((minha_lista[68])))
       TYTIMESET = get_element_by_name('TYTIMESET')
       if(minha_lista[69] !='' and minha_lista[69]!='-'):
         TYTIMESET.send_keys(((minha_lista[69])))
       if(minha_lista[70] !='' and minha_lista[70]!='-'):
        TYTIMEMAX = driver.find_element(By.NAME, "TYTIMEMAX")
        TYTIMEMAX.send_keys(((minha_lista[70])))


       UTIELE = (get_element_by_name('UTIELE'))
       if(minha_lista[71] !=''):
         UTIELE.send_keys(((minha_lista[71])))
         UTIAIR = (get_element_by_name('UTIAIR'))
       if(minha_lista[72] !=''):
         UTIAIR.send_keys(((minha_lista[72])))
       UTISTEAM = (get_element_by_name('UTISTEAM'))    
       if(minha_lista[73] !=''):
         UTISTEAM.send_keys(((minha_lista[73])))
       UTIWATER = (get_element_by_name('UTIWATER'))
       if(minha_lista[74] !=''):
         UTIWATER.send_keys(((minha_lista[74])))
       UTIVACUUM = (get_element_by_name('UTIVACUUM'))
       if(minha_lista[75] !=''):
         UTIVACUUM.send_keys(((minha_lista[75])))
         
       save = driver.find_element(By.XPATH, "//button[text()='Save']")
       save.click()
       time.sleep(5)
     





driver.quit()
     
