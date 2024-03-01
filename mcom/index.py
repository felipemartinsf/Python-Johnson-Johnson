import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
import random
import time
import shutil
import os 

def wayOut():
   close = driver.find_element(By.XPATH,"//button[.//span[text()='close']]")
   close.click()
   time.sleep(5)
   confirm = driver.find_element(By.XPATH,"//button[text()='Cancel?']")
   confirm.click()
def download_implemented():
   GenerateReport = driver.find_element(By.XPATH,"//button[text()=' Generate Report ']")
   time.sleep(1)
   GenerateReport.click()
   time.sleep(1)
   HazardList = driver.find_element(By.XPATH,"//button[text()=' Hazard List ']")
   HazardList.click()
   time.sleep(20)
   wait = WebDriverWait(driver, 20)
   HazardState = Select(wait.until(EC.visibility_of_element_located((By.ID, 'implementationState'))))
   HazardState.select_by_value("implemented")
   time.sleep(1)
   GenerateFile = driver.find_element(By.XPATH,"//button[text()=' Generate ']")
   GenerateFile.click()
   time.sleep(5)
   try:
        # Verifica se a mensagem de erro está presente na página
        error = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'No data for report found. Please consider updating report criteria.')]"))
        )
        # Se a mensagem de erro estiver presente, chama a função wayOut()
        wayOut()
        return 1
   except TimeoutException:
        # Se a mensagem de erro não estiver presente, continua normalmente
        return 0

def download_open():
   GenerateReport = driver.find_element(By.XPATH,"//button[text()=' Generate Report ']")
   time.sleep(1)
   GenerateReport.click()
   time.sleep(1)
   HazardList = driver.find_element(By.XPATH,"//button[text()=' Hazard List ']")
   HazardList.click()
   time.sleep(20)
   wait = WebDriverWait(driver, 20)
   HazardState = Select(wait.until(EC.visibility_of_element_located((By.ID, 'implementationState'))))
   HazardState.select_by_value("outstanding")
   GenerateFile = driver.find_element(By.XPATH,"//button[text()=' Generate ']")
   time.sleep(1)
   GenerateFile.click()
   time.sleep(1)
   try:
        # Verifica se a mensagem de erro está presente na página
        error = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'No data for report found. Please consider updating report criteria.')]"))
        )
        # Se a mensagem de erro estiver presente, chama a função wayOut()
        wayOut()
        return 1
   except TimeoutException:
        # Se a mensagem de erro não estiver presente, continua normalmente

        return 0

def lastFile():
   pasta = r'C:\Users\FFranci8\Downloads'
   arquivos = os.listdir(pasta)
   arquivos = [arquivo for arquivo in arquivos if os.path.isfile(os.path.join(pasta, arquivo))]
   caminhos = [os.path.join(pasta, arquivo) for arquivo in arquivos]
   datas_de_modificacao = [(arquivo, os.path.getmtime(caminho)) for arquivo, caminho in zip(arquivos, caminhos)]
   datas_de_modificacao.sort(key=lambda x: x[1], reverse=True)
   arquivo_mais_recente = datas_de_modificacao[0][0]
   caminho_completo = os.path.join(pasta, arquivo_mais_recente)
   numeros_aleatorios =str(random.randint(1, 3000))
   caminho_novo = os.path.join(os.path.dirname(caminho_completo), numeros_aleatorios + '.xlsx')
   os.rename(caminho_completo, caminho_novo)
   return (caminho_novo)

def throwFile(arquivo_mais_recente):
   caminho_destino = os.path.abspath('final')
   shutil.move(arquivo_mais_recente, caminho_destino)


def statusOpen(nome_arquivo_planilha):
    df_planilha1 = pd.read_excel(nome_arquivo_planilha)
    max_linhas = df_planilha1.shape[0]
    df_planilha1['Status'] = ['OPEN'] * max_linhas
    df_planilha1.to_excel(nome_arquivo_planilha, index=False)

def statusImplemented(nome_arquivo_planilha):
    df_planilha1 = pd.read_excel(nome_arquivo_planilha)
    max_linhas = df_planilha1.shape[0]
    df_planilha1['Status'] = ['IMPLEMENTED'] * max_linhas
    df_planilha1.to_excel(nome_arquivo_planilha, index=False)

def mergeAll():
   file_path = r'C:\Users\FFranci8\OneDrive - JNJ\Área de Trabalho\MCOM\final'
   file_path_end= r'C:\Users\FFranci8\OneDrive - JNJ\Área de Trabalho\MCOM'
   arquivos_excel = [arquivo for arquivo in os.listdir(file_path) if arquivo.endswith('.xlsx')]
   df_final = pd.DataFrame()
   for arquivo in arquivos_excel:
    caminho_arquivo = os.path.join(file_path, arquivo)
    df = pd.read_excel(caminho_arquivo)
    df_final = pd.concat([df_final, df], ignore_index=True)
   planilha_final = 'planilha_final.xlsx'
   df_final.to_excel(planilha_final, index=False)
   caminho_final = os.path.join(file_path_end, planilha_final)
   return caminho_final
def deleteFiles(caminho_da_pasta):
    for nome_arquivo in os.listdir(caminho_da_pasta):
        caminho_arquivo = os.path.join(caminho_da_pasta, nome_arquivo)
        try:
            os.remove(caminho_arquivo)
        except Exception as e:
            print(f'Erro ao remover o arquivo "{nome_arquivo}": {e}')

def putOldFile(file):
   df = pd.read_excel(file)
   df['Seção'] = [None] * len(df.index)
   df['Questão'] = [None] * len(df.index)
   df['LE 5'] = [None] * len(df.index)
   df['LE 6'] = [None] * len(df.index)
   df['SIF'] = [None] * len(df.index)
   df['Grupo / Tecnologia'] = [None] * len(df.index)
   df.to_excel(file, index=False)
   InitialRatingValue = df['Initial Rating Value'].astype(str).tolist()
   InitialRatingValueDescription = df['Initial Rating Value Description'].astype(str).str.upper().tolist()
   MachineType = df['MachineType'].astype(str).tolist()
   HazardDescription = df['Hazard Description'].astype(str).tolist()
   ControlMeasures = df['Control Measures'].astype(str).tolist()
   arquivos = os.listdir(os.path.dirname(os.path.abspath(__file__)))
   arquivos_xlsb = [arquivo for arquivo in arquivos if arquivo.endswith('.xlsb')]
   de = None
   for arquivo in arquivos_xlsb:
      print("arquivo que vai ler:"+arquivo)
      de = pd.read_excel(arquivo, sheet_name='Hazard Assessments')
      InitialRatingValueOld = de['Initial Rating Value'].astype(str).tolist()
      HazardDescriptionOld = de['Hazard Description'].astype(str).tolist()
      ControlMeasuresOld = de['Control Measures'].astype(str).tolist()
      InitialRatingValueDescriptionOld = de['Initial Rating Value Description'].astype(str).str.upper().tolist()
      MachineTypeOld = de['Machine Type'].astype(str).tolist()
      size = de.shape[0]
      size2 = df.shape[0]
      for i in range(size):
         for j in range(size2):
             if(HazardDescriptionOld[i]==HazardDescription[j] and ControlMeasuresOld[i]==ControlMeasures[j] and InitialRatingValueOld[i]==InitialRatingValue[j] and InitialRatingValueDescriptionOld[i]==InitialRatingValueDescription[j] and MachineTypeOld[i]==MachineType[j]):
                 secao = de.at[i, 'Seção']
                 df.at[j, 'Seção'] = secao
                 questao = de.at[i, 'Questão']
                 df.at[j, 'Questão'] = questao
                 le5 = de.at[i, 'LE 5']
                 df.at[j, 'LE 5'] = le5
                 le6 = de.at[i, 'LE 6']
                 df.at[j, 'LE 6'] = le6
                 sif = de.at[i,'SIF']
                 df.at[j,'SIF'] = sif
                 grupo = de.at[i,'Grupo / Tecnologia']
                 df.at[j,'Grupo / Tecnologia'] = grupo

      de = None
   df.to_excel(file, index=False)                  





if os.path.exists(r'C:\Users\FFranci8\OneDrive - JNJ\Área de Trabalho\MCOM\final'):
   deleteFiles(r'C:\Users\FFranci8\OneDrive - JNJ\Área de Trabalho\MCOM\final')
   os.rmdir(r'C:\Users\FFranci8\OneDrive - JNJ\Área de Trabalho\MCOM\final')
if os.path.exists('planilha_final.xlsx'):
    os.remove('planilha_final.xlsx')
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36'
}
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument(f'user-agent={headers["User-Agent"]}')

driver = webdriver.Chrome(options=chrome_options)
url = 'https://mcomone.tuev-sued.com/' #colocar o link do site aqui
driver.get(url)
time.sleep(10) #dar um tempo pra pagina carregar
driver.maximize_window()
time.sleep(3)

login = driver.find_element(By.XPATH,"//button[text()='Login with company account']")
login.click()
time.sleep(20)
projects = driver.find_element(By.XPATH,"//a[text()='Projects']")
projects.click()
time.sleep(2)


rows = driver.find_elements(By.XPATH,"//tr[.//td[@class='name']]")
os.makedirs('final')

i = 0 

for i in range(len(rows)):
   rows = driver.find_elements(By.XPATH,"//tr[.//td[@class='name']]")
   time.sleep(2)
   rows[i].click()
   time.sleep(2)
   mistake_open = download_open()
   time.sleep(5)
   if(mistake_open == 0):
      file = lastFile()
      statusOpen(file)
      throwFile(file)
   mistake_implemented = download_implemented()
   time.sleep(5)
   if(mistake_implemented == 0):
      file = lastFile()
      statusImplemented(file)
      throwFile(file)
   projects.click()
   time.sleep(3)
final = mergeAll()
putOldFile(final)
print('tudo foi salvo dentro da planilha: "planilha_final"')


#No data for report found. Please consider updating report criteria.

'''
 de.drop(columns=['Risk Notes'], inplace=True)
      ordem = {'Concluído': 1, 'n/a': 2, '-': 3}
      de['ordem_status'] = de['Status'].map(ordem)
      de.sort_values(by='ordem_status', inplace=True)
      de.drop(columns=['ordem_status'], inplace=True)
      de.to_excel(arquivo, index=False)
'''

#colunas_selecionadas = de[['Seção', 'Questões', 'LE 6','LE 5']]
'''def lookFileName(file):
   df = pd.read_excel(file)
   primeira_string_hazard = str(df.loc[1, 'ProjectName']).split()[0]
   primeira_string_hazard = primeira_string_hazard.lower()
   print('primeira string extracted: '+ primeira_string_hazard)
   arquivos = os.listdir(os.path.dirname(os.path.abspath(__file__)))
   arquivos_xlsx = [arquivo for arquivo in arquivos if arquivo.endswith('.xlsx')]
   print(arquivos_xlsx)
   for arquivo in arquivos_xlsx:
      de = pd.read_excel(arquivo,sheet_name='Hazard Assessments')
      primeira_string_excel = str(de.loc[1, 'Project']).split()[0]
      primeira_string_excel = primeira_string_excel.lower()
      print('primeira string excel: ' + primeira_string_excel)
      if primeira_string_excel == primeira_string_hazard:
            print('arquivo: '+ arquivo)
            return arquivo
      else: print('The extracted excel and the one you put it are not in the same area')

def mergeREV(file:str):
   arquivo = lookFileName(file)
   df = pd.read_excel(file)
   de = pd.read_excel(arquivo,sheet_name='Hazard Assessments')
   colunas_selecionadas = de[['Seção', 'Questões', 'LE 6','LE 5']]

   for coluna in colunas_selecionadas.columns:
      df[coluna] = colunas_selecionadas[coluna]

   df.to_excel(file, index=False)'''