import pyautogui
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.keys import Keys          
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
import traceback
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Carregar a planilha
Planilha_xml = pd.read_excel("aux expotação robo xml.xls")

# Redefinindo o cabeçalho para a segunda linha (índice 1)
Planilha_xml.columns = Planilha_xml.iloc[0]

# Removendo a segunda linha do DataFrame
Planilha_xml = Planilha_xml[1:].reset_index(drop=True)

# Filtrar as linhas que têm números na coluna "N° NF"
# Usar pd.to_numeric para converter e definir errors='coerce' para NaN valores não numéricos
Planilha_xml['N° NF'] = pd.to_numeric(Planilha_xml['N° NF'], errors='coerce')

# Manter apenas as linhas onde 'N° NF' não é NaN
Planilha_xml = Planilha_xml.dropna(subset=['N° NF'])
#print(Planilha_xml)

def click_selenium(selector, value):
    try:
        print("Clicando no botão...")
        elemento = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((selector, value)))
        elemento.click()
    except Exception as e:
        print(f"Erro ao clicar: {e}")

# Define o caminho para o diretório de downloads
download_dir = r"C:\Users\Usuario\Documents\XML"

# Configurações de opções para o Chrome
options = Options()

# Desabilita a detecção de automação
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# Configurações de preferências para downloads
prefs = {
    "download.default_directory": download_dir,  # Define o diretório de download
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True,  # Desativa a proteção contra downloads
}
options.add_experimental_option("prefs", prefs)

# # Inicializa o driver do Chrome  
#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver = webdriver.Chrome(options=options)
driver.get("https://www.nestle-parceiro.com.br/Portal/PortalNestle.aspx?contextId=fb8b300d-0f7d-4df7-8a63-672c07287d44#")
driver.maximize_window()

try:
    print("Inserir e-mail do gestor...")
    usuario = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'txtUsuario')))
    usuario.click()
    usuario.send_keys('jorge.fahl@dglnet.com.br')
    print("E-mail inserido com sucesso...")               
except Exception as e:
    print("Erro ao inserir o e-mail do gestor:", e)

try:
    print("Inserir e-mail do gestor...")
    email = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'txtSenha')))
    email.click()
    email.send_keys('Faturamento@2024.2')
    print("E-mail inserido com sucesso...")               
except Exception as e:
    print("Erro ao inserir o e-mail do gestor:", e)

try:
    print("Inserir e-mail do gestor...")
    botao_entrar = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'btnEntrar')))
    botao_entrar.click()
    print("E-mail inserido com sucesso...")               
except Exception as e:
    print("Erro ao inserir o e-mail do gestor:", e)

click_selenium(By.ID, 'chkTermo')
click_selenium(By.ID, 'btnEnviar')
click_selenium(By.ID, 'abrirmenu')

try:
    print("Mover o cursor sobre o botão...")
    # Espera até que o elemento esteja presente na página
    botao_entrar = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="lblMenu"]/li[3]/a')))
    # Mover o cursor sobre o elemento
    action = ActionChains(driver)
    action.move_to_element(botao_entrar).perform()
    print("Cursor movido sobre o botão com sucesso...")
except Exception as e:
    print("Erro ao mover o cursor sobre o botão:", e)

try:
    print("Obtendo o href da nota fiscal...")

    # Aguarda até que o elemento esteja presente na página
    campo_nota = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="listMenu-id-4"]/li[1]/a'))
    )
    # Obtém o atributo href do elemento
    href = campo_nota.get_attribute("href")
    print("Href da nota fiscal:", href)
 
except TimeoutException as e:
    print("Erro: Tempo de espera esgotado para encontrar o elemento:", e)
except NoSuchElementException as e:
    print("Erro: O elemento não foi encontrado na página:", e)
except Exception as e:
    print("Erro ao obter o href da nota fiscal:", e)
driver.get(href)

for i, linha in enumerate(Planilha_xml.index):
    nota = Planilha_xml.loc[linha, "N° NF"]
    nota = int(nota)
    data_inicial = Planilha_xml.loc[linha, "Data NF"]
    data_mes_anterior = data_inicial - relativedelta(weeks=1)
    data_mes_anterior_str = data_mes_anterior.strftime("%d/%m/%Y")
    data_final = Planilha_xml.loc[linha, "Data NF"]
    data_final = data_final.strftime("%d/%m/%Y")

    print(f'nota:{nota} data inicial:{data_mes_anterior_str} data final:{data_final}')
    try:
        print("Inserir nota fiscal...")
        # Aguarda o campo ficar clicável
        campo_nota = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, 'txtNotaFiscal')))
        campo_nota.click()
        campo_nota.clear()
        campo_nota.send_keys(nota)
        print("Nota inserida com sucesso...")               
    except Exception as e:
        print("Erro ao inserir a nota fiscal:", e)

    try:
        print("Inserir data inicial...")
        campo_data_inicial = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'txtDataEmissaoIni')))
        campo_data_inicial.clear()
        campo_data_inicial.send_keys(data_mes_anterior_str)
        print("data inicial inserida com sucesso...")                  
    except Exception as e:
        print("Erro ao inserir data inicial:", e)

    try:
        print("Inserir data final ...")
        campo_data_final = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txtDataEmissaoFim"]')))
        campo_data_final.clear()
        campo_data_final.send_keys(data_final)
        print("data final inserida com sucesso...")                  
    except Exception as e:
        print("Erro ao inserir data final:", e)

    click_selenium(By.ID, 'btnProcurar')
    if i ==0:
        pyautogui.sleep(0.1)
    click_selenium(By.ID, 'gvNotaFiscal_ctl02_hypDownload')

print('Terminando Salvamento')
pyautogui.sleep(10)
driver.quit()
print('Download dos XML finalizado')





 





