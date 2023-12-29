from webbrowser import Chrome
import pandas as pd
import time # Pausar a execução do código
from selenium import webdriver # Automatizar a interação com o navegador
from selenium.webdriver.common.by import By # Selecionar elementos de paginas web
from selenium.webdriver.common.keys import Keys # Capturar o "teclado"

# Ler o arquivo Excell e exibi-lo
tabela = pd.read_excel('challenge.xlsx', engine="openpyxl")
print(tabela)
print(tabela.head())
print(tabela.columns)


# Tempo de espera do site carregar
time.sleep(3)

#C Configurações do Chrome para não fechar automaticamente
chrome_opcoes = webdriver.ChromeOptions()
chrome_opcoes.add_experimental_option("detach", True)

# Iniciar o driver do Selenium para o Chrome com as opções configuradas
chrome = webdriver.Chrome(options=chrome_opcoes)

# Abrir o site
chrome.get("https://rpachallenge.com/")

# Clicar no botão de Start
chrome.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()

# Loopar o arquivo para percorrer e imprimir. ele pega o index e a linha da função iterrows (trazendo o objedo da linha interia)
for i, row in tabela.iterrows():
    company_name = row["Company Name"]
    phone_number = row["Phone Number"]
    email = row["Email"]
    role = row["Role in Company"]
    last_name = row["Last Name "]
    address = row["Address"]
    first_name = row["First Name"]
    

    #Preenchendo os campos
    chrome.find_element(By.CSS_SELECTOR,'[ng-reflect-name="labelCompanyName"]').send_keys(company_name)
    chrome.find_element(By.CSS_SELECTOR,'[ng-reflect-name="labelPhone"]').send_keys(str(phone_number))
    chrome.find_element(By.CSS_SELECTOR,'[ng-reflect-name="labelEmail"]').send_keys(email)
    chrome.find_element(By.CSS_SELECTOR,'[ng-reflect-name="labelRole"]').send_keys(role)
    chrome.find_element(By.CSS_SELECTOR,'[ng-reflect-name="labelLastName"]').send_keys(last_name)
    chrome.find_element(By.CSS_SELECTOR,'[ng-reflect-name="labelAddress"]').send_keys(address)
    chrome.find_element(By.CSS_SELECTOR,'[ng-reflect-name="labelFirstName"]').send_keys(first_name)
    chrome.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()

chrome.quit