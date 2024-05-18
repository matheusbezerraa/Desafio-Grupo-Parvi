from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os
import urllib.request

# Função para baixar o arquivo Excel
def download_excel(url, download_path):
    urllib.request.urlretrieve(url, download_path)

# URL do arquivo Excel
excel_url = "https://www.rpachallenge.com/assets/downloadFiles/challenge.xlsx"
download_dir = os.path.expanduser("~") + "/Downloads"  # Diretório padrão de downloads
filename = "challenge.xlsx"                            # Nome do arquivo baixado
filepath = os.path.join(download_dir, filename)

# Baixar o arquivo Excel
download_excel(excel_url, filepath)

# Verificar se o arquivo foi baixado corretamente
if not os.path.exists(filepath):
    raise FileNotFoundError(f"Arquivo {filename} não encontrado no diretório {download_dir}")

# Ler a planilha baixada
df = pd.read_excel(filepath)

# Configurar o webdriver usando webdriver-manager
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Navegar até o site do desafio
driver.get('https://www.rpachallenge.com/')

# Iniciar o desafio
start_button = driver.find_element(By.XPATH, '//button[text()="Start"]')
start_button.click()

# Mapeamento dos campos de entrada
field_mapping = {
    'First Name': 'First Name',
    'Last Name': 'Last Name',
    'Company Name': 'Company Name',
    'Role in Company': 'Role in Company',
    'Address': 'Address',
    'Email': 'Email',
    'Phone Number': 'Phone Number'
}

# Preencher o formulário
for index, row in df.iterrows():
    # Obter todos os campos do formulário na página
    fields = driver.find_elements(By.XPATH, '//input[@type="text"]')
    
    # Dicionário temporário para armazenar os campos e seus respectivos labels
    field_dict = {}
    for field in fields:
        label = field.find_element(By.XPATH, '../preceding-sibling::label').text
        field_dict[label] = field
    
    # Preencher cada campo com o dado correspondente
    for label, value in field_mapping.items():
        if label in field_dict:
            field_dict[label].send_keys(str(row[value]))

    # Submeter o formulário
    submit_button = driver.find_element(By.XPATH, '//input[@type="submit"]')
    submit_button.click()

    # Esperar um tempo para que o formulário seja processado e os campos mudem de posição
    time.sleep(2)

# Fechar o navegador após completar todas as rodadas
driver.quit()