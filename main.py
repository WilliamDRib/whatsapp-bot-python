#Planilhas
import pandas as pd
from openpyxl import Workbook, load_workbook

#Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

#Extras
import urllib
import time

#Lendo planilhas
contatos = pd.read_excel("base.xlsx")
planilha = load_workbook("base.xlsx")
aba_ativa = planilha.active

#Criar link do whatsapp com numero\mensagem
def create_link_with_message(name, number, message):
    text = urllib.parse.quote(f"Mensagem Teste: Oi {name}! {message}")
    return f"https://web.whatsapp.com/send?phone={number}&text={text}"

#Loop ate carregar site(whatsapp)
def site_loaded():
    while len(navegador.find_elements(By.ID, 'side')) < 1:
        time.sleep(1)
    time.sleep(2)

#Validar se o tamanho do numero é valido
def error_number_validator():
    try:
        if navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]'):
            return True
        return False
    except NoSuchElementException:
        return False
    
#Salva status na planilha
def saving_status(aba, id, status):
    aba[f"D{id+2}"] = status

#Sair da conta
def exit_whatsapp():
    navegador.get("https://web.whatsapp.com/")
    site_loaded()
    try_click('//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[4]/div/span')
    time.sleep(2)
    try_click('//*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[4]/span/div/ul/li[6]')
    time.sleep(2)
    try_click('//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div[3]/div/div[2]')
    time.sleep(2)

#Tenta realizar um click
def try_click(css_selector):
    try:
        navegador.find_element(By.XPATH, css_selector).click()
    except NoSuchElementException:
        return False
    

#Carregando o whatsapp
navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com/")
site_loaded()

for i, mensagem in enumerate(contatos['Mensagem']):
    
    #Entrando no site
    link = create_link_with_message(contatos.loc[i, "Pessoa"], contatos.loc[i, "Número"],mensagem)
    navegador.get(link)

    site_loaded()

    #Verificando erro de link
    if error_number_validator():
        saving_status(aba_ativa, i, "Erro: Numero invalido")
        continue
    
    #Enviando
    btn = navegador.find_element(By.CSS_SELECTOR, '#main > footer > div._2lSWV._3cjY2.copyable-area > div > span:nth-child(2) > div > div._1VZX7 > div._2xy_p._3XKXx > button').send_keys(Keys.ENTER)
    saving_status(aba_ativa, i, "Enviado")

    time.sleep(1)
        
#Planilha bkp
planilha.save("BasePy.xlsx")

exit_whatsapp()