# -*- coding: utf-8 -*-
import time
import openpyxl
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

try:
    from itertools import izip
except ImportError:  # python3.x
    izip = zip

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
gommobile = 'http://gomnet.ampla.com/Mobile/telaPrincipal.aspx'
username = login['A1'].value
password = login['A2'].value
mobUser = login['B1'].value
mobPass = login['B2'].value
wb = openpyxl.load_workbook('sobs.xlsx')
wb1 = openpyxl.load_workbook('sobs.xlsx')

driver = webdriver.Chrome()

if __name__ == '__main__':
    driver.get(url)
    # Faz login no sistema GOMNET
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()
    driver.get(gommobile)

    # Faz login no sistema GOMMOBILE
    mobLogin = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtLogin')
    mobLogin.send_keys(mobUser)
    mobSenha = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtSenha')
    mobSenha.send_keys(mobPass)
    entrarBtn = driver.find_element_by_id('ctl00_ContentPlaceHolder1_btnEnvia').click()

    # Sincroniza as tarefas
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_lbSincronizar'))).click()
    okBtn = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[3]/div/button/span'))).click()
    # Clica nas tarefas pendentes
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_lbTarefaPendente'))).click()

    # Acessa os dados na planilha 'sobs.xlsx' para começar a trabalhar.
    for sheet in wb.worksheets:
        try:
            # Busca o valor da SOB e clica
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "*//tr/td[contains(text(), '" + str(sheet['A1'].value) + "')]"))).click()
            # Pressina o TAB uma vez e depois ENTER, para abrir a janela de inserção de dados para a SOB.
            webdriver.ActionChains(driver).send_keys(Keys.TAB).perform()
            webdriver.ActionChains(driver).send_keys(Keys.RETURN).perform()
            # Seleciona todos os baremos executados
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="gvItens"]/tbody/tr[1]/th[1]'))).click()
            # Clica na aba "Geral"
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tabGeral"]'))).click()
            # Clica na caixa de informação de "Saída"
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtDtSaida"]'))).click()
            # Clica no botão "agora"
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ui-datepicker-div"]/div[3]/button[1]'))).click()
            # Clica na caixa de informação de "Início Tarefa"
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtIniTarefa"]'))).click()
            # Clica no botão "agora"
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ui-datepicker-div"]/div[3]/button[1]'))).click()
            try: # Verifica se a célula B1 da planilha 'sobs.xlsx' consta um 'X' para energizar a SOB. Caso não tenha, finaliza parcialmente.
                if str(sheet['B1'].value) == 'X':
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtEnergizacao"]'))).click()
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ui-datepicker-div"]/div[3]/button[1]'))).click()
            except NoSuchElementException:
                continue
            # Finaliza a SOB
            # WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/div/button[1]/span'))).click()
            time.sleep(5)
            # Cancela a SOB
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/div/button[2]/span'))).click()
        except NoSuchElementException:  # Caso não encontre, abre o arquivo txt e registra o número da SOB não movimentada.
            log = open("ErroSobs.txt", "a")
            log.write(str(sheet['A1'].value) + "\n")
            log.close()
            continue