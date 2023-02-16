from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from datetime import datetime,timedelta 
import numpy as np
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import gspread

filename = "service_account.json"

def data_ontem():
    data_ontem = datetime.now() - timedelta(1)
    ts = pd.Timestamp(data_ontem)
    data_ontem = data_ontem.strftime('%d/%m/%Y')

    return(data_ontem)

def data_hoje():
    data_hoje = datetime.now()
    ts = pd.Timestamp(data_hoje)
    data_hoje = data_hoje.strftime('%d/%m/%Y')
    
    return(data_hoje)

def hora_atual():
    hora_atual = datetime.now()
    ts = pd.Timestamp(hora_atual)
    hora_atual = hora_atual.strftime('%H:%M:%S')
    
    return(hora_atual)

def acessar_innovaro():
    
    #options = webdriver.ChromeOptions()
    #options.add_argument("--headless")
    
    #link1 = "http://192.168.3.141/"
    link1 = 'http://cemag.innovaro.com.br/sistema'
    #link1 = 'http://devcemag.innovaro.com.br:81/sistema'
    nav = webdriver.Chrome()#chrome_options=options)
    time.sleep(2)
    nav.get(link1)

    return(nav)

def listar(nav, classe):
    lista_menu = nav.find_elements(By.CLASS_NAME, classe)
    
    elementos_menu = []

    for x in range (len(lista_menu)):
        a = lista_menu[x].text
        elementos_menu.append(a)

    test_lista = pd.DataFrame(elementos_menu)
    test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

    return(lista_menu, test_lista)

########### LOGIN ###########

def login(nav):
    #logando 
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("Trainee - PCP")
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("cem@1605")

    time.sleep(2)

    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

    time.sleep(2)

########### MENU ###########

def menu_innovaro(nav):
    #abrindo menu

    try:
        nav.switch_to.default_content()
    except:
        pass

    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/table/tbody/tr/td[5]/div/table/tbody/tr/td[2]'))).click()
    time.sleep(2)

def menu_projeto(nav): #clicando em projeto e materiais e produtos

    try:
        nav.switch_to.default_content()
    except:
        pass

    #clicando em projeto
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Projeto'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em producao
    time.sleep(0.5)

def materiais(nav):
    #clicando em materiais e produtos

    try:
        nav.switch_to.default_content()
    except:
        pass

    #clicando em projeto
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Materiais e Produtos'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em producao
    time.sleep(0.5)

def expandir_1(peca, nav): #expande e procura a peça
    
    nav.switch_to.default_content()
    time.sleep(2)

    try:
        #mudando iframe
        iframe1 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        time.sleep(2)
        nav.switch_to.frame(iframe1)
    except:
        pass

    try:
        #mudando iframe
        iframe1 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[2]/div[2]/iframe')))
        time.sleep(2)
        nav.switch_to.frame(iframe1)
    except:
        pass


    time.sleep(2)

    #clicando em expandir
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div'))).click()

    #clicando na lupa
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div'))).click()

    #clicando no campo de código
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).click()

    #clicando em localizar
    time.sleep(0.5)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input'))).click()
    time.sleep(0.5)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input'))).send_keys(Keys.CONTROL + 'a')
    time.sleep(0.5)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input'))).send_keys(Keys.DELETE)

    #inserir código no campo localizar
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input'))).send_keys(peca)
    time.sleep(0.5)

    #clicando em próxima
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grSearchDoSearch_explorer"]'))).click()
    
    try:
        nav.switch_to.default_content()
        time.sleep(0.5)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button'))).click()
    except:
        pass

    try:
        #mudando iframe
        iframe1 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
    except:
        pass
        
def expandir_2(j, chapa, nav):
    
    nav.switch_to.default_content()

    try:
        #mudando iframe
        iframe1 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        time.sleep(2)
        nav.switch_to.frame(iframe1)
    except:
        pass

    try:
        #mudando iframe
        iframe1 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[2]/div[2]/iframe')))
        time.sleep(2)
        nav.switch_to.frame(iframe1)
    except:
        pass

    #clicando em expandir
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div'))).click()
    time.sleep(1.5)

    #coletando variáveis linha 1
    peca_atual = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[3]/div/div'))).text
    carac = peca_atual.find(' ')
    peca_atual = peca_atual[:carac]
    peso_atual = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[5]/div/div'))).text
    datai1 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[9]/div/div'))).text
    dataf1 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[10]/div/div'))).text

    #coletando variáveis linha 2
    try:
        peca_atual2 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[3]/div/div'))).text
        carac2 = peca_atual2.find(' ')
        peca_atual2 = peca_atual2[:carac2]
        peso_atual2 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[5]/div/div'))).text
        datai2 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[9]/div/div'))).text
        dataf2 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[10]/div/div'))).text
    except:
        peca_atual2 = ''

    #coletando variáveis linha 3
    try:
        peca_atual3 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[3]/div/div'))).text
        carac3 = peca_atual3.find(' ')
        peca_atual3 = peca_atual3[:carac3]
        peso_atual3 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[5]/div/div'))).text
    except:
        peca_atual3 = ''


    if peca_atual2 == '':

        if peca_atual != chapa:
            
            #data de ontem
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[10]/div/div'))).click()
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[10]/div/input'))).send_keys(data_ontem())
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[10]/div/input'))).send_keys(Keys.TAB)
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]'))).click()
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()
            time.sleep(2)

            #tab no índice
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[2]/div/input'))).send_keys('2')
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[2]/div/input'))).send_keys(Keys.TAB)
            time.sleep(2)
            
            #input recurso
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[3]/div/input'))).send_keys(chapa)
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[3]/div/input'))).send_keys(Keys.TAB)
            time.sleep(3)

            #peso atual
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[5]/div/input'))).send_keys(peso_atual)
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[5]/div/input'))).send_keys(Keys.TAB)
            time.sleep(3)

            #tab um para deposito
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[6]/div/input'))).send_keys(Keys.TAB)

            #preenchendo corte
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[7]/div/input'))).send_keys('Almox Corte e Estamparia')
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[7]/div/input'))).send_keys(Keys.TAB)
            time.sleep(3)

            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[9]/div/input'))).send_keys(data_hoje())
            time.sleep(3)

            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            time.sleep(3)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            time.sleep(3)

            print("Cadastro atualizado")
            wks1.update("c" + str(j + 1), 'OK ROBOSEXTA - ' + data_hoje() + ' ' + hora_atual())

        else:
            time.sleep(1.5)
            print("Peça igual ao cadastro")
            wks1.update("c" + str(j + 1), 'Peça igual ao cadastro - ' + data_hoje() + ' ' + hora_atual())

    if peca_atual3 == '' and peca_atual2 != '':

        if peca_atual != chapa:
            
            #data de ontem
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[10]/div/div'))).click()
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[10]/div/input'))).send_keys(data_ontem())
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[10]/div/input'))).send_keys(Keys.TAB)
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]'))).click()
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            time.sleep(4)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()
            time.sleep(2)

            #tab no índice
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[2]/div/input'))).send_keys('2')
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[2]/div/input'))).send_keys(Keys.TAB)
            time.sleep(2)
            
            #input recurso
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[3]/div/input'))).send_keys(chapa)
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[3]/div/input'))).send_keys(Keys.TAB)
            time.sleep(3)

            #peso atual
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[5]/div/input'))).send_keys(peso_atual)
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[5]/div/input'))).send_keys(Keys.TAB)
            time.sleep(3)

            #tab um para deposito
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[6]/div/input'))).send_keys(Keys.TAB)

            #preenchendo corte
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[7]/div/input'))).send_keys('Almox Corte e Estamparia')
            time.sleep(2)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[7]/div/input'))).send_keys(Keys.TAB)
            time.sleep(3)

            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[9]/div/input'))).send_keys(data_hoje())
            time.sleep(3)

            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            time.sleep(3)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            time.sleep(3)

            print("Cadastro atualizado")
            wks1.update("c" + str(j + 1), 'OK ROBOSEXTA - ' + data_hoje() + ' ' + hora_atual())

        else:
            time.sleep(1.5)
            print("Peça igual ao cadastro")
            wks1.update("c" + str(j + 1), 'Peça igual ao cadastro - ' + data_hoje() + ' ' + hora_atual())


    time.sleep(1.5)

def planilha(filename):

    sheet = 'Mudança chapa'
    worksheet1 = 'Cópia de para o robson'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(1)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1, inplace=False)[1:]

    ########### Tratando planilhas ###########

    base['Peça'] = base['Peça'].replace('-','_', regex=True)
    base['Peça'] = base['Peça'].replace('-','_', regex=True)
    base['Peça'] = base['Peça'].replace(' ','_', regex=True)
    base['n_underscore'] = base['Peça'].str.find('_')
    base['ponto'] = base['Peça'].str.find('.')
    base['peça_tratada'] = ''

    for i in range(len(base)+5):
        try:
            if base['n_underscore'][i] > 0:

                try:
                    base['peça_tratada'][i] = base['Peça'][i][:base['n_underscore'][i]]
                except:
                    pass
            
            else:

                try:
                    base['peça_tratada'][i] = base['Peça'][i]
                except:
                    pass
        except:
            pass
    
    base = base.fillna('')

    base = base.loc[(base['codigo'] != 'ok')]
    base = base.loc[(base['OK'] == '')]

    return(base, wks1)
    
########## LOOP ###########

nav = acessar_innovaro()

time.sleep(4)

login(nav)
                                   
menu_innovaro(nav)

menu_projeto(nav)

menu_innovaro(nav)

df_final, wks1 = planilha(filename)

df_final = df_final.reset_index()

for i in range(len(df_final)):
    print(i)

    try:
        peca = df_final['peça_tratada'][i]
        chapa = str(df_final['codigo'][i]) 
        j = df_final['index'][i]
        time.sleep(2)
        menu_innovaro(nav)
        print(i)
        time.sleep(2)
        materiais(nav)
        print("ok")
        print("ok2")
        time.sleep(4)
        expandir_1(peca, nav)
        time.sleep(2)
        expandir_2(j,chapa,nav) 
        time.sleep(1)
        nav.switch_to.default_content()

        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()
        
    except:

        try:

            nav.switch_to.default_content()

            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()
        except:
            pass

           