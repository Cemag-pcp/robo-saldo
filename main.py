from flask import Flask, render_template, jsonify
import threading
from flask import Flask, render_template, jsonify

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
# import chromedriver_autoinstaller
import warnings
import datetime
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from utils import *

from selenium.common.exceptions import SessionNotCreatedException

# https://googlechromelabs.github.io/chrome-for-testing/#stable
# https://github.com/GoogleChromeLabs/chrome-for-testing#json-api-endpoints

warnings.filterwarnings("ignore")

# chromedriver_autoinstaller.install()

filename = "service_account.json"

def todos_os_dias():

    from datetime import date, timedelta

    # Obtém a data atual
    data_atual = date.today()

    # Obtém o primeiro dia do mês atual
    primeiro_dia_do_mes = data_atual.replace(day=1)

    # Lista para armazenar as datas
    datas = []

    # Loop para adicionar as datas até a data atual
    while primeiro_dia_do_mes <= data_atual:
        datas.append(primeiro_dia_do_mes.strftime('%d/%m/%Y'))
        primeiro_dia_do_mes += timedelta(days=1)

    return datas

def dia_da_semana():
    
    today = datetime.datetime.now()
    today = today.isoweekday()
    return (today)

def data_sexta():
    data_sexta = datetime.datetime.now() - timedelta(3)
    ts = pd.Timestamp(data_sexta)
    data_sexta = data_sexta.strftime('%d/%m/%Y')

    return(data_sexta)

def data_sabado():
    data_sabado = datetime.datetime.now() - timedelta(2)
    ts = pd.Timestamp(data_sabado)
    data_sabado = data_sabado.strftime('%d/%m/%Y')

    return(data_sabado)

def data_ontem():
    data_ontem = datetime.datetime.now() - timedelta(1)
    ts = pd.Timestamp(data_ontem)
    data_ontem = data_ontem.strftime('%d/%m/%Y')

    return(data_ontem)

def data_antes_ontem():
    data_ontem = datetime.datetime.now() - timedelta(2)
    ts = pd.Timestamp(data_ontem)
    data_ontem = data_ontem.strftime('%d/%m/%Y')

    return(data_ontem)

def data_hoje():
    data_hoje = datetime.datetime.now()
    ts = pd.Timestamp(data_hoje)
    data_hoje = data_hoje.strftime('%d/%m/%Y')
    
    return(data_hoje)

def hora_atual():
    hora_atual = datetime.datetime.now()
    ts = pd.Timestamp(hora_atual)
    hora_atual = hora_atual.strftime('%H:%M:%S')
    
    return(hora_atual)

def mes_atual():    
    mes_atual = datetime.datetime.now().month
    
    return str(mes_atual)

def acessar_innovaro():
    
    link1 = "http://192.168.3.141/"
    #link1 = 'http://cemag.innovaro.com.br/sistema'
    #link1 = 'http://devcemag.innovaro.com.br:81/sistema'
    
    try:
        nav = webdriver.Chrome(r"C:\Users\pcp2\robo-saldo\chromedriver_extracted\chromedriver-win32\chromedriver.exe")
    except:
        chrome_driver_path = verificar_chrome_driver()
        nav = webdriver.Chrome(chrome_driver_path)

    #nav = webdriver.Chrome()
    nav.maximize_window()
    time.sleep(2)
    nav.get(link1)

    return(nav)

########### LOGIN ###########

def login(nav):
    #logando 
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("Ti.Prod")
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("Cem@@1600")

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

    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()

    time.sleep(2)

def listar(nav, classe):
    
    lista_menu = nav.find_elements(By.CLASS_NAME, classe)
    
    elementos_menu = []

    for x in range (len(lista_menu)):
        a = lista_menu[x].text
        elementos_menu.append(a)

    test_lista = pd.DataFrame(elementos_menu)
    test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

    return(lista_menu, test_lista)

def menu_apontamento(nav):

    try:
        nav.switch_to.default_content()
    except:
        pass

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Produção'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em producao
    time.sleep(0.5)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Controle de fábrica (SFC)'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em SFC
    time.sleep(0.5)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em SFC
    time.sleep(0.5)

def menu_transf(nav):

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Transferência'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em transf
    time.sleep(0.5)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Solicitação de transferência entre depósitos'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em solicitação de transf  
    time.sleep(0.5)
    
def menu_transf_2(nav):
    
    #clicando em transf
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[24]/span[2]'))).click()
    time.sleep(1.5)

    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[25]/span[2]'))).click()
    time.sleep(1.5)

def fechar_menu_consulta(nav):

    try:
        nav.switch_to.default_content()
    except:
        pass

    #fecha aba de consulta
    time.sleep(1.5)
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div"))).click()
    except:
        pass

    time.sleep(2)
    menu_innovaro(nav)
    time.sleep(2)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Consultas'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em consulta
    time.sleep(0.5)

def fechar_menu_transf(nav):
    
    try:
        nav.switch_to.default_content()
        time.sleep(2)
    except:
        pass

    #fecha aba de transf.
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div"))).click()
    
    time.sleep(1.5)
    menu_innovaro(nav)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Transferência'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em trasnferencia
    time.sleep(0.5)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Estoque'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em estoque
    time.sleep(0.5)

def fechar_menu_apont(nav):
    
    try:
        nav.switch_to.default_content()
    except:
        pass
    
    time.sleep(2)

    #fechar aba
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tabs"]/td[1]/table/tbody/tr/td[4]/span/div'))).click()
    time.sleep(2)

    menu_innovaro(nav)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Controle de fábrica (SFC)'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em Controle de fábrica (SFC)
    time.sleep(0.5)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Produção'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em Produção
    time.sleep(5)

    menu_innovaro(nav)

def fechar_estoque(nav):
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Estoque'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() 
    time.sleep(0.5)

def iframes(nav):
    
    iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try:
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass

########### ACESSANDO PLANILHAS DE TRANSFERÊNCIA ###########

def planilha_serra_transf(data, filename):

    #SERRA#

    # sheet = 'RQ PCP-001-000 (APONTAMENTO SERRA)'
    # worksheet1 = 'USO DO PCP'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    sheet_id = '1excOFHkBFOG_h5JKIEyqdBWnlemBT_nc3ey6E-cqUSE'
    worksheet1 = 'RQ PC-008-000(TRANSFERENCIA)'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(5)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1)[6:]

    ########### Tratando planilhas ###########

    #ajustando datas

    # for i in range(2, len(base)+2):
    #     if base['DATA'][i] == "":
    #         data_antes = base.at[i-1, 'DATA']
    #         base.at[i, 'DATA'] = data_antes

    #filtrando peças que não foram apontadas
    #base_filtrada = base[base['APONTAMENTO'].isnull()]
    #base_filtrada = base[base['APONTAMENTO'] == '']

    #filtrando data de hoje
    base_filtrada = base.loc[base.DATA == data]

    #filtrando por mat prima que não foi transferida
    base_filtrada = base_filtrada[(base_filtrada['PCP'].isnull()) | (base_filtrada['PCP'] == '')]

    #filtrando por mat prima que não foi transferida
    # base_filtrada = base_filtrada[base_filtrada['MAT PRIMA'] != '']
    base_filtrada = base_filtrada[(base_filtrada['MAT PRIMA'].notnull()) | (base_filtrada['MAT PRIMA'] != '')]

    base_filtrada['MAT PRIMA'] = base_filtrada['MAT PRIMA'].apply(lambda x: x.split(' ')[0])

    # base_filtrada = base_filtrada[(base_filtrada['OBSERVAÇÃO'] == '') | (base_filtrada['OBSERVAÇÃO'].isnull())]

    #filtrando data de hoje
    #base_filtrada = base_filtrada.loc[base_filtrada.TRANSFERÊNCIA == '']

    base_filtrada = base_filtrada[['DATA','MAT PRIMA','PESO']]

    # base_filtrada = base_filtrada[base_filtrada['MATERIAL'].notnull()]

    # base_filtrada = base_filtrada[base_filtrada['PESO'] != '']

    base_filtrada = base_filtrada.reset_index()

    transferidas = base_filtrada

    transferidas['PESO'] = transferidas['PESO'].apply(lambda x: float(x.replace(',','.')))

    # if len(base_filtrada) > 0:

    #     # quebrando_material = base_filtrada["MATERIAL"].str.split(" - ", n = 1, expand = True)

    #     # base_filtrada['MATERIAL'] = quebrando_material[0]
    #     base_filtrada = base_filtrada.reset_index(drop=True)

    #     for i in range(len(base_filtrada)):
    #         try:
    #             if len(base_filtrada['PESO'][i]) > 1:
    #                 base_filtrada['PESO'][i] = base_filtrada['PESO'][i].replace(',','')
    #                 base_filtrada['PESO'][i] = base_filtrada['PESO'][i].replace('.','')
    #         except:
    #             pass

    #     for j in range(len(base_filtrada)):
    #         base_filtrada['PESO'][j] = float(base_filtrada['PESO'][j]) / 100

    #     base_filtrada = base_filtrada.groupby(['index','DATA','MAT PRIMA']).sum().reset_index()

    base_filtrada = base_filtrada.loc[base_filtrada['PESO'] > 0].reset_index(drop=True)

    return(wks1, base, base_filtrada, transferidas)

def planilha_corte_transf(data, filename):

    #CORTE#

    # sheet = 'RQ PCP-012-000 (Banco de dados OP)'
    # worksheet1 = 'Transferência'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    sheet_id = '1t7Q_gwGVAEwNlwgWpLRVy-QbQo7kQ_l6QTjFjBrbWxE'
    worksheet1 = 'RQ PCP-003-000 (Transferencia)'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(5)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1)[2:]

    # base['Data'] = base['Data'].str[:10]
    #filtrando data de hoje
    base = base.loc[base.Data == data]

    ########### Tratando planilhas ###########

    #filtrando peças que não foram apontadas
    base_filtrada = base[base['Status'] == '']

    #filtrando chapas que existem apenas 1 código
    base_filtrada = base_filtrada.loc[base_filtrada['Código Chapa'] != '']

    #peso diferente de 0
    base_filtrada = base_filtrada.loc[base_filtrada['Peso'] != '0,00']

    base_filtrada = base_filtrada.reset_index()

    base_filtrada =  base_filtrada[['index','Data','Código Chapa','Peso']]

    base_filtrada['Peso'] = base_filtrada['Peso'].apply(lambda x: float(x.replace(',','.')))

    # if len(base_filtrada) > 0:

    #     for i in range(len(base)):
    #         try:
    #             if len(base_filtrada['Peso'][i]) > 1:
    #                 base_filtrada['Peso'][i] = base_filtrada['Peso'][i].replace(',','.')
    #         except:
    #             pass
        
        # base_filtrada['Peso'] = base_filtrada['Peso'].astype(float)

    base_filtrada = base_filtrada.groupby(['index','Data','Código Chapa']).sum().reset_index()

    # base_filtrada = base_filtrada.iloc[0:6]

    return(wks1, base, base_filtrada)

########### CONSULTAR SALDO ###########

def consulta_saldo(data, nav):
    
    # try:
    #     nav.switch_to.default_content()
    # except:
    #     pass

    lista_menu, test_lista = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_lista.loc[test_lista[0] == 'Estoque'].index[0]
    
    lista_menu[click_producao+1].click() ##clicando em estoque
    time.sleep(1)

    lista_menu, test_lista = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_lista.loc[test_lista[0] == 'Consultas'].index[0]
    
    lista_menu[click_producao+1].click() ##clicando em consulta
    time.sleep(1)

    lista_menu, test_lista = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_lista.loc[test_lista[0] == 'Saldos de recursos'].index[0]
    
    lista_menu[click_producao+1].click() ##clicando em apontamento
    time.sleep(1)

    iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try:
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass

    time.sleep(1.5)

    carregou = 0
    
    while carregou == 1:
        if WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))):
            carregou = 1
    
    #data base
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).click()

    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    time.sleep(3)
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
    time.sleep(3)
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(data)
    time.sleep(3)
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)
    time.sleep(3)
    
    #recursos
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).click()
    time.sleep(2)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)

    wks1, base, base_filtrada, transferidas = planilha_serra_transf(data, filename)

    try:
        if len(base_filtrada)>0:

            base_filtrada = base_filtrada.reset_index(drop=True)

            qtde_itens = len(base_filtrada.drop_duplicates(subset=['MAT PRIMA']))

            for i in range(len(base_filtrada)):
                recurso = base_filtrada['MAT PRIMA'][i]
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(recurso + ';')
            
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)
            time.sleep(1)

            try:
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]/span[2]/p'))).click()
            except:
                pass
    
            iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

            for iframe in range(len(iframe_list)):
                time.sleep(1)
                try:
                    nav.switch_to.default_content()
                    nav.switch_to.frame(iframe_list[iframe])
                    print(iframe)
                except:
                    pass

            #Botão de executar
            time.sleep(2)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).click()
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + Keys.SHIFT + "E")    

            try:
                while WebDriverWait(nav, 2).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[9]/table/tbody/tr/td[2]/div'))):
                    print("Carregando")
            except:
                print("Carregou")  
            
            #mudando iframe
            iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

            for iframe in range(len(iframe_list)):
                time.sleep(1)
                try:
                    nav.switch_to.default_content()
                    nav.switch_to.frame(iframe_list[iframe])
                    print(iframe)
                except:
                    pass

            table_prod = WebDriverWait(nav, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table')))
            table_html_prod = table_prod.get_attribute('outerHTML')
                
            tabelona = pd.read_html(str(table_html_prod), header=None)
            tabelona = tabelona[0]
            tabelona = tabelona.droplevel(level=0,axis=1)
            tabelona = tabelona.droplevel(level=0,axis=1)

            tabelona = tabelona[['Unid. Medida','Saldo']]
            #tabelona = tabelona[['Unnamed: 0_level_2','Saldo']]
            tabelona['Saldo'] = tabelona.Saldo.shift(-1)
            tabelona = tabelona.dropna()
            tabelona = tabelona.reset_index(drop=True)

            tabelona['Saldo'] = tabelona['Saldo'].apply(lambda x: x.replace('.','').replace(',',''))
            tabelona['Saldo'] = pd.to_numeric(tabelona['Saldo'], errors='coerce') / 10000
            tabelona = tabelona.dropna()
            # tabelona = tabelona.iloc[:-3]

            # if qtde_itens == 1:
            #     tabelona = tabelona[:1]
            # else:
            #     tabelona = tabelona[:len(tabelona)-2]

            #quebrando_material = tabelona["Unnamed: 0_level_2"].str.split(" ", n = 1, expand = True)

            #tabelona['Unnamed: 0_level_2'] = quebrando_material[0]

            # for i in range(len(tabelona)):
            #     if len(tabelona['Saldo'][i]) > 6 :
            #         tabelona['Saldo'][i] = tabelona['Saldo'][i].replace(',','')
            #         tabelona['Saldo'][i] = tabelona['Saldo'][i].replace('.','')

            # for saldo in range(len(tabelona)):
            #     try:
            #         tabelona['Saldo'][saldo] = tabelona['Saldo'][saldo][:len(tabelona['Saldo'][saldo])-4] + "." + tabelona['Saldo'][saldo][-4:]  
            #     except:
            #         pass

            # try:
            #     for j in range(len(tabelona)):
            #         if tabelona['Saldo'][j][:1] == '0' :
            #             tabelona['Saldo'][j] = tabelona['Saldo'][j][:1] + '.' + tabelona['Saldo'][j][1:3]

            # except:
            #     pass


            # try:
            #     for j in range(len(tabelona)):
            #         if len(tabelona['Saldo'][j]) >= 6 :
            #             tabelona['Saldo'][j] = float(tabelona['Saldo'][j]) / 10000

            # except:
            #     pass

            # tabelona['Saldo'] = tabelona['Saldo'].astype(float)

            #tabelona = tabelona.rename(columns={'Unnamed: 0_level_2':'MATERIAL'})
            tabelona = tabelona.rename(columns={'Unid. Medida':'MAT PRIMA'})
            tabelona['MAT PRIMA'] = tabelona['MAT PRIMA'].apply(lambda x: x.split(' ')[0])

            # for i in range(len(tabelona)):
            #     tabelona['MATERIAL'][i] = tabelona['MATERIAL'][i][:len(tabelona['MATERIAL'][i]) - 5]
            
            df_final = pd.merge(tabelona,base_filtrada,on='MAT PRIMA')

            lista_material = df_final['MAT PRIMA'].unique()           

            df_final['saldo2'] = ''
            
            df_final1 = pd.DataFrame()
            
            for material in lista_material:
                df_filtro_material = df_final[df_final['MAT PRIMA'] == material].reset_index(drop=True)
                
                for i in range(len(df_filtro_material)):
                    try:
                        df_filtro_material['saldo2'][i] = df_filtro_material['saldo2'][i-1] - df_filtro_material['PESO'][i]
                    except:
                        df_filtro_material['saldo2'][i] = df_filtro_material['Saldo'][i] - df_filtro_material['PESO'][i]
                    
                df_final1 = pd.concat([df_final1,df_filtro_material])

            df_final = df_final1

            df_final['comparar'] = df_final['saldo2'] >= 0 

            df_final = df_final.loc[df_final['comparar'] == True].reset_index(drop=True)

        else:
            
            df_final = pd.DataFrame()
    
    except:
        df_final = pd.DataFrame()

    return(df_final)

def consulta_saldo_chapas(data, nav):

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Transferência'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em transf
    time.sleep(0.5)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Consultas'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em consultas
    time.sleep(0.5)
    
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Saldos de recursos'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em consulta
    time.sleep(3)

    iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try:
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass
    
    #data base
    time.sleep(3)
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    time.sleep(3)
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
    time.sleep(3)
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(data)
    time.sleep(3)
    WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)
    time.sleep(3)
    
    #recursos
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).click()
    time.sleep(2)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)

    wks1, base, base_filtrada = planilha_corte_transf(data, filename)

    try:
        if len(base_filtrada)>0:

            base_filtrada = base_filtrada.reset_index(drop=True)

            qtde_itens = len(base_filtrada)

            for i in range(len(base_filtrada)):
                recurso = base_filtrada['Código Chapa'][i]
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(recurso + ';')
            
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)
            time.sleep(1)

            try:
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]/span[2]/p'))).click()
            except:
                pass

            iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

            for iframe in range(len(iframe_list)):
                time.sleep(1)
                try:
                    nav.switch_to.default_content()
                    nav.switch_to.frame(iframe_list[iframe])
                    print(iframe)
                except:
                    pass

            #Botão de executar
            time.sleep(2)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).click()
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + Keys.SHIFT + "E")

            #esperando carregar
            try:
                while WebDriverWait(nav, 2).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[9]/table/tbody/tr/td[2]/div'))):
                    print("Carregando")
            except:
                print("Carregou")           
                                                                                            
            #mudando iframe
            iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

            for iframe in range(len(iframe_list)):
                time.sleep(1)
                try:
                    nav.switch_to.default_content()
                    nav.switch_to.frame(iframe_list[iframe])
                    print(iframe)
                except:
                    pass

            table_prod = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table')))
            table_html_prod = table_prod.get_attribute('outerHTML')
                
            tabelona = pd.read_html(str(table_html_prod), header=None)
            tabelona = tabelona[0]
            tabelona = tabelona.droplevel(level=0,axis=1)
            tabelona = tabelona.droplevel(level=0,axis=1)

            tabelona = tabelona[['Unid. Medida','Saldo']]
            #tabelona = tabelona[['Unnamed: 0_level_2','Saldo']]
            tabelona['Saldo'] = tabelona.Saldo.shift(-1)
            tabelona = tabelona.dropna()
            tabelona = tabelona.reset_index(drop=True)

            tabelona['Saldo'] = tabelona['Saldo'].apply(lambda x: x.replace('.','').replace(',',''))
            tabelona['Saldo'] = pd.to_numeric(tabelona['Saldo'], errors='coerce') / 10000
            tabelona = tabelona.dropna()

            #tabelona = tabelona[['Unnamed: 0_level_2','Saldo']]
            # tabelona['Saldo'] = tabelona.Saldo.shift(-2)
            # tabelona = tabelona.dropna()
            # tabelona = tabelona.reset_index(drop=True)
            
            # if qtde_itens == 1:
            #     tabelona = tabelona[:1]
            # else:
            #     tabelona = tabelona[:len(tabelona)-2]

            #quebrando_material = tabelona["Unnamed: 0_level_2"].str.split(" ", n = 1, expand = True)

            #tabelona['Unnamed: 0_level_2'] = quebrando_material[0]

            # for i in range(len(tabelona)):
            #     if len(tabelona['Saldo'][i]) > 6 :
            #         tabelona['Saldo'][i] = tabelona['Saldo'][i].replace(',','')
            #         tabelona['Saldo'][i] = tabelona['Saldo'][i].replace('.','')

            # try:
            #     for j in range(len(tabelona)):
            #         if len(tabelona['Saldo'][j]) >= 6 :
            #             tabelona['Saldo'][j] = float(tabelona['Saldo'][j]) / 10000
            # except:
            #     pass

            # tabelona['Saldo'] = tabelona['Saldo'].astype(float)

            #tabelona = tabelona.rename(columns={'Unnamed: 0_level_2':'MATERIAL'})
            tabelona = tabelona.rename(columns={'Unid. Medida':'Código Chapa'})
            
            tabelona['Código Chapa'] = tabelona['Código Chapa'].apply(lambda x: x.split(' ')[0])

            # for i in range(len(tabelona)):
            #     tabelona['Código Chapa'][i] = tabelona['Código Chapa'][i][:len(tabelona['Código Chapa'][i]) - 5]

            df_final = pd.merge(tabelona,base_filtrada,on='Código Chapa')
            
            lista_material = df_final['Código Chapa'].unique()           

            df_final['saldo2'] = ''

            df_final1 = pd.DataFrame()

            for material in lista_material:
                df_filtro_material = df_final[df_final['Código Chapa'] == material].reset_index(drop=True)
                
                for i in range(len(df_filtro_material)):
                    try:
                        df_filtro_material['saldo2'][i] = df_filtro_material['saldo2'][i-1] - df_filtro_material['Peso'][i]
                    except:
                        df_filtro_material['saldo2'][i] = df_filtro_material['Saldo'][i] - df_filtro_material['Peso'][i]

                df_final1 = pd.concat([df_final1,df_filtro_material])

            df_final = df_final1

            df_final['comparar'] = df_final['saldo2'] >= 0 

            df_final = df_final.loc[df_final['comparar'] == True].reset_index(drop=True)

        else:
            
            df_final = pd.DataFrame()
    except:
        df_final = pd.DataFrame()

    return(df_final)

def fechar_tabs(nav):

    nav.switch_to.default_content()

    try:
        tab1 = nav.find_elements(By.CLASS_NAME, 'process-tab-right-active') #listar abas ativas (aba que está selecinada)
        tab2 = nav.find_elements(By.CLASS_NAME, 'process-tab-right-inactive') #listar abas inativas

        for apagar in range(len(tab2)):
            tab2[apagar].click()
            time.sleep(1)

        for apagar in range(len(tab1)):
            tab1[apagar].click()
            time.sleep(1)

    except:
        print("nenhuma aba aberta")

########## VERIFICAR CHECKBOX #########

def checkbox_apontamentos(filename):

    sheet = 'CENTRAL DE APONTAMENTO'
    worksheet1 = 'PAINEL'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks2 = sh.worksheet(worksheet1)

    base = wks2.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,16:]
    base = base.iloc[6:12,0:2]
    base = base.set_axis(['Setor','Ativador'], axis=1)
    base = base[base['Ativador'] == 'TRUE']

    lista_checkbox = base

    return lista_checkbox

def fechar_tabs(nav):

    nav.switch_to.default_content()

    try:
        tab1 = nav.find_elements(By.CLASS_NAME, 'process-tab-right-active') #listar abas ativas (aba que está selecinada)
        tab2 = nav.find_elements(By.CLASS_NAME, 'process-tab-right-inactive') #listar abas inativas

        for apagar in range(len(tab2)):
            tab2[apagar].click()
            time.sleep(1)

        for apagar in range(len(tab1)):
            tab1[apagar].click()
            time.sleep(1)

    except:
        print("nenhuma aba aberta")

########## VERIFICAR CHECKBOX #########

def checkbox_apontamentos(filename):

    sheet = 'CENTRAL DE APONTAMENTO'
    worksheet1 = 'PAINEL'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks2 = sh.worksheet(worksheet1)

    base = wks2.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,16:]
    base = base.iloc[6:12,0:2]
    base = base.set_axis(['Setor','Ativador'], axis=1)
    base = base[base['Ativador'] == 'TRUE']

    lista_checkbox = base

    return lista_checkbox

def fechar_tabs(nav):

    nav.switch_to.default_content()

    try:
        tab1 = nav.find_elements(By.CLASS_NAME, 'process-tab-right-active') #listar abas ativas (aba que está selecinada)
        tab2 = nav.find_elements(By.CLASS_NAME, 'process-tab-right-inactive') #listar abas inativas

        for apagar in range(len(tab2)):
            tab2[apagar].click()
            time.sleep(1)

        for apagar in range(len(tab1)):
            tab1[apagar].click()
            time.sleep(1)

    except:
        print("nenhuma aba aberta")

########## VERIFICAR CHECKBOX #########

def checkbox_apontamentos(filename):

    sheet = 'CENTRAL DE APONTAMENTO'
    worksheet1 = 'PAINEL'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks2 = sh.worksheet(worksheet1)

    base = wks2.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,16:]
    base = base.iloc[6:12,0:2]
    base = base.set_axis(['Setor','Ativador'], axis=1)
    base = base[base['Ativador'] == 'TRUE']

    lista_checkbox = base

    return lista_checkbox

########## LOOP ###########

##### onde o robô ta? #####

# sheet = 'CENTRAL DE APONTAMENTO'
# worksheet1 = 'PAINEL'

# sa = gspread.service_account(filename)
# sh = sa.open(sheet)

# wks2 = sh.worksheet(worksheet1)

#data_dia_1 = '12/07/2023'

while True:

    try:

        today = dia_da_semana()

        if today != 1:

            datas = [data_hoje(),  data_ontem()]#, data_sabado()]
            #datas = [data_hoje()]#, data_sabado()]

        else:

            datas = [data_hoje(), data_ontem(), data_sabado(), data_sexta()]

        #datas = ['25/04/2023','26/04/2023','27/04/2023','28/04/2023','29/04/2023','30/04/2023']

        agora = datetime.datetime.now().time()
        inicio = datetime.time(7, 0, 0)  # 8 da manhã
        fim = datetime.time(19, 0, 0)    # 17 da tarde

        if inicio <= agora <= fim:

            while True:

                print("Indo para saldo mp")
                nav = acessar_innovaro()

                #download saldo de recurso para central apenas materia prima
                login(nav)

                # Realizar download saldo de recurso
                iframes(nav)
                menu_innovaro(nav)
                time.sleep(1)

                navegar_consulta(nav)
                time.sleep(1)

                iframes(nav)
                time.sleep(1)

                input_data(nav)
                time.sleep(1)
                
                input_deposito_apenas_central(nav)
                time.sleep(1)

                limpar_recursos(nav)
                time.sleep(1)

                inserir_agrupamentos_central_mp(nav)
                time.sleep(1)

                inserir_agrupamentos_almox(nav)
                time.sleep(1)

                exportar_2(nav)
                time.sleep(10)

                inserir_gspread_saldo_central_mp()
                time.sleep(1)

                fechar_menu_consulta(nav)
                fechar_estoque(nav)
                fechar_tabs(nav)

                nav.close()

                # print("indo para saldo eric")
                # nav = acessar_innovaro()

                # time.sleep(4)

                # login(nav)

                # # Realizar download saldo de recurso
                # iframes(nav)
                # menu_innovaro(nav)
                # time.sleep(1)

                # navegar_consulta(nav)
                # time.sleep(1)

                # iframes(nav)
                # time.sleep(1)
                
                # input_data(nav)
                # time.sleep(1)

                # input_deposito(nav)
                # time.sleep(1)

                # apagar_mat_prima(nav)
                # time.sleep(1)

                # limpar_recursos(nav)
                # time.sleep(1)
                
                # inserir_agrupamentos_almox(nav)
                # time.sleep(1)
                
                # inputar_recurso(nav)
                # time.sleep(5)
                
                # inserir_gspread()
                # time.sleep(5)

                # fechar_menu_consulta(nav)
                # fechar_estoque(nav)
                # fechar_tabs(nav)
                
                # nav.close()

                # time.sleep(3)

                print("Indo para saldo levantamento")
                nav = acessar_innovaro()

                #download saldo de recurso para levantamento
                login(nav)

                # Realizar download saldo de recurso
                iframes(nav)
                menu_innovaro(nav)

                time.sleep(1)
                
                navegar_consulta(nav)
                time.sleep(1)

                iframes(nav)
                time.sleep(1)
                
                input_deposito_levantamento(nav)
                time.sleep(1)

                limpar_recursos(nav)
                time.sleep(1)

                apagar_mat_prima(nav)
                time.sleep(1)

                inserir_agrupamentos_levantamento(nav)
                time.sleep(1)

                exportar_2(nav)
                time.sleep(1)

                inserir_gspread_saldo_levantamento()
                time.sleep(1)
                
                fechar_menu_consulta(nav)
                fechar_estoque(nav)
                fechar_tabs(nav)

                nav.close()

                time.sleep(3)
                


                time.sleep(1800)

    except:
        nav.close()


