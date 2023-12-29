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
# https://googlechromelabs.github.io/chrome-for-testing/#stable
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
    nav = webdriver.Chrome(r"C:\Users\Engine\chromedriver.exe")
    nav.maximize_window()
    time.sleep(2)
    nav.get(link1)

    return(nav)

########### LOGIN ###########

def login(nav):
    #logando 
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("luan araujo")
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("luanaraujo4")

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

    sheet_id = '1Kiupceopsz-HHZ9TJgWhxADNtFHVptLszsDeJyxkiYk'
    worksheet1 = 'RQ PCP-001-001'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(6)

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
    base_filtrada = base_filtrada[(base_filtrada['TRANSF'].isnull()) | (base_filtrada['TRANSF'] == '')]

    #filtrando por mat prima que não foi transferida
    base_filtrada = base_filtrada[base_filtrada['CÓDIGO'] != '']

    base_filtrada = base_filtrada[(base_filtrada['OBSERVAÇÃO'] == '') | (base_filtrada['OBSERVAÇÃO'].isnull())]

    #filtrando data de hoje
    #base_filtrada = base_filtrada.loc[base_filtrada.TRANSFERÊNCIA == '']

    base_filtrada =  base_filtrada[['DATA','MATERIAL','PESO BARRAS']]

    base_filtrada = base_filtrada[base_filtrada['MATERIAL'].notnull()]

    base_filtrada = base_filtrada[base_filtrada['PESO BARRAS'] != '']

    base_filtrada = base_filtrada.reset_index()

    transferidas = base_filtrada

    if len(base_filtrada) > 0:

        quebrando_material = base_filtrada["MATERIAL"].str.split(" - ", n = 1, expand = True)

        base_filtrada['MATERIAL'] = quebrando_material[0]
        base_filtrada = base_filtrada.reset_index(drop=True)

        for i in range(len(base_filtrada)):
            try:
                if len(base_filtrada['PESO BARRAS'][i]) > 1:
                    base_filtrada['PESO BARRAS'][i] = base_filtrada['PESO BARRAS'][i].replace(',','')
                    base_filtrada['PESO BARRAS'][i] = base_filtrada['PESO BARRAS'][i].replace('.','')
            except:
                pass

        for j in range(len(base_filtrada)):
            base_filtrada['PESO BARRAS'][j] = float(base_filtrada['PESO BARRAS'][j]) / 10

        base_filtrada = base_filtrada.groupby(['index','DATA','MATERIAL']).sum().reset_index()

    base_filtrada = base_filtrada.loc[base_filtrada['PESO BARRAS'] > 0].reset_index(drop=True)

    return(wks1, base, base_filtrada, transferidas)

def planilha_corte_transf(data, filename):

    #CORTE#

    # sheet = 'RQ PCP-012-000 (Banco de dados OP)'
    # worksheet1 = 'Transferência'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    sheet_id = '1t7Q_gwGVAEwNlwgWpLRVy-QbQo7kQ_l6QTjFjBrbWxE'
    worksheet1 = 'RQ PCP-015-000'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(5)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1)[2:]

    base['Data'] = base['Data'].str[:10]

    ########### Tratando planilhas ###########

    #filtrando peças que não foram apontadas
    base_filtrada = base[base['Status'] == '']

    #filtrando data de hoje
    base_filtrada = base_filtrada.loc[base_filtrada.Data == data]

    #filtrando chapas que existem apenas 1 código
    base_filtrada = base_filtrada.loc[base_filtrada['Código Chapa'] != '']

    #peso diferente de 0
    base_filtrada = base_filtrada.loc[base_filtrada['Peso'] != '0,00']

    base_filtrada = base_filtrada.reset_index()

    base_filtrada =  base_filtrada[['index','Data','Código Chapa','Peso']]

    if len(base_filtrada) > 0:

        for i in range(len(base)):
            try:
                if len(base_filtrada['Peso'][i]) > 1:
                    base_filtrada['Peso'][i] = base_filtrada['Peso'][i].replace(',','.')
            except:
                pass
        
        base_filtrada['Peso'] = base_filtrada['Peso'].astype(float)

        base_filtrada = base_filtrada.groupby(['index','Data','Código Chapa']).sum().reset_index()

    base_filtrada = base_filtrada.iloc[0:6]

    return(wks1, base, base_filtrada)

########### ACESSANDO PLANILHAS DE APONTAMENTO ###########

def planilha_serra(filename, data):

    #SERRA#

    sheet_id = '1Kiupceopsz-HHZ9TJgWhxADNtFHVptLszsDeJyxkiYk'
    worksheet1 = 'RQ PCP-001-001'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    # sheet = 'RQ PCP-001-000 (APONTAMENTO SERRA)'
    # worksheet1 = 'USO DO PCP'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(6)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1)[6:]

    ########### Tratando planilhas ###########

    #ajustando datas
    # for i in range(len(base)+2):
    #     try:
    #         if base['DATA'][i] == "":
    #             data_antes = base.at[i-1, 'DATA']
    #             base.at[i, 'DATA'] = data_antes
    #     except:
    #         pass

    i = None

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['APONTAMENTO'].isnull()]

    base_filtrada = base[(base['APONTAMENTO'] == '') | (base['APONTAMENTO'].isna())]
    base_filtrada = base_filtrada[base_filtrada['DATA'] != '']
    base_filtrada = base_filtrada[base_filtrada['OBSERVAÇÃO'] == '']

    # base_filtrada['mes_data'] = ''
    # base_filtrada['dia'] = ''
    
    # for i in range(base.index[-1]+1):
    #     try:
    #         base_filtrada['mes_data'][i] = base_filtrada['DATA'][i][3:5]
    #         base_filtrada['dia'][i] = base_filtrada['DATA'][i][0:2]
    #     except:
    #         pass
        
    # mes = mes_atual()

    # if len(mes) == 1:
    #     mes = '0' + mes

    base_filtrada = base_filtrada[base_filtrada['DATA'] == data]
    # base_filtrada['dia'] = base_filtrada['dia'].astype(int)
    # base_filtrada = base_filtrada.sort_values(by='dia', ascending=False)

    base_filtrada = base_filtrada.fillna('')

    #filtrando peças que faltam apontar
    base_filtrada = base_filtrada.loc[base_filtrada.TRANSF != '']

    #inserindo 0 antes do código da peca
    base_filtrada['CÓDIGO'] = base_filtrada['CÓDIGO'].astype(str)

    for i in range(len(base)):
        try:
            if len(base_filtrada['CÓDIGO'][i]) == 5:
                base_filtrada['CÓDIGO'][i] = "0" + base_filtrada['CÓDIGO'][i] 
        except:
            pass

    i = None

    base_filtrada = base_filtrada[['DATA','CÓDIGO','QNT', 'APONTAMENTO', 'PEÇAS MORTA']]

    #base_filtrada = base_filtrada.reset_index(drop=True)

    pessoa = '4209'

    return(wks1, base, base_filtrada, pessoa)

def planilha_usinagem(filename, data):

    #USINAGEM#

    # sheet = 'RQ PCP-001-000 (APONTAMENTO SERRA)'
    # worksheet1 = 'USINAGEM 2022'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    sheet_id = '1Kiupceopsz-HHZ9TJgWhxADNtFHVptLszsDeJyxkiYk'
    worksheet1 = 'RQ PCP-014-001'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)   

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(5)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:11]
    base = base.set_axis(headers, axis=1)[5:]

    ########### Tratando planilhas ###########

    #ajustando datas
    # for i in range(len(base)+5):
    #     try:
    #         if base['DATA'][i] == "":
    #             base['DATA'][i] = base['DATA'][i-1]
    #     except:
    #         pass

    for i in range(len(base)+5):
        try:
            if base['DATA'][i] == "":
                data_antes = base.at[i-1, 'DATA']
                base.at[i, 'DATA'] = data_antes
        except:
            pass

    i = None

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['PCP'].isnull()]

    # base['mes_data'] = ''
    # base['dia'] = ''
    
    # for i in range(base.index[-1]+1):
    #     try:
    #         base['mes_data'][i] = base['DATA'][i][3:5]
    #         base['dia'][i] = base['DATA'][i][0:2]
    #     except:
    #         pass
        
    # mes = mes_atual()

    # if len(mes) == 1:
    #     mes = '0' + mes

    base_filtrada = base[base['DATA'] == data]
    # base_filtrada['dia'] = base_filtrada['dia'].astype(int)
    # base_filtrada = base_filtrada.sort_values(by='dia', ascending=False)

    base_filtrada = base_filtrada.fillna('')

    #filtrando linhas sem observação
    base_filtrada = base_filtrada[base_filtrada['PCP'] == '']
    #base_filtrada = base_filtrada[base_filtrada['OBSERVAÇÃO'].isnull()]
    #base_filtrada = base_filtrada[base_filtrada['PCP'].isnull()]
    
    #inserindo 0 antes do código da peca
    base_filtrada['CÓDIGO'] = base_filtrada['CÓDIGO'].astype(str)

    for i in range(len(base)):
        try:
            if len(base_filtrada['CÓDIGO'][i]) == 5:
                base_filtrada['CÓDIGO'][i] = "0" + base_filtrada['CÓDIGO'][i] 
        except:
            pass

    i = None

    base_filtrada = base_filtrada[['DATA','CÓDIGO','QNT', 'QNTD MORTA']]

    pessoa = '4057'

    return(wks1, base, base_filtrada, pessoa)

def planilha_corte(filename, data):

    #CORTE#

    # sheet = 'RQ PCP-012-000 (Banco de dados OP)'
    # worksheet1 = 'Finalizadas'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    sheet_id = '1t7Q_gwGVAEwNlwgWpLRVy-QbQo7kQ_l6QTjFjBrbWxE'
    worksheet1 = 'RQ PCP-012-000'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    #base = base.iloc

    headers = wks1.row_values(5)#[0:3]

    base = base.set_axis(headers, axis=1)[5:]

    base['Data finalização'] = base['Data finalização'].str[:10]

    ########### Tratando planilhas ###########

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['PCP'].isnull()]

    # base['mes_data'] = ''
    # base['dia'] = ''
    
    # for i in range(base.index[-1]+1):
    #     try:
    #         base['mes_data'][i] = base['Data finalização'][i][3:5]
    #         base['dia'][i] = base['Data finalização'][i][0:2]
    #     except:
    #         pass
        
    # mes = mes_atual()

    # if len(mes) == 1:
    #     mes = '0' + mes

    base_filtrada = base[base['Data finalização'] == data]
    # base_filtrada['dia'] = base_filtrada['dia'].astype(int)
    # base_filtrada = base_filtrada.sort_values(by='dia', ascending=False)

    base_filtrada = base_filtrada.fillna('')

    #filtrando linhas que foram transferidas
    base_filtrada = base_filtrada.loc[base_filtrada['Transf. chapa'] != '']

    #filtrando linhas que foram transferidas
    base_filtrada = base_filtrada.loc[base_filtrada['Apont. peças'] == '']

    #filtrando linhas que tem código de chapa
    base_filtrada = base_filtrada[base_filtrada['Código Chapa'] != '']

    #extraindo código
    base_filtrada["Peça"] = base_filtrada["Peça"].str[:6]

    for i in range(len(base)):
        try:
            if len(base_filtrada['Peça'][i]) == 5:
                base_filtrada['Peça'][i] = "0" + base_filtrada['Peça'][i] 
        except:
            pass

    base_filtrada = base_filtrada[base_filtrada['Erros'] == '']

    base_filtrada = base_filtrada[['Data finalização','Peça','Total Prod.','Mortas']]

    pessoa = '4161'

    return(wks1, base, base_filtrada, pessoa)

def planilha_estamparia(filename, data):

    #ESTAMPARIA#

    # sheet = 'RQ PCP-003-001 (APONTAMENTO ESTAMPARIA) / RQ PCP-009-000 (SEQUENCIAMENTO ESTAMPARIA) / RQ CQ-008-000 (Inspeção do Corte) / RQ CQ-015-000 (Inspeção da Estamparia)'
    # worksheet1 = 'APONTAMENTO PCP (RQ PCP 003 001)'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    sheet_id = '1AE8qNIxtAUVIx4B8vmLnIp5WBnNMNweWE38rJbczDXg'
    worksheet1 = 'RQ PCP-003-002'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:12]

    headers = wks1.row_values(5)[0:12]

    base = base.set_axis(headers, axis=1)[5:]

    ########### Tratando planilhas ###########

    # base['mes_data'] = ''
    # base['dia'] = ''
   
    # for i in range(base.index[-1]+1):
    #     try:
    #         base['mes_data'][i] = base['DATA'][i][3:5]
    #         base['dia'][i] = base['DATA'][i][0:2]
    #     except:
    #         pass
        
    # mes = mes_atual()

    # if len(mes) == 1:
    #     mes = '0' + mes

    base_filtrada = base[base['DATA'] == data]
    # base_filtrada['dia'] = base_filtrada['dia'].astype(int)
    # base_filtrada = base_filtrada.sort_values(by='dia', ascending=False)

    #filtrando data de hoje
    #base_filtrada = base.loc[base.DATA == data]

    base_filtrada = base_filtrada.fillna('')

    #filtrando linhas que não tem status de ok
    base_filtrada = base_filtrada.loc[base_filtrada.STATUS == '']

    #filtrando linhas sem observação
    base_filtrada = base_filtrada.loc[base_filtrada.CÓDIGO != '']

    #MATRICULA ALEX: 4322
    for i in range(len(base)):
        try:
            if base_filtrada['MATRÍCULA'][i] == '': 
                base_filtrada['MATRÍCULA'][i] = '4322'
        except:
            pass

    base_filtrada['MATRÍCULA'] = base_filtrada['MATRÍCULA'].str[:4]

    base_filtrada = base_filtrada.loc[(base_filtrada['QTD'] != '')]
    
    base_filtrada = base_filtrada[['DATA','MATRÍCULA','CÓDIGO','QTD']]

    return(wks1, base, base_filtrada)

def planilha_montagem(filename, data):

    #MONTAGEM#

    # sheet = 'RQ PCP-004-000 APONTAMENTO MONTAGEM M22'
    # worksheet1 = 'APONTAMENTO'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    sheet_id = '14h59mwOICMx92bPkS-mIdDZmM_-NoeSAE1D418zTKfo'
    worksheet1 = 'RQ PCP-004-000'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:9]

    headers = wks1.row_values(5)[0:9]

    base = base.set_axis(headers, axis=1)[5:]

    ########### Tratando planilhas ###########
    
    base['CARIMBO'] = base['CARIMBO'].str[:10]

    base['FUNCIONÁRIO'] = base['FUNCIONÁRIO'].str[:4]
    
    # base['mes_data'] = ''
    # base['dia'] = ''

    base = base[base['CARIMBO'] != '#REF!']
    base = base[base['CARIMBO'] != '']
    
    # for i in range(base.index[-1]+1):
    #     try:
    #         base['mes_data'][i] = base['CARIMBO'][i][3:5]
    #         base['dia'][i] = base['CARIMBO'][i][0:2]
    #     except:
    #         pass
        
    # mes = mes_atual()

    # if len(mes) == 1:
    #     mes = '0' + mes

    base_filtrada = base[base['CARIMBO'] == data]
    # base_filtrada['dia'] = base_filtrada['dia'].astype(int)
    # base_filtrada = base_filtrada.sort_values(by='dia', ascending=False)

    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace('-','_', regex=True)
    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace('-','_', regex=True)
    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace(' ','_', regex=True)
    base_filtrada['n_underscore'] = base_filtrada['CONJUNTO'].str.find('_')
    base_filtrada = base_filtrada.reset_index()

    for i in range(len(base_filtrada)+5):
        try:
            base_filtrada['CONJUNTO'][i] = base_filtrada['CONJUNTO'][i][:base_filtrada['n_underscore'][i]]
        except:
            pass

    base_filtrada = base_filtrada.fillna('')

    #filtrando linhas que não tem status de ok
    base_filtrada = base_filtrada[base_filtrada['STATUS'] == ''] 

    base_filtrada = base_filtrada[['index', 'CARIMBO','FUNCIONÁRIO','CONJUNTO','QUANTIDADE']]
    
    base_filtrada = base_filtrada.set_index('index')

    return(wks1, base, base_filtrada)

def planilha_sucata(filename, data):

    #SUCATA#

    sheet = 'RQ PCP-012-000 (Banco de dados OP)'
    worksheet1 = 'Finalizadas'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    #base = base.iloc

    headers = wks1.row_values(5)#[0:3]

    base = base.set_axis(headers, axis=1)[5:]

    ########### Tratando planilhas ###########

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['PCP'].isnull()]

    #filtrando data de hoje
    base_filtrada = base.loc[base['Data finalização'] == data]

    #filtrando linhas que foram transferidas
    base_filtrada = base_filtrada.loc[base_filtrada['Transf. chapa'] != '']

    #filtrando linhas que foram transferidas
    base_filtrada = base_filtrada.loc[base_filtrada['Apont. peças'] == '']

    #filtrando linhas que tem código de chapa
    base_filtrada = base_filtrada[base_filtrada['Código Chapa'].notnull()]
    base_filtrada = base_filtrada[base_filtrada['Código Chapa'] != '']

    #extraindo código
    base_filtrada["Peça"] = base_filtrada["Peça"].str[:6]

    for i in range(len(base)):
        try:
            if len(base_filtrada['Peça'][i]) == 5:
                base_filtrada['Peça'][i] = "0" + base_filtrada['Peça'][i] 
        except:
            pass

    base_filtrada = base_filtrada[['Data finalização','Peça','Total Prod.','Mortas']]

    pessoa = '4161'

    return(wks1, base, base_filtrada, pessoa)

def planilha_pintura(filename, data):

    #PINTURA#

    # sheet = 'BANCO DE DADOS ÚNICO - PINTURA / RQ CQ-019-000 (Inspeção da Pintura)'
    # worksheet1 = 'RQ PCP-005-003 (APONTAMENTO PINTURA)'

    # sa = gspread.service_account(filename)
    # sh = sa.open(sheet)

    sheet_id = '1RJH3k5brgO3nmEQPNOKdp_HFqVbJMUcwGEQ1XQyHtCs'
    worksheet1 = 'RQ PCP-005-003 (APONTAMENTO PINTURA)'

    sa = gspread.service_account(filename)
    sh = sa.open_by_key(sheet_id)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    #base = base.iloc

    headers = wks1.row_values(5)#[0:3]

    base = base.iloc[0:,0:24]

    base = base.set_axis(headers, axis=1)[5:]

    ########### Tratando planilhas ###########
    # base['mes_data'] = ''
    # base['dia'] = ''
    
    # for i in range(base.index[-1]+1):
    #     try:
    #         base['mes_data'][i] = base['Carimbo'][i][3:5]
    #         base['dia'][i] = base['Carimbo'][i][0:2]
    #     except:
    #         pass
        
    # mes = mes_atual()

    # if len(mes) == 1:
    #     mes = '0' + mes

    base_filtrada = base[base['Carimbo'] == data]
    # base_filtrada['dia'] = base_filtrada['dia'].astype(int)
    # base_filtrada = base_filtrada.sort_values(by='dia', ascending=False)

    #filtrando data de hoje
    #base_filtrada = base.loc[base['Carimbo'] == data]

    base_filtrada = base_filtrada.fillna('')

    #filtrando STATUS VAZIO
    base_filtrada = base_filtrada.loc[base_filtrada['STATUS'] == '']

    #filtrando linhas que estão ok na pintura
    base_filtrada = base_filtrada.loc[base_filtrada['PINTURA'] != '']

    base_filtrada['CÓDIGO'] = base_filtrada['CÓDIGO'].replace('-','_', regex=True)
    base_filtrada['CÓDIGO'] = base_filtrada['CÓDIGO'].replace(' ','_', regex=True)
    base_filtrada['n_underscore'] = base_filtrada['CÓDIGO'].str.find('_')
    
    for i in range(len(base)+5):
        try:
            base_filtrada['CÓDIGO'][i] = base_filtrada['CÓDIGO'][i][:base_filtrada['n_underscore'][i]]
        except:
            pass

    base_filtrada = base_filtrada[['Carimbo','CÓDIGO','Qtd', 'Tipo','Cor','CONSUMO PU','CONSUMO PÓ','CATALISADOR']]

    pessoa = '4271'

    return(wks1, base, base_filtrada, pessoa)

########### PREENCHIMENTO TRANSFERÊNCIA DE MP ###########

def preenchendo_serra_transf(nav, data, peca, qtde, wks1, c, i):

    iframes(nav)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #Classe
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[4]/div/input'))).send_keys(Keys.TAB)
    
    #Solicitante
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[6]/div/input'))).send_keys(Keys.TAB)
    
    #Deposito de origem
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[8]/div/input'))).send_keys('central')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[8]/div/input'))).send_keys(Keys.TAB)
    time.sleep(1)

    #Deposito de destino
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[10]/div/input'))).send_keys('serra')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[10]/div/input'))).send_keys(Keys.TAB)
    time.sleep(1)

    #Recurso
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[12]/div/input'))).send_keys(peca)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[12]/div/input'))).send_keys(Keys.TAB)
    time.sleep(1)

    #Lote
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[14]/div/input'))).send_keys(Keys.TAB)

    #Campo vazio
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[16]/div/input'))).send_keys(Keys.TAB)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[17]/div/input'))).send_keys(Keys.TAB)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[19]/div/input'))).send_keys(Keys.TAB)

    #Quantidade
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[21]/div/input'))).send_keys(qtde)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[21]/div/input'))).send_keys(Keys.TAB)
    
    #click em campo fantasma
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

    time.sleep(2)

    #click em confirmar
    try:
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
    except:
        pass

    time.sleep(2)

    c = c+2
    
    print(c)
    return(c)

def preenchendo_corte_transf(nav, data, peca, qtde, wks1, c, i):

    iframes(nav)

    #Insert
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()
    time.sleep(1.5)

    #Classe
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[4]/div/input'))).send_keys(Keys.TAB)
    time.sleep(1.5)

    #Solicitante
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[6]/div/input'))).send_keys(Keys.TAB)
    
    #Deposito de origem
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[8]/div/input'))).send_keys('central')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[8]/div/input'))).send_keys(Keys.TAB)
    time.sleep(1.5)

    #Deposito de destino
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[10]/div/input'))).send_keys('corte')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[10]/div/input'))).send_keys(Keys.TAB)
    time.sleep(1.5)

    #Recurso
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[12]/div/input'))).send_keys(peca)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[12]/div/input'))).send_keys(Keys.TAB)
    time.sleep(1.5)
    
    #Lote
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[14]/div/input'))).send_keys(Keys.TAB)

    #Campo vazio
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[16]/div/input'))).send_keys(Keys.TAB)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[17]/div/input'))).send_keys(Keys.TAB)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[19]/div/input'))).send_keys(Keys.TAB)

    #Quantidade
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[21]/div/input'))).send_keys(qtde)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[21]/div/input'))).send_keys(Keys.TAB)
    
    #click em campo fantasma
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

    time.sleep(2)

    #click em confirmar
    try:
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
    except:
        pass
    
    time.sleep(2)

    c = c+2
    
    print(c)
    return(c)

def selecionar_todos(nav,data):

    #selecinar todos os campos
    time.sleep(2)
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[1]/div'))).click()
    time.sleep(1)

    #aprovar
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/span[2]/p'))).click()
    time.sleep(1)
    
    time.sleep(1.5)
    #fechar pop-up
    nav.switch_to.default_content()
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.ID, 'confirm'))).click()

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

    #baixar 
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[3]/span[2]/p'))).click()
    time.sleep(1)

    #data 
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(data)
    time.sleep(3)

    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    time.sleep(0.5)
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + Keys.SHIFT + 'f')
    time.sleep(0.5)
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + Keys.SHIFT + 'f')

    try:
        while WebDriverWait(nav, 2).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[9]/table/tbody/tr/td[2]/div'))):
            print("Carregando")
    except:
        print("Carregou")

    iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try:
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass

    time.sleep(5)

    # try:
    # #confirmar baixa
    #     WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span[2]/p'))).click()
    #     WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span[2]/p'))).click()
    #     WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span[2]/p'))).click()
    #     time.sleep(10)
    # except:
    #     pass
    
    #texto erro
    try:
        text_erro = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button'))).click()
    except:
        #gravar
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

        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[4]/div/input'))).send_keys(Keys.CONTROL + Keys.SHIFT + 'g')
        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[4]/div/input'))).send_keys(Keys.CONTROL + Keys.SHIFT + 'g')
        # WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span[2]/p'))).click()
        # WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span[2]/p'))).click()
        
        try:
            while WebDriverWait(nav, 2).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[9]/table/tbody/tr/td[2]/div'))):
                print("Carregando")
        except:
            print("Carregou")

########### PREENCHIMENTO APONTAMENTO DE PEÇA ###########

def preenchendo_serra(nav, data, pessoa, peca, qtde, wks1, c, i, mortas):

    iframes(nav)
    
    erro = 0

    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    if c <= 3:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
        time.sleep(1.2)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
        time.sleep(1.2)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
        time.sleep(1.2)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
        time.sleep(1.2)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    else:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)

    try:
        nav.switch_to.default_content()

        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframes(nav)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    iframes(nav)

    #data
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(data)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.TAB)

    #inicio
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[6]/div/input"))).send_keys(Keys.TAB)

    #Fim
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[7]/div/input"))).send_keys(Keys.TAB)

    #pessoa
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(pessoa)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #peça
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(peca)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.TAB)

    try:
        iframes(nav)
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        wks1.update('U' + str(i+1), 'Código não encontrado')
        return(c)
    except:

        iframes(nav)
    
        #Etapa
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)
            
        #processo
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys('S')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
        time.sleep(2)

        try:
            iframes(nav)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
            wks1.update('U' + str(i+1), 'Processo não encontrado')
            return(c)

        except:
            
            try:
                iframes(nav)
            except:
                pass


            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)

            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            
            time.sleep(2)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)
            
            time.sleep(2)

            qtd_corfirmar = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).text
            
            time.sleep(2)

            if qtd_corfirmar=="":
                print("Quantidade vazia")
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).click()    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)    
                print("Quantidade preenchida")

            if mortas != '':

                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[19]/div/input"))).send_keys(Keys.TAB)
                time.sleep(3)

                #Mortas
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[21]/div/input"))).send_keys(mortas)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[21]/div/input"))).send_keys(Keys.TAB)
                
                #deposito desviado
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[22]/div/input"))).send_keys('Almox Sucata')
                
            time.sleep(1.5)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
            
            iframes(nav)

            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")
            
            time.sleep(2)

            try:
                WebDriverWait(nav, 20).until(EC.presence_of_element_located((By.ID, 'errorMessageBox')))
                
                # volta p janela principal (fora do iframe)
                time.sleep(1)
                nav.switch_to.default_content()
                
                time.sleep(5)
                
                texto_erro = WebDriverWait(nav, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text

                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()

                time.sleep(1)
                iframes(nav)
                
                time.sleep(1)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/div"))).click()

                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/input"))).send_keys(Keys.ESCAPE)
                time.sleep(1.5)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/input"))).send_keys(Keys.ENTER)
                time.sleep(1.5)
                                
                wks1.update('U' + str(i+1), texto_erro + ' ' + data_hoje() + ' ' + hora_atual())
                time.sleep(2)
                
                fechar_tabs(nav)

                iframes(nav)

                menu_innovaro(nav)
                time.sleep(2)
                
                lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                time.sleep(1.5)
                click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                
                lista_menu[click_producao].click() ##clicando em Apontamento da produção
 
                time.sleep(1.5)
                                
                c = 3
                return(c)
            
            except TimeoutException:
            
                wks1.update('S' + str(i+1), 'OK ROBINHO - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)
            time.sleep(1.5)

    return(c)

def preenchendo_usinagem(nav, data, pessoa, peca, qtde, wks1, c, i, mortas):

    erro = 0
    # hora = datetime.now()
    # hora = hora.strftime("%H:%M:%S")
    # wks1.update('E' + str(i+2), hora)
   
    iframes(nav)

    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    if c <= 3:

        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    else:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)

    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    iframes(nav)

    #data
    
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(data)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.TAB)

    #inicio
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[6]/div/input"))).send_keys(Keys.TAB)

    #Fim
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[7]/div/input"))).send_keys(Keys.TAB)

    #pessoa
    
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(pessoa)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #peça
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(peca)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.TAB)
    time.sleep(3)

    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        wks1.update('K' + str(i+1), 'Código não encontrado')
        return(c)
    except:

        iframes(nav)
        
        #Etapa
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)
            
        #processo
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys('S Usi')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
        time.sleep(3)

        try:
            nav.switch_to.default_content()
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
            wks1.update('K' + str(i+1), 'Processo não encontrado')
            return(c)
        except:
            print("deu ruim")

            iframes(nav)

            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)

            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)

            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)
            
            time.sleep(2)

            qtd_corfirmar = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).text
            
            if qtd_corfirmar=="":
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).click()    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)    

            if mortas != '':

                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[19]/div/input"))).send_keys(Keys.TAB)
                time.sleep(3)

                #Mortas
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[21]/div/input"))).send_keys(mortas)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[21]/div/input"))).send_keys(Keys.TAB)
                
                #deposito desviado
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[22]/div/input"))).send_keys('Almox Sucata')
                
            time.sleep(1.5)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()

            nav.switch_to.default_content()

            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")
            
            time.sleep(2)

            erro = 0
            try:
                WebDriverWait(nav, 20).until(EC.presence_of_element_located((By.ID, 'errorMessageBox')))

                time.sleep(1)
                nav.switch_to.default_content()
                
                time.sleep(5)
                
                texto_erro = WebDriverWait(nav, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text

                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()

                time.sleep(1)
                iframes(nav)
                
                time.sleep(1)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/div"))).click()

                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/input"))).send_keys(Keys.ESCAPE)
                time.sleep(1.5)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/input"))).send_keys(Keys.ENTER)
                time.sleep(1.5)
                
                wks1.update('K' + str(i+1), texto_erro + ' ' + data_hoje() + ' ' + hora_atual())

                time.sleep(2)

                fechar_tabs(nav)

                nav.switch_to.default_content()

                menu_innovaro(nav)

                time.sleep(3)
                
                lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                time.sleep(1.5)
                click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                
                lista_menu[click_producao].click() ##clicando em Apontamento da produção
                time.sleep(1.5)
                                
                c = 3

                return(c)
            
            except TimeoutException:
                
                wks1.update('J' + str(i+1), 'OK ROBINHO - ' + ' ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)
            time.sleep(1.5)

    return(c)

def preenchendo_corte(nav, data, pessoa, peca, qtde, wks1, c, i, mortas):

    erro = 0

    #mudando iframe
    iframes(nav)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    if c <= 3:

        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    else:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)

    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframes(nav)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    iframes(nav)

    #data
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(data)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.TAB)

    #inicio
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[6]/div/input"))).send_keys(Keys.TAB)

    #Fim
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[7]/div/input"))).send_keys(Keys.TAB)

    #pessoa
    if c == 3:
    
        time.sleep(1)
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(pessoa)
        time.sleep(1)
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    else:
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #peça
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(peca)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.TAB)

    try:
        time.sleep(3)
        nav.switch_to.default_content()
        time.sleep(1.5)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        wks1.update('O' + str(i+1), 'Código não encontrado')
        return(c)
    except:
        print('deu ruim')
    
    time.sleep(1)

    try:
        time.sleep(1)
        iframes(nav)
    except:
        print('deu ruim')

    try:
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[3]/span[2]"))).click()
    except:
        pass

    if WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[10]/div/div"))).text == '':
        
        time.sleep(2)

        fechar_tabs(nav)

        nav.switch_to.default_content()
        menu_innovaro(nav)
        time.sleep(3)
        
        lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
        time.sleep(1.5)
        click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
        
        lista_menu[click_producao].click() ##clicando em Apontamento da produção
        time.sleep(1.5)
                                
        c = 3

        time.sleep(1.5)

    else:
        #Etapa
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)    
        time.sleep(1)

        #Processo
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).click()
        time.sleep(1)
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys('S')
        time.sleep(4)                                                       
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)

        processo_texto = WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/div"))).text

        if processo_texto == '':
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/div"))).click()
            time.sleep(1)                                                                   
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys('S')
            time.sleep(4)                                                       
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
        else:
            pass

        time.sleep(1.5)

        #saindo do erro caso nao ache o processo
        try:
            nav.switch_to.default_content()
            time.sleep(2)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(2)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(2)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
            wks1.update('O' + str(i+1), 'Processo não encontrado')
            return(c)
        except:
            print('Processo encontrado')

            iframes(nav)
                
            time.sleep(1)

            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)
            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)
            
            time.sleep(2)

            qtd_corfirmar = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).text
            
            if qtd_corfirmar=="":
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).click()    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)    

            if mortas != '':

                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[19]/div/input"))).send_keys(Keys.TAB)
                time.sleep(3)

                #Mortas
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[21]/div/input"))).send_keys(mortas)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[21]/div/input"))).send_keys(Keys.TAB)
                
                #deposito desviado
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[22]/div/input"))).send_keys('Almox Sucata')
                
            time.sleep(1.5)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()

            nav.switch_to.default_content()

            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")

            time.sleep(5)
            erro = 0

            try:
                WebDriverWait(nav, 20).until(EC.presence_of_element_located((By.ID, 'errorMessageBox')))
                
                time.sleep(1)
                nav.switch_to.default_content()
                
                time.sleep(5)
                
                texto_erro = WebDriverWait(nav, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
                erro = 1
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()

                time.sleep(1)
                iframes(nav)
                
                time.sleep(1)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/div"))).click()

                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/input"))).send_keys(Keys.ESCAPE)
                time.sleep(1.5)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/input"))).send_keys(Keys.ENTER)
                time.sleep(1.5)
                
                wks1.update('O' + str(i+1), texto_erro + ' ' + data_hoje() + ' ' + hora_atual())

                time.sleep(2)

                fechar_tabs(nav)

                nav.switch_to.default_content()
                menu_innovaro(nav)
                time.sleep(3)
                
                lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                time.sleep(1.5)
                click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                
                lista_menu[click_producao].click() ##clicando em Apontamento da produção
                time.sleep(1.5)
                                        
                c = 3

                time.sleep(1.5)
                return(c)

            except TimeoutException:
            
                wks1.update('L' + str(i+1), 'OK ROBINHO - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)
            time.sleep(1.5)

    return(c)

def preenchendo_estamparia(nav, data, pessoa, peca, qtde, wks1, c, i):

    erro = 0

    print(c)

    iframes(nav)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass

    if c <= 3:

        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    else:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)

    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    iframes(nav)

    #data
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(data)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.TAB)

    #inicio
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[6]/div/input"))).send_keys(Keys.TAB)

    #Fim
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[7]/div/input"))).send_keys(Keys.TAB)

    #pessoa
    #if c == 3:
    
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(pessoa)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #else:
    #    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #peça
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(peca)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.TAB)
    
    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        wks1.update('K' + str(i+1), 'Código não encontrado')
        return(c)
    except:
        print("deu ruim")

        iframes(nav)

        #Etapa
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)    

        #processo
        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys('S Est')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)

        try:
            nav.switch_to.default_content()
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
            wks1.update('K' + str(i+1), 'Processo não encontrado')
            return(c)
        except:
            print("deu ruim")
            
            iframes(nav)

            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)
            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            
            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)
            
            time.sleep(2)

            qtd_corfirmar = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).text
            
            if qtd_corfirmar=="":
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).click()    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)    

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()

            nav.switch_to.default_content()

            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")

            time.sleep(2)
            erro = 0

            try:
                WebDriverWait(nav, 20).until(EC.presence_of_element_located((By.ID, 'errorMessageBox')))
                
                texto_erro = WebDriverWait(nav, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text

                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()

                time.sleep(1)
                iframes(nav)
                
                time.sleep(1)
                # WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[19]/div/div"))).click()

                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[19]/div/input"))).send_keys(Keys.ESCAPE)
                time.sleep(1.5)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[19]/div/input"))).send_keys(Keys.ENTER)
                time.sleep(1.5)

                wks1.update('K' + str(i+1), texto_erro + ' ' + data_hoje() + ' ' + hora_atual())
                wks1.update('L' + str(i+1), peca + ' - ' + qtde)

                time.sleep(2)

                fechar_tabs(nav)

                nav.switch_to.default_content()

                menu_innovaro(nav)

                time.sleep(3)
                
                lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                time.sleep(1.5)
                click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                
                lista_menu[click_producao].click() ##clicando em Apontamento da produção
                time.sleep(1.5)
                                
                c = 3
                
                return(c)

            except TimeoutException:
                
                wks1.update('J' + str(i+1), 'OK ROBINHO - ' + data_hoje() + ' ' + hora_atual())
                wks1.update('L' + str(i+1), peca + ' - ' + qtde)

                print('deu bom')

                c = c + 2

            print(c)
            time.sleep(1.5)


    return(c)

def preenchendo_montagem(nav, data, pessoa, peca, qtde, wks1, c, i):

    erro = 0

    #mudando iframe
    iframes(nav)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    if c <= 3:

        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    
    else:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)

    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    iframes(nav)

    #data
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(data)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.TAB)

    #inicio
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[6]/div/input"))).send_keys(Keys.TAB)

    #Fim
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[7]/div/input"))).send_keys(Keys.TAB)

    #pessoa
    #if c == 3:
    
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(pessoa)
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #else:
    #    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #peça
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(peca)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.TAB)

    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        wks1.update('J' + str(i+1), 'Código não encontrado')
        return(c)
    except:

        iframes(nav)
        
        #Etapa
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)
            
        #processo
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys('S Mont')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)

        try:
            nav.switch_to.default_content()
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
            wks1.update('J' + str(i+1), 'Processo não encontrado')
            return(c)
        except:

            iframes(nav)

            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)
            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)
            
            time.sleep(2)

            qtd_corfirmar = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).text
            
            if qtd_corfirmar=="":
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).click()    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)    

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
            
            nav.switch_to.default_content()
            
            try:
                while WebDriverWait(nav, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")

            time.sleep(2)
            erro = 0
            try:
                WebDriverWait(nav, 20).until(EC.presence_of_element_located((By.ID, 'errorMessageBox')))
                
                time.sleep(5)
                
                texto_erro = WebDriverWait(nav, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text

                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()

                iframes(nav)
                
                time.sleep(1)
                #WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/div"))).click()

                WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.NAME, "DEPOSITODEST"))).send_keys(Keys.ESCAPE)
                time.sleep(1.5)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.NAME, "DEPOSITODEST"))).send_keys(Keys.ENTER)
                time.sleep(1.5)

                wks1.update('J' + str(i+1), texto_erro + '' + data_hoje() + ' ' + hora_atual())
                time.sleep(2)

                fechar_tabs(nav)

                nav.switch_to.default_content()

                menu_innovaro(nav)

                time.sleep(3)
                
                lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                time.sleep(1.5)
                click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                
                lista_menu[click_producao].click() ##clicando em Apontamento da produção
                time.sleep(1.5)
                                
                c = 3

                return(c)

            except TimeoutException:

                wks1.update('I' + str(i+1), 'OK ROBINHO - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            time.sleep(1.5)
            print(c)
            time.sleep(1.5)


    return(c)

def preenchendo_pintura(nav, data, pessoa, peca, qtde,tipo, cor, wks1, c, i):

    print(c)

    erro = 0

    iframes(nav)
    
    #Insert
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    if c <= 3:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    else:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)

    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    iframes(nav)

    #data
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(data)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(Keys.TAB)  

    #inicio
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[6]/div/input"))).send_keys(Keys.TAB)

    #Fim
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[7]/div/input"))).send_keys(Keys.TAB)

    #pessoa
    #if c == 3:
    
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(pessoa)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #else:
    #    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(Keys.TAB)

    #peça
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).click()
    time.sleep(0.5)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(peca)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.TAB)

    try:
        if WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/form/table"))):
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.ENTER)
            time.sleep(1)
            #WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.ENTER)
            time.sleep(1)
        else:
            #WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(Keys.TAB)
            time.sleep(1)
            WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).click()
    except:
        pass
                                                                    
    try:
        time.sleep(2)
        nav.switch_to.default_content()
        time.sleep(1.5)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        wks1.update('T' + str(i+1), 'Código não encontrado')
        
        c = 3
        return(c)
    except:
        
        iframes(nav)
        
        #Etapa
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)
            
        #processo
        # if c <= 3:
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys('S Pint')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
        # else:
        #     WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)

        try:
            time.sleep(3)
            nav.switch_to.default_content()
            time.sleep(1.5)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
            wks1.update('T' + str(i+1), 'Processo não encontrado')
            return(c)
        except:

            iframes(nav)

            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)
            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)
            
            time.sleep(2)

            qtd_corfirmar = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).text
            
            if qtd_corfirmar=="":
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div"))).click()    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)    
                time.sleep(2)
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(Keys.TAB)    

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)
            
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()

            nav.switch_to.default_content()

            try:
                while WebDriverWait(nav, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")
            
            time.sleep(5)
            erro = 0

            try:
                WebDriverWait(nav, 20).until(EC.presence_of_element_located((By.ID, 'errorMessageBox')))
                               
                time.sleep(5)
                
                texto_erro = WebDriverWait(nav, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
                erro = 1
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
                
                time.sleep(1)
                iframes(nav)

                # WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr["+ str(c) +"]/td[21]/div/div"))).click()

                WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.NAME, "DEPOSITODEST"))).send_keys(Keys.ESCAPE)
                time.sleep(1.5)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.NAME, "DEPOSITODEST"))).send_keys(Keys.ENTER)
                time.sleep(1.5)
                
                wks1.update('T' + str(i+1), texto_erro + '' + data_hoje() + ' ' + hora_atual())

                time.sleep(2)

                nav.switch_to.default_content()

                fechar_tabs(nav)

                menu_innovaro(nav)
                
                time.sleep(3)
                
                lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                time.sleep(1.5)
                click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                
                lista_menu[click_producao].click() ##clicando em Apontamento da produção
                time.sleep(1.5)

                c = 3

                return(c)

            except TimeoutException:
               
                wks1.update('O' + str(i+1), 'OK ROBS - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            time.sleep(2)

            print(c)
            print(erro)
            time.sleep(1.5)

    return(c)

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

            qtde_itens = len(base_filtrada.drop_duplicates(subset=['MATERIAL']))

            for i in range(len(base_filtrada)):
                recurso = base_filtrada['MATERIAL'][i]
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
            tabelona['Saldo'] = tabelona.Saldo.shift(-2)
            tabelona = tabelona.dropna()
            tabelona = tabelona.reset_index(drop=True)
            
            if qtde_itens == 1:
                tabelona = tabelona[:1]
            else:
                tabelona = tabelona[:len(tabelona)-2]

            #quebrando_material = tabelona["Unnamed: 0_level_2"].str.split(" ", n = 1, expand = True)

            #tabelona['Unnamed: 0_level_2'] = quebrando_material[0]

            for i in range(len(tabelona)):
                if len(tabelona['Saldo'][i]) > 6 :
                    tabelona['Saldo'][i] = tabelona['Saldo'][i].replace(',','')
                    tabelona['Saldo'][i] = tabelona['Saldo'][i].replace('.','')

            for saldo in range(len(tabelona)):
                try:
                    tabelona['Saldo'][saldo] = tabelona['Saldo'][saldo][:len(tabelona['Saldo'][saldo])-4] + "." + tabelona['Saldo'][saldo][-4:]  
                except:
                    pass

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

            tabelona['Saldo'] = tabelona['Saldo'].astype(float)

            #tabelona = tabelona.rename(columns={'Unnamed: 0_level_2':'MATERIAL'})
            tabelona = tabelona.rename(columns={'Unid. Medida':'MATERIAL'})

            for i in range(len(tabelona)):
                tabelona['MATERIAL'][i] = tabelona['MATERIAL'][i][:len(tabelona['MATERIAL'][i]) - 5]
            
            df_final = pd.merge(tabelona,base_filtrada,on='MATERIAL')

            lista_material = df_final['MATERIAL'].unique()           

            df_final['saldo2'] = ''
            
            df_final1 = pd.DataFrame()
            
            for material in lista_material:
                df_filtro_material = df_final[df_final['MATERIAL'] == material].reset_index(drop=True)
                
                for i in range(len(df_filtro_material)):
                    try:
                        df_filtro_material['saldo2'][i] = df_filtro_material['saldo2'][i-1] - df_filtro_material['PESO BARRAS'][i]
                    except:
                        df_filtro_material['saldo2'][i] = df_filtro_material['Saldo'][i] - df_filtro_material['PESO BARRAS'][i]
                    
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
            tabelona['Saldo'] = tabelona.Saldo.shift(-2)
            tabelona = tabelona.dropna()
            tabelona = tabelona.reset_index(drop=True)
            
            if qtde_itens == 1:
                tabelona = tabelona[:1]
            else:
                tabelona = tabelona[:len(tabelona)-2]

            #quebrando_material = tabelona["Unnamed: 0_level_2"].str.split(" ", n = 1, expand = True)

            #tabelona['Unnamed: 0_level_2'] = quebrando_material[0]

            for i in range(len(tabelona)):
                if len(tabelona['Saldo'][i]) > 6 :
                    tabelona['Saldo'][i] = tabelona['Saldo'][i].replace(',','')
                    tabelona['Saldo'][i] = tabelona['Saldo'][i].replace('.','')

            try:
                for j in range(len(tabelona)):
                    if len(tabelona['Saldo'][j]) >= 6 :
                        tabelona['Saldo'][j] = float(tabelona['Saldo'][j]) / 10000
            except:
                pass

            tabelona['Saldo'] = tabelona['Saldo'].astype(float)

            #tabelona = tabelona.rename(columns={'Unnamed: 0_level_2':'MATERIAL'})
            tabelona = tabelona.rename(columns={'Unid. Medida':'Código Chapa'})

            for i in range(len(tabelona)):
                tabelona['Código Chapa'][i] = tabelona['Código Chapa'][i][:len(tabelona['Código Chapa'][i]) - 5]

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

########## LOOP ###########

##### onde o robô ta? #####

sheet = 'CENTRAL DE APONTAMENTO'
worksheet1 = 'PAINEL'

sa = gspread.service_account(filename)
sh = sa.open(sheet)

wks2 = sh.worksheet(worksheet1)

#data_dia_1 = '12/07/2023'

def funcao_main():
    
    while True:

        try:

            today = dia_da_semana()

            if today != 1:

                datas = [data_hoje(),  data_ontem()]#, data_sabado()]
                #datas = [data_hoje()]#, data_sabado()]

            else:

                datas = [data_hoje(), data_ontem(), data_sabado(), data_sexta()]

            #datas = ['25/04/2023','26/04/2023','27/04/2023','28/04/2023','29/04/2023','30/04/2023']

            nav = acessar_innovaro()

            time.sleep(4)

            login(nav)

            while True:
                
                for d in range(len(datas)):
                    
                    time.sleep(3)
                    
                    data = datas[d]
                    #data = data_dia_1
                    #data = data_ontem()
                    #data = '01/02/2023'
                    
                    ######### CONSULTAR SALDO ###########
                    
                    menu_innovaro(nav)

                    print("Verificando saldo da serra")

                    wks2.update("E" + "5", 'SALDO SERRA: ' + data) 

                    time.sleep(2)

                    df_final = consulta_saldo(data, nav)
                    df_final = df_final.reset_index(drop=True)
                    
                    time.sleep(2)

                    fechar_menu_consulta(nav)

                    ######## LOOP TRANSFERÊNCIA ###########

                    print("Indo para transferencia de tubos")

                    wks2.update("E" + "5", 'TRANSF. SERRA: ' + data) 

                    time.sleep(2)

                    menu_transf(nav)

                    wks1, base, base_filtrada, transferidas = planilha_serra_transf(data, filename)
                    
                    transferidas = transferidas.reset_index()

                    c = 3

                    i = 0

                    if not len(df_final) == 0:

                        if not int(len(transferidas)) == 0:
                                
                                df_final = df_final.reset_index(drop=True)
                                df_final['comparar2'] = ''

                                for saldo in range(len(df_final)):
                                    
                                    try:
                                        if df_final['MATERIAL'][saldo] == df_final['MATERIAL'][saldo-1]:
                                            df_final['comparar2'][saldo] = df_final['comparar2'][saldo-1] - df_final['PESO BARRAS'][saldo]
                                        
                                            if df_final['comparar2'][saldo] >= df_final['PESO BARRAS'][saldo]:
                                                df_final['comparar'][saldo] = 'True'
                                            else:
                                                df_final['comparar'][saldo] = 'False'
                                        else:
                                            df_final['comparar2'][saldo] = df_final['Saldo'][saldo] - df_final['PESO BARRAS'][saldo]
                                            
                                            if df_final['Saldo'][saldo] >= df_final['PESO BARRAS'][saldo]:
                                                df_final['comparar'][saldo] = 'True'
                                            else:
                                                df_final['comparar'][saldo] = 'False'
                                    except:
                                        df_final['comparar2'][saldo] = df_final['Saldo'][saldo] - df_final['PESO BARRAS'][saldo]
                                        
                                        if df_final['Saldo'][saldo] >= df_final['PESO BARRAS'][saldo]:
                                            df_final['comparar'][saldo] = 'True'
                                        else:
                                            df_final['comparar'][saldo] = 'False'

                                df_final = df_final[df_final['comparar'] == 'True']
                                
                                for i in range(len(base)+1): # serra

                                    print("i: ", i)
                                    try:
                                        peca = df_final['MATERIAL'][i]
                                        qtde = str(df_final['PESO BARRAS'][i])
                                        data = df_final['DATA'][i]
                                        c = preenchendo_serra_transf(nav,data,peca,qtde,wks1,c,i) 
                                        time.sleep(1.5)           
                                        print("c: ", c)

                                        if c == 23:
                                            c = 21           

                                    except:
                                        pass

                                time.sleep(1.5)
                                try:
                                    selecionar_todos(data,nav)

                                    time.sleep(1.5)

                                    for j in range(len(df_final)):
                                        try:
                                            linha_transferida = df_final['index'][j]
                                            wks1.update("T" + str(linha_transferida+1), 'OK TRANSF - ' + data_hoje() + ' ' + hora_atual()) 
                                            time.sleep(1.5)
                                        except:
                                            pass

                                except:
                                    pass

                    ######### CONSULTAR SALDO CORTE ###########

                    print("Verificando saldo de corte")
                    
                    wks2.update("E" + "5", 'SALDO CORTE: ' + data) 

                    time.sleep(2)

                    #fechando aba anterior
                    nav.switch_to.default_content()
                    time.sleep(1.5)

                    menu_innovaro(nav)

                    time.sleep(1)

                    df_final = consulta_saldo_chapas(data, nav)
                    df_final = df_final.reset_index(drop=True)
                    
                    time.sleep(2)

                    fechar_menu_consulta(nav)

                    print("indo para transferencia de chapas")

                    wks2.update("E" + "5", 'TRANSF. CORTE: ' + data) 

                    c = 3

                    i = 0

                    menu_transf(nav)

                    time.sleep(2)

                    wks1, base, base_filtrada = planilha_corte_transf(data, filename)

                    if not len(df_final) == 0:

                        df_final['comparar2'] = ''
                            
                        for saldo in range(len(df_final)):
                            
                            try:
                                if df_final['Código Chapa'][saldo] == df_final['Código Chapa'][saldo-1]:
                                    df_final['comparar2'][saldo] = df_final['comparar2'][saldo-1] - df_final['Peso'][saldo]
                                
                                    if df_final['comparar2'][saldo] >= df_final['Peso'][saldo]:
                                        df_final['comparar'][saldo] = 'True'
                                    else:
                                        df_final['comparar'][saldo] = 'False'
                                else:
                                    df_final['comparar2'][saldo] = df_final['Saldo'][saldo] - df_final['Peso'][saldo]
                                    
                                    if df_final['Saldo'][saldo] >= df_final['Peso'][saldo]:
                                        df_final['comparar'][saldo] = 'True'
                                    else:
                                        df_final['comparar'][saldo] = 'False'
                            except:
                                df_final['comparar2'][saldo] = df_final['Saldo'][saldo] - df_final['Peso'][saldo]
                                
                                if df_final['Saldo'][saldo] >= df_final['Peso'][saldo]:
                                    df_final['comparar'][saldo] = 'True'
                                else:
                                    df_final['comparar'][saldo] = 'False'

                        df_final = df_final[df_final['comparar'] == 'True']
                        
                        for i in range(len(df_final)+1):
                            print("i: ", i)
                            try:
                                peca = df_final['Código Chapa'][i]
                                qtde = str(df_final['Peso'][i])
                                data = df_final['Data'][i]
                                c = preenchendo_corte_transf(nav,data,peca,qtde,wks1,c,i)   
                                time.sleep(1.5)         
                                print("c: ", c)

                                if c == 23:
                                    c = 21
                                
                                nav.delete_all_cookies() 

                            except:
                                pass

                        time.sleep(1.5)
                        
                        try:
                            selecionar_todos(data,nav)
                            
                            time.sleep(1.5)

                            for j in range(len(df_final)):
                                try:
                                    linha_transferida = df_final['index'][j]
                                    wks1.update("L" + str(linha_transferida+1), 'OK ROBS ' + data_hoje() + ' ' + hora_atual()) 
                                    time.sleep(1.5)
                                except:
                                    pass

                        except:
                            pass

                    fechar_menu_transf(nav)

                    ########### LOOP APONTAMENTOS ###########

                    time.sleep(2)

                    menu_apontamento(nav)

                    lista_ativadores = checkbox_apontamentos(filename)

                    if len(lista_ativadores[lista_ativadores['Setor'] == 'Serra']) > 0:

                        print('Indo para serra')

                        wks2.update("E" + "5", 'APONT. SERRA: ') 

                        wks1, base, base_filtrada, pessoa  = planilha_serra(filename, data)

                        base_filtrada = base_filtrada.reset_index()

                        if not len(base_filtrada) == 0:
                            
                            time.sleep(1)
                            nav.switch_to.default_content()
                            menu_innovaro(nav)
                            time.sleep(3)
                            
                            lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                            time.sleep(1.5)
                            click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                            
                            lista_menu[click_producao].click() ##clicando em producao
                            time.sleep(1.5)
                            
                            c = 3

                            i = 0

                            for i in range(len(base)+5): # serra
                                print("i: ", i)
                                nav.delete_all_cookies()
                                try:
                                    peca = base_filtrada['CÓDIGO'][i]
                                    qtde = str(base_filtrada['QNT'][i])
                                    data = base_filtrada['DATA'][i]
                                    mortas = base_filtrada['PEÇAS MORTA'][i]
                                    pessoa = pessoa
                                    linha = base_filtrada['index'][i]
                                    c = preenchendo_serra(nav,data,pessoa,peca,qtde,wks1,c,linha,mortas)
                                    time.sleep(1.5)
                                    print("c: ", c)
                                    if c == 23:
                                        c = 21

                                    nav.delete_all_cookies() 

                                except:
                                    break
                    
                    if len(lista_ativadores[lista_ativadores['Setor'] == 'Usinagem']) > 0:
                                
                        print('Indo para usinagem')

                        time.sleep(2)

                        wks2.update("E" + "5", 'APONT. USINAGEM: ') 

                        wks1, base, base_filtrada, pessoa  = planilha_usinagem(filename, data)
                        
                        base_filtrada = base_filtrada.reset_index()

                        if not len(base_filtrada) == 0:
                            
                            time.sleep(1)

                            nav.switch_to.default_content()
                            menu_innovaro(nav)
                            time.sleep(3)
                            
                            lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                            time.sleep(1.5)
                            click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                            
                            lista_menu[click_producao].click() ##clicando em producao
                            time.sleep(1.5)
                            
                            c = 3

                            i = 0

                            for i in range(len(base)+5):# usinagem
                                print("i: ", i)
                                nav.delete_all_cookies()
                                try:
                                    peca = base_filtrada['CÓDIGO'][i]
                                    qtde = str(base_filtrada['QNT'][i])
                                    data = base_filtrada['DATA'][i]
                                    mortas = base_filtrada['QNTD MORTA'][i]
                                    pessoa = pessoa
                                    linha = base_filtrada['index'][i]
                                    c = preenchendo_usinagem(nav,data,pessoa,peca,qtde,wks1,c,linha, mortas)
                                    time.sleep(1.5)
                                    print("c: ", c)
                                    if c == 23:
                                        c = 21
                                    
                                    nav.delete_all_cookies() 
                                
                                except:
                                    break

                    if len(lista_ativadores[lista_ativadores['Setor'] == 'Corte']) > 0:

                        print('indo para corte')

                        time.sleep(2)

                        wks2.update("E" + "5", 'APONT. CORTE: ') 

                        wks1, base, base_filtrada, pessoa  = planilha_corte(filename, data)
                        base_filtrada = base_filtrada.reset_index()

                        if not len(base_filtrada) == 0:
                            
                            time.sleep(1)

                            nav.switch_to.default_content()
                            menu_innovaro(nav)
                            time.sleep(3)
                            
                            lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                            time.sleep(1.5)
                            click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                            
                            lista_menu[click_producao].click() ##clicando em producao

                            time.sleep(1.5)
                            
                            c = 3

                            i = 1

                            for i in range(len(base)+5):
                                nav.delete_all_cookies()
                                print("i: ", i)
                                try:
                                    peca = base_filtrada['Peça'][i]
                                    qtde = str(base_filtrada['Total Prod.'][i])
                                    data = base_filtrada['Data finalização'][i]
                                    mortas = base_filtrada['Mortas'][i]
                                    pessoa = pessoa
                                    linha = base_filtrada['index'][i]
                                    c = preenchendo_corte(nav,data,pessoa,peca,qtde,wks1,c,linha, mortas)
                                    time.sleep(1.5)
                                    print("c: ", c)
                                    
                                    if c == 23:
                                        c = 21
                                    
                                    nav.delete_all_cookies() 

                                except:
                                    break

                    if len(lista_ativadores[lista_ativadores['Setor'] == 'Estamparia']) > 0:

                        print('Indo para estamparia')

                        time.sleep(2)

                        wks2.update("E" + "5", 'APONT. ESTAMPARIA: ') 

                        wks1, base, base_filtrada = planilha_estamparia(filename, data)
                        base_filtrada = base_filtrada.reset_index()
                        
                        if not len(base_filtrada) == 0:
                            
                            time.sleep(1)

                            nav.switch_to.default_content()
                            menu_innovaro(nav)
                            time.sleep(3)
                            
                            lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                            time.sleep(1.5)
                            click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                            
                            lista_menu[click_producao].click() ##clicando em producao
                            time.sleep(1.5)
                            
                            c = 3

                            i = 0

                            for i in range(len(base)+5):
                                nav.delete_all_cookies()
                                print("i: ", i)
                                try:
                                    peca = base_filtrada['CÓDIGO'][i]
                                    qtde = str(base_filtrada['QTD'][i])
                                    data = base_filtrada['DATA'][i]
                                    pessoa = base_filtrada['MATRÍCULA'][i]
                                    linha = base_filtrada['index'][i]
                                    c = preenchendo_estamparia(nav,data,pessoa,peca,qtde,wks1,c,linha)
                                    time.sleep(1.5)
                                    print("c: ", c)
                                    if c == 23:
                                        c = 21               

                                    nav.delete_all_cookies() 
                                
                                except:
                                    break

                    if len(lista_ativadores[lista_ativadores['Setor'] == 'Montagem']) > 0:                
                    
                        print('indo para montagem')

                        time.sleep(2)

                        wks2.update("E" + "5", 'APONT. MONTAGEM: ') 

                        wks1, base, base_filtrada = planilha_montagem(filename, data)
                        base_filtrada = base_filtrada.reset_index()

                        if not len(base_filtrada) == 0:

                            time.sleep(2)
                            nav.switch_to.default_content()
                            menu_innovaro(nav)
                            time.sleep(3)
                            
                            lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                            time.sleep(1.5)
                            click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                            
                            lista_menu[click_producao].click() ##clicando em producao
                            time.sleep(1.5)
                            
                            c = 3
                            i = 0

                            for i in range(len(base)+5):
                                print("i: ", i)
                                nav.delete_all_cookies()
                                try:
                                    peca = base_filtrada['CONJUNTO'][i]
                                    qtde = str(base_filtrada['QUANTIDADE'][i])
                                    data = base_filtrada['CARIMBO'][i]
                                    pessoa = base_filtrada['FUNCIONÁRIO'][i]
                                    linha = base_filtrada['index'][i]
                                    c = preenchendo_montagem(nav,data,pessoa,peca,qtde,wks1,c,linha)
                                    time.sleep(1.5)
                                    print("c: ", c)
                                    if c == 23:
                                        c = 21        

                                    nav.delete_all_cookies() 
                            
                                except:
                                    break
                        
                    if len(lista_ativadores[lista_ativadores['Setor'] == 'Pintura']) > 0:

                        print('indo para pintura')

                        time.sleep(2)            

                        wks2.update("E" + "5", 'APONT. PINTURA: ') 

                        wks1, base, base_filtrada, pessoa = planilha_pintura(filename, data)
                        
                        base_filtrada = base_filtrada.reset_index()

                        if not len(base_filtrada) == 0:

                            time.sleep(2)
                            nav.switch_to.default_content()
                            menu_innovaro(nav)
                            time.sleep(3)
                            
                            lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
                            time.sleep(2)
                            click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
                            
                            lista_menu[click_producao].click() ##clicando em producao
                            time.sleep(1.5)
                            
                            c = 3

                            i = 0

                            for i in range(len(base)+5):
                                print("i: ", i)
                                try:
                                    peca = base_filtrada['CÓDIGO'][i]
                                    qtde = str(base_filtrada['Qtd'][i])
                                    data = base_filtrada['Carimbo'][i]
                                    tipo = base_filtrada['Tipo'][i]
                                    cor = base_filtrada['Cor'][i]
                                    pessoa = pessoa
                                    linha = base_filtrada['index'][i]
                                    print(peca, '', '',qtde, '', '',data, '', '',pessoa)
                                    c = preenchendo_pintura(nav,data, pessoa, peca, qtde,tipo, cor, wks1, c, linha)
                                    time.sleep(1.5)
                                    print("c: ", c)
                                    if c == 23:
                                        c = 21       
                                    
                                    nav.delete_all_cookies() 
                                
                                except:
                                    break

                    time.sleep(2)            

                    try:
                        fechar_menu_apont(nav)
                    except:
                        pass

                    try:
                        fechar_tabs(nav)
                    except:
                        pass

        except:
            nav.close()

app = Flask(__name__)

# Variável para rastrear o estado de execução
execucao_em_andamento = False

# Função para executar o script Selenium
def executar_script():
    global execucao_em_andamento
    execucao_em_andamento = True

    # início do script

    funcao_main()

    # fim do script

    execucao_em_andamento = False

@app.route('/')
def index():
    return render_template('index.html', execucao_em_andamento=execucao_em_andamento)

@app.route('/acionar-script')
def acionar_script():
    global execucao_em_andamento

    # Verifica se o script já está em execução
    if not execucao_em_andamento:
        # Inicia um thread para executar o script (evita bloqueio do servidor Flask)
        thread = threading.Thread(target=executar_script)
        thread.start()

    return jsonify({'success': True})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5050)
