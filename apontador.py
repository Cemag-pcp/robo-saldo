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

def dia_da_semana():
    
    today = datetime.now()
    today = today.isoweekday()
    return (today)

def data_sexta():
    data_sexta = datetime.now() - timedelta(3)
    ts = pd.Timestamp(data_sexta)
    data_sexta = data_sexta.strftime('%d/%m/%Y')

    return(data_sexta)

def data_sabado():
    data_sabado = datetime.now() - timedelta(2)
    ts = pd.Timestamp(data_sabado)
    data_sabado = data_sabado.strftime('%d/%m/%Y')

    return(data_sabado)

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

########### LOGIN ###########

def login(nav):
    #logando 
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("luan araujo")
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("luanaraujo1")

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

def menu_apontamento(nav):

    #clicando em produção
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[31]/span[2]'))).click()
    time.sleep(2)

    #clicando em SFC
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[37]/span[2]'))).click()
    time.sleep(2)

    #clicando em apontamento
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
    time.sleep(3)

def menu_transf(nav):

    #clicando em transferencia
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[24]/span[1]'))).click()

    time.sleep(2)

    #clicando em transf entre deposito
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[25]/span[2]'))).click()

    time.sleep(3)

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
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div"))).click()
    
    time.sleep(2)
    menu_innovaro(nav)
    time.sleep(2)

    #fecha consulta
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[1]'))).click()

def fechar_menu_transf(nav):
    
    try:
        nav.switch_to.default_content()
    except:
        pass

    #fecha aba de transf.
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div"))).click()
    
    time.sleep(1.5)
    menu_innovaro(nav)

    #fecha aba de consulta
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[24]/span[2]'))).click()

    time.sleep(1.5)
    #fecha estoque
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div[2]/div[4]/span[1]"))).click()

def fechar_menu_apont(nav):
    
    try:
        nav.switch_to.default_content()
    except:
        pass
    
    time.sleep(2)

    menu_innovaro(nav)

    time.sleep(1.5)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[36]/span[2]'))).click()
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[31]/span[2]'))).click()
    time.sleep(4)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/table/tbody/tr/td[1]/div/table/tbody/tr/td[2]"))).click()
    time.sleep(5)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div"))).click()
    time.sleep(2)

########### ACESSANDO PLANILHAS DE TRANSFERÊNCIA ###########

def planilha_serra_transf(data, filename):

    #SERRA#

    sheet = 'RQ PC-001-000 (APONTAMENTO SERRA)'
    worksheet1 = 'USO DO PCP'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(2)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1, inplace=False)[2:]

    ########### Tratando planilhas ###########

    #ajustando datas

    for i in range(2, len(base)+2):
        if base['DATA'][i] == "":
            base['DATA'][i] = base['DATA'][i-1]

    #filtrando peças que não foram apontadas
    base_filtrada  = base[base['APONTAMENTO'].isnull()]

    #filtrando data de hoje
    base_filtrada = base_filtrada.loc[base_filtrada.DATA == data]

    #filtrando por mat prima que não foi transferida
    base_filtrada  = base_filtrada[base_filtrada['TRANSF'].isnull()]

    #filtrando data de hoje
    #base_filtrada = base_filtrada.loc[base_filtrada.TRANSFERÊNCIA == '']

    base_filtrada =  base_filtrada[['DATA','MATERIAL','PESO BARRAS']]

    base_filtrada = base_filtrada[base_filtrada['MATERIAL'].notnull()]

    base_filtrada = base_filtrada[base_filtrada['PESO BARRAS'] != '']

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

        base_filtrada = base_filtrada.groupby(['DATA','MATERIAL']).sum().reset_index()

    base_filtrada = base_filtrada.loc[base_filtrada['PESO BARRAS'] > 0].reset_index(drop=True)

    return(wks1, base, base_filtrada, transferidas)

def planilha_corte_transf(data, filename):

    #CORTE#

    sheet = 'Banco de dados OP'
    worksheet1 = 'Transferência'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(5)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1, inplace=False)[2:]

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

    base_filtrada =  base_filtrada[['Data','Código Chapa','Peso']]
    
    if len(base_filtrada) > 0:

        for i in range(len(base)):
            try:
                if len(base_filtrada['Peso'][i]) > 1:
                    base_filtrada['Peso'][i] = base_filtrada['Peso'][i].replace(',','.')
            except:
                pass
        
        base_filtrada['Peso'] = base_filtrada['Peso'].astype(float)

        base_filtrada = base_filtrada.groupby(['Data','Código Chapa']).sum().reset_index()

    return(wks1, base, base_filtrada)

########### ACESSANDO PLANILHAS DE APONTAMENTO ###########

def planilha_serra(data, filename):

    #SERRA#

    sheet = 'RQ PC-001-000 (APONTAMENTO SERRA)'
    worksheet1 = 'USO DO PCP'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(2)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1, inplace=False)[2:]

    ########### Tratando planilhas ###########

    #ajustando datas
    for i in range(len(base)+2):
        try:
            if base['DATA'][i] == "":
                base['DATA'][i] = base['DATA'][i-1]
        except:
            pass

    i = None

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['APONTAMENTO'].isnull()]

    base_filtrada = base[base['APONTAMENTO'] == '']

    #filtrando data de hoje
    base_filtrada = base_filtrada.loc[base_filtrada.DATA == data]

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

    base_filtrada = base_filtrada[['DATA','CÓDIGO','QNT', 'APONTAMENTO']]

    #base_filtrada = base_filtrada.reset_index(drop=True)

    pessoa = '4209'

    return(wks1, base, base_filtrada, pessoa)

def planilha_usinagem(data, filename):

    #USINAGEM#

    sheet = 'RQ PC-001-000 (APONTAMENTO SERRA)'
    worksheet1 = 'USINAGEM 2022'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    headers = wks1.row_values(5)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:10]
    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    ########### Tratando planilhas ###########

    #ajustando datas
    for i in range(len(base)+5):
        try:
            if base['DATA'][i] == "":
                base['DATA'][i] = base['DATA'][i-1]
        except:
            pass

    i = None

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['PCP'].isnull()]

    #filtrando data de hoje
    base_filtrada = base.loc[base.DATA == data]

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

    base_filtrada = base_filtrada[['DATA','CÓDIGO','QNT']]

    pessoa = '4057'

    return(wks1, base, base_filtrada, pessoa)

def planilha_corte(data, filename):

    #CORTE#

    sheet = 'Banco de dados OP'
    worksheet1 = 'Finalizadas'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    #base = base.iloc

    headers = wks1.row_values(5)#[0:3]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    base['Data finalização'] = base['Data finalização'].str[:10]

    ########### Tratando planilhas ###########

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['PCP'].isnull()]

    #filtrando data de hoje
    base_filtrada = base.loc[base['Data finalização'] == data]

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

    base_filtrada = base_filtrada[['Data finalização','Peça','Total Prod.','Mortas']]

    pessoa = '4161'

    return(wks1, base, base_filtrada, pessoa)

def planilha_estamparia(data, filename):

    #ESTAMPARIA#

    sheet = 'RQ PC-003-000 (APONTAMENTO ESTAMPARIA)'
    worksheet1 = 'FICHA DE APONTAMENTO'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:10]

    headers = wks1.row_values(5)[0:10]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    ########### Tratando planilhas ###########

    #filtrando data de hoje
    base_filtrada = base.loc[base.DATA == data]

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

    base_filtrada = base_filtrada[['DATA','MATRÍCULA','CÓDIGO','QTD']]

    return(wks1, base, base_filtrada)

def planilha_montagem(data, filename):

    #MONTAGEM#

    sheet = 'RQ PC-004-000 APONTAMENTO MONTAGEM M22'
    worksheet1 = 'APONTAMENTO'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:9]

    headers = wks1.row_values(5)[0:9]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    ########### Tratando planilhas ###########
    
    base['CARIMBO'] = base['CARIMBO'].str[:10]

    base['FUNCIONÁRIO'] = base['FUNCIONÁRIO'].str[:4]

    base_filtrada = base.loc[base.CARIMBO == data]

    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace('-','_', regex=True)
    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace('-','_', regex=True)
    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace(' ','_', regex=True)
    base_filtrada['n_underscore'] = base_filtrada['CONJUNTO'].str.find('_')
    
    for i in range(len(base)+5):
        try:
            base_filtrada['CONJUNTO'][i] = base_filtrada['CONJUNTO'][i][:base_filtrada['n_underscore'][i]]
        except:
            pass

    base_filtrada = base_filtrada.fillna('')

    #filtrando linhas que não tem status de ok
    base_filtrada = base_filtrada[base_filtrada['STATUS'] == ''] 

    base_filtrada = base_filtrada[['CARIMBO','FUNCIONÁRIO','CONJUNTO','QUANTIDADE']]

    return(wks1, base, base_filtrada)

def planilha_sucata(data, filename):

    #SUCATA#

    sheet = 'Banco de dados OP'
    worksheet1 = 'Finalizadas'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    #base = base.iloc

    headers = wks1.row_values(5)#[0:3]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]

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

def planilha_pintura(data, filename):

    #PINTURA#

    sheet = 'BANCO DE DADOS ÚNICO - PINTURA'
    worksheet1 = 'RQ PC-005-003 (APONTAMENTO PINTURA)'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks1 = sh.worksheet(worksheet1)

    base = wks1.get()
    base = pd.DataFrame(base)
    #base = base.iloc

    headers = wks1.row_values(5)#[0:3]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    ########### Tratando planilhas ###########

    #filtrando data de hoje
    base_filtrada = base.loc[base['Carimbo'] == data]

    base_filtrada = base_filtrada.fillna('')

    #filtrando STATUS VAZIO
    base_filtrada = base_filtrada.loc[base_filtrada['STATUS'] == '']

    #filtrando linhas que estão ok na pintura
    base_filtrada = base_filtrada.loc[base_filtrada['PINTURA'] != '']

    base_filtrada = base_filtrada[['Carimbo','CÓDIGO','Qtd']]

    pessoa = '4271'

    return(wks1, base, base_filtrada, pessoa)

########### PREENCHIMENTO TRANSFERÊNCIA DE MP ###########

def preenchendo_serra_transf(data, peca, qtde, wks1, c, i):

    try:
        nav.switch_to.default_content()
    except:
        pass

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)
    
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

    #Deposito de destino
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[10]/div/input'))).send_keys('serra')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[10]/div/input'))).send_keys(Keys.TAB)

    #Recurso
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[12]/div/input'))).send_keys(peca)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[' + str(c) + ']/td[12]/div/input'))).send_keys(Keys.TAB)
    
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
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()

    try:
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
    except:
        pass

    time.sleep(2)

    c = c+2
    
    print(c)
    return(c)

def preenchendo_corte_transf(data, peca, qtde, wks1, c, i):

    try:
        nav.switch_to.default_content()
    except:
        pass

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()
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
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
    
    try:
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
    except:
        pass
    
    time.sleep(2)

    c = c+2
    
    print(c)
    return(c)

def selecionar_todos(nav):

    #selecinar todos os campos
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[1]/div'))).click()
    time.sleep(1)

    #aprovar
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/span[2]/p'))).click()
    time.sleep(1)

    try:
        nav.switch_to.default_content()
    except:
        pass

    #fechar pop-up
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button'))).click()

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)

    #baixar
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[3]/span[2]/p'))).click()
    time.sleep(1)

    #data 
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input'))).send_keys(data)
    time.sleep(1)

    try:
        nav.switch_to.default_content()
    except:
        pass

    #confirmar baixa
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span[2]/p'))).click()
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span[2]/p'))).click()
    time.sleep(10)
    
    #texto erro
    try:
        text_erro = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button'))).click()
    except:
        #gravar
        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span[2]/p'))).click()
        time.sleep(10)

########### PREENCHIMENTO APONTAMENTO DE PEÇA ###########

def preenchendo_serra(data, pessoa, peca, qtde, wks1, c, i):

    try:
        nav.switch_to.default_content()
    except:
        pass

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    
    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    try:
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
    except:
        pass

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
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
    except:

        #processo
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys('S')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)
        
        try:
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        except:

            #Etapa
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
            
            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)

            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            
            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
            
            nav.switch_to.default_content()

            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")
            
            try:
                # volta p janela principal (fora do iframe)

                texto_erro = WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
                WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
                wks1.update('S' + str(i+1), texto_erro + ' ' + data_hoje() + ' ' + hora_atual())
                time.sleep(2)
                
                nav.switch_to.default_content()
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()
                
                c = 3
            except:
                print('Funcionou')

                wks1.update('Q' + str(i+1), 'OK ROBINHO - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)

    return(c)

def preenchendo_usinagem(data, pessoa, peca, qtde, wks1, c, i):

    # hora = datetime.now()
    # hora = hora.strftime("%H:%M:%S")
    # wks1.update('E' + str(i+2), hora)

    try:
        nav.switch_to.default_content()
    except:
        pass

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    
    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    try:
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
    except:
        pass

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
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
    except:

        #processo
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys('S Usi')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)
        time.sleep(1)
        try:
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        except:

            #Etapa
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
            
            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)

            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)

            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()

            nav.switch_to.default_content()

            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")
            
            try:
                # volta p janela principal (fora do iframe)

                nav.switch_to.default_content()
                texto_erro = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
                wks1.update('J' + str(i+1), texto_erro + ' ' + data_hoje() + ' ' + hora_atual())

                time.sleep(2)
                nav.switch_to.default_content()
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()
                
                c = 3

            except:
                wks1.update('I' + str(i+1), 'OK ROBINHO - ' + ' ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)
    return(c)

def preenchendo_corte(data, pessoa, peca, qtde, wks1, c, i, mortas):

    try:
        nav.switch_to.default_content()
    except:
        pass

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    
    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    try:
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
    except:
        pass

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
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
    except:
        print('deu ruim')
        #processo
        try:
            iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
            nav.switch_to.frame(iframe1)
        except:
            print('deu ruim')

        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).click()
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys('S')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)

        #saindo do erro caso nao ache o processo
        try:
            nav.switch_to.default_content()
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        except:
            print('deu ruim')

            try:
                iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
                nav.switch_to.frame(iframe1)
            except:
                print('deu ruim')
                
            time.sleep(1)
            #Etapa
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
            time.sleep(1)
            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)
            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            time.sleep(3)

            #branco
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

            try:

                # volta p janela principal (fora do iframe)

                nav.switch_to.default_content()
                texto_erro = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
                wks1.update('O' + str(i+1), texto_erro + ' ' + data_hoje() + ' ' + hora_atual())
                
                time.sleep(2)
                nav.switch_to.default_content()
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()
                        
                c = 3

            except:
                wks1.update('L' + str(i+1), 'OK ROBINHO - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)
        
    return(c)

def preenchendo_estamparia(data, pessoa, peca, qtde, wks1, c, i):

    try:
        nav.switch_to.default_content()
    except:
        pass

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    
    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    try:
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
    except:
        pass

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
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
    except:
        print("deu ruim")

        try:
            iframe1 = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
            nav.switch_to.frame(iframe1)
        except:
            pass

        #processo
        WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys('S Est')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)

        try:
            nav.switch_to.default_content()
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        except:
            print("deu ruim")
            
            try:
                iframe1 = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
                nav.switch_to.frame(iframe1)
            except:
                pass

            #Etapa
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
            
            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)
            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()

            nav.switch_to.default_content()

            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")

            try:

                # volta p janela principal (fora do iframe)

                nav.switch_to.default_content()
                texto_erro = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
                wks1.update('K' + str(i+1), texto_erro + ' ' + data_hoje() + ' ' + hora_atual())

                time.sleep(2)
                nav.switch_to.default_content()
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()
                
                c = 3

            except:
                wks1.update('J' + str(i+1), 'OK ROBINHO - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)

    return(c)

def preenchendo_montagem(data, pessoa, peca, qtde, wks1, c, i):

    try:
        nav.switch_to.default_content()
    except:
        pass

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    
    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    try:
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
    except:
        pass

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
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
    except:

        #processo
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys('S Mont')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)
        
        try:
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        except:

            #Etapa
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
            
            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)
            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()
            
            nav.switch_to.default_content()
            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")

            try:

                # volta p janela principal (fora do iframe)

                texto_erro = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
                wks1.update('J' + str(i+1), texto_erro + '' + data_hoje() + ' ' + hora_atual())

                time.sleep(2)
                nav.switch_to.default_content()
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()
                
                c = 3

            except:
                wks1.update('I' + str(i+1), 'OK ROBINHO - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)

    return(c)

def preenchendo_pintura(data, pessoa, peca, qtde, wks1, c, i):

    try:
        nav.switch_to.default_content()
    except:
        pass

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)
    
    #Insert
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

    #chave
    try:
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/div/input'))).send_keys(Keys.TAB)
    except:
        pass
    
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).click()
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.CONTROL + 'a')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys('Produção por Máquina')
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    
    try:
        nav.switch_to.default_content()
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[1]/div[2]'))).click()
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys(Keys.TAB)
    except:
        pass
    
    try:
        iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
    except:
        pass

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
        time.sleep(2)
        nav.switch_to.default_content()
        time.sleep(1.5)
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
        time.sleep(1.5)
        webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
    except:
        print('deu ruim')
        
        try:
            iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
            nav.switch_to.frame(iframe1)
        except:
            pass
        
        #processo
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys('S Pint')
        time.sleep(1)
        WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys(Keys.TAB)

        try:
            time.sleep(3)
            nav.switch_to.default_content()
            time.sleep(1.5)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ESCAPE).perform()
            time.sleep(1.5)
            webdriver.ActionChains(nav).send_keys(Keys.ENTER).perform()
        except:

            try:
                iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
                nav.switch_to.frame(iframe1)
            except:
                pass

            #Etapa
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[14]/div/input"))).send_keys(Keys.TAB)
            
            #Máquina
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[16]/div/input"))).send_keys(Keys.TAB)
            time.sleep(3)

            #qtde
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
            time.sleep(3)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]"))).click()

            time.sleep(1)

            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div"))).click()

            nav.switch_to.default_content()

            try:
                while WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
                    print("Carregando")
            except:
                print("Carregou")

            try:

                # volta p janela principal (fora do iframe)

                nav.switch_to.default_content()
                texto_erro = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
                wks1.update('T' + str(i+1), texto_erro + '' + data_hoje() + ' ' + hora_atual())

                time.sleep(2)
                nav.switch_to.default_content()
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()
                
                c = 3

            except:
                wks1.update('O' + str(i+1), 'OK ROBS - ' + data_hoje() + ' ' + hora_atual())
                print('deu bom')
                c = c + 2

            print(c)
    return(c)

########### CONSULTAR SALDO ###########

def consulta_saldo(data, nav):
    
    time.sleep(0.5)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[4]/span[1]'))).click()
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[1]'))).click()
    time.sleep(0.5)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@title='-1897148410, Saldos de recursos.il']"))).click()
    time.sleep(1)

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    time.sleep(1)
    nav.switch_to.frame(iframe1)
    
    #data base
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(data)
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)
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

            qtde_itens = len(base_filtrada)

            for i in range(len(base_filtrada)):
                recurso = base_filtrada['MATERIAL'][i]
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(recurso + ';')
            
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)
            time.sleep(1)

            try:
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]/span[2]/p'))).click()
            except:
                pass

            try:
                nav.switch_to.default_content()
            except:
                pass

            time.sleep(2)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]/p'))).click()
            try:
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]/p'))).click()
            except:
                pass    
            
            #mudando iframe
            iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
            nav.switch_to.frame(iframe1)

            table_prod = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table')))
            table_html_prod = table_prod.get_attribute('outerHTML')
                
            tabelona = pd.read_html(str(table_html_prod), header=None)
            tabelona = tabelona[0]
            tabelona = tabelona.droplevel(level=0,axis=1)
            tabelona = tabelona.droplevel(level=0,axis=1)
            tabelona = tabelona[['Código','Saldo']]
            #tabelona = tabelona[['Unnamed: 0_level_2','Saldo']]
            #tabelona['Saldo'] = tabelona.Saldo.shift(-1)
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
            tabelona = tabelona.rename(columns={'Código':'MATERIAL'})
            
            df_final = pd.merge(tabelona,base_filtrada,on='MATERIAL')
            
            df_final['comparar'] = df_final['Saldo'] >= df_final['PESO BARRAS'] 

            df_final = df_final.loc[df_final['comparar'] == True]

        else:
            
            df_final = pd.DataFrame()
    except:
        df_final = pd.DataFrame()

    return(df_final)

def consulta_saldo_chapas(data, nav):
    
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[24]/span[2]'))).click()
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[1]'))).click()
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@title='-1897148410, Saldos de recursos.il']"))).click()
    time.sleep(2)
    
    #fechando aba anterior
    nav.switch_to.default_content()
    time.sleep(1)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tabs"]/td[1]/table/tbody/tr/td[4]/span/div'))).click()
    
    try:
        #nav.switch_to.default_content()
        iframe1 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe1)
    except:
        pass

    #data base
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(data)
    time.sleep(3)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)
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

            try:
                nav.switch_to.default_content()
            except:
                pass

            time.sleep(2)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]/p'))).click()
            try:
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]/p'))).click()
            except:
                pass    
            
            #mudando iframe
            iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
            nav.switch_to.frame(iframe1)

            table_prod = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table')))
            table_html_prod = table_prod.get_attribute('outerHTML')
                
            tabelona = pd.read_html(str(table_html_prod), header=None)
            tabelona = tabelona[0]
            tabelona = tabelona.droplevel(level=0,axis=1)
            tabelona = tabelona.droplevel(level=0,axis=1)
            tabelona = tabelona[['Código','Saldo']]
            #tabelona = tabelona[['Unnamed: 0_level_2','Saldo']]
            #tabelona['Saldo'] = tabelona.Saldo.shift(-1)
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
            tabelona = tabelona.rename(columns={'Código':'Código Chapa'})

            df_final = pd.merge(tabelona,base_filtrada,on='Código Chapa')
            
            df_final['comparar'] = df_final['Saldo'] >= df_final['Peso'] 

            df_final = df_final.loc[df_final['comparar'] == True]

        else:
            
            df_final = pd.DataFrame()
    except:
        df_final = pd.DataFrame()

    return(df_final)

########## LOOP ###########

while 'a' == 'a':

    try:

        w = 0

        today = dia_da_semana()

        if today != 1:

            datas = [data_hoje(), data_sabado(), data_sexta()]
        
        else:

            datas = [data_hoje(), data_sabado(), data_sexta()]

        nav = acessar_innovaro()

        time.sleep(4)

        login(nav)

        while w < 3:

            for d in range(len(datas)):

                data = datas[d]
                #data = data_hoje()
                #data = '01/02/2023'
                
                ########## CONSULTAR SALDO ###########

                print("Verificando saldo da serra")
                
                time.sleep(2)
                
                menu_innovaro(nav)

                time.sleep(1)

                df_final = consulta_saldo(data, nav)

                time.sleep(2)

                fechar_menu_consulta(nav)

                ########## LOOP TRANSFERÊNCIA ###########

                print("Indo para transferencia de tubos")

                time.sleep(2)

                menu_transf(nav)

                wks1, base, base_filtrada, transferidas = planilha_serra_transf(data, filename)
                
                transferidas = transferidas.reset_index()

                c = 3

                i = 0

                if not len(df_final) == 0:

                    if not int(len(transferidas)) == 0:

                            for i in range(len(base)+1): # serra

                                print("i: ", i)
                                try:
                                    peca = df_final['MATERIAL'][i]
                                    qtde = str(df_final['PESO BARRAS'][i])
                                    data = df_final['DATA'][i]
                                    c = preenchendo_serra_transf(data,peca,qtde,wks1,c,i)            
                                    print("c: ", c)
                                    j = 0

                                    for j in range(len(transferidas)):
                                        try:
                                            filtrado = transferidas.loc[transferidas.MATERIAL == peca]
                                            ok = filtrado['index'][j]
                                            wks1.update("R" + str(ok+1), 'OK TRANSF - ' + data_hoje() + ' ' + hora_atual()) 
                                        except:
                                            pass
                                except:
                                    pass
                            
                            try:
                                selecionar_todos(nav)
                            except:
                                pass

                ########## CONSULTAR SALDO CORTE ###########

                print("Verificando saldo de corte")
                
                time.sleep(2)
                
                menu_innovaro(nav)

                time.sleep(1)

                df_final = consulta_saldo_chapas(data, nav)

                time.sleep(2)

                fechar_menu_consulta(nav)

                print("indo para transferencia de chapas")

                c = 3

                i = 0

                menu_transf_2(nav)

                wks1, base, base_filtrada = planilha_corte_transf(data, filename)

                if not len(df_final) == 0:

                    for i in range(len(base)+1): # serra

                        print("i: ", i)
                        try:
                            peca = df_final['Código Chapa'][i]
                            qtde = str(df_final['Peso'][i])
                            data = df_final['Data'][i]
                            c = preenchendo_corte_transf(data,peca,qtde,wks1,c,i)            
                            print("c: ", c)
                            j = 0

                            for j in range(len(base)):
                                try:
                                    filtrado = base.loc[base['Código Chapa'] == peca]
                                    filtrado = filtrado.loc[filtrado.Data == data]
                                    filtrado = filtrado.loc[filtrado.Status == '']
                                    filtrado = filtrado.reset_index()
                                    ok = filtrado['index'][j]
                                    wks1.update("L" + str(ok+1), 'OK ROBS ' + data_hoje() + ' ' + hora_atual()) 
                                except:
                                    pass
                        except:
                            pass
                    
                    try:
                        selecionar_todos(nav)
                    except:
                        pass

                fechar_menu_transf(nav)

                ########### LOOP APONTAMENTOS ###########

                time.sleep(2)

                menu_apontamento(nav)

                print('Indo para serra')

                wks1, base, base_filtrada, pessoa  = planilha_serra(data, filename)

                if not len(base_filtrada) == 0:

                    nav.switch_to.default_content()
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()

                    c = 3

                    i = 0

                    for i in range(len(base)+5): # serra

                        print("i: ", i)
                        try:
                            peca = base_filtrada['CÓDIGO'][i]
                            qtde = str(base_filtrada['QNT'][i])
                            data = base_filtrada['DATA'][i]
                            pessoa = pessoa
                            c = preenchendo_serra(data,pessoa,peca,qtde,wks1,c,i)
                            print("c: ", c)
                        except:
                            pass

                print('Indo para usinagem')

                wks1, base, base_filtrada, pessoa  = planilha_usinagem(data, filename)

                if not len(base_filtrada) == 0:

                    nav.switch_to.default_content()
                    time.sleep(2)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()

                    c = 3

                    i = 0

                    for i in range(len(base)+5):# usinagem
                        
                        print("i: ", i)
                        try:
                            peca = base_filtrada['CÓDIGO'][i]
                            qtde = str(base_filtrada['QNT'][i])
                            data = base_filtrada['DATA'][i]
                            pessoa = pessoa
                            c = preenchendo_usinagem(data,pessoa,peca,qtde,wks1,c,i)
                            print("c: ", c)
                        except:
                            pass

                print('indo para corte')

                wks1, base, base_filtrada, pessoa  = planilha_corte(data, filename)

                if not len(base_filtrada) == 0:

                    time.sleep(2)
                    nav.switch_to.default_content()
                    time.sleep(2)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()

                    c = 3

                    i = 0

                    for i in range(len(base)+5):
                        
                        print("i: ", i)
                        try:
                            peca = base_filtrada['Peça'][i]
                            qtde = str(base_filtrada['Total Prod.'][i])
                            data = base_filtrada['Data finalização'][i]
                            mortas = base_filtrada['Mortas'][i]
                            pessoa = pessoa
                            c = preenchendo_corte(data,pessoa,peca,qtde,wks1,c,i, mortas)
                            print("c: ", c)
                        except:
                            pass

                print('Indo para estamparia')

                wks1, base, base_filtrada  = planilha_estamparia(data, filename)

                if not len(base_filtrada) == 0:

                    time.sleep(2)
                    nav.switch_to.default_content()
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()

                    c = 3

                    i = 0

                    for i in range(len(base)+5):
                        
                        print("i: ", i)
                        try:
                            peca = base_filtrada['CÓDIGO'][i]
                            qtde = str(base_filtrada['QTD'][i])
                            data = base_filtrada['DATA'][i]
                            pessoa = base_filtrada['MATRÍCULA'][i]
                            c = preenchendo_estamparia(data,pessoa,peca,qtde,wks1,c,i)
                            print("c: ", c)
                        except:
                            pass

                print('indo para montagem')

                wks1, base, base_filtrada = planilha_montagem(data, filename)

                if not len(base_filtrada) == 0:

                    time.sleep(2)
                    nav.switch_to.default_content()
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()

                    c = 3
                    i = 0

                    for i in range(len(base)+5):
                        
                        print("i: ", i)
                        try:
                            peca = base_filtrada['CONJUNTO'][i]
                            qtde = str(base_filtrada['QUANTIDADE'][i])
                            data = base_filtrada['CARIMBO'][i]
                            pessoa = base_filtrada['FUNCIONÁRIO'][i]
                            c = preenchendo_montagem(data,pessoa,peca,qtde,wks1,c,i)
                            print("c: ", c)
                        except:
                            pass

                time.sleep(2)            

                print('indo para pintura')

                wks1, base, base_filtrada, pessoa = planilha_pintura(data, filename)

                if not len(base_filtrada) == 0:

                    time.sleep(2)
                    nav.switch_to.default_content()
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[39]/span[2]'))).click()
                    time.sleep(3)
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()

                    c = 3

                    i = 0

                    for i in range(len(base)+5):
                        
                        print("i: ", i)
                        try:
                            peca = base_filtrada['CÓDIGO'][i]
                            qtde = str(base_filtrada['Qtd'][i])
                            data = base_filtrada['Carimbo'][i]
                            pessoa = pessoa
                            c = preenchendo_pintura(data,pessoa,peca,qtde,wks1,c,i)
                            print("c: ", c)
                        except:
                            pass

                time.sleep(2)            

                fechar_menu_apont(nav)
    
    except:
        pass