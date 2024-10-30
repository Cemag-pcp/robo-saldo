from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
import pandas as pd
from selenium.webdriver.common.keys import Keys
import time
import gspread
from selenium.webdriver.common.action_chains import ActionChains
import os
import glob
import datetime
from google.oauth2 import service_account
from flask import Flask, render_template

import subprocess
import zipfile
import requests

# https://googlechromelabs.github.io/chrome-for-testing/#stable
# Defiindo o tempo máximo de espera pelo elemento
wait = 20

app = Flask(__name__)

# # Configurando o logger para gerar logs
# logging.basicConfig(filename=r'C:\Users\Engine\automacao_saldo_almoc\automation_log.log', level=logging.INFO,
#                     format='%(asctime)s - %(levelname)s - %(message)s')


def acessar_innovaro():

    try:

        link1 = "http://192.168.3.141/"
        # link1 = 'http://cemag.innovaro.com.br/sistema'
        # link1 = 'http://devcemag.innovaro.com.br:81/sistema'
        nav = webdriver.Chrome(r'C:\Users\Engine\robo-saldo\robo-saldo\chromedriver_extracted\chromedriver-win32\chromedriver.exe')
        # nav = webdriver.Chrome(r'C:\Users\Engine\chromedriver.exe')
        # nav = webdriver.Chrome(r'C:\Users\Engine\chromedriver.exe')

        nav.maximize_window()
        nav.get(link1)

        logging.info("Acessou a página de login")

    except Exception as e:
        logging.error(f"Ocorreu um erro no inicio: {e}")

    return (nav)

def login(nav):

    try:
        # logando
        WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="username"]'))).send_keys("Ti.prod")
        WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys("Cem@@1600")
        WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

        logging.info("Acessou a página de login")

    except Exception as e:
        logging.error(f"Ocorreu um erro durante o login: {e}")


def menu_innovaro(nav):

    try:
        # abrindo menu

        try:
            nav.switch_to.default_content()
        except:
            pass

        WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()

        logging.info("Clicou no menu")

    except Exception as e:
        logging.error(f"Ocorreu um erro durante o click no botão de menu: {e}")


def listar(nav, classe):

    try:

        lista_menu = nav.find_elements(By.CLASS_NAME, classe)

        elementos_menu = []

        for x in range(len(lista_menu)):
            a = lista_menu[x].text
            elementos_menu.append(a)

        test_lista = pd.DataFrame(elementos_menu)
        test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

        logging.info("listou as opções do menu")

    except Exception as e:
        logging.error(f"Ocorreu um erro durante a listagem de opções: {e}")

    return (lista_menu, test_lista)


def iframes(nav):

    try:

        iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

        for iframe in range(len(iframe_list)):
            try:
                nav.switch_to.default_content()
                nav.switch_to.frame(iframe_list[iframe])
                print(iframe)
            except:
                pass

        logging.info("Navegou entre os iframes existentes")

    except Exception as e:
        logging.error(f"Ocorreu um erro durante a mudanca de iframe: {e}")


def input_data(nav):

    try:
        data_base = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
        time.sleep(0.5)
        data_base.send_keys(Keys.CONTROL + 'a')
        time.sleep(0.5)
        data_base.send_keys('h')
        time.sleep(0.5)
        data_base.send_keys(Keys.ENTER)

        logging.info("Navegou entre os iframes existentes")

    except Exception as e:
        logging.error(f"Ocorreu um erro durante o click no botão de menu: {e}")


def input_deposito(nav):

    try:

        deposito = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[8]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
        time.sleep(0.5)
        deposito.send_keys(Keys.CONTROL + 'a')
        time.sleep(0.5)
        deposito.send_keys('Almox Central')
        time.sleep(0.5)
        deposito.send_keys(Keys.TAB)
        time.sleep(0.5)
        deposito.send_keys(Keys.TAB)
        time.sleep(0.5)

        logging.info("Inputando o depósito")

    except Exception as e:
        logging.error(f"Ocorreu um erro no input do depósito: {e}")


def input_deposito_levantamento(nav):

    try:

        deposito = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[8]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
        time.sleep(0.5)
        deposito.send_keys(Keys.CONTROL + 'a')
        time.sleep(0.5)
        deposito.send_keys('Almox Mont Carretas;Almox Cx Acessórios;Almox Pintura - Embalagem;Almox Prod Especiais;Almox Corte e Estamparia;Almox Serra;Almox Usinagem')
        time.sleep(0.5)
        deposito.send_keys(Keys.TAB)
        time.sleep(0.5)
        
        try:
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]/span[2]'))).click()
            time.sleep(0.5)
            deposito.send_keys(Keys.TAB)
            time.sleep(0.5)
        except:
            pass

        logging.info("Inputando o depósito")

    except Exception as e:
        logging.error(f"Ocorreu um erro no input do depósito: {e}")

def input_deposito_apenas_central(nav):

    try:

        deposito = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[8]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
        time.sleep(0.5)
        deposito.send_keys(Keys.CONTROL + 'a')
        time.sleep(0.5)
        deposito.send_keys('Almox Central')
        time.sleep(0.5)
        deposito.send_keys(Keys.TAB)
        time.sleep(0.5)
        
        try:
            WebDriverWait(nav, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]/span[2]'))).click()
            time.sleep(0.5)
            deposito.send_keys(Keys.TAB)
            time.sleep(0.5)
        except:
            pass

        logging.info("Inputando o depósito")

    except Exception as e:
        logging.error(f"Ocorreu um erro no input do depósito: {e}")


def limpar_recursos(nav):

    try:
        recursos = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
        time.sleep(0.5)
        recursos.send_keys(Keys.CONTROL + 'a')
        time.sleep(0.5)
        recursos.send_keys(Keys.BACKSPACE)
        time.sleep(0.5)
        recursos.send_keys(Keys.TAB)
        time.sleep(0.5)

        logging.info("Limpou os recursos caso existissem")

    except Exception as e:
        logging.error(f"Ocorreu um erro limpeza dos recursos: {e}")


def inserir_agrupamentos_levantamento(nav):

    iframes(nav)
    agrupamento_1 = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[20]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')))
    time.sleep(0.5)
    agrupamento_1.send_keys(Keys.CONTROL + 'a')
    time.sleep(0.5)
    agrupamento_1.send_keys('Etapa')
    time.sleep(0.5)
    agrupamento_1.send_keys(Keys.TAB)
    time.sleep(0.5)


def inserir_agrupamentos_central_mp(nav):

    iframes(nav)
    agrupamento_1 = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[20]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')))
    time.sleep(0.5)
    agrupamento_1.send_keys(Keys.CONTROL + 'a')
    time.sleep(0.5)
    agrupamento_1.send_keys('Etapa')
    time.sleep(0.5)
    agrupamento_1.send_keys(Keys.TAB)
    time.sleep(0.5)


def inserir_agrupamentos_almox(nav):

    agrupamento_1 = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[20]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')))
    time.sleep(0.5)
    agrupamento_1.send_keys(Keys.CONTROL + 'a')
    time.sleep(0.5)
    agrupamento_1.send_keys('Classe de Recursos')
    time.sleep(0.5)
    agrupamento_1.send_keys(Keys.TAB)
    time.sleep(0.5)


def guardado_codigos():

    scope = ['https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive"]

    # credentials = service_account.Credentials.from_service_account_file(r'C:\Users\Engine\automacao_saldo_almoc\service_account_cemag.json', scopes=scope)
    credentials = service_account.Credentials.from_service_account_file('service_account.json',scopes=scope)

    client = gspread.authorize(credentials)
    # filename = r'C:\Users\Engine\automacao_saldo_almoc\service_account_cemag.json'
    filename = 'service_account.json'

    sa = gspread.service_account(filename)

    spreadsheet_id = '1dE89oIzFrldjr6cDk01lbVxFopxcGSvguTSnYHZ3fw4'
    worksheet = 'Ajuste de Estoque - Almoxafirado'

    sh = client.open_by_key(spreadsheet_id)
    wks1 = sh.worksheet(worksheet)

    cell_value = wks1.cell(1, 17).value

    return cell_value


def inputar_recurso(nav):

    try:
        iframes(nav)

        recursos = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
        time.sleep(0.5)
        recursos.click()
        time.sleep(0.5)
        codigos = guardado_codigos()

        recursos.send_keys(Keys.CONTROL, 'a')
        time.sleep(1)
        recursos.send_keys(Keys.DELETE)

        recursos.send_keys('')

        time.sleep(0.5)
        recursos.send_keys(Keys.TAB)

        time.sleep(20)

        recursos.send_keys(Keys.CONTROL, Keys.SHIFT + 't')

        try:
            WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]/span[2]"))).click()
        except:
            pass

        logging.info("Inputou recursos da planilha")

        texto = ''

        recursos = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))

        time.sleep(3)

        # Clicar em exportar
        recursos.send_keys(Keys.CONTROL, Keys.SHIFT + 'x')

        # Objeto WebDriverWait para aguardar elementos
        wait_webdriver = WebDriverWait(nav, 60)

        saida_iframe(nav)

        try:
            while WebDriverWait(nav, wait).until(EC.presence_of_element_located(
                (By.ID, "statusMessageBox"))):

                print("carregando")
        except:
            print('carregou')

        # iframes(nav)

        # clicar em exportar
        wait_webdriver.until(EC.presence_of_element_located(
            (By.XPATH, '//*[starts-with(@id, "buttonsContainer_")]'))).click()

       # wait_webdriver.until(EC.presence_of_element_located(
       #     (By.ID, '_lbl_dadosGeradosComSucessoSelecionaAOpcaoExportarParaSelecionarOFormatoDeExportacao')))

        
        # if texto != '':

            # js_code = """
            # var element = document.querySelector(".button.default");
            # element.click();
            # """
            
            # nav.execute_script(js_code)

        logging.info("Clicou em exportar")

        # opção de download
        wait_webdriver.until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[8]/table/tbody/tr/td[2]/div/div/div[2]'))).click()
        
        logging.info("Escolheu opções de download")

        time.sleep(2)

        iframes(nav)

        botao_exec = wait_webdriver.until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
        time.sleep(1)
        botao_exec.send_keys(Keys.CONTROL, Keys.SHIFT + 'e')

        try:
            WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[4]/div[2]/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]"))).click()
        except:
            pass

        texto_final = wait_webdriver.until(
            EC.presence_of_element_located((By.ID, '_lbl__instructions')))

        if texto_final:

            botao_exec = wait_webdriver.until(
                EC.presence_of_element_located((By.ID, '_download_elt')))
            time.sleep(1)
            botao_exec.click()

            logging.info("Tarefa finalizada com sucesso")
        
        # else:
        #     logging.error(
        #         f"Ocorreu um erro ao clicar em exportar dentro do else")

    except Exception as e:
        logging.error(f"Ocorreu um erro inputar recursos: {e}")


def exportar(nav):

    try:
        texto = ''

        iframes(nav)

        recursos = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))

        time.sleep(1)
        recursos.send_keys(Keys.CONTROL, Keys.SHIFT + 'x')

        # Objeto WebDriverWait para aguardar elementos
        wait_webdriver = WebDriverWait(nav, 20)

        texto = wait_webdriver.until(EC.presence_of_element_located(
            (By.ID, '_lbl_dadosGeradosComSucessoSelecionaAOpcaoExportarParaSelecionarOFormatoDeExportacao')))

        if texto != '':

            saida_iframe(nav)

            # Localiza o elemento tr pelo id
            button_element = wait_webdriver.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'span.button.default')))

            # Cria uma instância de ActionChains
            actions = ActionChains(nav)

            # Pressiona as teclas Shift + Ctrl + X no elemento do botão
            actions.move_to_element(button_element).key_down(Keys.SHIFT).key_down(
                Keys.CONTROL).send_keys('x').key_up(Keys.SHIFT).key_up(Keys.CONTROL).perform()

            # opção de download
            wait_webdriver.until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[9]/table/tbody/tr/td[2]/div/div/div[2]'))).click()

            time.sleep(2)

            iframes(nav)

            botao_exec = wait_webdriver.until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
            time.sleep(1)
            botao_exec.send_keys(Keys.CONTROL, Keys.SHIFT + 'e')
            botao_exec.send_keys(Keys.CONTROL, Keys.SHIFT + 'e')

            texto_final = wait_webdriver.until(
                EC.presence_of_element_located((By.ID, '_lbl__instructions')))

            if texto_final:

                botao_exec = wait_webdriver.until(
                    EC.presence_of_element_located((By.ID, '_download_elt')))
                time.sleep(1)
                botao_exec.click()

                logging.info("Clicou em exportar")

        else:
            logging.error(
                f"Ocorreu um erro ao clicar em exportar dentro do else")

    except Exception as e:
        logging.error(f"Ocorreu um erro ao clicar em exportar: {e}")


def exportar_2(nav):

    saida_iframe(nav)

    # exportar primeira tela
    WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span[2]'))).click()
    
    try:
        WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span[2]'))).click()
    except:
        pass

    loading = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="content_statusMessageBox"]')))

    while loading:
        print('carregando')
        
        try:
            loading = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="content_statusMessageBox"]')))
        except:
            loading = None
            break
    
    # loading_2 = WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
    #     (By.XPATH, '//*[@id="progressMessageBox"]')))
    
    # while loading_2:
    #     print('carregando2')
        
    #     try:
    #         loading_2 = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
    #             (By.XPATH, '//*[@id="progressMessageBox"]')))
    #     except:
    #         loading_2 = None
    #         break

    # exportar da segunda tela
    WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
        (By.XPATH, "//span[@class='wf-button default' and contains(., 'Exportar')]"))).click()
    
    try:
        WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, "//span[@class='wf-button default' and contains(., 'Exportar')]"))).click()
    except:
        pass

    # opção de excel
    WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[8]/table/tbody/tr/td[2]/div/div/div[2]'))).click()

    # executar
    WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[4]/div[2]/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]'))).click()
    
    try:
        WebDriverWait(nav, wait).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[4]/div[2]/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]'))).click()
    except:
        pass

    # clique aqui para fazer download do arquivo
    iframes(nav)
    time.sleep(10)

    botao_download = WebDriverWait(nav, 999).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/span')))

    botao_download.click()


def saida_iframe(nav):
    try:
        nav.switch_to.default_content()
        logging.info("Saiu do iframe")

    except Exception as e:
        logging.error("Erro ao sair do iframe")


def ultimo_arquivo():

    # Caminho para a pasta "Downloads"
    caminho_downloads = os.path.expanduser("~") + "/Downloads"

    # Lista de todos os arquivos na pasta "Downloads", ordenados por data de modificação (o mais recente primeiro)
    lista_arquivos = glob.glob(caminho_downloads + "/*", recursive=False)
    lista_arquivos.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    # Pegue o caminho do último arquivo baixado (o arquivo mais recente)
    ultimo_arquivo_baixado = lista_arquivos[0]

    print("Caminho do último arquivo baixado:", ultimo_arquivo_baixado)

    df = pd.read_csv(ultimo_arquivo_baixado, sep=';', encoding='latin-1')
    
    df['data'] = datetime.datetime.today()
    df['data'] = df['data'].dt.strftime('%d/%m/%Y %H:%M:%S')

    df.rename(columns=lambda x: x.replace('="', '').replace('"', ''), inplace=True)

    # df = df.drop(columns={'=" "'})
    # df.columns
    df["1o. Agrupamento"] = df["1o. Agrupamento"].apply(lambda x: str(x).replace("=", "").replace('"', ''))
    df["2o. Agrupamento"] = df["2o. Agrupamento"].apply(lambda x: str(x).replace("=", "").replace('"', ''))
    
    return df


def inserir_gspread():

    df = ultimo_arquivo()

    # Autentique-se com a API do Google Sheets (configure o caminho para suas credenciais)
    # gc = gspread.service_account(filename=r'C:\Users\Engine\automacao_saldo_almoc\service_account_cemag.json')
    gc = gspread.service_account(filename='service_account.json')

    # Abra a planilha com base no ID
    planilha = gc.open_by_key("1dE89oIzFrldjr6cDk01lbVxFopxcGSvguTSnYHZ3fw4")

    # Acessar a aba "BD_saldo_diario"
    aba = planilha.worksheet("BD_saldo_diario")

    # Defina o intervalo (range) que você deseja apagar (por exemplo, A2:H5)
    range_to_clear = "A2:H"
    
    # Obtém a lista de células no intervalo especificado
    cell_list = aba.range(range_to_clear)
    
    # Define o valor de todas as células no intervalo como uma string vazia ('')
    for cell in cell_list:
        cell.value = ""
    
    # Atualiza as células no intervalo com os valores vazios
    aba.update_cells(cell_list)
    
    df.rename(columns=lambda x: x.replace('="', '').replace('"', ''), inplace=True)
    
    df = df.applymap(lambda x: str(x).replace('="', '').replace('"', ''))

    df = df[['3o. Agrupamento',	'2o. Agrupamento','Recurso#Unid. Medida','Saldo','Custo#Total','Custo#Médio','data']]
    
    df['codigo'] = df['3o. Agrupamento'].apply(lambda x: x.split()[0])
    df['descricao'] = df['3o. Agrupamento'].apply(lambda x: x.split('-')[1])
    
    df_values = df.values.tolist()
    
    planilha.values_append("BD_saldo_diario", {'valueInputOption': 'RAW'}, {'values': df_values})

    return 'sucess'


def inserir_gspread_saldo_levantamento():

    df = ultimo_arquivo()

    # Autentique-se com a API do Google Sheets (configure o caminho para suas credenciais)
    # gc = gspread.service_account(filename=r'C:\Users\Engine\automacao_saldo_almoc\service_account_cemag.json')
    gc = gspread.service_account(filename='service_account.json')

    # Abra a planilha com base no ID
    planilha = gc.open_by_key("1u2Iza-ocp6ROUBXG9GpfHvEJwLHuW7F2uiO583qqLIE")

    # Acessar a aba "BD_saldo_diario"
    aba = planilha.worksheet("saldo de recurso")

    # Defina o intervalo (range) que você deseja apagar (por exemplo, A2:H5)
    range_to_clear = "A2:Z"
    
    df.rename(columns=lambda x: x.replace('="', '').replace('"', ''), inplace=True)
    
    df = df.applymap(lambda x: str(x).replace('="', '').replace('"', ''))

    df['Saldo'] = df['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
    df['Custo#Total'] = df['Custo#Total'].apply(lambda x: float(x.replace(".","").replace(",",".")))
    df['Custo#Médio'] = df['Custo#Médio'].apply(lambda x: float(x.replace(".","").replace(",",".")))
    
    df = df[df['2o. Agrupamento'] == 'nan']

    df['codigo'] = df['3o. Agrupamento'].apply(lambda x: x.split()[0])
    df['descricao'] = df['3o. Agrupamento'].apply(lambda x: x.split('-')[1])
    
    df = df[['1o. Agrupamento','codigo','descricao','3o. Agrupamento','Saldo','Custo#Total','Custo#Médio','data']]
    df_values = df.values.tolist()

    if len(df_values) > 0:
    
        # Obtém a lista de células no intervalo especificado
        cell_list = aba.range(range_to_clear)
        
        # Define o valor de todas as células no intervalo como uma string vazia ('')
        for cell in cell_list:
            cell.value = ""
        
        # Atualiza as células no intervalo com os valores vazios
        aba.update_cells(cell_list)

        planilha.values_append("saldo de recurso", {'valueInputOption': 'RAW'}, {'values': df_values})

    return 'sucess'

def inserir_gspread_saldo_central_mp():

    df = ultimo_arquivo()

    # Autentique-se com a API do Google Sheets (configure o caminho para suas credenciais)
    # gc = gspread.service_account(filename=r'C:\Users\Engine\automacao_saldo_almoc\service_account_cemag.json')
    gc = gspread.service_account(filename='service_account.json')

    # Abra a planilha com base no ID
    planilha = gc.open_by_key("1u2Iza-ocp6ROUBXG9GpfHvEJwLHuW7F2uiO583qqLIE")

    # Acessar a aba "BD_saldo_diario"
    aba = planilha.worksheet("saldo central")

    # Defina o intervalo (range) que você deseja apagar (por exemplo, A2:H5)
    range_to_clear = "A2:Z"
    
    df.rename(columns=lambda x: x.replace('="', '').replace('"', ''), inplace=True)
    
    df = df.applymap(lambda x: str(x).replace('="', '').replace('"', ''))

    df['Saldo'] = df['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
    df['Custo#Total'] = df['Custo#Total'].apply(lambda x: float(x.replace(".","").replace(",",".")))
    df['Custo#Médio'] = df['Custo#Médio'].apply(lambda x: float(x.replace(".","").replace(",",".")))
    
    df['codigo'] = df['3o. Agrupamento'].apply(lambda x: x.split()[0])
    df['descricao'] = df['3o. Agrupamento'].apply(lambda x: x.split('-')[1])

    df = df[['1o. Agrupamento', 'codigo', 'descricao', 'Saldo','Custo#Total', 'Custo#Médio','data']]

    df_values = df.values.tolist()

    if len(df_values) > 0:
        
        # Obtém a lista de células no intervalo especificado
        cell_list = aba.range(range_to_clear)
        
        # Define o valor de todas as células no intervalo como uma string vazia ('')
        for cell in cell_list:
            cell.value = ""
        
        # Atualiza as células no intervalo com os valores vazios
        aba.update_cells(cell_list)
        
        planilha.values_append("saldo central", {'valueInputOption': 'RAW'}, {'values': df_values})

    return 'sucess'

def navegar_consulta(nav):
    
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(0.5)
    click_producao = test_list.loc[test_list[0] == 'Estoque'].reset_index(drop=True)[
        'index'][0]
    lista_menu[click_producao].click()
    
    time.sleep(1)
    
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(0.5)
    click_producao = test_list.loc[test_list[0] == 'Consultas'].reset_index(drop=True)[
        'index'][0]
    lista_menu[click_producao].click()
    
    time.sleep(1)
    
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(0.5)
    click_producao = test_list.loc[test_list[0] == 'Saldos de Recursos - CEMAG'].reset_index(drop=True)[
        'index'][0]
    lista_menu[click_producao].click()


def add_classe_recurso(nav):

    deposito = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[15]/td[2]/table/tbody/tr/td[1]/input')))
    time.sleep(0.5)
    deposito.send_keys(Keys.CONTROL + 'a')
    time.sleep(0.5)
    deposito.send_keys('Mat Prima')
    time.sleep(0.5)
    deposito.send_keys(Keys.TAB)
    time.sleep(0.5)

def apagar_mat_prima(nav):

    deposito = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[15]/td[2]/table/tbody/tr/td[1]/input')))
    time.sleep(0.5)
    deposito.send_keys(Keys.CONTROL + 'a')
    time.sleep(0.5)
    deposito.send_keys(Keys.DELETE)
    deposito.send_keys(Keys.BACKSPACE)
    time.sleep(0.5)
    deposito.send_keys(Keys.TAB)
    time.sleep(0.5)

def apagar_classe_recurso(nav):

    deposito = WebDriverWait(nav, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[20]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')))
    time.sleep(0.5)
    deposito.send_keys(Keys.CONTROL + 'a')
    time.sleep(0.5)
    deposito.send_keys(Keys.BACKSPACE)
    time.sleep(0.5)
    deposito.send_keys(Keys.TAB)
    time.sleep(0.5)

# webdriver automático

def verificar_chrome_driver():
    # URL que você deseja acessar
    url = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"

    # Faz uma solicitação GET à URL
    response = requests.get(url)

    if response.status_code == 200:
        data = response.json()

        # Versão que você está procurando
        target_version = get_chrome_version()

        def version_key(version):
            # Converte a versão em uma lista de inteiros para comparação
            return list(map(int, version.split('.')))

        # Filtra as versões disponíveis
        versions = [v['version'] for v in data['versions']]

        # Encontra a versão mais próxima
        closest_version = None
        min_diff = float('inf')

        for version in versions:
            # Calcula a diferença em cada nível de versão
            current_diff = sum(
                abs(a - b) for a, b in zip(version_key(version), version_key(target_version)))

            if current_diff < min_diff:
                min_diff = current_diff
                closest_version = version

        # Filtra as versões disponíveis
        versions_info = data['versions']

        # Armazena versões que correspondem à versão desejada
        matching_versions = []

        for version_info in versions_info:
            version = version_info['version']

            # Verifica se a versão corresponde à versão desejada
            if version == closest_version:
                matching_versions.append(version_info)

        # Exibe todos os registros encontrados
        if matching_versions:
            print("Registros encontrados:")
            for match in matching_versions:
                download_chrome_driver = match['downloads']['chromedriver']
        else:
            print("Nenhum registro encontrado.")

        def get_url_by_platform(data, platform):
            for item in data:
                if item['platform'] == platform:
                    return item['url']
            return None

        # Exemplo de uso:
        url = get_url_by_platform(download_chrome_driver, 'win32')
        filename = 'chromedriver-win32.zip'
        download_file(url, filename)

        extract_to = 'chromedriver_extracted'
        unzip_file(filename, extract_to)

        chromedriver_path = find_chromedriver(extract_to)
        if chromedriver_path:
            print(f"chromedriver.exe encontrado em: {chromedriver_path}")
        else:
            print("chromedriver.exe não encontrado.")

        return chromedriver_path

    else:
        print(f"Falha ao acessar a URL: {response.status_code}")

def download_file(url, filename):
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, 'wb') as file:
            file.write(response.content)
        print(f"Download concluído: {filename}")
    else:
        print(f"Falha no download: {response.status_code}")

def unzip_file(zip_path, extract_to='.'):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
        print(f"Arquivos extraídos para: {extract_to}")

def find_chromedriver(extract_to='.'):
    for root, dirs, files in os.walk(extract_to):
        if 'chromedriver.exe' in files:
            return os.path.join(root, 'chromedriver.exe')
    return None

def get_chrome_version():
    try:
        # Executa o comando para obter a versão do Chrome
        version = subprocess.check_output(
            ['reg', 'query', r'HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon',
                '/v', 'version'],
            stderr=subprocess.STDOUT
        )
        # Processa a saída para extrair a versão
        version = version.decode().strip().split()[-1]
        return version
    except subprocess.CalledProcessError:
        return "Chrome não está instalado."


# ______ Inicio _______

# @app.route('/')
# def hello_world():
#     return render_template("index.html")

# # botão
# @app.route('/acionar-robo')
# def acionar_robo():
#     try:
#         logging.info("Iniciou")
        
#         nav = acessar_innovaro()
        
#         login(nav)
#         time.sleep(1)
        
#         menu_innovaro(nav)
#         time.sleep(1)
        
#         iframes(nav)
#         time.sleep(1)
        
#         input_data(nav)
#         time.sleep(1)
        
#         input_deposito(nav)
#         time.sleep(1)
        
#         limpar_recursos(nav)
#         time.sleep(1)
        
#         inputar_recurso(nav)
#         time.sleep(5)
        
#         inserir_gspread()
        
#         nav.close()
        
#     except:
#         nav.close()
    
#     return 'Robô rodou'


# if __name__ == '__main__':
    # app.run(host='0.0.0.0', port=5000)