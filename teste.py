from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from datetime import datetime,timedelta 
import numpy as np
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


#options = webdriver.ChromeOptions()
#options.add_argument("--headless")

#link1 = "http://192.168.3.141/"
#link1 = 'http://cemag.innovaro.com.br/sistema'
link1 = 'http://devcemag.innovaro.com.br:81/sistema'
nav = webdriver.Chrome()#chrome_options=options)
time.sleep(2)
nav.get(link1)

try:
    while WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="content_statusMessageBox"]'))):
        print("carregando")
except:
    pass
try:
    WebDriverWait(nav,2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="errorMessageBox"]')))
    print('erro')
except:
    print('Funcionou')