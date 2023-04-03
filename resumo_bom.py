from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#from gspread_dataframe import set_with_dataframe
import time

link1 = "http://cemag.innovaro.com.br/sistema"
nav = webdriver.Chrome(r'C:\Users\pcp2\apontador\chromedriver.exe')
nav.get(link1)

#logando 
WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("Trainee - PCP")
WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("cem@1606")



WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/main/section/main/section[1]/form/button[1]'))).click()



while True:

    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/table/tbody/tr/td[5]/div/table/tbody/tr/td[2]'))).click()
    time.sleep(5)