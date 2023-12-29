from selenium import webdriver
import threading
import time
from apontador import *

def executar_selenium(instancia):
    # Configuração do WebDriver (substitua 'chrome' por 'firefox' se preferir o Firefox)
    # driver = webdriver.Chrome(executable_path='chromedriver.exe')

    # Realize as operações desejadas na instância do WebDriver
    # driver.get(url)

    # Exemplo: Cada instância faz algo diferente com base no número da instância
    if instancia == 1:
        funcao_main()

    elif instancia == 2:
        funcao_transferencia()

    # Aguarde um pouco para visualizar a execução
    # time.sleep(5)

    # Feche o WebDriver
    # driver.quit()

# Crie duas threads para executar o Selenium concorrentemente
# url_instancia1 = 'https://www.exemplo1.com'
# url_instancia2 = 'https://www.exemplo2.com'

thread1 = threading.Thread(target=executar_selenium, args=(1,))
thread2 = threading.Thread(target=executar_selenium, args=(2,))

# Inicie as threads
thread1.start()
thread2.start()

# Aguarde até que ambas as threads tenham concluído
thread1.join()
thread2.join()
