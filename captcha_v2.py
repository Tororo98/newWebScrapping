# Nuevo codigo para automatizar busqueda de censo con Documentos activos/inactivos
# Sebastian Toro
# 1. Usar pip install para instalar pandas.
# 2. Tambien se puede usar pip install para xlrd, sirve igual. (requisito en ambos, archivo en formato "xls")
# 3. Usar pip install para xlswriter
# 4. Usar pip install para xlwt
# 5. Utilizar un archivo existente en formato xlsx para crear uno nuevo en xls
# Falta realizar el co.add(zenVPN.crx)
# Cambiar el webdriver actual por el de chrome anterior.
# Falta la funcion de escribir en el doc pero esa la saco yo.
# 6. pip install pywin32

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.proxy import Proxy, ProxyType
from selenium.webdriver.chrome import service
from collections import deque

import os, sys
import time,requests
from bs4 import BeautifulSoup
# import pywin32

import pandas as pd
import xlrd
import xlsxwriter
import xlwt

global maximumIterations
maximumIterations = 3

delayTime = 2
audioToTextDelay = 10
filename = 'test.mp3'
byPassUrl = 'https://wsp.registraduria.gov.co/censo/consultar/'
# googleIBMLink = 'https://speech-to-text-demo.ng.bluemix.net/'

# option = webdriver.ChromeOptions()
# option.add_argument('--disable-notifications')
# option.add_argument("--mute-audio")
# # option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
# option.add_argument("user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 10_3 like Mac OS X) AppleWebKit/602.1.50 (KHTML, like Gecko) CriOS/56.0.2924.75 Mobile/14E5239e Safari/602.1")

def createDoc():
    path = 'censo.xls'
    path2 = 'nuevo_censo.xls'

    # Leyendo archivo viejo
    inputWorkBook = xlrd.open_workbook(path)
    inputWorkSheet = inputWorkBook.sheet_by_index(0)

    # Creando un nuevo archivo
    outWorkBook = xlsxwriter.Workbook(path2)
    outSheet = outWorkBook.add_worksheet()

    maxIteration = maximumIterations
    i = 0
    allRows = []

    # Copiar los datos del archivo viejo
    for row in range(2, maxIteration):
        # Las coordenadas en excel son: (y,x) en vez de (x,y)
        currRow = inputWorkSheet.cell_value(row,0)
        allRows.append(currRow)

    # Pegando los datos al archivo nuevo
    for newRow in range(2, maxIteration):
        outSheet.write(newRow, 0, allRows[i])
        i+=1

    outWorkBook.close()


def readDoc(row):
    path = 'censo.xls'

    # Leyendo archivo viejo
    inputWorkBook = xlrd.open_workbook(path)
    inputWorkSheet = inputWorkBook.sheet_by_index(0)

    # Obtener ID dentro del archivo
    currRow = inputWorkSheet.cell_value(row,0)

    return currRow


def writeDoc(row):
    pass
        

def audioToText(mp3Path):

    driver.execute_script('''window.open("","_blank");''')
    driver.switch_to.window(driver.window_handles[1])

    driver.get(googleIBMLink)

    # Upload file 
    time.sleep(delayTime)
    # cookie = driver.find_element_by_
    root = driver.find_element_by_id('root').find_elements_by_class_name('dropzone _container _container_large')
    btn = driver.find_element(By.XPATH, '//*[@id="root"]/div/input')
    btn.send_keys(mp3Path)

    # Audio to text is processing
    time.sleep(audioToTextDelay)

    # accessing to the text
    text = driver.find_element_by_class_name('tab-panels--tab-content').text
    result = text

    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    return result


def saveFile(content,filename):
    with open(filename, "wb") as handle:
        for data in content.iter_content():
            handle.write(data)


def captcha(row, driver):
    # driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=option)
    # driver.get(byPassUrl)

    try:
        # Ingresar ID al campo
        numberField = driver.find_element_by_id('nuip')        
        currValue = readDoc(row)        
        numberField.send_keys(str(int(currValue)))
        numberField.submit()
    except Exception as e:
        print("No puedo ingresar el id. No halle el campo.")

    # Accediendo al frame de captcha para hacer click en el cuadro de verify
    googleClass = driver.find_elements_by_class_name('g-recaptcha')[0]
    outeriframe = googleClass.find_element_by_tag_name('iframe')
    outeriframe.click()

    allIframesLen = driver.find_elements_by_tag_name('iframe')
    audioBtnFound = False
    audioBtnIndex = -1

    for index in range(len(allIframesLen)):
        # Buscando el frame donde estan las opciones de solucionar el captcha
        driver.switch_to_default_content()
        iframe = driver.find_elements_by_tag_name('iframe')[index]
        driver.switch_to.frame(iframe)
        driver.implicitly_wait(delayTime)
        try:
            # Click en el boton del audio
            audioBtn = driver.find_element_by_id('recaptcha-audio-button') or driver.find_element_by_id('recaptcha-anchor')
            audioBtn.click()
            audioBtnFound = True
            audioBtnIndex = index
            break
        except Exception as e:
            print("Couldn't find the button.")
            pass

    if audioBtnFound:
        try:
            while True:
                # Descargar audio del challenge
                href = driver.find_element_by_id('audio-source').get_attribute('src')
                response = requests.get(href, stream=True)
                saveFile(response,filename)
                response = audioToText(os.getcwd() + '/' + filename)
                print(response)

                # Reubicandonos en el frame de introducir texto
                driver.switch_to_default_content()
                iframe = driver.find_elements_by_tag_name('iframe')[audioBtnIndex]
                driver.switch_to.frame(iframe)

                # Ingresar texto
                inputbtn = driver.find_element_by_id('audio-response')
                inputbtn.send_keys(response)
                inputbtn.send_keys(Keys.ENTER)

                # Verificar si acepto el texto
                time.sleep(2)
                errorMsg = driver.find_elements_by_class_name('rc-audiochallenge-error-message')[0]

                if errorMsg.text == "" or errorMsg.value_of_css_property('display') == 'none':
                    print("Success")
                    break
                 
        except Exception as e:
            print(e)
            print('Something went wrong with the answer. Or captcha was blocked')
    else:
        print('Button not found. This should not happen.')


def get_proxies():
    co = Options()
    co.add_argument("log-level=3")
    co.add_argument("--headless")
    co.add_argument('--disable-notifications')
    
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options = co)
    driver.get("https://sslproxies.org/")

    #Pile of proxies is created
    PROXIES = deque()
    proxies = driver.find_elements_by_css_selector("tr[role='row']")
    for p in proxies:
        result = p.text.split(" ")
        temporal_index = len(result)
        if result[temporal_index-1] == "yes":
            PROXIES.append(result[0]+":"+result[1])

    driver.close()
    return PROXIES       


def proxy_driver():
    global ALL_PROXIES, my_ip

    co = Options()
    prox = Proxy()

    if len(ALL_PROXIES) == 0:
        print("--- Proxies used up (%s)" % len(ALL_PROXIES))
        ALL_PROXIES = get_proxies()
        
    # temporal_index = len(ALL_PROXIES)
    # Accessing and removing last element of deque
    else:
        pxy = ALL_PROXIES.pop()
        my_ip = pxy
        print('Proxy Actual:', pxy)

        prox.proxy_type = ProxyType.MANUAL
        prox.autodetect = False
        prox.httpProxy = prox.sslProxy = pxy #prox.socksProxy = pxy

        capabilities = webdriver.DesiredCapabilities.CHROME
        prox.add_to_capabilities(capabilities)

        #print('Proxy Options', prox)
        co.Proxy = prox
        co.add_argument("ignore-certificate-errors")

        co.add_argument("start-maximized")
        co.add_experimental_option("excludeSwitches", ["enable-automation"])
        co.add_experimental_option('useAutomationExtension', True)
        ua = UserAgent()
        userAgent = ua.chrome
        co.add_argument(f'user-agent={userAgent}')
        co.add_argument('--disable-notifications')
        
        # Se agrega el add-on Buster para validar los captchas
        co.add_extension('./buster_extension.crx')
        co.add_extension('./vpn.crx')

    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options = co)
    # driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    return driver    


def main():
    global ALL_PROXIES
    ALL_PROXIES = []
    myRow = 2
    createDoc()

    for i in range(0, 1000):
        driver = proxy_driver()      
        driver.get(byPassUrl)
        captcha(myRow, driver)
    
        driver.quit()


main()