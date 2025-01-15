from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
import time



# iniciar o chrome
s = Service(r"chromedriverPath.exe")
driver = webdriver.Chrome(service=s)
# abrir a aldo
driver.get("https://www.aldo.com.br")
# clicar em ENTRAR
driver.find_element("xpath", '//*[@id="header"]/div/div/div[2]/div/div[2]/ul/li[1]/a').click()

#esperar o campo de login aparecer e preencher
login = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#frm-login-uname")))
login.send_keys("username")

#esperar o campo senha aparecer e preencher
senha = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#frm-login-pass")))
senha.send_keys("password")

#esperar o campo entrar aparecer e clicar
WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.NAME, "submit"))).click()

WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="header"]/div/div/div[2]/div/div[2]/ul/li[3]/a')))


driver.find_element("xpath", '//*[@id="mercado_main"]/li[2]/a').click()

time.sleep(5)
button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/div/div[2]/div[2]/div/div/div/ul/li[3]/ul/li[3]/a')))
button.click()

time.sleep(5)
button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/div/div[2]/div[2]/div/div/div/ul/li[4]/ul/li[4]/a')))
button.click()

time.sleep(5)
button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/div/div[2]/div[2]/div/div/div/ul/li[5]/ul/li[2]/a')))
button.click()

time.sleep(5)
button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/div/div[2]/div[1]/div[4]/div[2]/div/button')))
driver.execute_script('arguments[0].click()', button)

#while True:
#    try :
#        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, carregarmais)))
#        driver.execute_script('arguments[0].click()', button)
#    except :
#        print("No more pages left")
#        break

time.sleep(5)
kits = WebDriverWait(driver, 20).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "product")))

produto = []

for kit in kits:
    nome = kit.find_element(By.CLASS_NAME, "product-name").text
    desc = kit.find_element(By.CLASS_NAME, "product-description").text
    preco = kit.find_element(By.CLASS_NAME, "product-price").text
    
    if 'ON GRID' in nome:
        if 'WALLBOX' not in nome:
            matriz = desc.split(' ')
            for row in matriz:
                row = row + ' '
                if 'KWP' in row:
                    idx_placa = row.find('K')
                    kwp = row[:idx_placa:]
                if 'KW ' in row:
                    idx_inv = row.find('K')
                    inv = row[:idx_inv:]
            lista = [kwp, inv, preco]
            produto.append(lista)

produto = sorted(produto, key = lambda item : item[0])

#for row in produto:
#    kwp = row[0]
#    idx = [x for x, i in enumerate(produto[0]) if i == kwp]
#    print (idx)

#print(produto)
