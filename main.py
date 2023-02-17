import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

username = "30030644"
password = "T0Ea2SxG"
DNI = "52149074R"
url = "https://checkout.naturgy.es/login"
driver = webdriver.Chrome()
driver.maximize_window()
driver.get(url)
time.sleep(20)
driver.find_element(By.XPATH, "/html/body/div[4]/div[3]/div/div/div[2]/div/div/button").click()
time.sleep(10)
driver.find_element(By.XPATH, "/html/body/div[1]/section/div[2]/div/div/button").click()
time.sleep(10)
driver.find_element(By.XPATH,
                    "/html/body/div/div/div[2]/form/div/div/div/div/div[2]/div[2]/span/div/div/div/div/div/div/div/div/a/div[2]").click()
time.sleep(10)
driver.find_element(By.XPATH, "/html/body/div/div/div/div[2]/div[1]/form/input[3]").send_keys(username)
driver.find_element(By.XPATH, "/html/body/div/div/div/div[2]/div[1]/form/input[4]").send_keys(password)
time.sleep(10)
driver.find_element(By.XPATH, "/html/body/div/div/div/div[2]/div[1]/form/input[5]").click()
time.sleep(30)
canal = driver.find_element(By.XPATH, "/html/body/div[1]/section/div[2]/div/div/div/div[1]/div/div/div/select")
selector_canal = Select(canal)
selector_canal.select_by_visible_text("Fuerza de Ventas con Verificacion")
time.sleep(10)
driver.find_element(By.XPATH, "/html/body/div[1]/section/div[2]/div/div/div/div[3]/button").click()
time.sleep(10)
driver.find_element(By.XPATH,
                    "/html/body/div[1]/section/div[2]/div/div/div[3]/div[1]/div[2]/div[1]/div/div/div/div/form/div[1]/div/div[2]/div/div/input").send_keys(
    DNI)
time.sleep(5)
driver.find_element(By.XPATH,
                    "/html/body/div[1]/section/div[2]/div/div/div[3]/div[1]/div[2]/div[1]/div/div/div/div/form/div[1]/div/div[3]/button").click()
time.sleep(10)
driver.find_element(By.XPATH,
                    "/html/body/div[6]/div[3]/div/div[2]/div[2]/div[1]/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div[1]/span/span/input").click()
time.sleep(5)
driver.find_element(By.XPATH, "/html/body/div[6]/div[3]/div/div[2]/div[2]/div[2]/button").click()
nombre = driver.find_element(By.XPATH,
                             "/html/body/div[1]/section/div[2]/div/div/div[3]/div[1]/div[2]/div[1]/div/div/div/div/form/div[1]/div/div[4]/div/div/input")
text_nombre = nombre.get_attribute('value')

print(text_nombre)
time.sleep(20)
