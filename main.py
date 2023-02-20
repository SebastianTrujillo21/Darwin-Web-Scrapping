import time
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

username = "30030644"
password = "T0Ea2SxG"

filename = "DNI.xlsx"
filesheet = pd.read_excel(filename, usecols='A')
workbook = load_workbook(filename)
worksheet = workbook.active

dniArray = list(filesheet.loc[:, 'DNI'])
nombreArray = []
primerApellidoArray = []
segundoApellidoArray = []
correoArray = []
telefonoArray = []


def tarea():
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
    time.sleep(10)
    try:
        driver.find_element(By.XPATH, "//input[@value='Continue']").click()
    except NoSuchElementException:
        print("error")
    time.sleep(30)
    canal = driver.find_element(By.XPATH, "/html/body/div[1]/section/div[2]/div/div/div/div[1]/div/div/div/select")
    selector_canal = Select(canal)
    selector_canal.select_by_visible_text("Fuerza de Ventas con Verificacion")
    time.sleep(10)
    driver.find_element(By.XPATH, "/html/body/div[1]/section/div[2]/div/div/div/div[3]/button").click()
    time.sleep(10)
    for i in range(len(dniArray)):
        driver.find_element(By.XPATH,
                            "/html/body/div[1]/section/div[2]/div/div/div[3]/div[1]/div[2]/div[1]/div/div/div/div/form/div[1]/div/div[2]/div/div/input").send_keys(
            dniArray[i])
        time.sleep(5)
        driver.find_element(By.XPATH,
                            "/html/body/div[1]/section/div[2]/div/div/div[3]/div[1]/div[2]/div[1]/div/div/div/div/form/div[1]/div/div[3]/button").click()
        time.sleep(10)
        try:
            driver.find_element(By.XPATH,
                                "/html/body/div[6]/div[3]/div/div[2]/div[2]/div[1]/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div[1]/span/span/input").click()
        except NoSuchElementException:
            driver.refresh()
            continue

        time.sleep(5)
        try:
            driver.find_element(By.XPATH, "/html/body/div[6]/div[3]/div/div[2]/div[2]/div[2]/button").click()
        except NoSuchElementException:
            driver.refresh()
            continue

        time.sleep(5)
        try:
            nombre = driver.find_element(By.XPATH,
                                         "/html/body/div[1]/section/div[2]/div/div/div[3]/div[1]/div[2]/div[1]/div/div/div/div/form/div[1]/div/div[4]/div/div/input")
        except NoSuchElementException:
            driver.refresh()
            continue

        primer_apellido = driver.find_element(By.XPATH, "//input[@id='lastName']")
        segundo_apellido = driver.find_element(By.XPATH, "//input[@id='secondLastName']")
        correo = driver.find_element(By.XPATH, "//input[@id='email']")
        telefono = driver.find_element(By.XPATH, "//input[@id='phone']")

        text_nombre = nombre.get_attribute('value')
        text_Papellido = primer_apellido.get_attribute('value')
        text_Sapellido = segundo_apellido.get_attribute('value')
        text_correo = correo.get_attribute('value')
        text_telefono = telefono.get_attribute('value')

        nombreArray.append(text_nombre)
        primerApellidoArray.append(text_Papellido)
        segundoApellidoArray.append(text_Sapellido)
        correoArray.append(text_correo)
        telefonoArray.append(text_telefono)

        for a, value in enumerate(nombreArray):
            worksheet.cell(row=i+2, column=2, value=value)
        for a, value in enumerate(primerApellidoArray):
            worksheet.cell(row=i+2, column=3, value=value)
        for a, value in enumerate(segundoApellidoArray):
            worksheet.cell(row=i+2, column=4, value=value)
        for a, value in enumerate(correoArray):
            worksheet.cell(row=i+2, column=5, value=value)
        for a, value in enumerate(telefonoArray):
            worksheet.cell(row=i+2, column=6, value=value)

        workbook.save(filename)
        driver.refresh()


url = "https://checkout.naturgy.es/login"
driver = webdriver.Chrome()
driver.maximize_window()
driver.get(url)
time.sleep(20)
tarea()