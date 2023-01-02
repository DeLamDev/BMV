import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import xlrd
import xlwt
from xlutils.copy import copy

def refresh(driver):
    formato = driver.find_elements(By.XPATH, "//span[@class='value']")
    lista_opciones = driver.find_elements(By.XPATH, "//ul[@class='options']")
    return formato, lista_opciones

def main_choice(form):
    for e in range(len(form)):
        if form[e].text != "":
            print(f"{e}: {form[e].text}")
    return input("Opción deseada: ")

def verificador(elec):
    try:
        return int(elec)
    except Exception as e:
        print("Favor de ingresar el número de la opción únicamente.")
        print(e)
        exit()

def sub_choice(lista, elec):
    sub = lista[elec].find_elements(By.XPATH, "li/label")
    for e in range(len(sub)):
        print(f"{e}: {amp_fixer(sub[e].get_attribute('innerHTML'))}")
    return input("Opción deseada: ")

def final_download(form, driver, link):
    print("Comenzando descarga...")
    search = driver.find_element(By.XPATH, "//input[@id='btnSearch']")
    search.click()
    sleep(3)
    constructor = ['/en/Grupo_BMV/Informacion_de_emisora/_rid/541/_mto/3/_mod/doDownload',
                   '?idTipoMercado=', '&idTipoInstrumento=','&idTipoEmpresa=','&idSector=',
                   '&idSubsector=','&idRamo=','&idSubramo=','&random=9493']
    choices = []
    final_link = constructor[0]
    counter = 0
    lista = driver.find_elements(By.XPATH, "//ul[@class='options']")
    for e in lista:
        valores = e.find_elements(By.XPATH, "li")
        for i in valores:
            if amp_fixer(i.find_element(By.XPATH, "label").get_attribute("innerHTML")) == form[counter].text:
                choices.append(i.find_element(By.XPATH, "input").get_attribute('value'))
            if form[counter].text == "":
                choices.append("")
        counter += 1
    v3, v1, v4, v2 = choices[1], choices[2], choices[3], choices[4]
    choices[1], choices[2], choices[3], choices[4] = v1, v2, v3, v4
    choices.pop()
    for n in range(len(choices)):
        final_link = final_link + constructor[n + 1] + choices[n]
    final_link = final_link + constructor[-1]
    descarga(link + final_link)
    print("Terminado!")

def amp_fixer(broken):
    return broken.replace("amp;", "")

def descarga(url):
    local = "listaEmpresas.xls"
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

def verificar_exito():
    try:
        book = xlrd.open_workbook("listaEmpresas.xls")
        sh = book.sheet_by_index(0)
        return sh
    except Exception as e:
        print(e)
        print("Error en la descarga atribuible a la selección de opciones.")
        ticker = "TICKER SYMBOL"
        issuer = "ISSUER'S NAME"
        book = xlwt.Workbook()
        ws = book.add_sheet('DATA')
        ws.write(0, 0, ticker)
        ws.write(0, 1, issuer)
        book.save('listaEmpresas.xls')
        return False

# extraer las empresas del listado de la página y añadirlas al excel 
def fixer(driver):
    search = driver.find_element(By.XPATH, "//input[@id='btnSearch']")
    search.click()
    print("Arreglando documento manualmente...")
    sleep(2)
    table = driver.find_element(By.XPATH, "//tbody[@class='pages']")
    sub = table.find_elements(By.XPATH, "tr/td")
    book = xlrd.open_workbook("listaEmpresas.xls")
    wb = copy(book)
    s = wb.get_sheet(0)
    count = 1
    for i in range(0, len(sub), 2):
        full = sub[i + 1].text
        tic = sub[i].text
        s.write(count, 0, tic)
        s.write(count, 1, full)
        count += 1
    wb.save("listaEmpresas.xls")
    adder(driver)

# extraer los perfiles y añadirlos al lado de cada empresa
def adder(driver):
    perfiles = driver.find_elements(By.XPATH, "//tbody[@class='pages']/tr/td/a")
    book = xlrd.open_workbook("listaEmpresas.xls")
    wb = copy(book)
    s = wb.get_sheet(0)
    count = 1
    for i in perfiles:
        if '%' not in i.get_attribute('href'):
            s.write(count, 2, i.get_attribute('href'))
        count += 1
    wb.save("listaEmpresas.xls")

def lista_empresas():
    driver = webdriver.Firefox()
    driver.get("https://www.bmv.com.mx/en/issuers/issuers-information")
    link = "https://www.bmv.com.mx"
    while True:
        print("Seleccione alguna de las opciones o escriba 'listo' en su lugar para descargar la lista.")
        formato, lista = refresh(driver)
        eleccion = main_choice(formato)
        if eleccion.strip().lower() == "listo":
            break
        eleccion = verificador(eleccion)
        formato[eleccion].click()
        sleep(2)
        sub_eleccion = sub_choice(lista, eleccion)
        if sub_eleccion.strip().lower() == "listo":
            break
        sub_eleccion = verificador(sub_eleccion)
        sub = lista[eleccion].find_elements(By.XPATH, "li/label")
        sleep(2)
        sub[sub_eleccion].click()
        sleep(2)
    formato, _ =refresh(driver)
    final_download(formato, driver, link)
    sh = verificar_exito()
    if sh == False:
        fixer(driver)
    else:
        adder(driver)
    print("Listado completado.")

