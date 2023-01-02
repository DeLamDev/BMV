#!/home/habib/environments/main_venv/bin/python3

# Requests para descargar pagina de la BMV
# bs4 para buscar dento de la página descargada
# docx para crear el documento de word
# datetime para obtener la fecha del día de hoy

import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date, datetime
from extr_empresas import lista_empresas
import xlrd
from xlutils.copy import copy
import xlwt

meses = ['enero', 'febrero', 'marzo', 'abril', 
         'mayo', 'junio', 'julio', 'agosto', 
         'septiembre', 'octubre', 'noviembre', 
         'diciembre']

def fecha_actual():
    return date.today()


def fecha_elegida():
    print("¿De qué fecha desea extraer los eventos?\n")
    print("Ingrese la fecha con el formato: día-mes-año (ejem. 31-01-2000)")
    print("Si desea la fecha de hoy favor de escribir 'h'.")
    fecha_select = input("Fecha: ")
    return fecha_select
    

def validar_fecha(fecha, meses):
    format = "%d-%m-%Y"
    se_fecha = fecha.strip()
    if fecha.strip().lower() == 'h':
        print(f"Usted ingresó: {fecha_actual().strftime(format)}\n")
        return fecha_actual(), meses[int(fecha_actual().strftime('%m')) - 1]
    try:
        se_fecha = datetime.strptime(se_fecha, format)
        print(f"Usted ingresó: {fecha}\n")
        return se_fecha, meses[int(se_fecha.strftime('%m')) -1]
    except Exception as e:
        print("No ha ingresado la fecha con el formato correcto.")
        print("Recuerde ingresar el día, mes y año separados por un '-' sin espacios entre ellos.")
        print("Ejemplo: 01-10-2000 (primero de octubre del año 2000).\n")
        print(e)
        quit()

def titulo_y_sub():
    titulo = input("Favor de ingresar título del documento: ")
    subtitulo = input("Favor de ingresar subtítulo del documento: ")
    print(f"El título elegido es: {titulo}")
    print(f"El subtítulo elegido es: {subtitulo}")
    return titulo, subtitulo

def links_EvRel(perfil):
    base = "https://www.bmv.com.mx"
    r = requests.get(perfil)
    soup = BeautifulSoup(r.text, "html.parser")
    for e in soup.find_all("div", class_='tabs-area')[0].find_all('a'):
        if 'relevantevents' in e.get('href'):
            return base + e.get('href')

def extractorXls():
    book = xlrd.open_workbook("listaEmpresas.xls")
    sh = book.sheet_by_index(0)
    final_dict = {}
    for r in range(1, sh.nrows):
        final_dict[sh.cell_value(rowx=r, colx=0)] = []
        for c in range(1, sh.ncols):
            final_dict[sh.cell_value(rowx=r, colx=0)].append(sh.cell_value(rowx=r, colx=c))
    return final_dict

def relv_adder():
    b = xlrd.open_workbook("listaEmpresas.xls")
    book = copy(b)
    sh = book.get_sheet(0)
    counter = 1
    dict = extractorXls()
    for e in dict:
        sh.write(counter, 3, links_EvRel(dict[e][1]))
        counter += 1
    book.save('listaEmpresas.xls')

def doc_creator(fecha, mes, title, subtitle, dict):
    documento = Document()
    fecha_doc = documento.add_paragraph(fecha.strftime('%d') + ' de ' + mes + ' de ' + fecha.strftime('%Y'))
    fecha_doc.runs[0].bold = True
    fecha_doc.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    titulo = documento.add_paragraph(title)
    if title != "":
        titulo.runs[0].bold = True
    documento.add_paragraph(subtitle)
    tabla = documento.add_table(rows=len(dict) + 1, cols=5)
    tabla.style = 'TableGrid'
    filas = tabla.rows[0].cells
    filas[0].text = "Número"
    filas[1].text = "Clave emisora"
    filas[2].text = "Razón social"
    filas[3].text = "Eventos relevantes"
    for n in range(1, len(dict) + 1):
        tabla.rows[n].cells[0].text = str(n)
        tabla.rows[n].cells[1].text = list(dict.keys())[n - 1]
        tabla.rows[n].cells[2].text = dict[list(dict.keys())[n - 1]][0]
    documento.save("test.docx")

def doc_updater():
    pass

def loop_events(tables, number):
    busqueda = {}
    for event in tables[number].find_all("tr"):
        for count, data in enumerate(event.find_all("td")):
            if count == 0:
                busqueda[data.text] = []
            if count == 2:
                busqueda[list(busqueda)[-1]].append([link["href"] for link in data.find_all("a")])
            elif count == 1:
                busqueda[list(busqueda)[-1]].append(data.text)
    return busqueda

def rel_event_extractor(link):
    re = requests.get(link)
    soup = BeautifulSoup(re.text, "html.parser")
    tables = soup.find_all("tbody")
    results = []
    if len(tables) > 1:
        for e in range(1, len(tables)):
            results.append(loop_events(tables, e))
        return results
    else:
        return None

def searcher(dict, fecha):
    lista_word = []
    for empresa in dict:
        lista_word.append({empresa: []})
        resultados = rel_event_extractor(dict[empresa][2])
        if resultados == None:
            continue
        for e in resultados:
            for i in e:
                if i.split()[0] == fecha.strftime("%d-%m-%Y"):
                    print("Eureka")
                    print(i.split()[0])
                    print(e[i])
                    lista_word[-1][empresa].append(i)
                    lista_word[-1][empresa].extend(e[i])
    return lista_word


def checker_link_ER():
    book = xlrd.open_workbook('listaEmpresas.xls')
    sh = book.sheet_by_index(0)
    for e in range(1, sh.nrows):
        if sh.cell_value(rowx=e, colx=3) == "":
            print("Faltan links en el excel.")
            print("Favor de utilizar la función para completarlos o hacerlo de forma manual.")
            exit()




"""
for i in range(1, 101):
    print('Procesando empresa ' + str(i) + ' de 100')
    fila = tabla.rows[i - 1].cells
    fila[0].text = str(i)
    if i <= len(entidades_solas):
        fechas, asuntos, links, nombre = extraer_info(empresas_bmv[i -1])
        fila[1].text = entidades_solas[i - 1]
        fila[2].text = nombre
        if len(fechas) == len(asuntos) and len(fechas) == len(links):
            for f in range(len(fechas)):
                if fechas[f][:10] == fecha_actual().strftime('%d-%m-%Y'):
                    texto = fila[3].add_paragraph()
                    texto.add_run('ASUNTO').bold = True
                    texto.add_run('\n' + asuntos[f] + '\n')
                    texto.add_run('EVENTO RELEVANTE').bold = True
                    texto.add_run('\n' + 'https://www.bmv.com.mx' + links[f] + '\n')
                elif fila[3].text == '' and f == (len(fechas) - 1):
                    fila[3].text = 'Sin publicación'
        else:
            print('ERROR!!!!!!!!!!')
    else:
        fila[1].text = 'A completar'
        fila[2].text = 'A completar'


documento.save('prueba_bmx.docx')
"""
if __name__ == "__main__":
    print("Extractor de eventos relevantes de empresas de la BMV.\n\n")
    while True:
        elec_fecha = fecha_elegida()
        fecha, mes = validar_fecha(elec_fecha, meses)
        proseguir = input("Confirma la fecha (s/n): ")
        if proseguir == "s":
            break
        elif proseguir == "n":
            continue
        else:
            print("Favor de contestar solo 's' o 'n'.\n")
            exit()
    elec_lista = input("Si no desea crear una lista de empresas escriba 'n': ")
    if elec_lista == 'n':
        pass
    else:
        lista_empresas()
    titulo, subtitulo = titulo_y_sub()
    dec_EVREL = input("Desea añadir los links de eventos relevantes (s/n): ")
    if dec_EVREL.strip().lower() == "s":
        relv_adder()
        new_dict = extractorXls()
        resultados = searcher(new_dict, fecha)
    elif dec_EVREL.strip().lower() == "n":
        checker_link_ER()
        new_dict = extractorXls()
        resultados = searcher(new_dict, fecha)
        print(resultados)
    else:
        print("Favor de solo seleccionar (s / n).")
        exit()
    doc_creator(fecha, mes, titulo, subtitulo, new_dict)

