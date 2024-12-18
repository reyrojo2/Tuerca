from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from subprocess import CREATE_NO_WINDOW
import datetime
import time
import os
import glob
import fnmatch
import pandas as pd
from bs4 import BeautifulSoup
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import datetime
import sys


class DescargaEventHandler(FileSystemEventHandler):
    def __init__(self, observer):
        self.observer = observer
        self.descarga_completada = False

    def on_created(self, event):
        # Verificar si se ha creado un archivo en el directorio de descargas
        if event.is_directory:
            return
        if event.src_path.endswith('.xls'):
            # Si se ha creado un archivo con extensión '.xls', se considera la descarga completa
            self.descarga_completada = True
            # Detener el observador después de la descarga completa
            self.observer.stop()

def searchdate(usuario, clave, ticket):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless=new')
    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False
    })
    chrome_service = ChromeService (executable_path="chromedriver.exe")
    # Iniciar el controlador de ChromeDriver sin mostrar la terminal
    chrome_service.creation_flags = CREATE_NO_WINDOW
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
    # Acceder a la página de inicio
    driver.get('Insert here your URL')

    # Esperar un poco para asegurar la carga de la página
    time.sleep(2)

    # Rellenar el formulario de inicio de sesión
    username_input = driver.find_element(By.XPATH, '//input[@name="Login1$UserName"]')
    username_input.send_keys(usuario)

    password_input = driver.find_element(By.XPATH, '//input[@name="Login1$Password"]')
    password_input.send_keys(clave)

    # Enviar el formulario de inicio de sesión
    password_input.send_keys(Keys.RETURN)

    # Esperar un poco para asegurar el inicio de sesión
    time.sleep(2)

    # Esperar un poco más para asegurar que el elemento esté presente en el DOM
    time.sleep(2)

    if "Portal XXXXX" in driver.title:
        print("Inicio de sesión exitoso")
        # Encontrar y completar el campo de fecha de inicio
        ticket_input = driver.find_element(By.XPATH, '//input[@name="ctl00$ctl00$ContentPlaceHolder1$corporativo$txt_id_ticket"]')
        ticket_input.clear()
        ticket_input.send_keys(ticket)

        # Hacer clic en el botón de Consulta
        descargar_button = driver.find_element(By.ID, 'ContentPlaceHolder1_corporativo_btnConsultar')
        descargar_button.click()

        # Esperar hasta que el archivo se descargue completamente
        time.sleep(2)

        # Encontrar la tabla ContentPlaceHolder1_corporativo_GridView1
        tabla = driver.find_element(By.ID, 'ContentPlaceHolder1_corporativo_GridView1')

        # Encontrar el tercer td en la tabla
        tds = tabla.find_elements(By.TAG_NAME, 'td')
        if len(tds) >= 3:
            fecha = tds[2].text.strip()  # Obtener el texto del tercer td
            print(fecha)

            # Rellenar el campo de fecha inicial
            fecha_input = driver.find_element(By.ID, 'ContentPlaceHolder1_corporativo_Fecha_i')
            fecha_input.clear()
            fecha_input.send_keys(fecha)
            print(fecha)
            # Rellenar el campo de fecha final
            fecha_input = driver.find_element(By.ID, 'ContentPlaceHolder1_corporativo_Fecha_f')
            fecha_input.clear()
            fecha_input.send_keys(fecha)
            print(fecha)
            # Hacer clic en el botón de descarga
            descargar_button = driver.find_element(By.ID, 'ContentPlaceHolder1_corporativo_btnReporte')
            descargar_button.click()
            print("Ya dió clic al botón Descargar")
            
            # Iniciar el observador para monitorear el directorio de descargas
            directorio_descargas = os.path.expanduser('~/Downloads')
            observer = Observer()
            event_handler = DescargaEventHandler(observer)
            observer.schedule(event_handler, path=directorio_descargas, recursive=False)
            observer.start()

            # Esperar hasta que se complete la descarga o se alcance el tiempo máximo de espera (30 segundos)
            tiempo_espera = 25  # Tiempo máximo de espera en segundos
            tiempo_transcurrido = 0

            while not event_handler.descarga_completada and tiempo_transcurrido < tiempo_espera:
                time.sleep(3)  # Esperar 5 segundos
                tiempo_transcurrido += 5

            # Detener el observador si no se encontró el archivo dentro del tiempo de espera
            observer.stop()
            observer.join()

            # Cerrar el navegador
            driver.quit()

    else:
    # Manejar cualquier excepción que ocurra durante el proceso
        print("Error. Validar las credenciales o el acceso a la web.")
        messagebox.showinfo("ERROR", "Validar las credenciales o el acceso a la web.")
        # Cerrar el navegador
        driver.quit()
        # Terminar la ejecución del programa
        sys.exit()

def CrearyBuscarArchivo(hora_actual, minuto_actual, fecha_actual):
    #hora_actual = "00"
    # Obtener la hora y el minuto en el formato "hh_mm" o "h_mm"
    hora_actual = datetime.datetime.now().strftime('%H')
    minuto_actual = datetime.datetime.now().strftime('%M')

    # Obtener la fecha actual en el formato "dd_mm_aaaa"
    fecha_actual = datetime.datetime.now().strftime('%d_%m_%Y')

    # Eliminar el primer cero si la hora comienza con él
    if hora_actual.startswith('0'):
        hora_actual = hora_actual[1:]

    # Reemplazar "00" por "0" si la hora es medianoche
    if hora_actual == '00':
        hora_actual = '0'

    # Convertir las cadenas de texto a enteros
    fecha_actual = int(fecha_actual.replace("_", ""))
    hora_actual = int(hora_actual)
    minuto_actual = int(minuto_actual)

    # Obtener la ruta de descarga predeterminada
    ruta_descargas = os.path.expanduser('~/Downloads')
    # Obtener la lista de archivos en la carpeta de descargas
    archivos_descargados = glob.glob(os.path.join(ruta_descargas, '*'))

    # Buscar el archivo que coincide con el nombre esperado
    #patron = "Reporte_general_05_07_2023 14_05_32"
    patron = "Reporte_general_{:02d}_{:02d}_{:04d} {:01d}_{:02d}_**.xls".format(
        fecha_actual // 1000000, (fecha_actual // 10000) % 100, fecha_actual % 10000,
        hora_actual, minuto_actual
    )
    print(patron)
    for archivo in archivos_descargados:
        if fnmatch.fnmatch(os.path.basename(archivo), patron):
            return archivo
    
    # Buscar el archivo con un minuto menos
    #patron = "Reporte_general_05_07_2023 14_05_32"
    patron = "Reporte_general_{:02d}_{:02d}_{:04d} {:01d}_{:02d}_**.xls".format(
        fecha_actual // 1000000, (fecha_actual // 10000) % 100, fecha_actual % 10000,
        hora_actual, minuto_actual - 1
    )
    print(patron)
    for archivo in archivos_descargados:
        if fnmatch.fnmatch(os.path.basename(archivo), patron):
            return archivo
        
    # Buscar el archivo descargado
    if archivo is None:
        mensaje = f"Error al descargar el archivo, intente nuevamente."
        messagebox.showinfo("Error", mensaje)
        print("Error: No se encontró el archivo descargado.")
    else:
        nombre_archivo = os.path.basename(archivo)
        print("Nombre del archivo descargado:", nombre_archivo)
    return nombre_archivo

# Función para recibir el nombre del archivo
def select_xls_file(nombre_archivo):
    downloads_folder = os.path.expanduser('~/Downloads')  # Obtener la ruta de la carpeta de descargas
    file_path = os.path.join(downloads_folder, nombre_archivo)  # Unir la ruta de descargas con el nombre del archivo
    if not file_path:
        print("No se proporcionó ningún archivo XLS.")
        return None
    else:
        return file_path

# Función para abrir el archivo XLS en Excel y obtener los datos
def open_excel_file_and_get_data(file_path, ticket):
    # Leer el contenido HTML del archivo seleccionado
    with open(file_path, 'r') as file:
        contenido_html = file.read()

    # Reemplazar los caracteres \t con espacios y eliminar las etiquetas <div>
    contenido_html = contenido_html.replace('\t', '').replace('<div>', '').replace('</div>', '')

    # Analizar el contenido HTML con BeautifulSoup
    soup = BeautifulSoup(contenido_html, 'html.parser')

    # Encontrar todas las etiquetas 'tr' que representan las filas de la tabla
    rows = soup.find_all('tr')

    # Convertir las filas en una lista de listas que representan las celdas de la tabla
    data_list = []
    for row in rows:
        cells = row.find_all(['td', 'th'])
        row_data = [cell.text.strip() for cell in cells]
        data_list.append(row_data)

    # Obtener el encabezado de las columnas a partir de la primera fila de datos
    header = data_list[0]
    data_list = data_list[1:]  # Omitir la primera fila ya que es el encabezado

    # Convertir la lista de datos en un DataFrame
    data = pd.DataFrame(data_list, columns=header)
    # Comparar el valor en la columna "Ticket" con otro valor
    valor_a_comparar = ticket
    if valor_a_comparar in data['Ticket'].values:
        print(f"El valor {valor_a_comparar} está presente en la columna 'Ticket'.")
        # Obtener la fila que coincide con el valor a comparar
        fila_coincidente = data.loc[data['Ticket'] == valor_a_comparar].values.tolist()[0]

        # Obtener los valores de la fila sin el encabezado
        #fila_coincidente_sin_encabezado = fila_coincidente.values[0]

    else:
        print(f"El valor {valor_a_comparar} no se encuentra en la columna 'Ticket'.")
        fila_coincidente = []

    return fila_coincidente  # Retorna la lista con los datos

# Función para formatear los datos con la fecha y hora actual y para Google Sheets
def format_data_with_datetime(fila_coincidente):
    current_datetime = datetime.datetime.now()
    current_datetime_formatted = current_datetime.strftime("%d-%m-%Y %H:%M:%S")

    formatted_data = [current_datetime_formatted]
    for i, cell in enumerate(fila_coincidente):  # Comenzamos a contar desde 0 (valor predeterminado para 'start')
        if i == 11 or i == 14:  # Formateamos las celdas 11 y 14 (contando desde 0)
            if isinstance(cell, str):
                try:
                    date_obj = datetime.datetime.strptime(cell, "%d/%m/%Y %H:%M:%S")
                    formatted_data.append(date_obj.strftime('%d-%m-%Y %H:%M:%S'))
                except ValueError:
                    formatted_data.append(cell)  # Si hay un error en el formato, mantén el valor original
            else:
                formatted_data.append('')
        else:
            formatted_data.append(cell)

    return formatted_data

# Función para obtener la siguiente fila vacía en una hoja de Google Sheets y pegar los datos
def paste_data_to_google_sheets(sheet, data):
    next_empty_row = len(sheet.get_all_values()) + 1
    start_cell = f'A{next_empty_row}'
    num_rows = len(data)
    num_cols = len(data[0])
    end_cell = f'{chr(64 + num_cols)}{next_empty_row + num_rows - 1}'
    values = [data]  # Los datos están en una lista, lo convertimos en una lista de listas (1 fila)
    sheet.update(start_cell, values, value_input_option='RAW', major_dimension='ROWS')

# Función principal
def mainup(usuario, clave, ticket):
    # Configurar las credenciales
    scope = ['Here goes your spreadsheets',
             'Here goes your API']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('Insert your .json file', scope)
    client = gspread.authorize(credentials)

    # Llamar a la función searchdate
    searchdate(usuario, clave, ticket)

    hora_actual = datetime.datetime.now().strftime('%H')
    minuto_actual = datetime.datetime.now().strftime('%M')
    fecha_actual = datetime.datetime.now().strftime('%d_%m_%Y')
    # Llamar a la función CrearyBuscarArchivo
    nombre_archivo = CrearyBuscarArchivo(hora_actual, minuto_actual, fecha_actual)

    # Obtener la ruta del archivo XLS seleccionado
    xls_file_path = select_xls_file(nombre_archivo)
    if xls_file_path is None:
        return  # Salir si no se seleccionó ningún archivo XLS

    # Obtener los datos del archivo XLS
    fila_coincidente = open_excel_file_and_get_data(xls_file_path, ticket)
    if fila_coincidente is None:
        mensaje = f"Error al cargar el ticket, inténtelo nuevamente."
        messagebox.showinfo("Error", mensaje)
        return  # Salir si no se pudo abrir el archivo XLS

    # Abrir el libro de Google Sheets
    spreadsheet_id = 'Here goes your spreadsheet ID'
    sheet_name = 'Here goes your spreadsheets name sheet'
    sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)

    # Formatear los datos con la fecha y hora actual y para Google Sheets
    formatted_data = format_data_with_datetime(fila_coincidente)
    print(formatted_data)
    # Pegar los datos en la hoja de Google Sheets
    paste_data_to_google_sheets(sheet, formatted_data)

    print("Los datos se han copiado correctamente en Google Sheets.")
    # Mostrar el pop-up con la ruta y el nombre del archivo descargado
    mensaje = f"Los datos se han copiado correctamente en Google Sheets."
    messagebox.showinfo("Carga completada", mensaje)


