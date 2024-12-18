from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import gspread
from subprocess import CREATE_NO_WINDOW
import time
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import tkinter.messagebox as messagebox

def searchdate(clave, valor1, valor2):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('headless')
    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False })
    chrome_service = ChromeService(executable_path="chromedriver.exe")
    chrome_service.creation_flags = CREATE_NO_WINDOW
    # Iniciar el controlador de ChromeDriver sin mostrar la terminal
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
    # Acceder a la página de inicio
    driver.get('Here goes your URL')

    # Esperar un poco para asegurar la carga de la página
    time.sleep(2)

    # Rellenar el formulario de inicio de sesión
    username_input = driver.find_element(By.XPATH, '//input[@name="Login1$UserName"]')
    username_input.send_keys(valor2)

    password_input = driver.find_element(By.XPATH, '//input[@name="Login1$Password"]')
    password_input.send_keys(clave)

    # Enviar el formulario de inicio de sesión
    password_input.send_keys(Keys.RETURN)

    # Esperar un poco para asegurar el inicio de sesión
    time.sleep(2)

    # Esperar un poco más para asegurar que el elemento esté presente en el DOM
    time.sleep(2)

    #Hacer la validación de si los datos de usuario y clave son correctos, de ser así procede con la búsqueda
    if "Portal XXXXX" in driver.title:
        print("Inicio de sesión exitoso")
        # Cerrar el navegador
        driver.quit()
        # Configura las credenciales y abre el archivo de Google Sheets
        scope = ['Here goes your spreadsheets',
                'Here goes your API']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('Insert your .json file here', scope)  # Reemplaza con el nombre de tu archivo de credenciales
        client = gspread.authorize(credentials)
        spreadsheet = client.open('BaseTickets')
        worksheet = spreadsheet.sheet1  # Reemplaza con el nombre de tu hoja
        data = worksheet.get_all_values()  # Obtener todos los valores de la hoja de cálculo
        worksheet_values = data[1:]  # Excluir la primera fila que contiene encabezados

        resultados = []

        # Convertir valor1 a objeto datetime
        fecha_busqueda = datetime.strptime(valor1, '%d-%m-%Y')

        # Buscar coincidencias en las columnas A y Q
        for row in worksheet_values:
            if row[0]:
                # Verificar que la fila no esté completamente vacía
                if any(row[1:]):  
                    fecha = datetime.strptime(row[15][:10], '%d-%m-%Y')  # Columna 16 está en el índice 15 (0-indexed)
                    if fecha.date() == fecha_busqueda.date() and row[16].strip().lower() == valor2.strip().lower():
                        # Obtener los valores específicos de las columnas B, D, I, O, P, Q, R y S
                        valores_filtrados = [row[1], row[3], row[8], row[14], row[15], row[16], row[17], row[18]]
                        resultados.append(valores_filtrados)
        print(resultados)
        return resultados
        
    else:
        messagebox.showinfo("ERROR", "Compruebe las credenciales o el acceso a la web.")
        print("Error")
        # Cerrar el navegador
        driver.quit()
