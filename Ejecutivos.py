from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
import time
from subprocess import CREATE_NO_WINDOW

def search1(identificacion):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('headless')
    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False })
    chrome_service = ChromeService('chromedriver')
    # Iniciar el controlador de ChromeDriver sin mostrar la terminal
    chrome_service.creation_flags = CREATE_NO_WINDOW
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options, service_args=['--silent'])

    driver.get("Here goes your URL")

    # Esperar un poco para asegurar la carga de la página
    time.sleep(2)

    Identificacion_input = driver.find_element(By.ID, "Identificacion")
    Identificacion_input.clear()
    Identificacion_input.send_keys(identificacion)

    driver.find_element(By.ID, "Button1").click()

    # Verificar si se encuentra el elemento GridView1
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "GridView1")))
    except:
        # No se encontró el elemento, mostrar mensaje y salir de la función
        driver.quit()
        return []

    # Iterar sobre las filas y obtener los valores de los campos deseados
    rows = driver.find_elements(By.XPATH, "//table[@id='GridView1']/tbody/tr")

    # Crear una lista para almacenar los datos de los resultados
    results = []

    # Variable para indicar si se encontraron resultados
    encontrado = False

    # Iterar sobre las filas y obtener los valores de los campos deseados
    for row in rows[1:]:
        Identificación = row.find_element(By.XPATH, "./td[1]").text
        Cliente = row.find_element(By.XPATH, "./td[2]").text
        Asesor = row.find_element(By.XPATH, "./td[3]").text
        Jefe = row.find_element(By.XPATH, "./td[4]").text
        Region = row.find_element(By.XPATH, "./td[5]").text
        Segmento = row.find_element(By.XPATH, "./td[6]").text
        Categoria = row.find_element(By.XPATH, "./td[7]").text
        Provincia = row.find_element(By.XPATH, "./td[8]").text
        Ejecutivo = row.find_element(By.XPATH, "./td[9]").text
        Coordinador = row.find_element(By.XPATH, "./td[10]").text
        Extensión = row.find_element(By.XPATH, "./td[11]").text
        Celular = row.find_element(By.XPATH, "./td[12]").text
        Gestor = row.find_element(By.XPATH, "./td[13]").text

        # Agregar los valores a la lista de resultados
        results.append([Identificación, Cliente, Asesor, Jefe, Region, Segmento, Categoria, Provincia, Ejecutivo, Coordinador, Extensión, Celular, Gestor])

        # Actualizar la variable 'encontrado'
        encontrado = True

    if not encontrado:
        results = [["Cliente no se encuentra en base"]]

    driver.quit()  # Cerrar la ventana de Chrome
    return results
