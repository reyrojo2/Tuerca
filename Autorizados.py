from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from subprocess import CREATE_NO_WINDOW
import time



def searchaut(identificacion, celular, correo):

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('headless')
    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False })
    chrome_service = ChromeService ('chromedriver')

    # Iniciar el controlador de ChromeDriver sin mostrar la terminal
    chrome_service.creation_flags = CREATE_NO_WINDOW
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options, service_args=['--silent'])

    driver.get("Insert here your URL")

    # Esperar un poco para asegurar la carga de la página
    time.sleep(2)

    Identificacion_input = driver.find_element(By.ID, "Identificacion")
    Identificacion_input.clear()
    if not identificacion:
        identificacion = "0"
    Identificacion_input.send_keys(identificacion)

    Celular_input = driver.find_element(By.ID, "cel_per_aut")
    Celular_input.clear()
    if not celular:
        celular = "0"
    Celular_input.send_keys(celular)

    Correo_input = driver.find_element(By.ID, "correo_per_aut")
    Correo_input.clear()
    if not correo:
        correo = "0"
    Correo_input.send_keys(correo)

    driver.find_element(By.ID, "Button1").click()

        # Verificar si se encuentran los campos deseados
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "GridView1")))
        Identificacion_input = driver.find_element(By.ID, "Identificacion")
        Celular_input = driver.find_element(By.ID, "cel_per_aut")
        Correo_input = driver.find_element(By.ID, "correo_per_aut")
    except:
        # No se encontraron los campos, mostrar mensaje y salir de la función
        driver.quit()
        print("Campos no encontrados: Identificación, cel_per_aut, correo_per_aut")
        return []

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
        Región = row.find_element(By.XPATH, "./td[2]").text
        Categoria = row.find_element(By.XPATH, "./td[3]").text
        Cliente = row.find_element(By.XPATH, "./td[4]").text
        persona_aut = row.find_element(By.XPATH, "./td[5]").text
        cedula_persona_aut = row.find_element(By.XPATH, "./td[6]").text
        correo_aut = row.find_element(By.XPATH, "./td[7]").text
        telefonos = row.find_element(By.XPATH, "./td[8]").text
        Rep_Legal = row.find_element(By.XPATH, "./td[9]").text
        Ced_Rep_Legal = row.find_element(By.XPATH, "./td[10]").text
        Correo_Rep_Legal = row.find_element(By.XPATH, "./td[11]").text
        Celular_Rep_Legal = row.find_element(By.XPATH, "./td[12]").text
        Fecha_Cad_Nom = row.find_element(By.XPATH, "./td[13]").text
        Estado_Nom = row.find_element(By.XPATH, "./td[14]").text
        Fecha_Actual = row.find_element(By.XPATH, "./td[15]").text

        # Agregar los valores a la lista de resultados
        results.append([Identificación, Región, Categoria, Cliente, persona_aut, cedula_persona_aut, correo_aut, telefonos, Rep_Legal, Ced_Rep_Legal, Correo_Rep_Legal, Celular_Rep_Legal, Fecha_Cad_Nom, Estado_Nom, Fecha_Actual])

        # Actualizar la variable 'encontrado'
        encontrado = True

        

    if not encontrado:
        results = [["Cliente no se encuentra en base"]]
    print(results)
    driver.quit()  # Cerrar la ventana de Chrome
    return results