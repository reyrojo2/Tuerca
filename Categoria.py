import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configura las credenciales y abre el archivo de Google Sheets
scope = ['Here goes your spreadsheets',
         'Here goes your API']
credentials = ServiceAccountCredentials.from_json_keyfile_name('Insert your .json file', scope)  # Reemplaza con el nombre de tu archivo de credenciales
client = gspread.authorize(credentials)
spreadsheet = client.open('Base de Categorias')
worksheet = spreadsheet.sheet1  # Reemplaza con el nombre de tu hoja
cell = "E2"
# Función para buscar la identificación y devolver el valor correspondiente
def buscar_categoria(identificacion):
    data = worksheet.get('A:C')  # Obtener el rango de celdas A:C
    values = data
    for row in values:
        if row[0] == identificacion:
            return row[2]  # Devolver el valor de la tercera columna
    return "Sin categoría"

# Función para obtener la la fecha de actualización de la base
def buscar_actualizacion():
    # Obtiene el contenido de la celda especificada.
    cell_value = worksheet.acell("E2").value
    print(cell_value)
    return cell_value