import Autorizados
import threading
import Ejecutivos
import Recomendados
import Categoria
import DescargarTicket
import ProduAs
import keyboard
import tkinter as tk
from tkcalendar import DateEntry
from tkinter import ttk
from tkinter import Scrollbar
from tkinter import Menu
from tkinter import font
import tkinter.messagebox as messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

def mostrar_pestana(pestana):
    # Obtener el índice de la pestaña seleccionada
    indice = notebook.index(pestana)

    # Mostrar la pestaña seleccionada
    notebook.select(indice)

    # Ocultar el resto de las pestañas
    for i in range(notebook.index("end")):
        if i != indice:
            notebook.tab(i, state="hidden")

#AQUÍ ESTOY IMPORTANDO EL CÓDIGO PARA LA BÚSQUEDA DE AUTORIZADOS
#PARA QUE NO SE CUELGE USO LA BÚSQUEDA EN HILOS
#TAMBIÉN SE CONFIGURA LA INTERFAZ (PESTAÑA 1)
class BusquedaAutorizados(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.root = pestana1
        self.entry_identificacion = ttk.Entry(pestana1)
        self.entry_celular = ttk.Entry(pestana1)
        self.entry_correo = ttk.Entry(pestana1)
        self.treeview_aut = None
        self.create_treeview()

    def create_treeview(self):
        # Crear el contenedor para las barras de desplazamiento
        scrollbar_frame = ttk.Frame(pestana1)
        scrollbar_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Crear la barra de desplazamiento vertical
        y_scrollbar = Scrollbar(scrollbar_frame, orient="vertical")
        y_scrollbar.pack(side="right", fill="y")

        # Crear la barra de desplazamiento horizontal
        x_scrollbar = Scrollbar(scrollbar_frame, orient="horizontal")
        x_scrollbar.pack(side="bottom", fill="x")

        # Crear el treeview con las barras de desplazamiento
        self.treeview_aut = ttk.Treeview(scrollbar_frame, columns=("col1", "col2", "col3", "col4", "col5", "col6", "col7", "col8", "col9", "col10", "col11", "col12", "col13", "col14", "col15"), show="headings", yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set, height=4)
        self.treeview_aut.pack(fill="both", expand=True)

        # Configurar las barras de desplazamiento para sincronizarse con el treeview
        y_scrollbar.config(command=self.treeview_aut.yview)
        x_scrollbar.config(command=self.treeview_aut.xview)

        # Configurar los encabezados de las columnas
        self.treeview_aut.heading("col1", text="Identificación")
        self.treeview_aut.heading("col2", text="Región")
        self.treeview_aut.heading("col3", text="Categoria")
        self.treeview_aut.heading("col4", text="Cliente")
        self.treeview_aut.heading("col5", text="Persona Aut")
        self.treeview_aut.heading("col6", text="Cédula Persona Aut")
        self.treeview_aut.heading("col7", text="Correo Aut")
        self.treeview_aut.heading("col8", text="Teléfonos")
        self.treeview_aut.heading("col9", text="Rep Legal")
        self.treeview_aut.heading("col10", text="Ced Rep Legal")
        self.treeview_aut.heading("col11", text="Correo Rep Legal")
        self.treeview_aut.heading("col12", text="Celular Rep Legal")
        self.treeview_aut.heading("col13", text="Fecha Cad Nom")
        self.treeview_aut.heading("col14", text="Estado Nom")
        self.treeview_aut.heading("col15", text="Fecha Actual")

        # Establecer el ancho de cada columna
        self.treeview_aut.column("col1", width=100, minwidth=100, stretch=True)
        self.treeview_aut.column("col2", width=50, minwidth=50, stretch=True)
        self.treeview_aut.column("col3", width=150, minwidth=150, stretch=True)
        self.treeview_aut.column("col4", width=300, minwidth=280, stretch=True)
        self.treeview_aut.column("col5", width=300, minwidth=280, stretch=True)
        self.treeview_aut.column("col6", width=120, minwidth=120, stretch=True)
        self.treeview_aut.column("col7", width=200, minwidth=200, stretch=True)
        self.treeview_aut.column("col8", width=100, minwidth=100, stretch=True)
        self.treeview_aut.column("col9", width=300, minwidth=280, stretch=True)
        self.treeview_aut.column("col10", width=100, minwidth=80, stretch=True)
        self.treeview_aut.column("col11", width=200, minwidth=200, stretch=True)
        self.treeview_aut.column("col12", width=100, minwidth=100, stretch=True)
        self.treeview_aut.column("col13", width=100, minwidth=100, stretch=True)
        self.treeview_aut.column("col14", width=120, minwidth=120, stretch=True)
        self.treeview_aut.column("col15", width=80, minwidth=80, stretch=True)

        # Desactivar el redimensionamiento manual de las columnas al hacer doble clic
        self.treeview_aut.bind("<Double-Button-1>", lambda event: "break")

        # Asociar el menú contextual con el treeview
        self.treeview_aut.bind("<Double-1>", self.on_double_click)

        # Crear un contenedor para el label y el botón
        container = ttk.Frame(pestana1)
        container.pack(anchor='w', fill='x')

        # Crear el label para mostrar la cantidad de filas encontradas
        self.label_resultado = ttk.Label(container, text="Cantidad de registros encontrados: 0", anchor="w", justify="left")
        self.label_resultado.pack(side='left', padx=5, pady=5)

        # Crear el label para mostrar información referente a la categoría de la empresa
        self.label_categoria = ttk.Label(container, text="Categoría: ")
        self.label_categoria.pack(fill='both', padx=(200, 5), pady=5, anchor='center')

        self.label_actualizacion_categoria = ttk.Label(container, text="Actualizado: ")
        self.label_actualizacion_categoria.pack(fill='both', padx=(200, 5), pady=0, anchor='center')

        # Crear los campos de entrada y el botón de búsqueda
        label_identificacion = ttk.Label(pestana1, text="Identificación")
        label_identificacion.pack()
        self.entry_identificacion = ttk.Entry(pestana1)
        self.entry_identificacion.pack(padx=20, pady=0)

        label_celular = ttk.Label(pestana1, text="Celular")
        label_celular.pack()
        self.entry_celular = ttk.Entry(pestana1)
        self.entry_celular.pack(padx=20, pady=0)

        label_correo = ttk.Label(pestana1, text="Correo")
        label_correo.pack()
        self.entry_correo = ttk.Entry(pestana1)
        self.entry_correo.pack(padx=20, pady=0)

        self.button = ttk.Button(pestana1, text="Buscar", command=self.ejecutar_busqueda)
        self.button.pack(padx=20, pady=10)

    def update_table(self, data):
        # Eliminar todas las filas existentes
        self.treeview_aut.delete(*self.treeview_aut.get_children())

        # Insertar los nuevos datos en la tabla
        for item in data:
            self.treeview_aut.insert("", "end", values=item)

        self.treeview_aut.configure(height=1)

        # Actualizar la cantidad de filas encontradas en el label
        self.actualizar_cantidad_filas(self.treeview_aut)

    def actualizar_cantidad_filas(self):
        cantidad_filas = len(self.treeview_aut.get_children())
        texto = f"Cantidad de registros encontrados: {cantidad_filas}"
        self.label_resultado.config(text=texto)

    def on_double_click(self, event):
        # Obtener la celda seleccionada en el treeview
        row_id = self.treeview_aut.identify_row(event.y)
        column_id = self.treeview_aut.identify_column(event.x)

        if row_id and column_id:
            # Obtener el valor de la celda seleccionada
            cell_value = self.treeview_aut.set(row_id, column_id)

            # Verificar si la columna seleccionada es la columna "Teléfonos"
            if column_id == "#8" and (cell_value == "99999999" or cell_value == "99999999"):
                messagebox.showwarning("¡NO LLAMAR!", "Este es un número de XXXXXX, no de un cliente.")
                return  # Salir de la función si se detecta el número

            # Copiar el contenido de la celda al portapapeles
            ventana.clipboard_clear()
            ventana.clipboard_append(cell_value)
            messagebox.showinfo("Mensaje informativo", "El contenido de la celda se copió al portapapeles.")

    def buscar_cat(self):
        identificacion = self.entry_identificacion.get()
        resultados = Categoria.buscar_categoria(identificacion)
        
        if resultados is None:
            resultados = "Sin categoría"
        
        return resultados
    
    def mostrar_fecha_act(self):
        resultados2 = Categoria.buscar_actualizacion()
        return(resultados2)
    
    def buscar_autorizados(self):
        identificacion = self.entry_identificacion.get()
        celular = self.entry_celular.get()
        correo = self.entry_correo.get()
        resultados = Autorizados.searchaut(identificacion, celular, correo)
        return(resultados)
    
    def enable_button(self):
        self.button.config(state=tk.NORMAL)  # Habilitar el botón

    def ejecutar_busqueda(self):
        identificacion = self.entry_identificacion.get()
        celular = self.entry_celular.get()
        correo = self.entry_correo.get()
        if not identificacion and not celular and not correo:
            # Al menos uno de los campos está vacío, no ejecutar la búsqueda
            messagebox.showerror("ERROR","¡Llena al menos un campo para la búsqueda!")
            return
        # Realizar la búsqueda de la categoría
        resultados = self.buscar_cat()
        resultados2 = self.mostrar_fecha_act()
        # Construir el texto completo
        texto_etiqueta = "Categoría: " + resultados
        texto_etiqueta2 = "Actualización: " + resultados2
        # Actualizar el label con los resultados
        self.label_categoria.config(text=texto_etiqueta)
        self.label_actualizacion_categoria.config(text=texto_etiqueta2)
        # Deshabilitar el botón "Buscar"
        self.button.config(state=tk.DISABLED)
        # Crear hilos para realizar las búsquedas
        autorizados_thread = threading.Thread(target=lambda: self.mostrar_resultados(self.buscar_autorizados))

        # Iniciar los hilos
        autorizados_thread.start()

    def mostrar_resultados(self, funcion_busqueda1):
        resultados = funcion_busqueda1()

        # Actualizar la interfaz gráfica en el hilo principal
        self.root.after(0, self.actualizar_interfaz1, resultados)
        # Habilitar el botón
        self.after(0, self.enable_button)
        messagebox.showwarning("RECUERDA","Este programa ejecuta la búsqueda de autorizados en XXXXXX. DEBES BUSCAR EN XXXXX POR TU CUENTA.")

    def actualizar_interfaz1(self, resultados):
        # Limpiar la tabla
        self.treeview_aut.delete(*self.treeview_aut.get_children())

        # Verificar si no hay resultados
        if not resultados:
            self.treeview_aut.insert("", "end", values=("", "", "", "Cliente no se encuentra en la base.", "", "", "", "", "", "", "", "", "", "", ""))
            cantidad_filas = 0
        else:
            # Agregar los resultados a la tabla
            for resultado in resultados:
                self.treeview_aut.insert("", "end", values=resultado)
        self.actualizar_cantidad_filas()
    
#CÓDIGO PARA CREAR LA BÚSQUEDA DE EJECUTIVOS Y RECOMENDADOS, JUNTO CON SU INTERFAZ (PESTAÑA 2)
class BusquedaEjeReco(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.root = pestana2
        self.entry_identificacion = ttk.Entry(pestana2)
        self.tabla_ejecutivos = None
        self.tabla_recomendados = None
        self.create_interface()

    def create_interface(self):
        # Crear marco para la tabla de ejecutivos
        marco_ejecutivos = ttk.Frame(self.root)
        marco_ejecutivos.pack(fill="both", expand=True, padx=10, pady=10)
        # Etiqueta para el nombre de la tabla de ejecutivos
        label_ejecutivos = ttk.Label(marco_ejecutivos, text="Búsqueda de ejecutivos:")
        label_ejecutivos.pack(side="top", anchor="w")

        # Agregar barras de desplazamiento
        self.tabla_ejecutivos = ttk.Treeview(marco_ejecutivos, columns=("identificacion", "cliente", "asesor", "jefe", "region", "segmento", "categoria", "provincia", "ejecutivo", "coordinador", "extension", "celular", "gestor"), height=3)

        scrollbar_ejecutivos_y = ttk.Scrollbar(marco_ejecutivos, orient="vertical", command=self.tabla_ejecutivos.yview)
        scrollbar_ejecutivos_y.pack(side="right", fill="y")
        self.tabla_ejecutivos.configure(yscrollcommand=scrollbar_ejecutivos_y.set)

        scrollbar_ejecutivos_x = ttk.Scrollbar(self.root, orient="horizontal", command=self.tabla_ejecutivos.xview)
        scrollbar_ejecutivos_x.pack(fill="x")
        self.tabla_ejecutivos.configure(xscrollcommand=scrollbar_ejecutivos_x.set)

        # Configurar el ancho de la primera columna a cero
        self.tabla_ejecutivos.column("#0", width=0, stretch=tk.NO)

        # Cabeceras de las columnas para la tabla de ejecutivos
        self.tabla_ejecutivos.heading("identificacion", text="Identificación")
        self.tabla_ejecutivos.heading("cliente", text="Cliente")
        self.tabla_ejecutivos.heading("asesor", text="Asesor")
        self.tabla_ejecutivos.heading("jefe", text="Jefe")
        self.tabla_ejecutivos.heading("region", text="Región")
        self.tabla_ejecutivos.heading("segmento", text="Segmento")
        self.tabla_ejecutivos.heading("categoria", text="Categoría")
        self.tabla_ejecutivos.heading("provincia", text="Provincia")
        self.tabla_ejecutivos.heading("ejecutivo", text="Ejecutivo")
        self.tabla_ejecutivos.heading("coordinador", text="Coordinador")
        self.tabla_ejecutivos.heading("extension", text="Extensión")
        self.tabla_ejecutivos.heading("celular", text="Celular")
        self.tabla_ejecutivos.heading("gestor", text="Gestor")

        self.tabla_ejecutivos.pack(side="left", fill="both", expand=True)
        #Configurar ancho de las columnas
        self.tabla_ejecutivos.column("identificacion", width=100, minwidth=100, stretch=True)
        self.tabla_ejecutivos.column("cliente", width=300, minwidth=280, stretch=True)
        self.tabla_ejecutivos.column("asesor", width=150, minwidth=150, stretch=True)
        self.tabla_ejecutivos.column("jefe", width=150, minwidth=50, stretch=True)
        self.tabla_ejecutivos.column("region", width=50, minwidth=50, stretch=True)
        self.tabla_ejecutivos.column("segmento", width=110, minwidth=110, stretch=True)
        self.tabla_ejecutivos.column("categoria", width=150, minwidth=150, stretch=True)
        self.tabla_ejecutivos.column("provincia", width=150, minwidth=150, stretch=True)
        self.tabla_ejecutivos.column("ejecutivo", width=150, minwidth=150, stretch=True)
        self.tabla_ejecutivos.column("coordinador", width=150, minwidth=50, stretch=True)
        self.tabla_ejecutivos.column("extension", width=70, minwidth=70, stretch=True)
        self.tabla_ejecutivos.column("celular", width=100, minwidth=100, stretch=True)
        self.tabla_ejecutivos.column("gestor", width=150, minwidth=150, stretch=True)

        # Crear marco para la tabla de recomendados y agregar barras de desplazamiento
        marco_recomendados = ttk.Frame(self.root)
        marco_recomendados.pack(fill="both", expand=True, padx=10, pady=10)

        # Etiqueta para el nombre de la tabla de recomendados
        label_recomendados = ttk.Label(marco_recomendados, text="Búsqueda de clientes recomendados:")
        label_recomendados.pack(side="top", anchor="w")

        # Agregar barras de desplazamiento
        self.tabla_recomendados = ttk.Treeview(marco_recomendados, columns=("cuenta", "nombre", "identificacion", "calificacion", "gestor", "correo", "ases"),height=3)

        scrollbar_recomendados_y = ttk.Scrollbar(marco_recomendados, orient="vertical", command=self.tabla_recomendados.yview)
        scrollbar_recomendados_y.pack(side="right", fill="y")
        self.tabla_recomendados.configure(yscrollcommand=scrollbar_recomendados_y.set)

        scrollbar_recomendados_x = ttk.Scrollbar(self.root, orient="horizontal", command=self.tabla_recomendados.xview)
        scrollbar_recomendados_x.pack(fill="x")
        self.tabla_recomendados.configure(xscrollcommand=scrollbar_recomendados_x.set)

        # Configurar el ancho de la primera columna a cero
        self.tabla_recomendados.column("#0", width=0, stretch=tk.NO)

        # Cabeceras de las columnas para la tabla de recomendados
        self.tabla_recomendados.heading("cuenta", text="Cuenta")
        self.tabla_recomendados.heading("nombre", text="Nombre")
        self.tabla_recomendados.heading("identificacion", text="Identificación")
        self.tabla_recomendados.heading("calificacion", text="Calificación")
        self.tabla_recomendados.heading("gestor", text="Gestor")
        self.tabla_recomendados.heading("correo", text="Correo")
        self.tabla_recomendados.heading("ases", text="Ases")

        self.tabla_recomendados.pack(side="left", fill="both", expand=True)
        #Configurar ancho de las columnas
        self.tabla_recomendados.column("cuenta", width=100, minwidth=100, stretch=True)
        self.tabla_recomendados.column("nombre", width=300, minwidth=280, stretch=True)
        self.tabla_recomendados.column("identificacion", width=100, minwidth=100, stretch=True)
        self.tabla_recomendados.column("calificacion", width=100, minwidth=100, stretch=True)
        self.tabla_recomendados.column("gestor", width=150, minwidth=150, stretch=True)
        self.tabla_recomendados.column("correo", width=200, minwidth=200, stretch=True)
        self.tabla_recomendados.column("ases", width=150, minwidth=150, stretch=True)

        # Asociar el menú contextual con el treeview
        self.tabla_ejecutivos.bind("<Double-1>", self.on_double_click1)

        # Asociar el menú contextual con el treeview
        self.tabla_recomendados.bind("<Double-1>", self.on_double_click2)

        label_frame = ttk.Frame(self.root)
        label_frame.pack(pady=10)

        self.label_identificacion = ttk.Label(label_frame, text="Identificación:")
        self.label_identificacion.pack(side="left", padx=20, pady=0)

        self.entry_identificacion = ttk.Entry(label_frame)
        self.entry_identificacion.pack(side="left", padx=20, pady=0)

        self.boton_ejecutar = ttk.Button(pestana2, text="Buscar", command=self.ejecutar_busqueda)
        self.boton_ejecutar.pack(padx=20, pady=0)

    def buscar_ejecutivos(self):
        identificacion = self.entry_identificacion.get()
        resultados = Ejecutivos.search1(identificacion)
        return resultados

    def buscar_recomendados(self):
        identificacion = self.entry_identificacion.get()
        resultados = Recomendados.search2(identificacion)
        return resultados
    
    def enable_boton(self):
        self.boton_ejecutar.config(state=tk.NORMAL)  # Habilitar el botón

    def ejecutar_busqueda(self):
        identificacion = self.entry_identificacion.get()
        if not identificacion:
            # El campo vacío, no ejecutar la búsqueda
            messagebox.showerror("ERROR","¡Llena el número de RUC para la búsqueda!")
            return
        # Deshabilitar el botón "Buscar" 
        self.boton_ejecutar.config(state=tk.DISABLED)
        # Crear hilos para realizar las búsquedas
        ejecutivos_thread = threading.Thread(target=lambda: self.mostrar_resultados(self.buscar_ejecutivos, self.tabla_ejecutivos))
        recomendados_thread = threading.Thread(target=lambda: self.mostrar_resultados(self.buscar_recomendados, self.tabla_recomendados))

        # Iniciar los hilos
        ejecutivos_thread.start()
        recomendados_thread.start()
        
    def mostrar_resultados(self, funcion_busqueda, tabla_resultados):
        resultados = funcion_busqueda()

        # Actualizar la interfaz gráfica en el hilo principal
        self.root.after(0, self.actualizar_interfaz, tabla_resultados, resultados)
        # Habilitar el botón de búsqueda
        self.after(0, self.enable_boton)

    def actualizar_interfaz(self, tabla_resultados, resultados):
        # Limpiar la tabla
        tabla_resultados.delete(*tabla_resultados.get_children())

        # Verificar si no hay resultados
        if not resultados:
            tabla_resultados.insert("", "end", values=("","Cliente no se encuentra en la base.","","","","",""))
        else:
            # Agregar los resultados a la tabla
            for resultado in resultados:
                tabla_resultados.insert("", "end", values=resultado)
    
    def on_double_click1(self, event):
        # Obtener la celda seleccionada en el treeview
        row_id = self.tabla_ejecutivos.identify_row(event.y)
        column_id = self.tabla_ejecutivos.identify_column(event.x)

        if row_id and column_id:
            # Obtener el valor de la celda seleccionada
            cell_value = self.tabla_ejecutivos.set(row_id, column_id)

            # Copiar el contenido de la celda al portapapeles
            ventana.clipboard_clear()
            ventana.clipboard_append(cell_value)
            messagebox.showinfo("Mensaje informativo", "El contenido de la celda se copió al portapapeles.")
    
    def on_double_click2(self, event):

        # Obtener la celda seleccionada en el treeview
        row_id = self.tabla_recomendados.identify_row(event.y)
        column_id = self.tabla_recomendados.identify_column(event.x)

        if row_id and column_id:
            # Obtener el valor de la celda seleccionada
            cell_value = self.tabla_recomendados.set(row_id, column_id)

            # Copiar el contenido de la celda al portapapeles
            ventana.clipboard_clear()
            ventana.clipboard_append(cell_value)
            messagebox.showinfo("Mensaje informativo", "El contenido de la celda se copió al portapapeles.")
    
#ESTE ES EL CÓDIGO PARA EL TIPIFICADOR DE ACDMAIL, JUNTO CON SU INTERFAZ (PESTAÑA 3)
class TemplateCreatorACD(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Crear cuadro de entrada con barra de desplazamiento
        frame = ttk.Frame(self)
        frame.grid(padx=10, pady=10)
        
        rq_label = ttk.Label(frame, text="RQ:")
        rq_label.grid(row=0, column=0, sticky="w")
        
        # Crear cuadro de texto para RQ con barra de desplazamiento
        self.rq_text = tk.Text(frame, height=4, width=30)
        self.rq_text.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        # Agregar barra de desplazamiento vertical
        rq_scrollbar = ttk.Scrollbar(frame, command=self.rq_text.yview)
        rq_scrollbar.grid(row=0, column=2, sticky="ns")
        self.rq_text.config(yscrollcommand=rq_scrollbar.set)

        # Crear etiquetas para colocar los ID Inicio y Final
        idi_label = ttk.Label(frame, text="ID Inicio:")
        idi_label.grid(row=1, column=0, sticky="w")
        # Configura una función de validación personalizada
        validate_numeric_idi = self.register(self.validate_input)
        self.idi_entry = ttk.Entry(frame, width=40, validate="key", validatecommand=(validate_numeric_idi, "%P"))
        self.idi_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        
        idf_label = ttk.Label(frame, text="ID Final:")
        idf_label.grid(row=2, column=0, sticky="w")
        validate_numeric_idf = self.register(self.validate_input2)
        self.idf_entry = ttk.Entry(frame, width=40, validate="key", validatecommand=(validate_numeric_idf, "%P"))
        self.idf_entry.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        
        # Crear botón para crear la plantilla
        create_button = ttk.Button(self, text="Crear plantilla", command=self.create_template)
        create_button.grid(pady=1)
        
        # Crear cuadro de texto para la plantilla
        self.template_text = tk.Text(self, height=6, width=40)
        self.template_text.grid(padx=10, pady=10)
        self.template_text.config(state=tk.DISABLED)  # Deshabilitar la edición

        # Agregar barra de desplazamiento vertical al cuadro de texto
        template_scrollbar = ttk.Scrollbar(self, command=self.template_text.yview)
        template_scrollbar.grid(row=2, column=1, sticky="ns")
        self.template_text.config(yscrollcommand=template_scrollbar.set)

        # Crear botón para copiar la plantilla
        copy_button = ttk.Button(self, text="Copiar", command=self.copiar)
        copy_button.grid(row=3, column=0, padx=5, pady=10, sticky="w")
        
        # Crear botón para borrar la plantilla
        delete_button = ttk.Button(self, text="Borrar", command=self.borrar)
        delete_button.grid(row=3, column=0, padx=5, pady=10, sticky="e")
    
    def validate_input(self, P):
        self.idi_entry.delete(0, tk.END)
        self.idi_entry.insert(0, P.strip())
        return True
    
    def validate_input2(self, P):
        self.idf_entry.delete(0, tk.END)
        self.idf_entry.insert(0, P.strip())
        return True
    
    def create_template(self):
        self.template_text.config(state=tk.NORMAL)
        RQ = self.rq_text.get("1.0", tk.END).strip()
        ID_Inicio = self.idi_entry.get()
        ID_Final = self.idf_entry.get()
        if not RQ or not ID_Inicio or not ID_Final:
            messagebox.showerror("ERROR","¡Llena los campos!")
            self.template_text.config(state=tk.DISABLED)
            return
        template = f"RQ: {RQ}\nID Inicio: {ID_Inicio}\nID Final: {ID_Final}"
        self.template_text.delete("1.0", tk.END)  # Borrar el contenido actual
        self.template_text.insert(tk.END, template)
        self.template_text.config(state=tk.DISABLED)  # Deshabilitar la edición después de insertar
    
    def copiar(self):
        template = self.template_text.get("1.0", tk.END)
        self.clipboard_clear()
        self.clipboard_append(template)
        messagebox.showinfo("Copiado", "Plantilla copiada al portapapeles.")
    
    def borrar(self):
        self.template_text.config(state=tk.NORMAL)
        self.template_text.delete("1.0", tk.END)
        self.template_text.config(state=tk.DISABLED)  # Deshabilitar la edición después de borrar
        self.idi_entry.delete(0,"end")
        self.idf_entry.delete(0,"end")
        self.rq_text.delete("1.0", "end")

#ESTE ES EL CÓDIGO PARA EL TIPIFICADOR DE SEGUIMIENTO, JUNTO CON SU INTERFAZ (PESTAÑA 4)
class TemplateCreatorSEG(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Crear cuadro de entrada con barra de desplazamiento
        frame = ttk.Frame(self)
        frame.grid(padx=10, pady=10)
        
        rq_label = ttk.Label(frame, text="RQ:")
        rq_label.grid(row=0, column=0, sticky="w")
        
        # Crear cuadro de texto para RQ con barra de desplazamiento
        self.rq_text = tk.Text(frame, height=5, width=30)
        self.rq_text.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        # Agregar barra de desplazamiento vertical
        rq_scrollbar = ttk.Scrollbar(frame, command=self.rq_text.yview)
        rq_scrollbar.grid(row=0, column=2, sticky="ns")
        self.rq_text.config(yscrollcommand=rq_scrollbar.set)
        
        # Crear etiqueta para colocar el ID Final
        idf_label = ttk.Label(frame, text="ID Final:")
        idf_label.grid(row=2, column=0, sticky="w")
        validate_numeric = self.register(self.validate_input)
        self.idf_entry = ttk.Entry(frame, width=40, validate="key", validatecommand=(validate_numeric, "%P"))
        self.idf_entry.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        
        # Crear botón para crear la plantilla
        create_button = ttk.Button(self, text="Crear plantilla", command=self.create_template)
        create_button.grid(pady=1)
        
        # Crear cuadro de texto para la plantilla
        self.template_text = tk.Text(self, height=7, width=40)
        self.template_text.grid(padx=10, pady=10)
        self.template_text.config(state=tk.DISABLED)  # Deshabilitar la edición

        # Agregar barra de desplazamiento vertical al cuadro de texto
        template_scrollbar = ttk.Scrollbar(self, command=self.template_text.yview)
        template_scrollbar.grid(row=2, column=1, sticky="ns")
        self.template_text.config(yscrollcommand=template_scrollbar.set)

        # Crear botón para copiar la plantilla
        copy_button = ttk.Button(self, text="Copiar", command=self.copiar)
        copy_button.grid(row=3, column=0, padx=5, pady=10, sticky="w")
        
        # Crear botón para borrar la plantilla
        delete_button = ttk.Button(self, text="Borrar", command=self.borrar)
        delete_button.grid(row=3, column=0, padx=5, pady=10, sticky="e")
        
    def validate_input(self, P):
        self.idf_entry.delete(0, tk.END)
        self.idf_entry.insert(0, P.strip())
        return True
    
    def create_template(self):
        self.template_text.config(state=tk.NORMAL)
        RQ = self.rq_text.get("1.0", tk.END).strip()
        ID_Final = self.idf_entry.get()
        if not RQ or not ID_Final:
            messagebox.showerror("ERROR","¡Llena los campos!")
            self.template_text.config(state=tk.DISABLED)
            return
        template = f"RQ: {RQ}\nID Final: {ID_Final}"
        self.template_text.delete("1.0", tk.END)  # Borrar el contenido actual
        self.template_text.insert(tk.END, template)
        self.template_text.config(state=tk.DISABLED)  # Deshabilitar la edición después de insertar
    
    def copiar(self):
        template = self.template_text.get("1.0", tk.END)
        self.clipboard_clear()
        self.clipboard_append(template)
        messagebox.showinfo("Copiado", "Plantilla copiada al portapapeles.")
    
    def borrar(self):
        self.template_text.config(state=tk.NORMAL)
        self.template_text.delete("1.0", tk.END)
        self.template_text.config(state=tk.DISABLED)  # Deshabilitar la edición después de borrar
        self.idf_entry.delete(0,"end")
        self.rq_text.delete("1.0", "end")

#ESTE ES EL CÓDIGO PARA EL TIPIFICADOR DE ESIM, JUNTO CON SU INTERFAZ (PESTAÑA 5)
class RepoEsim(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets()
        
    def create_widgets(self):
        self.create_frame_dinamicos()
        self.create_frame_fijos()
        self.create_cuadro_texto()
        self.create_frame_botones()

    def create_frame_dinamicos(self):
        self.frame_dinamicos = ttk.Frame(pestana5)
        self.frame_dinamicos.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        self.label_num_labels = tk.Label(self.frame_dinamicos, text="Número de chips:")
        self.label_num_labels.grid(row=0, column=0, padx=5, pady=2)

        self.entry_num_labels = ttk.Entry(self.frame_dinamicos)
        self.entry_num_labels.grid(row=0, column=1, padx=5, pady=2)

        self.boton_generar_labels = ttk.Button(self.frame_dinamicos, text="Generar campos", command=self.generar_labels_dinamicos)
        self.boton_generar_labels.grid(row=0, column=2, padx=5, pady=2)

        self.frame_canvas = ttk.Frame(self.frame_dinamicos)
        self.frame_canvas.grid(row=1, column=0, columnspan=3, padx=5, pady=0, sticky="ew")

        self.canvas = tk.Canvas(self.frame_canvas)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.frame_dinamicos_interior = ttk.Frame(self.canvas, height=100)
        self.canvas.create_window((0, 0), window=self.frame_dinamicos_interior, anchor="nw")

        self.scrollbar_vertical = ttk.Scrollbar(self.frame_canvas, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(yscrollcommand=self.scrollbar_vertical.set)
        self.frame_dinamicos_interior.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def generar_labels_dinamicos(self):
        num_labels_str = (self.entry_num_labels.get())
        if not num_labels_str:
            # Mostrar un mensaje de error si el campo está vacío
            messagebox.showerror("ERROR", "Debe ingresar un número válido.")
            return

        try:
            num_labels = int(num_labels_str)
            if num_labels <= 0:
                raise ValueError("Debe ingresar un número válido.")
        except ValueError as e:
            # Mostrar un mensaje de error si el valor ingresado no es válido
            messagebox.showerror("ERROR", str(e))
            return

        # Continuar con la generación de etiquetas dinámicas
        for widget in self.frame_dinamicos_interior.winfo_children():
            widget.destroy()

        self.entrys_dinamicos = []
        for i in range(num_labels):
            frame_labels_dinamicos = tk.Frame(self.frame_dinamicos_interior)
            frame_labels_dinamicos.pack(fill=tk.X, padx=5, pady=2)

            label_dinamico_1 = tk.Label(frame_labels_dinamicos, text=f"Esim {i + 1}:")
            label_dinamico_1.pack(side=tk.LEFT, padx=(0, 5))

            entry_dinamico_1 = tk.Entry(frame_labels_dinamicos)
            entry_dinamico_1.pack(side=tk.LEFT)

            label_dinamico_2 = tk.Label(frame_labels_dinamicos, text=f"Línea {i + 1}:")
            label_dinamico_2.pack(side=tk.LEFT, padx=(10, 5))

            entry_dinamico_2 = tk.Entry(frame_labels_dinamicos)
            entry_dinamico_2.pack(side=tk.LEFT)

            self.entrys_dinamicos.extend([entry_dinamico_1, entry_dinamico_2])

        self.frame_dinamicos_interior.update_idletasks()
        interior_height = self.frame_dinamicos_interior.winfo_reqheight()

        self.frame_dinamicos_interior.config(height=100)
        self.canvas.config(scrollregion=self.canvas.bbox("all"), height=min(100, interior_height))

        self.scrollbar_vertical.config(command=self.canvas.yview)
        self.canvas.yview_moveto(0)
        if interior_height > 200:
            self.scrollbar_vertical.set(0, 200 / interior_height)
        else:
            self.scrollbar_vertical.set(0, 1)

    def create_frame_fijos(self):
        self.frame_fijos = ttk.Frame(pestana5)
        self.frame_fijos.pack(padx=10, pady=5, anchor="w")

        self.label_fijo_1 = ttk.Label(self.frame_fijos, text="Cantidad de Esims:")
        self.label_fijo_1.pack(side=tk.LEFT)

        self.entry_fijo_1 = ttk.Entry(self.frame_fijos)
        self.entry_fijo_1.pack(side=tk.LEFT, padx=5)

        self.label_fijo_2 = ttk.Label(self.frame_fijos, text="RUC:")
        self.label_fijo_2.pack(side=tk.LEFT)

        self.entry_fijo_2 = ttk.Entry(self.frame_fijos)
        self.entry_fijo_2.pack(side=tk.LEFT, padx=5)

        self.boton_crear_plantilla = ttk.Button(self.frame_fijos, text="Crear plantilla", command=self.crear_plantilla)
        self.boton_crear_plantilla.pack(side=tk.LEFT, padx=5)

    def create_cuadro_texto(self):
        self.cuadro_texto = tk.Text(pestana5, wrap=tk.WORD, height=7, width=70)
        self.cuadro_texto.pack(padx=10, pady=(0, 20))
        self.cuadro_texto.config(state=tk.DISABLED)  # Deshabilitar la edición
        

    def crear_plantilla(self):
        cantidad_esims = self.entry_fijo_1.get()
        ruc = self.entry_fijo_2.get()

        try:
            cantidad_esims = int(cantidad_esims)
            numero_total = cantidad_esims * 5.35
        except ValueError:
            numero_total = 0

        valores_dinamicos = [entry.get() for entry in self.entrys_dinamicos]

        texto_plantilla = ""
        for i, valor_dinamico in enumerate(valores_dinamicos):
            if i % 2 == 0:
                texto_plantilla += f"Esim {i//2 + 1}: {valor_dinamico}    "
            else:
                texto_plantilla += f"Línea {i//2 + 1}: {valor_dinamico}\n"

        self.cuadro_texto.config(state=tk.NORMAL)
        texto_plantilla += f"Cargo a factura: {numero_total:.2f}$ + IMP\tRUC: {ruc}"
        self.cuadro_texto.delete("1.0", tk.END)
        self.cuadro_texto.insert(tk.END, texto_plantilla)
        self.cuadro_texto.config(state=tk.DISABLED)  # Deshabilitar la edición

    def create_frame_botones(self):
        self.frame_botones = ttk.Frame(pestana5)
        self.frame_botones.pack(pady=5)

        self.boton_copiar_plantilla = ttk.Button(self.frame_botones, text="Copiar Plantilla", command=self.copiar_plantilla)
        self.boton_copiar_plantilla.pack(side=tk.LEFT, padx=20, pady=(0,2))

        self.boton_borrar_datos = ttk.Button(self.frame_botones, text="Borrar Datos", command=self.borrar_datos)
        self.boton_borrar_datos.pack(side=tk.LEFT, padx=20, pady=(0,2))

    def copiar_plantilla(self):
        texto_plantilla = self.cuadro_texto.get("1.0", tk.END)

        self.clipboard_clear()
        self.clipboard_append(texto_plantilla)
        messagebox.showinfo("Copiado", "Plantilla copiada al portapapeles.")

    def borrar_datos(self):
        self.cuadro_texto.config(state=tk.NORMAL)
        self.cuadro_texto.delete("1.0", tk.END)
        self.cuadro_texto.config(state=tk.DISABLED)  # Deshabilitar la edición

#ESTE ES EL CÓDIGO PARA EL TIPIFICADOR DE REPOSICIÓN A DOMICILIO, JUNTO CON SU INTERFAZ (PESTAÑA 6)
class RepoDom(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets()
        
    def create_widgets(self):
        self.create_frame_dinamicos()
        self.create_frame_fijos()
        self.create_cuadro_texto()

    def create_frame_dinamicos(self):
        self.frame_dinamicos = ttk.Frame(pestana6)
        self.frame_dinamicos.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        self.label_num_labels = tk.Label(self.frame_dinamicos, text="Número de chips:")
        self.label_num_labels.grid(row=0, column=0, padx=5, pady=2)

        self.entry_num_labels = ttk.Entry(self.frame_dinamicos)
        self.entry_num_labels.grid(row=0, column=1, padx=5, pady=2)

        self.boton_generar_labels = ttk.Button(self.frame_dinamicos, text="Generar campos", command=self.generar_labels_dinamicos)
        self.boton_generar_labels.grid(row=0, column=2, padx=5, pady=2)

        self.frame_canvas = ttk.Frame(self.frame_dinamicos)
        self.frame_canvas.grid(row=1, column=0, columnspan=3, padx=5, pady=0, sticky="ew")

        self.canvas = tk.Canvas(self.frame_canvas)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.frame_dinamicos_interior = ttk.Frame(self.canvas, height=100)
        self.canvas.create_window((0, 0), window=self.frame_dinamicos_interior, anchor="nw")

        self.scrollbar_vertical = ttk.Scrollbar(self.frame_canvas, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(yscrollcommand=self.scrollbar_vertical.set)
        self.frame_dinamicos_interior.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def generar_labels_dinamicos(self):
        num_labels_str = (self.entry_num_labels.get())
        if not num_labels_str:
            # Mostrar un mensaje de error si el campo está vacío
            messagebox.showerror("ERROR", "Debe ingresar un número válido.")
            return
        try:
            num_labels = int(num_labels_str)
            if num_labels <= 0:
                raise ValueError("Debe ingresar un número válido.")
        except ValueError as e:
            # Mostrar un mensaje de error si el valor ingresado no es válido
            messagebox.showerror("ERROR", str(e))
            return
        messagebox.showinfo("¡RECUERDA!", "Los asesores son:\n AXXXX XXXX (Y2 - XXXX)\n WXXXX XXXX (Y1 - XXXXX)")
        for widget in self.frame_dinamicos_interior.winfo_children():
            widget.destroy()

        self.entrys_dinamicos = []
        label_names = [f"Línea {i}" for i in range(1, num_labels + 1)]

        for label_name in label_names:
            frame_labels_dinamicos = tk.Frame(self.frame_dinamicos_interior)
            frame_labels_dinamicos.pack(fill=tk.X, padx=5, pady=2)

            label_dinamico = tk.Label(frame_labels_dinamicos, text=f"{label_name}:")
            label_dinamico.pack(side=tk.LEFT, padx=(0, 5))

            entry_dinamico = tk.Entry(frame_labels_dinamicos)
            entry_dinamico.pack(side=tk.LEFT)

            self.entrys_dinamicos.append(entry_dinamico)

        self.frame_dinamicos_interior.update_idletasks()
        interior_height = self.frame_dinamicos_interior.winfo_reqheight()

        self.frame_dinamicos_interior.config(height=100)
        self.canvas.config(scrollregion=self.canvas.bbox("all"), height=min(100, interior_height))

        self.scrollbar_vertical.config(command=self.canvas.yview)
        self.canvas.yview_moveto(0)
        if interior_height > 200:
            self.scrollbar_vertical.set(0, 200 / interior_height)
        else:
            self.scrollbar_vertical.set(0, 1)
    
    def mostrar_alerta(self, event):
        motivo_seleccionado = self.motivo_var.get()
        if motivo_seleccionado == "Robo":
            messagebox.showwarning("¡RECUERDA!", "Si el motivo es 'Robo', las líneas en axis deben estar suspendidas por robo.")

    def create_frame_fijos(self):
        self.frame_fijos = ttk.Frame(pestana6)
        self.frame_fijos.pack(padx=10, pady=5, anchor="w")

        # Columna 1
        self.label_fijo_1 = ttk.Label(self.frame_fijos, text="Motivo de reposición:")
        self.label_fijo_1.grid(row=0, column=0, padx=5, pady=2, sticky="w")
        # Lista desplegable (combobox) con opciones "Robo" y "Daño"
        self.motivo_var = tk.StringVar()
        self.motivo_var.set("")  # Opción seleccionada por defecto
        self.motivo_combobox = ttk.Combobox(self.frame_fijos, textvariable=self.motivo_var, values=["", "Robo", "Daño"])
        self.motivo_combobox.grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.label_fijo_2 = ttk.Label(self.frame_fijos, text="RUC:")
        self.label_fijo_2.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        self.entry_fijo_2 = ttk.Entry(self.frame_fijos)
        self.entry_fijo_2.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        # Enlazar el método a la selección de la lista desplegable
        self.motivo_combobox.bind("<<ComboboxSelected>>", self.mostrar_alerta)

        # Columna 2
        self.label_fijo_3 = ttk.Label(self.frame_fijos, text="Persona/N° Contacto 1:")
        self.label_fijo_3.grid(row=0, column=2, padx=5, pady=2, sticky="w")
        self.entry_fijo_3 = ttk.Entry(self.frame_fijos)
        self.entry_fijo_3.grid(row=1, column=2, padx=5, pady=2, sticky="w")

        self.label_fijo_4 = ttk.Label(self.frame_fijos, text="Persona/N° Contacto 2:")
        self.label_fijo_4.grid(row=0, column=3, padx=5, pady=2, sticky="w")
        self.entry_fijo_4 = ttk.Entry(self.frame_fijos)
        self.entry_fijo_4.grid(row=1, column=3, padx=5, pady=2, sticky="w")

        self.boton_crear_plantilla = ttk.Button(self.frame_fijos, text="Crear plantilla", command=self.crear_plantilla)
        self.boton_crear_plantilla.grid(row=2, column=0, columnspan=2, padx=5, pady=2, sticky="ew")

        self.boton_copiar_plantilla = ttk.Button(self.frame_fijos, text="Copiar Plantilla", command=self.copiar_plantilla)
        self.boton_copiar_plantilla.grid(row=2, column=2, padx=5, pady=2, sticky="ew")

        self.boton_borrar_datos = ttk.Button(self.frame_fijos, text="Borrar Datos", command=self.borrar_datos)
        self.boton_borrar_datos.grid(row=2, column=3, padx=5, pady=2, sticky="ew")

    def create_cuadro_texto(self):
        self.cuadro_texto = tk.Text(pestana6, wrap=tk.WORD, height=7, width=70)
        self.cuadro_texto.pack(padx=10, pady=(0, 20))
        self.cuadro_texto.config(state=tk.DISABLED)

    def crear_plantilla(self):
        motivo_reposicion = self.motivo_var.get()
        ruc = self.entry_fijo_2.get()
        persona_contacto_1 = self.entry_fijo_3.get()
        persona_contacto_2 = self.entry_fijo_4.get()

        valores_dinamicos = [entry.get() for entry in self.entrys_dinamicos]

        lineas = ", ".join(valores_dinamicos)
        texto_plantilla = f"Líneas: {lineas}, "
        texto_plantilla += f"Motivo de reposición: {motivo_reposicion}, Ruc: {ruc}, "
        texto_plantilla += f"Persona/Numero de contacto 1: {persona_contacto_1}, "
        texto_plantilla += f"Persona/Número de Contacto 2: {persona_contacto_2}\n"

        self.cuadro_texto.config(state=tk.NORMAL)
        self.cuadro_texto.delete("1.0", tk.END)
        self.cuadro_texto.insert(tk.END, texto_plantilla)
        self.cuadro_texto.config(state=tk.DISABLED)

    def copiar_plantilla(self):
        texto_plantilla = self.cuadro_texto.get("1.0", tk.END)

        self.clipboard_clear()
        self.clipboard_append(texto_plantilla)
        messagebox.showinfo("Copiado", "Plantilla copiada al portapapeles.")

    def borrar_datos(self):
        self.cuadro_texto.config(state=tk.NORMAL)
        self.cuadro_texto.delete("1.0", tk.END)
        self.cuadro_texto.config(state=tk.DISABLED)

#ESTE ES EL CÓDIGO PARA EL TIPIFICADOR DE REPOSICIÓN DE STOCK A DOMICILIO, JUNTO CON SU INTERFAZ (PESTAÑA 8)
class RepoStock(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Crear el frame
        frame = ttk.Frame(self)
        frame.grid(padx=10, pady=10)
        
        # Crear los label con las entry
        sim_label = ttk.Label(frame, text="Cant. Simcard:")
        sim_label.grid(row=0, column=0, sticky="w")
        self.sim_entry = ttk.Entry(frame)
        self.sim_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")   

        bscs_label = ttk.Label(frame, text="Cuenta Bscs:")
        bscs_label.grid(row=0, column=2, sticky="w")
        self.bscs_entry = ttk.Entry(frame)
        self.bscs_entry.grid(row=0, column=3, padx=10, pady=5, sticky="w")  

        num_ligado_label = ttk.Label(frame, text="Número ligado:")
        num_ligado_label.grid(row=1, column=0, sticky="w")
        self.num_ligado_entry = ttk.Entry(frame)
        self.num_ligado_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")  

        ruc_label = ttk.Label(frame, text="RUC:")
        ruc_label.grid(row=1, column=2, sticky="w")
        self.ruc_entry = ttk.Entry(frame)
        self.ruc_entry.grid(row=1, column=3, padx=10, pady=5, sticky="w")  

        persona_1_label = ttk.Label(frame, text="Persona/N° Contacto 1:")
        persona_1_label.grid(row=2, column=0, sticky="w")
        self.persona_1_entry = ttk.Entry(frame)
        self.persona_1_entry.grid(row=2, column=1, padx=10, pady=5, sticky="w")  

        persona_2_label = ttk.Label(frame, text="Persona/N° Contacto 2:")
        persona_2_label.grid(row=2, column=2, sticky="w")
        self.persona_2_entry = ttk.Entry(frame)
        self.persona_2_entry.grid(row=2, column=3, padx=10, pady=5, sticky="w")  
        
        # Crear botón para crear la plantilla
        create_button = ttk.Button(self, text="Crear plantilla", command=self.generate_template)
        create_button.grid(pady=1)
        
        # Crear cuadro de texto para la plantilla
        self.template_text = tk.Text(self, height=8, width=65)
        self.template_text.grid(padx=10, pady=10)
        self.template_text.config(state=tk.DISABLED)

        # Agregar barra de desplazamiento vertical al cuadro de texto
        template_scrollbar = ttk.Scrollbar(self, command=self.template_text.yview)
        template_scrollbar.grid(row=2, column=1, sticky="ns")
        self.template_text.config(yscrollcommand=template_scrollbar.set)

        # Crear botón para copiar la plantilla
        copy_button = ttk.Button(self, text="Copiar", command=self.copiar)
        copy_button.grid(row=3, column=0, padx=5, pady=10, sticky="w")
        
        # Crear botón para borrar la plantilla
        delete_button = ttk.Button(self, text="Borrar", command=self.borrar)
        delete_button.grid(row=3, column=0, padx=5, pady=10, sticky="e")
    
    def generate_template(self):
        cantidad_simcards = self.sim_entry.get()
        cuenta_bscs = self.bscs_entry.get()
        numero_ligado = self.num_ligado_entry.get()
        ruc = self.ruc_entry.get()
        persona_contacto_1 = self.persona_1_entry.get()
        persona_contacto_2 = self.persona_2_entry.get()

        template = f"Cantidad de simcard(s): {cantidad_simcards}, Motivo de reposición: Stock, " \
                   f"Cuenta XXXXX: {cuenta_bscs}, Número ligado a la cuenta XXXXX: {numero_ligado}, " \
                   f"Ruc: {ruc}, Persona/Numero de contacto 1: {persona_contacto_1}, " \
                   f"Persona/Numero de contacto 2: {persona_contacto_2}"
        
        self.template_text.config(state=tk.NORMAL)
        self.template_text.delete("1.0", tk.END)
        self.template_text.insert("1.0", template)
        self.template_text.config(state=tk.DISABLED)

    def copiar(self):
        template = self.template_text.get("1.0", tk.END)
        self.clipboard_clear()
        self.clipboard_append(template)
        messagebox.showinfo("Copiado", "Plantilla copiada al portapapeles.")
    
    def borrar(self):
        self.template_text.config(state=tk.NORMAL)
        self.template_text.delete("1.0", tk.END)
        self.template_text.config(state=tk.DISABLED)

#ESTE ES EL CÓDIGO PARA EL TIPIFICADOR DE MALAS DERIVACIONES, JUNTO CON SU INTERFAZ (PESTAÑA 9)
class MalDerivado(ttk.Frame):
    def __init__(self, pestana9):
        super().__init__(pestana9)
        self.inicializar_google_sheets()

        # Agregar el texto arriba del frame
        texto_arriba = "ALERTA DE CASOS SIN LLAMADA O GESTIONADOS INCORRECTAMENTE."
        texto_label = ttk.Label(self, text=texto_arriba, font=("Arial", 11))
        texto_label.grid(row=0, column=0, columnspan=5, padx=10, pady=10)

        # Crear el frame
        frame = ttk.Frame(self)
        frame.grid(padx=10, pady=10)
        
        # Crear los label con las entry
        self.psv_implicado_label = ttk.Label(frame, text="Usuario CRM:")
        self.psv_implicado_label.grid(row=0, column=0, sticky="w")
        self.psv_implicado_entry = ttk.Entry(frame, width=23)
        self.psv_implicado_entry.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        
        self.tramite_crm_label = ttk.Label(frame, text="Trámite CRM:")
        self.tramite_crm_label.grid(row=0, column=2, sticky="w")
        validate_numeric = self.register(self.validate_numeric_input)
        self.tramite_crm_entry = ttk.Entry(frame, width=23, validate="key", validatecommand=(validate_numeric, "%P"))
        self.tramite_crm_entry.grid(row=0, column=3, padx=5, pady=2, sticky="w")
        
        self.mala_gestion_label = ttk.Label(frame, text="Mala gestión:")
        self.mala_gestion_label.grid(row=1, column=0, sticky="w")
        # Lista desplegable (combobox) con opciones "No llamadas" y "Gestiona Mal"
        self.mala_gestion_var = tk.StringVar()
        self.mala_gestion_var.set("")  # Opción seleccionada por defecto
        self.mala_gestion_combobox = ttk.Combobox(frame, width=20, textvariable=self.mala_gestion_var, values=["", "No realiza llamadas", "Gestión incorrecta"])
        self.mala_gestion_combobox.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # Crear cuadro de texto para la observación
        self.observacion_label = ttk.Label(frame, text="Observación:")
        self.observacion_label.grid(row=2, column=0, sticky="w")
        self.observacion_text = tk.Text(frame, height=7, width=50)
        self.observacion_text.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        # Agregar barra de desplazamiento vertical
        rq_scrollbar = ttk.Scrollbar(frame, command=self.observacion_text.yview)
        rq_scrollbar.grid(row=2, column=4, sticky="ns")
        self.observacion_text.config(yscrollcommand=rq_scrollbar.set)

        # Botón "Reportar" para enviar los datos a Google Sheets
        self.boton_reportar = ttk.Button(frame, text="Reportar", command=self.obtener_valores)
        self.boton_reportar.grid(row=3, column=0, columnspan=2, pady=10)  # columnspan=4 indica que el botón ocupa cuatro columnas

        # Botón "Limpiar Campos" para borrar los datos
        self.boton_limpiar = ttk.Button(frame, text="Limpiar Campos", command=self.limpiar_campos)
        self.boton_limpiar.grid(row=3, column=3, columnspan=2, pady=10)
    
    def validate_numeric_input(self, value):
        return value.isdigit() or value == ""
    
    def limpiar_campos(self):
        # Borra los datos ingresados en los widgets
        self.psv_implicado_entry.delete(0, tk.END)  # Borra el contenido del Entry
        self.tramite_crm_entry.delete(0, tk.END)
        self.mala_gestion_combobox.set("")  # Reinicia la selección del Combobox
        self.observacion_text.delete("1.0", tk.END)  # Borra el contenido del Text

    def obtener_valores(self):
        # Obtener la siguiente fila vacía en la hoja
        siguiente_fila_vacia = len(self.hoja.get_all_values()) + 1

        valor_psv_implicado = self.psv_implicado_entry.get()
        valor_tramite_crm = self.tramite_crm_entry.get()
        valor_mala_gestion = self.mala_gestion_var.get()
        valor_observacion = self.observacion_text.get("1.0", tk.END)

        # Validar que todos los campos estén llenos antes de enviar a Google Sheets
        if valor_psv_implicado and valor_tramite_crm and valor_mala_gestion and valor_observacion:
            self.enviar_a_google_sheets(siguiente_fila_vacia)
            messagebox.showinfo("MENSAJE", "El caso fue reportado exitosamente.")
        else:
            messagebox.showerror("ERROR", "Por favor, complete todos los campos.")
    
    def inicializar_google_sheets(self):
        try:
            # Definir alcance y obtener las credenciales desde el archivo JSON
            scope = ["Here goes your spreadsheets", "Here goes your API"]
            credentials = ServiceAccountCredentials.from_json_keyfile_name('Insert your .json file here', scope)  # Reemplaza con el nombre de tu archivo de credenciales

            # Autenticar y abrir el libro "Base Reportes"
            self.gc = gspread.authorize(credentials)
            self.nombre_libro = "Base Reportes"
            self.libro = self.gc.open(self.nombre_libro)

            # Obtener la hoja "Mal derivados"
            self.nombre_hoja = "Mal derivados"
            self.hoja = self.libro.worksheet(self.nombre_hoja)

        except Exception as e:
            print("Error al inicializar Google Sheets:", str(e))

    def enviar_a_google_sheets(self, siguiente_fila_vacia):
        try:
            # Obtener la fecha y hora actual en formato dd/mm/yyyy hh:mm
            fecha_hora_actual = datetime.now().strftime('%d/%m/%Y %H:%M')

            valor_psv_implicado = self.psv_implicado_entry.get()
            valor_tramite_crm = self.tramite_crm_entry.get()
            valor_mala_gestion = self.mala_gestion_var.get()
            valor_observacion = self.observacion_text.get("1.0", tk.END)

            # Enviar los datos a la hoja de Google Sheets
            fila = [fecha_hora_actual, valor_psv_implicado, valor_tramite_crm, valor_mala_gestion, valor_observacion]
            self.hoja.append_row(fila)

            print("Datos enviados a Google Sheets.")

        except Exception as e:
            print("Error al enviar a Google Sheets:", str(e))

#ESTE ES EL CÓDIGO PARA DESCARGAR EL TICKET Y SUBIRLO AL DRIVE (PESTAÑA 7)
class BackUpTicket(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.usuario_var = tk.StringVar()
        self.clave_var = tk.StringVar()
        self.ticket_var = tk.StringVar()
        self.label_fechaEscogida_var = tk.StringVar()
        self.create_tree()

        # Configurar el comportamiento de expansión de filas y columnas
        self.pack(expand=True, fill='both')

        # Frame izquierdo
        frame_izquierdo = ttk.Frame(self)
        frame_izquierdo.pack(side='left', padx=10, pady=5)

        # Etiqueta y campo de entrada para el usuario
        self.label_usuario = ttk.Label(frame_izquierdo, text="Usuario:")
        self.label_usuario.pack(anchor='w')
        self.entry_usuario = ttk.Entry(frame_izquierdo, textvariable=self.usuario_var)
        self.entry_usuario.pack(anchor='e', padx=10, pady=5)

        # Etiqueta y campo de entrada para la clave
        self.label_clave = ttk.Label(frame_izquierdo, text="Clave:")
        self.label_clave.pack(anchor='w')
        self.entry_clave = ttk.Entry(frame_izquierdo, textvariable=self.clave_var, show="*")
        self.entry_clave.pack(anchor='e', padx=10, pady=5)

        # Etiqueta y campo de entrada para el ticket
        validate_numeric = self.register(self.validate_input)
        self.label_ticket = ttk.Label(frame_izquierdo, text="Ticket:")
        self.label_ticket.pack(anchor='w')
        self.entry_ticket = ttk.Entry(frame_izquierdo, textvariable=self.ticket_var, validate="key", validatecommand=(validate_numeric, "%P"))
        self.entry_ticket.pack(anchor='e', padx=10, pady=5)

        self.btn_buscar = ttk.Button(frame_izquierdo, text="Cargar", command=self.run_search)
        self.btn_buscar.pack(anchor='w', padx=35, pady=1)

        # Frame derecho
        frame_derecho = ttk.Frame(self)
        frame_derecho.pack(side='right', padx=10, pady=0)

        label_seleccionar = ttk.Label(frame_derecho, text="Seleccionar Fecha:")
        label_seleccionar.pack(side='top', anchor='w', padx=0, pady=0)

        # Frame para contener label_seleccionar y self.date_entry
        frame_fecha = ttk.Frame(frame_derecho)
        frame_fecha.pack(side='top', anchor='w')

        self.date_entry = DateEntry(frame_fecha, date_pattern="dd/mm/yyyy", width=12, locale="es_ES")
        self.date_entry.pack(side='left', padx=(0,20), pady=1)

        # Crear botón de búsqueda
        self.botonazo = ttk.Button(frame_fecha, text="Buscar", command=self.GenerarbusquedaAsync)
        self.botonazo.pack(side="left", padx=(10, 20), pady=1)

        self.label_gestionados_var = tk.StringVar()
        label_gestionados = ttk.Label(frame_derecho, textvariable=self.label_gestionados_var)
        label_gestionados.pack(side='top', anchor='w', pady=(10, 0))

        self.label_finalizados_var = tk.StringVar()
        label_finalizados = ttk.Label(frame_derecho, textvariable=self.label_finalizados_var)
        label_finalizados.pack(side='top', anchor='w', pady=(5, 0))

        self.label_derivados_var = tk.StringVar()
        label_derivados = ttk.Label(frame_derecho, textvariable=self.label_derivados_var)
        label_derivados.pack(side='top', anchor='w', pady=(5, 0))

        self.label_efectividad_var = tk.StringVar()
        label_efectividad = ttk.Label(frame_derecho, textvariable=self.label_efectividad_var)
        label_efectividad.pack(side='top', anchor='w', pady=(5, 0))

        self.actualizar_labels()

        self.label_fechaEscogida_var = tk.StringVar()
        label_fechaEscogida = ttk.Label(frame_derecho, textvariable=self.label_fechaEscogida_var)
        label_fechaEscogida.pack(side='top', anchor='w', pady=(5, 0))
        # Ocultamos el label haciendo que se olvide de él
        label_fechaEscogida.pack_forget()
        # Asignar la función select_date al evento de selección de fecha
        self.date_entry.bind("<<DateEntrySelected>>", self.select_date)
    
    def actualizar_labels(self):
        # Obtener los valores de la columna "Estado Actual trámite" en el Treeview
        estados = [self.tree.set(item, "Estado Actual trámite") for item in self.tree.get_children()]

        # Contar la cantidad de ocurrencias de cada estado
        conteo_finalizados = estados.count("Finalizado")
        conteo_gestionados = sum(1 for estado in estados if estado in ["Escalado", "Finalizado", "Agendamientos", "En proceso", "Pendiente"])
        conteo_derivados = estados.count("Derivado")

        # Actualizar los labels
        self.label_gestionados_var.set(f"Gestionados: {conteo_gestionados}")
        self.label_finalizados_var.set(f"Finalizados: {conteo_finalizados}")
        self.label_derivados_var.set(f"Derivados: {conteo_derivados}")

        # Calcular efectividad y actualizar el label correspondiente
        efectividad = conteo_finalizados / conteo_gestionados if conteo_gestionados != 0 else 0
        self.label_efectividad_var.set(f"Efectividad F/G: {efectividad:.2f}")

    def create_tree(self):
        # Frame inferior
        frame_inferior = ttk.Frame(self)
        frame_inferior.pack(fill='both', padx=10, pady=5)

        # Crear el contenedor para las barras de desplazamiento
        scrollbar_frame = ttk.Frame(frame_inferior)
        scrollbar_frame.pack(expand=True, fill='both')
        
        # Crear la barra de desplazamiento vertical
        y_scrollbar = Scrollbar(scrollbar_frame, orient="vertical")
        y_scrollbar.pack(side='right', fill='y')

        # Crear la barra de desplazamiento horizontal
        x_scrollbar = Scrollbar(scrollbar_frame, orient="horizontal")
        x_scrollbar.pack(side='bottom', fill='x')

        # Crear el treeview con las barras de desplazamiento
        self.tree = ttk.Treeview(scrollbar_frame, columns=("Ticket", "Cliente", "CI/RUC", "Area Derivación", "Fecha Actual trámite", "Usr. Últ. Modificación", "Estado Actual trámite", "Subestado"), show="headings", yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set, height=6)
        self.tree.pack(expand=True, fill='both')

        # Configurar las barras de desplazamiento para sincronizarse con el treeview
        y_scrollbar.config(command=self.tree.yview)
        x_scrollbar.config(command=self.tree.xview)

        # Encabezados del TreeView
        self.tree.heading("Ticket", text="Ticket")
        self.tree.heading("Cliente", text="Cliente")
        self.tree.heading("CI/RUC", text="Ruc")
        self.tree.heading("Area Derivación", text="Area Derivación")
        self.tree.heading("Fecha Actual trámite", text="Fecha Actual trámite")
        self.tree.heading("Usr. Últ. Modificación", text="Usr. Últ. Modificación")
        self.tree.heading("Estado Actual trámite", text="Estado Actual trámite")
        self.tree.heading("Subestado", text="Subestado")

        # Establecer el ancho de cada columna utilizando weight en lugar de un ancho fijo
        self.tree.column("Ticket", width=1, minwidth=50, stretch=True)
        self.tree.column("Cliente", width=1, minwidth=200, stretch=True)
        self.tree.column("CI/RUC", width=1, minwidth=150, stretch=True)
        self.tree.column("Area Derivación", width=1, minwidth=150, stretch=True)
        self.tree.column("Fecha Actual trámite", width=1, minwidth=150, stretch=True)
        self.tree.column("Usr. Últ. Modificación", width=1, minwidth=80, stretch=True)
        self.tree.column("Estado Actual trámite", width=1, minwidth=150, stretch=True)
        self.tree.column("Subestado", width=1, minwidth=150, stretch=True)

        # Ajustar el ancho del treeview al ancho disponible en el frame_inferior
        scrollbar_frame.grid_columnconfigure(0, weight=1)

    def enable_search_button(self):
        self.btn_buscar.config(state=tk.NORMAL)  # Habilitar el botón
    
    def enable_search_button2(self):
        self.botonazo.config(state=tk.NORMAL)  # Habilitar el botón

    def validate_input(self, P):
        self.entry_ticket.delete(0, tk.END)
        self.entry_ticket.insert(0, P.strip())
        return True
    
    def run_search(self):
        usuario = self.usuario_var.get()
        clave = self.clave_var.get()
        ticket = self.ticket_var.get()

        if not usuario or not clave or not ticket:
            # Al menos uno de los campos está vacío, no ejecutar la búsqueda
            messagebox.showerror("ERROR","¡Llena todos los campos!")
            return

        # Deshabilitar el botón "Cargar"
        self.btn_buscar.config(state=tk.DISABLED)

        # Crear un hilo para ejecutar la búsqueda
        search_thread = threading.Thread(target=self.execute_search, args=(usuario, clave, ticket))
        search_thread.start()
  
    def execute_search(self, usuario, clave, ticket):
        try:
            DescargarTicket.mainup(usuario, clave, ticket)
            ticket = str(self.ticket_var.get)
            print(ticket)
        except Exception as e:
            print(f"Error al descargar: {e}")
            # Puede personalizar este mensaje de error según sus necesidades
            messagebox.showerror("ERROR", "Error al descargar. Valida e intenta nuevamente.")
        finally:
            # Habilitar el botón "Buscar" independientemente de si hubo un error o no
            self.after(0, self.enable_search_button)
           
    def select_date(self, event=None):
        date = self.date_entry.get_date().strftime("%d-%m-%Y")
        self.label_fechaEscogida_var.set(date)
        print("Fecha seleccionada:", date)
        return date
    
    def GenerarbusquedaAsync(self):
        # Obtener los valores necesarios
        valor1 = self.select_date()  # "07/07/2023"
        valor2 = self.usuario_var.get()  # "PSVEGODG"
        clave = self.clave_var.get()
        # Variable global para indicar si se debe cancelar la búsqueda
        cancelar_busqueda = False
        # Eliminar los resultados anteriores
        self.tree.delete(*self.tree.get_children())
        
        def buscar_y_actualizar():
            nonlocal cancelar_busqueda  # Hacemos referencia a la variable global
            if cancelar_busqueda:  # Verifica si se debe cancelar la búsqueda
                return

            if not valor2 or not clave:
                # Al menos uno de los campos está vacío, no ejecutar la búsqueda
                messagebox.showerror("ERROR","¡Llena el usuario y clave!")
                self.botonazo.config(state=tk.NORMAL)
                return
            # Deshabilitar el botón "Buscar"
            self.botonazo.config(state=tk.DISABLED)
            # Ejecutar la búsqueda y mostrar los resultados en el Treeview
            resultado_busqueda = ProduAs.searchdate(clave, valor1, valor2)
            print(resultado_busqueda)
            self.after(0, self.enable_search_button2)
            if resultado_busqueda is None:
                # Terminar la ejecución si no se obtienen resultados
                self.after(0, self.enable_search_button2)
                return
            if len(resultado_busqueda) == 0:
                self.tree.insert("", "end", values=("", "No se encontraron resultados.", "", "", "", "", "", ""))
            else:
                for resultado in resultado_busqueda:
                    self.tree.insert("", "end", values=resultado)

            # Obtener los valores de la columna "Estado Actual trámite" en el Treeview
            estados = [self.tree.set(item, "Estado Actual trámite") for item in self.tree.get_children()]

            # Contar la cantidad de ocurrencias de cada estado
            conteo_finalizados = estados.count("Finalizado")
            conteo_gestionados = sum(1 for estado in estados if estado in ["Escalado", "Finalizado", "Agendamientos", "En proceso", "Pendiente"])
            conteo_derivados = estados.count("Ingresado")

            # Actualizar los labels
            self.label_gestionados_var.set(f"Gestionados: {conteo_gestionados}")
            self.label_finalizados_var.set(f"Finalizados: {conteo_finalizados}")
            self.label_derivados_var.set(f"Derivados: {conteo_derivados}")

            # Calcular efectividad y actualizar el label correspondiente
            efectividad = conteo_finalizados / conteo_gestionados * 100 if conteo_gestionados != 0 else 0
            self.label_efectividad_var.set(f"Efectividad F/G: {efectividad:.2f}%")
            
        # Función para manejar la pulsación de F5 y cancelar la búsqueda
        def manejar_tecla_F5(e):
            nonlocal cancelar_busqueda
            if e.event_type == keyboard.KEY_DOWN:
                messagebox.showwarning("Alerta", "Se ha presionado F5. Cancelando la búsqueda...")
            cancelar_busqueda = True
            self.after(0, self.enable_search_button2)
        # Elimina todos los manejadores de tecla antes de agregar uno nuevo
        keyboard.unhook_all()
        # Configura la detección de la tecla F5
        keyboard.on_press_key("F5", manejar_tecla_F5)

        # Ejecutar la función de búsqueda en un hilo separado
        thread = threading.Thread(target=buscar_y_actualizar)
        thread.start()

#ESTE ES EL CÓDIGO QUE CONTIENE LA INFORMACIÓN DEL PROGRAMA (PESTAÑA 10)
class Informacion(ttk.Frame):
    def __init__(self, pestana10):
        super().__init__(pestana10)

        self.tab_text = "Información"
        # Crear el widget Text
        self.text_widget = tk.Text(pestana10, wrap='word')
        self.text_widget.pack(fill='both', expand=True)

        # Texto largo con sangrías
        long_text = """
    Esta es la versión 2.5 del programa "Gestor del Asesor". Corrige errores de la segunda versión.
    El objetivo del programa es facilitar la consulta en bases, brindar tipificadores y ser de ayuda para una gestión más ágil del operador. 
    Fue desarrollado en Python 3.11. Utiliza Chromedriver (para Google Chrome 114).

    A continuación lo que permite realizar cada pestaña:

    MENÚS: 
        1. BÚSQUEDA EN BASES:
            1.1. Autorizados. Permite generar la búsqueda de autorizados en la base de XXXXXXX, mediante RUC, celular o correo.
            1.2. Ejec. y Recomend. Permite generar la búsqueda de los ejecutivos asignados a una empresa, así como de si el cliente es recomendado o no. Esto mediante el RUC.

        2. TIPIFICADORES:
            2.1. ACDXXXX. Es el tipificador para asesores de ACDXXXXX. Este se usa para colocar el comentario en el CRM.
            2.2. SegXXXXX. Es el tipificador para asesores de SegXXXXXX. Este se usa para colocar el comentario en el CRM.
            2.3 Casos Mal Gestionados. Se reporta aquellos casos donde se valide una mala gestión de algún asesor.
            2.4 Reposición de Simcard: Esim. Es el tipificador para reposiciones de esim, de una a varios servicios.
            2.5 Reposición de Simcard: A domicilio. Es el tipificador para reposiciones de chip a domicilio, de uno a varios servicios.
            2.6 Reposición de Simcard: Stock a domicilio. Es el tipificador para reposiciones de chip en stock con envío a domicilio, de uno a varios servicios.

        3. PRODUCTIVIDAD:
            3.1. Productividad. En esta pestaña se puede cargar los tickets que van gestionando. Una vez modificado, se carga y de esta forma se puede obtener un registro de la productividad. De igual forma, permite realizar la consulta de la misma, pero hace una validación de usuario y clave previamente. Se puede validar los gestionados, finalizados, derivados y el portcentaje de efectividad sobre lo gestionado.
    """

        # Eliminar sangrías en los niveles que no deben tener
        formatted_text = "\n".join(line.strip() for line in long_text.splitlines())

        # Insertar el texto en el widget Text
        self.text_widget.insert('1.0', formatted_text)

        # Cambiar el tipo de letra
        custom_font = font.Font(family="Arial", size=9)
        self.text_widget.configure(font=custom_font)

        # Configurar comportamiento de redimensionamiento
        self.text_widget.config(state='disabled')  # Para evitar la edición del texto
        self.text_widget.bind("<Configure>", lambda e: self.text_widget.configure(width=e.width))

# Crear la ventana principal
ventana = tk.Tk()
ventana.geometry("600x400")
ventana.iconbitmap('setting.ico')
ventana.title("Gestor de ayuda del asesor")

# Crear el control de pestañas
notebook = ttk.Notebook(ventana)

# Crear las pestañas
pestana1 = ttk.Frame(notebook)
pestana2 = ttk.Frame(notebook)
pestana3 = ttk.Frame(notebook)
pestana4 = ttk.Frame(notebook)
pestana5 = ttk.Frame(notebook)
pestana6 = ttk.Frame(notebook)
pestana7 = ttk.Frame(notebook)
pestana8 = ttk.Frame(notebook)
pestana9 = ttk.Frame(notebook)
pestana10 = ttk.Frame(notebook)

# Agregar contenido a las pestañas
#CONTENIDO DE LA PESTAÑA 1
# Crear la instancia de BusquedaEjeReco
template_creator = BusquedaAutorizados(pestana1)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 2
# Crear la instancia de BusquedaEjeReco
template_creator = BusquedaEjeReco(pestana2)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 3
# Agregar la instancia de TemplateCreatorACD a la pestaña 3
template_creator = TemplateCreatorACD(pestana3)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 4
# Agregar la instancia de TemplateCreatorSEG a la pestaña 4
template_creator = TemplateCreatorSEG(pestana4)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 5
# Agregar la instancia de RepoEsim a la pestaña 5
template_creator = RepoEsim(pestana5)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 6
# Agregar la instancia de RepoDom a la pestaña 6
template_creator = RepoDom(pestana6)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 7
# Agregar la instancia de BackUpTicket a la pestaña 7
template_creator = BackUpTicket(pestana7)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 8
# Agregar la instancia de BackUpTicket a la pestaña 8
template_creator = RepoStock(pestana8)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 9
# Agregar la instancia de BackUpTicket a la pestaña 9
template_creator = MalDerivado(pestana9)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

#CONTENIDO DE LA PESTAÑA 10
# Agregar la instancia de BackUpTicket a la pestaña 9
template_creator = Informacion(pestana10)
template_creator.pack(fill=tk.BOTH, expand=True)  # Agregar el controlador de eventos de tamaño a la pestaña

# Agregar las pestañas al control de pestañas
notebook.add(pestana1, text="Autorizados", state="hidden")
notebook.add(pestana2, text="Ejecutivos y Recomendados", state="hidden")
notebook.add(pestana3, text="ACDMail", state="hidden")
notebook.add(pestana4, text="Seguimiento", state="hidden")
notebook.add(pestana5, text="Reposición de Simcard: Esim", state="hidden")
notebook.add(pestana6, text="Reposición de Simcard: A domicilio", state="hidden")
notebook.add(pestana8, text="Reposición de Simcard: Stock a domicilio", state="hidden")
notebook.add(pestana9, text="Casos Mal Gestionados", state="hidden")
notebook.add(pestana7, text="Productividad", state="hidden")
notebook.add(pestana10, text="Información", state="hidden")
notebook.pack(fill=tk.BOTH, expand=True)

# Crear el control de menú
menu = tk.Menu(ventana)
ventana.config(menu=menu)
sub_menu = tk.Menu(menu, tearoff=True)

# Crear el menú desplegable con las opciones de las pestañas
opciones_desplegables = tk.Menu(menu, tearoff=False)

def mostrar_mensaje_casos_mal_gestionados():
    mensaje = "Alertar aquellos casos donde estando dentro del horario de llamada el asesor no ejecute el requerimiento; o casos donde el requerimiento sea claro, sin embargo no se procese o se haga de forma incorrecta."
    messagebox.showinfo("ALERTA", mensaje)

def on_tab_selected(event):
    # Obtener el índice de la pestaña seleccionada
    selected_index = event.widget.index("current")
    # Verificar si la pestaña seleccionada es la pestaña 9 (índice 8)
    if selected_index == 7:
        mostrar_mensaje_casos_mal_gestionados()

# Asociar el evento de selección de pestaña a la función on_tab_selected
notebook.bind("<<NotebookTabChanged>>", on_tab_selected)

def mostrar_pestanas():
    for index in range(notebook.index("end")):
        notebook.tab(index, state="normal")
    # Verificar si la pestaña seleccionada es la pestaña 9 (índice 8)
    selected_index = notebook.index(notebook.select())
    if selected_index == 7:
        mostrar_mensaje_casos_mal_gestionados()

#Primera opción del menú
opciones_desplegables.add_command(label="Autorizados", command=lambda: mostrar_pestana(pestana1))
opciones_desplegables.add_command(label="Ejec. y Recomend.", command=lambda: mostrar_pestana(pestana2))

#Segunda opción del menú
opciones_desplegables_segunda = tk.Menu(menu, tearoff=False)
opciones_desplegables_segunda.add_command(label="ACDXXXX", command=lambda: mostrar_pestana(pestana3))
opciones_desplegables_segunda.add_command(label="SegXXXXXXX", command=lambda: mostrar_pestana(pestana4))
opciones_desplegables_segunda.add_command(label="Casos Mal Gestionados", command=lambda: mostrar_pestana(pestana9))

#Subopciones de reposición de la segunda opción del menú
reposicion_menu = tk.Menu(opciones_desplegables_segunda, tearoff=False)
reposicion_menu.add_command(label="Esim", command=lambda: mostrar_pestana(pestana5))
reposicion_menu.add_command(label="A domicilio", command=lambda: mostrar_pestana(pestana6))
reposicion_menu.add_command(label="Stock a domicilio", command=lambda: mostrar_pestana(pestana8))

#Tercera opción del menú
opciones_desplegables_tercera = tk.Menu(menu, tearoff=False)
opciones_desplegables_tercera.add_command(label="Productividad", command=lambda: mostrar_pestana(pestana7))

#Cuarta opción del menú
opciones_desplegables_cuarta = tk.Menu(menu, tearoff=False)
opciones_desplegables_cuarta.add_command(label="Información", command=lambda: mostrar_pestana(pestana10))

# Agregar los menú desplegable al control de menú
menu.add_cascade(label="Búsqueda en bases", menu=opciones_desplegables)
menu.add_cascade(label="Tipificadores", menu=opciones_desplegables_segunda)
menu.add_cascade(label="Productividad", menu=opciones_desplegables_tercera)
menu.add_cascade(label="Acerca de", menu=opciones_desplegables_cuarta)
opciones_desplegables_segunda.add_cascade(label="Reposición de Simcard", menu=reposicion_menu)

# Mostrar la ventana
ventana.mainloop()