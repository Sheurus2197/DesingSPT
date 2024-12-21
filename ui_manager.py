#Contiene la clase UIManager, responsable de la interfaz gráfica.

import tkinter as tk
from tkinter import Menu, messagebox, ttk
from geometry_manager import GeometryManager
from resistivity_manager import ResistivityManager
from transformer_manager import TransformerManager

class UIManager:
    """
    Clase responsable de gestionar la interfaz gráfica de usuario (GUI).
    """

    def __init__(self, excel_manager):
        """
        Inicializa el gestor de la interfaz gráfica.

        :param excel_manager: Instancia de ExcelManager para manipular archivos Excel.
        """
        self.excel_manager = excel_manager
        self.root = tk.Tk()
        self.geometry_manager = GeometryManager(self.excel_manager)
        self.resistivity_manager = ResistivityManager(self.excel_manager, self.root)
        self.transformer_manager = TransformerManager(self.excel_manager)

    def iniciar_interfaz(self):
        """
        Configura e inicia la ventana principal.
        """
        # Configurar la ventana principal
        self.configurar_ventana()

        # Configurar el frame principal y las columnas/filas
        self.configurar_main_frame()

        # Menús
        self.crear_menus()

        # Arrancar el bucle principal
        self.root.mainloop()

    def configurar_ventana(self):
        """
        Configura las propiedades de la ventana principal.
        """
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight() - 80
        self.root.geometry(f"{screen_width}x{screen_height}")
        self.root.iconbitmap('icon.ico')

    def configurar_main_frame(self):
        """
        Configura el frame principal con columnas y filas ajustables.
        """
        # Crear un Frame principal
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True)

        # Configurar columnas
        self.main_frame.grid_columnconfigure(0, weight=1)  # Columna 1
        self.main_frame.grid_columnconfigure(1, weight=4)  # Columna 2
        self.main_frame.grid_columnconfigure(5, weight=1)  # Columna 3

        # Configurar filas
        self.main_frame.grid_rowconfigure(0, weight=1)  # Fila 0

        # Crear frames para cada columna
        self.columna_1 = tk.Frame(self.main_frame)
        self.columna_1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.columna_2 = tk.Frame(self.main_frame)
        self.columna_2.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        self.columna_3 = tk.Frame(self.main_frame)
        self.columna_3.grid(row=0, column=5, padx=5, pady=5, sticky="nsew")

        # Canvas dentro de la columna 2
        self.configurar_canvas_columna_2()

        # Agregar frames adicionales
        self.configurar_informacion_proyecto(self.columna_1)
        self.configurar_objetivos(self.columna_1)

        # Configurar frames dentro de Columna 3
        self.configurar_frame_resistencia(self.columna_3)
        self.configurar_frame_seguridad(self.columna_3)
        self.configurar_frame_sugerencias(self.columna_3)

    def configurar_informacion_proyecto(self, parent_frame):
        """
        Configura el frame de información del proyecto.
        """
        info_frame = tk.LabelFrame(parent_frame, text="Información del Proyecto", pady=5, padx=5, font=("Arial", 14))
        info_frame.grid(row=0, column=0, padx=5, pady=5, rowspan=2, sticky="w")

        # Título del Proyecto
        tk.Label(info_frame, text="Título del proyecto: ", font=("Arial", 12)).grid(row=0, column=0, sticky="e", pady=5)
        self.titulo_proyecto_val = tk.Label(info_frame, text="Proyecto 1", font=("Arial", 12))
        self.titulo_proyecto_val.grid(row=0, column=1, sticky="w")

        # Tensión Primario
        tk.Label(info_frame, text="Tensión Primario: ", font=("Arial", 12)).grid(row=1, column=0, sticky="e", pady=5)
        self.tension_primario_val = tk.Label(info_frame, text="", font=("Arial", 12))
        self.tension_primario_val.grid(row=1, column=1, sticky="w")

        # Tensión Secundario
        tk.Label(info_frame, text="Tensión Secundario: ", font=("Arial", 12)).grid(row=2, column=0, sticky="e", pady=5)
        self.tension_secundario_val = tk.Label(info_frame, text="", font=("Arial", 12))
        self.tension_secundario_val.grid(row=2, column=1, sticky="w")

        # Potencia Nominal
        tk.Label(info_frame, text="Potencia Nominal:", font=("Arial", 12)).grid(row=3, column=0, sticky="e", pady=5)
        self.potencia_nominal_val = tk.Label(info_frame, text="", font=("Arial", 12))
        self.potencia_nominal_val.grid(row=3, column=1, sticky="w")

    def configurar_objetivos(self, parent_frame):
        """
        Configura el frame de objetivos.
        """
        objetivo_frame = tk.LabelFrame(parent_frame, text="Objetivos", pady=6, padx=5, font=("Arial", 14))
        objetivo_frame.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        # Crear etiquetas y checkboxes dinámicamente
        self.objetivos = []
        for i in range(1, 14):  # Objetivos del 1 al 13
            tk.Label(objetivo_frame, text=f"Objetivo {i}:", font=("Arial", 12)).grid(row=i - 1, column=0, sticky="e",
                                                                                     pady=5)
            objetivo_var = tk.BooleanVar(value=False)
            objetivo_checkbox = tk.Checkbutton(objetivo_frame, variable=objetivo_var, state="disabled")
            objetivo_checkbox.grid(row=i - 1, column=1, sticky="w")
            self.objetivos.append(objetivo_var)

    def configurar_canvas_columna_2(self):
        """
        Configura el canvas dentro de la columna 2.
        """
        # Frame para el canvas
        canvas_frame = tk.Frame(self.columna_2, pady=5, padx=5)
        canvas_frame.grid(row=0, column=0, padx=5, pady=5, rowspan=7, columnspan=2, sticky="nsew")

        # Agregar el canvas
        self.canvas = tk.Canvas(canvas_frame, bg="lightgray")
        self.canvas.pack(fill="both", expand=True)  # Asegura que el canvas ocupe todo el espacio

    def crear_menus(self):
        """
        Configura los menús principales de la aplicación.
        """
        menu_bar = Menu(self.root)
        self.root.config(menu=menu_bar)

        # Menú Archivo
        archivo_menu = Menu(menu_bar, tearoff=0)
        archivo_menu.add_command(label="Abrir", command=self.abrir_archivo)
        archivo_menu.add_command(label="Guardar", command=self.guardar_archivo)
        archivo_menu.add_command(label="Nuevo", command=self.crear_nuevo_archivo)
        archivo_menu.add_separator()
        archivo_menu.add_command(label="Salir", command=self.root.quit)
        menu_bar.add_cascade(label="Archivo", menu=archivo_menu)

        # Menú Transformador
        proyecto_menu = Menu(menu_bar, tearoff=0)
        proyecto_menu.add_command(label="Datos del Transformador", command=self.abrir_datos_transformador)
        menu_bar.add_cascade(label="Transformador", menu=proyecto_menu)

        # Menú Geometría
        geometria_menu = Menu(menu_bar, tearoff=0)
        geometria_menu.add_command(label="Geometría Definida", command=self.geometry_manager.abrir_ventana_spt)
        menu_bar.add_cascade(label="Geometría", menu=geometria_menu)

        # Menú Resistividad
        resistividad_menu = Menu(menu_bar, tearoff=0)
        resistividad_menu.add_command(label="Datos de Resistividad", command=self.resistivity_manager.abrir_datos_resistencia)
        menu_bar.add_cascade(label="Resistividad", menu=resistividad_menu)

        # Menú Reportes
        reportes_menu = Menu(menu_bar, tearoff=0)
        reportes_menu.add_command(label="Reporte de Transformador", command=self.mostrar_info)
        reportes_menu.add_command(label="Reporte de Resistividad", command=self.mostrar_info)
        reportes_menu.add_command(label="Reporte de Geometria", command=self.mostrar_info)
        menu_bar.add_cascade(label="Reportes", menu=reportes_menu)

        # Menú Ayuda
        ayuda_menu = Menu(menu_bar, tearoff=0)
        ayuda_menu.add_command(label="Acerca de", command=self.mostrar_info)
        menu_bar.add_cascade(label="Ayuda", menu=ayuda_menu)

    def abrir_archivo(self):
        """
        Llama al método de ExcelManager para abrir un archivo y realiza validaciones.
        """
        if self.excel_manager.abrir_archivo():
            self.actualizar_datos_principal()

    def guardar_archivo(self):
        """
        Llama al método de ExcelManager para guardar el archivo actual.
        """
        self.excel_manager.guardar_archivo()

    def crear_nuevo_archivo(self):
        """
        Llama al método de ExcelManager para crear un nuevo archivo y actualiza la interfaz.
        """
        self.excel_manager.crear_nuevo_archivo()
        self.actualizar_datos_principal()

    def abrir_datos_transformador(self):
        """
        Abre la ventana para gestionar datos del transformador.
        """
        self.transformer_manager.abrir_datos_transformador()

    def mostrar_info(self):
        """
        Muestra información sobre la aplicación.
        """
        messagebox.showinfo("Acerca de", "DesingSPT\nVersión 1.0\nCreado por tu equipo.")

    def actualizar_datos_principal(self):
        """
        Actualiza la ventana principal con datos clave del archivo Excel.
        """
        if not self.excel_manager.archivo_cargado():
            print("No hay ningún archivo cargado.")
            return

        try:
            # Título del proyecto
            if "Información" in self.excel_manager.workbook.sheetnames:
                info_sheet = self.excel_manager.workbook["Información"]
                titulo_proyecto = info_sheet.cell(row=2, column=1).value
            else:
                titulo_proyecto = "No disponible"

            self.root.title(f"DesingSPT - {titulo_proyecto}")

        except Exception as e:
            print(f"Error al actualizar datos: {e}")

    def abrir_ventana_geometria(self):
        """
        Abre la ventana de geometría definida utilizando GeometryManager.
        """
        self.geometry_manager.abrir_ventana_geometria(self.root)

    def configurar_frame_resistencia(self, parent_frame):
        """
        Configura el Frame de Resistencia dentro de la columna 3.
        """
        resistencia_frame = tk.LabelFrame(parent_frame, text="Resistencia", pady=5, padx=5, font=("Arial", 14))
        resistencia_frame.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        # Campos dentro del Frame Resistencia
        tk.Label(resistencia_frame, text="Profundidad:", font=("Arial", 12)).grid(row=0, column=0, pady=5, sticky="e")
        profundidad_spinbox = tk.Spinbox(resistencia_frame, from_=0.5, to=1.5, increment=0.1, width=5)
        profundidad_spinbox.grid(row=0, column=1, sticky="w")
        profundidad_spinbox.delete(0, "end")
        profundidad_spinbox.insert(0, "0.5")  # Valor por defecto

        tk.Label(resistencia_frame, text="Resistividad:", font=("Arial", 12)).grid(row=1, column=0, sticky="e", pady=5)
        lista_resistividad = ttk.Combobox(resistencia_frame, state="readonly", width=15)
        lista_resistividad.grid(row=1, column=1, sticky="w")
        lista_resistividad.bind("<<ComboboxSelected>>", self.actualizar_resistividad_seleccionada)

        tk.Label(resistencia_frame, text="Resistencia Malla:", font=("Arial", 12)).grid(row=2, column=0, sticky="e", pady=5)
        tk.Entry(resistencia_frame).grid(row=2, column=1, sticky="w")

    def actualizar_resistividad_seleccionada(self, event=None):
        """
        Actualiza la resistividad seleccionada en el combobox del Frame Resistencia.
        """
        resistividad = event.widget.get()
        messagebox.showinfo("Resistividad Seleccionada", f"Has seleccionado: {resistividad}")

    def configurar_frame_seguridad(self, parent_frame):
        """
        Configura el Frame de Seguridad dentro de la columna 3.
        """
        seguridad_frame = tk.LabelFrame(parent_frame, text="Seguridad", pady=5, padx=5, font=("Arial", 14))
        seguridad_frame.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        # Campos dentro del Frame Seguridad
        tk.Label(seguridad_frame, text="GPR:", font=("Arial", 12)).grid(row=0, column=0, sticky="e")
        tk.Entry(seguridad_frame).grid(row=0, column=1, sticky="w", pady=5)
        tk.Label(seguridad_frame, text="Tensión de paso:", font=("Arial", 12)).grid(row=1, column=0, sticky="e")
        tk.Entry(seguridad_frame).grid(row=1, column=1, sticky="w", pady=5)
        tk.Label(seguridad_frame, text="Tensión de toque:", font=("Arial", 12)).grid(row=2, column=0, sticky="e")
        tk.Entry(seguridad_frame).grid(row=2, column=1, sticky="w", pady=5)

    def configurar_frame_sugerencias(self, parent_frame):
        """
        Configura el Frame de Sugerencias dentro de la columna 3.
        """
        default_frame = tk.LabelFrame(parent_frame, text="Sugerencias", pady=5, padx=5, font=("Arial", 14))
        default_frame.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        # Campo dentro del Frame Default
        tk.Label(default_frame, text="Texto sugerencia:", font=("Arial", 12)).grid(row=0, column=0, sticky="e")
        tk.Label(default_frame, text="Este es un texto no editable.").grid(row=0, column=1, sticky="w", pady=5)



