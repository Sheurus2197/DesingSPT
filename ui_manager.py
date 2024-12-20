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
        self.root.title("DesingSPT")
        self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight() - 80}")
        self.root.iconbitmap('icon.ico')
        self.crear_menus()
        self.root.mainloop()

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

        # Menú Proyecto
        proyecto_menu = Menu(menu_bar, tearoff=0)
        proyecto_menu.add_command(label="Datos del Proyecto", command=self.abrir_datos_proyecto)
        proyecto_menu.add_command(label="Datos del Transformador", command=self.abrir_datos_transformador)
        menu_bar.add_cascade(label="Proyecto", menu=proyecto_menu)

        # Menú Geometría
        geometria_menu = Menu(menu_bar, tearoff=0)
        geometria_menu.add_command(label="Geometría Definida", command=self.geometry_manager.abrir_ventana_spt)
        menu_bar.add_cascade(label="Geometría", menu=geometria_menu)

        # Menú Resistividad
        resistividad_menu = Menu(menu_bar, tearoff=0)
        resistividad_menu.add_command(label="Datos de Resistividad", command=self.resistivity_manager.abrir_datos_resistencia)
        menu_bar.add_cascade(label="Resistividad", menu=resistividad_menu)

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

    def abrir_datos_proyecto(self):
        """
        Abre la ventana para gestionar datos del proyecto.
        """
        # Esta funcionalidad ya está implementada en el archivo original
        messagebox.showinfo("Info", "Funcionalidad: Datos del Proyecto (implementada en transformer_manager.py)")

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
