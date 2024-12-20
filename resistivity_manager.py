#Contiene la clase ResistivityManager para manejar datos de resistividad.

import tkinter as tk
from tkinter import ttk, messagebox


class ResistivityManager:
    """
    Clase para gestionar los datos de resistividad.
    """

    def __init__(self, excel_manager, root):
        """
        Inicializa el gestor de resistividad.

        :param excel_manager: Instancia de ExcelManager para interactuar con datos de resistividad.
        :param root: Ventana principal de la aplicación.
        """
        self.excel_manager = excel_manager
        self.root = root

    def abrir_datos_resistencia(self):
        """
        Abre una ventana para gestionar datos de resistividad.
        """
        ventana_resistividad = tk.Toplevel(self.root)
        ventana_resistividad.title("Gestión de Resistividad")
        ventana_resistividad.geometry("700x450")

        # Canvas para dibujo de perfiles
        canvas = tk.Canvas(ventana_resistividad, width=150, height=150, bg="white")
        canvas.grid(row=0, column=0, columnspan=5, pady=10)

        # Tabla para mostrar los datos
        tabla = ttk.Treeview(ventana_resistividad, columns=("Perfil", "1m", "2m", "3m", "4m"), show="headings", height=6)
        tabla.grid(row=0, column=6, rowspan=6, padx=10, pady=5)

        # Configurar encabezados de la tabla
        tabla.heading("Perfil", text="Perfil")
        tabla.heading("1m", text="1m")
        tabla.heading("2m", text="2m")
        tabla.heading("3m", text="3m")
        tabla.heading("4m", text="4m")

        # Ajustar ancho de columnas
        for col in ("Perfil", "1m", "2m", "3m", "4m"):
            tabla.column(col, width=80, anchor="center")

        # Botones
        tk.Button(ventana_resistividad, text="Añadir Datos", command=lambda: self.anadir_datos(tabla)).grid(row=1, column=0, pady=5)
        tk.Button(ventana_resistividad, text="Guardar Datos", command=lambda: self.guardar_datos_resistencia(tabla)).grid(row=2, column=0, pady=5)
        tk.Button(ventana_resistividad, text="Eliminar Fila", command=lambda: self.eliminar_fila(tabla)).grid(row=3, column=0, pady=5)

        # Cargar datos existentes en la tabla
        self.cargar_datos_existentes(tabla)

    def cargar_datos_existentes(self, tabla):
        """
        Carga los datos existentes de resistividad desde el archivo Excel a la tabla.
        """
        if not self.excel_manager.archivo_cargado():
            messagebox.showwarning("Advertencia", "No hay ningún archivo abierto para cargar datos.")
            return

        try:
            if "Resistencias" not in self.excel_manager.workbook.sheetnames:
                messagebox.showinfo("Información", "No se encontró la hoja 'Resistencias' en el archivo.")
                return

            # Acceder a la hoja "Resistencias"
            resistencias_sheet = self.excel_manager.workbook["Resistencias"]

            # Limpiar la tabla antes de cargar nuevos datos
            for item in tabla.get_children():
                tabla.delete(item)

            # Leer los datos de la hoja (a partir de la segunda fila)
            for row in resistencias_sheet.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):  # Solo añadir filas con datos
                    tabla.insert("", "end", values=row)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los datos: {e}")

    def anadir_datos(self, tabla):
        """
        Añade una nueva fila de datos a la tabla.
        """
        # Crear una ventana emergente para ingresar datos
        ventana_datos = tk.Toplevel(self.root)
        ventana_datos.title("Añadir Datos de Resistividad")
        ventana_datos.geometry("300x300")

        # Entradas para los valores
        entradas = {}
        for i, label in enumerate(["Perfil", "1m", "2m", "3m", "4m"]):
            tk.Label(ventana_datos, text=label).grid(row=i, column=0, padx=10, pady=5)
            entry = tk.Entry(ventana_datos)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entradas[label] = entry

        def guardar_datos():
            # Leer valores de las entradas
            fila = [entradas[label].get() for label in ["Perfil", "1m", "2m", "3m", "4m"]]
            if not fila[0]:  # Verificar que el perfil no esté vacío
                messagebox.showerror("Error", "El campo 'Perfil' no puede estar vacío.")
                return
            tabla.insert("", "end", values=fila)
            ventana_datos.destroy()

        tk.Button(ventana_datos, text="Guardar", command=guardar_datos).grid(row=5, column=0, columnspan=2, pady=10)

    def eliminar_fila(self, tabla):
        """
        Elimina la fila seleccionada de la tabla.
        """
        seleccion = tabla.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione una fila para eliminar.")
            return

        for item in seleccion:
            tabla.delete(item)

    def guardar_datos_resistencia(self, tabla):
        """
        Guarda los datos de la tabla en la hoja "Resistencias" del archivo Excel.
        """
        if not self.excel_manager.archivo_cargado():
            messagebox.showwarning("Advertencia", "No hay ningún archivo abierto para guardar datos.")
            return

        try:
            if "Resistencias" not in self.excel_manager.workbook.sheetnames:
                resistencias_sheet = self.excel_manager.workbook.create_sheet("Resistencias")
                encabezados = ["Perfil", "1m", "2m", "3m", "4m"]
                for col, encabezado in enumerate(encabezados, start=1):
                    resistencias_sheet.cell(row=1, column=col, value=encabezado)
            else:
                resistencias_sheet = self.excel_manager.workbook["Resistencias"]

            # Limpiar la hoja antes de guardar
            resistencias_sheet.delete_rows(2, resistencias_sheet.max_row)

            # Guardar datos de la tabla en la hoja Excel
            for i, item in enumerate(tabla.get_children(), start=2):
                valores = tabla.item(item)["values"]
                for j, valor in enumerate(valores, start=1):
                    resistencias_sheet.cell(row=i, column=j, value=valor)

            self.excel_manager.guardar_archivo()
            messagebox.showinfo("Éxito", "Datos guardados correctamente.")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar los datos: {e}")
