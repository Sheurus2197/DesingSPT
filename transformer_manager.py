#Contiene la clase TransformerManager para datos del transformador.

import tkinter as tk
from tkinter import messagebox


class TransformerManager:
    """
    Clase para gestionar los datos del transformador.
    """

    def __init__(self, excel_manager):
        """
        Inicializa el gestor de transformadores.

        :param excel_manager: Instancia de ExcelManager para interactuar con los datos del transformador.
        """
        self.excel_manager = excel_manager

    def abrir_datos_transformador(self):
        """
        Abre una ventana para gestionar los datos del transformador.
        """
        ventana_transformador = tk.Toplevel()
        ventana_transformador.title("Datos del Transformador")
        ventana_transformador.geometry("400x300")

        # Etiquetas y entradas para los datos del transformador
        etiquetas = [
            "Tensión Primaria (V):", "Tensión Secundaria (V):",
            "Potencia Nominal (kVA):", "Frecuencia (Hz):", "Tipo de Transformador:"
        ]
        entradas = {}

        for i, etiqueta in enumerate(etiquetas):
            tk.Label(ventana_transformador, text=etiqueta).grid(row=i, column=0, padx=10, pady=5, sticky="e")
            entrada = tk.Entry(ventana_transformador, width=25)
            entrada.grid(row=i, column=1, padx=10, pady=5)
            entradas[etiqueta] = entrada

        # Botón para guardar los datos
        tk.Button(
            ventana_transformador, text="Guardar", command=lambda: self.guardar_datos_transformador(entradas)
        ).grid(row=len(etiquetas), column=0, columnspan=2, pady=10)

        # Cargar datos existentes
        self.cargar_datos_transformador(entradas)

    def cargar_datos_transformador(self, entradas):
        """
        Carga los datos del transformador desde el archivo Excel a las entradas.
        """
        if not self.excel_manager.archivo_cargado():
            messagebox.showwarning("Advertencia", "No hay ningún archivo abierto para cargar datos.")
            return

        try:
            if "Transformador" not in self.excel_manager.workbook.sheetnames:
                messagebox.showinfo("Información", "No se encontró la hoja 'Transformador' en el archivo.")
                return

            transformador_sheet = self.excel_manager.workbook["Transformador"]

            # Leer datos desde las celdas y cargar en las entradas
            for i, (etiqueta, entrada) in enumerate(entradas.items(), start=1):
                entrada.insert(0, transformador_sheet.cell(row=i, column=2).value or "")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los datos: {e}")

    def guardar_datos_transformador(self, entradas):
        """
        Guarda los datos del transformador desde las entradas al archivo Excel.
        """
        if not self.excel_manager.archivo_cargado():
            messagebox.showwarning("Advertencia", "No hay ningún archivo abierto para guardar datos.")
            return

        try:
            # Crear la hoja si no existe
            if "Transformador" not in self.excel_manager.workbook.sheetnames:
                transformador_sheet = self.excel_manager.workbook.create_sheet("Transformador")
                encabezados = ["Parámetro", "Valor"]
                for col, encabezado in enumerate(encabezados, start=1):
                    transformador_sheet.cell(row=1, column=col, value=encabezado)
            else:
                transformador_sheet = self.excel_manager.workbook["Transformador"]

            # Guardar datos desde las entradas
            for i, (etiqueta, entrada) in enumerate(entradas.items(), start=1):
                transformador_sheet.cell(row=i + 1, column=1, value=etiqueta.rstrip(":"))
                transformador_sheet.cell(row=i + 1, column=2, value=entrada.get())

            self.excel_manager.guardar_archivo()
            messagebox.showinfo("Éxito", "Datos del transformador guardados correctamente.")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron guardar los datos: {e}")
