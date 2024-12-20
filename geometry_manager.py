#Contiene la clase GeometryManager para los cálculos geométricos y la manipulación de figuras.

import math
import tkinter as tk
from tkinter import messagebox


class GeometryManager:
    """
    Clase para gestionar los cálculos geométricos y la representación en canvas.
    """

    def __init__(self, excel_manager):
        """
        Inicializa el gestor de geometría.

        :param excel_manager: Instancia de ExcelManager para interactuar con datos geométricos en Excel.
        """
        self.excel_manager = excel_manager
        self.canvas = None
        self.geometry_type = None
        self.points = []

    def abrir_ventana_spt(self):
        """
        Abre la ventana para gestionar geometrías.
        """
        ventana_geometria = tk.Toplevel()
        ventana_geometria.title("Geometría Definida")
        ventana_geometria.geometry("500x500")

        # Canvas para dibujo
        self.canvas = tk.Canvas(ventana_geometria, width=400, height=400, bg="white")
        self.canvas.pack(pady=10)

        # Frame para controles
        control_frame = tk.Frame(ventana_geometria)
        control_frame.pack(pady=10)

        # Listbox para seleccionar geometría
        tk.Label(control_frame, text="Geometría").grid(row=0, column=0, padx=5)
        geometry_var = tk.StringVar(value="L")
        geometry_combo = tk.ttk.Combobox(
            control_frame, textvariable=geometry_var,
            values=["L", "Rectángulo", "Línea", "Triángulo", "Circunferencia"]
        )
        geometry_combo.grid(row=0, column=1, padx=5)

        # Botones
        tk.Button(control_frame, text="Editar", command=lambda: self.editar_geometria()).grid(row=0, column=2, padx=5)
        tk.Button(control_frame, text="Guardar", command=lambda: self.guardar_geometria()).grid(row=0, column=3, padx=5)

        # Dibujar figura inicial
        self.geometry_type = "L"
        self.points = self.obtener_dimensiones_predeterminadas(self.geometry_type)
        self.dibujar_figura(self.geometry_type, self.points)

        # Vincular evento al combobox
        geometry_combo.bind("<<ComboboxSelected>>", lambda e: self.cambiar_geometria(geometry_var.get()))

    def cambiar_geometria(self, nueva_geometria):
        """
        Cambia la geometría actual y redibuja en el canvas.
        """
        self.geometry_type = nueva_geometria
        self.points = self.obtener_dimensiones_predeterminadas(nueva_geometria)
        self.dibujar_figura(nueva_geometria, self.points)

    def obtener_dimensiones_predeterminadas(self, tipo):
        """
        Obtiene dimensiones predeterminadas para un tipo de geometría.

        :param tipo: Tipo de geometría.
        :return: Lista de dimensiones predeterminadas.
        """
        dimensiones = {
            "L": [150, 50],
            "Rectángulo": [100, 150],
            "Triángulo": [100, 100, 100],
            "Línea": [150],
            "Circunferencia": [100],
        }
        return dimensiones.get(tipo, [])

    def dibujar_figura(self, tipo, dimensiones):
        """
        Dibuja una figura en el canvas según su tipo y dimensiones.

        :param tipo: Tipo de figura.
        :param dimensiones: Dimensiones necesarias para la figura.
        """
        self.canvas.delete("all")  # Limpiar canvas

        if tipo == "L":
            vertical, horizontal = dimensiones
            points = [(50, 50), (50, 50 + vertical), (50 + horizontal, 50 + vertical)]
        elif tipo == "Rectángulo":
            base, altura = dimensiones
            points = [(50, 50), (50 + base, 50), (50 + base, 50 + altura), (50, 50 + altura)]
        elif tipo == "Triángulo":
            lado1, lado2, lado3 = dimensiones
            points = [(50, 50), (50 + lado1, 50), (50, 50 + lado2)]
        elif tipo == "Línea":
            longitud = dimensiones[0]
            points = [(50, 100), (50 + longitud, 100)]
        elif tipo == "Circunferencia":
            diametro = dimensiones[0]
            self.dibujar_circunferencia(diametro)
            return

        # Dibujar los segmentos
        for i, (x1, y1) in enumerate(points):
            x2, y2 = points[(i + 1) % len(points)]
            self.canvas.create_line(x1, y1, x2, y2, fill="blue", width=2)

    def dibujar_circunferencia(self, diametro):
        """
        Dibuja una circunferencia en el canvas.

        :param diametro: Diámetro de la circunferencia.
        """
        radius = diametro / 2
        center_x, center_y = 200, 200
        self.canvas.create_oval(
            center_x - radius, center_y - radius,
            center_x + radius, center_y + radius,
            outline="blue", width=2
        )
        self.canvas.create_line(
            center_x - radius, center_y,
            center_x + radius, center_y,
            fill="red", width=2
        )

    def guardar_geometria(self):
        """
        Guarda los datos de la geometría actual en el archivo Excel.
        """
        if not self.excel_manager.archivo_cargado():
            messagebox.showwarning("Advertencia", "No hay ningún archivo abierto para guardar.")
            return

        try:
            if "Área" not in self.excel_manager.workbook.sheetnames:
                area_sheet = self.excel_manager.workbook.create_sheet("Área")
                encabezados = ["Tipo de Geometría", "Dimensiones"]
                for col, encabezado in enumerate(encabezados, start=1):
                    area_sheet.cell(row=1, column=col, value=encabezado)
            else:
                area_sheet = self.excel_manager.workbook["Área"]

            fila = 2
            area_sheet.cell(row=fila, column=1, value=self.geometry_type)
            area_sheet.cell(row=fila, column=2, value=str(self.points))

            self.excel_manager.guardar_archivo()
            messagebox.showinfo("Éxito", "Geometría guardada correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la geometría: {e}")

    def editar_geometria(self):
        """
        Abre una ventana para editar las dimensiones de la figura actual.
        """
        ventana_editar = tk.Toplevel()
        ventana_editar.title(f"Editar {self.geometry_type}")
        dimensiones_actuales = self.points

        entries = []
        for i, dim in enumerate(dimensiones_actuales):
            tk.Label(ventana_editar, text=f"Dimensión {i + 1}:").grid(row=i, column=0)
            entry = tk.Entry(ventana_editar)
            entry.insert(0, dim)
            entry.grid(row=i, column=1)
            entries.append(entry)

        def guardar_cambios():
            try:
                nuevas_dimensiones = [float(entry.get()) for entry in entries]
                self.points = nuevas_dimensiones
                self.dibujar_figura(self.geometry_type, self.points)
                ventana_editar.destroy()
                messagebox.showinfo("Éxito", "Dimensiones actualizadas.")
            except ValueError:
                messagebox.showerror("Error", "Ingrese valores numéricos válidos.")

        tk.Button(ventana_editar, text="Guardar", command=guardar_cambios).grid(row=len(entries), column=0, columnspan=2)
