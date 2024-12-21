#Contiene la clase GeometryManager para los cálculos geométricos y la manipulación de figuras.

import math
import tkinter as tk
from tkinter import messagebox, ttk


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
        Abre la ventana de "Geometría Definida" para seleccionar y dibujar geometrías.
        Si existen datos en la hoja "Área" del archivo Excel, los muestra en el canvas.
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
        geometry_combo = ttk.Combobox(
            control_frame, textvariable=geometry_var,
            values=["L", "Rectángulo", "Línea", "Triángulo", "Circunferencia"]
        )
        geometry_combo.grid(row=0, column=1, padx=5)

        # Botones
        tk.Button(control_frame, text="Editar", command=self.editar_geometria).grid(row=0, column=2, padx=5)
        tk.Button(control_frame, text="Guardar", command=self.guardar_geometria).grid(row=0, column=3, padx=5)

        # Dibujar figura inicial (desde Excel o predeterminada)
        self.cargar_figura_desde_excel()
        if not self.figura_desde_excel:
            self.geometry_type = "L"
            self.points = self.obtener_dimensiones_predeterminadas(self.geometry_type)
            self.dibujar_figura(self.geometry_type, self.points)

        # Vincular evento al combobox
        geometry_combo.bind("<<ComboboxSelected>>", lambda e: self.cambiar_geometria(geometry_var.get()))

        # Mostrar longitudes y perímetro
        dimensiones_frame = tk.Frame(ventana_geometria)
        dimensiones_frame.pack(pady=5)
        self.mostrar_dimensiones_y_perimetro(dimensiones_frame)

        # Dibujar automáticamente si se cambia la geometría
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
        global points
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
        Guarda los datos de la geometría actual en la hoja "Área" del archivo Excel.
        """
        if not self.excel_manager.archivo_cargado():
            messagebox.showwarning("Advertencia", "No hay un archivo abierto para guardar los datos.")
            return

        try:
            # Acceder o crear la hoja "Área"
            if "Área" not in self.excel_manager.workbook.sheetnames:
                area_sheet = self.excel_manager.workbook.create_sheet("Área")
                # Crear encabezados
                encabezados = ["Tipo de Geometría", "Puntos", "Longitudes de Segmentos", "Perímetro", "Área",
                               "Radio Equivalente"]
                for col, encabezado in enumerate(encabezados, start=1):
                    area_sheet.cell(row=1, column=col, value=encabezado)
            else:
                area_sheet = self.excel_manager.workbook["Área"]

            # Calcular las longitudes, perímetro y área
            longitudes, perimetro = self.calcular_longitudes_y_perimetro(self.points, self.geometry_type)
            area = self.calcular_area(self.geometry_type, longitudes)
            radio_equivalente = round(math.sqrt(area / math.pi), 4)

            # Escribir datos en la segunda fila de la hoja
            area_sheet.cell(row=2, column=1, value=self.geometry_type)  # Tipo de geometría
            area_sheet.cell(row=2, column=2, value=str(self.points))  # Puntos
            area_sheet.cell(row=2, column=3, value=str(longitudes))  # Longitudes de segmentos
            area_sheet.cell(row=2, column=4, value=round(perimetro, 2))  # Perímetro
            area_sheet.cell(row=2, column=5, value=round(area, 4))  # Área
            area_sheet.cell(row=2, column=6, value=radio_equivalente)  # Radio equivalente

            # Guardar el archivo
            self.excel_manager.guardar_archivo()
            messagebox.showinfo("Éxito", "Los datos de la geometría han sido guardados correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la geometría: {e}")

    def calcular_area(self, tipo_geometria, longitudes):
        """
        Calcula el área de la figura según su tipo.

        :param tipo_geometria: Tipo de geometría ('Rectángulo', 'L', etc.).
        :param longitudes: Lista de longitudes de los segmentos.
        :return: Área calculada.
        """
        if tipo_geometria == "Rectángulo":
            # Área = Base × Altura
            base, altura = longitudes[:2]
            return base * altura / 10000
        elif tipo_geometria == "Línea":
            # El área de una línea es igual a 0
            return 0
        elif tipo_geometria == "L":
            # Área aproximada = suma de los segmentos
            return sum(longitudes[:2]) / 100
        else:
            return 0

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

    def mostrar_dimensiones_y_perimetro(self, parent_frame):
        """
        Muestra las dimensiones de los segmentos y el perímetro de la figura actual.
        """
        tk.Label(parent_frame, text="Longitudes de segmentos:", font=("Arial", 10)).grid(row=0, column=0, sticky="w",
                                                                                         padx=5)
        self.longitudes_label = tk.Label(parent_frame, text="-", font=("Arial", 10))
        self.longitudes_label.grid(row=0, column=1, sticky="w", padx=5)

        tk.Label(parent_frame, text="Perímetro:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=5)
        self.perimetro_label = tk.Label(parent_frame, text="-", font=("Arial", 10))
        self.perimetro_label.grid(row=1, column=1, sticky="w", padx=5)

        # Actualizar las longitudes y el perímetro con los datos actuales
        self.actualizar_dimensiones_y_perimetro()

    def actualizar_dimensiones_y_perimetro(self):
        """
        Calcula y actualiza las longitudes de los segmentos y el perímetro.
        """
        if not self.points:
            self.longitudes_label.config(text="-")
            self.perimetro_label.config(text="-")
            return

        longitudes, perimetro = self.calcular_longitudes_y_perimetro(self.points, self.geometry_type)
        self.longitudes_label.config(text=", ".join(map(str, longitudes)))
        self.perimetro_label.config(text=f"{perimetro:.2f}")

    def calcular_longitudes_y_perimetro(self, puntos, tipo_geometria):
        """
        Calcula las longitudes de los segmentos y el perímetro de la figura según su tipo.

        :param puntos: Lista de puntos [(x1, y1), (x2, y2), ...] que definen la figura.
        :param tipo_geometria: Tipo de geometría ('L', 'Rectángulo', 'Triángulo', etc.).
        :return: Una lista de longitudes y el perímetro total.
        """
        longitudes = []
        perimetro = 0

        for i in range(len(puntos) - (1 if tipo_geometria in ["L", "Línea"] else 0)):
            x1, y1 = puntos[i]
            x2, y2 = puntos[(i + 1) % len(puntos)]  # Conexión cíclica para figuras cerradas
            distancia = round(math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2), 2)
            longitudes.append(distancia)
            perimetro += distancia

        return longitudes, perimetro

    def cargar_figura_desde_excel(self):
        """
        Carga datos de geometría desde la hoja "Área" del archivo Excel si está disponible.
        """
        if not self.excel_manager.archivo_cargado():
            return

        try:
            # Accede a la hoja "Área" en el archivo Excel
            if "Área" in self.excel_manager.workbook.sheetnames:
                area_sheet = self.excel_manager.workbook["Área"]

                # Leer datos de la primera fila de la hoja
                tipo_geometria = area_sheet.cell(row=2, column=1).value
                puntos_str = area_sheet.cell(row=2, column=2).value

                if tipo_geometria and puntos_str:
                    # Convertir los puntos desde el string al formato adecuado
                    puntos = eval(puntos_str)
                    self.validar_puntos(puntos)
                    self.geometry_type = tipo_geometria
                    self.points = puntos

                    # Dibuja la figura cargada
                    self.dibujar_figura(tipo_geometria, puntos)
                    self.figura_desde_excel = True
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los datos del área: {e}")

    def validar_puntos(self, puntos):
        """
        Valida que los puntos tengan el formato adecuado.

        :param puntos: Lista de puntos a validar.
        :raises ValueError: Si los puntos no tienen el formato esperado.
        """
        if not isinstance(puntos, list) or not all(isinstance(p, tuple) and len(p) == 2 for p in puntos):
            raise ValueError("Los puntos no tienen el formato esperado. Deben ser una lista de tuplas (x, y).")
