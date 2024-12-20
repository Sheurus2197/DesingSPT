#Funciones auxiliares y genéricas.

import tkinter as tk
from tkinter import messagebox


def validar_numero(valor):
    """
    Verifica si un valor es numérico.

    :param valor: Valor a validar.
    :return: True si el valor es un número válido (entero o flotante), False en caso contrario.
    """
    try:
        float(valor)
        return True
    except ValueError:
        return False


def mostrar_error(mensaje):
    """
    Muestra un cuadro de diálogo con un mensaje de error.

    :param mensaje: Mensaje a mostrar.
    """
    messagebox.showerror("Error", mensaje)


def mostrar_info(titulo, mensaje):
    """
    Muestra un cuadro de diálogo con un mensaje informativo.

    :param titulo: Título de la ventana.
    :param mensaje: Mensaje a mostrar.
    """
    messagebox.showinfo(titulo, mensaje)


def normalizar_texto(texto):
    """
    Normaliza una cadena de texto eliminando espacios innecesarios.

    :param texto: Cadena de texto a normalizar.
    :return: Texto normalizado.
    """
    return texto.strip()


def configurar_ventana(root, titulo, ancho, alto):
    """
    Configura la ventana principal con un tamaño específico y centrado en la pantalla.

    :param root: Ventana principal (objeto Tkinter).
    :param titulo: Título de la ventana.
    :param ancho: Ancho de la ventana.
    :param alto: Alto de la ventana.
    """
    root.title(titulo)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    pos_x = (screen_width // 2) - (ancho // 2)
    pos_y = (screen_height // 2) - (alto // 2)
    root.geometry(f"{ancho}x{alto}+{pos_x}+{pos_y}")


def manejar_excepcion(exception, contexto=""):
    """
    Maneja una excepción, mostrando un mensaje en consola y un cuadro de error.

    :param exception: Objeto de excepción capturado.
    :param contexto: Contexto o descripción adicional del error.
    """
    print(f"Error {contexto}: {exception}")
    mostrar_error(f"Error {contexto}: {exception}")
