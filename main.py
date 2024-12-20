#Punto de entrada del programa.
#Inicializa la interfaz principal y dependencias clave.

from ui_manager import UIManager
from excel_manager import ExcelManager

def main():
    """
    Punto de entrada principal de la aplicación.
    Inicializa la interfaz gráfica y las dependencias clave.
    """
    # Crear una instancia del gestor de archivos Excel
    excel_manager = ExcelManager()

    # Crear una instancia del gestor de la interfaz gráfica, pasándole el gestor de Excel
    ui_manager = UIManager(excel_manager)

    # Iniciar la interfaz gráfica
    ui_manager.iniciar_interfaz()

if __name__ == "__main__":
    main()
