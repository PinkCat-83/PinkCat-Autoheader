"""
Autopdf - Convertidor DOCX a PDF con Encabezado y Pie de Página
Punto de entrada principal de la aplicación
"""

import tkinter as tk
from src.gui import GUI
from src.controller import AppController

try:
    from tkinterdnd2 import TkinterDnD
    DRAG_DROP_DISPONIBLE = True
except ImportError:
    DRAG_DROP_DISPONIBLE = False
    print("tkinterdnd2 no disponible - drag & drop deshabilitado")


def main():
    """Inicializa y ejecuta la aplicación"""
    # Crear ventana principal (con o sin drag & drop según disponibilidad)
    if DRAG_DROP_DISPONIBLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    # Crear controlador (primero, sin GUI)
    controller = AppController()
    
    # Crear GUI y pasarle el controlador
    gui = GUI(root, controller)
    
    # Enlazar GUI con el controlador
    controller.set_gui(gui)
    
    # Iniciar loop principal
    root.mainloop()


if __name__ == "__main__":
    main()