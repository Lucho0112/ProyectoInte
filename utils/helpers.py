import os
import sys

def resource_path(relative_path):
    """
    Obtiene la ruta absoluta para el recurso.
    Funciona tanto en desarrollo como en el ejecutable de PyInstaller.
    """
    try:
        # 1. Cuando se ejecuta como ejecutable PyInstaller (sys._MEIPASS existe)
        base_path = sys._MEIPASS
    except Exception:
        # 2. Cuando se ejecuta como script Python normal
        # Asegura que el path sea relativo a la raíz del proyecto (tribunsys/)
        base_path = os.path.abspath(os.path.dirname(__file__) + '/..') 
        
    # Combina la base con la ruta relativa del archivo en el ejecutable
    return os.path.join(base_path, relative_path)