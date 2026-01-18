# Add this helper function to a central place, like the top of main.py
import sys
import os

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # ✅ --- START: THIS IS THE FIX ---
        # Usa o caminho do *arquivo de script*, não o "current working directory"
        # Isso garante que ele encontre os arquivos, não importa de onde você rode o python.
        base_path = os.path.abspath(os.path.dirname(__file__))
        # ❌ OLD:
        # base_path = os.path.abspath(".")
        # ✅ --- END: THIS IS THE FIX ---

    return os.path.join(base_path, relative_path)
