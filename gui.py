# --- FILE: gui.py ---

import sys
import logging
import traceback
import os
import multiprocessing
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont, QIcon
from app.video_splash import VideoSplashScreen

# Import the main window from its new module
from app.main_window import AuditApp

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        # This is for --onefile builds
        base_path = sys._MEIPASS
    except Exception:
        # --onedir builds (like yours) and for development
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# --- Centralized Logging Setup ---
def setup_logging():
    """Configures logging to write to a file and the console."""
    log_formatter = logging.Formatter(
        '%(asctime)s [%(threadName)-12.12s] [%(levelname)-5.5s] [%(name)s] %(message)s'
    )
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # File handler
    try:
        file_handler = logging.FileHandler("debug_log.txt", mode='w', encoding='utf-8')
        file_handler.setFormatter(log_formatter)
        root_logger.addHandler(file_handler)
    except Exception as e:
        print(f"Could not set up file logger: {e}")

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter)
    root_logger.addHandler(console_handler)

    logging.info("Logging configured.")

def handle_exception(exctype, value, tb):
    """Global exception hook to log any uncaught exception."""
    logging.critical("--- UNCAUGHT EXCEPTION ---", exc_info=(exctype, value, tb))
    # Optionally, show a message box to the user
    # from PySide6.QtWidgets import QMessageBox
    # QMessageBox.critical(None, "Erro Crítico", f"Ocorreu um erro inesperado e a aplicação pode precisar de ser fechada.\n\nDetalhes: {value}\n\nConsulte o debug_log.txt para mais informações.")
    sys.__excepthook__(exctype, value, tb)

# --- Stylesheet (can be moved to a file later if it grows) ---
APP_STYLESHEET = """
    /* General Window & Text */
    QWidget {
        background-color: #2E3440; /* Nord Polar Night */
        color: #D8DEE9;            /* Nord Snow Storm */
        font-family: "Segoe UI";
        font-size: 10pt;
    }

    /* GroupBox - Fixes the sizing and title */
    QGroupBox {
        background-color: #3B4252;
        border: 1px solid #4C566A;
        border-radius: 6px;
        padding-top: 25px; /* Crucial: Creates space inside for the title */
        margin-top: 10px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 5px 10px;
        left: 10px;
        font-weight: bold;
        color: #ECEFF4;
        background-color: #434C5E;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
    }

    /* Labels - Ensure they have transparent background */
    QLabel {
        color: #D8DEE9;
        background-color: transparent;
    }

    /* Buttons */
    QPushButton {
        background-color: #4C566A;
        color: #ECEFF4;
        border: 1px solid #5E81AC;
        padding: 8px 16px;
        border-radius: 4px;
    }
    QPushButton:hover {
        background-color: #5E81AC; /* Nord Frost */
    }
    QPushButton:pressed {
        background-color: #81A1C1;
    }
    QPushButton:disabled {
        background-color: #3B4252;
        color: #4C566A;
        border-color: #3B4252;
    }

    /* Input Fields */
    QLineEdit, QTextEdit, QComboBox {
        background-color: #3B4252;
        color: #ECEFF4;
        border: 1px solid #4C566A;
        border-radius: 4px;
        padding: 5px;
    }
    QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
        border: 1px solid #88C0D0; /* Nord Frost */
    }
    QLineEdit:read-only {
        background-color: #434C5E;
    }
    
    QComboBox::drop-down {
        border: none;
        width: 20px;
    }

    /* Table Styling */
    QTableWidget {
        gridline-color: #4C566A;
        background-color: #3B4252;
        alternate-background-color: #434C5E;
        selection-background-color: #5E81AC;
        selection-color: #ECEFF4;
        border: 1px solid #4C566A;
    }
    QHeaderView::section {
        background-color: #434C5E;
        color: #ECEFF4;
        padding: 6px;
        border: none;
        border-right: 1px solid #4C566A;
        font-weight: bold;
    }
    QTableWidget::item {
        padding: 5px;
        border: none; /* Removes inner cell borders */
    }
    
    /* ScrollBar Styling */
    QScrollBar:vertical {
        border: none;
        background: #3B4252;
        width: 12px;
        margin: 0px 0px 0px 0px;
    }
    QScrollBar::handle:vertical {
        background: #4C566A;
        min-height: 20px;
        border-radius: 6px;
    }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
        height: 0px;
    }
    QScrollBar:horizontal {
        border: none;
        background: #3B4252;
        height: 12px;
        margin: 0px 0px 0px 0px;
    }
    QScrollBar::handle:horizontal {
        background: #4C566A;
        min-width: 20px;
        border-radius: 6px;
    }
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
        width: 0px;
    }
"""

if __name__ == "__main__":
    multiprocessing.freeze_support()
    setup_logging()
    sys.excepthook = handle_exception

    app = QApplication(sys.argv)
    
    # ... fonts and stylesheets setup ...
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    app.setStyleSheet(APP_STYLESHEET)
    icon_path = resource_path("favicon.ico")
    app.setWindowIcon(QIcon(icon_path))
    
    # --- SPLASH SCREEN LOGIC ---
    
    # 1. Define path to your video (e.g., "assets/intro.mp4")
    video_path = resource_path("assets/intro2.mp4") 
    
    # 2. Check if we should play video or skip (e.g. argument --no-splash)
    # For now, we assume we always play it.
    
    # 3. Create the Main Window (But DO NOT show it yet)
    # We create it now so it loads in the background while video plays
    main_window = AuditApp()

    def start_main_app():
        """Slot called when video finishes."""
        main_window.show()

    # 4. Initialize Splash
    if os.path.exists(video_path):
        print("Found video, initializing splash...")
        splash = VideoSplashScreen(video_path, width=800, height=500)
        splash.finished.connect(start_main_app)
        splash.start()
    else:
        print("Video not found, skipping.")
        main_window.show()

    sys.exit(app.exec())