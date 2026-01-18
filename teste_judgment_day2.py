import time
import pytesseract
from PIL import Image
from pywinauto import Application
import pyautogui

# --- CONFIGURATION ---
PYTESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = PYTESSERACT_PATH

MAIN_TITLE_RE = r"^Protocolos.*" 
CHILD_TITLE_RE = r"^Tela de Controle de Documentos"

def find_and_open_process(target_name_part):
    """
    Scans the grid for a specific name and double-clicks it.
    target_name_part: A unique part of the name (e.g., "CAMPANHOLO")
    """
    print(f"\n--- HUNTING FOR: '{target_name_part}' ---")
    
    try:
        app = Application(backend="win32").connect(title_re=MAIN_TITLE_RE, timeout=10)
        main_win = app.window(title_re=MAIN_TITLE_RE)
        
        # 1. Find Window & Grid
        target_window = main_win.child_window(title_re=CHILD_TITLE_RE)
        if not target_window.exists():
            print("❌ Error: Window 'Tela de Controle de Documentos' not found.")
            return

        target_window.set_focus()
        
        # Target the "Documento(s) Eletrônico(s)" group (Top Grid)
        group_box = target_window.child_window(title="Documento(s) Eletrônico(s)", class_name="TGroupBox")
        grid_data_area = group_box.child_window(class_name="TcxGridSite")
        
        # 2. Capture Screenshot
        print("   -> 📸 Capturing grid image...")
        # Get absolute screen coordinates of the grid (Anchor)
        grid_rect = grid_data_area.rectangle()
        
        img = grid_data_area.capture_as_image()
        
        # 3. Analyze with Tesseract (Get Coordinates)
        print("   -> 🧠 Analyzing text positions...")
        
        # image_to_data returns a dictionary with lists: 'text', 'left', 'top', 'width', 'height'
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        
        found_index = -1
        
        # Iterate through all words found
        n_boxes = len(data['text'])
        for i in range(n_boxes):
            word = data['text'][i].strip()
            
            # Check if this word is part of our target name
            # We use case-insensitive check
            if word and target_name_part.upper() in word.upper():
                found_index = i
                print(f"   -> 🎯 FOUND MATCH: '{word}'")
                break
        
        if found_index != -1:
            # 4. Calculate Absolute Screen Coordinates
            # OCR coordinates are relative to the Image (The Grid)
            ocr_x = data['left'][found_index]
            ocr_y = data['top'][found_index]
            ocr_w = data['width'][found_index]
            ocr_h = data['height'][found_index]
            
            # Absolute X = Grid_Left + OCR_Left + Half_Width
            final_x = grid_rect.left + ocr_x + (ocr_w // 2)
            final_y = grid_rect.top + ocr_y + (ocr_h // 2)
            
            print(f"   -> 📍 Coordinates: ({final_x}, {final_y})")
            
            # 5. Perform Action
            print("   -> 🚀 Moving mouse and clicking...")
            pyautogui.moveTo(final_x, final_y, duration=0.8)
            
            # Verify we are there (Ghost Mode)
            pyautogui.moveRel(5, 0, duration=0.1)
            pyautogui.moveRel(-5, 0, duration=0.1)
            
            # Double Click to open
            pyautogui.doubleClick()
            print("   -> ✅ Double Click Sent.")
            return True
            
        else:
            print(f"   -> ❌ Could not find text '{target_name_part}' in the visible grid.")
            return False

    except Exception as e:
        print(f"❌ Critical Error: {e}")
        return False

# ---------------------------------------------------------
# EXECUTION
# ---------------------------------------------------------
if __name__ == "__main__":
    # REPLACE THIS WITH A NAME VISIBLE ON YOUR SCREEN FOR TESTING
    TEST_NAME = "CAMPANHOLO" 
    
    find_and_open_process(TEST_NAME)