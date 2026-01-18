import time
import sys
import re
import os
import pyperclip
import pyautogui
import pytesseract
from PIL import Image, ImageOps, ImageEnhance
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

# --- CONFIGURATION ---
PYTESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
TITLE_RE = r"^Comércio.*"
pytesseract.pytesseract.tesseract_cmd = PYTESSERACT_PATH

# ---------------------------------------------------------
# HELPER: VISUAL OCR CLICK (Updated with X and Y Offset)
# ---------------------------------------------------------
def click_text_via_ocr(window_obj, text_to_find, exact_match=False, x_offset=0, y_offset=0):
    """
    Captures window, finds text, and clicks relative to it.
    x_offset: Shift click Left/Right
    y_offset: Shift click Up/Down (Positive = Down)
    """
    print(f"      📷 Visual Scan: Looking for text '{text_to_find}'...")
    try:
        img = window_obj.capture_as_image()
        # psm 11 is 'Sparse Text', good for UI labels
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT, config='--psm 11')
        
        n_boxes = len(data['text'])
        target_found = False
        
        for i in range(n_boxes):
            detected_text = data['text'][i].strip()
            
            match = False
            if exact_match:
                match = (detected_text.lower() == text_to_find.lower())
            else:
                match = (text_to_find.lower() in detected_text.lower())
            
            if match and int(data['conf'][i]) > 40:
                x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
                
                # Calculate coordinates (Center of text)
                click_x = x + (w // 2)
                click_y = y + (h // 2)
                
                # Apply Offsets
                final_x = click_x + x_offset
                final_y = click_y + y_offset
                
                print(f"      ✅ Found '{detected_text}'. Clicking at Offset: ({x_offset}, {y_offset}).")
                
                window_obj.click_input(coords=(final_x, final_y))
                target_found = True
                break
        
        if not target_found:
            print(f"      ❌ Could not visually find text '{text_to_find}'.")
            return False
        return True

    except Exception as e:
        print(f"      ❌ Visual Click Error: {e}")
        return False

# ---------------------------------------------------------
# HELPER: SAVE PDF ROUTINE (Updated for Multiple Window Titles)
# ---------------------------------------------------------
def save_report_pdf(app, filename):
    """
    Handles the report window to save the PDF.
    Supports both 'frmVisualizador' (Comunicado) and 'Emissão DAM' (DAM).
    """
    print(f"   -> 💾 Initiating Save Routine for: {filename}...")
    
    try:
        # 1. Connect to the Viewer Window
        print("      ⏳ Waiting for report window to appear (Max 60s)...")
        
        viewer = None
        # Use regex to match either "frmVisualizador" OR "Emissão DAM"
        # The ^ char ensures it starts with one of those, avoiding false positives
        title_pattern = r"^(frmVisualizador|Emissão DAM).*"
        
        for _ in range(5):
            try:
                viewer = app.window(title_re=title_pattern)
                if viewer.exists(timeout=2):
                    print(f"      ✅ Detected window: '{viewer.window_text()}'")
                    break
            except: pass
            time.sleep(1)
            
        if not viewer or not viewer.exists(timeout=60):
            print("      ❌ Error: Report window (Visualizador/DAM) did not appear within 60s.")
            return False
            
        viewer.wait('visible', timeout=30)
        viewer.wait('ready', timeout=30)
        viewer.set_focus()
        time.sleep(2) # Extra buffer for rendering

        # 2. Click the "Save/Export" button (First button in toolbar)
        # Both windows share the "Arquivo" menu text.
        if not click_text_via_ocr(viewer, "Arquivo", y_offset=35):
            print("      ⚠️ OCR failed to find 'Arquivo'. Trying blind coordinate click...")
            # Fallback: Top-Left corner relative to window (Approx 25px Right, 60px Down)
            viewer.click_input(coords=(25, 60))
        
        # 3. Handle "Salvar Saída de Impressão como" Dialog
        print("      ⏳ Waiting for Save Dialog...")
        save_dlg = app.window(title_re=".*Salvar.*") 
        save_dlg.wait('ready', timeout=15)
        
        # 4. Type Filename and Save
        print(f"      ⌨️ Typing filename: {filename}")
        save_dlg.type_keys(filename, with_spaces=True)
        time.sleep(0.5)
        save_dlg.type_keys("{ENTER}")
        
        # 5. Handle "File Already Exists" (Overwrite)
        try:
            confirm_dlg = app.window(title="Confirmar Salvar Como")
            if confirm_dlg.exists(timeout=2):
                confirm_dlg.type_keys("%s") # Alt+S for 'Sim'
                print("      ⚠️ Overwriting existing file.")
        except: pass

        # 6. Close the Viewer Window
        print("      ❌ Closing Viewer...")
        viewer.close()
        
        # 7. CRITICAL: Wait for it to actually vanish
        try:
            viewer.wait_not('visible', timeout=15)
            print("      ✅ Viewer window closed successfully.")
        except:
            print("      ⚠️ Viewer window might still be open (wait timed out).")
        
        return True

    except Exception as e:
        print(f"      ❌ Save Error: {e}")
        return False

# ---------------------------------------------------------
# EXTRACTION HELPERS (Unchanged)
# ---------------------------------------------------------
def extract_idd_from_image(window_obj):
    try:
        full_img = window_obj.capture_as_image()
        w, h = full_img.size
        crop_box = (w - 400, 100, w - 10, 300) 
        header_img = full_img.crop(crop_box)
        if header_img.mode == 'RGB': r, g, b = header_img.split(); header_img = g
        else: header_img = ImageOps.grayscale(header_img)
        header_img = header_img.resize((header_img.width * 3, header_img.height * 3), Image.Resampling.LANCZOS)
        enhancer = ImageEnhance.Contrast(header_img)
        header_img = enhancer.enhance(3.0) 
        header_img = header_img.point(lambda x: 0 if x < 140 else 255)
        text = pytesseract.image_to_string(header_img, config=r'--psm 6')
        numbers = re.findall(r"\d{4,8}", text)
        if numbers: return numbers[-1]
        return None
    except: return None

def extract_text_from_element(element, config_type="number"):
    try:
        img = element.capture_as_image()
        img = img.resize((img.width * 3, img.height * 3), Image.Resampling.LANCZOS)
        img = ImageOps.grayscale(img)
        img = img.point(lambda x: 0 if x < 140 else 255)
        img = ImageOps.expand(img, border=20, fill='white')
        if config_type == "code": cfg = r'--psm 7 -c tessedit_char_whitelist=0123456789-/'
        else: cfg = r'--psm 7 -c tessedit_char_whitelist=0123456789.,'
        return pytesseract.image_to_string(img, config=cfg).strip()
    except: return ""

def clean_protocolo_string(text):
    text = text.strip().replace(" ", "")
    parts = text.split('/')
    if len(parts) > 2:
        year = parts[-1]
        code_body = "".join(parts[:-1])
        return f"{code_body}/{year}"
    return text

# ---------------------------------------------------------
# NAVIGATION HELPERS
# ---------------------------------------------------------
def trigger_documents_menu(app, window_obj, item_name):
    """
    Clicks 'Documentos' ARROW using OCR offset, then uses Keyboard.
    """
    print(f"   -> 🖱️ Triggering Menu: {item_name}...")
    
    menu_open = False
    
    # Try Visual Click with OFFSET to hit the arrow (Right 40px)
    if click_text_via_ocr(window_obj, "Documentos", x_offset=40):
        menu_open = True
    
    if not menu_open:
        print("      ❌ Failed to click 'Documentos' arrow.")
        return False

    time.sleep(1.5) # Wait for animation

    # 2. Select Item using Keyboard
    print(f"      ⌨️ Navigating to '{item_name}' via keyboard...")
    
    keys_to_press = ""
    # Menu structure: [Termo, Relatório, Comunicado, DAM]
    if item_name == "Comunicado ISS":
        keys_to_press = "{DOWN 3}{ENTER}"
    elif item_name == "DAM":
        keys_to_press = "{DOWN 4}{ENTER}"
    else:
        print(f"      ⚠️ Unknown menu item '{item_name}'")
        return False

    try:
        window_obj.type_keys(keys_to_press, pause=0.5)
        print(f"      ✅ Sent keys for '{item_name}'")
        return True
    except Exception as e:
        print(f"      ❌ Keyboard Error: {e}")
        return False

# ---------------------------------------------------------
# MAIN TEST LOGIC
# ---------------------------------------------------------
def test_emission_window():
    print("⏳ Connecting to application...")
    # Define a test IMU number (or get it from extraction)
    
    try:
        app = Application(backend="win32").connect(title_re=TITLE_RE, timeout=10)
        target_window = app.window(class_name="TfrmIDDEmissao")
        
        if not target_window.exists():
            print("❌ Error: Window 'TfrmIDDEmissao' not found.")
            return

        print("   -> 🪟 Maximizing...")
        try: target_window.maximize()
        except: pass
        time.sleep(1)
        target_window.set_focus()

        # 1. EXTRACT IDD
        print("\n--- 1. EXTRACTING IDD ---")
        idd_num = extract_idd_from_image(target_window)
        if idd_num:
            print(f"✅ IDD Found: {idd_num}")
        else:
            idd_num = "UNKNOWN_IDD"
            print("❌ IDD Failed")

        # 2. EXTRACT PROTOCOLO
        print("\n--- 2. EXTRACTING PROTOCOLO ---")
        try:
            proto_box = target_window.child_window(class_name="TMTProtocolo")
            if proto_box.exists():
                raw = extract_text_from_element(proto_box, "code")
                print(f"✅ Protocolo Found: {clean_protocolo_string(raw)}")
        except: pass

        # 3. TEST DOCUMENT GENERATION
        print("\n--- 3. TESTING DOCUMENT GENERATION ---")
        
        # Test A: Comunicado
        if trigger_documents_menu(app, target_window, "Comunicado ISS"):
            # Call the save routine
            filename = f"{idd_num}_Comunicado.pdf"
            if save_report_pdf(app, filename):
                print(f"✅ Successfully saved {filename}")
            else:
                print(f"❌ Failed to save {filename}")
        
        # Refocus main window
        try: target_window.set_focus()
        except: pass
        
        print("   -> 🛑 Pausing 3s before next document to allow app reset...")
        time.sleep(3) # Increased pause

        # Test B: DAM
        if trigger_documents_menu(app, target_window, "DAM"):
            filename = f"{idd_num}_DAM.pdf"
            if save_report_pdf(app, filename):
                print(f"✅ Successfully saved {filename}")
            else:
                print(f"❌ Failed to save {filename}")

    except Exception as e:
        print(f"❌ Critical Error: {e}")

if __name__ == "__main__":
    test_emission_window()