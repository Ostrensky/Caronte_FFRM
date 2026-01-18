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
# EXTRACTION HELPERS
# ---------------------------------------------------------
def extract_idd_from_image(window_obj):
    try:
        full_img = window_obj.capture_as_image()
        w, h = full_img.size
        # Top-Right Region
        crop_box = (w - 400, 100, w - 10, 300) 
        header_img = full_img.crop(crop_box)
        
        if header_img.mode == 'RGB':
            r, g, b = header_img.split()
            header_img = g
        else:
            header_img = ImageOps.grayscale(header_img)

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

        if config_type == "code":
            cfg = r'--psm 7 -c tessedit_char_whitelist=0123456789-/'
        else:
            cfg = r'--psm 7 -c tessedit_char_whitelist=0123456789.,'

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
# MAIN TEST LOGIC
# ---------------------------------------------------------
def test_emission_window():
    print("⏳ Connecting to application...")
    try:
        app = Application(backend="win32").connect(title_re=TITLE_RE, timeout=10)
        
        target_window = app.window(class_name="TfrmIDDEmissao")
        if not target_window.exists():
            print("❌ Error: Window 'TfrmIDDEmissao' not found. Please open it first.")
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
            print("❌ IDD Extraction Failed.")

        # 2. EXTRACT PROTOCOLO
        print("\n--- 2. EXTRACTING PROTOCOLO ---")
        try:
            proto_box = target_window.child_window(class_name="TMTProtocolo")
            if proto_box.exists():
                raw = extract_text_from_element(proto_box, "code")
                print(f"✅ Protocolo Found: {clean_protocolo_string(raw)}")
        except: pass

       

    except Exception as e:
        print(f"❌ Critical Error: {e}")

if __name__ == "__main__":
    test_emission_window()