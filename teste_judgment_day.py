import sys
import pytesseract
from PIL import Image, ImageOps, ImageEnhance
from pywinauto import Application

# --- CONFIG ---
PYTESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
TITLE_RE = r"^Comércio.*"
pytesseract.pytesseract.tesseract_cmd = PYTESSERACT_PATH

def debug_ocr_extraction():
    print("⏳ Connecting for DEBUG OCR...")
    app = Application(backend="win32").connect(title_re=TITLE_RE, timeout=10)
    main_win = app.window(title_re=TITLE_RE)
    tab = main_win.child_window(class_name="TfrmIDDValidador", found_index=0)

    # 1. FIND PANEL
    target_panel = None
    tab_rect = tab.rectangle()
    search_top = tab_rect.bottom - 300 
    search_left = tab_rect.left + (tab_rect.width() // 2)

    candidates = []
    for element in tab.descendants(class_name="TPanel"):
        try:
            r = element.rectangle()
            if r.top > search_top and r.left > search_left:
                candidates.append(element)
        except: pass

    if not candidates:
        print("❌ No panel found.")
        return

    target_panel = sorted(candidates, key=lambda e: e.rectangle().left, reverse=True)[0]
    
    # 2. CAPTURE RAW (Save this!)
    img = target_panel.capture_as_image()
    img.save("debug_1_raw.png")
    print("📸 Saved 'debug_1_raw.png' (Check this file!)")

    # 3. PROCESS IMAGE (The New "Soft" Method)
    # Resize 3x
    img = img.resize((img.width * 3, img.height * 3), Image.Resampling.LANCZOS)
    
    # Grayscale
    img = ImageOps.grayscale(img)
    
    # Invert? (Sometimes text is white on gray. Inverting makes it black on white)
    # Uncomment the next line if your text is Light-colored
    # img = ImageOps.invert(img)

    # Auto Contrast (Instead of manual Threshold)
    # This maximizes the difference between text and background automatically
    img = ImageOps.autocontrast(img, cutoff=2)
    
    # Add Border
    img = ImageOps.expand(img, border=20, fill='white')

    # Save Processed (Save this!)
    img.save("debug_2_processed.png")
    print("📸 Saved 'debug_2_processed.png' (Does this look readable?)")

    # 4. RUN OCR (No Whitelist - Let's see what it sees)
    # We remove the whitelist to see if it detects garbage like "R$ _" or "ii"
    # We use --psm 6 (Assume a block of text) instead of 7
    config = r'--psm 6' 
    
    try:
        raw_text = pytesseract.image_to_string(img, config=config)
        print(f"\n🧐 OCR RAW OUTPUT: '{raw_text.strip()}'")
        
        # Simple Cleaner
        clean = raw_text.replace("R$", "").replace(" ", "").strip()
        clean = clean.replace(".", "").replace(",", ".")
        print(f"💵 Interpretation: {clean}")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    debug_ocr_extraction()