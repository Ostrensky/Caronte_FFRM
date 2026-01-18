import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import re
import os
import pandas as pd
from datetime import datetime

# --- CONFIGURATION ---
# 1. Path to Tesseract (Keep this as is based on your previous success)
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# 2. The Main Folder containing all your subfolders
# CHANGE THIS to your actual folder path
ROOT_FOLDER = r"C:\Users\vostrensky\Documents\IDD\OIDD_XX_25" 

# 3. Target Year
TARGET_YEAR = 2021
# ---------------------

def parse_date(date_str):
    if not date_str: return None
    try:
        return datetime.strptime(date_str, "%d/%m/%Y")
    except ValueError:
        return None

def get_text_via_ocr(pdf_path):
    """
    Extracts text from PDF images using Tesseract (English mode to be safe).
    """
    if not os.path.exists(TESSERACT_PATH):
        return "ERROR_MISSING_TESSERACT"

    full_text = ""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            pix = page.get_pixmap(dpi=300)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # Using 'eng' to avoid the 'por' language pack crash
            text = pytesseract.image_to_string(img, lang='eng') 
            full_text += " " + text
    except Exception as e:
        return "" # Return empty string on corrupt file
        
    return full_text

def check_status_for_year(pdf_path, target_year):
    """
    Determines Simples Nacional status (Full, Partial, Not) for the target year.
    """
    y_start = datetime(target_year, 1, 1)
    y_end = datetime(target_year, 12, 31)
    
    # 1. Get Text
    text = get_text_via_ocr(pdf_path)
    
    if text == "ERROR_MISSING_TESSERACT":
        return "Error: Tesseract not found"
    if not text.strip():
        return "Error: OCR failed (Empty text)"

    flat_text = " ".join(text.split())
    periods = []

    # 2. Logic: Current Situation
    # Matches "Optante... desde dd/mm/yyyy" (Using fuzzy matching for OCR errors)
    current_match = re.search(r"Simples Nacional.*?desde\s*(\d{2}/\d{2}/\d{4})", flat_text, re.IGNORECASE)
    if current_match:
        start_date = parse_date(current_match.group(1))
        if start_date:
            periods.append({'start': start_date, 'end': None}) # None = Active

    # 3. Logic: History Table
    # Matches "01/07/2007 31/12/2008"
    history_matches = re.findall(r"(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})", flat_text)
    for d1, d2 in history_matches:
        s_date = parse_date(d1)
        e_date = parse_date(d2)
        if s_date and e_date and s_date <= e_date:
            periods.append({'start': s_date, 'end': e_date})

    if not periods:
        return "Not"

    # 4. Calculate Overlap
    is_optant = False
    overlap_s = None
    overlap_e = None

    for p in periods:
        p_start = p['start']
        p_end = p['end'] if p['end'] else datetime(9999, 12, 31)

        if p_start <= y_end and p_end >= y_start:
            is_optant = True
            overlap_s = max(p_start, y_start)
            overlap_e = min(p_end, y_end)
            break

    if not is_optant:
        return "Not"

    if overlap_s <= y_start and overlap_e >= y_end:
        return "Full"
    else:
        return f"Partial ({overlap_s.strftime('%d/%m/%Y')} to {overlap_e.strftime('%d/%m/%Y')})"

def process_folder_structure():
    print(f"--- Starting Batch Process for Year {TARGET_YEAR} ---")
    print(f"Scanning root: {ROOT_FOLDER}")
    
    results = []

    # os.walk goes through every subfolder recursively
    for root, dirs, files in os.walk(ROOT_FOLDER):
        
        # Look for a PDF file that looks like "consulta optantes"
        target_file = None
        for file in files:
            lower_name = file.lower()
            if "consulta" in lower_name and "optantes" in lower_name and lower_name.endswith(".pdf"):
                target_file = file
                break # Process only the first valid file found in this folder
        
        if target_file:
            full_path = os.path.join(root, target_file)
            folder_name = os.path.basename(root)
            
            print(f"Processing: {folder_name} -> {target_file}...")
            
            # Calculate Status
            status = check_status_for_year(full_path, TARGET_YEAR)
            
            # Append to results
            results.append({
                "Folder Name": folder_name,
                "File Used": target_file,
                f"Status {TARGET_YEAR}": status
            })

    # Generate Excel
    if results:
        df = pd.DataFrame(results)
        output_filename = f"Simples_Status_Results_{TARGET_YEAR}.xlsx"
        df.to_excel(output_filename, index=False)
        print(f"\n--- DONE! ---")
        print(f"Results saved to: {output_filename}")
        print(df)
    else:
        print("No matching files found.")

# --- EXECUTE ---
if __name__ == "__main__":
    process_folder_structure()