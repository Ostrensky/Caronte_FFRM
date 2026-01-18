# --- app/ferramentas/simples_reader.py ---

import fitz  # PyMuPDF
import pytesseract
from PIL import Image, ImageOps, ImageEnhance
import io
import re
import os
import pandas as pd
from datetime import datetime, timedelta
from PySide6.QtCore import QThread

# --- CONFIGURATION ---
DEFAULT_TESS_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(DEFAULT_TESS_PATH):
    pytesseract.pytesseract.tesseract_cmd = DEFAULT_TESS_PATH

def check_stop_flag():
    """Checks if the user clicked Stop in the UI."""
    current_thread = QThread.currentThread()
    if hasattr(current_thread, 'check_stop') and current_thread.check_stop():
        return True
    return False

def verify_tesseract_installed():
    """Checks if Tesseract is available for scanned files."""
    try:
        pytesseract.get_tesseract_version()
        return True, "OK"
    except:
        return False, "Tesseract OCR não instalado."

# --- DATE PARSING ENGINE ---
def parse_date_strict(date_str):
    """
    Parses strictly formatted dates dd/mm/yyyy.
    Used when we have clean digital text.
    """
    try:
        return datetime.strptime(date_str, "%d/%m/%Y")
    except ValueError:
        return None

def parse_date_fuzzy(text):
    """
    Parses broken OCR dates (e.g. '0l/01/2O22').
    Returns a list of datetime objects found in the string.
    """
    # 1. Map common OCR typos to digits
    subs = {
        'O': '0', 'o': '0', 'D': '0', 'Q': '0',
        'l': '1', 'I': '1', 'i': '1', 'L': '1', '|': '1', '!': '1',
        'Z': '2', 'z': '2',
        'S': '5', 's': '5', '$': '5',
        'B': '8', 'E': '8',
        'g': '9',
        '.': '/', '-': '/', ' ': '/'  # Normalize separators
    }
    
    clean_text = ""
    for char in text:
        if char.isdigit() or char == '/':
            clean_text += char
        elif char in subs:
            clean_text += subs[char]
        else:
            clean_text += " " # Replace garbage with space to separate numbers

    # 2. Extract standard pattern
    matches = re.finditer(r"(\d{2})/(\d{2})/(\d{4})", clean_text)
    dates = []
    for m in matches:
        try:
            d, m, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1990 <= y <= 2100 and 1 <= m <= 12 and 1 <= d <= 31:
                dates.append(datetime(y, m, d))
        except:
            continue
    return dates

# --- EXTRACTION ENGINE ---
def get_pdf_content_hybrid(pdf_path):
    """
    OVERHAUL: Hybrid Strategy.
    1. Try reading text directly (fast, 100% accurate for digital PDFs).
    2. If text is empty/garbage, fallback to heavy OCR (scanned PDFs).
    """
    full_text = ""
    used_ocr = False
    
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            if check_stop_flag():
                doc.close()
                return "STOPPED"
            
            # A. Try Digital Text
            text = page.get_text("text")
            
            # If page has < 50 chars, it's likely an image scan. Switch to OCR.
            if len(text.strip()) < 50:
                # B. OCR Fallback (4x Zoom for maximum precision)
                pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), dpi=300)
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))
                
                # Pre-processing
                img = img.convert('L')
                enhancer = ImageEnhance.Contrast(img)
                img = enhancer.enhance(2.0)
                
                # We do NOT use binary thresholding anymore (it kills faint text)
                
                ocr_text = pytesseract.image_to_string(img, lang='por', config='--psm 6')
                full_text += " " + ocr_text
                used_ocr = True
            else:
                # Use the perfect digital text
                full_text += " " + text

        doc.close()
    except Exception:
        return "ERROR"
        
    return full_text

# --- LOGIC ENGINE ---
def analyze_simples_data(text, target_year):
    """
    Decides status based on strict block analysis.
    """
    if not text or "ERROR" in text: return "Error"
    if "STOPPED" in text: return "STOPPED"
    
    # Normalize
    text_upper = text.upper()
    text_clean = " ".join(text_upper.split()) # Flatten newlines

    # 1. SANITIZE: Remove "Data da Consulta"
    # This prevents the print date from being confused with a Start Date
    text_clean = re.sub(r"DATA DA CONSULTA.*?(\d{2}/\d{2}/\d{4})", " ", text_clean)

    # 2. IDENTIFY CURRENT STATUS
    # We look for "Situação no Simples Nacional"
    # Logic: Defaults to None. If we find "NÃO OPTANTE", set False. If "OPTANTE", set True.
    is_currently_optant = False
    if "SITUAÇÃO NO SIMPLES NACIONAL: NÃO OPTANTE" in text_clean or "NAO OPTANTE" in text_clean:
        is_currently_optant = False
    elif "SITUAÇÃO NO SIMPLES NACIONAL: OPTANTE" in text_clean:
        is_currently_optant = True
    elif "OPTANTE PELO SIMPLES NACIONAL" in text_clean: # Fallback
        is_currently_optant = True

    # 3. EXTRACT PERIODS (The Core Logic)
    periods = []

    # A. Active "Desde" (Since)
    # Looks for "Desde dd/mm/yyyy"
    desde_matches = re.findall(r"DESDE\s*(\d{2}/\d{2}/\d{4})", text_clean)
    for d_str in desde_matches:
        d = parse_date_strict(d_str)
        if not d: d = parse_date_fuzzy(d_str)[0] if parse_date_fuzzy(d_str) else None
        
        if d:
            periods.append({'start': d, 'end': datetime(2999, 12, 31)})

    # B. History Table
    # We isolate the text between "Períodos Anteriores" and "Eventos Futuros"
    # This prevents us from reading garbage dates elsewhere.
    history_block = ""
    try:
        start_idx = text_clean.find("PERÍODOS ANTERIORES")
        end_idx = text_clean.find("EVENTOS FUTUROS")
        if start_idx != -1:
            if end_idx != -1:
                history_block = text_clean[start_idx:end_idx]
            else:
                history_block = text_clean[start_idx:] # Read to end
    except:
        pass

    # If "Não Existem" is in the history block, we know there is no history.
    history_exists = "NÃO EXISTEM" not in history_block and "NAO EXISTEM" not in history_block

    if history_exists:
        # Find all dates inside the history block
        raw_dates = parse_date_fuzzy(history_block)
        
        # We expect pairs: Start, End. 
        # Logic: If we have an even number of dates, we pair them.
        # If odd, the last one might be an open start, but usually Simples history is closed.
        if len(raw_dates) >= 2:
            # Sort to be safe, though usually they appear in order
            # Actually, don't sort yet, preserve table order (Start... End)
            
            # Simple pairing: Date 1 = Start, Date 2 = End
            # The table format is usually: Start [junk] End [junk] Reason
            # So looking for (Date... Date) patterns is safest.
            
            # We iterate and verify strict Start <= End logic
            i = 0
            while i < len(raw_dates) - 1:
                d1 = raw_dates[i]
                d2 = raw_dates[i+1]
                
                if d1 <= d2:
                    periods.append({'start': d1, 'end': d2})
                    i += 2
                else:
                    # If d1 > d2, something is wrong (maybe End Date came first? or broken OCR).
                    # Skip d1 and try d2 as start
                    i += 1

    # 4. CALCULATE OVERLAP
    y_start = datetime(target_year, 1, 1)
    y_end = datetime(target_year, 12, 31)
    
    # Flag: Did we match any period for this year?
    matched = False
    ov_s, ov_e = None, None

    for p in periods:
        if p['start'] <= y_end and p['end'] >= y_start:
            matched = True
            this_s = max(p['start'], y_start)
            this_e = min(p['end'], y_end)
            
            if ov_s is None:
                ov_s, ov_e = this_s, this_e
            else:
                ov_s = min(ov_s, this_s)
                ov_e = max(ov_e, this_e)

    # 5. FINAL VERDICT
    if matched:
        if ov_s <= y_start and ov_e >= y_end:
            return "Full"
        else:
            return f"Partial ({ov_s.strftime('%d/%m/%Y')} - {ov_e.strftime('%d/%m/%Y')})"
    
    # If no periods matched the target year:
    if is_currently_optant:
        # It says "Optante" now, but we didn't find a "Desde" date that covers this year?
        # This is rare. Usually "Optante" implies "Desde [past date]".
        # If we missed the date but it says Optante, we default to FULL to avoid False Negative.
        return "Full"
        
    return "Not"

def consolidate_status(year_statuses):
    """
    Merges {2022: Not, 2023: Full} into ranges.
    """
    ranges = []
    years = sorted(year_statuses.keys())
    
    for y in years:
        st = year_statuses[y]
        if st in ["Not", "STOPPED", "Error"]: continue
            
        s, e = None, None
        if st == "Full":
            s, e = datetime(y, 1, 1), datetime(y, 12, 31)
        elif "Partial" in st:
            m = re.search(r"(\d{2}/\d{2}/\d{4}).*?(\d{2}/\d{2}/\d{4})", st)
            if m:
                s = datetime.strptime(m.group(1), "%d/%m/%Y")
                e = datetime.strptime(m.group(2), "%d/%m/%Y")
        
        if s and e: ranges.append((s, e))
        
    if not ranges: return "Not"
    
    # Merge
    ranges.sort(key=lambda x: x[0])
    merged = []
    if ranges:
        curr_s, curr_e = ranges[0]
        for next_s, next_e in ranges[1:]:
            if next_s <= curr_e + timedelta(days=1):
                curr_e = max(curr_e, next_e)
            else:
                merged.append((curr_s, curr_e))
                curr_s, curr_e = next_s, next_e
        merged.append((curr_s, curr_e))
        
    return ", ".join([f"{s.strftime('%d/%m/%Y')}-{e.strftime('%d/%m/%Y')}" for s, e in merged])

def run_simples_reader(root_folder, target_years, progress_callback):
    # 1. Install Check
    installed, msg = verify_tesseract_installed()
    if not installed:
        progress_callback.emit(f"❌ {msg}")
        return None

    if not isinstance(target_years, list): target_years = [target_years]
    
    progress_callback.emit("--- Iniciando Leitura (Overhaul Híbrido) ---")
    progress_callback.emit("Estratégia: Texto Digital (Prioridade) -> OCR (Backup)")
    
    results = []
    stopped = False
    
    for root, dirs, files in os.walk(root_folder):
        if check_stop_flag(): stopped = True; break
        
        target_file = None
        for f in files:
            if f.lower().endswith(".pdf") and ("optante" in f.lower() or "simples" in f.lower()):
                target_file = f
                break
        
        if target_file:
            path = os.path.join(root, target_file)
            progress_callback.emit(f"Processando: {target_file}")
            
            # STEP 1: Get Content (Hybrid)
            text = get_pdf_content_hybrid(path)
            
            if text == "STOPPED": stopped = True; break
            
            # STEP 2: Analyze
            yr_stats = {}
            for y in target_years:
                yr_stats[y] = analyze_simples_data(text, y)
                
            final = consolidate_status(yr_stats)
            
            results.append({
                "Pasta": os.path.basename(root),
                "Arquivo": target_file,
                "Status": final
            })
            
    if stopped:
        progress_callback.emit("🛑 Interrompido.")
        return None
        
    if results:
        df = pd.DataFrame(results)
        ts = datetime.now().strftime("%H%M%S")
        out = os.path.join(root_folder, f"Resultado_Simples_{ts}.xlsx")
        df.to_excel(out, index=False)
        progress_callback.emit(f"Salvo em: {out}")
        return out
    else:
        progress_callback.emit("Nenhum arquivo encontrado.")
        return None
