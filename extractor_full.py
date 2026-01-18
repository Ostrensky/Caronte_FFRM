
import time
import sys
import re
import os
import pyautogui
import pytesseract
from PIL import Image, ImageOps, ImageEnhance
from pywinauto import Application

# --- CONFIGURATION ---
PYTESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
TITLE_RE = r"^Comércio.*"
pytesseract.pytesseract.tesseract_cmd = PYTESSERACT_PATH

# ---------------------------------------------------------
# HELPER: IDD EXTRACTION (Green Channel Method)
# ---------------------------------------------------------
def extract_idd_from_image(window_obj):
    try:
        full_img = window_obj.capture_as_image()
        w, h = full_img.size
        # Adjusted crop to ensure we catch the IDD in the header
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
    except Exception as e:
        print(f"   [IDD OCR] Error: {e}")
        return None

# ---------------------------------------------------------
# HELPER: GENERIC TEXT EXTRACTION
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# HELPER: VISUAL OCR CLICK
# ---------------------------------------------------------
def click_text_via_ocr(window_obj, text_to_find, exact_match=False, x_offset=0, y_offset=0):
    print(f"      📷 Visual Scan: Looking for text '{text_to_find}'...")
    try:
        img = window_obj.capture_as_image()
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
                click_x = x + (w // 2)
                click_y = y + (h // 2)
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
# HELPER: CLEANERS
# ---------------------------------------------------------
def clean_protocolo_string(text):
    text = text.strip().replace(" ", "")
    parts = text.split('/')
    if len(parts) > 2:
        year = parts[-1]
        code_body = "".join(parts[:-1])
        return f"{code_body}/{year}"
    return text

def parse_money(text):
    clean = text.replace("R$", "").replace(" ", "").strip()
    clean = clean.replace(".", "").replace(",", ".")
    try: return float(clean)
    except: return 0.0

# ---------------------------------------------------------
# OCR FUNCTION
# ---------------------------------------------------------
def get_value_via_ocr(tab_window):
    print("   [OCR] Scanning bottom-right for value...")
    candidates = []
    tab_rect = tab_window.rectangle()
    search_top = tab_rect.bottom - 300 
    search_left = tab_rect.left + (tab_rect.width() // 2)

    for element in tab_window.descendants(class_name="TPanel"):
        try:
            r = element.rectangle()
            if r.top > search_top and r.left > search_left:
                candidates.append(element)
        except: pass

    if not candidates: return 0.0
    target = sorted(candidates, key=lambda e: e.rectangle().left, reverse=True)[0]
    
    img = target.capture_as_image()
    img = img.crop((0, 0, img.width, int(img.height * 0.75)))
    img = img.resize((img.width * 4, img.height * 4), Image.Resampling.LANCZOS)
    img = ImageOps.grayscale(img)
    img = img.point(lambda x: 0 if x < 115 else 255)
    img = ImageOps.expand(img, border=20, fill='white')
    text = pytesseract.image_to_string(img, config=r'--psm 7 -c tessedit_char_whitelist=0123456789.,')
    return parse_money(text)

# ---------------------------------------------------------
# NAVIGATION HELPER
# ---------------------------------------------------------
def trigger_documents_menu(app, window_obj, item_name):
    print(f"   -> 🖱️ Triggering Menu: {item_name}...")
    menu_open = False
    
    # --- SMART RETRY LOOP ---
    max_retries = 10
    
    for attempt in range(max_retries):
        try:
            if attempt > 0: 
                window_obj.set_focus()
                time.sleep(0.3)

            # Try to click 'Documentos'
            if click_text_via_ocr(window_obj, "Documentos", x_offset=40):
                menu_open = True
                break 
            
            print(f"      ⏳ Waiting for 'Documentos' button (Attempt {attempt+1}/{max_retries})...")
            time.sleep(0.5) 
                
        except Exception as e:
            print(f"      ⚠️ Attempt {attempt} Error: {e}")
            time.sleep(0.5)

    if not menu_open:
        print("      ❌ Failed to click 'Documentos' arrow.")
        return False

    # Wait for dropdown to technically render
    time.sleep(1.0)

    print(f"      ⌨️ Navigating to '{item_name}' via keyboard...")
    keys_to_press = ""
    if item_name == "Comunicado ISS":
        keys_to_press = "{DOWN 3}{ENTER}"
    elif item_name == "DAM":
        keys_to_press = "{DOWN 4}{ENTER}"
    else:
        return False

    try:
        window_obj.type_keys(keys_to_press, pause=0.5)
        return True
    except Exception as e:
        print(f"      ❌ Keyboard Error: {e}")
        return False

# ---------------------------------------------------------
# HELPER: SAVE PDF ROUTINE (ROBUST)
# ---------------------------------------------------------
def save_report_pdf(app, filename):
    filename = os.path.normpath(filename)
    
    # Ensure no stale file exists
    if os.path.exists(filename):
        try: os.remove(filename)
        except: pass

    MAX_RETRIES = 3
    
    for attempt in range(1, MAX_RETRIES + 1):
        print(f"   -> 💾 Initiating Save Routine for: {os.path.basename(filename)} (Attempt {attempt})...")
        
        try:
            print("      ⏳ Waiting for report window...")
            viewer = None
            title_pattern = r"^(frmVisualizador|Emissão DAM).*"
            
            # Fast detection loop
            for _ in range(15):
                try:
                    viewer = app.window(title_re=title_pattern)
                    if viewer.exists(timeout=1):
                        print(f"      ✅ Detected window: '{viewer.window_text()}'")
                        break
                except: pass
                time.sleep(0.5)
                
            if not viewer or not viewer.exists():
                print("      ❌ Error: Report window not found.")
                # If we failed to find the window, we can't save. Return False so the main loop can decide what to do.
                # However, maybe the menu click failed? The main loop logic will just fail this file.
                return False
                
            viewer.wait('visible', timeout=10)
            try: viewer.maximize()
            except: pass
            viewer.set_focus()
            time.sleep(1) 

            # Click File/Archive
            if not click_text_via_ocr(viewer, "Arquivo", y_offset=35):
                print("      ⚠️ OCR failed to find 'Arquivo'. Trying blind click...")
                viewer.click_input(coords=(25, 60))
            
            print("      ⏳ Waiting for Save Dialog...")
            save_dlg = app.window(title_re=".*Salvar.*") 
            
            if not save_dlg.exists(timeout=5):
                 # Retry click if dialog didn't appear
                 viewer.click_input(coords=(25, 60))
                 
            save_dlg.wait('ready', timeout=10)
            
            save_dlg.type_keys("^a{DELETE}", pause=0.1)
            save_dlg.type_keys(filename, with_spaces=True)
            time.sleep(0.5)
            save_dlg.type_keys("{ENTER}")
            
            # Handle overwrite popup
            try:
                confirm_dlg = app.window(title="Confirmar Salvar Como")
                if confirm_dlg.exists(timeout=2):
                    confirm_dlg.type_keys("%s")
            except: pass
            
            # 🛑 DISK VERIFICATION 🛑
            print("      ⏳ Verifying file on disk...")
            file_saved = False
            for _ in range(20): # Check for 10 seconds
                if os.path.exists(filename) and os.path.getsize(filename) > 0:
                    file_saved = True
                    break
                time.sleep(0.5)
            
            print("      ❌ Closing Viewer...")
            viewer.close()
            try: viewer.wait_not('visible', timeout=5)
            except: pass
            
            if file_saved:
                print("      ✅ File verification successful!")
                return True
            else:
                print("      ⚠️ File not found on disk after save operation. Retrying...")
                # The loop continues to next attempt

        except Exception as e:
            print(f"      ❌ Save Error (Attempt {attempt}): {e}")
            try: app.window(title_re=title_pattern).close()
            except: pass
            
    print("      ❌ All save attempts failed.")
    return False

# ---------------------------------------------------------
# MAIN PROCESS
# ---------------------------------------------------------
def process_company(imu, year, expected_value=None, run_emission=False, output_folder=None):
    imu_formatted = str(imu).strip().zfill(9)
    print(f"\n--- PROCESSING: IMU {imu_formatted} | Year {year} ---")
    
    result_data = {
        "status": "Failed",
        "extracted_original": 0.0,
        "idd_number": None,
        "protocolo": None,
        "final_corrected_value": 0.0
    }

    try:
        app = Application(backend="win32").connect(title_re=TITLE_RE, timeout=10)
        main_win = app.window(title_re=TITLE_RE)
        
        tab = main_win.child_window(class_name="TfrmIDDValidador", found_index=0)
        if not tab.exists():
            print("❌ Error: Main tab not open.")
            return result_data
        
        print("   -> 🧹 Clearing screen (Alt+L)...")
        try:
            tab.set_focus()
            tab.type_keys("%l")
            time.sleep(1.0) # Wait for clear action
        except Exception as e:
            print(f"   -> ⚠️ Warning: Failed to clear screen: {e}")

        print("   -> Filling Form...")
        imu_field = tab.child_window(class_name="TCSInscricaoMunicipal")
        imu_field.click_input(double=True)
        imu_field.type_keys(imu_formatted)
        imu_field.type_keys("{TAB 2}") 
        tab.type_keys("+{END}{BACKSPACE}") 
        tab.type_keys(str(year), pause=0.1)
        
        month_dropdown = tab.child_window(class_name="TcxComboBox", found_index=0)
        month_dropdown.click_input()
        month_dropdown.type_keys("{HOME}+{END}{DELETE}{TAB}")

        print("   -> Clicking Consultar (Alt+C)...")
        tab.type_keys("%c") 
        time.sleep(4) 

        extracted_value = get_value_via_ocr(tab)
        result_data["extracted_original"] = extracted_value
        print(f"✅ OCR Result: {extracted_value}")

        if expected_value is not None:
            diff = abs(extracted_value - expected_value)
            tolerance = max(extracted_value, expected_value) * 0.001
            is_valid_value = expected_value > 0.01

            if diff <= tolerance and is_valid_value:
                print(f"   -> ✅ MATCH! Values aligned.")
                result_data["status"] = "Validated" 

                if run_emission:
                    print("   -> 🚀 Sending Alt+E (Emitir)...")
                    tab.type_keys("%e")
                    
                    print("   -> ⏳ Waiting for Emission Window...")
                    # Small wait for window init, but loop handles the rest
                    time.sleep(2)
                    
                    try: 
                        emission_win = app.window(class_name="TfrmIDDEmissao")
                        if emission_win.exists(timeout=10):
                            
                            print("   -> 🪟 Maximizing Emission Window...")
                            try: emission_win.maximize()
                            except: pass
                            time.sleep(1)
                            emission_win.set_focus()
                            
                            print("   -> ⌨️ Sending Alt+I (Inscrever)...")
                            emission_win.type_keys("%i")
                            time.sleep(1)
                            
                            print("   -> ⌨️ Sending ENTER to confirm...")
                            emission_win.type_keys("{ENTER}")
                            
                            # --- 🚀 SMART LOOP (Wait for IDD) ---
                            # Interleaved check: Look for Popup OR Look at Screen Header
                            print("   -> ⏳ Waiting for IDD (Popup or Header)...")
                            
                            idd_found_in_loop = False
                            start_time = time.time()
                            
                            while time.time() - start_time < 90:
                                # 1. Check for Popups
                                current_popup = None
                                for title in ["Atenção", "Informação", "Information", "Aviso"]:
                                    dlg = app.window(title=title)
                                    if dlg.exists(timeout=0.2) and dlg.is_visible():
                                        current_popup = dlg
                                        break
                                
                                if current_popup:
                                    popup_title = current_popup.window_text()
                                    print(f"   -> 🔔 Popup detected: '{popup_title}'")
                                    
                                    try:
                                        children_texts = [c.window_text() for c in current_popup.descendants(control_type="Static")]
                                        full_text = popup_title + " " + " ".join(children_texts)
                                    except:
                                        full_text = popup_title

                                    match = re.search(r"N[ºo].*?(\d+)", full_text)
                                    if match:
                                        result_data["idd_number"] = match.group(1)
                                        print(f"      ✅ IDD Found in Popup: {result_data['idd_number']}")
                                        idd_found_in_loop = True
                                    
                                    print("      🖱️ Closing Popup...")
                                    current_popup.set_focus()
                                    time.sleep(0.2)
                                    try: current_popup.type_keys("{ENTER}")
                                    except: pass
                                    
                                    # If found, break immediately
                                    if idd_found_in_loop:
                                        break 
                                else:
                                    # 2. IF NO POPUP -> Check Background Header
                                    # This enables "click immediately" behavior if popup was missed
                                    print("      📷 Checking header for IDD...")
                                    idd_visual = extract_idd_from_image(emission_win)
                                    if idd_visual:
                                        result_data["idd_number"] = idd_visual
                                        print(f"      ✅ IDD Found on Screen: {idd_visual}")
                                        idd_found_in_loop = True
                                        break
                                    
                                    time.sleep(0.5) 

                            if not idd_found_in_loop:
                                print("   -> ⚠️ Timeout: IDD not found in popup or on screen.")
                            # --- END SMART LOOP ---

                            print(f"   -> 🔢 Final IDD: {result_data['idd_number']}")

                            # B. EXTRACT PROTOCOLO (CRITICAL STEP)
                            try:
                                proto_box = emission_win.child_window(class_name="TMTProtocolo")
                                if proto_box.exists():
                                    raw = extract_text_from_element(proto_box, config_type="code")
                                    result_data["protocolo"] = clean_protocolo_string(raw)
                                    print(f"   -> 🆔 Protocolo: {result_data['protocolo']}")
                            except: pass

                            # C. EXTRACT FINAL VALUE
                            try:
                                win_rect = emission_win.rectangle()
                                search_y = win_rect.height() - 100 
                                search_x = win_rect.width() * 0.7 
                                candidates = []
                                for p in emission_win.descendants(class_name="TPanel"):
                                    r = p.rectangle()
                                    if r.width() < 10 or r.height() < 10: continue
                                    rel_top = r.top - win_rect.top
                                    rel_left = r.left - win_rect.left
                                    if rel_top > search_y and rel_left > search_x:
                                        candidates.append(p)
                                
                                if candidates:
                                    target_panel = sorted(candidates, key=lambda e: e.rectangle().left, reverse=True)[0]
                                    raw_final = extract_text_from_element(target_panel, config_type="number")
                                    result_data["final_corrected_value"] = parse_money(raw_final)
                                    print(f"   -> 💰 Final Value: {result_data['final_corrected_value']}")
                            except: pass
                            
                            if result_data["idd_number"]:
                                result_data["status"] = "Success"
                            
                            # --- PHASE 4: GENERATE DOCUMENTS ---
                            if result_data["idd_number"]:
                                print("\n   --- 📄 GENERATING DOCUMENTS ---")
                                idd_num = result_data["idd_number"]
                                
                                com_filename = f"{idd_num}_Comunicado.pdf"
                                dam_filename = f"{idd_num}_DAM.pdf"
                                
                                if output_folder and os.path.isdir(output_folder):
                                    com_path = os.path.join(output_folder, com_filename)
                                    dam_path = os.path.join(output_folder, dam_filename)
                                else:
                                    com_path = com_filename
                                    dam_path = dam_filename
                                
                                # trigger_documents_menu now retries quickly
                                if trigger_documents_menu(app, emission_win, "Comunicado ISS"):
                                    if not save_report_pdf(app, com_path):
                                        print("      ⚠️ Failed to save Comunicado.")
                                
                                try: emission_win.set_focus()
                                except: pass
                                time.sleep(1)

                                if trigger_documents_menu(app, emission_win, "DAM"):
                                    if not save_report_pdf(app, dam_path):
                                        print("      ⚠️ Failed to save DAM.")

                                # --- CLEANUP ---
                                print("   -> 🧹 Cleanup: Closing emission window (Alt+F, Alt+L)...")
                                try:
                                    emission_win.set_focus()
                                    time.sleep(0.5)
                                    emission_win.type_keys("%f")
                                    time.sleep(1)
                                    if tab.exists():
                                        tab.set_focus()
                                        tab.type_keys("%l")
                                except Exception as e:
                                    print(f"   -> ❌ Error during cleanup: {e}")

                        else:
                            print("   -> ❌ Emission Window not found.")
                    
                    except Exception as e: # CLOSE RPA TRY BLOCK
                        print(f"   -> ❌ Error inside Emission Logic: {e}")

                else: # ELSE FOR IF RUN_EMISSION
                    print(f"   -> 🛡️ VALIDATION SUCCESSFUL. Emission skipped.")

            else: # ELSE FOR IF DIFF <= TOLERANCE
                print(f"   -> ⚠️ DIVERGENCE! Diff ({diff:.2f}) > Tolerance.")
                result_data["status"] = "Divergence" 
                print("   -> 🧹 Pressing Alt+L (Limpar) to reset for next company...")
                try:
                    tab.type_keys("%l")
                except Exception as e:
                    print(f"   -> ⚠️ Failed to press Alt+L: {e}")

        return result_data

    except Exception as e:
        print(f"❌ Critical Error: {e}")
        return result_data
