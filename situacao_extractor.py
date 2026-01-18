import time
import sys
from pathlib import Path
import os
import random
import math
import pyperclip
import pyautogui

# --- Pywinauto Imports ---
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys
from PySide6.QtCore import QThread 

# Set a global timeout for pywinauto actions
timings.Timeout = 20

# --- CONFIGURATION ---
WINDOW_TITLE_RE = r"^Comércio.*" 
FINALIDADE_DROPDOWN_CLASS = "TmtDBLookupCombo"
IMU_INPUT_FIELD_CLASS = "TCSInscricaoMunicipal"

# List of titles to aggressively close. 
# "Confirmação" is usually the one that appears after Shift+F8
BLOCKING_POPUP_TITLES = ["Atenção", "Aviso", "Erro", "Confirmação", "Information", "Informação", "Question"]
SAVE_DIALOG_TITLE = "Salvar Saída de Impressão como" 
# --- END CONFIGURATION ---

# --------------------------------------------------------------------------------------
# ROBUST HELPER FUNCTIONS
# --------------------------------------------------------------------------------------

def dismiss_any_popup(app_instance, max_attempts=3):
    """
    Scans for any blocking popup and presses ENTER to dismiss/confirm it.
    Returns True if a popup was found and closed.
    """
    dismissed_something = False
    
    for _ in range(max_attempts):
        popup_found_in_pass = False
        
        for title in BLOCKING_POPUP_TITLES:
            try:
                # Fast check
                dlg = app_instance.window(title=title)
                if dlg.exists(timeout=0.3):
                    # Found one!
                    dlg.set_focus()
                    time.sleep(0.1)
                    send_keys("{ENTER}") # Confirms "Sim" or "OK"
                    
                    time.sleep(0.5) # Wait for it to vanish
                    dismissed_something = True
                    popup_found_in_pass = True
                    break 
            except Exception:
                pass
        
        # If we didn't find anything in this pass, stop trying
        if not popup_found_in_pass:
            break
            
    return dismissed_something

def wait_for_save_dialog_or_error(app_instance, save_title, timeout=15):
    """
    Waits for the Save Dialog. If an Error popup appears instead, it kills it and returns None.
    """
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        # 1. Check for Save Dialog
        try:
            save_win = app_instance.window(title=save_title)
            if save_win.exists(timeout=0.2):
                return save_win
        except: pass
            
        # 2. Check for Error/Blocking Popup
        if dismiss_any_popup(app_instance):
            # If we dismissed a popup, it means generation failed or finished with a message
            # We wait a split second to see if the Save window was just hidden behind it
            time.sleep(0.5)
            if save_win.exists(timeout=0.2): return save_win
            else: return None # The popup was likely an error, so no save window
        
        time.sleep(0.3)
        
    return None

def robust_paste_path(window, path_str):
    """Pastes path using Clipboard (Ctrl+V) for speed and accuracy."""
    try:
        pyperclip.copy(path_str)
        window.set_focus()
        time.sleep(0.3)
        send_keys("^v") 
        time.sleep(0.3)
        return True
    except Exception as e:
        return False

# --------------------------------------------------------------------------------------
# MAIN WORKER
# --------------------------------------------------------------------------------------

def run_situacao_extractor(data_list, progress_callback):
    
    def check_stop_flag():
        current_thread = QThread.currentThread()
        if hasattr(current_thread, 'check_stop') and current_thread.check_stop():
            return True
        return False

    if not data_list:
        progress_callback.emit("❌ Lista vazia.")
        return

    progress_callback.emit("\n--- RPA Iniciado (Modo Shift+F8) ---")
    
    # Connect to App
    try:
        app = Application(backend="win32").connect(title_re=WINDOW_TITLE_RE, timeout=10)
        main_window = app.window(title_re=WINDOW_TITLE_RE)
        main_window.maximize()
        main_window.set_focus()
    except Exception as e:
        progress_callback.emit(f"❌ Erro ao conectar: {e}")
        return
        
    try:
        pessoas_panel = main_window.child_window(title="Pessoas")
        input_field = pessoas_panel.child_window(class_name=IMU_INPUT_FIELD_CLASS, found_index=0)
    except Exception as e:
        progress_callback.emit(f"❌ Erro ao achar campos iniciais: {e}")
        return

    total = len(data_list)
    
    for i, record in enumerate(data_list):
        if check_stop_flag(): break

        imu_to_type = record.get('imu_id') or record.get('cnpj')
        dir_path = Path(record['dir_path'])
        
        if not imu_to_type: continue
        
        progress_callback.emit(f"\n--- {i+1}/{total}: IMU {imu_to_type} ---")

        try: 
            # 1. PREP: Dismiss any leftovers
            dismiss_any_popup(app)
            main_window.set_focus()

            # 2. SELECT DROPDOWN
            # Assuming 'd' selects the correct option.
            try:
                main_window.child_window(class_name=FINALIDADE_DROPDOWN_CLASS).click_input()
                send_keys("d{ENTER}")
                time.sleep(0.2)
            except: pass

            # 3. TYPE IMU
            input_field.set_focus()
            input_field.type_keys(imu_to_type, with_spaces=True)
            send_keys("{TAB}")
            
            # 4. GENERATE (Ctrl+E)
            send_keys("^E") 
            
            # 5. WAIT FOR SAVE
            save_win = wait_for_save_dialog_or_error(app, SAVE_DIALOG_TITLE, timeout=12)
            
            if save_win:
                new_filename = f"Situacao_{imu_to_type}.pdf"
                full_path = str(dir_path / new_filename)
                
                try:
                    robust_paste_path(save_win, full_path)
                    send_keys("{ENTER}")
                    
                    # Handle "Overwrite?" popup
                    if dismiss_any_popup(app): 
                        progress_callback.emit("-> Arquivo sobrescrito.")
                    
                    progress_callback.emit("-> Salvo.")
                except Exception as save_err:
                     progress_callback.emit(f"-> Erro ao salvar: {save_err}")
            else:
                progress_callback.emit("-> Erro: Janela de salvar não abriu.")

        except Exception as e:
            progress_callback.emit(f"Erro no Loop: {e}")

        finally:
            # --- CLEANUP PHASE (CRITICAL) ---
            # This ensures the screen is ready for the next company
            progress_callback.emit("   -> Limpando tela (Shift+F8)...")
            
            try:
                main_window.set_focus()
                
                # 1. Clear any 'Success' or 'Error' popups from the previous step
                dismiss_any_popup(app)
                
                # 2. Send Eraser Command
                send_keys("+{F8}") 
                
                # 3. Handle 'Confirmation' popup immediately
                # We loop briefly because the popup might take 500ms to animate
                time.sleep(0.5) 
                if dismiss_any_popup(app):
                    progress_callback.emit("   -> Limpeza confirmada.")
                
                # 4. Final safety clear just in case
                time.sleep(0.2)
                dismiss_any_popup(app)
                
            except Exception as cleanup_err:
                progress_callback.emit(f"⚠️ Erro na limpeza: {cleanup_err}")
                send_keys("{ESC}") # Panic button

    progress_callback.emit("\n✅ Concluído.")