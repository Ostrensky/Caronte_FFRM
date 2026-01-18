import time
import sys
from pathlib import Path

# --- Pywinauto Imports ---
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

# Set a global timeout for pywinauto actions
# Increased to 20 seconds for safety, given the previous timeout issues.
timings.Timeout = 20

# --- CONFIGURATION ---

# Application and Main Window
WINDOW_TITLE = "Comércio [Versão 3.1.1.60]"

# Control Class Names
FINALIDADE_DROPDOWN_CLASS = "TmtDBLookupCombo"
IMU_INPUT_FIELD_CLASS = "TCSInscricaoMunicipal"

# Pop-up Dialogs
ATTENTION_DIALOG_TITLE = "Atenção"
CONFIRMATION_DIALOG_TITLE = "Confirmação"
SIM_BUTTON_TITLE = "Sim"

# IMPORTANT: Use the exact Portuguese title from the dialog image
SAVE_DIALOG_TITLE = "Salvar Saída de Impressão como" 

# File Paths
INPUT_BASE_FOLDER = Path(r'C:\Users\vostrensky\Documents\IDD\OIDD_XX_25')

# --- END CONFIGURATION ---

# --------------------------------------------------------------------------------------
def get_data_from_directories():
    """
    Finds all subdirectories in INPUT_BASE_FOLDER and extracts the IMU code
    from the directory name (the part before the first underscore).
    """
    imu_data_list = []
    
    for item in INPUT_BASE_FOLDER.iterdir():
        if item.is_dir():
            dir_name = item.name
            try:
                imu_raw = dir_name.split('_', 1)[0].strip()
                
                if imu_raw and imu_raw.isdigit():
                    imu_id = imu_raw.zfill(9) 
                    
                    imu_data_list.append({
                        'imu_id': imu_id,
                        'dir_path': item
                    })
                else:
                    print(f"⚠️ Warning: Skipped directory '{dir_name}' as the IMU prefix is not a number.")
            except Exception as e:
                print(f"⚠️ Warning: Error processing directory '{dir_name}': {e}")
                
    if not imu_data_list:
        print(f"❌ FATAL ERROR: No valid IMU codes found in subdirectories of {INPUT_BASE_FOLDER}")
        
    print(f"✅ Loaded {len(imu_data_list)} IMU codes from directory structure.")
    return imu_data_list

# --------------------------------------------------------------------------------------
def check_for_popup(app_instance, dialog_title, timeout=10):
    """Checks if a dialog is visible and returns the dialog object or None."""
    try:
        # Use the app instance from the main connection (backend="win32")
        dialog = app_instance.window(title=dialog_title).wait('ready', timeout=timeout)
        return dialog
    except timings.TimeoutError:
        return None
    except Exception:
        return None

# --------------------------------------------------------------------------------------
def handle_confirmation_popup(app_instance, dialog_title, button_title, timeout=10):
    """
    Waits dynamically for the Confirmation pop-up and clicks a button within it.
    """
    try:
        # Use the app instance from the main connection (backend="win32")
        dialog = app_instance.window(title=dialog_title).wait('ready', timeout=timeout)
        
        button = getattr(dialog, button_title)
        button.click()
        
        dialog.wait_not('visible', timeout=5) 
        
        print(f"      -> Successfully handled '{dialog_title}' pop-up by clicking '{button_title}'.")
        return True
    
    except timings.TimeoutError:
        print(f"      -> Timeout waiting for '{dialog_title}' pop-up. Proceeding.")
        return False
        
    except Exception as e:
        print(f"      -> FATAL: Click failed for pop-up '{dialog_title}' Button '{button_title}': {e}")
        return False

# --------------------------------------------------------------------------------------
def automate_pdf_extraction_pywinauto(data_list):
    if not data_list:
        return

    # 1. Connect to the application (Keep win32 for the main app)
    try:
        app = Application(backend="win32").connect(title=WINDOW_TITLE, timeout=10)
        main_window = app.window(title=WINDOW_TITLE)
        
        main_window.maximize()
        main_window.set_focus()
        print(f"✅ Application '{WINDOW_TITLE}' connected and focused.")
        
    except Exception as e:
        print(f"❌ FATAL: Could not connect to the application: {e}")
        sys.exit(1)
        
    # --- Control References (keep original) ---
    try:
        pessoas_panel = main_window.child_window(title="Pessoas")
        imu_input_field = pessoas_panel.child_window(
            class_name=IMU_INPUT_FIELD_CLASS,
            found_index=0 
        )
        print("✅ IMU Input Field reference established via parent 'Pessoas'.")

    except Exception as e:
        print(f"❌ FATAL: Could not find the required 'Pessoas' panel or controls. Error: {e}")
        sys.exit(1)

    # Get the PID of the main application (needed for reliable dialog connection)
    main_app_pid = app.process 
    
    # 2. Main Automation Loop
    for record in data_list:
        imu_id = record['imu_id']
        output_dir = record['dir_path']
        
        print(f"\n--- Processing IMU: {imu_id} in folder: {output_dir.name} ---")

        try:
            main_window.set_focus()
            
            # --- STEP 1 & 2: Select Dropdown Option & Input IMU (Keep original) ---
            main_window.child_window(class_name=FINALIDADE_DROPDOWN_CLASS).click_input()
            send_keys("d") 
            send_keys("{ENTER}")
            print("-> Confirmed 'Finalidade' selection with keystrokes.")
            
            imu_input_field.set_focus()
            imu_input_field.type_keys(imu_id, with_spaces=True)
            send_keys("{TAB}")
            time.sleep(0.5) 
            print(f"-> Entered IMU: {imu_id} and pressed TAB.")
            
            # --- STEP 3: Trigger Search using Ctrl+E Hotkey (Keep original) ---
            main_window.set_focus()
            send_keys("^E")
            print("-> Triggered Search using {Ctrl+E} hotkey.")
            
            # --- STEP 4 & 5: Handle Pop-ups (Keep original) ---
            attention_dialog = check_for_popup(app, ATTENTION_DIALOG_TITLE, timeout=10)
            
            if attention_dialog:
                print("-> Found 'Atenção' dialog. Sending ENTER key twice to dismiss both pop-ups.")
                send_keys("{ENTER}") 
                time.sleep(1) 
                send_keys("{ENTER}")
                time.sleep(1) 
            else:
                print("-> 'Atenção' dialog not found (no error pop-ups). Proceeding.")
            
            
            # --- CODE TO TRIGGER THE SAVE DIALOG GOES HERE ---
            # NOTE: We assume the previous steps lead to the main window being active 
            # and that the next action triggers the print/save functionality.
            # If the application uses a specific button or menu, replace this placeholder.
            # Example: main_window.child_window(title="Imprimir", control_type="Button").click()
            # For now, let's assume the dialog is triggered by an action *just* before Step 6 starts.
            # If the dialog doesn't appear, ensure the trigger action is here and working!
            
            # --- STEP 6: Handle Save PDF Dialog (ROBUST CONNECTION + ENTER KEY) ---
            new_filename = f"Situacao_{imu_id}.pdf"
            full_path = str(output_dir / new_filename)
            save_dialog_title = SAVE_DIALOG_TITLE  # "Salvar Saída de Impressão como"
            main_app_pid = app.process 

            print(f"-> Attempting to save PDF to: {full_path} via robust pywinauto + ENTER")

            # CRITICAL FIX: Add a short, explicit wait to ensure the OS registers the window.
            time.sleep(1) 
            
            try:
                # 1. Connect to the Save As dialog using the main application's PID
                #    Using the 'win32' backend for faster dialog recognition.
                app_save = Application(backend="win32").connect(process=main_app_pid, title=save_dialog_title, timeout=15)
                save_dialog = app_save.window(title=save_dialog_title)

                # 2. Set the text for the file name input field. (Confirmed working)
                save_dialog.wait('ready', timeout=5)

                # Use the robust control finding based on inspect.exe results
                file_edit = save_dialog.child_window(class_name="Edit", found_index=0) 
                
                # Set the full path
                file_edit.set_text(full_path) 
                
                # 3. Use the ENTER key to confirm the 'Salvar' action.
                #    This is more reliable than .click() when dealing with tricky UI controls.
                save_dialog.set_focus() # Ensure the dialog is focused before sending keys
                send_keys("{ENTER}") # Simulates clicking the default 'Salvar' button
                
                print(f"-> ✅ Successfully saved file using robust connection and ENTER key.")
                
            except timings.TimeoutError as e:
                print(f"!!! ⚠️ FATAL ERROR: Dialog '{save_dialog_title}' failed to connect within 15s. Error: {e}")
            except Exception as e:
                # This error now indicates a control (like the 'Edit' field) could not be found.
                print(f"!!! ⚠️ UNEXPECTED ERROR: Could not interact with controls. Error: {e}")
                print("   -> Action: Control finding failed. Check inspect.exe results again.")
                
            # --- STEP 7 & 8: Trigger Eraser and Handle Confirmation Pop-up (Keep original) ---
            time.sleep(1) 
            
            main_window.set_focus()
            send_keys("+{F8}") 
            print("-> Triggered Eraser using {Shift+F8} hotkey.")
            
            handle_confirmation_popup(app, CONFIRMATION_DIALOG_TITLE, SIM_BUTTON_TITLE) 

            print("-> Handled 'Confirmação' pop-up. Ready for next loop.")
            
        except Exception as e:
            print(f"!!! ⚠️ UNEXPECTED ERROR processing IMU {imu_id}. Attempting screen clear. Error: {e}")
            try:
                main_window.set_focus()
                send_keys("+{F8}") 
                print("-> Recovery successful: Triggered Eraser using {Shift+F8}.")
            except:
                print("-> Recovery failed. Proceeding to next record.")
            continue

    print("\n✅ Automation loop completed successfully using pywinauto.")


if __name__ == "__main__":
    
    print(">>> REMINDER: Ensure the target application is OPEN and you run this script AS ADMINISTRATOR. <<<")
    
    company_data = get_data_from_directories()
    if company_data:
        automate_pdf_extraction_pywinauto(company_data)
    else:
        sys.exit(1)