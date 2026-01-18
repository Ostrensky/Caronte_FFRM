import pandas as pd
import keyboard
import sys
import pyperclip # Used to copy data to the clipboard

# --- CONFIGURATION ---
EXCEL_PATH = r'C:\Users\vostrensky\Documents\IDD\OIDD_XX_25\modelo_ficheiro_mestre.xlsx'
HOTKEY = 'f9'  # Press F9 to copy the next number to the clipboard
# --- END CONFIGURATION ---

# --- Global State ---
imu_data = [] # List to hold all IMU numbers
data_index = 0 # Current position in the list

# --- Setup Function ---

def setup_data(excel_path):
    """Reads Excel data and prepares the list of padded IMU numbers."""
    try:
        df = pd.read_excel(excel_path)
        imu_list_raw = df['imu'].astype(str).tolist()
        imu_list = [imu_id.strip().zfill(9) for imu_id in imu_list_raw]

        print(f"✅ Loaded {len(imu_list)} IMU numbers from Excel.")
        return imu_list

    except Exception as e:
        print(f"❌ FATAL ERROR during setup: {e}")
        sys.exit(1)

# --- Core Clipboard Logic ---

def copy_next_imu():
    """
    Function executed when the hotkey (F9) is pressed.
    It copies the next number to the clipboard ONLY.
    """
    global data_index
    global imu_data

    if not imu_data:
        print("Data not loaded. Check setup.")
        return

    if data_index >= len(imu_data):
        print("\n🚨 All IMU numbers processed. Press F10 to exit.")
        return

    imu_id = imu_data[data_index]
    
    # --- CRITICAL CLIPBOARD ACTION ---
    pyperclip.copy(imu_id)
    # --- END CRITICAL ACTION ---
    
    print(f"Copied IMU #{data_index + 1}/{len(imu_data)}: {imu_id} to clipboard.")

    # Move to the next record
    data_index += 1
    

# --- Main Execution ---

if __name__ == "__main__":
    imu_data = setup_data(EXCEL_PATH)
    
    if not imu_data:
        sys.exit(1)

    # Setup the hotkey listener for F9 to trigger the copy action
    keyboard.add_hotkey(HOTKEY, copy_next_imu)
    
    # Setup an exit hotkey for F10
    keyboard.add_hotkey('f10', lambda: (print("\n👋 Exiting script..."), keyboard.unhook_all(), sys.exit(0)))

    print("\n" + "="*70)
    print("🚀 SCRIPT RUNNING IN BACKGROUND (Copy-Only Mode)")
    print(f"   1. **Press {HOTKEY.upper()}** to copy the next number.")
    print(f"   2. **Manually press CTRL+V** in your application to paste it.")
    print(f"   3. **Press F10** at any time to exit the script.")
    print("="*70)
    
    # Keep the script running and listening for the hotkey
    keyboard.wait('f10')