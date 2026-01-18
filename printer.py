import tkinter as tk
from tkinter import filedialog
import win32print  # For listing printers
import win32api    # For sending the print job
import time

def get_available_printers():
    """
    Returns a list of names of all installed printers.
    """
    printers = []
    # EnumPrinters(2) lists local and network printers
    try:
        printer_info = win32print.EnumPrinters(2)
        printers = [name for flags, desc, name, comment in printer_info]
        return printers
    except Exception as e:
        print(f"Error getting printers: {e}")
        print("Please ensure 'pywin32' is installed: pip install pywin32")
        return None

def choose_files():
    """
    Opens a GUI dialog to select one or more files.
    Returns a list of file paths.
    """
    # We create a root Tkinter window but hide it
    root = tk.Tk()
    root.withdraw()
    
    print("Opening file dialog to choose files...")
    # Open the file dialog
    file_paths = filedialog.askopenfilenames(
        title="Select files to print"
    )
    
    # askopenfilenames returns a tuple, convert to list
    return list(file_paths)

def choose_printer(printers):
    """
    Asks the user to select a printer from a numbered list in the console.
    Returns the name of the selected printer.
    """
    if not printers:
        print("No printers found on this system.")
        return None
        
    print("\n--- Available Printers ---")
    for i, printer_name in enumerate(printers):
        print(f"  [{i+1}] {printer_name}")
    
    # Check if the default printer is in the list
    try:
        default_printer = win32print.GetDefaultPrinter()
        if default_printer in printers:
            default_index = printers.index(default_printer) + 1
            print(f"\n(Default printer is: [{default_index}] {default_printer})")
        else:
            default_index = -1
    except Exception:
        default_index = -1

    while True:
        try:
            choice_str = input(f"Which printer do you want to use? (Enter 1-{len(printers)}): ")
            choice = int(choice_str)
            if 1 <= choice <= len(printers):
                selected_printer = printers[choice - 1]
                return selected_printer
            else:
                print(f"Invalid choice. Please enter a number between 1 and {len(printers)}.")
        except ValueError:
            print("Invalid input. Please enter a number.")
        except EOFError:
            print("\nSelection cancelled.")
            return None

def print_files(file_paths, printer_name):
    """
    Sends each file to the specified printer.
    - For .txt files, it sends raw data directly to the printer (silent).
    - For other files (PDF, DOCX, etc.), it uses the 'printto' shell command,
      which relies on the default application (e..g, Adobe Acrobat)
      to handle the printing.
    """
    print(f"\nSending {len(file_paths)} file(s) to '{printer_name}'...")
    
    # Check if the printer is valid before starting
    try:
        h_printer_check = win32print.OpenPrinter(printer_name)
        win32print.ClosePrinter(h_printer_check)
    except Exception as e:
        print(f"    ERROR: Could not open printer '{printer_name}'.")
        print(f"    Please ensure the printer is online. Error: {e}")
        return

    global shell_print_used
    shell_print_used = False # Track if we used the non-silent method

    for file_path in file_paths:
        # Get the file extension
        file_extension = ""
        try:
            file_extension = file_path.lower().split('.')[-1]
        except IndexError:
            pass # No file extension

        # --- OPTION 1: Raw printing for .txt files ---
        # This is silent and does not open any application.
        if file_extension == 'txt':
            try:
                print(f"  -> Printing '{file_path}' (RAW mode - silent)...")
                
                # Open the file in binary read mode
                with open(file_path, "rb") as f:
                    file_data = f.read()
                
                # Open a handle to the printer
                h_printer = win32print.OpenPrinter(printer_name)
                try:
                    # Start a new print job
                    # DocInfo Level 1: (docName, pOutputFile, pDataType)
                    # "RAW" means we are sending data the printer understands directly.
                    doc_info = (f"Printing {file_path}", None, "RAW")
                    h_job = win32print.StartDocPrinter(h_printer, 1, doc_info)
                    if h_job > 0:
                        try:
                            win32print.StartPagePrinter(h_printer)
                            win32print.WritePrinter(h_printer, file_data)
                            win32print.EndPagePrinter(h_printer)
                        finally:
                            win32print.EndDocPrinter(h_printer)
                    else:
                        print(f"    ERROR: Could not start print job for {file_path}.")
                finally:
                    # Close the printer handle
                    win32print.ClosePrinter(h_printer)
                
                print(f"  -> Sent '{file_path}' successfully.")
                
            except Exception as e:
                print(f"    ERROR: Could not print raw file {file_path}. {e}")
        
        # --- OPTION 2: Shell command for other file types ---
        # This relies on the default application (Acrobat, Word, etc.)
        # and will likely open a print dialog.
        else:
            shell_print_used = True # Mark that we used this method
            try:
                print(f"  -> Printing '{file_path}' (Shell mode)...")
                print(f"    (Asking default app to print. This may open a window.)")
                
                win32api.ShellExecute(
                    0,
                    "printto",
                    file_path,
                    f'"{printer_name}"',
                    ".",
                    0
                )
                # Add a longer delay for applications to open and spool
                time.sleep(4) 
                
            except Exception as e:
                print(f"    ERROR: Could not print file {file_path}. {e}")
                print("    Please ensure the file type has an associated program that can print.")

    print("\nAll print jobs have been sent.")

def main():
    # 1. Get list of printers
    printers = get_available_printers()
    if not printers:
        input("Press Enter to exit.")
        return

    # 2. Get list of files
    files_to_print = choose_files()
    if not files_to_print:
        print("No files selected. Exiting.")
        return

    # 3. Ask user to choose a printer
    selected_printer = choose_printer(printers)
    if not selected_printer:
        print("No printer selected. Exiting.")
        return
    
    # 4. Print the files
    print_files(files_to_print, selected_printer)
    
    # Add a final note if we had to use the Shell method
    if 'shell_print_used' in globals() and shell_print_used:
        print("\n--- PLEASE NOTE ---")
        print("An application (like Adobe Acrobat) may have opened.")
        print("This is a security feature of that app to prevent silent printing.")
        print("Only .txt files can be printed silently and automatically.")
        print("-------------------")
        
    input("\nProcess finished. Press Enter to exit.")

if __name__ == "__main__":
    main()
