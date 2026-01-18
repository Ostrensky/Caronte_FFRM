import os
import re
import csv
import datetime

def generate_report(root_path, output_filename):
    # 1. Define Regex Patterns
    # Matches directory: "123_Name" -> Captures 123
    dir_pattern = re.compile(r'^(\d+)_.+$')
    
    # Matches file: "456_DAM.pdf" -> Captures 456
    file_pattern = re.compile(r'^(\d+)_DAM\.pdf$', re.IGNORECASE)

    rows = []

    print(f"Scanning directory: {root_path}...")

    # 2. Walk through the directory tree
    for current_root, dirs, files in os.walk(root_path):
        folder_name = os.path.basename(current_root)
        
        # Check if the current folder matches the "<number>_name" pattern
        dir_match = dir_pattern.match(folder_name)
        
        if dir_match:
            first_number = dir_match.group(1)
            
            # Search for the PDF file inside this folder
            for filename in files:
                file_match = file_pattern.match(filename)
                
                if file_match:
                    second_number = file_match.group(1)
                    full_file_path = os.path.join(current_root, filename)
                    
                    # 3. Get Modified Date
                    try:
                        timestamp = os.path.getmtime(full_file_path)
                        mod_date = datetime.datetime.fromtimestamp(timestamp)
                        formatted_date = mod_date.strftime('%d-%m-%Y')
                        
                        # Add to our list
                        rows.append([formatted_date, first_number, second_number])
                        print(f"Found: {filename} in {folder_name}")
                        
                    except OSError as e:
                        print(f"Error reading {filename}: {e}")

    # 4. Write to CSV (Excel compatible)
    if rows:
        with open(output_filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Header
            writer.writerow(['Modified Date', 'First Number (Dir)', 'Second Number (File)'])
            # Data
            writer.writerows(rows)
        print(f"\nSuccess! File saved as '{output_filename}' with {len(rows)} entries.")
    else:
        print("\nNo matching files found.")

if __name__ == "__main__":
    # --- CONFIGURATION ---
    # '.' means the current folder where this script is saved.
    # You can replace '.' with a full path like r'C:\Users\Name\Documents\Files'
    root_directory = 'C:/Users/vostrensky/Documents/IDD/OIDD_XX_25' 
    output_csv = 'file_report.csv'
    # ---------------------

    generate_report(root_directory, output_csv)