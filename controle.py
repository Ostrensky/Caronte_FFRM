import pandas as pd
import io
from datetime import datetime
import os

# ---------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------
# You can change this to your local path: "Z:/Controle_FFRM2/Operação_IDD/Operação_IDD_2021-Fase I.xlsx"
# I kept the uploaded filename here so it runs in this browser preview.
INPUT_FILE = "Z:/Controle_FFRM2/Operação_IDD/Operação_IDD_2021-Fase I.xlsx"

# Auditor Name (Case sensitive, must match exactly what is in the CSV/Excel)
TARGET_AUDITOR = "Thais Coimbra Nina" 

# Date Range for filtering (YYYY-MM-DD)
START_DATE = "2025-01-01"
END_DATE = "2025-12-31"

# ---------------------------------------------------------
# 1. HELPER FUNCTIONS
# ---------------------------------------------------------
def format_date_pt_br(date_obj):
    """
    Converts datetime back to '24 de novembro de 2025' format
    """
    if pd.isna(date_obj):
        return ""
    
    months = {
        1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril',
        5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
        9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
    }
    
    try:
        d = pd.to_datetime(date_obj)
        return f"{d.day} de {months[d.month]} de {d.year}"
    except:
        return ""

def determine_result(row):
    """
    Logic to determine the 'Resultado' field based on other columns.
    """
    if pd.notna(row.get('nº Protocolo')) and str(row.get('nº Protocolo')).strip() != '':
        return "Protocolo concluído"
    elif pd.notna(row.get('Observação')):
        return "Em análise / Com pendência"
    else:
        return ""

# ---------------------------------------------------------
# 2. PROCESSING LOGIC
# ---------------------------------------------------------
def process_audit_data():
    print(f"--- Loading file: {INPUT_FILE} ---")
    
    df = None
    
    # 1. Load Data (Smart Detection)
    try:
        if INPUT_FILE.lower().endswith('.xlsx'):
            print("📂 Detected Excel file. Reading with read_excel...")
            # Reads the Excel file. 
            # Note: If there are multiple sheets, this reads the first one by default.
            df = pd.read_excel(INPUT_FILE)
        else:
            print("📄 Detected CSV/Text file. Reading with read_csv...")
            # Added encoding='latin1' to handle special characters on Windows
            # Added sep=None and engine='python' to auto-detect ; or , separators
            df = pd.read_csv(INPUT_FILE, encoding='latin1', sep=None, engine='python')
            
        # 2. Header Search Logic
        # Sometimes headers are not on row 1. This searches for the 'Auditor' column.
        if 'Auditor' not in df.columns:
            print("🔍 Header 'Auditor' not found in first row. Searching subsequent rows...")
            header_found = False
            # Check first 20 rows for the header
            for i in range(min(20, len(df))):
                # Convert row to string to search for 'Auditor' safely
                row_values = df.iloc[i].astype(str).values
                if 'Auditor' in row_values:
                    print(f"✅ Found headers at row {i}")
                    
                    # Set the column names to this row's values
                    df.columns = df.iloc[i]
                    
                    # Remove the rows above and the header row itself from the data
                    df = df.iloc[i+1:].reset_index(drop=True)
                    header_found = True
                    break
            
            if not header_found:
                print("⚠️ Warning: Could not find 'Auditor' column in the first 20 rows.")
                print(f"Columns found: {df.columns.tolist()}")
                return

    except FileNotFoundError:
        print(f"❌ Error: File '{INPUT_FILE}' not found.")
        return
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return

    # 3. Filter Data
    # Convert the date column. Based on snippet, 'Distribuição' looks like the main date.
    if 'Distribuição' in df.columns:
        date_col = 'Distribuição'
    elif 'Data Limite' in df.columns:
        date_col = 'Data Limite'
    else:
        print("❌ Error: Could not find a Date column ('Distribuição' or 'Data Limite')")
        return

    print("Converting dates...")
    df['Date_Obj'] = pd.to_datetime(df[date_col], errors='coerce')
    
    start = pd.to_datetime(START_DATE)
    end = pd.to_datetime(END_DATE)
    
    # Filter
    mask = (
        (df['Auditor'] == TARGET_AUDITOR) & 
        (df['Date_Obj'] >= start) & 
        (df['Date_Obj'] <= end)
    )
    
    filtered_df = df[mask].copy()
    
    if filtered_df.empty:
        print(f"⚠️ No records found for auditor '{TARGET_AUDITOR}' in range {START_DATE} to {END_DATE}")
        # Print available auditors to help debug
        try:
            unique_auditors = df['Auditor'].dropna().unique()
            print(f"Available Auditors in file: {unique_auditors}")
        except:
            pass
        return

    # 4. Map Columns to Target Format
    output_df = pd.DataFrame()

    # Data
    output_df['Data'] = filtered_df['Date_Obj'].apply(format_date_pt_br)

    # Atividade Realizada (Hardcoded or mapped)
    output_df['Atividade Realizada'] = "Operação IDD"

    # Inscrição Municipal... (Mapping from IMU or CNPJ)
    # Check if 'IMU' exists, otherwise try 'Inscrição Municipal'
    col_imu = 'IMU' if 'IMU' in filtered_df.columns else 'Inscrição Municipal'
    if col_imu in filtered_df.columns:
        output_df['Inscrição Municipal/CNPJ/CPF ou Nº do(s) Alvará(s) - CVCO'] = filtered_df[col_imu]
    else:
        output_df['Inscrição Municipal/CNPJ/CPF ou Nº do(s) Alvará(s) - CVCO'] = ""

    # Verificações... (Mapping from Observação)
    output_df['Verificações e Análises Realizadas'] = filtered_df['Observação'].fillna("")

    # Resultado (Derived logic)
    output_df['Resultado'] = filtered_df.apply(determine_result, axis=1)

    # Nº Processo... (Mapping from nº Protocolo)
    output_df['Nº Processo ou Nº Certidão - CVCO'] = filtered_df['nº Protocolo'].fillna("")

    # Nº DAM... (Mapping from Nº IDD)
    output_df['Nº DAM / IDD / AI / Denúncia'] = filtered_df['Nº IDD'].fillna("")

    # Valor Original... (Mapping from Valor ISS Original)
    col_valor_orig = 'Valor ISS Original' if 'Valor ISS Original' in filtered_df.columns else 'ISS PREVISTO'
    if col_valor_orig in filtered_df.columns:
        output_df['Valor Original do ISS'] = filtered_df[col_valor_orig].fillna(0)
    else:
        output_df['Valor Original do ISS'] = 0

    # Valor Corrigido... (Mapping from Valor ISS Atualizado)
    col_valor_corr = 'Valor ISS Atualizado'
    if col_valor_corr in filtered_df.columns:
        output_df['Valor Corrigido do ISS'] = filtered_df[col_valor_corr].fillna(0)
    else:
        output_df['Valor Corrigido do ISS'] = 0

    # Horas Trabalhadas (Not in source, leaving empty or calculating)
    output_df['Horas Trabalhadas'] = "" 
    
    # 5. Export
    # Generate the tab-separated string
    output_text = output_df.to_csv(sep='\t', index=False, float_format='%.2f')
    
    filename = f"Extracao_{TARGET_AUDITOR.replace(' ', '_')}.txt"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(output_text)
        
    print(f"\n✅ Success! Extracted {len(output_df)} rows.")
    print(f"✅ Saved file: {filename}")
    print("\n--- PREVIEW ---")
    print(output_df[['Data', 'Nº Processo ou Nº Certidão - CVCO', 'Valor Original do ISS']].head().to_string(index=False))

# ---------------------------------------------------------
# 3. EXECUTION
# ---------------------------------------------------------
if __name__ == "__main__":
    process_audit_data()