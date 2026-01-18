import pyodbc
import pandas as pd
import os

# --- ⚙️ CONFIGURATION (EDIT HERE) ---
TARGET_IM = '6658120'   # Put the Inscrição Municipal here
TARGET_YEAR = 2023     # Put the specific year here
OUTPUT_FOLDER = 'dump_output'
# ------------------------------------

def extract_raw_table():
    # 1. Credentials (Matches your provided script)
    server = '172.19.210.187,1433'
    database = 'ISS_CURITIBA_RELATORIOS'
    username = 'vostrensky'
    password = 'T$&KzpUzUQ@yH4jchh' 

    conn_string = (
        f'DRIVER={{SQL Server}};' 
        f'SERVER={server};'
        f'DATABASE={database};'
        f'UID={username};'
        f'PWD={password};'
        f'TrustServerCertificate=yes;'
    )

    # 2. SQL Query - SELECT * (Raw Data)
    # We use COLLATE DATABASE_DEFAULT on the IM comparison to avoid collation conflicts
    sql_query = """
    SELECT *
    FROM [dbo].[ISSNFENota_Fiscal_Eletronica]
    WHERE 
        Num_IM_Prestador COLLATE DATABASE_DEFAULT = ?
        AND YEAR(Dta_Emissao_Nota_Fiscal) = ?
    ORDER BY 
        Dta_Emissao_Nota_Fiscal
    """

    print(f"--- STARTING RAW DUMP ---")
    print(f"TARGET: IM {TARGET_IM} | YEAR {TARGET_YEAR}")

    try:
        # 3. Connect and Execute
        print("🔌 Connecting to Database...")
        with pyodbc.connect(conn_string) as conn:
            
            print("⏳ Executing Query (SELECT *)...")
            # We pass parameters to prevent SQL injection and handle types correctly
            df = pd.read_sql(sql_query, conn, params=[str(TARGET_IM), int(TARGET_YEAR)])

            if df.empty:
                print(f"❌ No records found for IM {TARGET_IM} in {TARGET_YEAR}.")
                return

            print(f"✅ Found {len(df)} records.")

            # 4. Export to CSV (Best for raw inspection)
            if not os.path.exists(OUTPUT_FOLDER):
                os.makedirs(OUTPUT_FOLDER)

            filename = f"RAW_NFE_{TARGET_IM}_{TARGET_YEAR}.csv"
            filepath = os.path.join(OUTPUT_FOLDER, filename)

            print(f"💾 Saving to {filepath}...")
            
            # export using utf-8-sig to handle special characters in Excel correctly
            # index=False removes the pandas row numbers
            df.to_csv(filepath, index=False, sep=';', encoding='utf-8-sig')
            
            print("🚀 Done!")

    except pyodbc.Error as e:
        print(f"❌ Database Error: {e}")
    except Exception as e:
        print(f"❌ Unexpected Error: {e}")

if __name__ == "__main__":
    extract_raw_table()