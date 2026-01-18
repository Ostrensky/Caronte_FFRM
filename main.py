# main.py

import os
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pandas.tseries.offsets import MonthEnd
import sys
import rules_engine
from data_loader import create_context_for_generation
# ❌ REMOVIDO: from report_generator import generate_simple_document, generate_report, convert_to_pdf
from description_analyzer import DescriptionAnalyzer
# ❌ REMOVIDO: from pdf_reports_generator import generate_detailed_pdfs
import traceback
import logging
from collections import defaultdict
from utils import resource_path
from app.pgdas_loader import _load_and_process_pgdas
from pandas.tseries.offsets import MonthEnd, DateOffset # ✅ Import DateOffset
import time
# --- Local Imports for Configuration ---
from app.config import (
    get_aliquotas_path, 
    # ❌ REMOVIDO: (imports de template, pois não são mais usados aqui)
    get_output_dir
)
# ---

# ... (função run_rules_analysis_from_files permanece igual) ...
def run_rules_analysis_from_files(master_filepath, invoices_filepath, company_cnpj, status_callback=None):
    if status_callback: status_callback.emit("Carregando e preparando faturas...")
    company_invoices_df = load_and_prepare_invoices(master_filepath, invoices_filepath, company_cnpj, status_callback)
    
    if company_invoices_df.empty:
        if status_callback: status_callback.emit("Nenhuma fatura encontrada para a empresa.")
        return {}, company_invoices_df
        
    if status_callback: status_callback.emit("Executando análise de regras...")
    infraction_groups, df_with_analysis = perform_rules_analysis(company_invoices_df)
    return infraction_groups, df_with_analysis

# ... (função load_activity_data permanece igual) ...
def load_activity_data():
    """
    Loads activities and groups them by code to handle non-unique codes.
    Returns a dictionary: {'Codigo': [('Descrição', Aliquota, 'Sinonimos'), ...]}
    """
    try:
        aliquotas_path = get_aliquotas_path()
        if not os.path.exists(aliquotas_path):
            print(f"Error loading activity data: File not found at {aliquotas_path}")
            return {}
            
        df = pd.read_excel(aliquotas_path)
        
        df['Codigo'] = (df['Codigo'].astype(str)
                        .str.replace(r'\D', '', regex=True)
                        .str.strip()
                        .str.pad(4, side='left', fillchar='0'))
        
        df['Aliquota'] = df['Aliquota'].astype(str).str.replace(',', '.', regex=False)
        df['Aliquota'] = pd.to_numeric(df['Aliquota'], errors='coerce').fillna(0.0)

        if 'SINONIMOS_CHAVE' not in df.columns:
            df['SINONIMOS_CHAVE'] = ""
            print("AVISO: Coluna 'SINONIMOS_CHAVE' não encontrada em aliquotas.xlsx. Análise de atividade será limitada.")
            
        df['SINONIMOS_CHAVE'] = df['SINONIMOS_CHAVE'].astype(str).fillna("")

        activity_data = defaultdict(list)
        for _, row in df.iterrows():
            code = row['Codigo']
            description = row['Descrição da Atividade']
            aliquota = row['Aliquota']
            synonyms = row['SINONIMOS_CHAVE']
            activity_data[code].append((description, aliquota, synonyms))
            
        return activity_data
    except Exception as e:
        print(f"Error loading activity data: {e}")
        return {}


# ... (função load_and_prepare_invoices permanece igual) ...
def _find_column_by_synonyms(columns, targets):
    """Helper to find a column name given a list of synonyms."""
    upper_cols = {str(c).upper().strip(): c for c in columns}
    for t in targets:
        if t.upper() in upper_cols:
            return upper_cols[t.upper()]
    return None

def load_and_prepare_invoices(master_filepath, invoices_filepath, company_cnpj, status_callback=None, auto_resolve_conflicts=True):
    emit = status_callback.emit if status_callback else print
    try:
        emit("🔄 Carregando e preparando dados das Notas Fiscais...")
        all_invoices_df = pd.read_excel(
            invoices_filepath, 
            skiprows=2, 
            engine='calamine' 
        )
        emit(f"   - {len(all_invoices_df)} linhas lidas inicialmente do ficheiro de notas.")

        # --- ENSURE ESSENTIAL COLUMNS EXIST ---
        required_cols = ['VALOR', 'VALOR DEDUÇÃO', 'ALÍQUOTA', 'DESCONTO INCONDICIONAL', 'DATA EMISSÃO', 'DT. CANCELAMENTO', 'NÚMERO']
        for col in required_cols:
            if col not in all_invoices_df.columns:
                all_invoices_df[col] = None

        # --- Date Conversion ---
        all_invoices_df['DATA EMISSÃO'] = pd.to_datetime(all_invoices_df['DATA EMISSÃO'], errors='coerce')
        all_invoices_df['DT. CANCELAMENTO'] = pd.to_datetime(all_invoices_df['DT. CANCELAMENTO'], errors='coerce')

        cancelled_mask = all_invoices_df['DT. CANCELAMENTO'].notna()
        if cancelled_mask.any():
            # Get list of numbers that are cancelled in at least one row
            cancelled_numbers = all_invoices_df.loc[cancelled_mask, 'NÚMERO'].unique()
            
            initial_len = len(all_invoices_df)
            
            # Remove ANY row that matches these numbers (even if that specific row has no cancel date)
            all_invoices_df = all_invoices_df[~all_invoices_df['NÚMERO'].isin(cancelled_numbers)]
            
            removed_count = initial_len - len(all_invoices_df)
            if removed_count > 0:
                 emit(f"   - 🚫 Filtro de Segurança: {len(cancelled_numbers)} notas canceladas removidas completamente (incluindo duplicatas 'zumbis').")

        # --- Numeric Conversion ---
        numeric_cols = ['VALOR', 'VALOR DEDUÇÃO', 'ALÍQUOTA', 'DESCONTO INCONDICIONAL']
        for col in numeric_cols:
            all_invoices_df[col] = pd.to_numeric(all_invoices_df[col], errors='coerce').fillna(0.0)

        # --- Filter by Company CNPJ ---
        company_invoices = all_invoices_df[all_invoices_df['CNPJ PRESTADOR'] == company_cnpj].copy()
        
        # Initialize status_manual immediately
        company_invoices['status_manual'] = None 
        company_invoices['_is_conflict'] = False 

        if company_invoices.empty:
             emit(f"⚠️ Nenhuma nota fiscal encontrada para o CNPJ {company_cnpj}.")
             if 'VALOR_ORIGINAL' not in company_invoices.columns: company_invoices['VALOR_ORIGINAL'] = 0.0
             if 'status_legal' not in company_invoices.columns: company_invoices['status_legal'] = 'OK'
             return company_invoices

        unique_col = 'NÚMERO' 
    
        if unique_col in company_invoices.columns and not company_invoices.empty:
            initial_count = len(company_invoices)

            # 1. Remove Exact Duplicates
            company_invoices.drop_duplicates(inplace=True)

            # --- 🚀 NEW SMART SORT LOGIC START ---
            
            # Helper 1: Find RPS Column
            rps_col = _find_column_by_synonyms(company_invoices.columns, ['Nº RPS', 'RPS', 'NO RPS', 'NUMERO RPS'])
            
            # Helper 2: Find Regime Column
            regime_col = _find_column_by_synonyms(company_invoices.columns, ['REGIME DE TRIBUTAÇÃO', 'REGIME', 'REGIME TRIBUTACAO'])

            # Helper 3: Find Description Column
            desc_synonyms = ['DISCRIMINAÇÃO DOS SERVIÇOS', 'DISCRIMINACAO DOS SERVICOS', 'DESCRIÇÃO', 'DESC']
            desc_col = _find_column_by_synonyms(company_invoices.columns, desc_synonyms)

            # Create sorting scores (0 = Good/Keep, 1 = Bad/Remove)
            
            # A. RPS Criteria: If RPS == 1, it is "Bad" (1). Else "Good" (0).
            if rps_col:
                # Convert to numeric safely, handle strings "1" or ints 1
                company_invoices['_tmp_rps_val'] = pd.to_numeric(company_invoices[rps_col], errors='coerce').fillna(0)
                company_invoices['_sort_rps_bad'] = np.where(company_invoices['_tmp_rps_val'] == 1, 1, 0)
                company_invoices.drop(columns=['_tmp_rps_val'], inplace=True)
            else:
                company_invoices['_sort_rps_bad'] = 0 # Neutral if column missing

            # B. Regime Criteria: If "Optante pelo Simples", it is "Bad" (1).
            if regime_col:
                company_invoices['_sort_regime_bad'] = company_invoices[regime_col].astype(str).apply(
                    lambda x: 1 if "Optante pelo Simples Nacional" in x else 0
                )
            else:
                company_invoices['_sort_regime_bad'] = 0

            # C. Description Length (Tie breaker)
            if desc_col:
                company_invoices['_desc_len'] = company_invoices[desc_col].astype(str).str.len()
            else:
                company_invoices['_desc_len'] = 0

            # D. Apply Sort
            # Priority:
            # 1. Number (Group together)
            # 2. RPS Badness (Ascending -> 0 comes before 1. So NON-RPS-1 stays at top)
            # 3. Regime Badness (Ascending -> 0 comes before 1. So NON-Simples stays at top)
            # 4. Aliquot (Descending -> Higher Aliquot stays at top)
            # 5. Value (Descending -> Higher Value stays at top)
            
            company_invoices.sort_values(
                by=[unique_col, '_sort_rps_bad', '_sort_regime_bad', 'ALÍQUOTA', 'VALOR', '_desc_len'], 
                ascending=[True, True, True, False, False, False], 
                inplace=True
            )
            
            # Cleanup temp sort columns
            company_invoices.drop(columns=['_sort_rps_bad', '_sort_regime_bad', '_desc_len'], inplace=True)
            
            # --- 🚀 NEW SMART SORT LOGIC END ---

            # 3. Handle Conflicts
            # duplicated(keep='first') marks duplicates except for the first occurrence.
            # Since we sorted the "Best" one to the top, keep='first' works perfectly.
            duplicates_mask = company_invoices.duplicated(subset=[unique_col], keep='first')
            all_conflicts_mask = company_invoices.duplicated(subset=[unique_col], keep=False)

            if all_conflicts_mask.any():
                if auto_resolve_conflicts:
                    # 'first' keeps the top sorted row (Best RPS, Best Regime, Best Aliquot)
                    company_invoices.drop_duplicates(subset=[unique_col], keep='first', inplace=True)
                    diff = initial_count - len(company_invoices)
                    if diff > 0: emit(f"   - 🧹 Auto-Resolução: {diff} duplicatas removidas (Prioridade: RPS!=1 > Regime!=Simples > Maior Alíquota).")
                    company_invoices['_is_conflict'] = False
                else:
                    # Mark all involved rows as conflicts so the user can choose in the Dialog
                    company_invoices['_is_conflict'] = all_conflicts_mask
                    emit(f"   - ⚠️ Detetados conflitos para revisão.")
            else:
                company_invoices['_is_conflict'] = False

        # --- Final Calcs ---
        company_invoices['VALOR_ORIGINAL'] = company_invoices['VALOR']
        company_invoices['VALOR'] = company_invoices['VALOR'] - company_invoices['DESCONTO INCONDICIONAL']

        if 'CÓDIGO DA ATIVIDADE' in company_invoices.columns:
            company_invoices['CÓDIGO DA ATIVIDADE'] = (company_invoices['CÓDIGO DA ATIVIDADE'].astype(str)
                                                    .str.replace(r'\D', '', regex=True)
                                                    .str.strip().str.pad(4, side='left', fillchar='0'))
            
        # Decadence Logic
        today = pd.to_datetime(datetime.now().date())
        if 'PAGAMENTO' not in company_invoices.columns: company_invoices['PAGAMENTO'] = 'Não'
        company_invoices['PAGAMENTO'].fillna('Não', inplace=True)
        is_paid_mask = company_invoices['PAGAMENTO'].str.strip().str.lower().isin(['sim', 'idd'])
        
        company_invoices['status_legal'] = 'OK'
        paid_invoices_mask = is_paid_mask & company_invoices['DATA EMISSÃO'].notna()
        if paid_invoices_mask.any():
            cutoff = company_invoices.loc[paid_invoices_mask, 'DATA EMISSÃO'] + MonthEnd(0) + DateOffset(years=5)
            mask_dec = today > cutoff
            company_invoices.loc[mask_dec[mask_dec].index, 'status_legal'] = 'Decadente_Pago'

        emit("✅ Preparação das notas concluída.")
        return company_invoices

    except FileNotFoundError:
        emit(f"❌ ERRO: Ficheiro não encontrado: '{invoices_filepath}'")
        raise
    except Exception as e:
        emit(f"❌ ERRO CRÍTICO: {e}")
        raise
      
# ... (função perform_description_analysis permanece igual) ...
def perform_description_analysis(company_invoices_df, status_callback=None):
    if company_invoices_df.empty:
        return company_invoices_df

    try:
        activity_data = load_activity_data()
        if not activity_data:
            raise FileNotFoundError("O ficheiro de alíquotas está vazio ou não foi encontrado.")
            
    except Exception as e:
        if status_callback:
            status_callback.emit(f"❌ Erro ao carregar ficheiro de alíquotas: {e}")
        return company_invoices_df # Can't proceed

    desc_map = {}
    rate_map = {}
    for code, data_list in activity_data.items():
        if data_list:
            desc_map[code] = data_list[0][0]
            rate_map[code] = data_list[0][1]

    if 'activity_desc' not in company_invoices_df:
        company_invoices_df['activity_desc'] = 'N/A'
    if 'correct_rate' not in company_invoices_df:
        company_invoices_df['correct_rate'] = np.nan
    
    company_invoices_df['activity_desc'] = company_invoices_df['CÓDIGO DA ATIVIDADE'].map(desc_map).fillna('N/A')
    company_invoices_df['correct_rate'] = company_invoices_df['CÓDIGO DA ATIVIDADE'].map(rate_map).fillna(0.0)

    analyzer = DescriptionAnalyzer()
    if status_callback:
        analyzer.progress.connect(status_callback)
    
    analyzed_df_results = analyzer.analyze_invoices(company_invoices_df, activity_data)
    
    ai_columns = [
        'location_alert', 'activity_alert'
    ]
    df_with_ai = company_invoices_df.copy()
    for col in ai_columns:
        if col in analyzed_df_results.columns:
            df_with_ai[col] = analyzed_df_results[col]

    return df_with_ai

# ... (função perform_rules_analysis permanece igual) ...
def perform_rules_analysis(company_invoices_df, idd_mode=False):
    """
    Revised to use Vectorized Rules Engine.
    Significantly faster than previous list-of-dicts iteration.
    """
    if company_invoices_df.empty:
        return {}, company_invoices_df

    try:
        aliquotas_path = get_aliquotas_path()
        if not os.path.exists(aliquotas_path):
            print(f"Error loading aliquotas reference: File not found at {aliquotas_path}")
            return {}, company_invoices_df

        aliquotas_df = pd.read_excel(aliquotas_path)
        aliquotas_df['Codigo'] = (aliquotas_df['Codigo'].astype(str)
                                  .str.replace(r'\D', '', regex=True)
                                  .str.strip()
                                  .str.pad(4, side='left', fillchar='0'))

        # Build lookup (Still needed for the vector engine's setup phase)
        aliquotas_lookup = rules_engine.build_aliquotas_lookup(aliquotas_df)

    except (FileNotFoundError, KeyError) as e:
        print(f"Error loading aliquotas reference: {e}")
        return {}, company_invoices_df

    # Prepare DataFrame for analysis
    invoices_for_analysis = company_invoices_df.copy()
    
    # Pre-populate Description Analysis columns if missing (required for some rule logic)
    if 'activity_desc' not in invoices_for_analysis:
        invoices_for_analysis['activity_desc'] = 'N/A'

    # 🚀 VECTORIZED EXECUTION 🚀
    # We pass the entire DataFrame to the rules engine
    analyzed_df = rules_engine.process_invoices_vectorized(
        invoices_for_analysis, 
        aliquotas_lookup, 
        idd_mode=idd_mode
    )

    # Group results
    violating_invoices = analyzed_df[analyzed_df['primary_infraction_group'] != 'compliant']
    infraction_groups = {key: group for key, group in violating_invoices.groupby('primary_infraction_group')}

    return infraction_groups, analyzed_df

# ... (função run_ai_preparation permanece igual) ...
def run_ai_preparation(master_filepath, invoices_filepath, company_cnpj, status_callback=None):
    if status_callback: status_callback.emit("Carregando e preparando faturas...")
    company_invoices_df = load_and_prepare_invoices(master_filepath, invoices_filepath, company_cnpj, status_callback)

    if company_invoices_df.empty:
        if status_callback: status_callback.emit("Nenhuma fatura encontrada para a empresa.")

    return company_invoices_df
        
    analyzed_df = perform_description_analysis(company_invoices_df, status_callback)
    return analyzed_df


# ❌❌❌ A FUNÇÃO 'generate_final_documents' FOI MOVIDA PARA 'app/generation_task.py' ❌❌❌