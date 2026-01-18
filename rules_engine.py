# rules_engine.py
import pandas as pd
import numpy as np
import unicodedata
from datetime import datetime
from pandas.tseries.offsets import MonthEnd, DateOffset

# ... (Keep build_aliquotas_lookup and existing helper functions if needed for legacy support, 
# but the new logic relies on the function below) ...

def build_aliquotas_lookup(aliquotas_df):
    """
    Constructs a lookup dictionary. 
    (Kept for compatibility, though vectorized approach uses the DF directly/transformed)
    """
    lookup = {}
    if aliquotas_df.empty:
        return lookup

    # Sort to ensure default (first occurrence) logic is consistent
    # We prioritize rows with descriptions if your logic requires it, 
    # but based on previous code, the first one found was default.
    records = aliquotas_df.to_dict('records')

    for row in records:
        code = str(row['Codigo']).strip()
        desc = row.get('Descrição da Atividade')
        
        if code not in lookup:
            lookup[code] = {'default': None, 'by_desc': {}}
        
        if lookup[code]['default'] is None:
            lookup[code]['default'] = row
            
        if desc:
            lookup[code]['by_desc'][desc] = row
            
    return lookup

def process_invoices_vectorized(df, aliquotas_lookup, today=None, idd_mode=False):
    """
    Vectorized implementation of the rules engine.
    Drastically faster than row-by-row iteration.
    """
    if df.empty:
        return df

    # --- 1. PREPARATION & MAPPING (Data Enrichment) ---
    if today is None:
        today = pd.to_datetime(datetime.now().date())

    # Ensure types
    df['CÓDIGO DA ATIVIDADE'] = df['CÓDIGO DA ATIVIDADE'].astype(str).str.strip()
    df['DATA EMISSÃO'] = pd.to_datetime(df['DATA EMISSÃO'], errors='coerce')
    df['ALÍQUOTA'] = pd.to_numeric(df['ALÍQUOTA'], errors='coerce').fillna(0.0)
    df['VALOR DEDUÇÃO'] = pd.to_numeric(df['VALOR DEDUÇÃO'], errors='coerce').fillna(0.0)
    
    # Create Reference Columns (Default values based on Activity Code)
    # We flatten the lookup for vector mapping. 
    # Logic: 'default' key from lookup is the base.
    
    # Build simple dicts for mapping
    ref_rate = {}
    ref_desc = {}
    ref_deducao = {}
    ref_retencao = {}
    ref_local = {}
    ref_isencao = {}
    ref_imunidade = {}

    for code, data in aliquotas_lookup.items():
        def_row = data.get('default', {})
        if not def_row: continue
        ref_rate[code] = float(def_row.get('Aliquota', 0.0))
        ref_desc[code] = str(def_row.get('Descrição da Atividade', ''))
        ref_deducao[code] = str(def_row.get('Dedução', 'Não Habilita')).strip().lower()
        ref_retencao[code] = str(def_row.get('Retencao', 'Não Habilita')).strip().lower()
        ref_local[code] = str(def_row.get('Local', '')).strip().lower()
        ref_isencao[code] = str(def_row.get('Isencao', '')).strip().lower()
        ref_imunidade[code] = str(def_row.get('Imunidade', '')).strip().lower()

    # Map defaults
    df['ref_correct_rate'] = df['CÓDIGO DA ATIVIDADE'].map(ref_rate).fillna(0.0)
    df['ref_activity_desc'] = df['CÓDIGO DA ATIVIDADE'].map(ref_desc).fillna('N/A')
    df['ref_deducao'] = df['CÓDIGO DA ATIVIDADE'].map(ref_deducao).fillna('não habilita')
    df['ref_retencao'] = df['CÓDIGO DA ATIVIDADE'].map(ref_retencao).fillna('não habilita')
    df['ref_local'] = df['CÓDIGO DA ATIVIDADE'].map(ref_local).fillna('')
    df['ref_isencao'] = df['CÓDIGO DA ATIVIDADE'].map(ref_isencao).fillna('')
    df['ref_imunidade'] = df['CÓDIGO DA ATIVIDADE'].map(ref_imunidade).fillna('')

    # Handle "Description Overrides" (Specific Description matching)
    # Since vector mapping dictionary overrides is hard, we iterate only the codes that HAVE overrides.
    for code, data in aliquotas_lookup.items():
        if data.get('by_desc'):
            for specific_desc, row_data in data['by_desc'].items():
                # Boolean mask where Code AND Desc match
                mask_override = (df['CÓDIGO DA ATIVIDADE'] == code) & (df['activity_desc'] == specific_desc)
                if mask_override.any():
                    df.loc[mask_override, 'ref_correct_rate'] = float(row_data.get('Aliquota', 0.0))
                    df.loc[mask_override, 'ref_deducao'] = str(row_data.get('Dedução', 'Não Habilita')).lower()
                    df.loc[mask_override, 'ref_local'] = str(row_data.get('Local', '')).lower()
                    # ... update other refs if needed ...

    # Helper columns
    df['norm_natureza'] = df['NATUREZA DA OPERAÇÃO'].astype(str).str.strip().str.lower().str.replace(" ", "", regex=False)
    # Remove accents from nature for easier comparison
    df['norm_natureza_nfd'] = df['norm_natureza'].apply(lambda x: unicodedata.normalize('NFD', x).encode('ascii', 'ignore').decode("utf-8"))
    
    df['is_paid'] = df['PAGAMENTO'].astype(str).str.strip().str.lower().isin(['sim', 'idd'])
    df['regime_normal'] = df['REGIME DE TRIBUTAÇÃO'].astype(str).str.strip() == 'Contribuinte sujeito a tributação normal'
    
    # --- 2. BOOLEAN MASKS FOR RULES (The Engine) ---

    # Initialize rule masks (False = No Infraction)
    m_regime = pd.Series(False, index=df.index)
    m_aliquota = pd.Series(False, index=df.index)
    m_isencao_imu = pd.Series(False, index=df.index)
    m_natureza_local = pd.Series(False, index=df.index)
    m_deducao = pd.Series(False, index=df.index)
    m_retencao = pd.Series(False, index=df.index)
    
    # Logic: Skip if manual status 'Local_Tomador' or 'Decadente_Pago'
    mask_active = (
        (df['status_manual'] != 'Local_Tomador') & 
        (df['status_legal'] != 'Decadente_Pago')
    )

    # 2.1. Rule: Regime Incorreto
    if not idd_mode:
        m_regime = mask_active & (~df['regime_normal'])

    # 2.2. Rule: Alíquota Incorreta
    # Logic: If Correct > Declared (with float tolerance).
    if not idd_mode:
        # np.isclose returns boolean array. We negate it to ensure they aren't basically equal.
        # Check: correct > declared AND not close
        m_aliquota = mask_active & (
            (df['ref_correct_rate'] > df['ALÍQUOTA']) & 
            (~np.isclose(df['ref_correct_rate'], df['ALÍQUOTA']))
        )

    # 2.3. Rule: Isenção/Imunidade Indevida
    if not idd_mode:
        is_isencao = df['norm_natureza_nfd'].str.startswith('ise')
        is_imunidade = df['norm_natureza_nfd'].str.startswith('imu')
        
        m_isencao_imu = mask_active & (
            (is_isencao & (df['ref_isencao'] == 'não habilita')) |
            (is_imunidade & (df['ref_imunidade'] == 'não habilita'))
        )

    # 2.4. Rule: Natureza da Operação Incompatível (Local)
    # Special Exclusion: If 'tributacaoforamunicipio' and ref_local == 'tomador', it is VALID (stops other checks in original logic).
    # We implement this as a mask that forces compliance.
    is_fora_mun = df['norm_natureza'] == 'tributacaoforamunicipio'
    is_tomador_ref = df['ref_local'] == 'tomador'
    mask_tomador_ok = is_fora_mun & is_tomador_ref
    
    # Apply this exclusion to all currently active masks
    m_regime &= ~mask_tomador_ok
    m_aliquota &= ~mask_tomador_ok
    m_isencao_imu &= ~mask_tomador_ok

    if not idd_mode:
        # Actual Rule: If 'tributacaoforamunicipio' but ref says 'prestador' (implied by not tomador? logic check)
        # Original: if 'tributacaoforamunicipio' and permission == 'prestador' -> Error
        is_prestador_ref = df['ref_local'] == 'prestador'
        m_natureza_local = mask_active & (~mask_tomador_ok) & (is_fora_mun & is_prestador_ref)

    # 2.5. Rule: Dedução Indevida
    if not idd_mode:
        m_deducao = mask_active & (~mask_tomador_ok) & (
            (df['VALOR DEDUÇÃO'] > 0) & (df['ref_deducao'] == 'não habilita')
        )

    # 2.6. Rule: Retenção na Fonte a Verificar
    if not idd_mode:
        iss_retido_sim = df['ISS RETIDO'].astype(str).str.strip().str.lower() == 'sim'
        m_retencao = mask_active & (~mask_tomador_ok) & iss_retido_sim

    # --- 3. COMBINE INFRACTIONS & HANDLE DECADENCE (ART 173) ---
    
    # Identify rows that have ANY infraction
    df['has_infraction'] = (m_regime | m_aliquota | m_isencao_imu | m_natureza_local | m_deducao | m_retencao)

    # Decadence Calculation (Art 173): First day of invoice year + 6 years (logic from original: year + 6)
    # Original used: invoice_date.replace(day=1, month=1) -> year start
    dt_year_start = df['DATA EMISSÃO'].dt.to_period('Y').dt.to_timestamp() 
    cutoff_173 = dt_year_start + pd.DateOffset(years=6)
    
    m_decadente_173 = df['has_infraction'] & (today >= cutoff_173)

    # If decadent, we suppress the specific infractions in the output, just marking "Decadente"
    df.loc[m_decadente_173, 'status_legal'] = 'Decadente'
    # Clear infractions for decadent rows to avoid double reporting
    df.loc[m_decadente_173, 'has_infraction'] = False 
    
    # --- 4. IDD & PRESCRIPTION (ART 174) ---
    # Logic: Only check if NOT has_infraction AND NOT paid
    
    mask_check_idd = mask_active & (~df['has_infraction']) & (~df['is_paid']) & (~m_decadente_173) & (~mask_tomador_ok)

    # Prescription Date Calculation (Art 174)
    # Original: last_day_of_prev_month + 20 days + 5 years
    # Logic for "last day of prev month": (Date - MonthEnd) gets to prev month end?
    # safest: go to MonthBegin, subtract 1 day.
    prev_month_end = df['DATA EMISSÃO'].dt.to_period('M').dt.to_timestamp() - pd.Timedelta(days=1)
    due_date = prev_month_end + pd.Timedelta(days=20)
    cutoff_174 = due_date + pd.DateOffset(years=5)

    m_prescrito = mask_check_idd & (today >= cutoff_174)
    df.loc[m_prescrito, 'status_legal'] = 'Prescrito'
    
    # IDD Check (If not prescribed)
    mask_idd_candidates = mask_check_idd & (~m_prescrito)
    
    # Criteria for IDD
    # 1. Aliquota != 0
    # 2. Natureza == 'TributacaoMunicipio' (normalized check)
    # 3. ISS Retido == 'Não'
    # 4. Regime == Normal
    
    is_tributacao_mun = df['norm_natureza'].str.contains('tributacaomunicipio', na=False)
    iss_retido_nao = df['ISS RETIDO'].astype(str).str.strip().str.lower() == 'não'
    
    m_idd_nao_pago = mask_idd_candidates & (
        (df['ALÍQUOTA'] != 0) &
        is_tributacao_mun &
        iss_retido_nao &
        df['regime_normal']
    )

    # --- 5. BUILD OUTPUT COLUMNS ---

    # We need to construct the 'broken_rule_details' list column.
    # To do this effectively in pandas, we'll create string columns for each error and join them.
    
    df['err_regime'] = np.where(m_regime & (~m_decadente_173), 'Regime incorreto', '')
    
    # For Aliquota, we need dynamic text
    # "Alíquota Incorreta (Declarada: X%, Correta: Y%, Pagamento: Z)"
    mask_aliq_active = m_aliquota & (~m_decadente_173)
    
    # 2. Define specific conditions for Paid vs Unpaid
    cond_aliq_paid = mask_aliq_active & df['is_paid']
    cond_aliq_unpaid = mask_aliq_active & (~df['is_paid'])
    
    # 3. Create the messages
    # Message for PAID invoices (Preserve declared rate)
    msg_aliq_paid = (
        'Alíquota Incorreta (Declarada: ' + df['ALÍQUOTA'].map('{:.2f}'.format) + 
        '%, Correta: ' + df['ref_correct_rate'].map('{:.2f}'.format) + 
        '%, Pagamento: ' + df['PAGAMENTO'].astype(str) + ')'
    )
    
    # Message for UNPAID invoices (Generic string to merge groups)
    msg_aliq_unpaid = (
        'Alíquota Incorreta (Não Pago - Correta: ' + 
        df['ref_correct_rate'].map('{:.2f}'.format) + '%)'
    )
    
    # 4. Apply selection
    df['err_aliquota'] = np.select(
        [cond_aliq_paid, cond_aliq_unpaid],
        [msg_aliq_paid, msg_aliq_unpaid],
        default=''
    )
    
    df['err_isencao'] = np.where(m_isencao_imu & (~m_decadente_173), 'Isenção/Imunidade Indevida', '')
    df['err_nat_local'] = np.where(m_natureza_local & (~m_decadente_173), 'Natureza da Operação Incompatível', '')
    df['err_deducao'] = np.where(m_deducao & (~m_decadente_173), 'Dedução indevida', '')
    df['err_retencao'] = np.where(m_retencao & (~m_decadente_173), 'Retenção na Fonte (Verificar)', '')
    df['err_idd'] = np.where(m_idd_nao_pago, 'IDD (Não Pago)', '')

    # Consolidate 'correct_rate' and 'activity_desc' into the DF as required by main.py
    # (Already mapped to 'ref_correct_rate', 'ref_activity_desc')
    df['correct_rate'] = df['ref_correct_rate']

    
    
    df['activity_desc'] = df['ref_activity_desc']
    
    # --- CONSTRUCT THE LIST OF ERRORS ---
    # This part is effectively the "gathering" phase.
    
    # Define columns to aggregate
    err_cols = ['err_regime', 'err_aliquota', 'err_isencao', 'err_nat_local', 'err_deducao', 'err_retencao', 'err_idd']
    
    # Function to join non-empty strings into a list
    # While apply is technically a loop, doing it on specific text columns is much faster than the full logic engine.
    def gather_errors(row):
        return [row[c] for c in err_cols if row[c]]

    df['broken_rule_details'] = df[err_cols].apply(gather_errors, axis=1)
    
    # Determine Primary Infraction Group (First item in list or 'compliant')
    df['primary_infraction_group'] = df['broken_rule_details'].apply(lambda x: x[0] if x else 'compliant')

    # Cleanup temporary columns
    drop_cols = ['ref_rate', 'ref_desc', 'norm_natureza', 'norm_natureza_nfd', 
                 'is_paid', 'regime_normal', 'has_infraction'] + err_cols
    # Only drop what we created to avoid errors if cols didn't exist
    cols_to_drop = [c for c in drop_cols if c in df.columns]
    # We keep ref_correct_rate etc as they might be useful
    
    return df