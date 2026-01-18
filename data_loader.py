# --- FILE: data_loader.py ---

import os
import pandas as pd
from datetime import datetime
from document_parts import formatar_texto_multa, format_invoice_numbers, _format_currency_brl # ✅ Import currency formatterimport traceback
import logging
import locale
# Import the loader, as it's still used by the wizard (even if not here)
from app.pgdas_loader import _load_and_process_pgdas
from app.config import get_custom_general_texts
from statistics import mode # ✅ Import mode

def _load_and_process_dams(dam_filepath):
    """
    Loads the DAMs file (CSV) robustly, handling UTF-8-SIG (BOM), Latin1, 
    and both comma/semicolon separators.
    
    ✅ UPDATED: Now filters out DAMs that are not 'ISS Normal' or that have 
    specific invoice numbers associated (meaning they are not 'Avulsos').
    """
    df_dams = None
    
    attempts = [
        ('utf-8-sig', ','),
        ('utf-8-sig', ';'),
        ('latin1', ';'),
        ('latin1', ',')
    ]
    
    last_error = None
    
    for encoding, sep in attempts:
        try:
            df_dams = pd.read_csv(dam_filepath, sep=sep, decimal='.', thousands=',', encoding=encoding, on_bad_lines='skip')
            if len(df_dams.columns) > 1:
                break 
        except Exception as e:
            last_error = e
            continue
            
    if df_dams is None or len(df_dams.columns) < 2:
        try:
            df_dams = pd.read_csv(dam_filepath, sep=None, engine='python', encoding='utf-8-sig', on_bad_lines='skip')
        except:
            raise ValueError(f"Não foi possível ler o arquivo DAM. Erro: {last_error}")

    # --- CLEANUP HEADERS ---
    df_dams.columns = df_dams.columns.astype(str).str.strip().str.replace(r'^[^\w]+', '', regex=True)

    col_map = {
        'codigoVerificacao': 'codigoVerificacao', 'Código Verificação': 'codigoVerificacao', 'Codigo Verificacao': 'codigoVerificacao',
        'referenciaPagamento': 'referenciaPagamento', 'Competência': 'referenciaPagamento', 'Referência': 'referenciaPagamento',
        'receita': 'receita', 'Receita': 'receita',
        'totalRecolher': 'totalRecolher', 'Valor': 'totalRecolher', 'Valor Pago': 'totalRecolher',
        # ✅ Added mapping for filters
        'tributo': 'tributo', 'Tributo': 'tributo',
        'numerosDasNotas': 'numerosDasNotas', 'Notas': 'numerosDasNotas', 'Números das Notas': 'numerosDasNotas', 'Nota': 'numerosDasNotas'
    }
    
    new_cols = {}
    for col in df_dams.columns:
        c_str = str(col).strip()
        if c_str in col_map:
            new_cols[col] = col_map[c_str]
        else:
            for k, v in col_map.items():
                if k.lower() == c_str.lower():
                    new_cols[col] = v; break
    
    if new_cols:
        df_dams.rename(columns=new_cols, inplace=True)

    # Validation
    if 'codigoVerificacao' not in df_dams.columns:
        if len(df_dams.columns) >= 2: 
            df_dams.rename(columns={
                df_dams.columns[0]: 'codigoVerificacao',
                df_dams.columns[1]: 'referenciaPagamento'
            }, inplace=True)
            if len(df_dams.columns) >= 5:
                 df_dams.rename(columns={df_dams.columns[4]: 'totalRecolher'}, inplace=True)
        else:
            raise ValueError(f"Coluna 'codigoVerificacao' não encontrada.\nColunas lidas: {list(df_dams.columns)}")

    # Process Data
    dam_payments_map = {}
    
    for _, row in df_dams.iterrows():
        try:
            # ✅ FILTER 1: Tributo must be 'ISS Normal' (if column exists)
            tributo_val = str(row.get('tributo', '')).strip()
            if 'tributo' in df_dams.columns and tributo_val and tributo_val.lower() != 'iss normal':
                continue

            # ✅ FILTER 2: Must be 'Avulso' (Empty Invoice Numbers)
            # If 'numerosDasNotas' has content (e.g., "193"), it is linked to a note and is NOT a free credit.
            notas_val = row.get('numerosDasNotas')
            if pd.notna(notas_val) and str(notas_val).strip() != '':
                continue

            # Parse Date (MM/YYYY)
            ref_str = str(row.get('referenciaPagamento', '')).strip()
            if not ref_str or len(ref_str) < 6: continue
            
            try:
                if '/' in ref_str:
                    parts = ref_str.split('/')
                    key = f"{int(parts[0])}/{parts[1]}"
                elif len(ref_str) == 6: # 012021
                    m = int(ref_str[:2])
                    y = ref_str[2:]
                    key = f"{m}/{y}"
                else:
                    key = ref_str
            except:
                key = ref_str

            # Parse Value
            val_raw = row.get('totalRecolher', 0)
            if isinstance(val_raw, str):
                val_raw = val_raw.replace('R$', '').replace('.', '').replace(',', '.')
            
            val_float = float(val_raw)
            
            code = str(row.get('codigoVerificacao', '')).strip()

            if key not in dam_payments_map:
                dam_payments_map[key] = []
            
            dam_payments_map[key].append({
                'val': val_float,
                'code': code
            })
            
        except Exception:
            continue

    return dam_payments_map

def _load_all_dams_formatted(dam_filepath):
    """
    Loads the DAM CSV and formats it specifically for the Word report table.
    Filters: Only 'ISS Normal' and empty 'numerosDasNotas' (Avulsas).
    Uses robust loading to handle Brazilian CSV formats (semicolons, encoding).
    """
    if not dam_filepath or not os.path.exists(dam_filepath):
        return []

    try:
        # --- 1. ROBUST LOADING (Same as _load_and_process_dams) ---
        df = None
        attempts = [
            ('utf-8-sig', ','), 
            ('utf-8-sig', ';'), 
            ('latin1', ';'), 
            ('latin1', ',')
        ]
        
        for encoding, sep in attempts:
            try:
                df = pd.read_csv(dam_filepath, sep=sep, decimal='.', thousands=',', encoding=encoding, on_bad_lines='skip')
                if len(df.columns) > 1: break
            except: continue
            
        if df is None or len(df.columns) < 2:
             try: df = pd.read_csv(dam_filepath, sep=None, engine='python', encoding='utf-8-sig', on_bad_lines='skip')
             except: return []

        # --- 2. NORMALIZE HEADERS (Fixes spaces, BOM, case) ---
        df.columns = df.columns.astype(str).str.strip().str.replace(r'^[^\w]+', '', regex=True)
        
        col_map = {
            'tributo': 'tributo', 'Tributo': 'tributo',
            'numerosDasNotas': 'numerosDasNotas', 'Notas': 'numerosDasNotas', 'Números das Notas': 'numerosDasNotas',
            'receita': 'receita', 'Receita': 'receita',
            'totalRecolher': 'totalRecolher', 'Valor': 'totalRecolher', 'Valor Pago': 'totalRecolher',
            'codigoVerificacao': 'codigoVerificacao', 'Código Verificação': 'codigoVerificacao',
            'referenciaPagamento': 'referenciaPagamento', 'Competência': 'referenciaPagamento'
        }
        
        new_cols = {}
        for col in df.columns:
            c_str = str(col).strip()
            # Try exact match first, then case-insensitive
            if c_str in col_map: 
                new_cols[col] = col_map[c_str]
            else:
                for k, v in col_map.items():
                    if k.lower() == c_str.lower(): 
                        new_cols[col] = v; break
        
        if new_cols: df.rename(columns=new_cols, inplace=True)

        # --- 3. FILTER & FORMAT ---
        formatted_rows = []
        for _, row in df.iterrows():
            
            # Filter A: Must be "ISS Normal" (Ignore others)
            # We use .get() safely in case the column is missing in a specific file
            tributo_raw = str(row.get('tributo', '')).strip()
            if 'tributo' in df.columns and tributo_raw.lower() != "iss normal":
                continue

            # Filter B: Must NOT have Invoice Numbers (Avulsas only)
            notas_raw = row.get('numerosDasNotas')
            if pd.notna(notas_raw) and str(notas_raw).strip() != "":
                # If field is not empty, it's linked to an invoice, so we skip it
                continue
            
            # Prepare Values
            val_pago = pd.to_numeric(row.get('totalRecolher'), errors='coerce')
            val_receita = pd.to_numeric(row.get('receita'), errors='coerce')
            
            row_dict = {
                'codigo': str(row.get('codigoVerificacao', '-')).strip(),
                'competencia': str(row.get('referenciaPagamento', '-')),
                'receita': _format_currency_brl(val_receita) if pd.notna(val_receita) else "R$ 0,00",
                'valor_pago': _format_currency_brl(val_pago) if pd.notna(val_pago) else "R$ 0,00",
                'tributo': tributo_raw if tributo_raw else '-',
                'notas_associadas': "-" # Forced to dash since we filtered for empty
            }
            formatted_rows.append(row_dict)
            
        return formatted_rows

    except Exception as e:
        logging.error(f"Error loading full DAMs table: {e}")
        return []
    

def create_context_for_generation(master_filepath, company_cnpj,
                                   final_data, preview_context, 
                                   numero_multa, company_invoices_df, 
                                   company_imu=None, idd_mode=False, 
                                   dam_filepath=None):
    """
    Builds the final context, using the pre-calculated auto data
    from the wizard's preview_context and enriching it with
    data from final_data (like invoice lists).
    Filters invoice lists based on monthly compensation.
    """
    try:
        try:
            locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        except locale.Error:
            logging.warning("Locale 'pt_BR.utf8' not found. Month names may be in English.")

        hoje = datetime.now()
        data_por_extenso = hoje.strftime('%d de %B de %Y')
        dia_data = hoje.strftime('%d')
        mes_data = hoje.strftime('%m')
        ano_data = hoje.strftime('%Y')

        logging.info("Inside create_context_for_generation. Loading company data...")

        logging.info(f"Data Loader: Verificando master_filepath. O valor é: '{master_filepath}'")
        
        if not master_filepath or not isinstance(master_filepath, str):
            # Log o erro fatal e levanta a exceção
            err_msg = f"CRASH: 'master_filepath' é inválido. Esperava-se um string de caminho, mas o valor é: {master_filepath} (Tipo: {type(master_filepath)})"
            logging.error(err_msg)
            raise ValueError(err_msg)
        
        logging.info("Data Loader: master_filepath parece ser um string válido. A tentar processar o caminho...")

        base, ext = os.path.splitext(master_filepath)
        cadastro_path = f"{base}_cadastro{ext}"

        logging.info(f"Data Loader: Caminho do cadastro determinado como: '{cadastro_path}'")
        if not os.path.exists(cadastro_path):
            raise FileNotFoundError(f"File not found: {cadastro_path}")

        logging.info(f"Data Loader: Verificando se o ficheiro existe em: {cadastro_path}")
        if not os.path.exists(cadastro_path):
            logging.warning(f"Data Loader: Ficheiro de cadastro NÃO encontrado. A tentar carregar do mestre.")
            # (a lógica 'raise FileNotFoundError' original estava aqui, mas é melhor
            #  deixar a lógica de fallback 'else:' mais abaixo tratar disso)
        else:
            logging.info(f"Data Loader: Ficheiro de cadastro ENCONTRADO.")
        
        logging.info(f"Data Loader: A tentar ler o Excel: {cadastro_path}")

        # ✅ --- INÍCIO DO BLOCO DE DEBUG BLINDADO ---
        try:
            df_empresas = pd.read_excel(cadastro_path, dtype={'cnpj': str, 'imu': str, 'cep': str, 'epaf_numero': str})
        
        except PermissionError as pe:
            # Erro específico de ficheiro bloqueado
            logging.error(f"Data Loader: CRASH (PermissionError)! O ficheiro está bloqueado (aberto no Excel?): {cadastro_path}")
            logging.exception(pe)
            raise pe # Lança o erro para o worker
        except FileNotFoundError as fe:
            # Fallback (apesar de termos verificado)
            logging.error(f"Data Loader: CRASH (FileNotFoundError)! O ficheiro não foi encontrado em: {cadastro_path}")
            logging.exception(fe)
            raise fe
        except Exception as e:
            # Outros erros (ex: ficheiro corrompido, xlrd/openpyxl em falta)
            logging.error(f"Data Loader: CRASH (Exception)! Falha ao ler o Excel. O ficheiro pode estar corrompido.")
            logging.exception(e)
            raise e
        
        logging.info(f"Data Loader: Leitura do Excel concluída com sucesso. {len(df_empresas)} linhas lidas.")
        company_data = df_empresas[df_empresas['cnpj'] == company_cnpj]
        if company_data.empty:
            raise ValueError(f"ERROR: CNPJ '{company_cnpj}' not found.")
        context = company_data.iloc[0].to_dict()
        if company_imu:
            context['imu'] = company_imu
            logging.info(f"Using IMU from override: {company_imu}")
        context['data_por_extenso'] = data_por_extenso
        context['dia_data'] = dia_data
        context['mes_data'] = mes_data
        context['ano_data'] = ano_data
        general_texts = get_custom_general_texts()
        context['AUDITOR_NOME'] = general_texts.get("AUDITOR_NOME", "Nome Padrão")
        context['AUDITOR_MATRICULA'] = general_texts.get("AUDITOR_MATRICULA", "0000")
        logging.info(f"Company '{context.get('razao_social')}' loaded.")

        epaf_from_wizard = preview_context.get('epaf_numero')
        if epaf_from_wizard:
            context['epaf_numero'] = epaf_from_wizard
            logging.info(f"Using ePAF number from wizard override: {epaf_from_wizard}")

        context['idd_mode'] = idd_mode

        period_start_date = None
        period_end_date = None
        df_dates = pd.Series(dtype='datetime64[ns]')

        if not company_invoices_df.empty and 'DATA EMISSÃO' in company_invoices_df.columns:
            # Ensure DATA EMISSÃO is datetime before using .dt accessor
            if not pd.api.types.is_datetime64_any_dtype(company_invoices_df['DATA EMISSÃO']):
                company_invoices_df['DATA EMISSÃO'] = pd.to_datetime(company_invoices_df['DATA EMISSÃO'], errors='coerce')

            df_dates = company_invoices_df['DATA EMISSÃO'].dropna()
            if not df_dates.empty:
                min_date = df_dates.min()
                max_date = df_dates.max()
                period_start_date = datetime(min_date.year, 1, 1)
                period_end_date = max_date
                
                # ✅ --- START: FIX 1 (Periodo Fiscalizado) ---
                # Use "janeiro a dezembro de ANO" based on the max_date
                ano_fiscalizado = max_date.year
                context['periodo_fiscalizado'] = f"janeiro a dezembro de {ano_fiscalizado}"
                # ✅ --- END: FIX 1 ---
                
                context['ano'] = max_date.strftime('%Y')
        
        ano_str = context.get('ano', '____')
        if idd_mode:
            # TÍTULO IDD
            context['titulo_documento'] = "Informação Fiscal - Operação IDD"
            
            # INTRO IDD (Com substituição do ano)
            intro_idd_template = (
                "O presente procedimento teve como objeto a análise das informações constantes nas Notas "
                "Fiscais de Serviço Eletrônicas (NFS-e) emitidas pelo contribuinte no "
                f"decorrer do ano-calendário de {ano_str}, com vistas à formalização dos "
                "débitos de Imposto sobre Serviços de Qualquer Natureza (ISS) confessados e "
                "constituídos pelas NFS no sistema ISS-Curitiba, sem pagamento encontrado aos "
                "Cofres Municipais."
            )
            context['I_INTRO'] = intro_idd_template
        else:
            # TÍTULO PADRÃO (RECEITAS)
            context['titulo_documento'] = "Relatório do Procedimento Fiscal da Operação Receitas (monitoramento)"
            
            # INTRO PADRÃO (Mantém o que vier do config/default, não sobrescreve se não for IDD)
            # Se 'I_INTRO' não existir no config, o report_generator usará o default.
            pass
        
        # --- Get all infraction indices ---
        all_infraction_indices = []
        for auto_info in final_data.values():
            all_infraction_indices.extend(auto_info.get('invoices', []))
        unique_indices = list(set(all_infraction_indices))
        # 💡 FIX: Check if indices exist before filtering
        valid_indices = company_invoices_df.index.intersection(unique_indices)
        df_all_infractions = company_invoices_df.loc[valid_indices].copy()


        # --- Multa Logic ---
        instrumental_causes = [
            'Dedução indevida', 'Regime incorreto', 'Isenção/Imunidade Indevida',
            'Natureza da Operação Incompatível', 'Local da incidência incorreto',
            'Retenção na Fonte (Verificar)', 'Alíquota Incorreta'
        ]
        def check_for_instrumental(details_list):
            if not isinstance(details_list, list): return False
            for detail in details_list:
                for cause in instrumental_causes:
                    if str(detail).startswith(cause):
                        return True
            return False

        has_instrumental_infractions = False
        # 💡 FIX: Check if 'broken_rule_details' column exists
        if 'broken_rule_details' in df_all_infractions.columns:
            df_instrumental_infractions = df_all_infractions[
                df_all_infractions['broken_rule_details'].apply(check_for_instrumental)
            ]
            has_instrumental_infractions = not df_instrumental_infractions.empty
        else:
            logging.warning("Column 'broken_rule_details' not found. Cannot determine instrumental infractions.")
        
        texto_multa_final = preview_context.get('texto_multa', '')
        valor_multa_final = preview_context.get('valor_multa', '')
        if numero_multa and has_instrumental_infractions and texto_multa_final:
            context['multa_aplicada'] = {
                'aplicada': True, 'texto_multa': texto_multa_final,
                'numero': numero_multa, 'valor': valor_multa_final
            }
            context['multa_sem_infracao'] = False
        else:
            context['multa_aplicada'] = {'aplicada': False, 'texto_multa': ''}
            context['multa_sem_infracao'] = not has_instrumental_infractions


        # --- Other findings logic ---
        # 💡 FIX: Check required columns for 'achados'.
        if 'status_legal' in company_invoices_df.columns:
            df_decadente_nao_autuado = company_invoices_df[(company_invoices_df['status_legal'] == 'Decadente') & (~company_invoices_df.index.isin(all_infraction_indices))]
            # 💡 FIX: Check columns *inside* the block
            if not df_decadente_nao_autuado.empty and 'DATA EMISSÃO' in df_decadente_nao_autuado.columns and 'NÚMERO' in df_decadente_nao_autuado.columns:
                valid_dates = df_decadente_nao_autuado['DATA EMISSÃO'].dropna()
                if not valid_dates.empty:
                    min_date = valid_dates.min().strftime('%m/%Y')
                    max_date = valid_dates.max().strftime('%m/%Y')
                    periodo = min_date if min_date == max_date else f"{min_date} a {max_date}"
                    nfs_numeros = format_invoice_numbers(df_decadente_nao_autuado['NÚMERO'].astype(str).unique())
                    context['achado_decadencia_nao_autuado'] = {'periodo': periodo, 'nfs_numeros': nfs_numeros}
            
            df_prescrito_nao_autuado = company_invoices_df[(company_invoices_df['status_legal'] == 'Prescrito') & (~company_invoices_df.index.isin(all_infraction_indices))]
            if not df_prescrito_nao_autuado.empty and 'DATA EMISSÃO' in df_prescrito_nao_autuado.columns and 'NÚMERO' in df_prescrito_nao_autuado.columns:
                valid_dates = df_prescrito_nao_autuado['DATA EMISSÃO'].dropna()
                if not valid_dates.empty:
                    min_date = valid_dates.min().strftime('%m/%Y')
                    max_date = valid_dates.max().strftime('%m/%Y')
                    periodo = min_date if min_date == max_date else f"{min_date} a {max_date}"
                    nfs_numeros = format_invoice_numbers(df_prescrito_nao_autuado['NÚMERO'].astype(str).unique())
                    context['achado_prescrito_nao_autuado'] = {'periodo': periodo, 'nfs_numeros': nfs_numeros}
        
        # 💡 FIX: Check all required columns before filtering
        required_cols_fora = ['NATUREZA DA OPERAÇÃO', 'CÓDIGO DA ATIVIDADE', 'DATA EMISSÃO', 'NÚMERO']
        if all(col in company_invoices_df.columns for col in required_cols_fora):
            df_fora = company_invoices_df[(company_invoices_df['NATUREZA DA OPERAÇÃO'].str.contains("Fora do Município", case=False, na=False)) & (company_invoices_df['CÓDIGO DA ATIVIDADE'] == '0702')]
            if not df_fora.empty:
                valid_dates = df_fora['DATA EMISSÃO'].dropna()
                if not valid_dates.empty:
                    min_date = valid_dates.min().strftime('%m/%Y')
                    max_date = valid_dates.max().strftime('%m/%Y')
                    periodo = min_date if min_date == max_date else f"{min_date} a {max_date}"
                    nfs_numeros = format_invoice_numbers(df_fora['NÚMERO'].astype(str).unique())
                    context['achado_fora_municipio'] = {'periodo': periodo, 'nfs_numeros': nfs_numeros}
        
        # ✅ --- START: Report Text for 'Local Tomador' ---
        # 💡 FIX: Check required columns
        if 'status_manual' in company_invoices_df.columns and 'activity_desc' in company_invoices_df.columns and 'NÚMERO' in company_invoices_df.columns:
            df_tomador = company_invoices_df[company_invoices_df['status_manual'] == 'Local_Tomador']
            if not df_tomador.empty:
                achados_tomador_list = []
                # Agrupa por descrição da atividade para criar o texto
                for desc, group_df in df_tomador.groupby('activity_desc'):
                    nfs_numeros = format_invoice_numbers(group_df['NÚMERO'].astype(str).unique())
                    if not desc or pd.isna(desc):
                        desc = "Atividade Não Especificada"
                    achados_tomador_list.append(
                        f"NFes de nº(s) {nfs_numeros} foram desconsideradas pois o serviço '{desc}' tem tributação no local do tomador."
                    )
                context['achado_local_tomador'] = "\n".join(achados_tomador_list)
        # ✅ --- END: Report Text for 'Local Tomador' ---
        
        if period_start_date and period_end_date and not df_dates.empty:
            all_months = pd.date_range(period_start_date, period_end_date, freq='MS').strftime('%Y-%m').tolist()
            months_with_invoices = df_dates.dt.strftime('%Y-%m').unique().tolist()
            months_without = sorted(list(set(all_months) - set(months_with_invoices)))
            if months_without:
                formatted = [datetime.strptime(m, '%Y-%m').strftime('%m/%Y') for m in months_without]
                context['achado_sem_notas'] = {'periodo': ', '.join(formatted)}

        # --- Process Autos ---
        all_selected_autos = preview_context.get('autos', [])

        autos_para_relatorio = []
        autos_compensados = []
        for auto_data in all_selected_autos:
            total_credito_final = auto_data.get('totais', {}).get('iss_apurado_op', 0.0)
            
            # ✅ Check for manual override
            manual_credit = auto_data.get('user_defined_credito')
            if manual_credit is not None:
                total_credito_final = manual_credit
                
            if total_credito_final > 0.01:
                autos_para_relatorio.append(auto_data)
            else:
                autos_compensados.append(auto_data)
        
        context['autos'] = autos_para_relatorio
        # 💡 FIX: Safe list comprehension
        logging.info(f"Autos para relatório (valor > 0): {[a.get('numero', 'N/A') for a in autos_para_relatorio]}")
        
        if autos_compensados:
            # 💡 FIX: Safe list comprehension
            numeros_autos_compensados = [auto.get('numero', 'N/A').replace('AUTO-', '') for auto in autos_compensados]
            lista_autos_compensados_str = format_invoice_numbers(numeros_autos_compensados)
            context['achado_autos_compensados'] = {'lista_numeros': lista_autos_compensados_str}

        lista_strings_compensadas = []
        
        for auto_data in context['autos']:
            auto_key = auto_data.get('numero')
            if not auto_key: continue
            
            original_auto_info = final_data.get(auto_key)
            if not original_auto_info: continue

            invoice_indices = original_auto_info.get('invoices', [])
            valid_indices = company_invoices_df.index.intersection(invoice_indices)
            if len(valid_indices) == 0: continue
            
            df_invoices = company_invoices_df.loc[valid_indices].copy()
            
            # Prepare df_invoices for date filtering
            if 'DATA EMISSÃO' not in df_invoices.columns:
                continue # Cannot filter without dates
            
            if not pd.api.types.is_datetime64_any_dtype(df_invoices['DATA EMISSÃO']):
                df_invoices['DATA EMISSÃO'] = pd.to_datetime(df_invoices['DATA EMISSÃO'], errors='coerce')
            
            df_invoices['mes_ano_str'] = df_invoices['DATA EMISSÃO'].dt.strftime('%m/%Y')
            
            dados_anuais = auto_data.get('dados_anuais', [])
            
            # Identify compensated months AND their source
            for mes_data in dados_anuais:
                # Check if this specific month resulted in 0.00 debt
                if mes_data.get('iss_apurado_op', 0.0) <= 0.01:
                    mes_ano_val = mes_data.get('mes_ano')
                    if not mes_ano_val: continue
                    
                    # Determine source of payment for this month
                    sources = []
                    if mes_data.get('dam_iss_pago', 0.0) > 0.001:
                        sources.append("DAM")
                    if mes_data.get('das_iss_pago', 0.0) > 0.001:
                        sources.append("PGDAS")
                    
                    if not sources: continue # Should not happen if compensated, but safety check
                    
                    source_label = " e ".join(sources)
                    
                    # Filter invoices specifically for THIS month
                    invoices_mes_mask = df_invoices['mes_ano_str'] == mes_ano_val
                    invoices_mes_df = df_invoices[invoices_mes_mask]
                    
                    # Remove these invoices from the main listing dataframe (for the table header)
                    # We modify the auto_data['nfs_e_numeros'] later based on remaining
                    df_invoices = df_invoices[~invoices_mes_mask] 

                    if not invoices_mes_df.empty and 'NÚMERO' in invoices_mes_df.columns:
                        nums = format_invoice_numbers(invoices_mes_df['NÚMERO'].astype(str).tolist())
                        if nums:
                            # ✅ Format: "NFS-e 9 a 10 (Comp: 08/2021 via DAM) referente ao AUTO-001"
                            texto = f"NFS-e {nums} (Comp: {mes_ano_val} via {source_label}) referente ao {auto_key}"
                            lista_strings_compensadas.append(texto)

            # Recalculate the main list of invoices for the auto (excluding the compensated ones)
            if not df_invoices.empty and 'NÚMERO' in df_invoices.columns:
                auto_data['nfs_e_numeros'] = format_invoice_numbers(df_invoices['NÚMERO'].astype(str).tolist())
                
                # Recalculate Period for the remaining invoices
                valid_dates = df_invoices['DATA EMISSÃO'].dropna()
                if not valid_dates.empty:
                    min_d, max_d = valid_dates.min(), valid_dates.max()
                    if min_d.month == max_d.month and min_d.year == max_d.year:
                        auto_data['periodo'] = min_d.strftime('%m/%Y')
                    else:
                        auto_data['periodo'] = f"{min_d.strftime('%m/%Y')} a {max_d.strftime('%m/%Y')}"
            else:
                auto_data['nfs_e_numeros'] = "N/A" # Or handle as fully compensated auto logic (though logic above separates fully compensated autos)

            # Motive Logic (Preserved)
            default_aliquota_str = original_auto_info.get('correct_aliquota', '[N/A]')
            monthly_overrides = original_auto_info.get('monthly_overrides', {})
            final_aliquota_str_to_use = default_aliquota_str
            if monthly_overrides:
                vals = [v for v in monthly_overrides.values() if v > 0]
                if vals: 
                    try: final_aliquota_str_to_use = f"{mode(vals):.2f}"
                    except: final_aliquota_str_to_use = f"{vals[0]:.2f}"
            
            auto_data['motivo'] = {
                'tipo': original_auto_info.get('rule_name', 'DEFAULT_AUTO_FALLBACK'),
                'texto_simples': original_auto_info.get('motive_text', 'Motivo não especificado.'),
                'aliquota_correta': final_aliquota_str_to_use 
            }
            
            # Ensure defaults for table
            for mes in auto_data.get('dados_anuais', []):
                mes.setdefault('das_aliquota', '-'); mes.setdefault('das_identificacao', '-')
                mes.setdefault('dam_aliquota', '-'); mes.setdefault('dam_identificacao', '-')
                mes.setdefault('aliquota_declarada', '-'); mes.setdefault('iss_declarado_pago', 0.0)

        # ✅ Assign formatted list to context
        if lista_strings_compensadas:
            context['achado_invoices_compensadas'] = {
                'lista_formatada': lista_strings_compensadas
            }

        # Build Final Summary
        if preview_context and 'summary' in preview_context:
            context['summary'] = preview_context['summary']
            
            # ✅ --- START: FIX 3 (Summary Table Totals - COMPLETE REBUILD) ---
            # Instead of patching the potentially stale 'context['summary']['autos']',
            # we rebuild it completely from 'autos_para_relatorio'.
            # This ensures every auto that appears in the detail section appears here with the correct value.

            new_summary_autos = []
            summary_total_autos = 0.0
            
            for auto in autos_para_relatorio:
                auto_num = auto.get('numero')
                
                # Determine Final Value
                if auto.get('user_defined_credito') is not None:
                    val = float(auto.get('user_defined_credito'))
                else:
                    val = float(auto.get('totais', {}).get('iss_apurado_op', 0.0))
                
                # Get formatted NFS list
                nfs_str = auto.get('nfs_e_numeros', 'N/A')
                # Get Original ISS (Gross)
                iss_original = float(auto.get('totais', {}).get('iss_apurado_bruto', 0.0))
                # Get Motive
                motivo_str = auto.get('motive_text', '') or auto.get('motivo', {}).get('texto_simples', '')

                new_summary_autos.append({
                    'numero': auto_num,
                    'nfs_tributadas': nfs_str,
                    'iss_valor_original': iss_original,
                    'total_credito_tributario': val,
                    'motivo': motivo_str
                })
                
                summary_total_autos += val

            # Overwrite the old summary list
            context['summary']['autos'] = new_summary_autos
            
            # ✅ CHANGED: Handle Multiple Multas List
            summary_multas_list = context['summary'].get('multas', [])
            summary_total_multas = 0.0
            if summary_multas_list:
                for m in summary_multas_list:
                    summary_total_multas += float(m.get('valor_credito', 0.0))
            
            # Update Grand Total
            context['summary']['total_geral_credito'] = summary_total_autos + summary_total_multas
            
            # Keep backwards compatibility for templates expecting single 'multa' object
            context['multa_aplicada'] = {
                'aplicada': len(summary_multas_list) > 0,
                'texto_multa': preview_context.get('texto_multa', ''),
                'numero': summary_multas_list[0].get('numero') if summary_multas_list else '',
                'valor': summary_multas_list[0].get('valor_credito') if summary_multas_list else ''
            }
            
            logging.info("Successfully REBUILT 'summary' data in final context.")
            
        else:
            context['summary'] = {} 
            logging.warning("No 'summary' data found in preview_context. Conclusion table will be empty.")
            
        # --- Default fallbacks ---
        context.setdefault('idds', None); context.setdefault('pagamentos_avulsos', None)
        context.setdefault('infracao_existente', None); context.setdefault('das_existente', None)
        context.setdefault('achado_sem_pagamento', None); context.setdefault('multa_dispensada_is', None)
        context.setdefault('multa_dispensada_simples', None)
        context.setdefault('achado_autos_compensados', None)
        context.setdefault('achado_invoices_compensadas', None)
        context.setdefault('achado_sem_notas', None)
        context.setdefault('achado_fora_municipio', None)
        context.setdefault('achado_decadencia_nao_autuado', None)
        context.setdefault('achado_prescrito_nao_autuado', None)
        context.setdefault('achado_local_tomador', None) # ✅ Add Fallback

        # List auto numbers
        lista_autos_numeros = [auto.get('numero', 'N/A') for auto in context.get('autos', [])]
        context['lista_autos_numeros'] = lista_autos_numeros
        logging.info(f"Generated list of auto numbers for templates: {lista_autos_numeros}")
        context['pagamentos_avulsos'] = None
        if dam_filepath:
            dams_data = _load_all_dams_formatted(dam_filepath)
            if dams_data:
                context['pagamentos_avulsos'] = dams_data # ✅ Set correct key
                logging.info(f"Loaded {len(dams_data)} rows for DAMs table.")

        return context

    except Exception as e:
        logging.exception("CRITICAL ERROR in create_context_for_generation")
        raise