# --- FILE: app/workers.py ---
# (Updated MultiYearPrepWorker to fix "Missing Invoices" bug via safer deduplication)

import pandas as pd
import traceback
import logging
from PySide6.QtCore import QObject, Signal, QCoreApplication
import multiprocessing
from app.generation_task import run_generation_task
from app.shared_memory import share_dataframe
from app.updater import Updater
import os
import copy
from datetime import datetime
from app.constants import Columns
from data_loader import _load_and_process_dams
from app.pgdas_loader import _load_and_process_pgdas
from document_parts import format_invoice_numbers
import tempfile
import glob
import urllib.request

multiprocessing.freeze_support() 

class NewsFetcherWorker(QObject):
    """
    Fetches simple text content from a remote URL for the 'Parochial News' feature.
    """
    finished = Signal(str)
    
    def __init__(self, url):
        super().__init__()
        self.url = url

    def run(self):
        try:
            # Set a timeout so it doesn't hang if internet is slow
            with urllib.request.urlopen(self.url, timeout=5) as response:
                data = response.read().decode('utf-8')
                self.finished.emit(data)
        except Exception:
            # If offline or 404, emit empty string (UI will handle it)
            self.finished.emit("")

class UpdateCheckerWorker(QObject):
    finished = Signal(bool, str, str)
    error = Signal(str)
    def run(self):
        logger = logging.getLogger(__name__)
        try:
            updater = Updater()
            available, url, ver = updater.check_for_updates()
            self.finished.emit(available, url, ver)
        except Exception as e:
            logger.exception("UpdateCheckerWorker crashed")
            self.error.emit(str(e))

class UpdateDownloaderWorker(QObject):
    progress = Signal(str)
    error = Signal(str)
    reboot_requested = Signal()
    def __init__(self, download_url):
        super().__init__()
        self.download_url = download_url
    def run(self):
        logger = logging.getLogger(__name__)
        try:
            updater = Updater()
            success = updater.download_and_install(self.download_url, status_callback=self.progress.emit)
            if success: self.reboot_requested.emit()
        except Exception as e:
            logger.exception("UpdateDownloaderWorker crashed")
            self.error.emit(str(e))

class BaseWorker(QObject):
    finished = Signal() 
    error = Signal(str)
    progress = Signal(str)
    stop_signal = Signal()
    def __init__(self, *args, **kwargs):
        super().__init__()
        self._stop_requested = False
        self.stop_signal.connect(self.stop) 
    def stop(self):
        self._stop_requested = True
        self.progress.emit("⚠️ Tentativa de Parada... Aguardando ponto de interrupção seguro.")
    def check_stop(self):
        return self._stop_requested

class AIPrepWorker(BaseWorker):
    finished = Signal(pd.DataFrame) 
    def __init__(self, master_path, invoices_path, cnpj):
        super().__init__()
        self.master_path, self.invoices_path, self.cnpj = master_path, invoices_path, cnpj
        
    def run(self):
        try:
            # We import locally to avoid circular deps if any
            from main import load_and_prepare_invoices
            
            self.progress.emit("🔄 Iniciando carregamento (Modo Interativo)...")
            
            # ✅ DISABLE AUTO RESOLVE -> UI will handle it
            final_df = load_and_prepare_invoices(
                self.master_path, 
                self.invoices_path, 
                self.cnpj, 
                status_callback=self.progress,
                auto_resolve_conflicts=False 
            )
            self.finished.emit(final_df)
        except Exception:
            self.error.emit(f"❌ Erro durante a preparação e análise (IA):\n{traceback.format_exc()}")

class AIAnalysisWorker(BaseWorker):
    finished = Signal(pd.DataFrame)
    def __init__(self, invoices_df):
        super().__init__()
        self.invoices_df = invoices_df
    def run(self):
        try:
            from main import perform_description_analysis
            analyzed_df = perform_description_analysis(self.invoices_df, status_callback=self.progress)
            self.finished.emit(analyzed_df)
        except Exception:
            self.error.emit(f"❌ Erro durante a análise de descrição (IA):\n{traceback.format_exc()}")

class MultiYearPrepWorker(BaseWorker):
    """
    Scans a folder for Invoice Excels and DAM CSVs/Excels.
    Merges them into single datasets and PREPARES CONFLICTS for UI review.
    ✅ FIXED: Disable auto-resolve in inner loop to catch intra-file duplicates.
    """
    finished = Signal(pd.DataFrame, str) 

    def __init__(self, folder_path, master_path, cnpj):
        super().__init__()
        self.folder_path = folder_path
        self.master_path = master_path
        self.cnpj = cnpj

    def _normalize_dam_df(self, df):
        if df is None or df.empty: return df
        df.columns = df.columns.astype(str).str.strip().str.replace(r'^[^\w]+', '', regex=True)
        col_map = {
            'codigoVerificacao': 'codigoVerificacao', 'Código Verificação': 'codigoVerificacao',
            'Codigo Verificacao': 'codigoVerificacao', 'Código de Verificação': 'codigoVerificacao', 
            'Codigo': 'codigoVerificacao', 'Verificação': 'codigoVerificacao',
            'Competência': 'referenciaPagamento', 'referenciaPagamento': 'referenciaPagamento',
            'Referência': 'referenciaPagamento',
            'receita': 'receita', 'Receita': 'receita',
            'Valor': 'totalRecolher', 'Valor Pago': 'totalRecolher', 'totalRecolher': 'totalRecolher'
        }
        new_cols = {}
        for col in df.columns:
            c = str(col).strip()
            if c in col_map: 
                new_cols[col] = col_map[c]
            else:
                for k,v in col_map.items():
                    if k.lower() == c.lower(): 
                        new_cols[col] = v
                        break
        if new_cols: df.rename(columns=new_cols, inplace=True)
        return df

    def run(self):
        try:
            from main import load_and_prepare_invoices
            
            self.progress.emit(f"📂 Escaneando pasta: {self.folder_path}")

            # --- 1. INVOICE LOADING ---
            all_files = glob.glob(os.path.join(self.folder_path, "*.xls*"))
            invoice_files = []
            
            for f in all_files:
                fname = os.path.basename(f).lower()
                if "controle_de_atividades" in fname: continue
                if "~$" in fname: continue
                # Logic: Exclude only DAM reports
                if "dam" in fname and "relatorio" in fname: continue 
                
                invoice_files.append(f)

            if not invoice_files:
                self.error.emit("❌ Nenhum ficheiro de notas (Excel) encontrado na pasta.")
                return

            merged_df = pd.DataFrame()
            self.progress.emit(f"🔄 Carregando {len(invoice_files)} ficheiros de notas...")
            
            for f in invoice_files:
                self.progress.emit(f"   - Lendo: {os.path.basename(f)}")
                try:
                    # ✅ FIXED: auto_resolve_conflicts=False to PRESERVE duplicates for review
                    df = load_and_prepare_invoices(self.master_path, f, self.cnpj, auto_resolve_conflicts=False)
                    if not df.empty:
                        merged_df = pd.concat([merged_df, df], ignore_index=True)
                except Exception as e:
                    self.progress.emit(f"   ⚠️ Erro ao ler nota {os.path.basename(f)}: {e}")

            if merged_df.empty:
                self.error.emit("❌ Dados vazios após leitura dos arquivos de notas.")
                return

            # ✅ FIX: Ensure 'status_manual' exists
            if 'status_manual' not in merged_df.columns:
                merged_df['status_manual'] = None
            if '_is_conflict' not in merged_df.columns:
                merged_df['_is_conflict'] = False

            # --- 2. MULTI-YEAR CONFLICT DETECTION ---
            
            # ✅ FIX: Robust ID Column Finder
            id_col = None
            possible_ids = ['NÚMERO', 'NUMERO', 'NUMBER', 'NO', 'Nº']
            upper_cols = {c.upper(): c for c in merged_df.columns}
            
            for p in possible_ids:
                if p in upper_cols:
                    id_col = upper_cols[p]
                    break
            
            if id_col:
                self.progress.emit(f"   🔍 Analisando conflitos (Coluna ID: {id_col})...")
                
                # 1. Normalize Number
                merged_df[id_col] = (
                    merged_df[id_col]
                    .astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                )
                
                # 2. Smart Sort
                desc_synonyms = ['DISCRIMINAÇÃO DOS SERVIÇOS', 'DISCRIMINACAO DOS SERVICOS', 'DESCRIÇÃO', 'DESC']
                desc_col = None
                
                # Re-map upper cols in case they changed (unlikely but safe)
                upper_cols = {c.upper(): c for c in merged_df.columns}
                for t in desc_synonyms:
                    if t in upper_cols:
                        desc_col = upper_cols[t]
                        break

                if desc_col:
                    merged_df['_desc_len'] = merged_df[desc_col].astype(str).str.len()
                else:
                    merged_df['_desc_len'] = 0
                
                if 'ALÍQUOTA' not in merged_df.columns: merged_df['ALÍQUOTA'] = 0.0
                if 'VALOR' not in merged_df.columns: merged_df['VALOR'] = 0.0
                
                # Sort: ID (Asc), Aliquot (Desc), Value (Desc), DescLen (Desc)
                merged_df.sort_values(
                    by=[id_col, 'ALÍQUOTA', 'VALOR', '_desc_len'], 
                    ascending=[True, False, False, False], 
                    inplace=True
                )
                merged_df.drop(columns=['_desc_len'], inplace=True)

                # 3. Detect Conflicts (COMBINED)
                # Combine intra-file conflicts (already in _is_conflict) with inter-file conflicts (duplicated now)
                
                # Check for duplicates across the whole merged set
                inter_file_duplicates = merged_df.duplicated(subset=[id_col], keep=False)
                
                # OR logic: True if it was already a conflict OR if it is a conflict now in the merge
                # FillNA(False) ensures boolean operation works
                merged_df['_is_conflict'] = merged_df['_is_conflict'].fillna(False) | inter_file_duplicates
                
                conflict_count = merged_df[merged_df['_is_conflict']][id_col].nunique()
                
                if conflict_count > 0:
                    self.progress.emit(f"   ⚠️ Encontrados {conflict_count} conflitos entre arquivos/anos.")
                else:
                    self.progress.emit("   ✅ Nenhum conflito encontrado.")
            else:
                 self.progress.emit("   ⚠️ Aviso: Coluna NÚMERO não encontrada. Pulo verificação de conflitos.")
                 merged_df['_is_conflict'] = False

            # --- 3. DAM LOADING ---
            dam_files = glob.glob(os.path.join(self.folder_path, "*Relatorio_DAMS*.csv"))
            dam_files += glob.glob(os.path.join(self.folder_path, "*Relatorio_DAMS*.xls*"))
            merged_dams_path = ""
            
            if dam_files:
                self.progress.emit(f"🔄 Consolidando {len(dam_files)} ficheiros de DAMs...")
                dam_dfs = []
                for df_path in dam_files:
                    try:
                        df = None
                        if df_path.lower().endswith('.csv'):
                            try: df = pd.read_csv(df_path, on_bad_lines='skip', sep=None, engine='python')
                            except: pass
                        else:
                            df = pd.read_excel(df_path)
                        
                        if df is not None:
                            df = self._normalize_dam_df(df)
                            dam_dfs.append(df)
                    except Exception: pass

                if dam_dfs:
                    full_dam_df = pd.concat(dam_dfs, ignore_index=True)
                    # Deduplicate DAMs
                    subset_cols = ['codigoVerificacao']
                    if 'referenciaPagamento' in full_dam_df.columns: subset_cols.append('referenciaPagamento')
                    if 'codigoVerificacao' in full_dam_df.columns:
                         full_dam_df['codigoVerificacao'] = full_dam_df['codigoVerificacao'].astype(str).str.strip()
                    full_dam_df.drop_duplicates(subset=subset_cols, inplace=True)
                    
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix="_consolidated_auto_dams.csv", mode='w', encoding='utf-8')
                    full_dam_df.to_csv(temp.name, index=False, sep=',') 
                    temp.close()
                    merged_dams_path = temp.name
                    self.progress.emit(f"✅ {len(dam_dfs)} arquivos de DAMs unidos.")

            self.finished.emit(merged_df, merged_dams_path)

        except Exception as e:
            self.error.emit(f"❌ Erro na consolidação:\n{traceback.format_exc()}")

class RulesPrepWorker(BaseWorker):
    finished = Signal(dict, pd.DataFrame) 
    def __init__(self, all_invoices_df, idd_mode=False):
        super().__init__()
        self.all_invoices_df = all_invoices_df
        self.idd_mode = idd_mode

    def run(self):
        try:
            from main import perform_rules_analysis
            # 1. Run Standard Analysis
            infraction_groups, df_with_analysis = perform_rules_analysis(self.all_invoices_df, idd_mode=self.idd_mode)
            
            # 2. SPLIT GROUPS BY YEAR
            split_groups = {}
            for group_name, df_group in infraction_groups.items():
                if df_group.empty: continue
                
                if 'DATA EMISSÃO' not in df_group.columns:
                    split_groups[group_name] = df_group
                    continue
                
                # Temp year column
                df_group['__temp_year'] = pd.to_datetime(df_group['DATA EMISSÃO'], errors='coerce').dt.year
                unique_years = sorted(df_group['__temp_year'].dropna().unique())

                if len(unique_years) <= 1:
                    # Single year or no year -> Keep original name
                    split_groups[group_name] = df_group
                else:
                    # Multiple years -> Split
                    for year in unique_years:
                        new_key = f"{group_name} ({int(year)})"
                        year_df = df_group[df_group['__temp_year'] == year].copy()
                        del year_df['__temp_year']
                        split_groups[new_key] = year_df

            if '__temp_year' in df_with_analysis.columns:
                del df_with_analysis['__temp_year']

            self.finished.emit(split_groups, df_with_analysis)
        except Exception:
            self.error.emit(f"❌ Erro na análise de regras:\n{traceback.format_exc()}")

class GenerationWorker(QObject):
    progress = Signal(str)
    finished = Signal(bool, str) 
    error = Signal(str)
    stop_signal = Signal()

    def __init__(self, cnpj, master_path, final_data, preview_context, invoices_df, 
                 numero_multa, dam_filepath, pgdas_folder_path, output_dir, 
                 encerramento_version, company_imu, idd_mode=False):
        super().__init__()
        shared_df_info = share_dataframe(invoices_df)
        self.generation_args = {
            'cnpj': cnpj,
            'master_path': master_path,
            'final_data': final_data,
            'preview_context': preview_context,
            'shared_df_info': shared_df_info, 
            'numero_multa': numero_multa,
            'dam_filepath': dam_filepath,
            'pgdas_folder_path': pgdas_folder_path,
            'output_dir': output_dir,
            'encerramento_version': encerramento_version,
            'company_imu': company_imu,
            'idd_mode': idd_mode
        }
        self.output_dir = output_dir 
        try: multiprocessing.set_start_method('spawn', force=True)
        except RuntimeError: pass
        except Exception as e: logging.warning(f"Não foi possível forçar 'spawn': {e}")
        self.queue = multiprocessing.Queue()
        self.process = multiprocessing.Process(target=run_generation_task, args=(self.queue, self.generation_args))
        self.stop_signal.connect(self.stop) 

    def stop(self):
        self.progress.emit("⚠️ Tentando TERMINAR o subprocesso de Geração (Forçada)...")
        if self.process.is_alive():
            self.process.terminate()
            self.progress.emit("🛑 Subprocesso de Geração encerrado.")
        
    def run(self):
        try:
            self.process.start()
            success = False
            final_output_dir = self.output_dir
            from queue import Empty 
            while True:
                if not self.process.is_alive():
                    if self.process.exitcode == 0: break
                    elif self.process.exitcode != 0: 
                        self.error.emit(f"❌ CRASH: O subprocesso falhou (Exit Code: {self.process.exitcode}).")
                        success = False; break
                try:
                    message = self.queue.get(timeout=0.1) 
                    if isinstance(message, tuple):
                        if message[0] == "SUCCESS":
                            success = True; final_output_dir = message[1]; break 
                        elif message[0] == "ERROR":
                            self.error.emit(f"❌ Erro Crítico do Subprocesso:\n{message[1]}")
                            success = False; break 
                    else: self.progress.emit(str(message))
                except Empty: pass
                except Exception as e:
                    self.error.emit(f"Erro interno do worker: {e}")
                    success = False; 
                    if self.process.is_alive(): self.process.terminate()
                    break
            self.process.join(timeout=2)
            if self.process.is_alive(): self.process.terminate()
            self.queue.close()
            self.finished.emit(success, final_output_dir)
        except Exception as e:
            self.error.emit(f"❌ Erro Crítico no Worker (Thread):\n{e}")
            self.finished.emit(False, self.output_dir)

class SimplesReaderWorker(BaseWorker):
    finished = Signal(str)
    def __init__(self, root_folder, target_years):
        super().__init__()
        self.root_folder = root_folder
        self.target_years = target_years # Expect list
    def run(self):
        try:
            from app.ferramentas.simples_reader import run_simples_reader
            output_path = run_simples_reader(self.root_folder, self.target_years, self.progress)
            self.finished.emit(output_path)
        except Exception:
            self.error.emit(f"❌ Erro durante a Leitura (OCR) do Simples:\n{traceback.format_exc()}")

class SimplesDownloaderWorker(BaseWorker):
    def __init__(self, tasks):
        super().__init__()
        self.tasks = tasks
    def run(self):
        try:
            from app.ferramentas.simples_downloader import run_simples_downloader
            run_simples_downloader(self.tasks, self.progress)
            self.finished.emit()
        except Exception:
            self.error.emit(f"❌ Erro durante o Download (RPA) do Simples:\n{traceback.format_exc()}")

class DatabaseExtractorWorker(BaseWorker):
    def __init__(self, target_list, years_list, do_dams, do_nfse, do_das):
        super().__init__()
        self.target_list = target_list
        self.years_list = years_list # Expect list
        self.do_dams = do_dams
        self.do_nfse = do_nfse
        self.do_das = do_das # <--- NEW PARAMETER
        
    def run(self):
        try:
            from app.ferramentas.db_extractor import run_db_extraction
            # Added self.do_das to the function call
            run_db_extraction(
                self.target_list, 
                self.years_list, 
                self.do_dams, 
                self.do_nfse, 
                self.do_das, 
                self.progress
            )
            self.finished.emit()
        except Exception:
            self.error.emit(f"❌ Erro durante a Extração do Banco de Dados:\n{traceback.format_exc()}")
            
class SituacaoExtractorWorker(BaseWorker):
    finished = Signal()
    def __init__(self, tasks):
        super().__init__()
        self.tasks = tasks
    def run(self):
        try:
            from app.ferramentas.situacao_extractor import run_situacao_extractor
            run_situacao_extractor(self.tasks, self.progress)
            self.finished.emit()
        except Exception:
            self.error.emit(f"❌ Erro durante a Extração de Situação (RPA):\n{traceback.format_exc()}")

class ValidationExtractorWorker(BaseWorker):
    finished = Signal(dict) 
    def __init__(self, tasks):
        super().__init__()
        self.tasks = tasks
    def run(self):
        results = {}
        try:
            from app.ferramentas.extractor_full import process_company
            for i, task in enumerate(self.tasks):
                if self.check_stop(): break
                imu = task['imu']
                year = task['year']
                expected = task.get('expected_value')
                self.progress.emit(f"🔍 Validando {year} (Exp: {expected})...")
                val = process_company(imu, year, expected_value=expected)
                if val is not None: results[str(year)] = val
                else: results[str(year)] = 0.0
            self.finished.emit(results)
        except Exception:
            self.error.emit(f"❌ Erro Crítico:\n{traceback.format_exc()}")

class AutomaticIDDWorker(BaseWorker):
    finished = Signal(dict) 
    def __init__(self, imu, year, expected_value, output_folder=None): 
        super().__init__()
        self.imu = imu
        self.year = year
        self.expected_value = expected_value
        self.output_folder = output_folder
    def run(self):
        try:
            from app.ferramentas.extractor_full import process_company
            self.progress.emit(f"🚀 Iniciando Processo Automático para {self.year}...")
            self.progress.emit(f"💰 Valor Esperado: R$ {self.expected_value:,.2f}")
            if self.output_folder: self.progress.emit(f"📂 Salvando em: {self.output_folder}")
            result = process_company(self.imu, self.year, expected_value=self.expected_value, run_emission=True, output_folder=self.output_folder)
            self.progress.emit(f"📊 Resultado da Extração: {result.get('status')}")
            if result.get("status") == "Success":
                self.progress.emit("✅ Emissão realizada com sucesso!")
                self.progress.emit(f"📝 IDD: {result.get('idd_number')} | Protocolo: {result.get('protocolo')}")
            else:
                self.progress.emit(f"⚠️ Processo finalizado com status: {result.get('status')}")
            self.finished.emit(result)
        except Exception:
            self.error.emit(f"❌ Erro no Processo Automático:\n{traceback.format_exc()}")

# --- HEADLESS CALCULATOR (UNCHANGED) ---
class HeadlessTaxCalculator:
    def __init__(self, all_invoices_df, infraction_groups, dam_file_path=None):
        self.all_invoices_df = all_invoices_df
        self.autos = {}
        self.auto_counter = 1
        self.dam_payments_map = {}
        self.pgdas_payments_map = {}
        if dam_file_path and os.path.exists(dam_file_path):
            try: self.dam_payments_map = _load_and_process_dams(dam_file_path)
            except: pass
        self.motive_to_rule_map = {
             'IDD (Não Pago)': 'idd_nao_pago', 'Dedução indevida': 'deducao_indevida',
             'Alíquota Incorreta': 'aliquota_incorreta', 'Regime incorreto': 'regime_incorreto',
             'Isenção/Imunidade Indevida': 'isencao_imunidade_indevida',
             'Natureza da Operação Incompatível': 'natureza_operacao_incompativel',
             'Benefício Fiscal incorreto': 'beneficio_fiscal_incorreto',
             'Local da incidência incorreto': 'local_incidencia_incorreto',
             'Retenção na Fonte (Verificar)': 'retencao_na_fonte_a_verificar'
        }
        for group_name, group_df in infraction_groups.items():
            auto_id = f"AUTO-{self.auto_counter:03d}"
            base_motive = group_name.split(' (')[0]
            rule_name = self.motive_to_rule_map.get(base_motive, 'regra_desconhecida')
            self.autos[auto_id] = {
                'motive': group_name, 
                'df': group_df.copy(),
                'rule_name': rule_name,
                'user_defined_aliquota': None, 
                'user_defined_credito': None
            }
            self.auto_counter += 1

    def calculate_context(self):
        # ... (Same logic as in original file, reused here) ...
        final_data = {}
        for auto_id, auto_data in self.autos.items():
            df = auto_data.get('df')
            if df is None or df.empty: continue
            invoice_list = df.index.tolist()
            first_row = df.iloc[0]
            correct_aliquota_val = auto_data.get('user_defined_aliquota')
            if correct_aliquota_val is None:
                if 'correct_rate' in df.columns and pd.notna(first_row.get('correct_rate')):
                    correct_aliquota_val = first_row.get('correct_rate')
            correct_aliquota_str = f"{correct_aliquota_val:.2f}" if correct_aliquota_val is not None else "5.00"
            final_data[auto_id] = {
                'invoices': invoice_list,
                'auto_id': auto_id,
                'rule_name': auto_data['rule_name'],
                'motive_text': auto_data['motive'],
                'correct_aliquota': correct_aliquota_str,
                'user_defined_credito': auto_data.get('user_defined_credito'),
                'is_split_diff': False,
                'auto_text': '',
                'monthly_overrides': {}
            }
        
        all_valid_dates = []
        for auto_info in final_data.values():
            invoice_indices = auto_info.get('invoices', [])
            if not invoice_indices: continue
            df_inv = self.all_invoices_df.loc[invoice_indices].copy()
            df_inv['DATA EMISSÃO'] = pd.to_datetime(df_inv['DATA EMISSÃO'], errors='coerce')
            all_valid_dates.extend(df_inv['DATA EMISSÃO'].dropna().tolist())

        all_periods_list = []
        if all_valid_dates:
            min_year = min(d.year for d in all_valid_dates)
            max_year = max(d.year for d in all_valid_dates)
            for year in range(min_year, max_year + 1):
                for month in range(1, 13):
                    all_periods_list.append(pd.Period(year=year, month=month, freq='M'))
            all_periods_list.sort()
        else:
            current_year = datetime.now().year
            for month in range(1, 13):
               all_periods_list.append(pd.Period(year=current_year, month=month, freq='M'))

        available_credits = {'DAM': copy.deepcopy(self.dam_payments_map), 'PGDAS': {k: v[0] for k, v in self.pgdas_payments_map.items()}}
        autos_context = []
        
        for auto_key, auto_info in final_data.items():
            invoice_indices = auto_info.get('invoices', [])
            df_invoices = self.all_invoices_df.loc[invoice_indices].copy()
            df_invoices['DATA EMISSÃO'] = pd.to_datetime(df_invoices['DATA EMISSÃO'], errors='coerce')
            df_invoices['_month_str'] = df_invoices['DATA EMISSÃO'].dt.strftime('%m/%Y')
            
            try:
                aliquota_str = auto_info.get('correct_aliquota', '0.0')
                default_aliquota_pct = float(aliquota_str)
            except: default_aliquota_pct = 0.0

            if not df_invoices.empty:
                for idx, row in df_invoices.iterrows():
                    if 'correct_rate' in df_invoices.columns and pd.notna(row['correct_rate']) and row['correct_rate'] > 0:
                        target = row['correct_rate']
                    elif row.get('ALÍQUOTA', 0) > 0:
                        target = row.get('ALÍQUOTA')
                    else:
                        target = default_aliquota_pct
                    df_invoices.at[idx, '_target_rate_group'] = float(target)
            
            dados_anuais = []
            total_iss_liquido_auto = 0.0; total_iss_op = 0.0; total_iss_bruto_auto = 0.0
            
            for period_key in all_periods_list:
                period_str_mm_yyyy = period_key.strftime('%m/%Y')
                period_str_m_yyyy = f"{period_key.month}/{period_key.year}"
                mask = (df_invoices['DATA EMISSÃO'].dt.to_period('M') == period_key)
                df_month = df_invoices[mask]
                if df_month.empty: continue
                
                unique_rates = df_month['_target_rate_group'].unique()
                for rate_val in unique_rates:
                    group = df_month[df_month['_target_rate_group'] == rate_val]
                    if group.empty: continue
                    val_col = group['VALOR'].fillna(0.0)
                    ded_col = group.get('VALOR DEDUÇÃO', 0.0).fillna(0.0)
                    base_calculo = (val_col - ded_col).sum()
                    rate_dec = rate_val / 100.0
                    
                    iss_liquido_calc = 0.0; iss_correto_bruto = 0.0; iss_declarado_pago = 0.0
                    for _, row in group.iterrows():
                        v = row.get('VALOR', 0.0) - row.get('VALOR DEDUÇÃO', 0.0)
                        decl_rate = row.get('ALÍQUOTA', 0.0) / 100.0
                        is_paid = str(row.get('PAGAMENTO', '')).strip().lower() in ['sim', 'idd']
                        iss_correto_bruto += (v * rate_dec)
                        paid_amt = (v * decl_rate) if is_paid else 0.0
                        iss_declarado_pago += paid_amt
                        if is_paid: iss_liquido_calc += max(0, (rate_dec - decl_rate) * v)
                        else: iss_liquido_calc += (rate_dec * v)

                    dams_list = available_credits['DAM'].get(period_str_m_yyyy, [])
                    available_dam_total = sum(d['val'] for d in dams_list)
                    dam_utilizado = min(iss_liquido_calc, available_dam_total)
                    rem = dam_utilizado; codes = []
                    for d_obj in dams_list:
                        if rem <= 0.0001: break
                        if d_obj['val'] > 0:
                            deduct = min(d_obj['val'], rem)
                            d_obj['val'] -= deduct
                            rem -= deduct
                            codes.append(d_obj['code'])
                    dam_ident = ", ".join(sorted(set(codes))) if codes else "-"
                    iss_op = max(0, iss_liquido_calc - dam_utilizado)
                    total_iss_liquido_auto += iss_liquido_calc
                    total_iss_op += iss_op
                    total_iss_bruto_auto += iss_correto_bruto
                    
                    aliquota_op_display = f'{rate_val:.2f}%'
                    if base_calculo > 0.001: aliquota_declarada_display = f'{(iss_declarado_pago / base_calculo) * 100.0:.2f}%'
                    else: aliquota_declarada_display = "-"

                    dados_anuais.append({
                        'mes_ano': period_str_mm_yyyy,
                        'base_calculo': base_calculo,
                        'aliquota_op': aliquota_op_display,
                        'iss_apurado_bruto': iss_correto_bruto,
                        'aliquota_declarada': aliquota_declarada_display,
                        'iss_declarado_pago': iss_declarado_pago,
                        'base_calculo_op': base_calculo,
                        'iss_apurado': iss_liquido_calc,
                        'iss_apurado_op': iss_op,
                        'dam_iss_pago': dam_utilizado,
                        'dam_identificacao': dam_ident,
                        'das_iss_pago': 0.0, 'das_identificacao': "-", 'das_aliquota': "-", 'dam_aliquota': "-"
                    })

            autos_context.append({
                'numero': auto_key,
                'motive_text': auto_info['motive_text'],
                'dados_anuais': dados_anuais,
                'totais': {
                    'iss_apurado': total_iss_liquido_auto,
                    'iss_apurado_op': total_iss_op,
                    'iss_apurado_bruto': total_iss_bruto_auto,
                    'base_calculo': total_iss_bruto_auto / (default_aliquota_pct/100) if default_aliquota_pct else 0,
                    'base_calculo_op': total_iss_bruto_auto / (default_aliquota_pct/100) if default_aliquota_pct else 0,
                    'das_iss_pago': 0.0, 'dam_iss_pago': 0.0, 'iss_declarado_pago': 0.0
                }
            })

        summary_autos_list = []
        total_geral = 0.0
        nfs_map = {}
        for k, v in final_data.items():
            ids = self.all_invoices_df.loc[v['invoices'], Columns.INVOICE_NUMBER].astype(str).tolist() if v['invoices'] else []
            nfs_map[k] = ids

        for auto in autos_context:
            val = auto['totais']['iss_apurado_op']
            if val > 0.01:
                summary_autos_list.append({
                    'numero': auto['numero'],
                    'nfs_tributadas': format_invoice_numbers(nfs_map.get(auto['numero'], [])), 
                    'iss_valor_original': auto.get('totais', {}).get('iss_apurado_bruto', 0.0), 
                    'total_credito_tributario': val,
                    'motivo': auto.get('motive_text', '')
                })
                total_geral += val
        
        return {
            'autos': autos_context,
            'summary': { 'autos': summary_autos_list, 'multa': None, 'total_geral_credito': total_geral }
        }

    def get_final_data(self):
        data = {}
        for auto_id, auto_data in self.autos.items():
            df = auto_data.get('df')
            if df is None: continue
            correct_aliquota_val = auto_data.get('user_defined_aliquota')
            if correct_aliquota_val is None and not df.empty:
                first = df.iloc[0]
                if 'correct_rate' in df.columns: correct_aliquota_val = first.get('correct_rate')
            str_rate = f"{correct_aliquota_val:.2f}" if correct_aliquota_val else "5.00"
            data[auto_id] = {
                'invoices': df.index.tolist(),
                'auto_id': auto_id,
                'rule_name': auto_data['rule_name'],
                'motive_text': auto_data['motive'],
                'correct_aliquota': str_rate,
                'user_defined_credito': None,
                'is_split_diff': False,
                'auto_text': '',
                'monthly_overrides': {}
            }
        return data


class BatchIDDWorker(BaseWorker):
    """
    Updated Worker: Handles Divergence Loop Logic + Keys Update + ROBUST RETRY
    """
    finished = Signal(dict)

    def __init__(self, tasks, master_path, auditor_data):
        super().__init__()
        self.tasks = tasks 
        self.master_path = master_path
        self.auditor_data = auditor_data 

    def run(self):
        results = {}
        import traceback
        import os
        from datetime import datetime
        import glob
        
        # Lazy imports
        try:
            import pandas as pd
            from main import load_and_prepare_invoices, perform_rules_analysis
            from app.ferramentas.extractor_full import process_company
            from app.generation_task import _generate_final_documents_task
        except ImportError as e:
            self.error.emit(f"Erro de Importação: {e}")
            return

        total_tasks = len(self.tasks)
        
        for i, task in enumerate(self.tasks):
            if self.check_stop(): 
                self.progress.emit("🛑 Processo interrompido pelo usuário.")
                break
            
            imu = task.get('imu')
            cnpj = task.get('cnpj')
            name = task.get('name')
            inv_path = task.get('invoices_path')
            dam_path = task.get('dam_path')
            target_folder = task.get('folder_path')
            
            self.progress.emit(f"\n🚀 [{i+1}/{total_tasks}] Processando: {name} ({imu})...")
            
            # --- Temp file variable to ensure cleanup ---
            temp_pickle_path = None
            
            # --- Flag to control strict sequence ---
            step_success = False 

            try:
                # 1. Load Data
                self.progress.emit(f"   📂 Lendo notas...")
                prepped_df = load_and_prepare_invoices(self.master_path, inv_path, cnpj)
                
                if prepped_df.empty:
                    self.progress.emit("   ⛔ ERRO: Sem notas válidas. Interrompendo lote.")
                    results[imu] = "Falha: Sem Notas"
                    break # STRICT STOP

                # 2. Rules
                self.progress.emit("   ⚙️ Analisando regras (Modo IDD)...")
                infraction_groups, df_analyzed = perform_rules_analysis(prepped_df, idd_mode=True)

                # 3. Headless Calc
                self.progress.emit("   🧮 Calculando (Headless)...")
                calculator = HeadlessTaxCalculator(df_analyzed, infraction_groups, dam_path)
                preview_context = calculator.calculate_context()

                expected_val = preview_context['summary'].get('total_geral_credito', 0.0)
                
                years = []
                for auto in preview_context.get('autos', []):
                    for m in auto.get('dados_anuais', []):
                        try: years.append(m['mes_ano'].split('/')[1])
                        except: pass
                target_year = max(set(years), key=years.count) if years else str(datetime.now().year)

                self.progress.emit(f"   💰 Valor: R$ {expected_val:,.2f} (Ano {target_year})")

                # 4. RPA (IDD Emission & PDF Saving)
                self.progress.emit("   🤖 Executando RPA (Emissão e Download)...")
                rpa_result = process_company(
                    imu, 
                    target_year, 
                    expected_value=expected_val, 
                    run_emission=True,
                    output_folder=target_folder 
                )
                
                status = rpa_result.get("status")
                
                # --- CHECK 1: HANDLE DIVERGENCE (CONTINUE LOOP) ---
                if status == "Divergence":
                    self.progress.emit(f"   ⚠️ Divergência de valores detectada.")
                    self.progress.emit(f"   ⏭️ Pulando empresa {name} e continuando...")
                    results[imu] = "Pulado (Divergência)"
                    continue # Skip to next company in loop 

                # --- STRICT CHECK: DID RPA SUCCEED? ---
                if status != "Success":
                    self.progress.emit(f"   ⛔ PARADA DE EMERGÊNCIA: Falha no RPA (Status: {status}).")
                    self.progress.emit("   ⚠️ Motivo: DAM ou Comunicado não foram salvos ou valor divergiu.")
                    results[imu] = f"Falha RPA: {status}"
                    break # STRICT STOP if generic failure

                # If we are here, RPA worked (Files are on disk)
                idd_num = rpa_result.get("idd_number")
                protocolo = rpa_result.get("protocolo")
                self.progress.emit(f"   ✅ RPA Concluído! IDD: {idd_num}")
                
                # 5. Generate PDF (Informação Fiscal)
                self.progress.emit("   📄 Gerando Informação Fiscal (Word/PDF)...")
                
                # --- UPDATE CONTEXT ---
                preview_context['epaf_numero'] = protocolo
                for auto in preview_context.get('autos', []):
                    auto['numero'] = idd_num
                for summary_auto in preview_context.get('summary', {}).get('autos', []):
                    summary_auto['numero'] = idd_num

                raw_final_data = calculator.get_final_data()
                new_final_data = {}
                for key, val in raw_final_data.items():
                    val['auto_id'] = idd_num
                    new_final_data[idd_num] = val 
                
                # Create Temporary Pickle File
                # NOTE: Re-creating pickle inside loop if retrying is safer, 
                # but here we do it once before retry loop.
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pkl") as tmp:
                        df_analyzed.to_pickle(tmp.name)
                        temp_pickle_path = tmp.name
                except Exception as e:
                    self.progress.emit(f"   ⛔ Erro criando pickle temp: {e}")
                    results[imu] = "Erro Temp File"
                    break

                class MockQueue:
                    def put(self, msg): pass
                
                # --- GENERATION RETRY LOOP (Max 3 attempts) ---
                gen_success = False
                MAX_GEN_RETRIES = 3
                
                for gen_attempt in range(1, MAX_GEN_RETRIES + 1):
                    try:
                        self.progress.emit(f"   🔄 Tentativa de Geração {gen_attempt}/{MAX_GEN_RETRIES}...")
                        
                        # Verify pickle exists before call
                        if not os.path.exists(temp_pickle_path):
                            # Re-dump if missing (weird edge case)
                            df_analyzed.to_pickle(temp_pickle_path)

                        _generate_final_documents_task(
                            MockQueue(),
                            cnpj,
                            self.master_path,
                            new_final_data, 
                            preview_context, 
                            temp_pickle_path,
                            "", 
                            dam_path,
                            "", 
                            target_folder,
                            "AR", 
                            imu,
                            idd_mode=True 
                        )
                        
                        # --- STRICT CHECK 2: DID INFO FISCAL FILE APPEAR? ---
                        expected_pattern = os.path.join(target_folder, "*Informacao_Fiscal*")
                        files_found = glob.glob(expected_pattern)
                        
                        if not files_found:
                            raise Exception("Arquivo 'Informação Fiscal' não encontrado após execução.")
                        
                        gen_success = True
                        break # Success!
                    
                    except Exception as loop_error:
                        self.progress.emit(f"      ⚠️ Falha na geração (Tentativa {gen_attempt}): {loop_error}")
                        time.sleep(2.0) # Wait before retry

                if not gen_success:
                    self.progress.emit(f"   ⛔ PARADA DE EMERGÊNCIA: Geração falhou após {MAX_GEN_RETRIES} tentativas.")
                    results[imu] = "Falha Geração Docs"
                    break # STRICT STOP

                # If we got here, EVERYTHING is perfect.
                results[imu] = "Sucesso"
                step_success = True
                self.progress.emit(f"   🏆 Empresa {name} finalizada com sucesso. Avançando...")

            except Exception as e:
                self.progress.emit(f"   ⛔ CRÍTICO: Exceção não tratada: {e}")
                results[imu] = f"Erro Exceção"
                logging.error(f"Batch Error: {traceback.format_exc()}")
                break # STRICT STOP
            
            finally:
                if temp_pickle_path and os.path.exists(temp_pickle_path):
                    try: os.remove(temp_pickle_path)
                    except: pass
            
            # Final safety check: if we didn't mark step_success, break loop
            if not step_success:
                self.progress.emit("   🛑 Interrompendo o lote devido a falha na empresa atual.")
                break

        self.finished.emit(results)


class DeckerWorker(BaseWorker):
    def __init__(self, tasks, base_directory, credentials):
        super().__init__()
        self.tasks = tasks
        self.base_directory = base_directory
        self.credentials = credentials

    def run(self):
        try:
            from app.ferramentas.decker import run_decker_sender
            run_decker_sender(self.tasks, self.base_directory, self.credentials, self.progress)
            self.finished.emit()
        except Exception:
            self.error.emit(f"❌ Erro durante o envio de emails (Decker):\n{traceback.format_exc()}")

class AnalysisScannerWorker(BaseWorker):
    """
    Scans a root folder for 'Relatorio_NFSE_{IMU}_{YEAR}.xlsx' files,
    runs the rules engine + headless calculator for each,
    and produces a consolidated Excel report of potential debts (IDD/Auto).
    
    ✅ UPDATED: Now performs data cleaning and DECADENCE CHECKS exactly like the Main Window.
    """
    finished = Signal(str) # Returns path to the generated report

    def __init__(self, root_folder):
        super().__init__()
        self.root_folder = root_folder

    def run(self):
        try:
            import os
            import re
            import pandas as pd
            from main import perform_rules_analysis
            from app.constants import Columns
            from datetime import datetime

            results = []
            files_to_process = []

            # 1. Scan for files first
            self.progress.emit("🔍 Escaneando diretórios...")
            for root, dirs, files in os.walk(self.root_folder):
                for file in files:
                    if self.check_stop(): break
                    if file.startswith("Relatorio_NFSE_") and file.endswith(".xlsx"):
                        core = file.replace("Relatorio_NFSE_", "").replace(".xlsx", "")
                        parts = core.split('_')
                        
                        if len(parts) >= 2:
                            imu = parts[0]
                            year = parts[-1]
                            if len(year) == 4 and year.isdigit():
                                files_to_process.append({
                                    'path': os.path.join(root, file),
                                    'imu': imu,
                                    'year': year,
                                    'folder': os.path.basename(root)
                                })

            total_files = len(files_to_process)
            if total_files == 0:
                self.error.emit("❌ Nenhum arquivo 'Relatorio_NFSE_*.xlsx' encontrado.")
                return

            self.progress.emit(f"📄 Encontrados {total_files} arquivos para análise.")

            # 2. Process each file
            for i, task in enumerate(files_to_process):
                if self.check_stop(): break
                
                path = task['path']
                imu = task['imu']
                year = task['year']
                folder_name = task['folder']
                
                self.progress.emit(f"⚙️ [{i+1}/{total_files}] Analisando: {folder_name} (Ano {year})...")
                
                try:
                    # ==========================================================
                    # ⚠️  EXACT DATA LOADING REPLICATION (Main Window Logic) ⚠️
                    # ==========================================================
                    
                    # A. Load DataFrame with skiprows=2 (Standard GTM export format)
                    df = pd.read_excel(path, engine='openpyxl', skiprows=2)
                    
                    if df.empty:
                        results.append(self._make_error_row(imu, year, folder_name, "Arquivo Vazio (sem dados)"))
                        continue

                    # B. Column Cleaning & Type Conversion
                    # Matches load_and_prepare_invoices in main.py
                    numeric_cols = ['VALOR', 'VALOR DEDUÇÃO', 'ALÍQUOTA', 'DESCONTO INCONDICIONAL']
                    for col in numeric_cols:
                        if col in df.columns:
                            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                        else:
                            df[col] = 0.0 # Ensure column exists

                    # C. Date Conversion
                    if 'DATA EMISSÃO' in df.columns:
                        df['DATA EMISSÃO'] = pd.to_datetime(df['DATA EMISSÃO'], errors='coerce')
                    if 'DT. CANCELAMENTO' in df.columns:
                        df['DT. CANCELAMENTO'] = pd.to_datetime(df['DT. CANCELAMENTO'], errors='coerce')

                    # D. Remove Cancelled Invoices
                    if 'DT. CANCELAMENTO' in df.columns:
                        df = df[pd.isna(df['DT. CANCELAMENTO'])]
                        
                    if df.empty:
                         results.append(self._make_success_row(imu, year, folder_name, 0.0, ["Todas Canceladas"], 0, 0.0))
                         continue

                    # E. Apply Discount Logic (Net Value)
                    if 'VALOR' in df.columns and 'DESCONTO INCONDICIONAL' in df.columns:
                        df['VALOR'] = df['VALOR'] - df['DESCONTO INCONDICIONAL']

                    # F. Clean Activity Code (Standardize to 4 digits)
                    if 'CÓDIGO DA ATIVIDADE' in df.columns:
                        df['CÓDIGO DA ATIVIDADE'] = (df['CÓDIGO DA ATIVIDADE'].astype(str)
                                                                .str.replace(r'\D', '', regex=True)
                                                                .str.strip()
                                                                .str.pad(4, side='left', fillchar='0'))
                    
                    # G. Calculate Status Legal (DECADENCE)
                    # This is crucial so that 'rules_engine' skips older paid invoices.
                    df['status_legal'] = 'OK'
                    
                    if 'PAGAMENTO' not in df.columns: 
                        df['PAGAMENTO'] = 'Não'
                    df['PAGAMENTO'] = df['PAGAMENTO'].fillna('Não').astype(str)
                    
                    is_paid_mask = df['PAGAMENTO'].str.strip().str.lower().isin(['sim', 'idd'])
                    
                    if 'DATA EMISSÃO' in df.columns:
                        today = pd.to_datetime(datetime.now().date())
                        paid_invoices_mask = is_paid_mask & df['DATA EMISSÃO'].notna()
                        
                        if paid_invoices_mask.any():
                             # Art 150: 5 years from payment (approx invoice date here)
                             # Matches main.py logic: MonthEnd(0) + 5 Years
                             cutoff_dates_paid = df.loc[paid_invoices_mask, 'DATA EMISSÃO'] + pd.offsets.MonthEnd(0) + pd.DateOffset(years=5)
                             mask_decadente_paid = today > cutoff_dates_paid
                             df.loc[mask_decadente_paid[mask_decadente_paid].index, 'status_legal'] = 'Decadente_Pago'

                    # ==========================================================
                    # ⚠️  END OF REPLICATION ⚠️
                    # ==========================================================

                    # D. Run Rules (idd_mode=False for analytical scan)
                    infraction_groups, df_analyzed = perform_rules_analysis(df, idd_mode=False)
                    
                    if not infraction_groups:
                        results.append(self._make_success_row(imu, year, folder_name, 0.0, [], 0, 0.0))
                        continue

                    # E. Calculate Potentials
                    calculator = HeadlessTaxCalculator(df_analyzed, infraction_groups, dam_file_path=None)
                    context = calculator.calculate_context()
                    
                    summary = context.get('summary', {})
                    total_credito = summary.get('total_geral_credito', 0.0)
                    autos_summary = summary.get('autos', [])
                    
                    idd_amount = 0.0
                    motives = []
                    
                    for auto in autos_summary:
                        val = auto.get('total_credito_tributario', 0.0)
                        motive = auto.get('motivo', '')
                        motives.append(f"{motive} (R$ {val:.2f})")
                        
                        if 'IDD' in motive or 'Não Pago' in motive:
                            idd_amount += val

                    motives_str = "; ".join(motives)
                    num_autos = len(autos_summary)

                    results.append(self._make_success_row(
                        imu, year, folder_name, total_credito, 
                        motives_str, num_autos, idd_amount
                    ))

                except Exception as e:
                    import traceback
                    logging.error(f"Scanner Error on {path}: {traceback.format_exc()}")
                    results.append(self._make_error_row(imu, year, folder_name, str(e)))

            # 3. Save Report
            if results:
                self.progress.emit("💾 Salvando relatório final...")
                df_report = pd.DataFrame(results)
                
                output_filename = f"Relatorio_Analitico_Geral_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                output_path = os.path.join(self.root_folder, output_filename)
                
                # Format currency columns for easier reading
                try:
                    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                        df_report.to_excel(writer, index=False, sheet_name="Resumo")
                        sheet = writer.sheets["Resumo"]
                        # Auto-adjust column width (basic)
                        for column in sheet.columns:
                            sheet.column_dimensions[column[0].column_letter].width = 20
                except:
                    df_report.to_excel(output_path, index=False)

                self.finished.emit(output_path)
            else:
                self.error.emit("Nenhum resultado gerado.")

        except Exception as e:
            self.error.emit(f"❌ Erro Crítico no Scanner:\n{traceback.format_exc()}")

    def _make_success_row(self, imu, year, folder, total, motives, count, idd_val):
        return {
            'Pasta': folder,
            'IMU': imu,
            'Ano': int(year),
            'Status': 'Sucesso',
            'Potencial Total (R$)': total,
            'Potencial IDD (R$)': idd_val,
            'Qtd Autos': count,
            'Detalhes Infrações': motives if motives else "Nenhuma"
        }

    def _make_error_row(self, imu, year, folder, error_msg):
        return {
            'Pasta': folder,
            'IMU': imu,
            'Ano': int(year),
            'Status': 'Erro',
            'Potencial Total (R$)': 0.0,
            'Potencial IDD (R$)': 0.0,
            'Qtd Autos': 0,
            'Detalhes Infrações': error_msg
        }