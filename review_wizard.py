# --- FILE: app/review_wizard.py ---

import pandas as pd
import json
import os
from datetime import datetime
from PySide6.QtWidgets import (QDialog, QTabWidget, QDialogButtonBox, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QTableWidget,
                               QTableWidgetItem, QGroupBox, QFormLayout, 
                               QMessageBox, QListWidget, QListWidgetItem, 
                               QLineEdit, QHeaderView, QWidget,
                               QTextEdit, QCheckBox, QFileDialog, QLabel, 
                               QComboBox, QStyle, QScrollArea, QSpacerItem,
                               QMenu, QCompleter)
from PySide6.QtCore import Qt, QLocale, QStringListModel # ✅ Import QStringListModel
from PySide6.QtGui import QColor, QDoubleValidator, QAction
from PySide6.QtCore import QThread,  QSize
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                               QPushButton, QTableWidget, QTableWidgetItem, 
                               QHeaderView, QCheckBox, QMessageBox)
from PySide6.QtGui import QColor, QIcon
import traceback
import tempfile # <--- ADD THIS IMPORT AT THE TOP
# --- Local Application Imports ---
from .widgets import NumericTableWidgetItem, DateTableWidgetItem, ColumnSelectionDialog, SORT_ROLE
from .constants import Columns, SESSION_FILE_PREFIX
from document_parts import formatar_texto_multa
from .infraction_correction_dialog import InfractionCorrectionDialog
from .new_auto_dialog import NewAutoDialog
from .detail_viewer_dialog import InvoiceDetailViewerDialog
from .widgets import CollapsibleGroupBox
# --- Loaders for payment maps ---
from data_loader import _load_and_process_dams
from .pgdas_loader import _load_and_process_pgdas
# Import the new dialog
from .auto_text_dialog import AutoTextDialog
import logging
from statistics import mode # ✅ Import mode
from document_parts import formatar_texto_multa, format_invoice_numbers
import hashlib
import copy  # ✅ Required for deepcopy of list-based credits
from .workers import ValidationExtractorWorker # <--- Import the new worker
from app.excel_filter import FilterableHeaderView

from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import QTableView

class InvoiceTableModel(QAbstractTableModel):
    def __init__(self, df, visible_columns, parent=None):
        super().__init__(parent)
        self._df = df
        self._visible_columns = visible_columns
        
        # Pre-calculate column indices for speed
        self._col_indices = [df.columns.get_loc(c) for c in visible_columns]
        
        # Colors
        self.color_decadent = QColor(80, 80, 80)
        self.color_rule = QColor(100, 92, 63)
        self.color_ignored = QColor(220, 220, 220)
        self.color_text_ignored = QColor(100, 100, 100)

    def rowCount(self, parent=QModelIndex()):
        return len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return len(self._visible_columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row_idx = index.row()
        col_name = self._visible_columns[index.column()]

        # 1. Display Data (Text)
        if role == Qt.DisplayRole or role == Qt.ToolTipRole:
            val = self._df.iloc[row_idx][col_name]

            # 1. Handle Lists/Arrays (The cause of your crash)
            try:
                # If val is an array, this line raises ValueError
                if pd.isna(val):
                    return ""
            except ValueError:
                # If we catch the error, it means 'val' is an array/list.
                # Arrays are rarely "NA", so we treat it as valid data.
                pass
            
            if col_name == Columns.VALUE:
                return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            elif col_name in [Columns.RATE, Columns.CORRECT_RATE]:
                return f"{val:,.2f}"
            elif col_name == Columns.ISSUE_DATE:
                return val.strftime('%d/%m/%Y') if hasattr(val, 'strftime') else str(val)
            elif col_name == Columns.BROKEN_RULE_DETAILS:
                 return "; ".join(map(str, val)) if isinstance(val, list) else str(val)
            return str(val)

        # 2. Background Color (Status Logic)
        elif role == Qt.BackgroundRole:
            # We access the full row safely to check status columns
            # Note: We use the underlying dataframe index to get specific status columns
            # Using .iloc is faster for sequential access in views
            
            # Optimization: Check status columns only if they exist
            status_legal = self._df.iloc[row_idx].get(Columns.STATUS_LEGAL)
            if status_legal in ['Decadente', 'Prescrito']:
                return self.color_decadent

            # Check for ignored status (Manual)
            status_manual = self._df.iloc[row_idx].get('status_manual')
            if status_manual == 'Ignored':
                return self.color_ignored

            # Check for Broken Rules
            if col_name == Columns.BROKEN_RULE_DETAILS:
                val = self._df.iloc[row_idx][col_name]
                if val: # If list is not empty
                    return self.color_rule

        # 3. Foreground Color (Text Color)
        elif role == Qt.ForegroundRole:
             if self._df.iloc[row_idx].get('status_manual') == 'Ignored':
                 return self.color_text_ignored

        # 4. UserRole: Return the Original DataFrame Index (Vital for selection logic)
        elif role == Qt.UserRole:
            return self._df.index[row_idx]

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            col_name = self._visible_columns[section]
            # Simple map for nicer headers
            header_map = {
                Columns.INVOICE_NUMBER: "Nº Nota", Columns.ISSUE_DATE: "Data Emissão",
                Columns.VALUE: "Valor (R$)", Columns.RATE: "Alíq. Decl.",
                Columns.CORRECT_RATE: "Alíq. Corr.", Columns.SERVICE_DESCRIPTION: "Discriminação",
                Columns.STATUS_LEGAL: "Status Legal", Columns.BROKEN_RULE_DETAILS: "Infrações"
            }
            return header_map.get(col_name, col_name)
        return None

def normalize_dam_dataframe(df):
    """
    Standardizes column names, specifically stripping BOM artifacts (ï»¿)
    and mapping variations to 'codigoVerificacao'.
    """
    if df is None or df.empty:
        return df

    # 1. CRITICAL FIX: Clean column headers of BOM/garbage characters
    # This removes anything at the start that isn't a letter/number (like Ã¯Â»Â¿)
    df.columns = df.columns.astype(str).str.strip().str.replace(r'^[^\w]+', '', regex=True)

    col_map = {
        'codigoVerificacao': 'codigoVerificacao',
        'Código Verificação': 'codigoVerificacao',
        'Codigo Verificacao': 'codigoVerificacao',
        'Código de Verificação': 'codigoVerificacao',
        'Codigo': 'codigoVerificacao',
        'Verificação': 'codigoVerificacao',
        
        'referenciaPagamento': 'referenciaPagamento',
        'Competência': 'referenciaPagamento',
        'Competencia': 'referenciaPagamento',
        'Referência': 'referenciaPagamento',
        
        'receita': 'receita',
        'Receita': 'receita',
        
        'totalRecolher': 'totalRecolher',
        'Valor': 'totalRecolher',
        'Valor Pago': 'totalRecolher',
        'Valor Total': 'totalRecolher',
        'Total Recolher': 'totalRecolher',
        'totalRecolher': 'totalRecolher'
    }
    
    rename_dict = {}
    for col in df.columns:
        col_clean = str(col).strip()
        if col_clean in col_map:
            rename_dict[col] = col_map[col_clean]
        else:
            for k, v in col_map.items():
                if k.lower() == col_clean.lower():
                    rename_dict[col] = v
                    break
    
    if rename_dict:
        df.rename(columns=rename_dict, inplace=True)
        
    return df

def compute_dataframe_hash(df):
    """
    Creates a unique hash for the DataFrame to ensure session validity.
    We hash the index and the 'VALOR' column to detect if the dataset changed.
    """
    try:
        if df is None or df.empty:
            return "empty"
        
        # Create a string representation of key characteristics
        # Index length, first index, last index, and sum of values
        fingerprint = f"{len(df)}_{df.index[0]}_{df.index[-1]}_{df[Columns.VALUE].sum():.4f}"
        return hashlib.md5(fingerprint.encode('utf-8')).hexdigest()
    except Exception:
        # Fallback if something complex fails
        return "hash_error"

def format_summary_label(df, value_col=Columns.VALUE, iss_value=None, iss_label=None):
    """
    Formats the summary label, now with an optional ISS value.
    """
    if df is None or df.empty:
        return "Notas: 0 | Valor Total: R$ 0,00"
    
    count = len(df)
    total_value = df[value_col].sum()
    
    # Helper to format currency
    def fmtd(val):
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    label = f"Notas: {count} | Valor Total: R$ {fmtd(total_value)}"
    
    if iss_value is not None and iss_label is not None:
        label += f" | {iss_label}: R$ {fmtd(iss_value)}"
        
    return label

class ReviewWizard(QDialog):
    def __init__(self, all_invoices_df, infraction_groups, company_cnpj, 
                 company_razao_social, company_imu, company_epaf_initial, 
                 parent=None,dam_file_path=None):
        super().__init__(parent)
        self.setWindowTitle("Módulo de Revisão e Geração de Autos")
        self.setMinimumSize(1100, 650)
        self.setWindowState(Qt.WindowState.WindowMaximized)
        
        self.current_data_hash = compute_dataframe_hash(all_invoices_df)
        
        main_layout = QVBoxLayout(self)

        self.all_invoices_df = all_invoices_df
        # Ensure Matcher ID exists
        if Columns.INVOICE_NUMBER in self.all_invoices_df.columns:
            self.all_invoices_df['_matcher_id'] = self.all_invoices_df[Columns.INVOICE_NUMBER].astype(str).str.strip()
        else:
            self.all_invoices_df['_matcher_id'] = self.all_invoices_df.index.astype(str)

        self.company_cnpj = company_cnpj
        self.autos = {}
        self.auto_counter = 1
        self.company_razao_social = company_razao_social
        self.company_imu = company_imu
        self.company_epaf_initial = company_epaf_initial

        self.dam_payments_map = {}; self.pgdas_payments_map = {}
        self.preview_context = {}; self.available_credits_map = {} 
        self.fine_text_final = ""; self.fine_value_final = ""

        # ✅ NEW: Dirty Flag (Controls when to re-run heavy calculations)
        self.context_dirty = True 

        self.visible_columns = [
            Columns.INVOICE_NUMBER, Columns.ISSUE_DATE, Columns.VALUE, Columns.RATE,
            Columns.BROKEN_RULE_DETAILS, Columns.CORRECT_RATE, 
            Columns.SERVICE_DESCRIPTION, 'PAGAMENTO', 'NATUREZA DA OPERAÇÃO', Columns.STATUS_LEGAL
        ]
        self.motive_to_rule_map = {
             'IDD (Não Pago)': 'idd_nao_pago', 
             'IDD': 'idd_nao_pago', 
             'Dedução indevida': 'deducao_indevida',
             'Alíquota Incorreta': 'aliquota_incorreta', 
             'Alíquota': 'aliquota_incorreta', 
             'Regime incorreto': 'regime_incorreto',
             'Isenção/Imunidade Indevida': 'isencao_imunidade_indevida',
             'Natureza da Operação Incompatível': 'natureza_operacao_incompativel',
             'Benefício Fiscal incorreto': 'beneficio_fiscal_incorreto',
             'Local da incidência incorreto': 'local_incidencia_incorreto',
             'Retenção na Fonte (Verificar)': 'retencao_na_fonte_a_verificar'
        }
        
        sanitized_cnpj = "".join(filter(str.isdigit, self.company_cnpj))
        self.session_filepath = f"{SESSION_FILE_PREFIX}{sanitized_cnpj}.json"

        # 2. Create Pages
        self.assignment_page = AssignmentPage(self)
        self.fine_page = FineDetailsPage(self)
        self.preview_page = PreviewPage(self)
        self.confirm_page = ConfirmationPage(self)

        # 3. Create Tab Widget
        self.tab_widget = QTabWidget()
        
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #4C566A; background: #2E3440; }
            QTabBar::tab { background: #2E3440; color: #D8DEE9; border: 1px solid #4C566A; padding: 8px 20px; margin-right: 2px; }
            QTabBar::tab:selected { background: #5E81AC; color: #FFFFFF; font-weight: bold; border-bottom: 2px solid #88C0D0; }
            QTabBar::tab:!selected:hover { background: #3B4252; }
        """)

        self.tab_widget.addTab(self.assignment_page, "1. Atribuição de Notas")
        self.tab_widget.addTab(self.fine_page, "2. Detalhes (Multa e Créditos)")
        self.tab_widget.addTab(self.preview_page, "3. Pré-visualização")
        self.tab_widget.addTab(self.confirm_page, "4. Confirmação e Geração")
        
        # ✅ NEW: Track the previous tab index to save data on exit
        self.last_tab_index = 0 
        
        # 4. Build Top Bar
        top_bar = QHBoxLayout()
        self.restore_btn = QPushButton("📂 Restaurar Sessão Anterior")
        self.restore_btn.setToolTip("Tenta recuperar o trabalho salvo anteriormente para este CNPJ.")
        self.restore_btn.clicked.connect(self.force_load_session)
        
        self.save_btn = QPushButton("💾 Salvar Sessão Agora")
        self.save_btn.clicked.connect(self.handle_manual_save)
        
        top_bar.addWidget(self.restore_btn)
        top_bar.addWidget(self.save_btn)
        top_bar.addStretch()
        
        main_layout.addLayout(top_bar)
        main_layout.addWidget(self.tab_widget)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Gerar Documentos")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        main_layout.addWidget(buttons)
        
        # 6. Connect Signals
        self.tab_widget.currentChanged.connect(self.on_page_changed)
        
        self.populate_from_groups(infraction_groups)

        self.assignment_page.refresh_all_tables()
        self.fine_page.update_fine_text()
        self.update_subtitle() 
        self.assignment_page.epaf_numero_edit.textChanged.connect(self.update_subtitle) 
        self.update_tab_states()

        if dam_file_path and os.path.exists(dam_file_path):
            self.fine_page.load_dam_file_programmatically(dam_file_path)

    def force_load_session(self):
        """Called by button click to explicitly restore."""
        self.load_session(silent=False)
        
    def update_tab_states(self):
        """Enables/disables tabs based on progress."""
        # Page 1 (Assignment) is always enabled
        
        # Check if Page 1 is "complete" (has at least one auto with notes)
        can_use_fine_page = True
        self.tab_widget.setTabEnabled(1, can_use_fine_page) # Enable Fine page
        self.tab_widget.setTabEnabled(2, can_use_fine_page) # Enable Preview page
        self.tab_widget.setTabEnabled(3, can_use_fine_page) # Enable Confirm page
        
        if not can_use_fine_page:
            self.tab_widget.setTabText(1, "2. Detalhes (Bloqueado)")
            self.tab_widget.setTabText(2, "3. Pré-visualização (Bloqueado)")
            self.tab_widget.setTabText(3, "4. Confirmação (Bloqueado)")
        else:
            self.tab_widget.setTabText(1, "2. Detalhes (Multa e Créditos)")
            self.tab_widget.setTabText(2, "3. Pré-visualização")
            self.tab_widget.setTabText(3, "4. Confirmação e Geração")

    def update_subtitle(self):
        """Updates the subtitle on all wizard pages with current company info."""
        # ✅ CHANGED: Read from AssignmentPage
        epaf_num = self.assignment_page.epaf_numero_edit.text() if hasattr(self, 'assignment_page') and self.assignment_page.epaf_numero_edit else self.company_epaf_initial
        subtitle_text = f"Empresa: {self.company_razao_social} | IMU: {self.company_imu} | ePAF: {epaf_num}"
        
        if hasattr(self, 'assignment_page') and hasattr(self.assignment_page, 'subtitle_label'):
            self.assignment_page.subtitle_label.setText(subtitle_text)
        if hasattr(self, 'fine_page') and hasattr(self.fine_page, 'subtitle_label'):
            self.fine_page.subtitle_label.setText(subtitle_text)
        if hasattr(self, 'preview_page') and hasattr(self.preview_page, 'subtitle_label'):
            self.preview_page.subtitle_label.setText(subtitle_text)
        if hasattr(self, 'confirm_page') and hasattr(self.confirm_page, 'subtitle_label'):
            self.confirm_page.subtitle_label.setText(subtitle_text)

    def _fmtd(self, val):
        """Helper to format currency in BRL (1.234,56)"""
        if not isinstance(val, (int, float)):
            try:
                val = float(val)
            except (ValueError, TypeError):
                return "0,00"
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _get_auto_year(self, auto_id):
        """
        Determines the year associated with an auto based on:
        1. The year explicitly mentioned in the motive (e.g., 'Alíquota (2021)')
        2. The year of the invoices already assigned to it.
        Returns None if no year is strictly bound yet.
        """
        if not auto_id or auto_id not in self.autos:
            return None
            
        auto_data = self.autos[auto_id]
        
        # 1. Try to parse from Motive string (created by RulesPrepWorker)
        # Looks for pattern "(20XX)"
        import re
        motive = auto_data.get('motive', '')
        match = re.search(r'\((\d{4})\)', motive)
        if match:
            return int(match.group(1))
            
        # 2. Try to infer from existing invoices in the DataFrame
        df = auto_data.get('df')
        if isinstance(df, pd.DataFrame) and not df.empty:
            if 'DATA EMISSÃO' in df.columns:
                try:
                    # Return the year of the first invoice found
                    first_date = pd.to_datetime(df['DATA EMISSÃO'].iloc[0], errors='coerce')
                    if pd.notna(first_date):
                        return first_date.year
                except:
                    pass
                    
        return None
    
    def mark_dirty(self):
        """Called by other pages when source data changes (autos created, invoices moved, etc)."""
        self.context_dirty = True

    def on_page_changed(self, tab_index):
        current_widget = self.tab_widget.widget(tab_index)
        
        # ✅ FIX: Save Preview Page data (Manual Edits) into self.autos when LEAVING it
        if self.last_tab_index == 2:
            if hasattr(self, 'preview_page'):
                # This reads the table and updates 'self.autos' (the source of truth)
                self.preview_page._read_tables_into_context()

        self.last_tab_index = tab_index
        
        if current_widget == self.preview_page:
            self.preview_page.initializePage() 
        elif current_widget == self.confirm_page:
            self.confirm_page.populate_confirmation()

    def populate_from_groups(self, infraction_groups):
        for group_name, group_df in infraction_groups.items():
            auto_id = f"AUTO-{self.auto_counter:03d}"
            
            base_motive = group_name.split(' (')[0]
            rule_name = self.motive_to_rule_map.get(base_motive, 'regra_desconhecida')

            self.autos[auto_id] = {
                'motive': group_name, 
                'df': group_df.copy(),
                'auto_text': '',
                'rule_name': rule_name 
            }
            self.auto_counter += 1
        self.assignment_page.refresh_autos_list()

    def handle_manual_save(self):
        try:
            self.save_session()
            QMessageBox.information(self, "Sessão Salva", "O seu progresso foi salvo com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar a sessão:\n{e}")

    def save_session(self):
        fines_list = self.fine_page.get_fines_data()
        
        session_data = {
            "timestamp": datetime.now().isoformat(),
            "autos": {},
            "multas": fines_list,
            "dam_filepath": self.fine_page.dam_filepath_edit.text(),
            "pgdas_folder_path": self.fine_page.pgdas_folder_path_edit.text(),
            "texto_multa": self.fine_page.fine_text_edit.toPlainText(),
            # ✅ CHANGED: Read from AssignmentPage
            "epaf_numero": self.assignment_page.epaf_numero_edit.text() 
        }
        
        for auto_id, data in self.autos.items():
            df = data.get('df')
            invoice_ids = [] 
            if isinstance(df, pd.DataFrame) and not df.empty:
                if Columns.INVOICE_NUMBER in df.columns:
                    invoice_ids = df[Columns.INVOICE_NUMBER].astype(str).str.strip().tolist()
            
            session_data["autos"][auto_id] = {
                "motive": data["motive"],
                "rule_name": data.get("rule_name", ""),
                "invoice_ids": invoice_ids,
                "user_defined_aliquota": data.get("user_defined_aliquota"),
                "user_defined_credito": data.get("user_defined_credito"),
                "auto_text": data.get("auto_text", ""),
                "monthly_overrides": data.get("monthly_overrides", {})
            }
            
        try:
            with open(self.session_filepath, 'w', encoding='utf-8') as f:
                json.dump(session_data, f, indent=4)
        except Exception as e:
            print(f"Error saving session: {e}")
            raise e

    def load_session(self, silent=False):
        if not os.path.exists(self.session_filepath):
            if not silent: QMessageBox.information(self, "Sem Sessão", "Não foi encontrado nenhum ficheiro de sessão salvo para este CNPJ.")
            return False

        try:
            with open(self.session_filepath, 'r', encoding='utf-8') as f:
                session_data = json.load(f)

            if not silent:
                ts = session_data.get("timestamp", "Data desconhecida")
                reply = QMessageBox.question(self, "Restaurar Sessão?",
                    f"Existe um trabalho salvo de {ts}.\n\nDeseja restaurá-lo?\nIsso substituirá a configuração atual.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes)
                if reply == QMessageBox.StandardButton.No:
                    return False
            else:
                if not session_data.get("autos"): return False

            self.autos = {}
            max_id_num = 0
            invoice_map = {str(row[Columns.INVOICE_NUMBER]).strip(): idx 
                           for idx, row in self.all_invoices_df.iterrows()}
            restored_count = 0
            
            for auto_id, data in session_data.get("autos", {}).items():
                saved_ids = data.get("invoice_ids", []) 
                valid_indices = []
                for inv_num in saved_ids:
                    if inv_num in invoice_map:
                        valid_indices.append(invoice_map[inv_num])
                
                auto_df = self.all_invoices_df.loc[valid_indices] if valid_indices else pd.DataFrame()
                
                self.autos[auto_id] = {
                    "motive": data["motive"],
                    "rule_name": data.get("rule_name"),
                    "df": auto_df,
                    "user_defined_aliquota": data.get("user_defined_aliquota"),
                    "user_defined_credito": data.get("user_defined_credito"),
                    "auto_text": data.get("auto_text", ""),
                    "monthly_overrides": data.get("monthly_overrides", {})
                }
                restored_count += len(valid_indices)
                try:
                    num = int(auto_id.replace("AUTO-", ""))
                    if num >= max_id_num: max_id_num = num
                except: pass

            self.auto_counter = max_id_num + 1
            
            fines_list = session_data.get("multas", [])
            if not fines_list and session_data.get("valor_multa"):
                 fines_list.append({
                     'year': datetime.now().year,
                     'number': session_data.get("numero_multa", ""),
                     'value': session_data.get("valor_multa", "")
                 })
            self.fine_page.set_fines_data(fines_list)

            self.assignment_page.epaf_numero_edit.setText(session_data.get("epaf_numero", self.company_epaf_initial))
            self.fine_page.fine_text_edit.setText(session_data.get("texto_multa", ""))
            
            dam_path = session_data.get("dam_filepath", "")
            pgdas_path = session_data.get("pgdas_folder_path", "")
            self.fine_page.dam_filepath_edit.setText(dam_path)
            self.fine_page.pgdas_folder_path_edit.setText(pgdas_path)
            
            if dam_path and os.path.exists(dam_path):
                try: self.dam_payments_map = _load_and_process_dams(dam_path)
                except: pass
            if pgdas_path and os.path.exists(pgdas_path):
                try:
                    class DummyEmitter:
                        def emit(self, m): pass
                    self.pgdas_payments_map = _load_and_process_pgdas(pgdas_path, DummyEmitter())
                except: pass

            self.assignment_page.refresh_autos_list()
            self.assignment_page.refresh_all_tables()
            self.fine_page.update_fine_text()
            
            # ✅ Force recalculation next time Preview is opened so it uses the loaded values
            self.mark_dirty() 

            if not silent:
                QMessageBox.information(self, "Sucesso", f"Sessão restaurada. {restored_count} notas re-associadas aos autos.")
            return True

        except Exception as e:
            if not silent: QMessageBox.warning(self, "Erro", f"Falha ao carregar sessão: {e}")
            return False

    def accept(self):
        if hasattr(self, 'preview_page'):
            self.preview_page._read_tables_into_context()
            
            for auto_key, (_, edit) in self.preview_page.manual_credit_widgets.items():
                val_text = edit.text().strip()
                if not val_text:
                    QMessageBox.warning(
                        self, 
                        "Campo Obrigatório", 
                        f"O 'Valor Atualizado' para o {auto_key} é obrigatório e não pode ficar vazio.\n"
                        "Por favor, preencha-o na aba '3. Pré-visualização'."
                    )
                    self.tab_widget.setCurrentWidget(self.preview_page)
                    edit.setFocus()
                    return

        self.preview_context = self.preview_page.get_final_context()
        self.fine_text_final = self.fine_page.fine_text_edit.toPlainText()
        self.preview_context['texto_multa'] = self.fine_text_final
        # ✅ CHANGED: Read from AssignmentPage
        self.preview_context['epaf_numero'] = self.assignment_page.epaf_numero_edit.text()
        
        self.save_session()
        super().accept()

    def get_final_data_for_confirmation(self):
        """
        Gets the data format needed for calculations and final context building,
        including flags for split autos.
        """
        final_data = {}
        for auto_id, auto_data in self.autos.items():
            df = auto_data.get('df') 

            motive_text = auto_data['motive']
            base_rule = auto_data.get('rule_name', self.motive_to_rule_map.get(motive_text.split(' (')[0], motive_text))

            if isinstance(df, pd.DataFrame):
                invoice_list = df.index.tolist()
                df_is_empty = df.empty
                first_row_data = df.iloc[0] if not df_is_empty else {}
            else:
                invoice_list = []
                df_is_empty = True
                first_row_data = {}

            correct_aliquota_val = auto_data.get('user_defined_aliquota')
            
            if correct_aliquota_val is None and not df_is_empty:
                if 'correct_rate' in df.columns and pd.notna(first_row_data.get('correct_rate')):
                    correct_aliquota_val = first_row_data.get('correct_rate')

            correct_aliquota_str = f"{correct_aliquota_val:.2f}" if correct_aliquota_val is not None else "5.00"

            data_entry = {
                'invoices': invoice_list, 
                'auto_id': auto_id,
                'rule_name': base_rule,
                'motive_text': motive_text,
                'correct_aliquota': correct_aliquota_str, 
                'user_defined_credito': auto_data.get('user_defined_credito'),
                'is_split_diff': auto_data.get('is_split_diff', False),
                'auto_text': auto_data.get('auto_text', ''),
                'monthly_overrides': auto_data.get('monthly_overrides', {})
            }
            if auto_data.get('user_defined_aliquota') is not None:
                 data_entry['user_defined_aliquota'] = auto_data.get('user_defined_aliquota')

            final_data[auto_id] = data_entry
        return final_data

    def calculate_preview_context(self):
        # 1. Start Standard Calculation
        # We fetch data from self.autos via get_final_data_for_confirmation.
        # Since self.autos already contains the up-to-date 'user_defined_credito' (saved when leaving Tab 3),
        # we don't need to complex "preserved_overrides" logic for that field anymore.
        final_data = self.get_final_data_for_confirmation()
        company_invoices_df = self.all_invoices_df
        dam_map = self.dam_payments_map
        pgdas_map = self.pgdas_payments_map 

        all_valid_dates = []
        for auto_info in final_data.values():
            invoice_indices = auto_info.get('invoices', [])
            if not invoice_indices: continue
            df_inv = company_invoices_df.loc[invoice_indices].copy()
            df_inv['DATA EMISSÃO'] = pd.to_datetime(df_inv['DATA EMISSÃO'], errors='coerce')
            all_valid_dates.extend(df_inv['DATA EMISSÃO'].dropna().tolist())
        for date_str in pgdas_map.keys():
            try: all_valid_dates.append(pd.to_datetime(date_str, format='%m/%Y'))
            except (ValueError, TypeError): pass
        
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

        self.available_credits_map = {
            'DAM': copy.deepcopy(dam_map),
            'PGDAS': {k: v[0] for k, v in pgdas_map.items()}
        }

        autos_context = []
        for auto_key, auto_info in final_data.items():
            target_year = self._get_auto_year(auto_key)
            monthly_override_map = auto_info.get('monthly_overrides', {})
            invoice_indices = auto_info.get('invoices', [])
            df_invoices = company_invoices_df.loc[invoice_indices].copy()
            df_invoices['DATA EMISSÃO'] = pd.to_datetime(df_invoices['DATA EMISSÃO'], errors='coerce')
            df_invoices.dropna(subset=['DATA EMISSÃO'], inplace=True)
            df_invoices['_month_str'] = df_invoices['DATA EMISSÃO'].dt.strftime('%m/%Y')

            try:
                aliquota_str = auto_info.get('user_defined_aliquota', auto_info.get('correct_aliquota', '0.0'))
                default_aliquota_pct = float(aliquota_str)
            except (ValueError, TypeError):
                default_aliquota_pct = 0.0

            is_idd_auto = auto_info.get('rule_name') == 'idd_nao_pago'

            if not df_invoices.empty:
                for idx, row in df_invoices.iterrows():
                    period_str = row['_month_str']
                    if period_str in monthly_override_map:
                        target_rate = monthly_override_map[period_str]
                    elif is_idd_auto:
                        target_rate = row.get('ALÍQUOTA', 0)
                    elif 'correct_rate' in df_invoices.columns and pd.notna(row['correct_rate']) and row['correct_rate'] > 0:
                        target_rate = row['correct_rate']
                    elif row.get('ALÍQUOTA', 0) > 0:
                            target_rate = row.get('ALÍQUOTA')
                    else:
                        target_rate = default_aliquota_pct
                    df_invoices.at[idx, '_target_rate_group'] = float(target_rate)
            else:
                df_invoices['_target_rate_group'] = default_aliquota_pct

            dados_anuais = []
            dam_pago_auto_total = 0; das_pago_auto_total = 0; total_iss_final_calc = 0
            total_base_auto = 0.0; total_iss_bruto_auto = 0.0
            total_iss_pago_auto = 0.0; total_iss_liquido_auto = 0.0

            for period_key in all_periods_list:
                if target_year is not None:
                    if period_key.year != target_year:
                        continue 

                period_str_mm_yyyy = period_key.strftime('%m/%Y')
                period_str_m_yyyy = f"{period_key.month}/{period_key.year}"

                month_mask = (df_invoices['DATA EMISSÃO'].dt.to_period('M') == period_key)
                df_month = df_invoices[month_mask]

                if df_month.empty:
                    unique_rates = [default_aliquota_pct]
                else:
                    unique_rates = df_month['_target_rate_group'].unique()
                    unique_rates.sort()

                for rate_val in unique_rates:
                    if df_month.empty:
                        group = None
                        current_target_rate = monthly_override_map.get(period_str_mm_yyyy, default_aliquota_pct)
                    else:
                        group = df_month[df_month['_target_rate_group'] == rate_val]
                        current_target_rate = rate_val

                    dams_list = self.available_credits_map['DAM'].get(period_str_m_yyyy, [])
                    available_dam_total = sum(d['val'] for d in dams_list)
                    
                    available_pgdas = self.available_credits_map['PGDAS'].get(period_str_mm_yyyy, 0.0)
                    pgdas_decl_num = pgdas_map.get(period_str_mm_yyyy, (0.0, "-"))[1]

                    base_calculo = 0.0; iss_correto_bruto = 0.0
                    iss_declarado_pago = 0.0; iss_liquido_calc = 0.0
                    monthly_invoice_data = []

                    if group is not None and not group.empty:
                        val_col = group['VALOR'].fillna(0.0)
                        ded_col = group.get('VALOR DEDUÇÃO', pd.Series([0]*len(group), index=group.index)).fillna(0.0)
                        base_calculo = (val_col - ded_col).sum()
                        rate_dec = current_target_rate / 100.0

                        for _, row in group.iterrows():
                            raw_val = row.get('VALOR', 0.0)
                            deducao = row.get('VALOR DEDUÇÃO', 0.0)
                            if pd.isna(deducao): deducao = 0.0
                            val = raw_val - deducao
                            decl_rate = row.get('ALÍQUOTA', 0.0) / 100.0
                            paid_status = str(row.get('PAGAMENTO', 'Não')).strip().lower()
                            is_paid = paid_status in ['sim', 'idd']

                            monthly_invoice_data.append({
                                'valor': val, 'declared_rate': decl_rate, 'is_paid': is_paid
                            })

                            iss_correto_bruto += (val * rate_dec)
                            paid_amt = (val * decl_rate) if is_paid else 0.0
                            iss_declarado_pago += paid_amt
                            if is_paid:
                                iss_liquido_calc += max(0, (rate_dec - decl_rate) * val)
                            else:
                                iss_liquido_calc += (rate_dec * val)
                    
                    if base_calculo == 0.0 and len(unique_rates) > 1:
                        continue 

                    aliquota_op_display = f'{current_target_rate:.2f}%'
                    if base_calculo > 0.001:
                        aliquota_declarada_display = f'{(iss_declarado_pago / base_calculo) * 100.0:.2f}%'
                        effective_aliquota_pct = (iss_liquido_calc / base_calculo) * 100.0
                        aliquota_display = f'{effective_aliquota_pct:.2f}%'
                    else:
                        aliquota_declarada_display = "-"; aliquota_display = "-"

                    dam_utilizado = min(iss_liquido_calc, available_dam_total)
                    used_dam_codes = []
                    remainder_to_deduct = dam_utilizado
                    for dam_obj in dams_list:
                        if remainder_to_deduct <= 0.0001: break
                        if dam_obj['val'] > 0:
                            deduct = min(dam_obj['val'], remainder_to_deduct)
                            dam_obj['val'] -= deduct 
                            remainder_to_deduct -= deduct
                            used_dam_codes.append(dam_obj['code'])
                    
                    dam_ident_str = ", ".join(sorted(set(used_dam_codes))) if used_dam_codes else "-"
                    iss_restante = iss_liquido_calc - dam_utilizado
                    pgdas_utilizado = min(iss_restante, available_pgdas)
                    iss_apurado_op = max(0, iss_liquido_calc - dam_utilizado - pgdas_utilizado)

                    self.available_credits_map['PGDAS'][period_str_mm_yyyy] = available_pgdas - pgdas_utilizado
                    dam_pago_auto_total += dam_utilizado
                    das_pago_auto_total += pgdas_utilizado
                    total_iss_final_calc += iss_apurado_op
                    total_base_auto += base_calculo
                    total_iss_bruto_auto += iss_correto_bruto
                    total_iss_pago_auto += iss_declarado_pago
                    total_iss_liquido_auto += iss_liquido_calc

                    dados_anuais.append({
                        'mes_ano': period_str_mm_yyyy,
                        'base_calculo': base_calculo,
                        'aliquota_display': aliquota_display,
                        'aliquota_target_user': current_target_rate,
                        'aliquota_op': aliquota_op_display,
                        'iss_apurado_bruto': iss_correto_bruto,
                        'aliquota_declarada': aliquota_declarada_display,
                        'iss_declarado_pago': iss_declarado_pago,
                        'iss_apurado_liquido': iss_liquido_calc,
                        'iss_apurado': iss_liquido_calc, 
                        'base_calculo_op': base_calculo,
                        'iss_apurado_op': iss_apurado_op,
                        'dam_iss_pago': dam_utilizado, 'dam_identificacao': dam_ident_str,
                        'das_iss_pago': pgdas_utilizado, 'das_identificacao': pgdas_decl_num if pgdas_utilizado > 0 else "-",
                        '_monthly_invoices_data': monthly_invoice_data, 
                    })

            if total_base_auto > 0.001:
                total_effective_aliquota_pct = (total_iss_liquido_auto / total_base_auto) * 100.0
                total_aliquota_display = f'{total_effective_aliquota_pct:.2f}%'
            else:
                total_aliquota_display = "-"

            # ✅ KEY CHANGE: Use 'user_defined_credito' directly from the source (final_data/self.autos)
            # This ensures that even if you rename the auto (change key/id), the value persists
            # because 'edit_selected_auto' moves this value in self.autos.
            final_user_credit = auto_info.get('user_defined_credito')

            auto_data = {
                'numero': auto_key,
                'motive_text': auto_info.get('motive_text', 'N/A'),
                'nfs_e_numeros': "...", 
                'periodo': "...", 
                'motivo': { 'tipo': auto_info.get('rule_name', 'desconhecido') },
                'auto_text': auto_info.get('auto_text', ''),
                'is_split_diff': auto_info.get('is_split_diff', False),
                'tem_pagamento_das': das_pago_auto_total > 0,
                'tem_pagamento_dam': dam_pago_auto_total > 0,
                'totais': {
                    'base_calculo': total_base_auto,
                    'iss_apurado_bruto': total_iss_bruto_auto,
                    'iss_declarado_pago': total_iss_pago_auto,
                    'iss_apurado_liquido': total_iss_liquido_auto,
                    'iss_apurado': total_iss_liquido_auto,
                    'base_calculo_op': total_base_auto,
                    'iss_apurado_op': total_iss_final_calc,
                    'das_iss_pago': das_pago_auto_total,
                    'dam_iss_pago': dam_pago_auto_total,
                    '_total_aliquota_display': total_aliquota_display 
                },
                'dados_anuais': dados_anuais,
                'calculated_iss_original': total_iss_liquido_auto,
                'user_defined_credito': final_user_credit, 
                'monthly_overrides': monthly_override_map 
            }
            autos_context.append(auto_data)

        # --- Build Summary ---
        summary_autos_list = []
        total_credito_autos = 0.0
        nfs_numbers_map = {}
        for auto_key, auto_info in final_data.items():
            invoice_indices = auto_info.get('invoices', [])
            if invoice_indices:
                invoice_numbers = company_invoices_df.loc[invoice_indices, Columns.INVOICE_NUMBER].astype(str).tolist()
                nfs_numbers_map[auto_key] = invoice_numbers
            else:
                nfs_numbers_map[auto_key] = []

        for auto in autos_context:
            auto_total_credito = auto.get('totais', {}).get('iss_apurado_op', 0.0)
            user_credit = auto.get('user_defined_credito')
            if user_credit is not None:
                auto_total_credito = user_credit

            if auto_total_credito > 0.01: 
                auto_key = auto['numero']
                invoice_list = format_invoice_numbers(nfs_numbers_map.get(auto_key, []))
                summary_autos_list.append({
                    'numero': auto['numero'],
                    'nfs_tributadas': invoice_list,
                    'iss_valor_original': auto.get('totais', {}).get('iss_apurado_bruto', 0.0),
                    'total_credito_tributario': auto_total_credito,
                    'motivo': auto.get('motive_text', 'N/A')
                })
                total_credito_autos += auto_total_credito

        # Fines
        fines_list = self.fine_page.get_fines_data()
        total_multas_val = 0.0
        processed_fines = []
        for f in fines_list:
            try:
                val = float(str(f['value']).replace("R$", "").strip().replace(".", "").replace(",", "."))
                if val > 0.01:
                    total_multas_val += val
                    processed_fines.append({
                        'numero': f['number'],
                        'valor_credito': val,
                        'ano': f['year']
                    })
            except: pass

        total_geral = total_credito_autos + total_multas_val
        
        summary_data = {
            'autos': summary_autos_list,
            'multas': processed_fines,
            'multa': processed_fines[0] if processed_fines else None,
            'total_geral_credito': total_geral
        }
        
        self.preview_context = {
            'autos': autos_context,
            'available_credits_startup': {
                'DAM': copy.deepcopy(dam_map), 
                'PGDAS': {k: v[0] for k, v in pgdas_map.items()}
            },
            'summary': summary_data  
        }
        return self.preview_context
    

class AssignmentPage(QWidget): 
    def __init__(self, wizard):
        super().__init__()
        self.wizard = wizard
        self.current_available_df = pd.DataFrame() 
        
        self.assigned_filters = []
        self.available_filters = []

        main_page_layout = QVBoxLayout(self)

        # 2. Subtitle Label
        self.subtitle_label = QLabel()
        self.subtitle_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.subtitle_label.setStyleSheet("background-color: white; color: black; font-style: italic; padding: 5px; border: 1px solid #ddd;")
        main_page_layout.addWidget(self.subtitle_label)

        # ✅ NEW: Add EPAF Input here
        epaf_layout = QHBoxLayout()
        epaf_layout.addWidget(QLabel("Nº ePAF:"))
        self.epaf_numero_edit = QLineEdit()
        self.epaf_numero_edit.setPlaceholderText("Ex: 2025/1234567-8")
        self.epaf_numero_edit.setText(self.wizard.company_epaf_initial)
        epaf_layout.addWidget(self.epaf_numero_edit)
        epaf_layout.addStretch()
        main_page_layout.addLayout(epaf_layout)

        # 3. Create the original layout
        tab_layout = QHBoxLayout()
        
        # --- Autos List ---
        autos_group = QGroupBox("Autos de Infração")
        autos_layout = QVBoxLayout()
        self.autos_list_widget = QListWidget()
        self.autos_list_widget.currentItemChanged.connect(self.on_auto_selection_change)
        self.autos_list_widget.itemDoubleClicked.connect(self.edit_selected_auto)

        autos_buttons_layout = QHBoxLayout()
        new_auto_btn = QPushButton("Criar Novo Auto")
        new_auto_btn.clicked.connect(self.create_new_auto)
        edit_auto_btn = QPushButton("Editar Auto")
        edit_auto_btn.clicked.connect(self.edit_selected_auto)
        remove_auto_btn = QPushButton("Remover Auto")
        remove_auto_btn.clicked.connect(self.remove_selected_auto)
        autos_buttons_layout.addWidget(new_auto_btn)
        autos_buttons_layout.addWidget(edit_auto_btn)
        autos_buttons_layout.addWidget(remove_auto_btn)

        self.split_auto_button = QPushButton("Dividir Auto (Alíq. Incorreta)")
        self.split_auto_button.setToolTip(
            "Disponível apenas para autos de 'Alíquota Incorreta' Pura.\n"
            "Cria um auto 'IDD (Alíq. Declarada)' e um 'Diferença Alíquota'."
        )
        self.split_auto_button.setEnabled(False)
        self.split_auto_button.clicked.connect(self.split_aliquota_auto)
        
        self.edit_auto_text_btn = QPushButton("Editar Texto do Auto")
        self.edit_auto_text_btn.setToolTip(
            "Abrir o editor de texto detalhado para este auto de infração."
        )
        self.edit_auto_text_btn.clicked.connect(self.edit_auto_text)
        
        extra_buttons_layout = QHBoxLayout()
        extra_buttons_layout.addWidget(self.split_auto_button)
        extra_buttons_layout.addWidget(self.edit_auto_text_btn)

        export_btn_layout = QHBoxLayout()
        self.export_excel_btn = QPushButton("Exportar Autos (Excel)...")
        self.export_excel_btn.clicked.connect(self.export_autos_to_excel)
        export_btn_layout.addWidget(self.export_excel_btn)

        save_session_btn = QPushButton("Salvar Sessão")
        save_session_btn.clicked.connect(self.wizard.handle_manual_save)
        export_btn_layout.addWidget(save_session_btn)

        autos_layout.addWidget(self.autos_list_widget)
        autos_layout.addLayout(autos_buttons_layout)
        autos_layout.addLayout(extra_buttons_layout)
        autos_layout.addLayout(export_btn_layout)
        autos_group.setLayout(autos_layout)

        # --- Move Buttons ---
        move_buttons_layout = QVBoxLayout()
        move_buttons_layout.addStretch()
        add_btn = QPushButton(">>\nAdicionar")
        add_btn.clicked.connect(self.move_to_assigned)

        correct_btn = QPushButton("Corrigir\nInfrações")
        correct_btn.clicked.connect(self.correct_infractions)
        correct_btn.setToolTip("Selecione notas e clique para remover alertas de infração específicos.")

        remove_btn = QPushButton("<<\nRemover")
        remove_btn.clicked.connect(self.move_to_available)

        move_buttons_layout.addWidget(add_btn)
        move_buttons_layout.addWidget(correct_btn)
        move_buttons_layout.addWidget(remove_btn)
        move_buttons_layout.addStretch()

        # --- Tables ---
        tables_widget = QWidget()
        tables_layout = QVBoxLayout(tables_widget)
        
        configure_cols_btn = QPushButton("Configurar Colunas Visíveis")
        configure_cols_btn.clicked.connect(self.open_column_configuration)
        tables_layout.addWidget(configure_cols_btn, 0, Qt.AlignmentFlag.AlignRight)

        assigned_group = QGroupBox("Notas Atribuídas ao Auto Selecionado")
        assigned_layout = QVBoxLayout()

        assigned_filter_controls_layout = QHBoxLayout()
        self.add_assigned_filter_btn = QPushButton("Adicionar Filtro (Atribuídas)")
        assigned_filter_controls_layout.addWidget(self.add_assigned_filter_btn)
        assigned_filter_controls_layout.addStretch()
        
        self.assigned_filters_widget = QWidget()
        self.assigned_filters_layout = QVBoxLayout(self.assigned_filters_widget)
        self.assigned_filters_layout.setContentsMargins(0, 0, 0, 0)
        
        assigned_layout.addLayout(assigned_filter_controls_layout)
        assigned_layout.addWidget(self.assigned_filters_widget)

        self.assigned_invoices_table = self.create_invoice_table()
        assigned_layout.addWidget(self.assigned_invoices_table)

        assigned_bottom_layout = QHBoxLayout()
        self.view_assigned_details_btn = QPushButton("Ver Detalhes")
        self.view_assigned_details_btn.clicked.connect(self.view_assigned_details)
        assigned_bottom_layout.addWidget(self.view_assigned_details_btn)
        assigned_bottom_layout.addStretch()
        self.assigned_summary_label = QLabel("Notas: 0 | Valor Total: R$ 0,00")
        self.assigned_summary_label.setStyleSheet("font-weight: bold;")
        assigned_bottom_layout.addWidget(self.assigned_summary_label)
        assigned_layout.addLayout(assigned_bottom_layout)

        assigned_group.setLayout(assigned_layout)

        available_group = QGroupBox("Notas Disponíveis")
        available_layout = QVBoxLayout()
        
        available_filter_controls_layout = QHBoxLayout()
        self.add_avail_filter_btn = QPushButton("Adicionar Filtro (Disponíveis)")
        available_filter_controls_layout.addWidget(self.add_avail_filter_btn)
        available_filter_controls_layout.addStretch()
        
        self.available_filters_widget = QWidget()
        self.available_filters_layout = QVBoxLayout(self.available_filters_widget)
        self.available_filters_layout.setContentsMargins(0, 0, 0, 0)

        available_layout.addLayout(available_filter_controls_layout)
        available_layout.addWidget(self.available_filters_widget)
        
        self.available_invoices_table = self.create_invoice_table()
        available_layout.addWidget(self.available_invoices_table)

        available_bottom_layout = QHBoxLayout()
        self.view_available_details_btn = QPushButton("Ver Detalhes")
        self.view_available_details_btn.clicked.connect(self.view_available_details)
        available_bottom_layout.addWidget(self.view_available_details_btn)
        available_bottom_layout.addStretch()
        self.available_summary_label = QLabel("Notas: 0 | Valor Total: R$ 0,00")
        self.available_summary_label.setStyleSheet("font-weight: bold;")
        available_bottom_layout.addWidget(self.available_summary_label)
        available_layout.addLayout(available_bottom_layout)

        available_group.setLayout(available_layout)

        tables_layout.addWidget(assigned_group, 1)
        tables_layout.addWidget(available_group, 1)

        tab_layout.addWidget(autos_group, 3) 
        tab_layout.addLayout(move_buttons_layout, 1) 
        tab_layout.addWidget(tables_widget, 6)
        main_page_layout.addLayout(tab_layout)
        
        self.add_assigned_filter_btn.clicked.connect(self.add_assigned_filter_row)
        self.add_avail_filter_btn.clicked.connect(self.add_available_filter_row)

        self.assigned_invoices_table.doubleClicked.connect(self.show_full_cell_content)
        self.available_invoices_table.doubleClicked.connect(self.show_full_cell_content)

        self.ignore_btn = QPushButton("🗑️ Ignorar\n(Fora/Desc.)")
        self.ignore_btn.setToolTip("Marca as notas selecionadas como 'Ignoradas' (ex: Fora do Município) e remove da lista de disponíveis.")
        self.ignore_btn.setStyleSheet("background-color: #BF616A; color: white; font-weight: bold;")
        self.ignore_btn.clicked.connect(self.mark_as_ignored)
        
        move_buttons_layout.addWidget(self.ignore_btn)
        
        self.show_ignored_chk = QCheckBox("Mostrar Ignoradas")
        self.show_ignored_chk.stateChanged.connect(self.populate_available_table)
        
        available_filter_controls_layout.addWidget(self.show_ignored_chk)
    
    def refresh_autos_list(self):
        self.autos_list_widget.clear()
        
        # ✅ CHANGE: Sort items by key (which includes the Year suffix now)
        sorted_items = sorted(self.wizard.autos.items(), key=lambda x: x[0]) 
        
        for auto_id, data in sorted_items:
            full_text = f"{auto_id}: {data['motive']}" 
            item = QListWidgetItem(full_text)
            item.setData(Qt.ItemDataRole.UserRole, auto_id)
            item.setToolTip(full_text)
            self.autos_list_widget.addItem(item)
            
        if self.autos_list_widget.count() > 0:
            self.autos_list_widget.setCurrentRow(0)

    def show_full_cell_content(self, index): # Argument is now QModelIndex
        if not index.isValid(): return
        
        # Get column name from the model's visible columns list
        col_idx = index.column()
        try:
            column_name = self.wizard.visible_columns[col_idx]
        except IndexError: return

        if column_name != Columns.SERVICE_DESCRIPTION:
            return

        # Get DF index from UserRole
        invoice_index = index.data(Qt.UserRole)
        if invoice_index is None:
            return

        # 4. Busca o texto completo e bruto direto do DataFrame
        try:
            full_text = self.wizard.all_invoices_df.loc[invoice_index, Columns.SERVICE_DESCRIPTION]
            if pd.isna(full_text):
                full_text = "(Vazio)"
        except (KeyError, AttributeError):
            return # Deve ser um índice válido

        # 5. Cria um diálogo simples para exibir o texto
        dialog = QDialog(self)
        dialog.setWindowTitle("Discriminação do Serviço")
        dialog.setMinimumSize(700, 500) # Tamanho bom para leitura
        
        layout = QVBoxLayout(dialog)
        
        text_edit = QTextEdit()
        text_edit.setPlainText(str(full_text)) # Mostra o texto exatamente como está
        text_edit.setReadOnly(True)
        
        # Adiciona um checkbox para quebra de linha (para textos mal formatados)
        wrap_check = QCheckBox("Quebra automática de linha")
        wrap_check.setChecked(True) # Começa ativado
        
        # Ação do checkbox: altera o modo de quebra de linha do QTextEdit
        wrap_check.toggled.connect(lambda checked: text_edit.setLineWrapMode(
            QTextEdit.LineWrapMode.WidgetWidth if checked else QTextEdit.LineWrapMode.NoWrap
        ))
        
        close_button = QPushButton("Fechar")
        close_button.clicked.connect(dialog.accept) # .accept() fecha o diálogo
        
        layout.addWidget(wrap_check)
        layout.addWidget(text_edit)
        layout.addWidget(close_button)
        
        dialog.exec() # Mostra o diálogo (modal)
    
    def isComplete(self):
        if not self.wizard.autos:
            return False
        return any(
            isinstance(data.get('df'), pd.DataFrame) and not data['df'].empty 
            for data in self.wizard.autos.values()
        )
    
    def edit_auto_text(self):
        current_item = self.autos_list_widget.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Ação Inválida", "Nenhum auto selecionado.")
            return

        auto_id = current_item.data(Qt.ItemDataRole.UserRole)
        try:
            auto_data_original = self.wizard.autos.get(auto_id)
            if not auto_data_original:
                raise KeyError(f"Auto {auto_id} not found in wizard.autos")
            # Make a copy to pass to the dialog
            auto_data_for_dialog = auto_data_original.copy()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível encontrar dados para o auto {auto_id}: {e}")
            return

        df_invoices = auto_data_original.get('df')
        
        if not isinstance(df_invoices, pd.DataFrame) or df_invoices.empty:
            QMessageBox.warning(self, "Texto Não Disponível", 
                                "Atribua notas fiscais a este auto antes de gerar o texto.")
            return

        # --- ✅ START: Aliquot Override Logic for Dialog ---
        try:
            # 1. Get the monthly overrides from the original data
            monthly_overrides = auto_data_original.get('monthly_overrides', {})
            
            # 2. Get the default "correct" aliquot
            default_aliquota = auto_data_original.get('user_defined_aliquota')
            
            final_aliquota_to_use = default_aliquota
            final_aliquota_str = f"{default_aliquota:.2f}" if default_aliquota is not None else "[N/A]"
            
            if monthly_overrides:
                # 3. Find the most common (mode) overridden value
                override_values = [v for v in monthly_overrides.values() if v > 0]
                if override_values:
                    try:
                        final_aliquota_to_use = mode(override_values)
                    except Exception: # Handle case where there is no unique mode
                        final_aliquota_to_use = override_values[0] # Just pick the first one
                    
                    final_aliquota_str = f"{final_aliquota_to_use:.2f}"

            # 4. Set this "best" aliquot in the *copy* of the data
            if final_aliquota_to_use is not None:
                auto_data_for_dialog['user_defined_aliquota'] = final_aliquota_to_use

            # 5. ✅ Build the 'motivo' dict *now* for the dialog
            # This ensures the dialog has the same data structure as the final report
            auto_data_for_dialog['motivo'] = {
                'tipo': auto_data_original.get('rule_name', 'DEFAULT_AUTO_FALLBACK'),
                'texto_simples': auto_data_original.get('motive', 'Motivo não especificado.'),
                'aliquota_correta': final_aliquota_str 
            }
            
            # 6. Run calculation to get the calculated context
            self.wizard.calculate_preview_context()
            calculated_auto_data = None
            for auto in self.wizard.preview_context.get('autos', []):
                if auto['numero'] == auto_id:
                    calculated_auto_data = auto
                    break
            
            if not calculated_auto_data:
                print(f"Warning: No calculated data found for {auto_id} in preview context.")
                calculated_auto_data = {}
                
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Calcular", f"Erro ao preparar dados para o editor de texto: {e}")
            print(traceback.format_exc())
            return
        # --- ✅ END: Aliquot Override Logic for Dialog ---

        # Pass the *copy* (auto_data_for_dialog) to the dialog
        dialog = AutoTextDialog(
            auto_data_for_dialog, # Pass the copy with the corrected aliquot AND motivo dict
            df_invoices, 
            self.wizard.pgdas_payments_map, 
            calculated_auto_data,
            self
        )
        
        if dialog.exec():
            # Save the text back into the *original* data store
            self.wizard.autos[auto_id]['auto_text'] = dialog.get_text()
            print(f"Texto salvo para {auto_id}")

    def split_aliquota_auto(self):
        current_item = self.autos_list_widget.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Ação Inválida", "Nenhum auto selecionado.")
            return

        original_auto_id = current_item.data(Qt.ItemDataRole.UserRole)
        original_auto_data = self.wizard.autos.get(original_auto_id)

        if not original_auto_data or not original_auto_data['motive'].startswith('Alíquota Incorreta'):
            QMessageBox.warning(self, "Ação Inválida", "Esta ação só é válida para autos de 'Alíquota Incorreta'.")
            return

        df_invoices = original_auto_data.get('df')
        
        # ✅ --- START FIX: Check if df_invoices is a DataFrame before checking empty ---
        if not isinstance(df_invoices, pd.DataFrame) or df_invoices.empty:
        # ✅ --- END FIX ---
            QMessageBox.warning(self, "Ação Inválida", "O auto selecionado não contém notas fiscais para dividir.")
            return

        # **Crucial Check:** Ensure NO other infractions exist on ANY invoice in this auto
        for index, row in df_invoices.iterrows():
            details = self.wizard.all_invoices_df.loc[index, Columns.BROKEN_RULE_DETAILS]
            if isinstance(details, list):
                # Check if any detail DOES NOT start with 'Alíquota Incorreta'
                if any(not detail.startswith('Alíquota Incorreta') for detail in details):
                    QMessageBox.warning(self, "Divisão Não Permitida",
                                        f"A Nota Fiscal nº {row.get(Columns.INVOICE_NUMBER, index)} "
                                        f"contém outras infrações além da alíquota incorreta.\n"
                                        f"A divisão só é permitida para autos 'puros' de alíquota incorreta.")
                    return
            elif details: # Should be a list, but handle unexpected data
                 QMessageBox.warning(self, "Erro de Dados", f"Dados de infração inesperados para a nota {index}.")
                 return


        # --- Confirmation ---
        reply = QMessageBox.question(self, "Confirmar Divisão",
                                     f"Tem certeza que deseja dividir o auto '{original_auto_id}'?\n\n"
                                     "Isto irá REMOVER o auto original e criar DOIS novos autos:\n"
                                     f"1. IDD (Alíq. Declarada): Para o valor devido pela alíquota declarada.\n"
                                     f"2. Diferença Alíquota: Para o valor da diferença entre a alíquota correta e a declarada.\n\n"
                                     "As mesmas notas fiscais serão usadas em ambos os novos autos.",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.No:
            return

        # --- Perform Split ---
        try:
            # Extract necessary rates
            first_invoice_row = df_invoices.iloc[0]
            declared_rate = first_invoice_row.get(Columns.RATE, 0.0) # Get declared rate from data
            correct_rate = original_auto_data.get('user_defined_aliquota') # Get correct rate stored in auto
            if correct_rate is None: # Fallback if not user-defined
                correct_rate = first_invoice_row.get(Columns.CORRECT_RATE, 5.0) # Get from data or default

            # --- Create New IDD Auto ---
            new_idd_id = f"AUTO-{self.wizard.auto_counter:03d}"
            self.wizard.auto_counter += 1
            new_idd_data = {
                'motive': 'IDD (Alíq. Declarada)',
                'df': df_invoices.copy(), # Use the same invoices
                'rule_name': 'idd_nao_pago', # Use standard IDD rule name
                # Store the DECLARED rate as the 'correct' rate for THIS auto's calculation context
                'user_defined_aliquota': declared_rate,
                'auto_text': ''
            }
            self.wizard.autos[new_idd_id] = new_idd_data

            # --- Create New Difference Auto ---
            new_diff_id = f"AUTO-{self.wizard.auto_counter:03d}"
            self.wizard.auto_counter += 1
            new_diff_data = {
                'motive': 'Diferença Alíquota',
                'df': df_invoices.copy(), # Use the same invoices
                'rule_name': 'diferenca_aliquota', # Define a new internal rule name
                # Store the CORRECT rate for THIS auto's calculation context
                'user_defined_aliquota': correct_rate,
                # Flag for calculation logic adjustment
                'is_split_diff': True,
                'auto_text': ''
            }
            self.wizard.autos[new_diff_id] = new_diff_data

            # --- Remove Original Auto ---
            del self.wizard.autos[original_auto_id]

            # --- Refresh UI ---
            self.refresh_autos_list()
            new_diff_item = self.autos_list_widget.findItems(new_diff_id, Qt.MatchFlag.MatchStartsWith)
            if new_diff_item:
                self.autos_list_widget.setCurrentItem(new_diff_item[0])
            self.refresh_all_tables() # Triggers recalculation

            QMessageBox.information(self, "Divisão Concluída",
                                    f"Auto '{original_auto_id}' removido.\n"
                                    f"Novos autos criados: '{new_idd_id}' e '{new_diff_id}'.")

        except Exception as e:
            QMessageBox.critical(self, "Erro na Divisão", f"Ocorreu um erro inesperado:\n{e}")
            print(traceback.format_exc()) # Log the error

    # ✅ --- START: New Filter/Column Helper Methods ---
    
    def _add_filter_row(self, target_layout, target_filter_list, on_change_slot, source_df):
        """
        Helper to add a new filter row with Contains/Not Contains logic.
        """
        filter_row_layout = QHBoxLayout()
        
        # 1. Column Selection
        column_combo = QComboBox()
        column_combo.addItems(self.wizard.visible_columns) 
        
        # 2. ✅ NEW: Logic Selection (Contém / Não Contém)
        logic_combo = QComboBox()
        logic_combo.addItems(["Contém", "Não Contém"])
        logic_combo.setFixedWidth(110) # Keep it compact
        
        # 3. Text Input
        filter_edit = QLineEdit()
        filter_edit.setPlaceholderText("Texto...")
        
        
        
        # 5. Remove Button
        remove_btn = QPushButton("X")
        remove_btn.setToolTip("Remover filtro")
        remove_btn.setFixedWidth(30)
        remove_btn.setStyleSheet("color: red; font-weight: bold;")
        
        # --- Add Widgets to Layout ---
        filter_row_layout.addWidget(column_combo, 1) # Weight 1
        filter_row_layout.addWidget(logic_combo, 0)  # Fixed width
        filter_row_layout.addWidget(filter_edit, 2)  # Weight 2 (Gets more space)
        filter_row_layout.addWidget(remove_btn)
        
        target_layout.addLayout(filter_row_layout)
        
        # Store widget references
        filter_widgets = { 
            "layout": filter_row_layout, 
            "combo": column_combo, 
            "logic": logic_combo, # ✅ Save reference to logic combo
            "edit": filter_edit, 
            "button": remove_btn,
            "source_df": source_df
        }
        target_filter_list.append(filter_widgets)
        
        # --- Connect Signals ---
        # Update completer when column changes
        
        # Update filter when logic changes (Contém -> Não Contém)
        logic_combo.currentIndexChanged.connect(on_change_slot) 
        # Update filter when text changes
        filter_edit.textChanged.connect(on_change_slot)
        
        remove_btn.clicked.connect(lambda: self._remove_filter_row(
            filter_widgets, target_layout, target_filter_list, on_change_slot
        ))
        
        # Initialize
        on_change_slot()
    

    def _remove_filter_row(self, filter_widgets, target_layout, target_filter_list, on_change_slot):
        """Generic helper to remove a filter UI row."""
        # Remove from list
        if filter_widgets in target_filter_list:
            target_filter_list.remove(filter_widgets)
            
        # Delete widgets from layout
        for i in reversed(range(filter_widgets["layout"].count())):    
            widget = filter_widgets["layout"].itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()
        # Delete layout
        filter_widgets["layout"].deleteLater()
        
        on_change_slot() # Re-apply filters

    def _apply_filters_to_df(self, df, filter_list):
        """Helper to apply filters to a DataFrame with negative logic."""
        if df is None or df.empty:
            return df
            
        filtered_df = df.copy()
        
        for f in filter_list:
            column_name = f["combo"].currentText()
            filter_text = f["edit"].text().strip().lower()
            logic_mode = f["logic"].currentText() # ✅ Get "Contém" or "Não Contém"
            
            if filter_text and column_name in filtered_df.columns:
                try:
                    # Create the boolean mask (True where match found)
                    mask = filtered_df[column_name].astype(str).str.lower().str.contains(filter_text, na=False)
                    
                    if logic_mode == "Contém":
                        filtered_df = filtered_df[mask]   # Keep matches
                    else:
                        filtered_df = filtered_df[~mask]  # Keep NON-matches (Invert)
                        
                except Exception as e:
                    print(f"Error filtering column {column_name}: {e}")
                    
        return filtered_df

    # --- Public Slots for Buttons ---
    def add_assigned_filter_row(self):
        """Slot to add a filter to the 'Assigned' table."""
        self._add_filter_row(
            self.assigned_filters_layout, 
            self.assigned_filters, 
            self.refresh_all_tables,
            # ✅ Passa uma *função* que obtém o DF atual
            lambda: self.wizard.autos.get(self.get_current_auto_id(), {}).get('df', pd.DataFrame())
        )

    def add_available_filter_row(self):
        """Slot to add a filter to the 'Available' table."""
        self._add_filter_row(
            self.available_filters_layout, 
            self.available_filters, 
            self.populate_available_table,
            # ✅ Passa uma *função* que obtém o DF principal
            lambda: self.wizard.all_invoices_df
        )
    # ✅ --- END: New Filter/Column Helper Methods ---

    def populate_available_table(self):
        all_assigned_indices = pd.Index([])
        
        # ✅ --- START FIX: Check if 'df' is a DataFrame before accessing .index ---
        for auto_data in self.wizard.autos.values():
            df = auto_data.get('df')
            if isinstance(df, pd.DataFrame): # Check if it's a DataFrame
                all_assigned_indices = all_assigned_indices.union(df.index)
        # ✅ --- END FIX ---

        available_df = self.wizard.all_invoices_df.drop(all_assigned_indices, errors='ignore')

        if not self.show_ignored_chk.isChecked():
            # Ensure column exists
            if 'status_manual' not in available_df.columns:
                available_df['status_manual'] = ''
            
            available_df = available_df[available_df['status_manual'] != 'Ignored']

        current_auto_id = self.get_current_auto_id()
        target_year = self.wizard._get_auto_year(current_auto_id) # Use the helper we just made
        
        # Ensure DATA EMISSÃO is proper format in available_df
        if not available_df.empty:
             if not pd.api.types.is_datetime64_any_dtype(available_df['DATA EMISSÃO']):
                 available_df['DATA EMISSÃO'] = pd.to_datetime(available_df['DATA EMISSÃO'], errors='coerce')

        if target_year is not None and not available_df.empty:
            # Filter: Show only invoices from 'target_year'
            available_df = available_df[available_df['DATA EMISSÃO'].dt.year == target_year]

        # ✅ --- START: Apply new dynamic filters ---
        available_df = self._apply_filters_to_df(available_df, self.available_filters)
        # ✅ --- END: Apply new dynamic filters ---

        self.current_available_df = available_df.copy()
        
        self.available_invoices_table.horizontalHeader().set_dataframe(available_df)
        self.available_invoices_table.column_mapping = self.wizard.visible_columns # Sync mapping

        self.populate_table_with_df(self.available_invoices_table, available_df)
        
        # ❌ REMOVED: The manual iteration loop that caused the crash.
        # The InvoiceTableModel (lines 80-92) now handles the gray background/text color 
        # automatically based on 'status_manual' == 'Ignored'.
        
        iss_declarado = 0.0
        if not available_df.empty:
            iss_declarado = ((available_df[Columns.VALUE] * available_df[Columns.RATE]) / 100.0).sum()
        self.available_summary_label.setText(format_summary_label(available_df, iss_value=iss_declarado, iss_label="ISS Declarado"))

    def populate_table_with_df(self, table, df):
        if df is None: 
            df = pd.DataFrame(columns=self.wizard.all_invoices_df.columns)
        
        # Create the efficient model
        model = InvoiceTableModel(df, self.wizard.visible_columns, parent=table)
        
        # Assign model to view
        table.setModel(model)
        
        # Refresh header connection
        if table.horizontalHeader():
            table.horizontalHeader().set_dataframe(df)

    def refresh_all_tables(self):
        current_auto_id = self.get_current_auto_id()
        assigned_df = pd.DataFrame(columns=self.wizard.all_invoices_df.columns)
        calculated_iss_original = 0.0 # Default value

        if current_auto_id and current_auto_id in self.wizard.autos:
            auto_data = self.wizard.autos[current_auto_id]
            
            # ✅ --- START FIX: Check if 'df' is a DataFrame before proceeding ---
            assigned_df = auto_data.get('df') # Use .get for safety

            # Check if it's actually a DataFrame before proceeding
            if not isinstance(assigned_df, pd.DataFrame):
                 assigned_df = pd.DataFrame(columns=self.wizard.all_invoices_df.columns)
            # ✅ --- END FIX ---
            
            if not assigned_df.empty:
                # --- START: Updated Aliquot/Tax Calculation Logic (Matches Preview Page) ---
                
                # 1. Get the saved monthly overrides (from Page 3)
                monthly_overrides = auto_data.get('monthly_overrides', {})

                # 2. Get the default "correct" aliquot for this auto (Fallback)
                first_row = assigned_df.iloc[0]
                default_aliquota_pct = 5.0 # default
                if auto_data.get('user_defined_aliquota') is not None:
                    default_aliquota_pct = auto_data['user_defined_aliquota']
                elif 'correct_rate' in assigned_df.columns and pd.notna(first_row.get('correct_rate')):
                    default_aliquota_pct = first_row.get('correct_rate')
                elif (auto_data['motive'].startswith('IDD (Não Pago)') or 
                      auto_data['motive'].startswith('Alíquota incorreta')) and pd.notna(first_row.get(Columns.RATE)):
                    default_aliquota_pct = first_row.get(Columns.RATE)
                

                is_idd_auto = auto_data.get('rule_name') == 'idd_nao_pago'

                # 3. Apply calculation logic row-by-row
                temp_df = assigned_df.copy() 
                temp_df['iss_original_calculado'] = 0.0
                
                if not pd.api.types.is_datetime64_any_dtype(temp_df['DATA EMISSÃO']):
                     temp_df['DATA EMISSÃO'] = pd.to_datetime(temp_df['DATA EMISSÃO'], errors='coerce')

                for idx, row in temp_df.iterrows():
                    period_str = row['DATA EMISSÃO'].strftime('%m/%Y') if pd.notna(row['DATA EMISSÃO']) else None
                    target_rate = default_aliquota_pct 

                    # Priority 1: Manual Override (from Page 3)
                    if period_str in monthly_overrides:
                        target_rate = monthly_overrides[period_str]
                    # ✅ CHANGE: Priority 2: Force IDD to Declared Rate (ALÍQUOTA)
                    elif is_idd_auto:
                        target_rate = row.get('ALÍQUOTA', 0)
                    # Priority 3: Specific 'correct_rate' on Invoice (Normal Flow)
                    elif 'correct_rate' in temp_df.columns and pd.notna(row['correct_rate']) and row['correct_rate'] > 0:
                        target_rate = row['correct_rate']
                    # Priority 4: Declared Rate (Fallback)
                    elif row.get('ALÍQUOTA', 0) > 0:
                         target_rate = row.get('ALÍQUOTA')
                    
                    aliquota_dec = target_rate / 100.0

                    # 5. Calculate ISS for this row using the final aliquot
                    declared_rate_decimal = row.get('ALÍQUOTA', 0.0) / 100.0
                    raw_valor = row.get('VALOR', 0.0)
                    deducao = row.get('VALOR DEDUÇÃO', 0.0)
                    if pd.isna(deducao): deducao = 0.0
                    valor = raw_valor - deducao
                    payment_status = str(row.get('PAGAMENTO', 'Não')).strip().lower()
                    is_paid = payment_status in ['sim', 'idd']
                    
                    invoice_iss_original = 0.0
                    if is_paid:
                        aliquota_difference = aliquota_dec - declared_rate_decimal
                        invoice_iss_original = max(0, aliquota_difference * valor)
                    else: 
                        invoice_iss_original = aliquota_dec * valor
                        
                    temp_df.loc[idx, 'iss_original_calculado'] = invoice_iss_original
                
                calculated_iss_original = temp_df['iss_original_calculado'].sum()
                # --- END: Updated Aliquot/Tax Calculation Logic ---
        
        # ✅ --- START: Apply filters to ASSIGNED table ---
        filtered_assigned_df = self._apply_filters_to_df(assigned_df, self.assigned_filters)
        self.populate_table_with_df(self.assigned_invoices_table, filtered_assigned_df)
        # ✅ --- END: Apply filters ---
        
        # ✅ Summary label uses the calculated total which now respects individual invoice rates
        self.assigned_summary_label.setText(format_summary_label(assigned_df, iss_value=calculated_iss_original, iss_label="ISS Calculado"))
        
        self.populate_available_table()

    def create_invoice_table(self):
        # CHANGE: Use QTableView instead of QTableWidget
        table = QTableView()
        table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        table.setEditTriggers(QTableView.EditTriggers.NoEditTriggers)
        table.setAlternatingRowColors(False)
        table.setSortingEnabled(True)
        
        # Performance setting for QTableView
        table.verticalHeader().setVisible(False) # Hiding row numbers saves huge repaint time
        
        # Connect Header Filter (Updated below)
        from app.excel_filter import FilterableHeaderView
        header = FilterableHeaderView(table)
        header.filter_changed.connect(self.apply_excel_filters)
        table.setHorizontalHeader(header)
        
        table.column_mapping = self.wizard.visible_columns 
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        
        return table

    def apply_excel_filters(self, col_index, allowed_values):
        table = self.available_invoices_table
        model = table.model()
        
        if not model: return
        
        # Access the raw DF directly from the model for speed
        df = model._df 
        
        # Iterate DataFrame rows efficiently
        for row_idx in range(len(df)):
            should_show = True
            
            # Check against active filters
            for f_col_idx, allowed in table.horizontalHeader().filters.items():
                # Get value from DF safely
                col_name = self.wizard.visible_columns[f_col_idx]
                val = str(df.iloc[row_idx][col_name])
                
                if val not in allowed:
                    should_show = False
                    break
            
            table.setRowHidden(row_idx, not should_show)

    def mark_as_ignored(self):
        """
        Sets a flag on the selected invoices so they disappear from the 'Available' list
        unless 'Show Ignored' is checked.
        """
        indices = self.get_selected_indices_from_table(self.available_invoices_table)
        if not indices:
            QMessageBox.warning(self, "Aviso", "Selecione notas para ignorar.")
            return

        # Update DataFrame
        for idx in indices:
            self.wizard.all_invoices_df.at[idx, 'status_manual'] = 'Ignored'
            # Optional: Clear any existing auto assignment just in case
            # (Though if it's in available table, it shouldn't be assigned)

        # Refresh
        self.populate_available_table()

    # ✅ --- START: REQ 1: Novas funções para menu de colunas ---
    def show_column_context_menu(self, pos):
        """Cria e exibe um menu de clique-direito para mostrar/ocultar colunas."""
        header = self.sender() # Obtém o cabeçalho que emitiu o sinal
        if not isinstance(header, QHeaderView):
            return
            
        all_columns = self.wizard.all_invoices_df.columns
        visible_columns = self.wizard.visible_columns
        
        menu = QMenu(self)
        
        for column_name in all_columns:
            action = QAction(column_name, self)
            action.setCheckable(True)
            action.setChecked(column_name in visible_columns)
            
            # Conecta o sinal 'toggled' passando o nome da coluna
            action.toggled.connect(
                # Usamos uma expressão lambda para "travar" o valor de column_name
                lambda checked, name=column_name: self.handle_column_visibility_toggle(checked, name)
            )
            menu.addAction(action)
            
        menu.exec(header.mapToGlobal(pos))

    def handle_column_visibility_toggle(self, is_checked, column_name):
        """Adiciona ou remove uma coluna das colunas visíveis."""
        if is_checked:
            if column_name not in self.wizard.visible_columns:
                # Tenta inserir de forma ordenada, se possível, ou apenas adiciona
                try:
                    all_cols_list = list(self.wizard.all_invoices_df.columns)
                    insert_pos = all_cols_list.index(column_name)
                    
                    # Encontra a posição correta para inserir na lista visível
                    # (isso é complexo, vamos simplificar por agora)
                    # Simplificação: apenas adiciona no final
                    self.wizard.visible_columns.append(column_name)
                    # TODO: Reordenar self.wizard.visible_columns com base na ordem de all_cols_list
                except ValueError:
                    self.wizard.visible_columns.append(column_name) # Fallback
        else:
            if column_name in self.wizard.visible_columns:
                self.wizard.visible_columns.remove(column_name)
                
        self.refresh_all_tables_and_filters()
    # ✅ --- END: REQ 1 ---

    def on_auto_selection_change(self, current, previous):
        # Enable/Disable Split Button based on selection
        self.split_auto_button.setEnabled(False) # Default disable
        if current:
            auto_id = current.data(Qt.ItemDataRole.UserRole)
            if auto_id and auto_id in self.wizard.autos:
                auto_data = self.wizard.autos[auto_id]
                motive = auto_data.get('motive', '')
                # Enable only if motive starts with 'Alíquota Incorreta'
                if motive.startswith('Alíquota Incorreta'):
                    self.split_auto_button.setEnabled(True)

        self.refresh_all_tables() # Refresh tables as usual

    def create_new_auto(self, invoices_df=None):
        dialog = NewAutoDialog(self)
        dialog.auto_id_edit.setText(f"AUTO-{self.wizard.auto_counter:03d}")
        
        if dialog.exec():
            data = dialog.get_data()
            auto_id, motive = data['auto_id'], data['motive']
            if not auto_id or not motive:
                QMessageBox.warning(self, "Dados Inválidos", "O Nº do Auto e o Motivo não podem estar vazios.")
                return None 
            if auto_id in self.wizard.autos:
                QMessageBox.warning(self, "ID Duplicado", "Já existe um auto com este número.")
                return None 
            
            rule_name = self.wizard.motive_to_rule_map.get(motive, 'regra_desconhecida')
            
            df = invoices_df if invoices_df is not None else pd.DataFrame(columns=self.wizard.all_invoices_df.columns)
            
            new_auto_data = {
                'motive': motive, 
                'rule_name': rule_name, 
                'df': df,
                'user_defined_credito': data.get('user_defined_credito'),
                'auto_text': ''
            }
            if data['correct_aliquota'] is not None:
                new_auto_data['user_defined_aliquota'] = data['correct_aliquota']

            self.wizard.autos[auto_id] = new_auto_data
            item = QListWidgetItem(f"{auto_id}: {motive}")
            item.setData(Qt.ItemDataRole.UserRole, auto_id)
            self.autos_list_widget.addItem(item)
            self.autos_list_widget.setCurrentItem(item)
            self.wizard.auto_counter += 1
            
            self.wizard.update_tab_states()
            
            # ✅ MARK DIRTY
            self.wizard.mark_dirty()

            if hasattr(self.wizard, 'fine_page'):
                self.wizard.fine_page.update_fine_text()
            
            return auto_id

    def edit_selected_auto(self):
        current_item = self.autos_list_widget.currentItem()
        if not current_item: return
        original_auto_id = current_item.data(Qt.ItemDataRole.UserRole)
        auto_data = self.wizard.autos[original_auto_id]
        
        dialog = NewAutoDialog(
            self,
            auto_id=original_auto_id,
            motive=auto_data['motive'],
            correct_aliquota=auto_data.get('user_defined_aliquota')
        )

        if dialog.exec():
            new_data = dialog.get_data()
            new_auto_id, new_motive = new_data['auto_id'], new_data['motive']
            new_aliquota = new_data['correct_aliquota']
            
            if not new_auto_id:
                QMessageBox.warning(self, "ID Inválido", "O Nº do Auto não pode estar vazio.")
                return
            if new_auto_id != original_auto_id and new_auto_id in self.wizard.autos:
                QMessageBox.warning(self, "ID Duplicado", "Já existe um auto com este número.")
                return
            
            data_to_update = self.wizard.autos.pop(original_auto_id)
            data_to_update['motive'] = new_motive
            data_to_update['rule_name'] = self.wizard.motive_to_rule_map.get(new_motive.split(' (')[0], 'regra_desconhecida')
            
            if new_aliquota is not None:
                data_to_update['user_defined_aliquota'] = new_aliquota
            elif 'user_defined_aliquota' in data_to_update:
                del data_to_update['user_defined_aliquota']
            
            if new_motive != auto_data['motive']:
                data_to_update['auto_text'] = ''
                
            self.wizard.autos[new_auto_id] = data_to_update
            current_item.setText(f"{new_auto_id}: {new_motive}")
            current_item.setData(Qt.ItemDataRole.UserRole, new_auto_id)
            
            self.refresh_all_tables()
            
            # ✅ CRITICAL: Ensure dirty flag is set so Preview Page recalculates names
            self.wizard.mark_dirty()

    def remove_selected_auto(self):
        current_item = self.autos_list_widget.currentItem()
        if not current_item: return
        auto_id = current_item.data(Qt.ItemDataRole.UserRole)
        reply = QMessageBox.question(self, "Confirmar Remoção", f"Tem a certeza que quer remover o '{auto_id}'? As Notas atribuídas voltarão a estar disponíveis.")
        if reply == QMessageBox.StandardButton.Yes:
            del self.wizard.autos[auto_id]
            self.autos_list_widget.takeItem(self.autos_list_widget.row(current_item))
            self.refresh_all_tables()
            self.wizard.mark_dirty()

    def get_selected_indices_from_table(self, table):
        indices = set()
        selection_model = table.selectionModel()
        
        # Get all selected rows
        if selection_model:
            selected_rows = selection_model.selectedRows()
            for model_index in selected_rows:
                # We defined UserRole to return the original DF index in the model
                idx = model_index.data(Qt.UserRole)
                if idx is not None:
                    indices.add(idx)
                    
        return list(indices)

    def move_to_assigned(self):
        auto_id = self.get_current_auto_id()
        if not auto_id:
            QMessageBox.information(self, "Nenhum Auto Selecionado", "Crie ou selecione um auto primeiro.")
            return
        indices_to_move = self.get_selected_indices_from_table(self.available_invoices_table)
        if not indices_to_move: return

        new_motive = self.wizard.autos[auto_id].get('motive', 'compliant')
        if not new_motive or new_motive == 'compliant': 
            print(f"Warning: Attempting to assign to auto {auto_id} with invalid motive '{new_motive}'.")

        for index in indices_to_move:
            self.wizard.all_invoices_df.at[index, 'primary_infraction_group'] = new_motive 
            current_details_obj = self.wizard.all_invoices_df.at[index, Columns.BROKEN_RULE_DETAILS]
            current_details = list(current_details_obj) if isinstance(current_details_obj, list) else []
            if new_motive not in current_details:
                new_details = [new_motive] + current_details 
            else:
                new_details = current_details
            self.wizard.all_invoices_df.at[index, Columns.BROKEN_RULE_DETAILS] = new_details

        invoices_to_add = self.wizard.all_invoices_df.loc[indices_to_move]
        
        current_auto_df = self.wizard.autos[auto_id].get('df')
        if not isinstance(current_auto_df, pd.DataFrame): 
             current_auto_df = pd.DataFrame(columns=invoices_to_add.columns)

        combined_df = pd.concat([current_auto_df, invoices_to_add])
        self.wizard.autos[auto_id]['df'] = combined_df[~combined_df.index.duplicated(keep='first')]
        self.wizard.autos[auto_id]['auto_text'] = ''
        
        self.refresh_all_tables()
        
        # ✅ MARK DIRTY
        self.wizard.mark_dirty()

        if hasattr(self.wizard, 'fine_page'):
            self.wizard.fine_page.update_fine_text()

  

    def move_to_available(self):
        auto_id = self.get_current_auto_id()
        if not auto_id: return
        indices_to_remove = self.get_selected_indices_from_table(self.assigned_invoices_table)
        if not indices_to_remove: return

        auto_motive = self.wizard.autos[auto_id].get('motive')
        if not auto_motive: 
            print(f"Warning: Auto {auto_id} has no motive set.")
            return

        for index in indices_to_remove:
            current_details_obj = self.wizard.all_invoices_df.at[index, Columns.BROKEN_RULE_DETAILS]
            current_details = list(current_details_obj) if isinstance(current_details_obj, list) else []
            new_details = [detail for detail in current_details if detail != auto_motive]
            new_primary_group = new_details[0] if new_details else 'compliant'
            self.wizard.all_invoices_df.at[index, 'primary_infraction_group'] = new_primary_group
            self.wizard.all_invoices_df.at[index, Columns.BROKEN_RULE_DETAILS] = new_details

        current_df = self.wizard.autos[auto_id].get('df')
        if isinstance(current_df, pd.DataFrame):
            self.wizard.autos[auto_id]['df'] = current_df.drop(indices_to_remove, errors='ignore')
        else:
            self.wizard.autos[auto_id]['df'] = pd.DataFrame() 
        
        self.wizard.autos[auto_id]['auto_text'] = ''
        self.refresh_all_tables()
        
        # ✅ MARK DIRTY
        self.wizard.mark_dirty()

        if hasattr(self.wizard, 'fine_page'):
            self.wizard.fine_page.update_fine_text()

    def correct_infractions(self):
        selected_indices = self.get_selected_indices_from_table(self.assigned_invoices_table)
        selected_indices.extend(self.get_selected_indices_from_table(self.available_invoices_table))
        unique_indices = list(set(selected_indices))
        if not unique_indices:
            QMessageBox.information(self, "Nenhuma Nota Selecionada", "Selecione uma ou mais notas para corrigir.")
            return
        
        all_infractions = set()
        for index in unique_indices:
            details = self.wizard.all_invoices_df.loc[index, Columns.BROKEN_RULE_DETAILS]
            if isinstance(details, list):
                all_infractions.update(details)
        
        if not all_infractions:
            QMessageBox.information(self, "Nenhuma Infração", "As notas selecionadas não possuem infrações para corrigir.")
            return
            
        dialog = InfractionCorrectionDialog(all_infractions, self)
        if dialog.exec():
            infractions_to_keep = dialog.get_infractions_to_keep()
            for index in unique_indices:
                affected_auto_id = None
                
                for auto_id, auto_data in self.wizard.autos.items():
                    # ✅ --- START FIX: Check if 'df' is a DataFrame before checking index/dropping ---
                    df = auto_data.get('df')
                    if isinstance(df, pd.DataFrame) and index in df.index:
                        auto_data['df'] = df.drop(index)
                    # ✅ --- END FIX ---
                        affected_auto_id = auto_id
                        break
                
                if affected_auto_id:
                    self.wizard.autos[affected_auto_id]['auto_text'] = ''

                self.wizard.all_invoices_df.loc[index, Columns.BROKEN_RULE_DETAILS] = infractions_to_keep
                new_primary_group = infractions_to_keep[0] if infractions_to_keep else 'compliant'
                self.wizard.all_invoices_df.loc[index, 'primary_infraction_group'] = new_primary_group
                
                if new_primary_group != 'compliant':
                    target_auto_id = None
                    for auto_id, auto_data in self.wizard.autos.items():
                        if auto_data.get('motive') == new_primary_group:
                            target_auto_id = auto_id
                            break
                    
                    if not target_auto_id: 
                        mock_auto_data = {'motive': new_primary_group, 'auto_id': f"AUTO-{self.wizard.auto_counter:03d}", 'correct_aliquota': None}
                        target_auto_id = self.create_new_auto() 
                    
                    if target_auto_id:
                        invoice_to_move = self.wizard.all_invoices_df.loc[[index]]
                        
                        # ✅ --- START FIX: Check if 'df' is a DataFrame before concat ---
                        target_df = self.wizard.autos[target_auto_id].get('df')
                        if not isinstance(target_df, pd.DataFrame):
                            target_df = pd.DataFrame(columns=invoice_to_move.columns)
                        self.wizard.autos[target_auto_id]['df'] = pd.concat([target_df, invoice_to_move])
                        # ✅ --- END FIX ---
                        
                        self.wizard.autos[target_auto_id]['auto_text'] = ''

            self.refresh_all_tables()
            if hasattr(self.wizard, 'fine_page'):
                self.wizard.fine_page.update_fine_text()

    def open_column_configuration(self):
        """Abre o diálogo de fallback para configurar colunas."""
        all_possible_display_cols = list(self.wizard.all_invoices_df.columns)
        dialog = ColumnSelectionDialog(all_possible_display_cols, self.wizard.visible_columns, self)
        if dialog.exec():
            self.wizard.visible_columns = dialog.get_selected_columns()
            self.refresh_all_tables_and_filters() # ✅ Chama a nova função

    def refresh_all_tables_and_filters(self):
        """
        ✅ Nova função wrapper que atualiza as tabelas E as
        listas de colunas nos dropdowns de filtro.
        """
        # Atualiza os dropdowns de filtro
        all_filters = self.assigned_filters + self.available_filters
        for f in all_filters:
            combo = f["combo"]
            current_selection = combo.currentText()
            
            combo.blockSignals(True) # Impede a execução do filtro
            combo.clear()
            combo.addItems(self.wizard.visible_columns)
            if current_selection in self.wizard.visible_columns:
                combo.setCurrentText(current_selection)
            combo.blockSignals(False) # Reabilita os sinais
            
            # ✅ Atualiza o completer associado
            self._update_filter_completer(f, lambda: None) # Passa um slot vazio
            
        # Atualiza as tabelas
        self.refresh_all_tables()

    

    def get_current_auto_id(self):
        current_item = self.autos_list_widget.currentItem()
        return current_item.data(Qt.ItemDataRole.UserRole) if current_item else None

    def view_assigned_details(self):
        auto_id = self.get_current_auto_id()
        df = pd.DataFrame()
        if auto_id and auto_id in self.wizard.autos:
            # ✅ --- START FIX: Check if 'df' is a DataFrame ---
            df_check = self.wizard.autos[auto_id].get('df')
            if isinstance(df_check, pd.DataFrame):
                df = df_check
            # ✅ --- END FIX ---
        
        if df.empty:
            QMessageBox.information(self, "Sem Notas", "Nenhuma nota atribuída a este auto para ver em detalhe.")
            return
        
        all_cols = self.wizard.all_invoices_df.columns.tolist()
        dialog = InvoiceDetailViewerDialog(df, all_cols, self)
        dialog.exec()

    def view_available_details(self):
        df = self.current_available_df
        if df.empty:
            QMessageBox.information(self, "Sem Notas", "Nenhuma nota disponível para ver em detalhe (ou o filtro está muito restrito).")
            return
        
        all_cols = self.wizard.all_invoices_df.columns.tolist()
        dialog = InvoiceDetailViewerDialog(df, all_cols, self)
        dialog.exec()

    def flag_available_infraction(self):
        indices_to_flag = self.get_selected_indices_from_table(self.available_invoices_table)
        if not indices_to_flag:
            QMessageBox.warning(self, "Nenhuma Nota Selecionada", "Por favor, selecione uma ou mais notas da tabela 'Notas Disponíveis' para flagar.")
            return

        invoices_df = self.wizard.all_invoices_df.loc[indices_to_flag]
        
        new_auto_id = self.create_new_auto(invoices_df=invoices_df)
        
        if new_auto_id:
            new_motive = self.wizard.autos[new_auto_id].get('motive', 'compliant')
            for index in indices_to_flag:
                self.wizard.all_invoices_df.loc[index, 'primary_infraction_group'] = new_motive
                details = self.wizard.all_invoices_df.loc[index, Columns.BROKEN_RULE_DETAILS]
                if not isinstance(details, list): details = []
                if new_motive not in details:
                    details.insert(0, new_motive) 
                self.wizard.all_invoices_df.loc[index, Columns.BROKEN_RULE_DETAILS] = details
            
            self.wizard.autos[new_auto_id]['auto_text'] = ''
            
            self.refresh_all_tables()

            if hasattr(self.wizard, 'fine_page'):
                self.wizard.fine_page.update_fine_text()
            
    def export_autos_to_excel(self):
        if not self.wizard.autos:
            QMessageBox.warning(self, "Nada para Exportar", "Nenhum auto de infração foi criado.")
            return

        default_filename = f"export_autos_{self.wizard.company_cnpj.replace('/', '-')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        filepath, _ = QFileDialog.getSaveFileName(self, "Salvar Exportação de Autos", default_filename, "Excel Files (*.xlsx)")
        
        if not filepath:
            return

        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                self.populate_available_table() 
                if not self.current_available_df.empty:
                    self.current_available_df.to_excel(writer, sheet_name="Notas Nao Autuadas", index=False)
                
                for auto_id, auto_data in self.wizard.autos.items():
                    df = auto_data.get('df')
                    # ✅ --- START FIX: Check if 'df' is a DataFrame before exporting ---
                    if isinstance(df, pd.DataFrame) and not df.empty:
                    # ✅ --- END FIX ---
                        sheet_name = auto_id.replace(":", "").replace("/", "-").replace(" ", "_")
                        sheet_name = sheet_name[:31] 
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
            QMessageBox.information(self, "Exportação Concluída", f"Os autos de infração foram exportados com sucesso para:\n{filepath}")
        except Exception as e:
            QMessageBox.critical(self, "Erro na Exportação", f"Ocorreu um erro ao salvar o ficheiro Excel:\n{e}")

    def isComplete(self):
        if not self.wizard.autos:
            return False
        # ✅ --- START FIX: Check if 'df' is a DataFrame before checking .empty ---
        return any(
            isinstance(data.get('df'), pd.DataFrame) and not data['df'].empty 
            for data in self.wizard.autos.values()
        )
        # ✅ --- END FIX ---

# --- Page 2: Fine Details ---

class FineDetailsPage(QWidget):
    def __init__(self, wizard):
        super().__init__()
        self.wizard = wizard

        main_page_layout = QVBoxLayout(self)

        self.subtitle_label = QLabel()
        self.subtitle_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.subtitle_label.setStyleSheet("background-color: white; color: black; font-style: italic; padding: 5px; border: 1px solid #ddd;")
        main_page_layout.addWidget(self.subtitle_label)
        
        tab_layout = QHBoxLayout()
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        fine_group = QGroupBox("Dados da Multa")
        fine_layout = QVBoxLayout()

        # ✅ REMOVED: EPAF Field was here. Now it's gone.
        # (The old QFormLayout and epaf_numero_edit code is deleted)

        # --- Multas Table ---
        self.fines_table = QTableWidget()
        self.fines_table.setColumnCount(3)
        self.fines_table.setHorizontalHeaderLabels(["Ano Exercício", "Nº da Multa", "Valor (R$)"])
        self.fines_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.fines_table.verticalHeader().setVisible(False)
        fine_layout.addWidget(self.fines_table)
        
        # Buttons for Table
        fine_btns = QHBoxLayout()
        add_fine_btn = QPushButton("Adicionar Multa")
        add_fine_btn.clicked.connect(self.add_fine_row)
        remove_fine_btn = QPushButton("Remover Selecionada")
        remove_fine_btn.clicked.connect(self.remove_fine_row)
        fine_btns.addWidget(add_fine_btn)
        fine_btns.addWidget(remove_fine_btn)
        fine_layout.addLayout(fine_btns)

        # Hints
        self.meses_multa_label = QLabel("Meses com infrações: N/A")
        self.meses_multa_label.setWordWrap(True)
        fine_layout.addWidget(self.meses_multa_label)
        
        self.fine_hint_label = QLabel("")
        self.fine_hint_label.setStyleSheet("font-style: italic; color: #555;")
        fine_layout.addWidget(self.fine_hint_label)

        fine_group.setLayout(fine_layout)
        
        # DAM Group
        dam_group = QGroupBox("Ficheiro DAM Avulso (Opcional)")
        dam_layout = QHBoxLayout()
        self.dam_filepath_edit = QLineEdit()
        self.dam_filepath_edit.setPlaceholderText("Selecione o ficheiro CSV de DAMs pagos avulsos...")
        self.dam_filepath_edit.setReadOnly(True)
        browse_dam_btn = QPushButton("Procurar...")
        browse_dam_btn.clicked.connect(self.browse_dam_file)
        dam_layout.addWidget(self.dam_filepath_edit)
        dam_layout.addWidget(browse_dam_btn)
        dam_group.setLayout(dam_layout)

        # PGDAS Group
        pgdas_group = QGroupBox("Pasta PGDAS (Opcional)")
        pgdas_layout = QHBoxLayout()
        self.pgdas_folder_path_edit = QLineEdit()
        self.pgdas_folder_path_edit.setPlaceholderText("Selecione a pasta com os ficheiros PGDASD*.pdf...")
        self.pgdas_folder_path_edit.setReadOnly(True)
        browse_pgdas_btn = QPushButton("Procurar Pasta...")
        browse_pgdas_btn.clicked.connect(self.browse_pgdas_folder)
        pgdas_layout.addWidget(self.pgdas_folder_path_edit)
        pgdas_layout.addWidget(browse_pgdas_btn)
        pgdas_group.setLayout(pgdas_layout)
        
        left_layout.addWidget(fine_group)
        left_layout.addWidget(dam_group)
        left_layout.addWidget(pgdas_group)
        left_layout.addStretch()
        
        # Right Text Group
        fine_text_group = QGroupBox("Texto da Multa para Copiar")
        fine_text_layout = QVBoxLayout()
        self.fine_text_edit = QTextEdit()
        # Make text editable
        self.fine_text_edit.setReadOnly(False)
        fine_text_layout.addWidget(self.fine_text_edit)
        fine_text_group.setLayout(fine_text_layout)
        
        tab_layout.addWidget(left_panel, 1)
        tab_layout.addWidget(fine_text_group, 1)

        main_page_layout.addLayout(tab_layout)

        # Initialize with one empty row
        self.add_fine_row()

    def add_fine_row(self):
        row = self.fines_table.rowCount()
        self.fines_table.insertRow(row)
        
        # Year (Default to current year)
        self.fines_table.setItem(row, 0, QTableWidgetItem(str(datetime.now().year)))
        
        # Number
        self.fines_table.setItem(row, 1, QTableWidgetItem(""))
        
        # Value (LineEdit with validator)
        val_edit = QLineEdit()
        locale = QLocale(QLocale.Language.Portuguese, QLocale.Country.Brazil)
        validator = QDoubleValidator(0.0, 9999999.99, 2)
        validator.setLocale(locale)
        val_edit.setValidator(validator)
        val_edit.setPlaceholderText("0,00")
        # Connect text changed to update hint logic (for total calc)
        val_edit.textChanged.connect(self.update_fine_text) 
        self.fines_table.setCellWidget(row, 2, val_edit)

    def remove_fine_row(self):
        cur = self.fines_table.currentRow()
        if cur >= 0:
            self.fines_table.removeRow(cur)
            self.update_fine_text()

    def get_fines_data(self):
        """Returns list of dicts: [{'year': '2024', 'number': '...', 'value': '...'}]"""
        fines = []
        for r in range(self.fines_table.rowCount()):
            year_item = self.fines_table.item(r, 0)
            num_item = self.fines_table.item(r, 1)
            val_widget = self.fines_table.cellWidget(r, 2)
            
            y = year_item.text().strip() if year_item else ""
            n = num_item.text().strip() if num_item else ""
            v = val_widget.text().strip() if val_widget else "0,00"
            
            # Allow saving partial rows (will be filtered in context builder if needed)
            fines.append({'year': y, 'number': n, 'value': v})
        return fines

    def set_fines_data(self, fines_list):
        """Populates the table from a list of dicts."""
        self.fines_table.setRowCount(0)
        for f in fines_list:
            row = self.fines_table.rowCount()
            self.fines_table.insertRow(row)
            self.fines_table.setItem(row, 0, QTableWidgetItem(str(f.get('year', ''))))
            self.fines_table.setItem(row, 1, QTableWidgetItem(str(f.get('number', ''))))
            
            val_edit = QLineEdit(str(f.get('value', '')))
            locale = QLocale(QLocale.Language.Portuguese, QLocale.Country.Brazil)
            validator = QDoubleValidator(0.0, 9999999.99, 2)
            validator.setLocale(locale)
            val_edit.setValidator(validator)
            val_edit.textChanged.connect(self.update_fine_text)
            self.fines_table.setCellWidget(row, 2, val_edit)
        
        # If empty list passed, add one default row
        if self.fines_table.rowCount() == 0:
            self.add_fine_row()

    def load_dam_file_programmatically(self, file_path):
        self.dam_filepath_edit.setText(file_path)
        try:
            if os.path.exists(file_path):
                self.wizard.dam_payments_map = _load_and_process_dams(file_path)
                print(f"✅ Auto-loaded DAMs: {len(self.wizard.dam_payments_map)} records from {file_path}")
                # ✅ MARK DIRTY (DAMs affect calculation)
                self.wizard.mark_dirty()
            else:
                print(f"❌ File not found during auto-load: {file_path}")
        except Exception as e:
            print(f"❌ Error auto-loading DAMs: {e}")

    # 3. Update browse_dam_file to use the logic (optional cleanup, or leave as is)
    def browse_dam_file(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Selecionar Ficheiros de DAMs (Um ou Vários)", 
            "", 
            "DAM Files (*.csv *.xls *.xlsx)"
        )
        
        if not files:
            return

        try:
            dfs = []
            for f in files:
                try:
                    df = None
                    if f.lower().endswith('.csv'):
                        # Try combinations to handle encoding/separators correctly
                        attempts = [
                            {'sep': ',', 'encoding': 'utf-8-sig'}, # Best for BOM files
                            {'sep': ';', 'encoding': 'utf-8-sig'},
                            {'sep': ';', 'encoding': 'latin1'},
                            {'sep': ',', 'encoding': 'latin1'}
                        ]
                        
                        for params in attempts:
                            try:
                                temp_df = pd.read_csv(f, on_bad_lines='skip', **params)
                                # Basic validation: did we get columns?
                                if len(temp_df.columns) > 1:
                                    df = temp_df
                                    break
                            except: continue
                            
                        # Last resort fallback
                        if df is None:
                            df = pd.read_csv(f, sep=None, engine='python', on_bad_lines='skip')
                            
                    else:
                        df = pd.read_excel(f)
                    
                    if df is not None and not df.empty:
                        # Normalize immediately to fix the headers
                        df = normalize_dam_dataframe(df)
                        dfs.append(df)
                        
                except Exception as e:
                    print(f"Skipping corrupt DAM file {f}: {e}")

            if not dfs:
                QMessageBox.warning(self, "Erro", "Nenhum dado válido encontrado.")
                return

            full_df = pd.concat(dfs, ignore_index=True)
            full_df.drop_duplicates(inplace=True)

            # Final check before saving
            if 'codigoVerificacao' not in full_df.columns:
                # If still missing, try by index as last resort
                if len(full_df.columns) >= 1:
                    print("Warning: Forced renaming of column 0 to codigoVerificacao")
                    full_df.rename(columns={full_df.columns[0]: 'codigoVerificacao'}, inplace=True)

            # Save clean temp file with standard Comma separator
            temp = tempfile.NamedTemporaryFile(delete=False, suffix="_clean_dams.csv", mode='w', encoding='utf-8')
            full_df.to_csv(temp.name, sep=',', index=False)
            temp.close()

            self.dam_filepath_edit.setText(temp.name)
            
            # Load
            self.wizard.dam_payments_map = _load_and_process_dams(temp.name)
            
            count = sum(len(v) for v in self.wizard.dam_payments_map.values())
            QMessageBox.information(self, "DAMs Carregados", 
                                    f"Sucesso! {len(files)} arquivos processados.\n"
                                    f"{count} pagamentos carregados na memória.")

        except Exception as e:
            QMessageBox.critical(self, "Erro Fatal", f"Falha ao processar: {e}")
            traceback.print_exc()

    def browse_pgdas_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta PGDAS")
        if folder_path:
            self.pgdas_folder_path_edit.setText(folder_path)
            try:
                # Load map immediately
                # We need a dummy progress emitter
                class DummyEmitter:
                    def emit(self, msg): print(msg)
                
                self.wizard.pgdas_payments_map = _load_and_process_pgdas(folder_path, DummyEmitter())
                QMessageBox.information(self, "PGDAS Carregados", f"{len(self.wizard.pgdas_payments_map)} pagamentos PGDAS carregados.")
            except Exception as e:
                QMessageBox.critical(self, "Erro ao Ler PGDAS", f"Não foi possível processar a pasta: {e}")

    def update_fine_text(self):
        # ... (Existing logic for instrumental fines extraction remains the same) ...
        
        all_assigned_invoices = [
            auto_data['df'] for auto_data in self.wizard.autos.values() 
            if isinstance(auto_data.get('df'), pd.DataFrame) and not auto_data['df'].empty
        ]
        
        if not all_assigned_invoices:
            self.fine_text_edit.clear()
            self.meses_multa_label.setText("Meses com infrações: Nenhuma nota atribuída.")
            return
            
        combined_df = pd.concat(all_assigned_invoices).drop_duplicates(subset=[Columns.INVOICE_NUMBER])
        
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
            
        df_instrumental = combined_df[
            combined_df[Columns.BROKEN_RULE_DETAILS].apply(check_for_instrumental)
        ].copy() 
        
        fine_hints = {1: "R$ 229,96", 2: "R$ 459,92", 3: "R$ 689,88", 4: "R$ 919,84", 5: "R$ 1.149,80"}

        if df_instrumental.empty:
            self.meses_multa_label.setText("Meses com infrações: Nenhuma infração instrumental encontrada.")
            self.fine_hint_label.setText("") 
        else:
            try:
                df_instrumental['DATA EMISSÃO'] = pd.to_datetime(df_instrumental['DATA EMISSÃO'], errors='coerce')
                unique_periods = df_instrumental['DATA EMISSÃO'].dropna().dt.to_period('M').unique()
                sorted_periods = sorted(list(unique_periods))
                months_str = ", ".join([p.strftime('%m/%Y') for p in sorted_periods])
                self.meses_multa_label.setText(f"Meses com infrações: {months_str}")

                num_months = len(unique_periods)
                if num_months in fine_hints:
                    self.fine_hint_label.setText(f"Valor sugerido para {num_months} mês(es): {fine_hints[num_months]}")
                elif num_months > 5:
                    self.fine_hint_label.setText(f"Valor sugerido para {num_months} meses (limite 5): {fine_hints[5]}")
                else: 
                    self.fine_hint_label.setText("")
            except Exception as e:
                self.meses_multa_label.setText(f"Erro ao extrair meses: {e}")
                self.fine_hint_label.setText("") 

        # Sum values for placeholder text
        total_val = 0.0
        fines = self.get_fines_data() # ✅ Retrieve current table data
        for f in fines:
             try:
                 clean_v = str(f['value']).replace("R$", "").strip().replace(".", "").replace(",", ".")
                 total_val += float(clean_v)
             except: pass
        
        valor_multa_str = f"R$ {total_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        # ✅ CHANGED: Pass 'fines' list to formatar_texto_multa
        texto_multa_gerado = formatar_texto_multa({}, df_instrumental, valor_multa_str, multas_list=fines)
        
        self.fine_text_edit.setText(texto_multa_gerado)
        
# --- Page 3: Preview and Edit Credits ---

class PreviewPage(QWidget):
    def __init__(self, wizard):
        super().__init__()
        self.wizard = wizard
        
        self.main_layout = QHBoxLayout(self) 

        # --- Left Panel (Master) ---
        left_panel_widget = QWidget()
        left_panel_layout = QVBoxLayout(left_panel_widget)
        left_panel_widget.setMaximumWidth(400) 

        # Auto selection list
        autos_group = QGroupBox("Autos de Infração")
        autos_layout = QVBoxLayout()
        self.autos_list_widget = QListWidget()
        self.autos_list_widget.currentItemChanged.connect(self.on_auto_selected)
        autos_layout.addWidget(self.autos_list_widget)
        autos_group.setLayout(autos_layout)
        left_panel_layout.addWidget(autos_group)

        summary_group = QGroupBox("Créditos Disponíveis (Restante / Original)")
        self.summary_layout = QFormLayout() 
        summary_group.setLayout(self.summary_layout)
        left_panel_layout.addWidget(summary_group)
        
        grand_total_group = QGroupBox("Resumo do Lançamento")
        self.grand_total_layout = QFormLayout()
        self.total_iss_autos_label = QLabel("R$ 0,00")
        self.total_multa_label = QLabel("R$ 0,00")
        self.total_geral_label = QLabel("R$ 0,00")
        self.grand_total_layout.addRow("Total ISS (Autos):", self.total_iss_autos_label)
        self.grand_total_layout.addRow("Total Multa:", self.total_multa_label)
        self.grand_total_layout.addRow("Total Geral (Crédito):", self.total_geral_label)
        grand_total_group.setLayout(self.grand_total_layout)
        left_panel_layout.addWidget(grand_total_group)

        left_panel_layout.addStretch()
        self.main_layout.addWidget(left_panel_widget)

        # --- Right Panel (Detail) ---
        right_panel_widget = QWidget()
        right_panel_layout = QVBoxLayout(right_panel_widget)
        
        controls_layout = QHBoxLayout()
        self.recalculate_btn = QPushButton("Atualizar Totais e Créditos Restantes")
        self.recalculate_btn.clicked.connect(self.recalculate_and_redraw)
        self.hide_empty_months_check = QCheckBox("Ocultar meses sem atividade")
        self.hide_empty_months_check.toggled.connect(self.filter_tables_view)
        controls_layout.addWidget(self.recalculate_btn)
        controls_layout.addWidget(self.hide_empty_months_check)
        controls_layout.addStretch()
        right_panel_layout.addLayout(controls_layout)

        self.detail_stack_layout = QVBoxLayout()
        right_panel_layout.addLayout(self.detail_stack_layout)
        
        self.main_layout.addWidget(right_panel_widget)
        
        self.credit_labels = {} 
        self.auto_detail_widgets = {} 
        self.manual_credit_widgets = {} 
        self.current_auto_key = None

    def initializePage(self):
        """Called when the page is shown."""
        # ✅ FIX: Only recalculate (resetting tables) if something in Step 1 or 2 actually changed.
        # If we are just switching tabs without changes, we keep the current context.
        if self.wizard.context_dirty:
            self.wizard.calculate_preview_context()
            self.wizard.context_dirty = False
        
        self.redraw_all()

    def on_auto_selected(self, current_item, previous_item):
        if previous_item:
            prev_key = previous_item.data(Qt.ItemDataRole.UserRole)
            if prev_key in self.auto_detail_widgets:
                container, table = self.auto_detail_widgets[prev_key]
                container.setVisible(False)
        
        target_year = None
        
        if current_item:
            self.current_auto_key = current_item.data(Qt.ItemDataRole.UserRole)
            if self.current_auto_key in self.auto_detail_widgets:
                container, table = self.auto_detail_widgets[self.current_auto_key]
                container.setVisible(True)
            
            target_year = self.wizard._get_auto_year(self.current_auto_key)

        self._update_credit_summary(target_year)

    def _update_credit_summary(self, target_year):
        for label in self.credit_labels.values():
            label.setParent(None)
            label.deleteLater()
        self.credit_labels = {} 
        
        while self.summary_layout.count():
             item = self.summary_layout.takeAt(0)
             if item.widget(): item.widget().deleteLater()

        context = self.wizard.preview_context
        if not context: return

        locale = QLocale(QLocale.Language.Portuguese, QLocale.Country.Brazil)
        startup_dam = context.get('available_credits_startup', {}).get('DAM', {})
        startup_pgdas = context.get('available_credits_startup', {}).get('PGDAS', {})
        available_credits = self.wizard.available_credits_map

        all_credit_periods = set()

        for p in startup_dam: 
            try:
                y = int(p.split('/')[1])
                if target_year is None or y == target_year:
                    all_credit_periods.add(p)
            except: pass
            
        for p in startup_pgdas:
            try:
                y = int(p.split('/')[1])
                if target_year is None or y == target_year:
                    all_credit_periods.add(p)
            except: pass

        def sort_key(period_str):
            try: return datetime.strptime(period_str, "%m/%Y")
            except ValueError:
                try: return datetime.strptime(period_str, "%#m/%Y")
                except ValueError: return datetime.min
        
        sorted_periods = sorted(list(all_credit_periods), key=sort_key)

        for period in sorted_periods:
            try: display_period = sort_key(period).strftime("%m/%Y")
            except: display_period = period

            pgdas_key = display_period 
            dam_key = f"{int(display_period.split('/')[0])}/{display_period.split('/')[1]}" 
            
            pgdas_orig = startup_pgdas.get(pgdas_key, 0.0)
            pgdas_rem = available_credits.get('PGDAS', {}).get(pgdas_key, 0.0)
            
            dam_list_orig = startup_dam.get(dam_key, [])
            dam_orig = sum(d['val'] for d in dam_list_orig)
            
            dam_list_rem = available_credits.get('DAM', {}).get(dam_key, [])
            dam_rem = sum(d['val'] for d in dam_list_rem)

            if pgdas_orig > 0.001:
                key = f"PGDAS {display_period}"
                label_text = f"R$ {locale.toString(float(pgdas_rem), 'f', 2)} / R$ {locale.toString(float(pgdas_orig), 'f', 2)}"
                self.credit_labels[key] = QLabel(label_text)
                self.summary_layout.addRow(key, self.credit_labels[key])
                
            if dam_orig > 0.001:
                key = f"DAM {display_period}"
                label_text = f"R$ {locale.toString(float(dam_rem), 'f', 2)} / R$ {locale.toString(float(dam_orig), 'f', 2)}"
                self.credit_labels[key] = QLabel(label_text)
                self.summary_layout.addRow(key, self.credit_labels[key])

    def redraw_all(self):
        selected_key = None
        if self.autos_list_widget.currentItem():
            selected_key = self.autos_list_widget.currentItem().data(Qt.ItemDataRole.UserRole)
        
        self.autos_list_widget.clear()
        
        for key, (container, table) in self.auto_detail_widgets.items():
            container.setParent(None); container.deleteLater()
        self.auto_detail_widgets = {}
        self.manual_credit_widgets = {} 

        context = self.wizard.preview_context
        locale = QLocale(QLocale.Language.Portuguese, QLocale.Country.Brazil)

        if not context:
            self.detail_stack_layout.addWidget(QLabel("Erro: Contexto vazio."))
            return
            
        for i, auto_data in enumerate(context.get('autos', [])):
            auto_key = auto_data['numero']
            item = QListWidgetItem(f"{auto_key} - {auto_data['motive_text']}")
            item.setData(Qt.ItemDataRole.UserRole, auto_key)
            self.autos_list_widget.addItem(item)
            
            container_widget = QWidget(); container_layout = QVBoxLayout(container_widget)
            
            override_group = QGroupBox("Valor Atualizado (Obrigatório)")
            ol = QHBoxLayout()
            ol.addWidget(QLabel("Valor Atualizado (R$):"))
            edit_manual = QLineEdit()
            
            user_credit = auto_data.get('user_defined_credito')
            
            if user_credit is not None:
                edit_manual.setText(locale.toString(float(user_credit), 'f', 2))
            else:
                edit_manual.setText("") 
                edit_manual.setPlaceholderText("Obrigatório")
            
            edit_manual.setValidator(QDoubleValidator(0.0, 999999999.0, 2))
            ol.addWidget(edit_manual)
            override_group.setLayout(ol)
            container_layout.addWidget(override_group)
            
            self.manual_credit_widgets[auto_key] = (None, edit_manual)

            table = self.create_preview_table(auto_data)
            container_layout.addWidget(table)
            
            self.detail_stack_layout.addWidget(container_widget)
            container_widget.setVisible(False)
            self.auto_detail_widgets[auto_key] = (container_widget, table)
            
        target_year = None
        
        if selected_key:
            items = self.autos_list_widget.findItems(selected_key, Qt.MatchFlag.MatchStartsWith)
            if items:
                self.autos_list_widget.setCurrentItem(items[0])
                target_year = self.wizard._get_auto_year(selected_key)
            elif self.autos_list_widget.count() > 0:
                 self.autos_list_widget.setCurrentRow(0)
                 target_year = self.wizard._get_auto_year(self.autos_list_widget.item(0).data(Qt.ItemDataRole.UserRole))
        elif self.autos_list_widget.count() > 0:
            self.autos_list_widget.setCurrentRow(0)
            target_year = self.wizard._get_auto_year(self.autos_list_widget.item(0).data(Qt.ItemDataRole.UserRole))

        self._update_credit_summary(target_year)
    
        # ✅ FIX: Handle None values safely
        total_autos = 0.0
        for a in context.get('autos', []):
            val = a.get('user_defined_credito')
            if val is None:
                val = a.get('totais', {}).get('iss_apurado_op', 0.0)
            try:
                total_autos += float(val)
            except (ValueError, TypeError):
                pass

        multas_list = context.get('summary', {}).get('multas', [])
        total_multa = sum(float(m.get('valor_credito', 0.0)) for m in multas_list)
        
        self.total_iss_autos_label.setText(f"R$ {locale.toString(float(total_autos), 'f', 2)}")
        self.total_multa_label.setText(f"R$ {locale.toString(float(total_multa), 'f', 2)}")
        self.total_geral_label.setText(f"R$ {locale.toString(float(total_autos + total_multa), 'f', 2)}")
    
        if self.autos_list_widget.count() == 0:
            self.detail_stack_layout.addWidget(QLabel("Nenhum auto de infração encontrado."))

        self.filter_tables_view()

    def filter_tables_view(self):
        is_checked = self.hide_empty_months_check.isChecked()
        for container, table in self.auto_detail_widgets.values():
            for row_idx in range(table.rowCount() - 1): 
                item = table.item(row_idx, 1) 
                is_hidden = False
                if item:
                    base_val_str = item.text().strip().replace("R$", "")
                    try:
                        val, ok = QLocale().toDouble(base_val_str)
                        if ok and val < 0.001: is_hidden = True
                        elif not ok and not base_val_str: is_hidden = True
                    except Exception:
                        if not base_val_str: is_hidden = True
                else:
                    is_hidden = True
                table.setRowHidden(row_idx, is_checked and is_hidden)

    def create_preview_table(self, auto_data):
        table = QTableWidget()
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive) 

        dados = auto_data.get('dados_anuais', [])
        table.setRowCount(len(dados) + 1)

        headers = ['Mês/Ano', 'Base de Cálculo', 'Alíquota (%)', 
                   'Alíq. Efetiva (%)', 
                   'ISS Apurado', 'DAS Pago', 'DAM Pago', 'ISS Constituído']
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)

        locale = QLocale(QLocale.Language.Portuguese, QLocale.Country.Brazil)
        validator = QDoubleValidator(0.0, 9999999.99, 2)
        validator.setLocale(locale)
        validator.setNotation(QDoubleValidator.Notation.StandardNotation)

        col_idx_map = { 
            'aliquota': 2, 
            'aliquota_efetiva': 3, 
            'iss_apurado': 4, 
            'das': 5, 
            'dam': 6, 
            'iss_op': 7
        }

        for row_idx, mes_data in enumerate(dados):
            table.setItem(row_idx, 0, NumericTableWidgetItem(mes_data['mes_ano']))
            
            base_calculo_val = float(mes_data.get('base_calculo', 0.0))
            base_calculo_str = f"{base_calculo_val:.2f}".replace(".", ",")
            table.setItem(row_idx, 1, NumericTableWidgetItem(base_calculo_str))
            
            aliquota_edit = QLineEdit(locale.toString(float(mes_data.get('aliquota_target_user', 0.0)), 'f', 2))
            aliquota_edit.setValidator(validator)
            aliquota_edit.setStyleSheet("background-color: white; color: black; padding-right: 3px;")
            aliquota_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
            aliquota_edit.setProperty("auto_key", auto_data['numero'])
            aliquota_edit.setProperty("row_idx", row_idx)
            aliquota_edit.setProperty("col_key", "aliquota_target_user")
            table.setCellWidget(row_idx, col_idx_map['aliquota'], aliquota_edit)

            efetiva_display_str = str(mes_data.get('aliquota_display', '-'))
            aliquota_efetiva_item = NumericTableWidgetItem(efetiva_display_str)
            aliquota_efetiva_item.setFlags(aliquota_efetiva_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            aliquota_efetiva_item.setBackground(QColor(230, 230, 230))
            aliquota_efetiva_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row_idx, col_idx_map['aliquota_efetiva'], aliquota_efetiva_item)

            iss_apurado_item = NumericTableWidgetItem(locale.toString(float(mes_data.get('iss_apurado', 0.0)), 'f', 2))
            iss_apurado_item.setFlags(iss_apurado_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            iss_apurado_item.setBackground(QColor(230, 230, 230))
            iss_apurado_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row_idx, col_idx_map['iss_apurado'], iss_apurado_item)
            
            das_pago_edit = QLineEdit(locale.toString(float(mes_data.get('das_iss_pago', 0.0)), 'f', 2))
            das_pago_edit.setValidator(validator)
            das_pago_edit.setStyleSheet("background-color: white; color: black; padding-right: 3px;")
            das_pago_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
            das_pago_edit.setProperty("auto_key", auto_data['numero'])
            das_pago_edit.setProperty("row_idx", row_idx)
            das_pago_edit.setProperty("col_key", "das_iss_pago")
            table.setCellWidget(row_idx, col_idx_map['das'], das_pago_edit)

            dam_pago_edit = QLineEdit(locale.toString(float(mes_data.get('dam_iss_pago', 0.0)), 'f', 2))
            dam_pago_edit.setValidator(validator)
            dam_pago_edit.setStyleSheet("background-color: white; color: black; padding-right: 3px;")
            dam_pago_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
            dam_pago_edit.setProperty("auto_key", auto_data['numero'])
            dam_pago_edit.setProperty("row_idx", row_idx)
            dam_pago_edit.setProperty("col_key", "dam_iss_pago")
            table.setCellWidget(row_idx, col_idx_map['dam'], dam_pago_edit)

            iss_op_item = NumericTableWidgetItem(locale.toString(float(mes_data.get('iss_apurado_op', 0.0)), 'f', 2))
            iss_op_item.setFlags(iss_op_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            iss_op_item.setBackground(QColor(230, 230, 230))
            iss_op_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row_idx, col_idx_map['iss_op'], iss_op_item)


        totais = auto_data.get('totais', {})
        total_row = len(dados)
        
        table.setItem(total_row, 0, NumericTableWidgetItem("TOTAL"))
        
        total_base_calculo_val = float(totais.get('base_calculo', 0.0))
        total_base_calculo_str = f"{total_base_calculo_val:.2f}".replace(".", ",")
        table.setItem(total_row, 1, NumericTableWidgetItem(total_base_calculo_str))
        
        table.setItem(total_row, 2, NumericTableWidgetItem("-")) 
        
        total_efetiva_str = str(totais.get('_total_aliquota_display', '-'))
        table.setItem(total_row, 3, NumericTableWidgetItem(total_efetiva_str)) 
        
        table.setItem(total_row, 4, NumericTableWidgetItem(locale.toString(float(totais.get('iss_apurado', 0.0)), 'f', 2)))
        table.setItem(total_row, 5, NumericTableWidgetItem(locale.toString(float(totais.get('das_iss_pago', 0.0)), 'f', 2)))
        table.setItem(total_row, 6, NumericTableWidgetItem(locale.toString(float(totais.get('dam_iss_pago', 0.0)), 'f', 2)))
        table.setItem(total_row, 7, NumericTableWidgetItem(locale.toString(float(totais.get('iss_apurado_op', 0.0)), 'f', 2)))

        for col in range(len(headers)): 
            item = table.item(total_row, col)
            if item:
                font = item.font()
                font.setBold(True)
                item.setFont(font)
                item.setBackground(QColor(210, 210, 210))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        table.resizeColumnsToContents()
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setMinimumSectionSize(80)
        table.horizontalHeader().setStretchLastSection(True)

        table.setEditTriggers(QTableWidget.EditTriggers.AllEditTriggers)
        return table

    def get_final_context(self):
        """
        Reads all QLineEdits, updates the context, and returns it.
        This is the "state" that will be passed to the generation worker.
        """
        self._read_tables_into_context()
        self._update_context_summary()
        return self.wizard.preview_context

    def recalculate_and_redraw(self):
        """
        Reads values from tables, updates the context, and redraws everything.
        """
        self._read_tables_into_context()
        
        # ✅ FIX: Do NOT call wizard.calculate_preview_context() here.
        # Calling it would reset manual DAM edits back to auto-calculated defaults.
        # Instead, we just update the summary based on the current context state.
        self._update_context_summary()
        
        self.redraw_all() # This will now use the preserved manual edits
        QMessageBox.information(self, "Recalculado", "Totais e créditos restantes foram atualizados com base nos seus dados.")

    def _update_context_summary(self):
        """
        Rebuilds 'preview_context['summary']' based on the current state of 'autos',
        preserving manual edits. Matches logic from ReviewWizard.calculate_preview_context.
        """
        context = self.wizard.preview_context
        
        summary_autos_list = []
        total_credito_autos = 0.0
        
        for auto in context.get('autos', []):
            auto_key = auto['numero']
            
            # Recalculate credit value logic
            auto_total_credito = auto.get('totais', {}).get('iss_apurado_op', 0.0)
            user_credit = auto.get('user_defined_credito')
            if user_credit is not None:
                auto_total_credito = float(user_credit)

            if auto_total_credito > 0.01:
                # Get invoice numbers from the source of truth (wizard.autos)
                invoice_list_str = "..."
                if auto_key in self.wizard.autos:
                    df = self.wizard.autos[auto_key].get('df')
                    if isinstance(df, pd.DataFrame) and not df.empty and Columns.INVOICE_NUMBER in df.columns:
                        nums = df[Columns.INVOICE_NUMBER].astype(str).tolist()
                        invoice_list_str = format_invoice_numbers(nums)

                summary_autos_list.append({
                    'numero': auto['numero'],
                    'nfs_tributadas': invoice_list_str,
                    'iss_valor_original': auto.get('totais', {}).get('iss_apurado_bruto', 0.0),
                    'total_credito_tributario': auto_total_credito,
                    'motivo': auto.get('motive_text', 'N/A')
                })
                total_credito_autos += auto_total_credito

        # Fines
        fines_list = self.wizard.fine_page.get_fines_data()
        processed_fines = []
        total_multas_val = 0.0
        
        for f in fines_list:
            try:
                val = float(str(f['value']).replace("R$", "").strip().replace(".", "").replace(",", "."))
                if val > 0.01:
                    total_multas_val += val
                    processed_fines.append({
                        'numero': f['number'],
                        'valor_credito': val,
                        'ano': f['year']
                    })
            except: pass

        context['summary'] = {
            'autos': summary_autos_list,
            'multas': processed_fines,
            'multa': processed_fines[0] if processed_fines else None,
            'total_geral_credito': total_credito_autos + total_multas_val
        }

    def _read_tables_into_context(self):
        locale = QLocale(QLocale.Language.Portuguese, QLocale.Country.Brazil)
        context = self.wizard.preview_context
        
        if not context or 'available_credits_startup' not in context:
            logging.warning("PreviewPage._read_tables_into_context called with invalid context.")
            return
            
        original_pgdas_map = self.wizard.pgdas_payments_map

        for auto_key, (_, edit) in self.manual_credit_widgets.items():
            wizard_auto_data = self.wizard.autos.get(auto_key)
            if not wizard_auto_data:
                continue
            
            text_val = edit.text().strip()
            if not text_val:
                wizard_auto_data['user_defined_credito'] = None
            else:
                val_num, ok = locale.toDouble(text_val)
                wizard_auto_data['user_defined_credito'] = val_num if ok else 0.0
            
            for auto in context.get('autos', []):
                if auto['numero'] == auto_key:
                    auto['user_defined_credito'] = wizard_auto_data['user_defined_credito']
                    break

        self.wizard.available_credits_map = {
            'DAM': copy.deepcopy(context['available_credits_startup']['DAM']),
            'PGDAS': copy.deepcopy(context['available_credits_startup']['PGDAS'])
        }

        for auto_data in context.get('autos', []):
            auto_key = auto_data['numero']
            is_split_diff_case = auto_data.get('is_split_diff', False)
            
            wizard_auto_data = self.wizard.autos.get(auto_key)
            if not wizard_auto_data: continue
            
            if 'monthly_overrides' not in wizard_auto_data:
                 wizard_auto_data['monthly_overrides'] = {}

            table = None
            if auto_key in self.auto_detail_widgets:
                container, table = self.auto_detail_widgets[auto_key]
            
            if not table: continue

            for row_idx in range(table.rowCount() - 1): 
                aliquota_widget = table.cellWidget(row_idx, 2)
                das_widget = table.cellWidget(row_idx, 5)
                dam_widget = table.cellWidget(row_idx, 6)
                
                if not aliquota_widget or not das_widget or not dam_widget: continue
                if row_idx >= len(auto_data.get('dados_anuais', [])): continue

                mes_data = auto_data['dados_anuais'][row_idx]
                mes_ano_str = mes_data['mes_ano']

                al_val_str = aliquota_widget.text(); al_val_num, ok_a = locale.toDouble(al_val_str)
                das_val_str = das_widget.text(); das_val_num, ok_d = locale.toDouble(das_val_str)
                dam_val_str = dam_widget.text(); dam_val_num, ok_m = locale.toDouble(dam_val_str)
                
                if not ok_a: al_val_num = 0.0
                if not ok_d: das_val_num = 0.0
                if not ok_m: dam_val_num = 0.0
                
                mes_data['aliquota_target_user'] = al_val_num
                wizard_auto_data['monthly_overrides'][mes_ano_str] = al_val_num

                new_aliquota_dec = al_val_num / 100.0
                new_monthly_iss_apurado = 0.0
                
                for inv_data in mes_data.get('_monthly_invoices_data', []):
                    valor = inv_data['valor']
                    declared = inv_data['declared_rate']
                    is_paid = inv_data['is_paid']
                    
                    inv_iss = 0.0
                    if is_split_diff_case:
                        inv_iss = max(0, (new_aliquota_dec - declared) * valor)
                    elif is_paid:
                        inv_iss = max(0, (new_aliquota_dec - declared) * valor)
                    else:
                        inv_iss = new_aliquota_dec * valor
                    new_monthly_iss_apurado += inv_iss
                
                mes_data['iss_apurado'] = new_monthly_iss_apurado

                period_key_str_pgdas = mes_data['mes_ano'] 
                try:
                    dt = datetime.strptime(mes_data['mes_ano'], "%m/%Y")
                    period_key_str_dam = f"{dt.month}/{dt.year}"
                except:
                    period_key_str_dam = mes_data['mes_ano']

                dams_list = self.wizard.available_credits_map['DAM'].get(period_key_str_dam, [])
                available_dam_total = sum(d['val'] for d in dams_list)
                
                available_pgdas = self.wizard.available_credits_map['PGDAS'].get(period_key_str_pgdas, 0.0)
                pgdas_decl_num = original_pgdas_map.get(period_key_str_pgdas, (0.0, "-"))[1]

                dam_utilizado = min(dam_val_num, available_dam_total)
                iss_apos_dam = mes_data['iss_apurado'] - dam_utilizado
                das_utilizado = min(das_val_num, max(0, min(iss_apos_dam, available_pgdas)))
                
                iss_apurado_op = max(0, iss_apos_dam - das_utilizado)

                remainder_to_deduct = dam_utilizado
                used_dam_codes = []
                for dam_obj in dams_list:
                    if remainder_to_deduct <= 0.0001: break
                    if dam_obj['val'] > 0:
                        deduct = min(dam_obj['val'], remainder_to_deduct)
                        dam_obj['val'] -= deduct
                        remainder_to_deduct -= deduct
                        used_dam_codes.append(dam_obj['code'])
                
                dam_ident_str = ", ".join(sorted(set(used_dam_codes))) if used_dam_codes else "-"

                mes_data['dam_iss_pago'] = dam_utilizado
                mes_data['das_iss_pago'] = das_utilizado
                mes_data['iss_apurado_op'] = iss_apurado_op
                mes_data['das_identificacao'] = pgdas_decl_num if das_utilizado > 0 else "-"
                mes_data['dam_identificacao'] = dam_ident_str

                self.wizard.available_credits_map['PGDAS'][period_key_str_pgdas] = available_pgdas - das_utilizado

            auto_data['totais']['iss_apurado'] = sum(m['iss_apurado'] for m in auto_data['dados_anuais'])
            auto_data['totais']['dam_iss_pago'] = sum(m['dam_iss_pago'] for m in auto_data['dados_anuais'])
            auto_data['totais']['das_iss_pago'] = sum(m['das_iss_pago'] for m in auto_data['dados_anuais'])
            auto_data['totais']['iss_apurado_op'] = sum(m['iss_apurado_op'] for m in auto_data['dados_anuais'])

            auto_data['monthly_overrides'] = wizard_auto_data['monthly_overrides']


# --- Page 4: Confirmation ---
class ConfirmationPage(QWidget):
    def __init__(self, wizard):
        super().__init__()
        self.wizard = wizard
        
        self.layout = QVBoxLayout(self)
        self.layout.setSpacing(10)

        self.subtitle_label = QLabel()
        self.subtitle_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.subtitle_label.setStyleSheet("background-color: white; color: black; font-style: italic; padding: 5px; border: 1px solid #ddd;")
        self.layout.addWidget(self.subtitle_label)
        
        header_layout = QHBoxLayout()
        label = QLabel("Confirmação Final dos Autos:")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        header_layout.addWidget(label)
        header_layout.addStretch()
        self.layout.addLayout(header_layout)

        self.autos_table = QTableWidget()
        self.autos_table.setColumnCount(4)
        self.autos_table.setHorizontalHeaderLabels([
            "Gerar?", "Auto", "Motivo", "Valor Final (Atualizado)"
        ])
        
        self.autos_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.autos_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.autos_table.verticalHeader().setVisible(False)
        self.autos_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.layout.addWidget(self.autos_table)
        
        button_layout = QHBoxLayout()
        select_all_btn = QPushButton("Selecionar Todos")
        select_all_btn.clicked.connect(self.select_all)
        deselect_all_btn = QPushButton("Desselecionar Todos")
        deselect_all_btn.clicked.connect(self.deselect_all)
        
        button_layout.addStretch()
        button_layout.addWidget(select_all_btn)
        button_layout.addWidget(deselect_all_btn)
        self.layout.addLayout(button_layout)

        self.layout.addStretch()
        self.final_data = {}

    def initializePage(self):
        self.populate_confirmation()

    def populate_confirmation(self):
        self.autos_table.setRowCount(0)
        self.final_data = {
            auto['numero']: auto for auto in self.wizard.preview_context.get('autos', [])
        }
        
        if not self.final_data: return

        self.autos_table.setRowCount(len(self.final_data))
        row = 0
        for auto_key, auto_info in self.final_data.items():
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget); chk_layout.setContentsMargins(0,0,0,0); chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            checkbox = QCheckBox(); checkbox.setChecked(True); checkbox.setProperty("auto_key", auto_key)
            chk_layout.addWidget(checkbox)
            self.autos_table.setCellWidget(row, 0, chk_widget)
            
            self.autos_table.setItem(row, 1, QTableWidgetItem(auto_key))
            
            motive_text = auto_info.get('motive_text', auto_key)
            self.autos_table.setItem(row, 2, QTableWidgetItem(motive_text))
            
            # ✅ FIX: Handle NoneType for user_defined_credito
            user_val = auto_info.get('user_defined_credito')
            if user_val is None:
                # If manual override is empty, fallback to calculated value
                user_val = auto_info.get('totais', {}).get('iss_apurado_op', 0.0)
                
            try:
                final_val = float(user_val)
            except (ValueError, TypeError):
                final_val = 0.0
                
            fmt_calc = f"R$ {self.wizard._fmtd(final_val)}"
            item_calc = QTableWidgetItem(fmt_calc)
            item_calc.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.autos_table.setItem(row, 3, item_calc)
            
            row += 1

    # --- VALIDATION LOGIC ---

    def start_validation_process(self):
        tasks = []
        year_totals = {} # Stores { "2021": 2500.00, "2022": 500.00 }
        
        # 1. Calculate Expected Totals per Year
        # We look at all autos generated in the wizard
        for auto_data in self.wizard.preview_context.get('autos', []):
            
            # A. Get Value for this Auto
            # Use manual credit if defined, otherwise calculated total
            val = auto_data.get('totais', {}).get('iss_apurado_op', 0.0)
            if auto_data.get('user_defined_credito') is not None:
                val = float(auto_data.get('user_defined_credito'))
            
            # B. Identify Year
            # Check the months inside this auto
            years_found = set()
            for m in auto_data.get('dados_anuais', []):
                try:
                    y = m['mes_ano'].split('/')[1]
                    years_found.add(y)
                except: pass
            
            # Assign value to the year (Assumption: Auto doesn't span years)
            if years_found:
                # Use the first year found (Standard practice for autos)
                target_year = list(years_found)[0]
                year_totals[target_year] = year_totals.get(target_year, 0.0) + val

        if not year_totals:
            QMessageBox.warning(self, "Erro", "Não foi possível calcular os totais por ano.")
            return

        # 2. Build Tasks with Expected Value
        for year, expected_val in year_totals.items():
            tasks.append({
                'imu': self.wizard.company_imu,
                'year': year,
                'expected_value': expected_val # <--- PASSING THE TARGET
            })

        # 3. Start Worker (UI Logic)
        self.validate_btn.setEnabled(False)
        self.validate_btn.setText("⏳ Processando...")

        self.verification_thread = QThread()
        self.verification_worker = ValidationExtractorWorker(tasks) 
        self.verification_worker.moveToThread(self.verification_thread)
        
        self.verification_worker.progress.connect(lambda msg: self.validate_btn.setText(f"⏳ {msg}"))
        self.verification_thread.started.connect(self.verification_worker.run)
        self.verification_worker.finished.connect(self.on_validation_finished)
        self.verification_worker.error.connect(self.on_validation_error)
        
        self.verification_worker.finished.connect(self.verification_thread.quit)
        self.verification_worker.finished.connect(self.verification_worker.deleteLater)
        self.verification_thread.finished.connect(self.verification_thread.deleteLater)
        
        self.verification_thread.start()

    def on_validation_finished(self, results):
        """
        Called when the robot finishes.
        results: dict { "2021": { "extracted_original": 2500.00, ... }, ... }
        """
        self.validate_btn.setEnabled(True)
        self.validate_btn.setText("🔍 Validar Valores no Sistema (OCR)")
        
        # Iterate rows and compare values
        for row in range(self.autos_table.rowCount()):
            # Get Auto ID
            auto_key_item = self.autos_table.item(row, 1)
            if not auto_key_item: continue
            auto_key = auto_key_item.text()
            
            # Find which year this auto belongs to
            target_year = None
            auto_data = self.final_data.get(auto_key)
            if auto_data:
                # We assume an auto belongs to the year found in its first monthly record
                for m in auto_data.get('dados_anuais', []):
                    try:
                        target_year = m['mes_ano'].split('/')[1]
                        break
                    except: pass
            
            # Perform Comparison if we have data
            if target_year and str(target_year) in results:
                raw_data = results[str(target_year)]
                
                # --- FIX START: Handle Dictionary vs Float ---
                if isinstance(raw_data, dict):
                    # Extract the numeric value from the dictionary
                    system_val = raw_data.get("extracted_original", 0.0)
                else:
                    # Fallback if it somehow returns a direct float
                    system_val = float(raw_data)
                # --- FIX END ---

                calc_val = self.autos_table.item(row, 3).data(Qt.ItemDataRole.UserRole)
                
                # Update System Value Column
                self.autos_table.setItem(row, 4, QTableWidgetItem(f"R$ {self.wizard._fmtd(system_val)}"))
                
                # Logic: Is the difference acceptable? (0.1% tolerance)
                diff = abs(calc_val - system_val)
                tolerance = max(calc_val, system_val) * 0.001 
                
                status_item = QTableWidgetItem()
                status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                if diff <= tolerance:
                    status_item.setText("✅ Validado")
                    status_item.setBackground(QColor(200, 255, 200)) # Green
                    status_item.setToolTip(f"Diferença: R$ {diff:.2f} (Dentro da tolerância)")
                else:
                    status_item.setText("⚠️ Divergente")
                    status_item.setBackground(QColor(255, 200, 200)) # Red
                    status_item.setToolTip(f"Diferença: R$ {diff:.2f} (Maior que 0.1%)")
                
                self.autos_table.setItem(row, 5, status_item)
            else:
                self.autos_table.setItem(row, 4, QTableWidgetItem("-"))
                self.autos_table.setItem(row, 5, QTableWidgetItem("Ano n/a"))

        QMessageBox.information(self, "Validação Concluída", "Comparação finalizada. Verifique os status na tabela.")

    def on_validation_error(self, err_msg):
        self.validate_btn.setEnabled(True)
        self.validate_btn.setText("🔍 Validar Valores no Sistema (OCR)")
        QMessageBox.critical(self, "Erro na Validação", err_msg)

    # --- HELPER METHODS ---

    def get_checkbox_for_row(self, row):
        """Helper to find the QCheckBox widget inside the cell layout."""
        widget = self.autos_table.cellWidget(row, 0)
        if widget:
            return widget.findChild(QCheckBox)
        return None

    def select_all(self):
        for row in range(self.autos_table.rowCount()):
            cb = self.get_checkbox_for_row(row)
            if cb: cb.setChecked(True)

    def deselect_all(self):
        for row in range(self.autos_table.rowCount()):
            cb = self.get_checkbox_for_row(row)
            if cb: cb.setChecked(False)

    def get_selected_data(self):
        """
        Returns the final context filtered by the checkboxes in the table.
        Called by the wizard when clicking 'Gerar Documentos'.
        """
        selected_keys = set()
        for row in range(self.autos_table.rowCount()):
            cb = self.get_checkbox_for_row(row)
            if cb and cb.isChecked():
                # We stored auto_key in property
                auto_key = cb.property("auto_key")
                selected_keys.add(auto_key)

        # Filter the original context to include only selected autos
        final_context = self.wizard.preview_context.copy()
        final_context['autos'] = [
            auto for auto in final_context.get('autos', [])
            if auto['numero'] in selected_keys
        ]
        
        return final_context
