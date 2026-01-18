# --- FILE: app/main_window.py ---
import pandas as pd
import os
import shutil
import traceback
import re
import gc
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QComboBox, QFileDialog, QLineEdit,
                               QTextEdit, QLabel, QStyle, QGroupBox, QFormLayout,
                               QMessageBox, QMenu, QCheckBox, QDialog, QPlainTextEdit, 
                               QApplication,QDialogButtonBox)
from PySide6.QtCore import QThread, QSettings, QUrl, QTimer, Qt
from PySide6.QtGui import QAction, QIcon, QDesktopServices, QFont
import json 
from datetime import datetime 
from main import load_activity_data
# ✅ Import SESSION_FILE_PREFIX to handle session cleanup
from .constants import Columns, APP_NAME, SESSION_FILE_PREFIX, APP_VERSION
# ✅ Importar os novos workers (incluindo SituacaoExtractorWorker e AutomaticIDDWorker)
from .workers import (AIPrepWorker, RulesPrepWorker, AIAnalysisWorker, GenerationWorker,
                      SimplesReaderWorker, SimplesDownloaderWorker, DatabaseExtractorWorker,
                      SituacaoExtractorWorker, UpdateCheckerWorker, UpdateDownloaderWorker,
                      AutomaticIDDWorker, DeckerWorker, AnalysisScannerWorker, MultiYearPrepWorker,
                      NewsFetcherWorker) # <--- ADD DeckerWorker
from .relabeling_dialog import RelabelingWindow
from .review_wizard import ReviewWizard 
from .settings_dialog import SettingsDialog
from .text_editor_dialog import TextEditorDialog 
from app.config import (get_custom_general_texts, set_custom_general_texts,
                        get_custom_auto_texts, set_custom_auto_texts,
                        DEFAULT_GENERAL_TEXTS, DEFAULT_AUTO_TEXTS, NEWS_SOURCE_URL)

import logging
from .generation_summary_dialog import GenerationSummaryDialog
# ✅ Importar o novo diálogo de Ferramentas
from app.ferramentas.qt_dialogs import GetFolderAndYearsDialog, CNPJSelectionDialog # Importado CNPJSelectionDialog
from app.workers import UpdateCheckerWorker
from app.updater import Updater
from app.constants import APP_VERSION
from app.activity_review_dialog import ActivityReviewDialog
from .duplicate_review_dialog import DuplicateReviewDialog # ✅ Import New Dialog


class AuditApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        #(100, 100, 1400, 900)
        self.resize(1200, 700)
        self.setWindowIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))

        self.statusBar().showMessage("Pronto. Por favor, selecione os ficheiros mestre e de notas.")

        self.thread = None
        self.worker = None
        
        # ✅ NEW: separate clean copy from working copy
        self.clean_invoices_df = None 
        self.company_invoices_df = None
        
        self.infraction_groups = {}
        self.activity_list = []
        
        self.df_empresas = pd.DataFrame(columns=[
            Columns.CNPJ, Columns.RAZAO_SOCIAL, Columns.IMU, 
            Columns.ENDERECO, Columns.CEP, Columns.EPAF_NUMERO
        ])
        self.activity_data = {}
        
        self.workflow_state = {
            'load': 'pending', 
            'rules': 'pending',
            'ai': 'pending',
            'relabel': 'pending',
            'review': 'pending'
        }
        self._setup_icons()

        self._setup_menu_bar() 
        self._setup_ui()
        self.load_auditor_info()
        self.company_combo.currentIndexChanged.connect(self.on_company_selection_change)
        self.update_button_states() 
        self.save_cadastro_button.setEnabled(False)
        QTimer.singleShot(1000, self.check_updates_silently)
        QTimer.singleShot(1500, self.check_news_silently) 

    def check_news_silently(self):
        self.news_thread = QThread()
        self.news_worker = NewsFetcherWorker(NEWS_SOURCE_URL)
        self.news_worker.moveToThread(self.news_thread)
        self.news_worker.finished.connect(self.on_news_fetched)
        self.news_thread.started.connect(self.news_worker.run)
        self.news_thread.finished.connect(self.news_thread.deleteLater)
        self.news_thread.start()

    def on_news_fetched(self, news_text):
        self.news_thread.quit()
        if news_text and len(news_text.strip()) > 0:
            # Show the dialog
            dialog = NewsDialog(news_text, self)
            dialog.exec()

    def open_script_viewer(self):
        dialog = ScriptViewerDialog(self)
        dialog.exec()

    def create_master_template(self):
        """Generates an empty Excel template for the Master File."""
        columns = [
            "razao_social", "imu", "cnpj", "CNPJ_n", "Períodos de Opção", 
            "Situação Cadastral", "Reg Trib Diferenciado em 2020?", 
            "AI/IDD em 2020?", "Há Revisão em Aberto?", "pgdas", 
            "DAM'S Normais/Avulsos PAGOS", "Nfe", "e-PAF", "SUP", "DEC", 
            "CIÊNCIA", "Inserir Data da Ciência GTM", 
            "Para ffrm69 com Termo de Ciênica", "Data Limite", 
            "se NÃO IMPUGNOU tramitar para ffrm69:"
        ]
        
        default_filename = "modelo_ficheiro_mestre.xlsx"
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Salvar Modelo Mestre", default_filename, "Excel Files (*.xlsx)"
        )
        
        if not filepath:
            return

        try:
            df = pd.DataFrame(columns=columns)
            # Saving to sheet 'Empresas' as that is the default expected by load_companies
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Empresas', index=False)
            
            self.log_text_edit.append(f"✅ Modelo criado com sucesso em: {filepath}")
            QMessageBox.information(self, "Sucesso", f"Modelo criado em:\n{filepath}")
        except Exception as e:
            self.log_text_edit.append(f"❌ Erro ao criar modelo: {e}")
            QMessageBox.critical(self, "Erro", f"Falha ao criar ficheiro:\n{e}")
    
    def create_company_folders_tool(self):
        if self.df_empresas is None or self.df_empresas.empty:
            QMessageBox.warning(self, "Dados em Falta", "Por favor, carregue o 'Ficheiro Mestre'.")
            return
        output_dir = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", "", QFileDialog.Option.ShowDirsOnly)
        if not output_dir: return
        created_count = 0; existing_count = 0; errors = 0
        self.statusBar().showMessage("A criar pastas...")
        col_razao = Columns.RAZAO_SOCIAL; col_imu = Columns.IMU
        for index, row in self.df_empresas.iterrows():
            company_name = str(row.get(col_razao, '')); raw_imu = str(row.get(col_imu, ''))
            if not company_name.strip() or not raw_imu.strip(): continue
            try:
                imu_clean = re.sub(r'\D', '', raw_imu.split('.')[0])
                if not imu_clean: continue
                cleaned_name = re.sub(r'[^\w\s-]', '', company_name.strip())
                cleaned_name = re.sub(r'[\s-]+', '_', cleaned_name).upper()
                folder_name = f"{imu_clean}_{cleaned_name}"
                folder_path = os.path.join(output_dir, folder_name)
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path); created_count += 1
                else: existing_count += 1
            except Exception as e: errors += 1
        QMessageBox.information(self, "Concluído", f"Criadas: {created_count}\nExistentes: {existing_count}\nErros: {errors}")

    def _load_last_paths(self):
        settings = QSettings("MyAuditApp", "AuditApp")
        last_master = settings.value("paths/last_master_file", "")
        if os.path.exists(last_master):
            self.master_file_path_edit.setText(last_master)
            self.load_companies(last_master)
        self.check_if_ready_to_load()

    def _save_last_paths(self):
        settings = QSettings("MyAuditApp", "AuditApp")
        settings.setValue("paths/last_master_file", self.master_file_path_edit.text())
        settings.setValue("paths/last_invoices_file", self.invoices_file_path_edit.text())

    def _setup_icons(self):
        style = self.style()
        self.icons = {
            'pending': style.standardIcon(QStyle.StandardPixmap.SP_MediaSkipForward),
            'running': style.standardIcon(QStyle.StandardPixmap.SP_MediaPlay), 
            'completed': style.standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton),
            'stale': style.standardIcon(QStyle.StandardPixmap.SP_BrowserReload)
        }

    def _setup_menu_bar(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("&Ficheiro")
        
        # --- NOVO: Ação de Reinicialização ---
        reset_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogResetButton), "Reiniciar Aplicação (Reset)", self)
        reset_action.triggered.connect(self.reset_all_state)
        file_menu.addAction(reset_action)
        file_menu.addSeparator()
        # -------------------------------------
        
        settings_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView), "Preferências...", self)
        settings_action.triggered.connect(self.open_settings_dialog)
        file_menu.addAction(settings_action)
        file_menu.addSeparator()
        import_texts_action = QAction(QIcon.fromTheme("document-open"), "Importar Textos Personalizados...", self)
        import_texts_action.triggered.connect(self.import_custom_texts)
        file_menu.addAction(import_texts_action)
        export_texts_action = QAction(QIcon.fromTheme("document-save-as"), "Exportar Textos Personalizados...", self)
        export_texts_action.triggered.connect(self.export_custom_texts)
        file_menu.addAction(export_texts_action)
        log_action = QAction(QIcon.fromTheme("document-properties", self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)), "Ver Histórico", self)
        log_action.triggered.connect(self.open_activity_log)
        file_menu.addAction(log_action)
        file_menu.addSeparator()
        exit_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCloseButton), "Sair", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        edit_menu = menu_bar.addMenu("&Editar")
        edit_texts_action = QAction(QIcon.fromTheme("document-edit"), "Editar Textos...", self)
        edit_texts_action.triggered.connect(self.open_text_editor)
        edit_menu.addAction(edit_texts_action)

        tools_menu = menu_bar.addMenu("&Ferramentas")
        view_script_action = QAction(QIcon.fromTheme("edit-copy"), "Ver Script de Extração", self)
        view_script_action.triggered.connect(self.open_script_viewer)
        tools_menu.addAction(view_script_action)
        tools_menu.addAction(QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon),"Criar Modelo Mestre (Excel)",self,triggered=self.create_master_template))
        create_folders_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon), "Gerar Pastas", self)
        create_folders_action.triggered.connect(self.create_company_folders_tool)
        tools_menu.addAction(create_folders_action)
        
        # --- ✅ NOVAS FERRAMENTAS INTEGRADAS ---
        tools_menu.addSeparator()

        # 1. Ferramenta de Leitura de PDF (OCR) do Simples
        #read_simples_action = QAction(QIcon.fromTheme("document-open"), "Ler PDFs do Simples (OCR)", self)
        #read_simples_action.triggered.connect(self.open_simples_reader_tool)
        #tools_menu.addAction(read_simples_action)
        
        # 2. Ferramenta de Download (RPA) do Simples
        download_simples_action = QAction(QIcon.fromTheme("system-run"), "Baixar PDFs do Simples (RPA)", self)
        download_simples_action.triggered.connect(self.start_simples_downloader_tool)
        tools_menu.addAction(download_simples_action)

        # 3. Ferramenta de Extração do Banco de Dados (DAMS/NFSE)
        db_extractor_action = QAction(QIcon.fromTheme("utilities-data-base"), "Extrair Relatórios (DAMS/NFSE)", self)
        db_extractor_action.triggered.connect(self.open_db_extractor_tool)
        tools_menu.addAction(db_extractor_action)

        # 4. Ferramenta de Extração de Situação (pywinauto)
        situacao_extractor_action = QAction(QIcon.fromTheme("system-run"), "Extrair Situação do Sistema (RPA)", self)
        situacao_extractor_action.triggered.connect(self.start_situacao_extractor_tool)
        tools_menu.addAction(situacao_extractor_action)

        #Decker
        decker_action = QAction(QIcon.fromTheme("mail-message-new"), "Enviar Emails IDD (Decker)", self)
        decker_action.triggered.connect(self.open_decker_tool)
        tools_menu.addAction(decker_action)
        # --- NEW BATCH IDD TOOL ---
        tools_menu.addSeparator()
        # [NEW ACTION]
        batch_auto_idd_action = QAction(QIcon.fromTheme("system-run"), "Batch: IDD Automático (Diretórios)", self)
        batch_auto_idd_action.triggered.connect(self.open_batch_idd_tool_folders)
        tools_menu.addAction(batch_auto_idd_action)

        tools_menu.addSeparator()
        
        # [NEW ACTION]
        analysis_scanner_action = QAction(QIcon.fromTheme("system-search"), "Escanear Potencial (Analítico por Ano)", self)
        analysis_scanner_action.triggered.connect(self.open_analysis_scanner_tool)
        tools_menu.addAction(analysis_scanner_action)
        # --- FIM DAS NOVAS FERRAMENTAS ---

        help_menu = self.menuBar().addMenu("&Ajuda")
        
        check_updates_action = QAction("Verificar Atualizações", self)
        check_updates_action.triggered.connect(self.manual_update_check)
        help_menu.addAction(check_updates_action)
        
        # Add version label to menu or status bar
        help_menu.addAction(QAction(f"Versão: {APP_VERSION}", self, enabled=False))

        
    def open_analysis_scanner_tool(self):
        """
        Tool to scan a root directory for 'Relatorio_NFSE_{IMU}_{YEAR}.xlsx' files
        and generate a consolidated potential report.
        """
        # 1. Select Root Folder
        root_dir = QFileDialog.getExistingDirectory(self, "Selecione a Pasta Raiz (onde estão as pastas das empresas)", "", QFileDialog.Option.ShowDirsOnly)
        if not root_dir: return

        # 2. Confirm
        msg = QMessageBox(self)
        msg.setWindowTitle("Iniciar Scanner Analítico")
        msg.setText("O sistema irá procurar por arquivos <b>Relatorio_NFSE_*.xlsx</b> em todas as subpastas.<br><br>"
                    "Será realizada uma análise completa (Regras + Cálculos) para cada ano encontrado.<br>"
                    "Isso pode levar algum tempo.")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if msg.exec() != QMessageBox.StandardButton.Yes: return

        # 3. Start Worker
        self.log_text_edit.clear()
        self.log_text_edit.append(f"--- Iniciando Scanner Analítico ---")
        self.log_text_edit.append(f"Pasta Raiz: {root_dir}")
        self.statusBar().showMessage("Escaneando e analisando...")
        
        self._start_worker_thread(
            AnalysisScannerWorker,
            'tool_scanner',
            root_dir,
            on_finished_slot=self.on_analysis_scanner_finished
        )

    def on_analysis_scanner_finished(self, output_path):
        self.workflow_state['tool_scanner'] = 'completed'
        self.log_text_edit.append(f"✅ Análise concluída.")
        self.log_text_edit.append(f"📂 Relatório salvo em: {output_path}")
        
        QMessageBox.information(self, "Concluído", f"Relatório Analítico Gerado com Sucesso:\n\n{output_path}")
        
        # Open file automatically
        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(output_path))
        except:
            pass
            
        self.update_button_states()

    def check_updates_silently(self):
        self.start_update_worker(silent=True)

    def open_decker_tool(self):
        """
        Tool to automate email sending using DrissionPage (Decker).
        """
        # 1. Validation
        if self.df_empresas is None or self.df_empresas.empty:
            QMessageBox.warning(self, "Aviso", "Carregue o Ficheiro Mestre primeiro para ter a lista de CNPJs.")
            return

        # 2. Select Root Folder (Where the IDD folders are)
        base_dir = QFileDialog.getExistingDirectory(self, "Selecione a Pasta Raiz (com subpastas IDD_XX_25)", "", QFileDialog.Option.ShowDirsOnly)
        if not base_dir: return

        # 3. Credentials Dialog (Simple Input)
        # Using a quick QDialog to get User/Pass
        creds_dialog = QDialog(self)
        creds_dialog.setWindowTitle("Credenciais ADM-DEC")
        l = QFormLayout(creds_dialog)
        user_input = QLineEdit("") # Default from script
        pass_input = QLineEdit()
        pass_input.setEchoMode(QLineEdit.EchoMode.Password)
        pass_input.setText("") # Default from script
        l.addRow("Usuário:", user_input)
        l.addRow("Senha:", pass_input)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(creds_dialog.accept)
        bb.rejected.connect(creds_dialog.reject)
        l.addWidget(bb)
        
        if creds_dialog.exec() != QDialog.Accepted:
            return
            
        credentials = {'user': user_input.text(), 'pass': pass_input.text()}

        # 4. Scan Folders & Prepare Selection List
        self.statusBar().showMessage("Escaneando pastas para envio...")
        from pathlib import Path
        root_path = Path(base_dir)
        
        candidates = []
        
        try:
            for folder in root_path.iterdir():
                if folder.is_dir():
                    # Parse IMU from folder name (12345_Name)
                    folder_name = folder.name
                    parts = folder_name.split('_', 1)
                    possible_imu = re.sub(r'\D', '', parts[0])
                    
                    if possible_imu:
                        # Find CNPJ in Master DF
                        master_clean_imus = self.df_empresas[Columns.IMU].astype(str).str.replace(r'\D', '', regex=True)
                        matches = self.df_empresas[master_clean_imus == possible_imu]
                        
                        if not matches.empty:
                            row = matches.iloc[0]
                            # Check if PDF exists
                            pdf_exists = len(list(folder.glob("*_Comunicado.pdf"))) > 0
                            
                            candidates.append({
                                'imu': possible_imu,
                                'cnpj': str(row[Columns.CNPJ]),
                                'name': str(row[Columns.RAZAO_SOCIAL]),
                                'dir_path': str(folder), # Used by selection dialog logic
                                'is_selected': pdf_exists # Auto-select if file exists
                            })
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler diretórios: {e}")
            return

        if not candidates:
            QMessageBox.warning(self, "Sem Correspondências", "Nenhuma pasta correspondente encontrada.")
            return

        # 5. Selection Dialog
        # We reuse CNPJSelectionDialog. It expects list of dicts.
        # It returns list of dicts: {'cnpj': ID, 'dir_path': ...}
        
        # Mapping for dialog (it uses 'cnpj' key as display ID usually, we pass IMU there for clarity or CNPJ)
        dialog_list = []
        for c in candidates:
            dialog_list.append({
                'cnpj': c['cnpj'], 
                'name': c['name'],
                'dir_path': c['dir_path'],
                'is_selected': c['is_selected']
            })

        from app.ferramentas.qt_dialogs import CNPJSelectionDialog
        sel_dialog = CNPJSelectionDialog(dialog_list, self)
        sel_dialog.setWindowTitle(f"Selecione para Envio ({len(candidates)} pastas)")
        
        if sel_dialog.exec() != QDialog.Accepted: return
        
        selected_raw = sel_dialog.get_selected_cnpjs() # [{'cnpj': '...', 'dir_path': '...'}]
        
        # Remap to full task objects required by worker
        final_tasks = []
        for sel in selected_raw:
            # Find original IMU based on path or CNPJ
            orig = next((x for x in candidates if x['cnpj'] == sel['cnpj']), None)
            if orig:
                final_tasks.append(orig)

        if not final_tasks: return

        # 6. Confirm & Run
        msg = QMessageBox(self)
        msg.setWindowTitle("Iniciar Envio de Emails")
        msg.setText(f"Enviar emails para <b>{len(final_tasks)}</b> empresas?<br><br>"
                    "⚠️ O navegador será controlado automaticamente.<br>"
                    "⚠️ Não use o mouse/teclado.")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if msg.exec() != QMessageBox.StandardButton.Yes: return

        self.log_text_edit.clear()
        self.log_text_edit.append(f"--- Iniciando Decker (Envio de Emails) ---")
        
        self._start_worker_thread(
            DeckerWorker,
            'tool_decker',
            final_tasks,
            base_dir,
            credentials,
            on_finished_slot=self.on_decker_finished
        )

    def on_decker_finished(self):
        self.workflow_state['tool_decker'] = 'completed'
        self.log_text_edit.append("✅ Decker finalizado.")
        QMessageBox.information(self, "Concluído", "Processo de envio de emails finalizado.")
        self.update_button_states()

    def open_batch_idd_tool_folders(self):
        """
        1. User selects Root Folder.
        2. App scans subfolders matching IMU/CNPJ against loaded Master File.
        3. App looks for Excel files inside those folders.
        4. User selects companies.
        5. Worker runs IDD logic and saves inside each company folder.
        """
        # 1. Validation
        if self.df_empresas is None or self.df_empresas.empty:
            QMessageBox.warning(self, "Aviso", "Carregue o Ficheiro Mestre primeiro (Ficheiro > Abrir Mestre).")
            return

        # 2. Select Root Folder
        root_dir = QFileDialog.getExistingDirectory(self, "Selecione a Pasta Raiz (com subpastas das empresas)", "", QFileDialog.Option.ShowDirsOnly)
        if not root_dir: return

        self.statusBar().showMessage("Escaneando pastas e notas...")
        candidates = []
        
        from pathlib import Path
        root_path = Path(root_dir)
        
        # 3. Scan Directories
        try:
            for folder in root_path.iterdir():
                if folder.is_dir():
                    # Attempt to identify company by folder name (e.g. "12345_EmpresaX")
                    folder_name = folder.name
                    parts = folder_name.split('_', 1)
                    possible_imu = re.sub(r'\D', '', parts[0]) # Clean numeric part
                    
                    matched_row = None
                    
                    # Try finding in Master File
                    if possible_imu:
                        # Match by IMU (ignoring punctuation in master)
                        master_imus = self.df_empresas[Columns.IMU].astype(str).str.replace(r'\D', '', regex=True)
                        matches = self.df_empresas[master_imus == possible_imu]
                        if not matches.empty:
                            matched_row = matches.iloc[0]
                    
                    # If found in master, look for Invoice Excel inside
                    if matched_row is not None:
                        inv_path = None
                        dam_path = None
                        
                        # ✅ FIX: Scan ALL files to find Invoices (xls) and DAMs (csv/xls)
                        for f in folder.iterdir():
                            fname = f.name.lower()
                            if "~$" in fname: continue # Ignore temp files
                            if "controle_de_atividades" in fname: continue
                            
                            # Check for DAMs (Prioritize CSV, but check name)
                            if "dam" in fname and "relatorio" in fname:
                                dam_path = str(f)
                            
                            # Check for Invoices (Excel)
                            elif fname.endswith(('.xls', '.xlsx')):
                                inv_path = str(f)

                        # Only add if we found invoices (DAM is optional but highly recommended for reduction)
                        if inv_path:
                            candidates.append({
                                'imu': str(matched_row[Columns.IMU]),
                                'cnpj': str(matched_row[Columns.CNPJ]),
                                'name': str(matched_row[Columns.RAZAO_SOCIAL]),
                                'folder_path': str(folder),
                                'invoices_path': inv_path,
                                'dam_path': dam_path, # ✅ Now correctly populated
                                'display_text': f"{folder_name} {'(Com DAMs)' if dam_path else ''}"
                            })
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler diretórios: {e}")
            return

        if not candidates:
            QMessageBox.warning(self, "Sem Correspondências", "Nenhuma pasta correspondente ao Mestre contendo ficheiros Excel foi encontrada.")
            return

        # 4. Selection Dialog
        # Convert to format expected by CNPJSelectionDialog
        dialog_list = []
        for c in candidates:
            dialog_list.append({
                'cnpj': c['imu'], # Use IMU as ID
                'name': c['name'],
                'dir_path': c['folder_path'] # Pass folder path
            })
            
        from app.ferramentas.qt_dialogs import CNPJSelectionDialog
        dialog = CNPJSelectionDialog(dialog_list, self)
        dialog.setWindowTitle(f"Selecione as Empresas ({len(candidates)} encontradas)")
        
        if dialog.exec() != QDialog.Accepted: return
        
        selected_raw = dialog.get_selected_cnpjs() # Returns [{'cnpj': '...', 'dir_path': '...'}]
        
        # Map back to full task objects
        final_tasks = []
        for sel in selected_raw:
            # Find the original candidate data
            for c in candidates:
                if c['folder_path'] == sel['dir_path']:
                    final_tasks.append(c)
                    break
        
        if not final_tasks: return

        # 5. Confirm & Run
        msg = QMessageBox(self)
        msg.setWindowTitle("Iniciar Batch IDD")
        msg.setText(f"Processar <b>{len(final_tasks)}</b> empresas?<br>"
                    "Os arquivos (PDF do IDD, Relatório Fiscal) serão salvos dentro de cada pasta.<br><br>"
                    "⚠️ O mouse será controlado pelo robô.")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if msg.exec() != QMessageBox.StandardButton.Yes: return

        self.log_text_edit.clear()
        self.log_text_edit.append(f"--- Iniciando Batch IDD ({len(final_tasks)} tarefas) ---")
        
        from app.workers import BatchIDDWorker
        
        auditor_data = {'nome': self.auditor_name_edit.text(), 'matricula': self.auditor_matricula_edit.text()}
        
        self._start_worker_thread(
            BatchIDDWorker,
            'tool_batch_idd_folders',
            final_tasks,
            self.master_file_path_edit.text(),
            auditor_data,
            on_finished_slot=self.on_batch_idd_folders_finished
        )

    def on_batch_idd_folders_finished(self, results):
        self.workflow_state['tool_batch_idd_folders'] = 'completed'
        
        success = sum(1 for v in results.values() if v == "Sucesso")
        report = "\n📊 Resumo Batch:\n"
        for k, v in results.items():
            report += f"{k}: {v}\n"
            
        self.log_text_edit.append(report)
        QMessageBox.information(self, "Concluído", f"Processo finalizado.\nSucessos: {success}/{len(results)}")
        self.update_button_states()

    def on_update_check_result(self, available, url, version_tag):
        # We retrieve the 'silent' flag from the worker object itself if needed, 
        # or simply assume silent=False for manual checks. 
        # For simplicity, let's handle the UI logic here.
        
        is_silent_check = getattr(self, '_is_silent_check', False)
        
        if available:
            reply = QMessageBox.question(
                self, 
                "Atualização Disponível", 
                f"Nova versão <b>{version_tag}</b> disponível.<br>Deseja instalar?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.start_download_thread(url)
        else:
            if not is_silent_check:
                QMessageBox.information(self, "Tudo em dia", "Você já está na versão mais recente.")
                self.statusBar().showMessage("Sistema atualizado.")

    def manual_update_check(self):
        self.statusBar().showMessage("Verificando atualizações...")
        self.start_update_check_thread()

    def start_update_check_thread(self):
        self.update_thread = QThread()
        self.update_worker = UpdateCheckerWorker()
        self.update_worker.moveToThread(self.update_thread)
        
        # ⚠️ CRITICAL FIX: Add 'Qt.QueuedConnection' here ⚠️
        self.update_worker.finished.connect(self.on_update_check_finished, Qt.QueuedConnection)
        self.update_worker.error.connect(self.on_update_check_error, Qt.QueuedConnection)
        
        self.update_thread.started.connect(self.update_worker.run)
        self.update_thread.finished.connect(self.update_thread.deleteLater)
        self.update_thread.start()

    def on_update_check_error(self, error_msg):
        """Safe slot for errors"""
        self.update_thread.quit()
        QMessageBox.warning(self, "Erro", f"Erro ao verificar: {error_msg}")

    def on_update_check_finished(self, available, url, version_tag):
        """Safe slot for finished check"""
        self.update_thread.quit()
        
        if available:
            reply = QMessageBox.question(
                self, 
                "Atualização Disponível", 
                f"Nova versão <b>{version_tag}</b> disponível.<br><br>"
                "Deseja baixar e instalar agora?<br>"
                "<i>O aplicativo será fechado e reiniciado automaticamente.</i>",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.start_download_thread(url)
        else:
            QMessageBox.information(self, "Tudo em dia", "Você já está na versão mais recente.")
            self.statusBar().showMessage("Sistema atualizado.")

    def start_download_thread(self, url):
        self.dl_thread = QThread()
        self.dl_worker = UpdateDownloaderWorker(url)
        self.dl_worker.moveToThread(self.dl_thread)
        
        # ✅ USE Qt.QueuedConnection - This is vital for thread safety
        self.dl_worker.progress.connect(self.on_download_progress, Qt.QueuedConnection)
        self.dl_worker.error.connect(self.on_download_error, Qt.QueuedConnection)
        
        # ✅ Connect the new reboot signal
        self.dl_worker.reboot_requested.connect(self.finalize_update_and_restart, Qt.QueuedConnection)
        
        self.dl_thread.started.connect(self.dl_worker.run)
        self.dl_thread.finished.connect(self.dl_thread.deleteLater)
        self.dl_thread.start()

    def finalize_update_and_restart(self):
        """
        Called strictly by the Main Thread when the update is ready.
        """
        # 1. Stop the thread gracefully
        if hasattr(self, 'dl_thread') and self.dl_thread.isRunning():
            self.dl_thread.quit()
            self.dl_thread.wait(1000) # Wait up to 1s for cleanup

        self.log_text_edit.append("🔄 Fechando para atualização...")
        
        # 2. Force Qt to process any pending paint events (Fixes BackingStore warning)
        QApplication.processEvents()
        
        # 3. Clean Quit (Releases file locks for the batch script)
        QApplication.quit()

    def on_download_progress(self, msg):
        self.log_text_edit.append(msg)
        self.statusBar().showMessage(msg)

    def on_download_error(self, e):
        QMessageBox.critical(self, "Erro de Atualização", str(e))

    def start_update_worker(self, silent=False):
        # 1. Store the silent flag on the instance so the slot can read it
        self._is_silent_check = silent

        self.update_thread = QThread()
        self.update_worker = UpdateCheckerWorker()
        self.update_worker.moveToThread(self.update_thread)
        
        # 2. Connect using Qt.QueuedConnection (Fixes QObject::setParent / QBackingStore)
        self.update_worker.finished.connect(self.on_update_check_result, Qt.QueuedConnection)
        self.update_worker.error.connect(self.on_update_check_error, Qt.QueuedConnection)
        
        self.update_thread.started.connect(self.update_worker.run)
        self.update_thread.finished.connect(self.update_thread.deleteLater)
        self.update_thread.start()

    def on_update_checked(self, available, url, version_tag, silent):
        if available:
            reply = QMessageBox.question(
                self, 
                "Atualização Disponível", 
                f"Uma nova versão ({version_tag}) está disponível.\nDeseja baixar e instalar agora?\n\nO aplicativo será reiniciado.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.perform_update(url)
        else:
            if not silent:
                QMessageBox.information(self, "Atualizado", f"Você já está na versão mais recente ({APP_VERSION}).")
        
        self.statusBar().showMessage("Pronto.")

    def perform_update(self, url):
        # Since the updater calls sys.exit(), we can run it directly 
        # or inside a thread if we want a progress bar. 
        # For simplicity, let's run it carefully.
        try:
            updater = Updater()
            # Defining a simple callback to update log/status
            def update_status(msg):
                self.statusBar().showMessage(msg)
                self.log_text_edit.append(msg)
                QApplication.processEvents() # Keep UI responsive
            
            updater.download_and_install(url, status_callback=update_status)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha na atualização: {e}")

    def open_settings_dialog(self): SettingsDialog(self).exec()
    def export_custom_texts(self): self.export_custom_texts_logic() 
    def import_custom_texts(self): self.import_custom_texts_logic() 
    def open_text_editor(self): TextEditorDialog(self).exec()

    def _setup_ui(self):
        central_widget = QWidget(); self.setCentralWidget(central_widget); main_layout = QVBoxLayout(central_widget)
        top_layout = QHBoxLayout()
        left_panel = QWidget(); left_layout = QVBoxLayout(left_panel); left_layout.setContentsMargins(0,0,10,0)
        
        file_group = QGroupBox("Seleção de Ficheiros"); file_layout = QVBoxLayout()
        m_layout = QHBoxLayout(); self.master_file_path_edit = QLineEdit(); m_layout.addWidget(QLabel("Mestre:")); m_layout.addWidget(self.master_file_path_edit); 
        mb = QPushButton("Procurar..."); mb.clicked.connect(self.browse_master_file); m_layout.addWidget(mb)
        i_layout = QHBoxLayout(); self.invoices_file_path_edit = QLineEdit(); i_layout.addWidget(QLabel("Notas:")); i_layout.addWidget(self.invoices_file_path_edit); 
        ib = QPushButton("Procurar..."); ib.clicked.connect(self.browse_invoices_file); i_layout.addWidget(ib)
        file_layout.addLayout(m_layout); file_layout.addLayout(i_layout); file_group.setLayout(file_layout)
        
        comp_group = QGroupBox("Seleção da Empresa"); c_layout = QHBoxLayout()
        self.company_combo = QComboBox(); self.company_combo.setEnabled(False)
        self.add_company_button = QPushButton("Adicionar"); self.add_company_button.clicked.connect(self.prepare_add_new_company); self.add_company_button.setEnabled(False)
        c_layout.addWidget(QLabel("Empresa:")); c_layout.addWidget(self.company_combo, 1); c_layout.addWidget(self.add_company_button); comp_group.setLayout(c_layout)
        
        cad_group = QGroupBox("Cadastro"); cad_layout = QFormLayout()
        self.razao_social_edit = QLineEdit(); self.cnpj_edit = QLineEdit(); self.imu_edit = QLineEdit()
        self.endereco_edit = QLineEdit(); self.cep_edit = QLineEdit(); self.epaf_numero_edit = QLineEdit()
        cad_layout.addRow("Razão Social:", self.razao_social_edit); cad_layout.addRow("CNPJ:", self.cnpj_edit)
        cad_layout.addRow("IMU:", self.imu_edit); cad_layout.addRow("Endereço:", self.endereco_edit)
        cad_layout.addRow("CEP:", self.cep_edit); cad_layout.addRow("ePAF:", self.epaf_numero_edit)
        self.save_cadastro_button = QPushButton("Salvar Alterações"); self.save_cadastro_button.clicked.connect(self.save_cadastro_data); cad_layout.addRow(self.save_cadastro_button)
        cad_group.setLayout(cad_layout)
        
        left_layout.addWidget(file_group); left_layout.addWidget(comp_group); left_layout.addWidget(cad_group)
        
        right_panel = QWidget(); right_layout = QVBoxLayout(right_panel); right_layout.setContentsMargins(10,0,0,0)
        
        # ✅ RESTAURADO LAYOUT DO FLUXO (Botões Verticais / Originais)
        an_group = QGroupBox("Fluxo")
        an_layout = QHBoxLayout()
        
        # Coluna Principal (Fluxo Obrigatório)
        main_path = QVBoxLayout()
        load_layout = QHBoxLayout()
        load_layout.setContentsMargins(0,0,0,0)
        
        self.load_invoices_button = QPushButton("1. Carregar Notas (Arquivo)")
        self.load_invoices_button.clicked.connect(self.load_company_invoices)
        
        self.load_multi_year_button = QPushButton("📂 Multi-Ano (Pasta)")
        self.load_multi_year_button.setToolTip("Carregar múltiplos anos (faturas e DAMs) de uma pasta.")
        self.load_multi_year_button.setStyleSheet("background-color: #4C566A; border: 1px dashed #8FBCBB;")
        self.load_multi_year_button.clicked.connect(self.load_multi_year_invoices)
        
        load_layout.addWidget(self.load_invoices_button, 2)
        load_layout.addWidget(self.load_multi_year_button, 1)

        # Update main_path to use this layout
        # main_path.addWidget(self.load_invoices_button) <--- REMOVE THIS
        main_path.addLayout(load_layout)

        self.btn_review_activities = QPushButton("🔍 1.5. Revisar Atividades / Local")
        self.btn_review_activities.setToolTip("Visualizar resumo de códigos, descrições completas e definir Local Tomador.")
        self.btn_review_activities.clicked.connect(self.open_activity_review)
        self.btn_review_activities.setEnabled(False) # Default disabled until loaded
        main_path.addWidget(self.btn_review_activities)

        self.idd_mode_checkbox = QCheckBox("Modo IDD")
        self.run_rules_button = QPushButton("2. Análise de Regras"); self.run_rules_button.clicked.connect(self.start_rules_analysis_thread)
        self.review_button = QPushButton("3. Rever e Criar Autos"); self.review_button.clicked.connect(self.open_review_wizard)
        
        # --- ✅ NEW BUTTON FOR AUTOMATIC IDD ---
        self.auto_idd_button = QPushButton("🚀 Gerar IDD Automático (Beta)")
        self.auto_idd_button.setToolTip("Calcula, valida no site, emite e gera o PDF automaticamente.")
        self.auto_idd_button.setStyleSheet("background-color: #5e81ac; color: white; font-weight: bold;")
        self.auto_idd_button.clicked.connect(self.open_automatic_idd_tool)
        # ---------------------------------------

        main_path.addWidget(self.idd_mode_checkbox)
        main_path.addWidget(self.run_rules_button)
        main_path.addWidget(self.review_button)
        main_path.addWidget(self.auto_idd_button) # Add new button here
        
        # NOVO: Botão de Parada Central
        self.stop_button = QPushButton("🛑 Parar Processo")
        self.stop_button.clicked.connect(self.stop_process_thread)
        self.stop_button.setEnabled(False)
        self.stop_button.setStyleSheet("background-color: #f44336; color: white;")
        main_path.addWidget(self.stop_button) 
        
        main_path.addStretch() # Empurra botões para cima
        
        # Coluna Opcional (IA)
        ai_path = QVBoxLayout()
        self.run_ai_button = QPushButton("Opcional: Análise (IA)"); self.run_ai_button.clicked.connect(self.start_ai_analysis_thread)
        self.relabel_button = QPushButton("Opcional: Revisar (IA)"); self.relabel_button.clicked.connect(self.open_relabeling_window)
        
        ai_path.addWidget(self.run_ai_button)
        ai_path.addWidget(self.relabel_button)
        ai_path.addStretch() # Empurra botões para cima
        
        an_layout.addLayout(main_path)
        an_layout.addLayout(ai_path)
        an_group.setLayout(an_layout)
        
        right_layout.addWidget(an_group)

        auditor_group = QGroupBox("Identificação do Auditor")
        auditor_layout = QFormLayout()
        
        self.auditor_name_edit = QLineEdit()
        self.auditor_name_edit.setPlaceholderText("Nome Completo")
        # Save on editing finished (Enter or LostFocus)
        self.auditor_name_edit.editingFinished.connect(self.save_auditor_info)

        self.auditor_matricula_edit = QLineEdit()
        self.auditor_matricula_edit.setPlaceholderText("Matrícula")
        self.auditor_matricula_edit.editingFinished.connect(self.save_auditor_info)

        auditor_layout.addRow("Nome:", self.auditor_name_edit)
        auditor_layout.addRow("Matrícula:", self.auditor_matricula_edit)
        auditor_group.setLayout(auditor_layout)

        right_layout.addWidget(auditor_group)

        right_layout.addStretch(1) # Empurra o groupbox para cima, ocupando o resto do painel direito com vazio
        
        top_layout.addWidget(left_panel, 1); top_layout.addWidget(right_panel, 1); main_layout.addLayout(top_layout, 0)
        
        log_group = QGroupBox("Registo"); log_layout = QVBoxLayout()
        self.log_text_edit = QTextEdit(); self.log_text_edit.setReadOnly(True); log_layout.addWidget(self.log_text_edit)
        log_group.setLayout(log_layout); main_layout.addWidget(log_group, 1)

    def _get_cadastro_path(self, master_path): return f"{os.path.splitext(master_path)[0]}_cadastro{os.path.splitext(master_path)[1]}"
    
    def load_auditor_info(self):
        """Loads auditor info from config into the UI."""
        texts = get_custom_general_texts()
        self.auditor_name_edit.setText(texts.get("AUDITOR_NOME", ""))
        self.auditor_matricula_edit.setText(texts.get("AUDITOR_MATRICULA", ""))

    def save_auditor_info(self):
        """Saves UI auditor info back to config."""
        texts = get_custom_general_texts()
        
        # Update values
        new_name = self.auditor_name_edit.text()
        new_matricula = self.auditor_matricula_edit.text()
        
        # Only save if changed to avoid unnecessary writes
        if texts.get("AUDITOR_NOME") != new_name or texts.get("AUDITOR_MATRICULA") != new_matricula:
            texts["AUDITOR_NOME"] = new_name
            texts["AUDITOR_MATRICULA"] = new_matricula
            set_custom_general_texts(texts)
            self.statusBar().showMessage("Dados do auditor salvos.", 2000)

    def browse_master_file(self):
        f, _ = QFileDialog.getOpenFileName(self, "Master", "", "Excel (*.xlsx *.xls)")
        if f: 
            self.master_file_path_edit.setText(f); self.load_companies(f); self.activity_data = load_activity_data()
            self.check_if_ready_to_load(); self._save_last_paths()
            self.save_cadastro_button.setEnabled(True); self.add_company_button.setEnabled(True)
    
    def browse_invoices_file(self):
        f, _ = QFileDialog.getOpenFileName(self, "Notas", "", "Excel (*.xlsx *.xls)")
        if f:
            self.invoices_file_path_edit.setText(f); self.check_if_ready_to_load()
            self.auto_select_company_from_invoices(f); self._save_last_paths()

    def auto_select_company_from_invoices(self, file_path):
        if self.df_empresas is None or self.df_empresas.empty: return
        try:
            try: df_sample = pd.read_excel(file_path, skiprows=2, nrows=1000, usecols=['CNPJ PRESTADOR'])
            except: df_sample = pd.read_excel(file_path, skiprows=2, nrows=1000)
            if 'CNPJ PRESTADOR' not in df_sample.columns: return
            
            # ✅ FIX: Clean CNPJ from sample to match loaded cadastro
            found_raw = df_sample['CNPJ PRESTADOR'].dropna().unique()
            found_clean = [re.sub(r'\D', '', str(c)) for c in found_raw]
            
            # Match against cleaned cadastro CNPJs
            cadastro_cnpjs = self.df_empresas[Columns.CNPJ].astype(str).str.replace(r'\D', '', regex=True).tolist()
            
            matched_indices = [i for i, c in enumerate(cadastro_cnpjs) if c in found_clean]
            
            if len(matched_indices) == 1:
                # Get the original formatted CNPJ from the dataframe to set in combo
                original_cnpj = self.df_empresas.iloc[matched_indices[0]][Columns.CNPJ]
                
                idx = self.company_combo.findData(original_cnpj)
                if idx != -1: 
                    self.company_combo.setCurrentIndex(idx)
                    self.statusBar().showMessage("Auto-selecionado.", 3000)
        except Exception as e: 
            print(f"Auto-select error: {e}")

    def load_companies(self, file_path):
        cp = self._get_cadastro_path(file_path)
        sc = {Columns.CNPJ: str, Columns.RAZAO_SOCIAL: str, Columns.IMU: str, Columns.ENDERECO: str, Columns.CEP: str, Columns.EPAF_NUMERO: str}
        try:
            df = pd.read_excel(cp, dtype=sc) if os.path.exists(cp) else None
            if df is None and os.path.exists(file_path):
                xls = pd.ExcelFile(file_path)
                if 'Empresas' in xls.sheet_names: df = pd.read_excel(file_path, sheet_name='Empresas', dtype=sc); df.to_excel(cp, index=False)
            self.df_empresas = df if df is not None else pd.DataFrame(columns=sc.keys())
            for col in sc: 
                if col not in self.df_empresas.columns: self.df_empresas[col] = ''
                self.df_empresas[col] = self.df_empresas[col].fillna('').astype(str)
            
            # ⚠️ NOTE: We are NOT stripping CNPJ here anymore to preserve formatting for display/lookup
            # If your 'main.py' expects stripped, we will strip it ONLY when passing to the worker.
            
            self.company_combo.clear()
            for _, r in self.df_empresas.iterrows(): 
                # Store the CNPJ exactly as it is in the file (formatted) in UserData
                self.company_combo.addItem(f"{r[Columns.RAZAO_SOCIAL]} ({r[Columns.CNPJ]})", userData=r[Columns.CNPJ])
            
            self.company_combo.setEnabled(True); self.save_cadastro_button.setEnabled(True); self.add_company_button.setEnabled(True)
            
            # ✅ FIX: DO NOT auto-select the first company. Wait for invoice file to trigger selection.
            # self.company_combo.setCurrentIndex(0) 
            self.company_combo.setCurrentIndex(-1)
            self.update_cadastro_display()
            
        except Exception as e:
            self.df_empresas = pd.DataFrame(columns=sc.keys()); self.company_combo.clear()

    def prepare_add_new_company(self, cnpj_to_add=None):
        self.company_combo.setCurrentIndex(-1); self.razao_social_edit.clear(); self.cnpj_edit.clear()
        self.imu_edit.clear(); self.endereco_edit.clear(); self.cep_edit.clear(); self.epaf_numero_edit.clear()
        self.cnpj_edit.setReadOnly(False)
        if cnpj_to_add: self.cnpj_edit.setText(cnpj_to_add); self.razao_social_edit.setFocus()
        else: self.cnpj_edit.setFocus()
        self.save_cadastro_button.setText("Salvar Nova Empresa")

    def on_company_selection_change(self, index):
        if index == -1: self.save_cadastro_button.setText("Salvar Nova Empresa")
        else: self.update_cadastro_display(); self.save_cadastro_button.setText("Salvar Alterações no Cadastro")

    def update_cadastro_display(self):
        cnpj = self.company_combo.currentData()
        if cnpj and self.df_empresas is not None:
            row = self.df_empresas[self.df_empresas[Columns.CNPJ] == cnpj]
            if not row.empty:
                d = row.iloc[0].to_dict()
                self.razao_social_edit.setText(d.get(Columns.RAZAO_SOCIAL,'')); self.cnpj_edit.setText(d.get(Columns.CNPJ,''))
                self.imu_edit.setText(d.get(Columns.IMU,'')); self.endereco_edit.setText(d.get(Columns.ENDERECO,''))
                self.cep_edit.setText(d.get(Columns.CEP,'')); self.epaf_numero_edit.setText(d.get(Columns.EPAF_NUMERO,''))
                return
        self.razao_social_edit.clear(); self.cnpj_edit.clear(); self.imu_edit.clear()
        self.endereco_edit.clear(); self.cep_edit.clear(); self.epaf_numero_edit.clear()
        self.cnpj_edit.setReadOnly(self.company_combo.currentIndex() != -1)

    def save_cadastro_data(self):
        mp = self.master_file_path_edit.text()
        if not mp: return
        cp = self._get_cadastro_path(mp)
        
        # Get text exactly as entered (preserving format if user typed it)
        cur_cnpj = self.cnpj_edit.text().strip()
        
        if not cur_cnpj: QMessageBox.warning(self, "Erro", "CNPJ Obrigatório"); return
        
        try:
            # Determine if existing or new based on cleaned comparison, but save what is in text box
            clean_current = re.sub(r'\D', '', cur_cnpj)
            
            # Find based on cleaned version
            idx = self.df_empresas.index[self.df_empresas[Columns.CNPJ].astype(str).str.replace(r'\D', '', regex=True) == clean_current]
            
            data = {Columns.RAZAO_SOCIAL: self.razao_social_edit.text(), Columns.CNPJ: cur_cnpj,
                    Columns.IMU: self.imu_edit.text(), Columns.ENDERECO: self.endereco_edit.text(),
                    Columns.CEP: self.cep_edit.text(), Columns.EPAF_NUMERO: self.epaf_numero_edit.text()}
            
            new_comp = False
            if not idx.empty:
                # Update existing
                for k,v in data.items(): self.df_empresas.loc[idx[0], k] = v
                # Update combo item text
                ci = self.company_combo.findData(self.df_empresas.loc[idx[0], Columns.CNPJ]) # Find by old data if possible, might be tricky if key changed
                
                # Simpler: just reload list logic or update current
                # Since we match by Data, and Data is CNPJ, if CNPJ changed we have a problem.
                # For now, assume CNPJ edit creates new or updates in place. 
                
                # Re-find index in combo by iterating (safe way)
                for i in range(self.company_combo.count()):
                    item_clean_cnpj = re.sub(r'\D', '', str(self.company_combo.itemData(i)))
                    if item_clean_cnpj == clean_current:
                        self.company_combo.setItemText(i, f"{data[Columns.RAZAO_SOCIAL]} ({cur_cnpj})")
                        self.company_combo.setItemData(i, cur_cnpj)
                        break

            else:
                new_comp = True
                self.df_empresas = pd.concat([self.df_empresas, pd.DataFrame([data])], ignore_index=True)
                self.company_combo.addItem(f"{data[Columns.RAZAO_SOCIAL]} ({cur_cnpj})", userData=cur_cnpj)
                self.company_combo.setCurrentIndex(self.company_combo.count()-1)
            
            self.df_empresas.to_excel(cp, index=False)
            self.statusBar().showMessage("Salvo.", 3000)
            self.cnpj_edit.setReadOnly(False); self.save_cadastro_button.setText("Salvar Alterações")
            if new_comp: self.on_company_selection_change(self.company_combo.currentIndex())
        except Exception as e: self.log_text_edit.append(f"Erro ao salvar: {e}")

    def check_if_ready_to_load(self): self.update_button_states()

    def update_button_states(self):
        states = self.workflow_state
        any_run = any(s == 'running' for s in states.values())
        ready = bool(self.master_file_path_edit.text() and self.invoices_file_path_edit.text() and (self.company_combo.count()>0 or not self.cnpj_edit.isReadOnly()))
        
        # Stop Button State
        self.stop_button.setEnabled(any_run)
        
        b = self.load_invoices_button
        if states['load'] == 'running': b.setEnabled(False); b.setText("1. A Carregar...")
        else: b.setEnabled(ready and not any_run); b.setText("1. Carregar Notas")
        
        b = self.run_rules_button
        can = states['load'] in ['completed', 'stale']
        if states['rules'] == 'running': b.setEnabled(False); b.setText("2. A Analisar...")
        else: b.setEnabled(can and not any_run); b.setText("2. Re-executar" if states['rules']=='stale' else "2. Análise de Regras")
        
        b = self.review_button
        can = states['rules'] in ['completed', 'stale'] or states['load'] in ['completed', 'stale']
        if states['review'] == 'running': b.setEnabled(False); b.setText("3. A Gerar...")
        else: b.setEnabled(can and not any_run); b.setText("3. Rever e Criar Autos")
        
        # --- Auto Button State ---
        b = self.auto_idd_button
        # Only enable if rules are done and not running
        can_auto = states['rules'] in ['completed'] and not any_run
        if states.get('tool_auto_idd') == 'running':
            b.setEnabled(False); b.setText("🚀 Gerando Automático...")
        else:
            b.setEnabled(can_auto); b.setText("🚀 Gerar IDD Automático (Beta)")
        
        # Other buttons simplified...
        self.run_ai_button.setEnabled(states['load'] in ['completed', 'stale'] and not any_run)
        can_review = states['load'] in ['completed', 'stale']
        self.btn_review_activities.setEnabled(can_review and not any_run)

    def open_activity_review(self):
        """Opens the Activity & Location Review Dialog."""
        # Use company_invoices_df (working copy) if available, else clean
        if self.company_invoices_df is not None and not self.company_invoices_df.empty:
            target_df = self.company_invoices_df
        elif self.clean_invoices_df is not None:
            target_df = self.clean_invoices_df.copy()
        else:
            return

        # Ensure we have activity data
        if not self.activity_data:
            self.activity_data = load_activity_data() # From main.py util

        dialog = ActivityReviewDialog(target_df, self.activity_data, self)
        
        if dialog.exec():
            # Update the working dataframe with changes (manual status, new codes)
            self.company_invoices_df = dialog.get_updated_dataframe()
            self.log_text_edit.append("✅ Revisão de atividades/local aplicada. Re-execute as Regras para surtir efeito nos Autos.")
            
            # Invalidate Rules State so user knows to run them again
            self.workflow_state['rules'] = 'pending' 
            self.update_button_states()

    def update_log_from_thread(self, msg): self.log_text_edit.append(msg)
    
    def on_thread_error(self, msg):
        self.log_text_edit.append(msg); logging.error(msg)
        for k,v in self.workflow_state.items(): 
            if v == 'running': self.workflow_state[k] = 'pending'
        if self.thread and self.thread.isRunning(): 
            try:
                # Attempt graceful quit first
                self.thread.quit()
                self.thread.wait(500)
            except:
                pass
        self.update_button_states()
        
    def stop_process_thread(self):
        """
        Safely stops the currently running worker thread by sending the stop signal.
        """
        if self.thread and self.worker and self.thread.isRunning():
            self.log_text_edit.append("🛑 Recebida requisição de parada do usuário.")
            
            # 1. Emit the stop signal to the worker
            if hasattr(self.worker, 'stop_signal'):
                self.worker.stop_signal.emit()
            
            # 2. Wait for the thread to finish gracefully (up to 1 second)
            if not self.thread.wait(1000):
                # 3. Force cleanup if it didn't stop (necessary for RPA/pywinauto)
                self.thread.terminate() 
                self.thread.wait(1000)
                self.log_text_edit.append("⚠️ Thread terminada à força.")
            else:
                 self.log_text_edit.append("✅ Parada processada pela thread.")

            # 4. Reset state
            for k,v in self.workflow_state.items(): 
                if v == 'running': self.workflow_state[k] = 'pending'
            self.update_button_states()
        else:
            self.log_text_edit.append("ℹ️ Nenhum processo ativo para parar.")
            
    def reset_all_state(self):
        """
        Reinicializa o estado de todas as variáveis de análise, DFs e o fluxo de trabalho.
        """
        if self.thread and self.thread.isRunning():
            QMessageBox.warning(self, "Aviso", "Pare o processo em execução antes de reiniciar o aplicativo.")
            return

        reply = QMessageBox.question(self, 'Confirmar Reinicialização',
            "Tem certeza de que deseja reiniciar o estado do aplicativo? Todos os dados de análise de notas serão perdidos.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.clean_invoices_df = None 
            self.company_invoices_df = None
            self.infraction_groups = {}
            self.activity_list = []
            self._temp_multi_year_dams_path = None
            # Resetar o estado do fluxo de trabalho
            self.workflow_state = {
                'load': 'pending', 
                'rules': 'pending',
                'ai': 'pending',
                'relabel': 'pending',
                'review': 'pending'
            }
            
            # Limpar a UI de exibição de dados e logs
            self.log_text_edit.clear()
            self.razao_social_edit.clear()
            
            QMessageBox.information(self, "Sucesso", "Estado do aplicativo reiniciado com sucesso.")
            self.statusBar().showMessage("Pronto. Estado limpo.")
            self.update_button_states()


    def _start_worker_thread(self, worker_class, worker_key, *args, on_finished_slot):
        # 1. Force Cleanup of previous thread/worker to prevent zombies
        if self.worker:
            try: 
                self.worker.progress.disconnect()
                self.worker.error.disconnect()
                self.worker.finished.disconnect()
            except: 
                pass
            self.worker.deleteLater()
            self.worker = None
        
        if self.thread:
            try:
                if self.thread.isRunning(): 
                    self.thread.quit()
                    self.thread.wait()
                self.thread.deleteLater()
            except RuntimeError:
                pass
            self.thread = None
            
        gc.collect() 

        self.workflow_state[worker_key] = 'running'
        self.update_button_states()
        
        self.thread = QThread()
        self.worker = worker_class(*args)
        self.worker.moveToThread(self.thread)
        self.worker.progress.connect(self.update_log_from_thread)
        self.worker.error.connect(self.on_thread_error)
        self.worker.finished.connect(on_finished_slot)
        
        self.thread.started.connect(self.worker.run)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def load_company_invoices(self):
        self.log_text_edit.clear()
        cc = self.company_combo.currentData()
        if not cc and not self.cnpj_edit.isReadOnly(): cc = self.cnpj_edit.text().strip()
        
        if not cc: QMessageBox.warning(self, "Erro", "CNPJ Inválido"); return
        
        # To be safe, let's try to pass the FORMATTED version first (since that's what's in your cadastro now)
        self._temp_multi_year_dams_path = None
        self.log_text_edit.append(f"--- Carregar Notas (CNPJ Alvo: {cc}) ---")
        self.statusBar().showMessage("A carregar...")
        self._start_worker_thread(AIPrepWorker, 'load', self.master_file_path_edit.text(), self.invoices_file_path_edit.text(), cc, on_finished_slot=self.on_invoices_loaded)

    def on_invoices_loaded(self, df):
        if df.empty: 
            self.log_text_edit.append("Nenhuma nota encontrada..."); 
            self.workflow_state['load'] = 'pending'
            self.update_button_states()
            return

        # ✅ CHECK FOR CONFLICTS
        if '_is_conflict' in df.columns and df['_is_conflict'].any():
            conflict_count = df[df['_is_conflict']]['NÚMERO'].nunique()
            
            # User Feedback
            self.log_text_edit.append(f"⚠️ ATENÇÃO: {conflict_count} notas possuem duplicatas com dados conflitantes.")
            
            reply = QMessageBox.question(
                self, 
                "Conflito de Notas", 
                f"Foram detectadas <b>{conflict_count} notas duplicadas</b> com valores ou dados diferentes.<br><br>"
                "O sistema selecionou automaticamente a versão com maior valor/alíquota.<br>"
                "Deseja revisar e escolher manualmente qual versão manter?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                # Open Dialog
                dialog = DuplicateReviewDialog(df, parent=self)
                if dialog.exec():
                    # Get cleaned list of indices
                    indices_to_keep = dialog.get_resolved_indices()
                    df = df.loc[indices_to_keep].copy()
                    self.log_text_edit.append("✅ Revisão de duplicatas aplicada manualmente.")
                else:
                    self.log_text_edit.append("⚠️ Revisão cancelada. Mantendo seleção automática padrão.")
                    # Fallback to smart sort drop (first)
                    df.drop_duplicates(subset=['NÚMERO'], keep='first', inplace=True)
            else:
                # User chose NO -> Auto Resolve
                self.log_text_edit.append("ℹ️ Mantendo seleção automática (Maior Valor/Alíquota).")
                df.drop_duplicates(subset=['NÚMERO'], keep='first', inplace=True)
                
            # Clean up temp column
            df.drop(columns=['_is_conflict'], inplace=True)

        # ✅ Store BOTH Clean and Working copies (Standard Logic)
        self.clean_invoices_df = df.copy()
        self.company_invoices_df = df.copy()
        self.log_text_edit.append(f"✅ {len(df)} notas carregadas e validadas.")
        self.workflow_state['load'] = 'completed'
        self.workflow_state['rules'] = 'pending'
        self.workflow_state['review'] = 'pending'
        self.update_button_states()

    def start_rules_analysis_thread(self):
        # ✅ CHANGE: Prioritize 'company_invoices_df' (Working Copy) if it exists.
        # Only use 'clean_invoices_df' (Original Copy) if starting from scratch.
        
        df_to_use = None
        
        if self.company_invoices_df is not None and not self.company_invoices_df.empty:
            # Use the version that contains Relabeling changes
            df_to_use = self.company_invoices_df.copy()
            self.log_text_edit.append("ℹ️ Usando dados editados (com revisões).")
        elif self.clean_invoices_df is not None:
            # Use raw version
            df_to_use = self.clean_invoices_df.copy()
            self.log_text_edit.append("ℹ️ Usando dados originais (limpos).")
        else:
            return

        self.log_text_edit.append("\n--- Análise de Regras ---")
        
        self._start_worker_thread(
            RulesPrepWorker, 
            'rules', 
            df_to_use, # ✅ Passing the correct DF
            self.idd_mode_checkbox.isChecked(), 
            on_finished_slot=self.on_rules_analysis_finished
        )

    def on_rules_analysis_finished(self, groups, df):
        self.infraction_groups = groups
        self.company_invoices_df = df 
        self.log_text_edit.append(f"✅ Regras concluídas. {len(groups)} grupos.")
        # ✅ Logic to delete session file REMOVED. We rely on ReviewWizard.load_session to handle it.
        self.workflow_state['rules'] = 'completed'
        self.workflow_state['review'] = 'pending'
        self.update_button_states()

    def start_ai_analysis_thread(self):
        if self.clean_invoices_df is None: return
        self.log_text_edit.append("--- IA ---")
        # ✅ Use CLEAN COPY for AI
        self._start_worker_thread(AIAnalysisWorker, 'ai', self.clean_invoices_df.copy(), on_finished_slot=self.on_ai_analysis_finished)

    def on_ai_analysis_finished(self, df):
        self.company_invoices_df = df # Update working copy
        self.log_text_edit.append("✅ IA concluída.")
        self.workflow_state['ai'] = 'completed'; self.workflow_state['relabel'] = 'pending'
        self.update_button_states()

    def open_relabeling_window(self):
        if not self.activity_data: QMessageBox.warning(self, "Aviso", "Carregue o Mestre."); return
        if self.workflow_state['ai'] != 'completed': QMessageBox.warning(self, "Aviso", "Execute IA."); return
        
        d = RelabelingWindow(self.company_invoices_df, self.activity_data, self)
        if d.exec():
            self.company_invoices_df = d.get_updated_dataframe()
            self.log_text_edit.append("✅ Relabeling feito. Re-execute regras.")
            self.workflow_state['relabel'] = 'completed'
            self.workflow_state['rules'] = 'stale' # Mark rules as stale
            self.update_button_states()

    def start_generation_thread(self, final_data, preview_context, numero_multa, dam_filepath, pgdas_folder_path, output_dir_override, encerramento_version, company_imu):        
        cc = self.company_combo.currentData() or self.cnpj_edit.text().strip()
        if not cc: return
        self.log_text_edit.append("\n--- Gerar Docs ---")
        
        # ✅ Pass COPY
        self._start_worker_thread(GenerationWorker, 'review', cc, self.master_file_path_edit.text(), final_data, preview_context, self.company_invoices_df.copy(), numero_multa, dam_filepath, pgdas_folder_path, output_dir_override, encerramento_version, company_imu, self.idd_mode_checkbox.isChecked(), on_finished_slot=self.on_generation_finished)

    def on_generation_finished(self, success, output_dir):
        if not success: self.log_text_edit.append("❌ Falha."); self.workflow_state['review'] = 'stale'
        else:
            self.log_text_edit.append("✅ Sucesso.")
            self.workflow_state['review'] = 'completed'
            try:
                files = [f for f in os.listdir(output_dir) if not f.startswith('temp_') and f.endswith(('.docx','.pdf','.xlsx'))]
                GenerationSummaryDialog(output_dir, files, self).exec()
            except: pass
        self.update_button_states()

    def load_multi_year_invoices(self):
        self.log_text_edit.clear()

        self._temp_multi_year_dams_path = None
        
        # 1. Ask for folder FIRST (Moved up)
        folder_path = QFileDialog.getExistingDirectory(self, "Selecione a Pasta com Ficheiros dos Anos", "")
        if not folder_path: return

        # 2. Check for current company. If missing, try to auto-detect from files in the folder.
        cc = self.company_combo.currentData() or self.cnpj_edit.text().strip()
        
        if not cc: 
            self.statusBar().showMessage("Tentando detectar empresa nos arquivos...")
            try:
                # Look for Excel files in the folder to scan for CNPJ
                potential_files = [
                    f for f in os.listdir(folder_path) 
                    if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('~$')
                ]
                
                # Sort roughly to try "main" files first, though any invoice file should work
                potential_files.sort()

                for f in potential_files:
                    full_path = os.path.join(folder_path, f)
                    # Reuse the existing auto-select logic
                    self.auto_select_company_from_invoices(full_path)
                    
                    # Check if detection was successful
                    cc = self.company_combo.currentData() or self.cnpj_edit.text().strip()
                    if cc:
                        self.log_text_edit.append(f"ℹ️ Empresa detectada automaticamente: {self.company_combo.currentText()}")
                        break
            except Exception as e:
                print(f"Auto-detection error in folder: {e}")

        # 3. Final validation
        if not cc: 
            QMessageBox.warning(self, "Erro", "Não foi possível detectar a empresa automaticamente.\n\nSelecione a empresa manualmente na lista acima ou verifique se os arquivos na pasta contêm o CNPJ correto."); 
            return

        self.log_text_edit.append(f"--- Carregar Multi-Ano (CNPJ: {cc}) ---")
        self.statusBar().showMessage("Consolidando arquivos...")
        
        self._start_worker_thread(
            MultiYearPrepWorker, 
            'load', 
            folder_path, 
            self.master_file_path_edit.text(), 
            cc, 
            on_finished_slot=self.on_multi_year_loaded
        )

    def on_multi_year_loaded(self, df, dams_path):
        if df.empty:
            self.log_text_edit.append("❌ Erro: DataFrame consolidado vazio.")
            self.workflow_state['load'] = 'pending'
            self.update_button_states()
            return

        self._temp_multi_year_dams_path = None
        # ✅ CHECK FOR CONFLICTS (Same logic as on_invoices_loaded)
        if '_is_conflict' in df.columns and df['_is_conflict'].any():
            conflict_count = df[df['_is_conflict']]['NÚMERO'].nunique()
            
            self.log_text_edit.append(f"⚠️ ATENÇÃO: {conflict_count} conflitos encontrados entre os arquivos dos anos.")
            
            reply = QMessageBox.question(
                self, 
                "Conflito Multi-Ano", 
                f"Foram detectadas <b>{conflict_count} notas duplicadas</b> com dados diferentes entre os arquivos carregados.<br><br>"
                "O sistema selecionou automaticamente a melhor versão (Maior Valor/Alíquota).<br>"
                "Deseja revisar e escolher manualmente?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                dialog = DuplicateReviewDialog(df, parent=self)
                if dialog.exec():
                    indices_to_keep = dialog.get_resolved_indices()
                    df = df.loc[indices_to_keep].copy()
                    self.log_text_edit.append("✅ Revisão de duplicatas aplicada manualmente.")
                else:
                    self.log_text_edit.append("⚠️ Revisão cancelada. Usando automático.")
                    # Default: keep first (sorted best)
                    df.drop_duplicates(subset=['NÚMERO'], keep='first', inplace=True)
            else:
                self.log_text_edit.append("ℹ️ Mantendo seleção automática.")
                df.drop_duplicates(subset=['NÚMERO'], keep='first', inplace=True)
                
            df.drop(columns=['_is_conflict'], inplace=True)

        self.clean_invoices_df = df.copy()
        self.company_invoices_df = df.copy()
        self.log_text_edit.append(f"✅ Multi-Ano Carregado com Sucesso!")
        self.log_text_edit.append(f"📊 Total de Notas: {len(df)}")
        self.invoices_file_path_edit.setText(f"[Multi-Ano] {len(df)} registros")
        
        if dams_path:
            self._temp_multi_year_dams_path = dams_path 
            self.log_text_edit.append(f"💰 DAMs consolidados vinculados.")

        self.workflow_state['load'] = 'completed'
        self.workflow_state['rules'] = 'pending'
        self.workflow_state['review'] = 'pending'
        self.update_button_states()

    def open_review_wizard(self):
        if self.company_invoices_df is None or self.company_invoices_df.empty: 
            QMessageBox.warning(self,"Aviso","Carregue notas."); return
        
        # ✅ UPDATE: Check for multi-year temp path
        found_dam_path = getattr(self, '_temp_multi_year_dams_path', None)
        
        if not found_dam_path:
            # (Original logic for single file)
            try:
                invoices_path = self.invoices_file_path_edit.text()
                if invoices_path and os.path.exists(invoices_path) and "[Multi-Ano]" not in invoices_path:
                    folder = os.path.dirname(invoices_path)
                    for f in os.listdir(folder):
                        if f.startswith("Relatorio_DAMS"):
                            found_dam_path = os.path.join(folder, f)
                            break
            except Exception: pass

        # ✅ CHANGE 1: Store wizard in 'self' to keep it alive (garbage collection protection)
        self.current_review_wizard = ReviewWizard(
            self.company_invoices_df, 
            self.infraction_groups, 
            self.cnpj_edit.text(), 
            self.razao_social_edit.text(), 
            self.imu_edit.text(), 
            self.epaf_numero_edit.text(), 
            self, # Parent is self, but we set modality below
            dam_file_path=found_dam_path 
        )
        
        # ✅ CHANGE 2: Connect the "Accepted" signal (OK button) to the processing slot
        self.current_review_wizard.accepted.connect(self.on_review_wizard_accepted)
        
        # ✅ CHANGE 3: Set to NonModal and use .show() instead of .exec()
        # This allows you to click back on the Main Window while Wizard is open
        self.current_review_wizard.setWindowModality(Qt.NonModal)
        self.current_review_wizard.show()

    def on_review_wizard_accepted(self):
        """
        Called when the user clicks 'Gerar Documentos' (OK) in the ReviewWizard.
        This contains the logic that used to be inside 'if rw.exec():'
        """
        # Retrieve the instance we stored
        rw = self.current_review_wizard
        if not rw: return

        try:
            final_selected_context = rw.confirm_page.get_selected_data()
            final_data = rw.get_final_data_for_confirmation()
            selected_keys = {auto['numero'] for auto in final_selected_context.get('autos', [])}
            final_data_filtered = {key: val for key, val in final_data.items() if key in selected_keys}

            if not final_selected_context.get('autos'):
                if QMessageBox.question(self, "Sem Autos", "Gerar sem autos?", QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No: return

            nm = ""
            dp = rw.fine_page.dam_filepath_edit.text()
            pp = rw.fine_page.pgdas_folder_path_edit.text()
            
            # Modal for version selection is fine here because the Wizard is technically "done"
            mb = QMessageBox(self); mb.setText("Versão?"); ar=mb.addButton("AR", QMessageBox.ButtonRole.YesRole); dec=mb.addButton("DEC", QMessageBox.ButtonRole.NoRole); mb.addButton("Cancelar", QMessageBox.ButtonRole.RejectRole)
            mb.exec()
            if mb.clickedButton() == ar: ev = "AR"
            elif mb.clickedButton() == dec: ev = "DEC"
            else: return
            
            od = QFileDialog.getExistingDirectory(self, "Destino", "", QFileDialog.Option.ShowDirsOnly)
            if not od: return
            
            self.start_generation_thread(final_data_filtered, final_selected_context, nm, dp, pp, od, ev, self.imu_edit.text())
        
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao processar dados do Wizard: {e}")
            logging.error(traceback.format_exc())
        finally:
            # Cleanup reference
            self.current_review_wizard = None

    # --- ✅ INÍCIO: NOVAS FUNÇÕES DAS FERRAMENTAS ---

    def open_automatic_idd_tool(self):
        """
        Orchestrates the automatic IDD generation logic.
        """
        if self.company_invoices_df is None or self.company_invoices_df.empty:
            QMessageBox.warning(self, "Aviso", "Carregue notas e execute as regras primeiro.")
            return

        # 1. Try to find DAM file automatically
        found_dam_path = None
        try:
            invoices_path = self.invoices_file_path_edit.text()
            if invoices_path and os.path.exists(invoices_path):
                folder = os.path.dirname(invoices_path)
                # Clean IMU logic not strictly necessary for file search if we assume it's in the same folder
                # but good for verification if needed. 
                # For auto-mode, we prioritize finding ANY file looking like a DAM report in that folder.
                
                for f in os.listdir(folder):
                    fname_lower = f.lower()
                    # Look for standard keywords in filename
                    if "relatorio" in fname_lower and "dam" in fname_lower:
                        if fname_lower.endswith(('.csv', '.xls', '.xlsx')):
                            found_dam_path = os.path.join(folder, f)
                            self.log_text_edit.append(f"ℹ️ DAMs encontrados: {f}")
                            break
        except Exception as e: 
            print(f"Auto-DAM search error: {e}")

        # 2. Use Headless ReviewWizard for Logic Reuse
        headless_wizard = ReviewWizard(
            self.company_invoices_df, 
            self.infraction_groups, 
            self.cnpj_edit.text(), 
            self.razao_social_edit.text(), 
            self.imu_edit.text(), 
            self.epaf_numero_edit.text(), 
            self,
            dam_file_path=found_dam_path 
        )
        
        preview_context = headless_wizard.calculate_preview_context()
        
        if not preview_context or 'summary' not in preview_context:
            QMessageBox.warning(self, "Erro", "Falha ao calcular valores.")
            return
            
        total_val = preview_context['summary'].get('total_geral_credito', 0.0)
        
        self._temp_auto_context = {
            'wizard_instance': headless_wizard,
            'preview_context': preview_context,
            'dam_path': found_dam_path
        }

        # Determine YEAR
        years = []
        for auto in preview_context.get('autos', []):
            for m in auto.get('dados_anuais', []):
                try: years.append(m['mes_ano'].split('/')[1])
                except: pass
        target_year = max(set(years), key=years.count) if years else str(datetime.now().year)

        # 3. Confirmation
        msg = QMessageBox(self)
        msg.setWindowTitle("Confirmar IDD Automático")
        msg.setText(f"Valor Calculado pelo App: <b>R$ {total_val:,.2f}</b><br>"
                    f"Ano Base: <b>{target_year}</b><br><br>"
                    "O robô irá validar este valor no sistema e emitir o IDD se coincidir.<br>"
                    "Deseja continuar?")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if msg.exec() != QMessageBox.StandardButton.Yes: return

        # 4. Automatic Folder Selection (from Invoices)
        invoices_path = self.invoices_file_path_edit.text()
        od = None
        if invoices_path and os.path.exists(invoices_path):
            od = os.path.dirname(invoices_path)
            self.log_text_edit.append(f"📂 Pasta de saída definida automaticamente: {od}")
        else:
            od = QFileDialog.getExistingDirectory(self, "Salvar Documentos em:", "", QFileDialog.Option.ShowDirsOnly)
            if not od:
                self.log_text_edit.append("⚠️ Processo cancelado: Pasta não selecionada.")
                return
        
        self._temp_auto_context['output_folder'] = od

        # 5. Start Worker
        self.log_text_edit.clear()
        self.log_text_edit.append(f"--- Iniciando IDD Automático ---")
        self.statusBar().showMessage("Executando IDD Automático...")
        
        self._start_worker_thread(
            AutomaticIDDWorker,
            'tool_auto_idd',
            self.imu_edit.text(),
            target_year,
            total_val,
            od, # Pass output folder
            on_finished_slot=self.on_auto_idd_finished
        )

    def on_auto_idd_finished(self, result):
        """
        Called when AutomaticIDDWorker finishes.
        result = { 'status': 'Success'|'Failed', 'idd_number': '...', 'protocolo': '...' }
        """
        if result.get("status") == "Success":
            idd_num = result.get("idd_number")
            protocolo = result.get("protocolo")
            
            self.log_text_edit.append(f"✅ IDD Emitido: {idd_num}")
            self.log_text_edit.append(f"✅ Protocolo: {protocolo}")
            
            # --- Trigger PDF Generation ---
            if hasattr(self, '_temp_auto_context'):
                wizard = self._temp_auto_context['wizard_instance']
                context = self._temp_auto_context['preview_context']
                dam_path = self._temp_auto_context['dam_path']
                od = self._temp_auto_context['output_folder']
                
                # 1. Update Global Context with Real Protocol
                context['epaf_numero'] = protocolo
                
                # 2. Update Auto Numbers in Context with Real IDD
                for auto in context.get('autos', []):
                    old_num = auto['numero']
                    auto['numero'] = idd_num 
                    self.log_text_edit.append(f"🔄 Atualizando auto {old_num} para IDD {idd_num}")
                
                # 3. Update Summary to reflect new number
                if 'summary' in context:
                    for auto_sum in context['summary'].get('autos', []):
                        auto_sum['numero'] = idd_num

                # 4. Prepare Final Data and Update Keys
                final_data = wizard.get_final_data_for_confirmation()
                new_final_data = {}
                
                for key, val in final_data.items():
                    val['auto_id'] = idd_num
                    new_final_data[idd_num] = val

                # Start Generation
                self.start_generation_thread(
                    new_final_data, 
                    context, 
                    "", # No Multa manually entered in auto mode
                    dam_path, 
                    "", # No PGDAS folder manual entry
                    od, 
                    "AR", # Default to AR
                    self.imu_edit.text()
                )
                self.log_text_edit.append("🔄 Iniciando geração do relatório PDF (Informação Fiscal)...")
                
        else:
            self.log_text_edit.append(f"❌ Processo automático não concluído. Status: {result.get('status')}")
            if result.get("status") == "Validated":
                 self.log_text_edit.append("ℹ️ O valor foi validado, mas a emissão não ocorreu (verifique flags).")
            
            # ⚠️ CHANGED: Ensure we STOP and don't proceed with generation if failed
            QMessageBox.warning(self, "Atenção", "O processo não obteve IDD ou Protocolo.\nA geração do documento fiscal será cancelada.")
            self.workflow_state['tool_auto_idd'] = 'completed' # Mark as done (failed)
            self.update_button_states()
            return # 🛑 Stop execution here
        
        self.workflow_state['tool_auto_idd'] = 'completed'
        self.update_button_states()

    def open_simples_reader_tool(self):
        """
        Abre um diálogo para o usuário selecionar a pasta e MÚLTIPLOS ANOS
        para a ferramenta de Leitura de PDF (OCR) do Simples.
        """
        # Agora usa GetFolderAndYearsDialog (com 's')
        from app.ferramentas.qt_dialogs import GetFolderAndYearsDialog
        dialog = GetFolderAndYearsDialog("Ler PDFs do Simples (OCR) - Multi Ano", self)
        
        if dialog.exec():
            folder_path, years_list = dialog.get_values()
            
            if folder_path and years_list:
                str_years = ", ".join(map(str, years_list))
                
                self.log_text_edit.clear()
                self.log_text_edit.append(f"--- Iniciando Leitor de PDFs (OCR) ---")
                self.log_text_edit.append(f"Pasta: {folder_path} | Anos: {str_years}")
                self.statusBar().showMessage("Executando Leitura (OCR)...")
                
                # Inicia o worker com a LISTA de anos
                self._start_worker_thread(
                    SimplesReaderWorker, 
                    'tool_ocr', 
                    folder_path, 
                    years_list, # Passa lista
                    on_finished_slot=self.on_simples_reader_finished
                )

    def on_simples_reader_finished(self, output_path):
        """
        Chamado quando o SimplesReaderWorker termina.
        """
        self.workflow_state['tool_ocr'] = 'completed'
        if output_path:
            self.log_text_edit.append(f"✅ Leitura (OCR) concluída. Relatório salvo em: {output_path}")
            QMessageBox.information(self, "Sucesso", f"Relatório de OCR salvo em:\n{output_path}")
        else:
            self.log_text_edit.append("ℹ️ Leitura (OCR) concluída, mas nenhum relatório foi gerado.")
        self.statusBar().showMessage("Leitura (OCR) concluída.", 5000)
        self.update_button_states() # Reseta os botões principais

    def start_simples_downloader_tool(self):
        """
        Inicia a ferramenta de Download (RPA) do Simples.
        Versão Final Corrigida: Compatível com o novo qt_dialogs.py.
        """
        self.log_text_edit.clear()
        
        if self.df_empresas is None or self.df_empresas.empty:
            QMessageBox.warning(self, "Cadastro Vazio", "Por favor, carregue o Ficheiro Mestre primeiro.")
            return

        # 1. Selecionar Pasta Raiz
        base_output_dir = QFileDialog.getExistingDirectory(self, "Selecione a Pasta Raiz para Salvar", "", QFileDialog.Option.ShowDirsOnly)
        if not base_output_dir:
            return

        # 2. Preparar dados para o diálogo
        company_list_for_dialog = []
        for _, row in self.df_empresas.iterrows():
            cnpj_val = str(row.get(Columns.CNPJ, '')).strip()
            name_val = str(row.get(Columns.RAZAO_SOCIAL, '')).strip()
            if cnpj_val:
                company_list_for_dialog.append({'cnpj': cnpj_val, 'name': name_val})

        if not company_list_for_dialog:
            QMessageBox.warning(self, "Aviso", "Nenhum CNPJ encontrado no cadastro.")
            return

        # 3. Mostrar Diálogo
        # O CNPJSelectionDialog agora retorna [{'cnpj': '...', 'dir_path': '...'}, ...]
        dialog = CNPJSelectionDialog(company_list_for_dialog, self)
        if dialog.exec() != QDialog.Accepted:
            self.log_text_edit.append("Seleção cancelada.")
            return

        selected_items = dialog.get_selected_cnpjs()
        
        if not selected_items:
            QMessageBox.warning(self, "Aviso", "Nenhuma empresa foi marcada.")
            return

        # 4. Processar Tarefas
        tasks = []
        self.log_text_edit.append(f"Processando {len(selected_items)} itens selecionados...")
        
        for item in selected_items:
            # Extrai o CNPJ com segurança
            current_cnpj = item.get('cnpj') # Agora a chave 'cnpj' existe graças à correção no qt_dialogs
            
            if not current_cnpj:
                self.log_text_edit.append("⚠️ Item sem CNPJ ignorado.")
                continue

            # Limpa para apenas números (para busca no DF e input no site)
            clean_cnpj_input = re.sub(r'\D', '', str(current_cnpj))
            
            # Tenta recuperar dados completos do DataFrame para criar pasta bonita
            final_name = "Empresa_Sem_Nome"
            final_imu = "SEM_IMU"
            
            # Busca segura no DF
            found_row = None
            try:
                # Tenta match exato
                filtered = self.df_empresas[self.df_empresas[Columns.CNPJ] == current_cnpj]
                if filtered.empty:
                    # Tenta match numérico
                    clean_col = self.df_empresas[Columns.CNPJ].astype(str).str.replace(r'\D', '', regex=True)
                    filtered = self.df_empresas[clean_col == clean_cnpj_input]
                
                if not filtered.empty:
                    found_row = filtered.iloc[0]
                    final_name = str(found_row.get(Columns.RAZAO_SOCIAL, ''))
                    final_imu = str(found_row.get(Columns.IMU, ''))
            except Exception as e:
                print(f"Erro na busca do DF: {e}")

            # Monta caminho da pasta
            try:
                imu_clean = re.sub(r'\D', '', final_imu.split('.')[0])
                if not imu_clean: imu_clean = "SEM_IMU"
                
                cleaned_name = re.sub(r'[^\w\s-]', '', final_name.strip())
                cleaned_name = re.sub(r'[\s-]+', '_', cleaned_name).upper()
                
                folder_name = f"{imu_clean}_{cleaned_name}"
                folder_path = os.path.join(base_output_dir, folder_name)
                os.makedirs(folder_path, exist_ok=True)
                
                tasks.append({'cnpj': clean_cnpj_input, 'folder': folder_path})
                
            except Exception as e:
                self.log_text_edit.append(f"Erro pasta: {e}. Usando raiz.")
                tasks.append({'cnpj': clean_cnpj_input, 'folder': base_output_dir})

        if not tasks:
            QMessageBox.warning(self, "Erro", "Não foi possível gerar tarefas válidas.")
            return

        # 5. Iniciar Worker
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setWindowTitle("Atenção RPA")
        msg.setText(f"Iniciando download para {len(tasks)} empresas.\n\n"
                    "1. Mude para o Chrome AGORA.\n"
                    "2. Não mexa no mouse.")
        msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
        
        if msg.exec() == QMessageBox.StandardButton.Cancel:
            self.log_text_edit.append("Cancelado.")
            return

        self.log_text_edit.append(f"🚀 Iniciando RPA... ({len(tasks)} tarefas)")
        self.statusBar().showMessage("Executando RPA...")

        QTimer.singleShot(5000, lambda: self._start_worker_thread(
            SimplesDownloaderWorker,
            'tool_rpa', 
            tasks, 
            on_finished_slot=self.on_simples_downloader_finished
        ))

    def on_simples_downloader_finished(self):
        """
        Chamado quando o SimplesDownloaderWorker termina.
        """
        self.workflow_state['tool_rpa'] = 'completed'
        self.log_text_edit.append("✅ Download (RPA) concluído.")
        QMessageBox.information(self, "Sucesso", "Download (RPA) concluído.\n\nVocê pode usar o mouse e teclado novamente.")
        self.statusBar().showMessage("Download (RPA) concluído.", 5000)
        self.update_button_states()

    def open_db_extractor_tool(self):
        """
        Abre ferramenta de extração de BD com seleção de empresas, tipos de relatórios
        e MÚLTIPLOS ANOS. Agora inclui suporte a DAS (Simples).
        """
        # 1. Obter Pasta Raiz e Anos (Multi)
        from app.ferramentas.qt_dialogs import GetFolderAndYearsDialog, DbSelectionDialog
        
        dialog_path = GetFolderAndYearsDialog("Extrair Relatórios BD - Multi Ano", self)
        if not dialog_path.exec():
            return

        folder_path, years_list = dialog_path.get_values()
        if not folder_path or not years_list:
            return

        # 2. Ler subpastas para identificar empresas (IMU)
        self.statusBar().showMessage("Lendo pastas...")
        from pathlib import Path
        root = Path(folder_path)
        company_list = []
        
        try:
            for item in root.iterdir():
                if item.is_dir():
                    parts = item.name.split('_', 1)
                    possible_imu = re.sub(r'\D', '', parts[0]) 
                    
                    if possible_imu and len(possible_imu) > 0:
                        company_list.append({
                            'imu': possible_imu,
                            'name': item.name,
                            'path': str(item)
                        })
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler diretório:\n{e}")
            return

        if not company_list:
            QMessageBox.warning(self, "Aviso", "Nenhuma pasta com IMU encontrada.")
            return

        # 3. Mostrar Diálogo de Seleção (Existente)
        sel_dialog = DbSelectionDialog(company_list, self)
        if sel_dialog.exec() != QDialog.Accepted:
            return

        selected_items, do_dams, do_nfse = sel_dialog.get_data()

        if not selected_items:
            QMessageBox.warning(self, "Aviso", "Nenhuma empresa selecionada.")
            return
        
        # 4. Perguntar sobre DAS (Simples Nacional)
        # Como não temos acesso ao código do DbSelectionDialog para adicionar o checkbox lá,
        # perguntamos explicitamente aqui.
        do_das = False
        reply = QMessageBox.question(
            self, 
            "Incluir DAS?", 
            "Deseja extrair também os pagamentos do Simples Nacional (DAS)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            do_das = True

        if not do_dams and not do_nfse and not do_das:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos um tipo de relatório.")
            return

        # 5. Iniciar Worker
        str_years = ", ".join(map(str, years_list))
        
        self.log_text_edit.clear()
        self.log_text_edit.append(f"--- Iniciando Extrator BD ---")
        self.log_text_edit.append(f"Empresas: {len(selected_items)} | Anos: {str_years}")
        self.log_text_edit.append(f"Modo: {'DAMS ' if do_dams else ''}{'NFSE ' if do_nfse else ''}{'DAS ' if do_das else ''}")
        self.statusBar().showMessage("Executando Extração do BD...")
        
        # Passa a LISTA de anos para o worker
        # ⚠️ IMPORTANTE: O DatabaseExtractorWorker deve aceitar o argumento extra 'do_das'.
        # Se o worker apenas repassar *args para run_db_extraction, isso funcionará automaticamente.
        self._start_worker_thread(
            DatabaseExtractorWorker, 
            'tool_db', 
            selected_items, 
            years_list,     
            do_dams,        
            do_nfse,
            do_das,         # <--- Novo Argumento
            on_finished_slot=self.on_db_extractor_finished
        )

    def on_db_extractor_finished(self):
        """
        Chamado quando o DatabaseExtractorWorker termina.
        """
        self.workflow_state['tool_db'] = 'completed'
        self.log_text_edit.append("✅ Extração do Banco de Dados concluída.")
        QMessageBox.information(self, "Sucesso", "Extração do Banco de Dados concluída.")
        self.statusBar().showMessage("Extração do BD concluída.", 5000)
        self.update_button_states()

    def start_situacao_extractor_tool(self):
        """
        Inicia a ferramenta de Extração de Situação (pywinauto).
        Lógica:
        1. Lê as pastas.
        2. Pega o IMU do nome da pasta (antes do primeiro '_').
        3. Formata para 9 dígitos (zeros à esquerda).
        4. Envia esse IMU para ser digitado no robô.
        """
        # 1. Selecionar Pasta Raiz
        root_folder = QFileDialog.getExistingDirectory(self, "1. Selecione a Pasta Raiz com Subdiretórios (IMU_NOME)", "", QFileDialog.Option.ShowDirsOnly)
        if not root_folder:
            return

        # 2. Ler pastas e preparar dados
        data_list_for_dialog = []
        tasks_missing = 0
        
        self.log_text_edit.append("--- Lendo pastas para extração de IMU ---")

        try:
            from pathlib import Path
            base_path = Path(root_folder)
            
            for item in base_path.iterdir():
                if item.is_dir():
                    dir_name = item.name
                    # Extrai IMU da pasta (ex: "12345_EMPRESA") -> "12345"
                    imu_raw_part = dir_name.split('_', 1)[0].strip()
                    
                    # Remove caracteres não numéricos apenas por segurança
                    imu_numeric = re.sub(r'\D', '', imu_raw_part)
                    
                    if imu_numeric:
                        # REGRA: 9 dígitos, completar com 0 antes
                        imu_formatted = imu_numeric.zfill(9)
                        
                        # Verifica se o PDF já existe (para autoselect)
                        expected_filename = f"Situacao_{imu_formatted}.pdf"
                        expected_path = item / expected_filename
                        
                        is_selected = not expected_path.exists()
                        if is_selected:
                            tasks_missing += 1

                        data_list_for_dialog.append({
                            'cnpj': imu_formatted,   # Usamos a chave 'cnpj' para exibir no diálogo (coluna ID)
                            'imu_id': imu_formatted, # Chave explícita para o worker
                            'name': dir_name,
                            'dir_path': str(item),
                            'is_selected': is_selected
                        })
                    else:
                        self.log_text_edit.append(f"⚠️ Pasta '{dir_name}' ignorada: Não inicia com números.")
                        
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao processar pastas:\n{e}")
            return

        if not data_list_for_dialog:
            QMessageBox.warning(self, "Sem Pastas Válidas", "Nenhuma pasta com padrão 'NUMERO_NOME' encontrada.")
            return

        # 3. Mostrar Diálogo de Seleção
        dialog = CNPJSelectionDialog(data_list_for_dialog, self)
        dialog.setWindowTitle(f"2. Confirmar Extração ({tasks_missing} Faltantes)")
        
        if dialog.exec() != QDialog.Accepted:
            return

        tasks = dialog.get_selected_cnpjs() 
        
        if not tasks:
            return
            
        # 4. Iniciar Worker
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Warning)
        msg_box.setWindowTitle("Atenção RPA")
        msg_box.setText(f"Iniciando extração via IMU para {len(tasks)} pastas.\n\n"
                        "1. Certifique-se de que o GTM/Comércio está ABERTO.\n"
                        "2. Não use o mouse.")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
        
        if msg_box.exec() == QMessageBox.StandardButton.Cancel:
            return

        self.log_text_edit.append(f"--- Iniciando RPA (Modo IMU: 9 dígitos) ---")
        
        # Inicia o worker com um pequeno delay para o usuário trocar de janela se precisar
        QTimer.singleShot(2000, lambda: self._start_worker_thread(
            SituacaoExtractorWorker,
            'tool_situacao_rpa', 
            tasks, 
            on_finished_slot=self.on_situacao_extractor_finished
        ))

    def on_situacao_extractor_finished(self):
        """
        Chamado quando o SituacaoExtractorWorker termina.
        """
        self.workflow_state['tool_situacao_rpa'] = 'completed'
        self.log_text_edit.append("✅ Extração de Situação (RPA) concluída.")
        QMessageBox.information(self, "Sucesso", "Extração de Situação (RPA) concluída.\n\nVocê pode usar o mouse e teclado novamente.")
        self.statusBar().showMessage("Extração de Situação (RPA) concluída.", 5000)
        self.update_button_states()
        
    # --- ✅ FIM: NOVAS FUNÇÕES DAS FERRAMENTAS ---

    def export_custom_texts_logic(self):
        f, _ = QFileDialog.getSaveFileName(self, "Export", f"auditapp_textos_{datetime.now().strftime('%Y%m%d')}.json", "JSON (*.json)")
        if f: 
            try: 
                with open(f,'w',encoding='utf-8') as file: json.dump({"general_texts": get_custom_general_texts(), "auto_texts": get_custom_auto_texts()}, file, indent=4, ensure_ascii=False)
                QMessageBox.information(self, "Sucesso", "Exportado.")
            except Exception as e: QMessageBox.critical(self, "Erro", str(e))

    def import_custom_texts_logic(self):
        f, _ = QFileDialog.getOpenFileName(self, "Import", "", "JSON (*.json)")
        if f:
            try:
                with open(f,'r',encoding='utf-8') as file: d = json.load(file)
                set_custom_general_texts(d["general_texts"]); set_custom_auto_texts(d["auto_texts"])
                QMessageBox.information(self, "Sucesso", "Importado.")
            except Exception as e: QMessageBox.critical(self, "Erro", str(e))

    def open_activity_log(self):
        p = "controle_de_atividades.xlsx"
        if os.path.exists(p): QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(p)))
        else: QMessageBox.warning(self, "Aviso", "Ainda não existe.")

class NewsDialog(QDialog):
    def __init__(self, news_content, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📰 Informativo Caronte")
        self.resize(500, 400)
        
        layout = QVBoxLayout(self)
        
        # Title/Header
        title = QLabel("Atualizações & Avisos")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #88C0D0; margin-bottom: 5px;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Content Area
        self.text_browser = QTextEdit()
        self.text_browser.setReadOnly(True)
        self.text_browser.setHtml(f"<div style='font-size: 14px; color: #ECEFF4;'>{news_content.replace(chr(10), '<br>')}</div>")
        self.text_browser.setStyleSheet("""
            QTextEdit {
                background-color: #3B4252;
                border: 1px solid #4C566A;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        layout.addWidget(self.text_browser)

        # Close Button
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        ok_btn = QPushButton("Entendido")
        ok_btn.setCursor(Qt.PointingHandCursor)
        ok_btn.setStyleSheet("""
            QPushButton {
                background-color: #5E81AC;
                color: white;
                font-weight: bold;
                padding: 8px 15px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #81A1C1; }
        """)
        ok_btn.clicked.connect(self.accept)
        btn_box.addWidget(ok_btn)
        btn_box.addStretch()
        
        layout.addLayout(btn_box)

class ScriptViewerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent); self.setWindowTitle("Script Chrome"); self.resize(800,600); l=QVBoxLayout(self)
        l.addWidget(QLabel("Copie e cole no Console do Chrome (F12) na página do ISS:"))
        self.t = QPlainTextEdit(); self.t.setReadOnly(True); self.t.setFont(QFont("Consolas")); l.addWidget(self.t)
        self.t.setPlainText(r"""async function exportDataToCSVFile() {
    function waitForElement(selector, timeout = 5000) {
        return new Promise((resolve, reject) => {
            const intervalTime = 100;
            let timeWaited = 0;
            const interval = setInterval(() => {
                const element = document.querySelector(selector);
                if (element) {
                    clearInterval(interval);
                    resolve(element);
                } else {
                    timeWaited += intervalTime;
                    if (timeWaited >= timeout) {
                        clearInterval(interval);
                        reject(new Error(`Element "${selector}" not found after ${timeout}ms`));
                    }
                }
            }, intervalTime);
        });
    }
    const allRecords = [];
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    const confirmedPaymentsPanel = document.getElementById('pnlDocEmitPagConf');
    if (!confirmedPaymentsPanel) {
        console.error('The main panel with ID "pnlDocEmitPagConf" was not found.');
        return;
    }
    const mainTableRows = confirmedPaymentsPanel.querySelectorAll('tbody > tr[role="row"]');
    for (const row of mainTableRows) {
        const record = {};
        const detailButton = row.querySelector('button[data-action="detalhes"]');
        if (detailButton) {
            detailButton.click();
            await delay(1000);
            record.codigoVerificacao = document.getElementById('IdentificacaoBaixa')?.value || '';
            record.totalRecolher = document.getElementById('ValorImposto')?.value || '';
            record.referenciaPagamento = document.getElementById('ReferenciaPagamento')?.value || '';
            record.receita = document.getElementById('Receita')?.value || '';
            record.desconto = document.getElementById('ValorDesconto')?.value || '';
            record.tributo = document.getElementById('Tributo')?.value || '';
            const closeButton = document.querySelector('div.modal.show button[title="Fechar"]');
            if (closeButton) { closeButton.click(); await delay(500); }
        }
        const notasButton = row.querySelector('button[data-action="notas"]');
        record.numerosDasNotas = '';
        if (notasButton) {
            notasButton.click();
            try {
                await waitForElement('#tblDocsDam');
                const invoiceDetailRows = document.querySelectorAll('#tblDocsDam tbody tr');
                const invoiceNumbers = [];
                for (const invoiceRow of invoiceDetailRows) {
                    const numberCell = invoiceRow.querySelector('td:nth-child(3)');
                    if (numberCell) invoiceNumbers.push(numberCell.textContent.trim());
                }
                record.numerosDasNotas = invoiceNumbers.join('- ');
            } catch (error) {
                console.warn(`Could not find invoice table for a row or it has no invoices.`);
                record.numerosDasNotas = "N/A"; 
            }
            notasButton.click(); await delay(500);
        }
        if (Object.keys(record).length > 0) allRecords.push(record);
    }
    if (allRecords.length > 0) {
        const headers = Object.keys(allRecords[0]);
        const csvRows = [headers.join(',')];
        for (const record of allRecords) {
            const values = headers.map(header => {
                const value = record[header] || '';
                const escaped = ('' + value).replace(/"/g, '""');
                return `"${escaped}"`;
            });
            csvRows.push(values.join(','));
        }
        const csvString = csvRows.join('\n');
        const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'dados_extraidos.csv';
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        console.log(`✅ Success! The file "dados_extraidos.csv" with ${allRecords.length} records should be in your Downloads folder.`);
    } else { console.log('No data was extracted.'); }
}
exportDataToCSVFile();""") 
        h=QHBoxLayout(); b=QPushButton("Copiar"); b.clicked.connect(lambda: (QApplication.clipboard().setText(self.t.toPlainText()), QMessageBox.information(self,"Copiado","Copiado!"))); h.addWidget(b); h.addStretch(); c=QPushButton("Fechar"); c.clicked.connect(self.accept); h.addWidget(c); l.addLayout(h)
