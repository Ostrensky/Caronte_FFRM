# --- FILE: app/activity_review_dialog.py ---
import pandas as pd
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QLabel, QPushButton, QSplitter, QTextEdit, QHeaderView, QComboBox,
    QMessageBox, QGroupBox, QAbstractItemView, QMenu, QCheckBox,
    QWidget, QLineEdit, QFormLayout, QSizePolicy  # <--- Added QSizePolicy
)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QColor, QBrush, QAction, QFont

# --- New Imports for Enhanced Table ---
from app.excel_filter import FilterableHeaderView
from app.widgets import NumericTableWidgetItem, DateTableWidgetItem, SORT_ROLE
from app.constants import Columns

# Standard codes taxed at "Local da Prestação" (LC 116/03 - Art. 3 Exceções)
LOCAL_PRESTACAO_CODES = [
    '0305', '0702', '0704', '0705', '0709', '0710', '0711', 
    '0712', '0716', '0717', '0719', '1101', '1102', '1104', 
    '1705', '1710'
]

class ActivityReviewDialog(QDialog):
    """
    Dialog allowing the user to:
    1. See a summary of all Activity Codes in the invoices.
    2. Review full Service Descriptions (Master-Detail view).
    3. Bulk change Activity Codes.
    4. Flag invoices as 'Local Tomador' (Manual Status).
    5. Filter and Sort using Excel-like tools AND Advanced Logic Filters.
    """
    def __init__(self, invoices_df, activity_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Revisão de Atividades e Local da Prestação")
        self.resize(1300, 850)
        
        # Working copy of the dataframe
        self.df = invoices_df
        self.activity_data = activity_data # Dictionary from main.py
        
        # Ensure status column exists
        if 'status_manual' not in self.df.columns:
            self.df['status_manual'] = ''

        # Helper list for ComboBoxes (Code - Description)
        # We store the FULL text here
        self.code_list = []
        if self.activity_data:
            for code, details in self.activity_data.items():
                desc = details[0][0] if details else "Sem Descrição"
                self.code_list.append(f"{code} - {desc}")
            self.code_list.sort()

        # State for Table
        self.current_summary_code = None 
        self.current_display_df = self.df.copy() 
        
        # Filter State
        self.active_filters = [] 
        
        # Default Columns to Show
        self.visible_columns = [
            'DATA EMISSÃO', 
            'NÚMERO', 
            'DISCRIMINAÇÃO DOS SERVIÇOS', 
            'CÓDIGO DA ATIVIDADE', 
            'VALOR', 
            'status_manual'
        ]

        self._setup_ui()
        self.load_summary_table()
        # Initial Load (All Data)
        self.load_invoice_data(filter_code=None)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # --- 1. Top Section: Activity Summary ---
        summary_group = QGroupBox("Resumo por Código de Atividade (Clique para filtrar)")
        summary_layout = QVBoxLayout(summary_group)
        
        self.summary_table = QTableWidget()
        self.summary_table.setColumnCount(5)
        self.summary_table.setHorizontalHeaderLabels([
            "Código", "Descrição (Excel)", "Qtd Notas", "Valor Total", "Status Local"
        ])
        self.summary_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.summary_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.summary_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.summary_table.cellClicked.connect(self.on_summary_row_clicked)
        self.summary_table.setAlternatingRowColors(True)
        self.summary_table.setMaximumHeight(200)
        
        summary_layout.addWidget(self.summary_table)
        main_layout.addWidget(summary_group)

        # --- 2. Action Toolbar ---
        action_layout = QHBoxLayout()
        
        self.btn_reset_filter = QPushButton("Mostrar Todas as Notas")
        self.btn_reset_filter.clicked.connect(self.reset_summary_filter)
        
        self.btn_bulk_change = QPushButton("📝 Alterar Código (Selecionadas)")
        self.btn_bulk_change.clicked.connect(self.bulk_change_activity)
        self.btn_bulk_change.setStyleSheet("background-color: #5E81AC; color: white; font-weight: bold;")
        
        self.btn_set_local_tomador = QPushButton("📍 Definir como 'Local Tomador'")
        self.btn_set_local_tomador.setToolTip("Notas marcadas como Local Tomador serão ignoradas na geração de Autos/IDD.")
        self.btn_set_local_tomador.clicked.connect(self.set_local_tomador)
        self.btn_set_local_tomador.setStyleSheet("background-color: #BF616A; color: white;")

        self.btn_clear_status = QPushButton("🔄 Limpar Status Manual")
        self.btn_clear_status.clicked.connect(self.clear_manual_status)

        action_layout.addWidget(self.btn_reset_filter)
        action_layout.addStretch()
        action_layout.addWidget(self.btn_bulk_change)
        action_layout.addWidget(self.btn_set_local_tomador)
        action_layout.addWidget(self.btn_clear_status)
        
        main_layout.addLayout(action_layout)

        # --- 3. Splitter Section: Invoice List + Description Viewer ---
        splitter = QSplitter(Qt.Vertical)
        
        # A. Invoice List Container
        invoice_container = QWidget()
        invoice_layout = QVBoxLayout(invoice_container)
        invoice_layout.setContentsMargins(0, 0, 0, 0)

        # --- Advanced Filter Controls ---
        filter_control_layout = QHBoxLayout()
        self.btn_add_filter = QPushButton("Adicionar Filtro (Contém / Não Contém)")
        self.btn_add_filter.clicked.connect(self.add_filter_row)
        filter_control_layout.addWidget(self.btn_add_filter)
        filter_control_layout.addStretch()
        
        self.filters_container = QWidget()
        self.filters_layout = QVBoxLayout(self.filters_container)
        self.filters_layout.setContentsMargins(0, 0, 0, 0)
        self.filters_layout.setSpacing(2)

        invoice_layout.addLayout(filter_control_layout)
        invoice_layout.addWidget(self.filters_container)
        
        # --- The Table ---
        self.invoice_table = QTableWidget()
        self.invoice_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.invoice_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.invoice_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.invoice_table.setAlternatingRowColors(True)
        self.invoice_table.setSortingEnabled(True)
        
        # Apply Custom Filterable Header (Excel-style)
        header = FilterableHeaderView(self.invoice_table)
        header.filter_changed.connect(self.apply_excel_filters)
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self.show_column_context_menu)
        self.invoice_table.setHorizontalHeader(header)
        
        self.invoice_table.itemSelectionChanged.connect(self.on_invoice_selection_changed)
        
        # Enable Right-Click Context Menu on Rows
        self.invoice_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.invoice_table.customContextMenuRequested.connect(self.open_context_menu)

        invoice_layout.addWidget(self.invoice_table)
        splitter.addWidget(invoice_container)

        # B. Full Description Viewer
        viewer_group = QGroupBox("Discriminação dos Serviços (Texto Completo)")
        viewer_layout = QVBoxLayout(viewer_group)
        self.desc_viewer = QTextEdit()
        self.desc_viewer.setReadOnly(True)
        self.desc_viewer.setPlaceholderText("Selecione uma nota acima para ver o texto completo aqui...")
        self.desc_viewer.setStyleSheet("font-size: 11pt; background-color: #ECEFF4; color: #2E3440;")
        viewer_layout.addWidget(self.desc_viewer)
        
        splitter.addWidget(viewer_group)
        splitter.setSizes([500, 150]) # Initial ratio

        main_layout.addWidget(splitter)
        
        # Save Button
        self.btn_save = QPushButton("Salvar e Fechar")
        self.btn_save.clicked.connect(self.accept)
        self.btn_save.setStyleSheet("padding: 10px; font-weight: bold;")
        main_layout.addWidget(self.btn_save)

    # --- ADVANCED FILTER LOGIC ---

    def add_filter_row(self):
        """Adds a new row of filter controls."""
        filter_row_layout = QHBoxLayout()
        column_combo = QComboBox()
        column_combo.addItems(self.visible_columns)
        logic_combo = QComboBox()
        logic_combo.addItems(["Contém", "Não Contém"])
        logic_combo.setFixedWidth(110)
        filter_edit = QLineEdit()
        filter_edit.setPlaceholderText("Texto para filtrar...")
        remove_btn = QPushButton("X")
        remove_btn.setFixedWidth(30)
        remove_btn.setStyleSheet("color: red; font-weight: bold;")
        
        filter_row_layout.addWidget(column_combo, 1)
        filter_row_layout.addWidget(logic_combo, 0)
        filter_row_layout.addWidget(filter_edit, 2)
        filter_row_layout.addWidget(remove_btn)
        
        self.filters_layout.addLayout(filter_row_layout)
        
        filter_widgets = { 
            "layout": filter_row_layout, "combo": column_combo, 
            "logic": logic_combo, "edit": filter_edit, "button": remove_btn
        }
        self.active_filters.append(filter_widgets)
        
        column_combo.currentIndexChanged.connect(self.refresh_table)
        logic_combo.currentIndexChanged.connect(self.refresh_table) 
        filter_edit.textChanged.connect(self.refresh_table)
        remove_btn.clicked.connect(lambda: self._remove_filter_row(filter_widgets))

    def _remove_filter_row(self, filter_widgets):
        if filter_widgets in self.active_filters:
            self.active_filters.remove(filter_widgets)
        for i in reversed(range(filter_widgets["layout"].count())):    
            widget = filter_widgets["layout"].itemAt(i).widget()
            if widget is not None: widget.deleteLater()
        filter_widgets["layout"].deleteLater()
        self.refresh_table()

    def _apply_filters_to_df(self, df):
        if df is None or df.empty or not self.active_filters: return df
        filtered_df = df.copy()
        for f in self.active_filters:
            column_name = f["combo"].currentText()
            filter_text = f["edit"].text().strip().lower()
            logic_mode = f["logic"].currentText()
            if filter_text and column_name in filtered_df.columns:
                try:
                    col_series = filtered_df[column_name].astype(str).str.lower()
                    mask = col_series.str.contains(filter_text, na=False, regex=False)
                    if logic_mode == "Contém": filtered_df = filtered_df[mask]
                    else: filtered_df = filtered_df[~mask]
                except Exception: pass
        return filtered_df

    # --- TABLE LOADING LOGIC ---

    def load_summary_table(self):
        self.summary_table.setRowCount(0)
        if 'CÓDIGO DA ATIVIDADE' not in self.df.columns: return
        summary = self.df.groupby('CÓDIGO DA ATIVIDADE').agg({
            'VALOR': 'sum', 'NÚMERO': 'count'
        }).reset_index().sort_values('VALOR', ascending=False)
        self.summary_table.setRowCount(len(summary))
        
        for row_idx, row_data in summary.iterrows():
            code = str(row_data['CÓDIGO DA ATIVIDADE'])
            count = row_data['NÚMERO']
            value = row_data['VALOR']
            desc = "Não encontrado no Excel"
            if self.activity_data and code in self.activity_data:
                desc = self.activity_data[code][0][0]
            is_local = code in LOCAL_PRESTACAO_CODES
            local_status = "📍 LOCAL PRESTAÇÃO" if is_local else "Normal"
            items = [
                QTableWidgetItem(code), QTableWidgetItem(desc),
                QTableWidgetItem(str(count)), QTableWidgetItem(f"R$ {value:,.2f}"),
                QTableWidgetItem(local_status)
            ]
            if is_local:
                color = QColor("#D08770")
                for it in items: it.setBackground(color); it.setForeground(QColor("white"))
            for c, it in enumerate(items): self.summary_table.setItem(row_idx, c, it)

    def load_invoice_data(self, filter_code=None):
        self.current_summary_code = filter_code
        self.refresh_table()

    def refresh_table(self):
        self.invoice_table.blockSignals(True)
        self.invoice_table.setSortingEnabled(False)
        if self.current_summary_code:
            temp_df = self.df[self.df['CÓDIGO DA ATIVIDADE'] == self.current_summary_code].copy()
        else:
            temp_df = self.df.copy()

        self.current_display_df = self._apply_filters_to_df(temp_df)
        self.invoice_table.column_mapping = self.visible_columns
        self.invoice_table.horizontalHeader().set_dataframe(self.current_display_df)
        self.invoice_table.setColumnCount(len(self.visible_columns))
        self.invoice_table.setHorizontalHeaderLabels(self.visible_columns)
        self.invoice_table.setRowCount(len(self.current_display_df))
        
        row_idx = 0
        for original_idx, row in self.current_display_df.iterrows():
            for col_idx, col_name in enumerate(self.visible_columns):
                raw_val = row.get(col_name, '')
                item = None
                if col_name in ['VALOR', 'ALÍQUOTA', 'correct_rate', 'VALOR DEDUÇÃO']:
                    display_text = f"{raw_val:,.2f}" if pd.notna(raw_val) and isinstance(raw_val, (int, float)) else str(raw_val)
                    item = NumericTableWidgetItem(display_text)
                    item.setData(SORT_ROLE, float(raw_val) if pd.notna(raw_val) else 0.0)
                elif col_name == 'DATA EMISSÃO':
                    display_text = ""; sort_val = pd.NaT
                    if pd.notna(raw_val):
                        if hasattr(raw_val, 'strftime'): display_text = raw_val.strftime('%d/%m/%Y'); sort_val = raw_val
                        else: display_text = str(raw_val)
                    item = DateTableWidgetItem(display_text)
                    item.setData(SORT_ROLE, sort_val)
                else:
                    val_str = str(raw_val) if pd.notna(raw_val) else ""
                    display_str = val_str[:100] + "..." if col_name == 'DISCRIMINAÇÃO DOS SERVIÇOS' and len(val_str) > 100 else val_str
                    item = QTableWidgetItem(display_str)
                    item.setData(Qt.UserRole + 1, val_str) 

                if col_idx == 0: item.setData(Qt.UserRole, original_idx)
                if str(row.get('status_manual', '')) == 'Local_Tomador':
                    item.setForeground(QColor("#BF616A"))
                self.invoice_table.setItem(row_idx, col_idx, item)
            row_idx += 1
        self.invoice_table.setSortingEnabled(True)
        self.invoice_table.blockSignals(False)
        self.invoice_table.resizeColumnsToContents()
        try:
            desc_idx = self.visible_columns.index('DISCRIMINAÇÃO DOS SERVIÇOS')
            if self.invoice_table.columnWidth(desc_idx) > 400: self.invoice_table.setColumnWidth(desc_idx, 400)
        except: pass

    # --- ACTIONS & EVENTS ---

    def apply_excel_filters(self, col_index, allowed_values):
        header = self.invoice_table.horizontalHeader()
        active_filters = header.filters 
        for row in range(self.invoice_table.rowCount()):
            should_show = True
            for f_col_idx, allowed in active_filters.items():
                col_name = self.visible_columns[f_col_idx]
                item_0 = self.invoice_table.item(row, 0)
                if not item_0: continue
                original_idx = item_0.data(Qt.UserRole)
                try:
                    raw_val = self.current_display_df.at[original_idx, col_name]
                    val_str = str(raw_val) if pd.notna(raw_val) else ""
                except: val_str = ""
                if val_str not in allowed: should_show = False; break
            self.invoice_table.setRowHidden(row, not should_show)

    def show_column_context_menu(self, pos):
        menu = QMenu(self)
        all_cols = sorted(list(self.df.columns))
        for col in all_cols:
            action = QAction(col, self)
            action.setCheckable(True)
            action.setChecked(col in self.visible_columns)
            action.triggered.connect(lambda checked, c=col: self.toggle_column(checked, c))
            menu.addAction(action)
        menu.exec(self.invoice_table.mapToGlobal(pos))

    def toggle_column(self, checked, col_name):
        if checked:
            if col_name not in self.visible_columns: self.visible_columns.append(col_name)
        else:
            if col_name in self.visible_columns: self.visible_columns.remove(col_name)
        self.refresh_table()

    def on_summary_row_clicked(self, row, col):
        code = self.summary_table.item(row, 0).text()
        self.load_invoice_data(filter_code=code)
        self.desc_viewer.clear()

    def reset_summary_filter(self):
        self.summary_table.clearSelection()
        self.load_invoice_data(filter_code=None)
        self.desc_viewer.clear()

    def on_invoice_selection_changed(self):
        selected_rows = self.invoice_table.selectionModel().selectedRows()
        if not selected_rows: return
        last_row = selected_rows[-1].row()
        item = self.invoice_table.item(last_row, 0)
        if item:
            original_idx = item.data(Qt.UserRole)
            full_desc = self.df.loc[original_idx, 'DISCRIMINAÇÃO DOS SERVIÇOS']
            self.desc_viewer.setPlainText(str(full_desc))

    def get_selected_indices(self):
        selected_rows = self.invoice_table.selectionModel().selectedRows()
        indices = []
        for row_obj in selected_rows:
            r = row_obj.row()
            if not self.invoice_table.isRowHidden(r):
                item = self.invoice_table.item(r, 0)
                if item: indices.append(item.data(Qt.UserRole))
        return indices

    def set_local_tomador(self):
        indices = self.get_selected_indices()
        if not indices: QMessageBox.warning(self, "Aviso", "Selecione notas."); return
        self.df.loc[indices, 'status_manual'] = 'Local_Tomador'
        self.refresh_table() 
        QMessageBox.information(self, "Sucesso", f"{len(indices)} notas marcadas.")

    def clear_manual_status(self):
        indices = self.get_selected_indices()
        if not indices: QMessageBox.warning(self, "Aviso", "Selecione notas."); return
        self.df.loc[indices, 'status_manual'] = ''
        self.refresh_table()

    def bulk_change_activity(self):
        indices = self.get_selected_indices()
        if not indices: 
            QMessageBox.warning(self, "Aviso", "Selecione notas para alterar.")
            return

        # --- Improved Dialog Resolution and Styling ---
        dialog = QDialog(self)
        dialog.setWindowTitle("Alterar Código de Atividade")
        
        # ✅ FIX 1: Strict Fixed Width
        dialog.setFixedWidth(600)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        lbl_info = QLabel(f"Selecione o novo código para <b>{len(indices)}</b> notas selecionadas:")
        lbl_info.setFont(QFont("Arial", 11))
        
        # ✅ FIX 2: Truncate long descriptions to prevent horizontal expansion
        display_items = []
        for item in self.code_list:
            if len(item) > 100:
                display_items.append(item[:100] + "...")
            else:
                display_items.append(item)

        combo = QComboBox()
        combo.addItems(display_items)
        
        # ✅ FIX 3: Ignore content width, adhere to the fixed dialog width
        combo.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        
        combo.setStyleSheet("""
            QComboBox { padding: 8px; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; }
            QComboBox::drop-down { border: 0px; }
        """)
        
        btn_apply = QPushButton("Aplicar Mudança")
        btn_apply.setCursor(Qt.PointingHandCursor)
        btn_apply.setStyleSheet("""
            QPushButton { 
                background-color: #5E81AC; color: white; font-weight: bold; 
                font-size: 14px; padding: 10px; border-radius: 4px;
            }
            QPushButton:hover { background-color: #81A1C1; }
        """)
        btn_apply.clicked.connect(dialog.accept)

        layout.addWidget(lbl_info)
        layout.addWidget(combo)
        layout.addWidget(btn_apply)
        layout.addStretch()
        
        if dialog.exec():
            # Even if truncated, the code is at the start "1234 - Desc...", so split still works
            new_code = combo.currentText().split(' - ')[0].strip()
            
            # Update the MAIN DataFrame
            self.df.loc[indices, 'CÓDIGO DA ATIVIDADE'] = new_code
            
            if self.activity_data and new_code in self.activity_data:
                self.df.loc[indices, 'activity_desc'] = self.activity_data[new_code][0][0]
                self.df.loc[indices, 'correct_rate'] = self.activity_data[new_code][0][1]
            
            # Real-time update
            self.load_summary_table()
            self.refresh_table()
            QMessageBox.information(self, "Sucesso", "Atividades atualizadas com sucesso.")

    def open_context_menu(self, position):
        menu = QMenu()
        action_tomador = QAction("📍 Definir como Local Tomador", self)
        action_tomador.triggered.connect(self.set_local_tomador)
        menu.addAction(action_tomador)
        action_change = QAction("📝 Alterar Código de Atividade", self)
        action_change.triggered.connect(self.bulk_change_activity)
        menu.addAction(action_change)
        menu.exec(self.invoice_table.viewport().mapToGlobal(position))

    def get_updated_dataframe(self):
        return self.df