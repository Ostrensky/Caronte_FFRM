# --- FILE: app/relabeling_dialog.py ---

import pandas as pd
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QGroupBox, QHBoxLayout, 
                               QLabel, QComboBox, QPushButton, QTableWidget, 
                               QTableWidgetItem, QHeaderView, QCompleter,
                               QCheckBox, QTextEdit, QSplitter) # ✅ Import QTextEdit and QSplitter
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QColor

from .constants import Columns

class RelabelingWindow(QDialog):
    def __init__(self, df_invoices, activity_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Revisão de Atividade e Localização") # ✅ New Title
        self.setMinimumSize(1100, 650)
        self.setWindowState(Qt.WindowState.WindowMaximized)
        self.df = df_invoices.copy()
        self.activity_data = activity_data
        self.tomador_indices = [] # ✅ For Request 3
        
        main_layout = QVBoxLayout(self)
        
        # --- UI Setup ---
        controls_group = QGroupBox("Ações de Reclassificação")
        controls_layout = QHBoxLayout()
        controls_layout.addWidget(QLabel("Nova Categoria para Selecionados:"))
        
        self.activity_combo = QComboBox()
        self.activity_combo.setEditable(True)
        self.activity_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.activity_combo.lineEdit().setPlaceholderText("Digite para procurar...")

        for code, activities in sorted(self.activity_data.items()):
            for description, aliquot, _ in activities:
                display_text = f"{code} - [{aliquot:.2f}%] {description}"
                self.activity_combo.addItem(display_text, userData=(code, description))
        
        completer = QCompleter(self.activity_combo.model(), self)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.activity_combo.setCompleter(completer)
        
        controls_layout.addWidget(self.activity_combo, 1)
        relabel_button = QPushButton("Aplicar Nova Categoria")
        relabel_button.clicked.connect(self.relabel_selected)
        controls_layout.addWidget(relabel_button)
        
        self.tomador_button = QPushButton("Marcar como 'Local Tomador' (Não Tributar)")
        self.tomador_button.setStyleSheet("background-color: #BF616A; color: #ECEFF4;") # Reddish color
        self.tomador_button.clicked.connect(self._mark_as_tomador)
        controls_layout.addWidget(self.tomador_button)
        
        controls_group.setLayout(controls_layout)
        main_layout.addWidget(controls_group)
        
        # ✅ --- START: Filter Checkbox Update ---
        filter_layout = QHBoxLayout()
        self.filter_checkbox = QCheckBox("Mostrar apenas faturas com alertas (Local ou Atividade)")
        self.filter_checkbox.setChecked(True) # Default to filtered view
        self.filter_checkbox.stateChanged.connect(self.populate_table)
        filter_layout.addWidget(self.filter_checkbox)
        filter_layout.addStretch()
        main_layout.addLayout(filter_layout)
        # ✅ --- END: Filter Checkbox Update ---

        # ✅ --- START: Splitter (Table + Full Text) ---
        splitter = QSplitter(Qt.Orientation.Vertical)
        
        self.table = QTableWidget()
        self.activity_desc_col = 'activity_desc'
        
        # ✅ --- START: Add New Column ---
        self.columns = [
            Columns.INVOICE_NUMBER, Columns.SERVICE_DESCRIPTION, "CÓDIGO ORIGINAL",
            self.activity_desc_col.replace('_', ' ').title(),
            "ALÍQUOTA CORRETA (%)", 
            "ALERTA DE ATIVIDADE", # <-- NEW
            Columns.LOCATION_ALERT
        ]
        # ✅ --- END: Add New Column ---
        
        self.table.setColumnCount(len(self.columns))
        self.table.setHorizontalHeaderLabels(self.columns)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTriggers.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        
        self.table.itemSelectionChanged.connect(self._show_details_for_selected_row)
        
        splitter.addWidget(self.table)
        # ✅ --- END: Splitter (Table + Full Text) ---

        details_group = QGroupBox("Discriminação Completa do Serviço (Nota Selecionada)")
        details_layout = QVBoxLayout()
        self.details_text_edit = QTextEdit()
        self.details_text_edit.setReadOnly(True)
        self.details_text_edit.setFontPointSize(12)
        details_layout.addWidget(self.details_text_edit)
        details_group.setLayout(details_layout)
        
        splitter.addWidget(details_group)
        splitter.setSizes([600, 300]) # Set initial sizes (600px for table, 300px for text)
        main_layout.addWidget(splitter, 1) # Add splitter to main layout

        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()
        save_button = QPushButton("Salvar Alterações e Fechar")
        save_button.clicked.connect(self.accept)
        bottom_layout.addWidget(save_button)
        main_layout.addLayout(bottom_layout)
        
        self.populate_table()

    def _show_details_for_selected_row(self):
        """Populates the text box with the full description."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            self.details_text_edit.clear()
            return
            
        # Get the first selected row
        first_row_model_index = selected_rows[0]
        original_index = self.table.item(first_row_model_index.row(), 0).data(Qt.ItemDataRole.UserRole)
        
        # Get full description from the original DataFrame
        full_description = self.df.loc[original_index].get(Columns.SERVICE_DESCRIPTION, "Descrição não encontrada.")
        self.details_text_edit.setText(full_description)

    def populate_table(self):
        self.table.setSortingEnabled(False)
        self.table.clearContents()
        self.details_text_edit.clear()
        
        # --- Filter Logic ---
        df_to_show = self.df
        
        # Exclude rows already marked as 'Local Tomador'
        df_to_show = df_to_show[~df_to_show.index.isin(self.tomador_indices)]
        
        if self.filter_checkbox.isChecked():
            # ✅ --- START: Update Filter Logic ---
            mask_loc_alert = df_to_show.get(Columns.LOCATION_ALERT, '').astype(str).str.strip() != ''
            mask_act_alert = df_to_show.get('activity_alert', '').astype(str).str.strip() != ''
            final_mask = mask_loc_alert | mask_act_alert
            df_to_show = df_to_show[final_mask]
            # ✅ --- END: Update Filter Logic ---

        self.table.setRowCount(df_to_show.shape[0])
        alert_color = QColor(100, 92, 63) # Use the same color for both alerts

        for row_idx, (index, row_data) in enumerate(df_to_show.iterrows()):
            location_alert_text = str(row_data.get(Columns.LOCATION_ALERT, ''))
            has_location_alert = bool(location_alert_text)
            
            # ✅ --- START: Get New Alert Text ---
            activity_alert_text = str(row_data.get('activity_alert', ''))
            has_activity_alert = bool(activity_alert_text)
            # ✅ --- END: Get New Alert Text ---
            
            original_code = str(row_data.get(Columns.ACTIVITY_CODE, ''))
            correct_aliquot = row_data.get(Columns.CORRECT_RATE, 0.0)
            activity_desc = str(row_data.get(self.activity_desc_col, 'N/A'))

            items = {
                0: str(row_data.get(Columns.INVOICE_NUMBER, '')),
                1: str(row_data.get(Columns.SERVICE_DESCRIPTION, ''))[:200] + "...", # Truncate in table
                2: original_code,
                3: activity_desc,
                4: f"{correct_aliquot:.2f}",
                5: activity_alert_text, # ✅ New item
                6: location_alert_text  # ✅ Index shifted
            }
            
            for col_idx, text in items.items():
                item = QTableWidgetItem(text)
                item.setData(Qt.ItemDataRole.UserRole, index) # Store the ORIGINAL index
                
                if col_idx == 1: # Service Description
                    item.setToolTip(str(row_data.get(Columns.SERVICE_DESCRIPTION, '')))
                else:
                    item.setToolTip(text)
                
                # ✅ Color row if *either* alert is present
                if has_location_alert or has_activity_alert:
                    item.setBackground(alert_color)

                if col_idx == 4: # Aliquota
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(row_idx, col_idx, item)
        
        self.table.setSortingEnabled(True)
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(1, 400) # Description
        self.table.setColumnWidth(3, 300) # Activity Desc
        self.table.setColumnWidth(5, 300) # Activity Alert

    def relabel_selected(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows: return
        
        selected_data = self.activity_combo.currentData()
        if not selected_data:
            return
        
        new_code, selected_description = selected_data
        
        new_aliquot = 0.0
        for description, aliquot, _ in self.activity_data.get(new_code, []):
            if description == selected_description:
                new_aliquot = aliquot
                break

        for row_model_index in selected_rows:
            original_index = self.table.item(row_model_index.row(), 0).data(Qt.ItemDataRole.UserRole)
            self.df.loc[original_index, Columns.ACTIVITY_CODE] = new_code
            self.df.loc[original_index, Columns.CORRECT_RATE] = new_aliquot
            self.df.loc[original_index, self.activity_desc_col] = selected_description
            # ✅ Clear the alert, since it has been manually reviewed
            self.df.loc[original_index, 'activity_alert'] = ""
        
        self.populate_table() # Re-draws the table (which will re-filter)

    def _mark_as_tomador(self):
        """Marks selected invoices as 'Local Tomador' and removes them from the view."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows: return
        
        indices_to_mark = []
        for row_model_index in selected_rows:
            original_index = self.table.item(row_model_index.row(), 0).data(Qt.ItemDataRole.UserRole)
            indices_to_mark.append(original_index)
            
        self.tomador_indices.extend(indices_to_mark)
        self.tomador_indices = list(set(self.tomador_indices)) # Remove duplicates
        
        # Refresh the table, which will now exclude these indices
        self.populate_table() 

    def get_tomador_indices(self):
        """Returns the list of indices marked as 'Local Tomador'."""
        return self.tomador_indices

    def get_updated_dataframe(self):
        """Returns the dataframe with activity/rate changes."""
        return self.df