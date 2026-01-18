# --- FILE: app/detail_viewer_dialog.py ---

import pandas as pd
from datetime import datetime
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                               QTableWidget, QDialogButtonBox, QComboBox, 
                               QLineEdit, QLabel, QWidget, QTableWidgetItem,
                               QApplication, QMessageBox, QMenu, QCompleter, # ✅ Import QMenu, QCompleter
                               QHeaderView, QStyle) # ✅ Import QHeaderView
from PySide6.QtGui import QAction # ✅ Import QAction
from PySide6.QtCore import Qt, QStringListModel # ✅ Import QStringListModel
import numpy as np # ✅ Import NumPy

# ✅ Import widgets/constants for consistent table formatting
from .widgets import NumericTableWidgetItem, DateTableWidgetItem, ColumnSelectionDialog, SORT_ROLE
from .constants import Columns
from PySide6.QtGui import QColor


class InvoiceDetailViewerDialog(QDialog):
    def __init__(self, df, all_columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Visualizador de Detalhes da Nota Fiscal")
        self.setMinimumSize(1100, 650)
        self.setWindowState(Qt.WindowState.WindowMaximized)
        self.df = df
        self.all_columns = all_columns
        # ✅ Inicializa as colunas visíveis com as do wizard, se possível
        if hasattr(parent, 'wizard'):
             self.visible_columns = parent.wizard.visible_columns[:]
        else:
             self.visible_columns = all_columns[:]
             
        self.active_filters = []
        self.current_filtered_df = self.df.copy() 

        main_layout = QVBoxLayout(self)
        
        controls_layout = QHBoxLayout()
        self.add_filter_btn = QPushButton("Adicionar Filtro")
        self.add_filter_btn.clicked.connect(self.add_filter_row)
        controls_layout.addWidget(self.add_filter_btn)
        
        self.select_cols_btn = QPushButton("Selecionar Colunas Visíveis")
        self.select_cols_btn.clicked.connect(self.open_column_selection)
        controls_layout.addWidget(self.select_cols_btn)
        
        self.copy_btn = QPushButton("Copiar para Área de Transferência")
        self.copy_btn.clicked.connect(self.copy_table_to_clipboard)
        controls_layout.addWidget(self.copy_btn)
        
        controls_layout.addStretch(1)
        main_layout.addLayout(controls_layout)
        
        self.filters_widget = QWidget()
        self.filters_layout = QVBoxLayout(self.filters_widget)
        self.filters_layout.setContentsMargins(0, 0, 0, 0) # ✅ Adiciona margem
        main_layout.addWidget(self.filters_widget)
        
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTriggers.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        
        # ✅ --- REQ 1: Adicionar menu de clique-direito no cabeçalho ---
        header = self.table.horizontalHeader()
        header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        header.customContextMenuRequested.connect(self.show_column_context_menu)
        # ✅ --- FIM REQ 1 ---
        
        main_layout.addWidget(self.table)
        
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        main_layout.addWidget(buttons)
        
        self.populate_table_from_df(self.df) # Initial population

    # ✅ --- REQ 2: Funções de filtro com Completer ---
    def add_filter_row(self):
        filter_row_layout = QHBoxLayout()
        column_combo = QComboBox()
        column_combo.addItems(self.visible_columns)
        
        filter_edit = QLineEdit()
        filter_edit.setPlaceholderText("Digite o texto...")
        
        completer = QCompleter()
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        filter_edit.setCompleter(completer)
        
        remove_btn = QPushButton()
        remove_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCancelButton))
        remove_btn.setToolTip("Remover este filtro")
        remove_btn.setFixedWidth(30)
        
        filter_row_layout.addWidget(QLabel("Filtrar Coluna:"))
        filter_row_layout.addWidget(column_combo, 1)
        filter_row_layout.addWidget(QLabel("contendo:"))
        filter_row_layout.addWidget(filter_edit, 2)
        filter_row_layout.addWidget(remove_btn)
        self.filters_layout.addLayout(filter_row_layout)
        
        filter_widgets = { 
            "layout": filter_row_layout, 
            "combo": column_combo, 
            "edit": filter_edit, 
            "button": remove_btn,
            "completer": completer
        }
        self.active_filters.append(filter_widgets)
        
        column_combo.currentIndexChanged.connect(
            lambda: self._update_filter_completer(filter_widgets)
        )
        filter_edit.textChanged.connect(self.apply_all_filters)
        remove_btn.clicked.connect(lambda: self.remove_filter_row(filter_widgets))
        
        self._update_filter_completer(filter_widgets)

    def _update_filter_completer(self, filter_widgets):
        try:
            column_name = filter_widgets["combo"].currentText()
            completer = filter_widgets["completer"]
            edit_widget = filter_widgets["edit"]
            
            if column_name not in self.df.columns:
                completer.setModel(None)
                return

            if pd.api.types.is_numeric_dtype(self.df[column_name].dtype):
                completer.setModel(None) 
                edit_widget.setPlaceholderText("Digite o número...")
            else:
                unique_values = self.df[column_name].astype(str).unique()
                model = QStringListModel(unique_values)
                completer.setModel(model)
                edit_widget.setPlaceholderText("Digite ou selecione da lista...")
                
        except Exception as e:
            print(f"Erro ao atualizar completer: {e}")
            if 'completer' in filter_widgets:
                filter_widgets['completer'].setModel(None)
        
        self.apply_all_filters()

    def remove_filter_row(self, filter_widgets):
        if filter_widgets in self.active_filters:
            self.active_filters.remove(filter_widgets)
        for i in reversed(range(filter_widgets["layout"].count())):    
            widget = filter_widgets["layout"].itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()
        filter_widgets["layout"].deleteLater()
        self.apply_all_filters()
    # ✅ --- FIM REQ 2 ---
    
    def _fmtd(self, val):
        """Helper to format currency in BRL (1.234,56)"""
        # ... (função inalterada) ...
        if not isinstance(val, (int, float)):
            try:
                val = float(val)
            except (ValueError, TypeError):
                return "0,00"
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def populate_table_from_df(self, df_to_show):
        """Populates the table from a given DataFrame."""
        # ... (função inalterada, como na última modificação) ...
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        self.table.setColumnCount(len(self.visible_columns))
        self.table.setHorizontalHeaderLabels(self.visible_columns)

        if df_to_show is None or df_to_show.empty:
            self.table.setSortingEnabled(True)
            return

        self.table.setRowCount(len(df_to_show))
        rule_color = QColor(100, 92, 63)
        expired_color = QColor(80, 80, 80)

        for row_idx, (index, row_data) in enumerate(df_to_show.iterrows()):
            status = row_data.get(Columns.STATUS_LEGAL)
            for col_idx, col_name in enumerate(self.visible_columns):
                raw_value = row_data.get(col_name)
                item = None
                tooltip_text = ""

                if col_name == Columns.BROKEN_RULE_DETAILS:
                    display_text = "; ".join(map(str, raw_value)) if isinstance(raw_value, list) and raw_value else ""
                    tooltip_text = "Infrações:\n- " + "\n- ".join(map(str, raw_value)) if display_text else ""
                    item = QTableWidgetItem(display_text)
                
                elif col_name == Columns.VALUE:
                    display_text = self._fmtd(raw_value)
                    item = NumericTableWidgetItem(display_text)
                    item.setData(SORT_ROLE, float(raw_value) if pd.notna(raw_value) else 0.0)
                
                elif col_name in [Columns.RATE, Columns.CORRECT_RATE, 'ALÍQUOTA']:
                    display_text = f"{raw_value:,.2f}" if pd.notna(raw_value) else "0.00"
                    item = NumericTableWidgetItem(display_text)
                    item.setData(SORT_ROLE, float(raw_value) if pd.notna(raw_value) else 0.0)
                
                elif col_name == Columns.ISSUE_DATE and pd.notna(raw_value):
                    display_text = raw_value.strftime('%d/%m/%Y')
                    item = DateTableWidgetItem(display_text)
                    item.setData(SORT_ROLE, raw_value)
                
                else:
                    display_text = str(raw_value if pd.notna(raw_value) else '')
                    item = QTableWidgetItem(display_text)
                
                if status in ['Decadente', 'Prescrito']:
                    item.setBackground(expired_color)
                    tooltip_text += f"\n(Nota {status.lower()}, desconsiderada para autuação)"
                if col_name == Columns.BROKEN_RULE_DETAILS and display_text:
                    item.setBackground(rule_color)

                item.setToolTip(tooltip_text or display_text)
                self.table.setItem(row_idx, col_idx, item)
        
        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)

    def apply_all_filters(self):
        """IMPROVED: Filters the DataFrame first and stores the result."""
        # ... (função inalterada) ...
        filtered_df = self.df.copy()
        for f in self.active_filters:
            column_name = f["combo"].currentText()
            filter_text = f["edit"].text().strip().lower()
            
            if filter_text and column_name in filtered_df.columns:
                try:
                    mask = filtered_df[column_name].astype(str).str.lower().str.contains(filter_text, na=False)
                    filtered_df = filtered_df[mask]
                except Exception as e:
                    print(f"Error filtering column {column_name}: {e}")
        
        self.current_filtered_df = filtered_df 
        self.populate_table_from_df(self.current_filtered_df) 
    
    def copy_table_to_clipboard(self):
        """Copies the currently visible (filtered) data to the clipboard."""
        # ... (função inalterada) ...
        if self.current_filtered_df is not None and not self.current_filtered_df.empty:
            df_to_copy = self.current_filtered_df[self.visible_columns]
            try:
                clipboard = QApplication.clipboard()
                clipboard.setText(df_to_copy.to_csv(sep='\t', index=False))
                QMessageBox.information(self, "Copiado", f"{len(df_to_copy)} linhas (apenas dados filtrados/visíveis) copiadas para a área de transferência.")
            except Exception as e:
                QMessageBox.warning(self, "Erro ao Copiar", f"Não foi possível copiar: {e}")
        else:
            QMessageBox.information(self, "Nada para Copiar", "A tabela (filtrada) está vazia.")
    
    # ✅ --- START: REQ 1: Novas funções para menu de colunas ---
    def show_column_context_menu(self, pos):
        """Cria e exibe um menu de clique-direito para mostrar/ocultar colunas."""
        header = self.sender()
        if not isinstance(header, QHeaderView):
            return
            
        menu = QMenu(self)
        
        for column_name in self.all_columns: # Usa all_columns aqui
            action = QAction(column_name, self)
            action.setCheckable(True)
            action.setChecked(column_name in self.visible_columns)
            action.toggled.connect(
                lambda checked, name=column_name: self.handle_column_visibility_toggle(checked, name)
            )
            menu.addAction(action)
            
        menu.exec(header.mapToGlobal(pos))

    def handle_column_visibility_toggle(self, is_checked, column_name):
        """Adiciona ou remove uma coluna das colunas visíveis."""
        if is_checked:
            if column_name not in self.visible_columns:
                self.visible_columns.append(column_name)
        else:
            if column_name in self.visible_columns:
                self.visible_columns.remove(column_name)
                
        self.refresh_all_filters_and_table()
    
    def open_column_selection(self):
        """Abre o diálogo de fallback para configurar colunas."""
        dialog = ColumnSelectionDialog(self.all_columns, self.visible_columns, self)
        if dialog.exec():
            self.visible_columns = dialog.get_selected_columns()
            self.refresh_all_filters_and_table()

    def refresh_all_filters_and_table(self):
        """Atualiza os dropdowns de filtro e recarrega a tabela."""
        for f in self.active_filters:
            combo = f["combo"]
            current_selection = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(self.visible_columns)
            if current_selection in self.visible_columns:
                combo.setCurrentText(current_selection)
            combo.blockSignals(False)
            
        self.apply_all_filters() # Recarrega a tabela
    # ✅ --- END: REQ 1 ---