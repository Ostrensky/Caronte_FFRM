# --- FILE: app/duplicate_review_dialog.py ---
import pandas as pd
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QListWidget, 
                               QTableWidget, QTableWidgetItem, QPushButton, 
                               QLabel, QSplitter, QHeaderView, QMessageBox, 
                               QWidget, QAbstractItemView)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor

class DuplicateReviewDialog(QDialog):
    """
    Dialog to manually review and resolve invoice conflicts.
    """
    def __init__(self, df_with_conflicts, conflict_col='_is_conflict', id_col='NÚMERO', parent=None):
        super().__init__(parent)
        self.setWindowTitle("Revisão de Duplicatas de Notas")
        self.resize(1150, 700) # Increased width for new column
        
        self.df = df_with_conflicts
        
        # 1. Normalize Column Names for the Dialog Logic
        self.col_map = self._map_columns(self.df.columns)
        
        # Use the mapped ID column or fallback
        self.id_col = self.col_map.get('id', id_col) 
        self.conflict_col = conflict_col
        
        # Ensure conflict column exists
        if self.conflict_col not in self.df.columns:
            # Fallback: calculate if missing
            self.df[self.conflict_col] = self.df.duplicated(subset=[self.id_col], keep=False)

        # Extract conflicts
        self.conflict_groups = self.df[self.df[self.conflict_col]].groupby(self.id_col)
        self.conflict_ids = sorted(list(self.conflict_groups.groups.keys()))
        
        # Store user decisions: {invoice_number: selected_index_in_original_df}
        self.decisions = {} 
        
        # Pre-fill decisions with the "first" (Smart Sort winner) as default
        for invoice_num, indices in self.conflict_groups.groups.items():
            self.decisions[invoice_num] = indices[0]

        self._setup_ui()
        self._load_list()

    def _map_columns(self, available_cols):
        """
        Smartly maps normalized concepts to actual DataFrame columns (Case Insensitive).
        """
        mapping = {}
        # Concepts we need
        targets = {
            'id': ['NÚMERO', 'NUMERO', 'NUMBER', 'No', 'Nº', 'NR'],
            'date': ['DATA EMISSÃO', 'DATA', 'DT EMISSAO', 'EMISSAO', 'DATA EMISSAO'],
            'value': ['VALOR', 'VALOR (R$)', 'VALOR TOTAL', 'VLR'],
            'rate': ['ALÍQUOTA', 'ALIQUOTA', 'ALIQ', 'ALIQ.'],
            'regime': ['REGIME DE TRIBUTAÇÃO', 'REGIME', 'REGIME TRIBUTACAO', 'TRIBUTACAO'],
            'rps': ['Nº RPS', 'RPS', 'NO RPS', 'NUMERO RPS'], # ✅ ADDED RPS MAPPING
            'desc': [
                'DISCRIMINAÇÃO DOS SERVIÇOS', 'DISCRIMINACAO DOS SERVICOS', 
                'DISCRIMINAÇÃO', 'DISCRIMINACAO', 
                'DESCRIÇÃO', 'DESCRICAO', 'DESC', 'DESCRICAO DOS SERVICOS',
                'HISTORICO', 'SERV. PRESTADO'
            ],
            'code': ['CÓDIGO DA ATIVIDADE', 'CODIGO DA ATIVIDADE', 'CODIGO', 'ATIVIDADE', 'COD. ATIVIDADE']
        }
        
        # Create uppercase lookup
        upper_cols = {str(c).upper().strip(): c for c in available_cols}
        
        for key, candidates in targets.items():
            for cand in candidates:
                cand_upper = cand.upper()
                if cand_upper in upper_cols:
                    mapping[key] = upper_cols[cand_upper]
                    break
        return mapping

    def _get_val(self, row, key, default=''):
        """Safely retrieves value using the column map."""
        col_name = self.col_map.get(key)
        if col_name and col_name in row:
            val = row[col_name]
            return val if pd.notna(val) else default
        return default

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        
        info_label = QLabel(f"⚠️ Foram encontradas {len(self.conflict_ids)} notas com números duplicados mas conteúdos diferentes.\n"
                            "O sistema pré-selecionou a 'melhor' versão priorizando: Sem RPS=1 > Sem Simples Nacional > Maior Alíquota.")
        info_label.setStyleSheet("color: #4C566A; font-size: 14px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(info_label)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.addWidget(QLabel("Notas com Conflito:"))
        self.list_widget = QListWidget()
        self.list_widget.currentRowChanged.connect(self._on_item_selected)
        left_layout.addWidget(self.list_widget)
        splitter.addWidget(left_widget)
        
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        self.detail_label = QLabel("Detalhes (Selecione a versão correta):")
        right_layout.addWidget(self.detail_label)
        
        self.table = QTableWidget()
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.itemClicked.connect(self._on_row_clicked)
        right_layout.addWidget(self.table)
        splitter.addWidget(right_widget)
        
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        layout.addWidget(splitter)

        btn_layout = QHBoxLayout()
        auto_btn = QPushButton("Aceitar Sugestão Automática (Todas)")
        auto_btn.clicked.connect(self.accept_defaults)
        
        save_btn = QPushButton("Confirmar Minha Seleção")
        save_btn.setStyleSheet("background-color: #5E81AC; color: white; font-weight: bold; padding: 8px;")
        save_btn.clicked.connect(self.accept)
        
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addWidget(auto_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def _load_list(self):
        self.list_widget.clear()
        for nf_id in self.conflict_ids:
            self.list_widget.addItem(f"Nota Nº {nf_id}")
        if self.list_widget.count() > 0:
            self.list_widget.setCurrentRow(0)

    def _on_item_selected(self, row_idx):
        if row_idx < 0: return
        nf_id = self.conflict_ids[row_idx]
        group_df = self.conflict_groups.get_group(nf_id)
        
        self.detail_label.setText(f"Variações da Nota {nf_id}:")
        
        # ✅ Added RPS to columns
        cols = ['Status', 'Nº RPS', 'Data Emissão', 'Valor (R$)', 'Alíquota', 'Regime Trib.', 'Descrição', 'Código Ativ.']
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        self.table.setRowCount(len(group_df))
        
        current_selection = self.decisions.get(nf_id)
        
        for i, (orig_idx, row) in enumerate(group_df.iterrows()):
            # 1. Status
            is_selected = (orig_idx == current_selection)
            status_text = "✅ MANTER" if is_selected else "🗑️ Descartar"
            
            item_status = QTableWidgetItem(status_text)
            item_status.setData(Qt.ItemDataRole.UserRole, orig_idx)
            if is_selected:
                item_status.setBackground(QColor("#A3BE8C"))
                item_status.setForeground(QColor("white"))
            else:
                item_status.setBackground(QColor("#BF616A"))
                item_status.setForeground(QColor("white"))
            self.table.setItem(i, 0, item_status)
            
            # 2. RPS (New)
            rps_val = str(self._get_val(row, 'rps', '-'))
            self.table.setItem(i, 1, QTableWidgetItem(rps_val))

            # 3. Date
            dt = self._get_val(row, 'date')
            try:
                if pd.notna(dt): dt = dt.strftime('%d/%m/%Y')
            except: pass
            self.table.setItem(i, 2, QTableWidgetItem(str(dt)))
            
            # 4. Value
            val = self._get_val(row, 'value', 0.0)
            try:
                val_fmt = f"{float(val):,.2f}"
            except: val_fmt = str(val)
            self.table.setItem(i, 3, QTableWidgetItem(val_fmt))
            
            # 5. Aliquot
            aliq = self._get_val(row, 'rate', 0.0)
            try:
                aliq_fmt = f"{float(aliq):.2f}%"
            except: aliq_fmt = str(aliq)
            self.table.setItem(i, 4, QTableWidgetItem(aliq_fmt))
            
            # 6. Regime
            regime = str(self._get_val(row, 'regime', '-'))
            self.table.setItem(i, 5, QTableWidgetItem(regime))
            
            # 7. Description (Truncated)
            desc_full = str(self._get_val(row, 'desc', ''))
            desc_trunc = desc_full[:70] + ("..." if len(desc_full) > 70 else "")
            desc_item = QTableWidgetItem(desc_trunc)
            desc_item.setToolTip(desc_full) 
            self.table.setItem(i, 6, desc_item)
            
            # 8. Activity Code
            self.table.setItem(i, 7, QTableWidgetItem(str(self._get_val(row, 'code', ''))))

        self.table.resizeColumnsToContents()

    def _on_row_clicked(self, item):
        row = item.row()
        nf_id_idx = self.list_widget.currentRow()
        if nf_id_idx < 0: return
        
        nf_id = self.conflict_ids[nf_id_idx]
        status_item = self.table.item(row, 0)
        original_idx = status_item.data(Qt.ItemDataRole.UserRole)
        
        self.decisions[nf_id] = original_idx
        self._on_item_selected(nf_id_idx)

    def accept_defaults(self):
        super().accept()

    def get_resolved_indices(self):
        non_conflict_indices = self.df[~self.df[self.conflict_col]].index.tolist()
        winner_indices = list(self.decisions.values())
        return sorted(non_conflict_indices + winner_indices)