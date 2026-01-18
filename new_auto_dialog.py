# --- FILE: app/new_auto_dialog.py ---

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QFormLayout, 
                               QLineEdit, QComboBox, QDialogButtonBox, 
                               QMessageBox, QLabel)  # Removed QDoubleValidator from here
from PySide6.QtGui import QDoubleValidator # ✅ Moved to QtGui
from PySide6.QtCore import Qt, QLocale
import pandas as pd
from datetime import datetime
import re

class NewAutoDialog(QDialog):
    def __init__(self, parent=None, auto_id=None, motive=None, correct_aliquota=None, current_year=None):
        super().__init__(parent)
        self.setWindowTitle("Configurar Auto/IDD")
        self.resize(400, 280)
        
        self.data = {}
        
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        
        # 1. Auto ID
        self.auto_id_edit = QLineEdit(auto_id or "")
        form_layout.addRow("Número:", self.auto_id_edit)
        
        # 2. Motive (Standard List)
        self.motive_combo = QComboBox()
        self.motive_combo.setEditable(True)
        standard_motives = [
            "Alíquota Incorreta",
            "Dedução Indevida",
            "IDD (Não Pago)", # This triggers the "IDD" label in the report
            "Natureza da Operação Incompatível",
            "Regime Incorreto",
            "Isenção/Imunidade Indevida",
            "Retenção na Fonte (Verificar)"
        ]
        self.motive_combo.addItems(standard_motives)
        
        # Clean motive text (remove old year suffix if editing)
        clean_motive = motive
        existing_year_in_motive = None
        
        if motive:
            # Check for (YYYY) pattern
            match = re.search(r'(.+?)\s*\((\d{4})\)$', motive)
            if match:
                clean_motive = match.group(1)
                existing_year_in_motive = match.group(2)
            self.motive_combo.setCurrentText(clean_motive)
            
        form_layout.addRow("Motivo:", self.motive_combo)
        
        # 3. Year Selector (NEW)
        self.year_combo = QComboBox()
        
        # Try to get years from the wizard data
        available_years = set()
        if parent and hasattr(parent, 'wizard') and hasattr(parent.wizard, 'all_invoices_df'):
            df = parent.wizard.all_invoices_df
            if not df.empty and 'DATA EMISSÃO' in df.columns:
                try:
                    dates = pd.to_datetime(df['DATA EMISSÃO'], errors='coerce').dropna()
                    if not dates.empty:
                        available_years.update(dates.dt.year.unique())
                except: pass
        
        # Fallback if no data loaded
        if not available_years:
            cy = datetime.now().year
            available_years = {cy, cy-1, cy-2, cy-3}
            
        sorted_years = sorted(list(available_years))
        self.year_combo.addItems([str(y) for y in sorted_years])
        
        # Set Selection Priority:
        # 1. Year passed in arg (current_year)
        # 2. Year extracted from existing motive string
        # 3. Last year in the list
        if current_year:
            self.year_combo.setCurrentText(str(current_year))
        elif existing_year_in_motive:
            self.year_combo.setCurrentText(existing_year_in_motive)
        elif sorted_years:
            self.year_combo.setCurrentIndex(len(sorted_years)-1)
        
        form_layout.addRow("Ano de Competência:", self.year_combo)

        # 4. Aliquot
        self.aliquota_edit = QLineEdit()
        self.aliquota_edit.setPlaceholderText("Ex: 5.00 (Opcional)")
        locale = QLocale(QLocale.Language.Portuguese, QLocale.Country.Brazil)
        validator = QDoubleValidator(0.0, 100.0, 2)
        validator.setLocale(locale)
        self.aliquota_edit.setValidator(validator)
        
        if correct_aliquota is not None:
            self.aliquota_edit.setText(locale.toString(float(correct_aliquota), 'f', 2))
            
        form_layout.addRow("Alíquota Correta (%):", self.aliquota_edit)
        
        layout.addLayout(form_layout)
        
        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def validate_and_accept(self):
        auto_id = self.auto_id_edit.text().strip()
        raw_motive = self.motive_combo.currentText().strip()
        year = self.year_combo.currentText().strip()
        
        if not auto_id:
            QMessageBox.warning(self, "Erro", "ID é obrigatório.")
            return
        if not raw_motive:
            QMessageBox.warning(self, "Erro", "Motivo é obrigatório.")
            return
            
        # Combine Motive + Year to create the filtering key
        final_motive_key = f"{raw_motive} ({year})"
        
        al_val = None
        if self.aliquota_edit.text():
            try:
                locale = QLocale(QLocale.Language.Portuguese, QLocale.Country.Brazil)
                al_val = locale.toDouble(self.aliquota_edit.text())[0]
            except: pass

        self.data = {
            'auto_id': auto_id,
            'motive': final_motive_key, # "IDD (Não Pago) (2022)"
            'correct_aliquota': al_val,
            'year': int(year)
        }
        self.accept()

    def get_data(self):
        return self.data