# --- FILE: app/infraction_correction_dialog.py ---

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QListWidget, QListWidgetItem,
                               QDialogButtonBox, QLabel)
from PySide6.QtCore import Qt

class InfractionCorrectionDialog(QDialog):
    def __init__(self, infractions, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Corrigir Infrações da Nota")
        self.setMinimumWidth(500)
        
        layout = QVBoxLayout(self)
        
        label = QLabel("Desmarque as infrações que deseja remover das notas selecionadas:")
        layout.addWidget(label)
        
        self.infraction_list = QListWidget()
        for infraction in sorted(infractions):
            item = QListWidgetItem(infraction)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            self.infraction_list.addItem(item)
            
        layout.addWidget(self.infraction_list)
        
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_infractions_to_keep(self):
        """Returns a list of the infractions that remained checked."""
        to_keep = []
        for i in range(self.infraction_list.count()):
            item = self.infraction_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                to_keep.append(item.text())
        return to_keep