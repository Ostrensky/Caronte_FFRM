# --- FILE: app/settings_dialog.py ---

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QGroupBox, QFormLayout, 
                               QLineEdit, QPushButton, QFileDialog, QDialogButtonBox,
                               QHBoxLayout, QWidget)
from PySide6.QtCore import QSettings

# Import the config getters and setters
from .config import (
    get_aliquotas_path, set_aliquotas_path,
    get_template_inicio_path, set_template_inicio_path,
    get_template_relatorio_path, set_template_relatorio_path,
    # ✅ --- Correct imports are already here ---
    get_template_encerramento_dec_path,
    get_template_encerramento_ar_path,
    set_template_encerramento_dec_path,
    set_template_encerramento_ar_path,
    # ✅ --- End ---
    get_output_dir, set_output_dir
)

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preferências")
        self.setMinimumWidth(700)
        
        main_layout = QVBoxLayout(self)
        
        # --- File Paths Group ---
        paths_group = QGroupBox("Caminhos dos Ficheiros")
        paths_layout = QFormLayout()
        
        # ✅ --- FIX: Call with `self.` and rename method ---
        self.aliquotas_edit = self._create_path_selector(
            self.browse_file, "Ficheiros Excel (*.xlsx *.xls)"
        )
        paths_layout.addRow("Ficheiro de Alíquotas:", self.aliquotas_edit)
        
        self.inicio_edit = self._create_path_selector(
            self.browse_file, "Documentos Word (*.docx)"
        )
        paths_layout.addRow("Template Termo de Início:", self.inicio_edit)
        
        self.relatorio_edit = self._create_path_selector(
            self.browse_file, "Documentos Word (*.docx)"
        )
        paths_layout.addRow("Template Relatório:", self.relatorio_edit)
        
        # ✅ --- START: MODIFIED Encerramento paths ---
        # ❌ REMOVED: self.encerramento_edit
        
        # ✅ ADDED: DEC version
        self.encerramento_dec_edit = self._create_path_selector(
            self.browse_file, "Documentos Word (*.docx)"
        )
        paths_layout.addRow("Template Encerramento (DEC):", self.encerramento_dec_edit)
        
        # ✅ ADDED: AR version
        self.encerramento_ar_edit = self._create_path_selector(
            self.browse_file, "Documentos Word (*.docx)"
        )
        paths_layout.addRow("Template Encerramento (AR):", self.encerramento_ar_edit)
        # ✅ --- END: MODIFIED Encerramento paths ---
        
        self.output_edit = self._create_path_selector(self.browse_dir)
        paths_layout.addRow("Pasta de Saída:", self.output_edit)
        # ✅ --- END OF FIX ---
        
        paths_group.setLayout(paths_layout)
        main_layout.addWidget(paths_group)
        
        # --- Dialog Buttons ---
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel |
            QDialogButtonBox.StandardButton.Apply
        )
        buttons.accepted.connect(self.save_and_accept)
        buttons.rejected.connect(self.reject)
        buttons.button(QDialogButtonBox.StandardButton.Apply).clicked.connect(self.save_settings)
        main_layout.addWidget(buttons)
        
        self.load_settings()

    # ✅ --- FIX: Renamed method to be more conventional ---
    def _create_path_selector(self, browse_slot, file_filter=None):
        """Helper to create a QLineEdit + QPushButton combo."""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        
        line_edit = QLineEdit()
        browse_btn = QPushButton("Procurar...")
        
        if file_filter:
            browse_btn.clicked.connect(lambda: browse_slot(line_edit, file_filter))
        else:
            browse_btn.clicked.connect(lambda: browse_slot(line_edit))
            
        layout.addWidget(line_edit)
        layout.addWidget(browse_btn)
        
        # Store the line_edit as a dynamic property of the container widget
        widget.line_edit = line_edit
        return widget

    def browse_file(self, line_edit, file_filter):
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar Ficheiro", line_edit.text(), file_filter)
        if path:
            line_edit.setText(path)
            
    def browse_dir(self, line_edit):
        path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta", line_edit.text())
        if path:
            line_edit.setText(path)

    def load_settings(self):
        """Load settings from config into the UI."""
        self.aliquotas_edit.line_edit.setText(get_aliquotas_path())
        self.inicio_edit.line_edit.setText(get_template_inicio_path())
        self.relatorio_edit.line_edit.setText(get_template_relatorio_path())
        
        # ✅ --- START: MODIFIED to load both paths ---
        self.encerramento_dec_edit.line_edit.setText(get_template_encerramento_dec_path())
        self.encerramento_ar_edit.line_edit.setText(get_template_encerramento_ar_path())
        # ✅ --- END: MODIFIED ---
        
        self.output_edit.line_edit.setText(get_output_dir())

    def save_settings(self):
        """Save settings from UI back to config."""
        set_aliquotas_path(self.aliquotas_edit.line_edit.text())
        set_template_inicio_path(self.inicio_edit.line_edit.text())
        set_template_relatorio_path(self.relatorio_edit.line_edit.text())
        
        # ✅ --- START: MODIFIED to save both paths ---
        set_template_encerramento_dec_path(self.encerramento_dec_edit.line_edit.text())
        set_template_encerramento_ar_path(self.encerramento_ar_edit.line_edit.text())
        # ✅ --- END: MODIFIED ---
        
        set_output_dir(self.output_edit.line_edit.text())
        self.statusBar().showMessage("Preferências salvas.", 3000) # Assuming parent has statusBar

    def save_and_accept(self):
        self.save_settings()
        self.accept()
        
    def statusBar(self):
        # Dummy statusBar to prevent crash if parent has none
        class DummyStatusBar:
            def showMessage(self, *args):
                pass
        return getattr(self.parent(), 'statusBar', DummyStatusBar)()