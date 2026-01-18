# --- FILE: app/text_editor_dialog.py ---


from app.config import (get_custom_general_texts, set_custom_general_texts,
                        get_custom_auto_texts, set_custom_auto_texts,
                        DEFAULT_GENERAL_TEXTS, DEFAULT_AUTO_TEXTS)
import copy # To deep copy dictionaries
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QDialogButtonBox,
                               QPushButton, QWidget, QLabel, QTextEdit, QScrollArea,
                               QTabWidget, QComboBox, QMessageBox, QSplitter, QLineEdit)
from PySide6.QtCore import Qt, Signal

class TextEditorDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Editar Textos Padrão do Relatório")
        self.setMinimumSize(900, 600)

        # Load initial data (make deep copies to avoid modifying defaults)
        self.current_general_texts = copy.deepcopy(get_custom_general_texts())
        self.current_auto_texts = copy.deepcopy(get_custom_auto_texts())

        # Main Layout
        main_layout = QVBoxLayout(self)

        # Tabs
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # -- General Texts Tab --
        general_tab = QWidget()
        general_layout = QVBoxLayout(general_tab)
        general_scroll = QScrollArea()
        general_scroll.setWidgetResizable(True)
        general_scroll_widget = QWidget()
        self.general_edits_layout = QVBoxLayout(general_scroll_widget)
        general_scroll.setWidget(general_scroll_widget)
        general_layout.addWidget(general_scroll)
        self.tab_widget.addTab(general_tab, "Textos Gerais do Relatório")

        # -- Auto Texts Tab --
        auto_tab = QWidget()
        auto_layout = QVBoxLayout(auto_tab)
        # Splitter for better layout
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left side: ComboBox for auto type
        auto_select_widget = QWidget()
        auto_select_layout = QVBoxLayout(auto_select_widget)
        auto_select_layout.addWidget(QLabel("Selecionar Tipo de Infração:"))
        self.auto_type_combo = QComboBox()
        # Add sorted keys + fallback
        auto_keys = sorted(list(DEFAULT_AUTO_TEXTS.keys()))
        if "DEFAULT_AUTO_FALLBACK" in auto_keys: # Move fallback to end
            auto_keys.remove("DEFAULT_AUTO_FALLBACK")
            auto_keys.append("DEFAULT_AUTO_FALLBACK")
        self.auto_type_combo.addItems(auto_keys)
        self.auto_type_combo.currentTextChanged.connect(self.load_auto_text)
        auto_select_layout.addWidget(self.auto_type_combo)
        auto_select_layout.addStretch()
        splitter.addWidget(auto_select_widget)

        # Right side: TextEdit for the selected auto type
        self.auto_text_edit = QTextEdit()
        splitter.addWidget(self.auto_text_edit)
        splitter.setSizes([300, 700]) # Adjust initial sizes

        auto_layout.addWidget(splitter)
        self.tab_widget.addTab(auto_tab, "Textos por Tipo de Infração")

        # Buttons
        button_layout = QHBoxLayout()
        import_button = QPushButton("Importar...")
        import_button.clicked.connect(self.handle_import) # Chama função interna
        export_button = QPushButton("Exportar...")
        export_button.clicked.connect(self.handle_export) # Chama função interna
        reset_button = QPushButton("Restaurar Padrões")
        reset_button.clicked.connect(self.reset_to_defaults)
        button_layout.addWidget(reset_button)
        button_layout.addStretch()
        dialog_buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        dialog_buttons.accepted.connect(self.save_and_accept)
        dialog_buttons.rejected.connect(self.reject)
        button_layout.addWidget(dialog_buttons)
        main_layout.addLayout(button_layout)

        # Populate initial fields
        self.general_text_edits = {} # Store QTextEdit widgets for general texts
        self.populate_general_texts()
        self.load_auto_text(self.auto_type_combo.currentText()) # Load initial auto text

    def handle_import(self):
        # A janela principal tem as funções de import/export
        if hasattr(self.parent(), 'import_custom_texts'):
            # Chama a função de importação da janela pai
            self.parent().import_custom_texts()
            # Recarrega os textos no dialog APÓS a importação ter sido feita
            self.reload_texts_from_config()

    def handle_export(self):
        if hasattr(self.parent(), 'export_custom_texts'):
            self.parent().export_custom_texts() # Chama a função da janela pai

    # ✅ Nova função para recarregar textos
    def reload_texts_from_config(self):
        self.current_general_texts = copy.deepcopy(get_custom_general_texts())
        self.current_auto_texts = copy.deepcopy(get_custom_auto_texts())
        # Repopulate UI
        self.populate_general_texts()
        self.load_auto_text(self.auto_type_combo.currentText())
        self.texts_imported.emit() # Emite sinal (opcional)

    def populate_general_texts(self):
        # Clear existing widgets first
        for widget in self.general_text_edits.values():
            widget.deleteLater()
        self.general_text_edits = {}
        # Clear layout (important if repopulating)
        while self.general_edits_layout.count():
            item = self.general_edits_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        # Add editors for each general text
        for key, text in self.current_general_texts.items():
            label = QLabel(f"<b>{key}:</b>")
            label.setWordWrap(True)
            text_edit = QTextEdit(text)
            text_edit.setFixedHeight(80) # Adjust height as needed
            self.general_edits_layout.addWidget(label)
            self.general_edits_layout.addWidget(text_edit)
            self.general_text_edits[key] = text_edit
        self.general_edits_layout.addStretch()

    def load_auto_text(self, auto_key):
        if auto_key in self.current_auto_texts:
            self.auto_text_edit.setPlainText(self.current_auto_texts[auto_key])

    def save_current_auto_text(self):
        """Saves the text from the editor back to the internal dict."""
        current_key = self.auto_type_combo.currentText()
        if current_key:
            self.current_auto_texts[current_key] = self.auto_text_edit.toPlainText()

    def save_and_accept(self):
        # Salva textos gerais (incluindo auditor)
        temp_general = {}
        for key, editor_widget in self.general_text_edits.items():
            if isinstance(editor_widget, QLineEdit):
                temp_general[key] = editor_widget.text()
            elif isinstance(editor_widget, QTextEdit):
                temp_general[key] = editor_widget.toPlainText()
        set_custom_general_texts(temp_general) # Salva o que está na UI

        # Salva o texto do auto atual antes de salvar tudo
        self.save_current_auto_text() # Atualiza self.current_auto_texts
        set_custom_auto_texts(self.current_auto_texts) # Salva o dict atualizado

        QMessageBox.information(self, "Sucesso", "Textos personalizados salvos.")
        self.accept()

    def reset_to_defaults(self):
        reply = QMessageBox.question(self, "Restaurar Padrões",
                                     "Tem certeza que deseja restaurar todos os textos para os padrões originais?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            # Reverte para os defaults no config E atualiza a UI
            set_custom_general_texts(DEFAULT_GENERAL_TEXTS)
            set_custom_auto_texts(DEFAULT_AUTO_TEXTS)
            self.reload_texts_from_config() # Recarrega e atualiza UI
            QMessageBox.information(self, "Restaurado", "Textos restaurados para o padrão.")

    # Override currentTextChanged handler to save previous text first
    def eventFilter(self, obj, event):
         # Standard event processing
         return super().eventFilter(obj, event)

    def connect_signals(self):
         # Save text when switching auto types
         self.auto_type_combo.activated.connect(self.save_current_auto_text_before_switch)

    def save_current_auto_text_before_switch(self, index):
        # This approach is simpler: save before loading the new one
        self.save_current_auto_text()
        # The load_auto_text is already connected to currentTextChanged