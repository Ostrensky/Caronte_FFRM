# --- FILE: app/generation_summary_dialog.py ---
# --- [ARQUIVO NOVO] ---

import os
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QListWidget, QListWidgetItem,
    QDialogButtonBox, QLabel, QStyle # --- [NOVA IMPORTAÇÃO] ---
)
from PySide6.QtGui import QDesktopServices, QIcon
from PySide6.QtCore import QUrl, Qt, QSize

class GenerationSummaryDialog(QDialog):
    """
    Um diálogo que aparece após a geração, listando os arquivos criados
    e permitindo que o usuário os abra.
    """
    def __init__(self, output_dir, file_list, parent=None):
        super().__init__(parent)
        self.output_dir = output_dir
        self.file_list = file_list
        
        self.setWindowTitle("Geração Concluída")
        self.setMinimumSize(600, 400)

        # --- Ícone ---
        # --- [LINHA MODIFICADA] ---
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))

        # --- Layout ---
        layout = QVBoxLayout(self)
        
        label = QLabel("Os seguintes documentos foram gerados com sucesso:")
        label.setStyleSheet("font-weight: bold;")
        layout.addWidget(label)
        
        dir_label = QLabel(f"Localização: {self.output_dir}")
        dir_label.setTextInteractionFlags(Qt.TextSelectableByMouse) # Permite copiar o caminho
        dir_label.setStyleSheet("font-style: italic; color: #aaa;")
        layout.addWidget(dir_label)

        # --- Lista de Arquivos ---
        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self.open_file)
        
        # --- [BLOCO MODIFICADO] ---
        # Carrega os ícones corretos usando QStyle
        file_icon = self.style().standardIcon(QStyle.SP_FileIcon)
        pdf_icon = self.style().standardIcon(QStyle.SP_MessageBoxInformation) # Usando um ícone diferente para PDF
        # --- [FIM DO BLOCO MODIFICADO] ---
        
        for filename in sorted(self.file_list):
            item = QListWidgetItem(filename)
            item.setData(Qt.ItemDataRole.UserRole, os.path.join(self.output_dir, filename)) # Armazena o caminho completo
            
            # Define o ícone com base na extensão
            if filename.endswith(".pdf"):
                item.setIcon(pdf_icon)
            else:
                item.setIcon(file_icon)
                
            self.list_widget.addItem(item)
            
        layout.addWidget(self.list_widget)

        # --- Botões ---
        info_label = QLabel("Dê um clique duplo em um arquivo para abri-lo.")
        layout.addWidget(info_label)
        
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

    def open_file(self, item):
# ... (o resto do arquivo está correto e permanece o mesmo) ...
        """
        Abre o arquivo selecionado usando o aplicativo padrão do sistema.
        """
        file_path = item.data(Qt.ItemDataRole.UserRole)
        if file_path:
            try:
                QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))
            except Exception as e:
                # Se houver um erro, exibe-o no diálogo (raro)
                error_label = QLabel(f"Erro ao abrir {os.path.basename(file_path)}: {e}")
                error_label.setStyleSheet("color: #BF616A;") # Cor de erro
                self.layout().addWidget(error_label)