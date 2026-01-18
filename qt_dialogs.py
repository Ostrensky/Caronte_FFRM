# --- app/ferramentas/qt_dialogs.py ---

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLineEdit,
                               QPushButton, QFileDialog, QLabel, QFormLayout,
                               QComboBox, QDialogButtonBox, QListWidget, 
                               QListWidgetItem, QCheckBox, QWidget, QGroupBox,
                               QAbstractItemView)
from PySide6.QtCore import QDate, Qt

class GetFolderAndYearsDialog(QDialog):
    """
    Dialogo para selecionar uma pasta raiz e MÚLTIPLOS ANOS.
    """
    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(450, 300)
        
        self.folder_path = ""
        self.selected_years = []

        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        # 1. Seletor de Pasta
        folder_layout = QHBoxLayout()
        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("Selecione a pasta raiz...")
        self.folder_edit.setReadOnly(True)
        folder_button = QPushButton("Procurar...")
        folder_button.clicked.connect(self.browse_folder)
        folder_layout.addWidget(self.folder_edit)
        folder_layout.addWidget(folder_button)
        form_layout.addRow("Pasta Raiz:", folder_layout)

        # 2. Seletor de Anos (Lista de Seleção Múltipla)
        lbl_years = QLabel("Selecione os Anos (Segure Ctrl ou Shift para múltiplos):")
        layout.addWidget(lbl_years)
        
        self.years_list = QListWidget()
        self.years_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        
        current_year = QDate.currentDate().year()
        # Listar ultimos 10 anos + proximo ano
        for y in range(current_year + 1, current_year - 15, -1):
            item = QListWidgetItem(str(y))
            self.years_list.addItem(item)
            # Pré-selecionar o ano passado como padrão
            if y == current_year - 1:
                item.setSelected(True)
        
        self.years_list.setFixedHeight(150)
        layout.addWidget(self.years_list)

        # Layout Final
        layout.addLayout(form_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecionar Pasta Raiz", "", QFileDialog.Option.ShowDirsOnly)
        if folder:
            self.folder_path = folder
            self.folder_edit.setText(folder)

    def accept(self):
        self.folder_path = self.folder_edit.text()
        
        # Coletar anos selecionados
        self.selected_years = []
        for item in self.years_list.selectedItems():
            try:
                self.selected_years.append(int(item.text()))
            except ValueError:
                pass
        
        # Sort desc
        self.selected_years.sort(reverse=True)

        if not self.folder_path:
            return
        if not self.selected_years:
            return # Deve selecionar pelo menos um
            
        super().accept()

    def get_values(self):
        """Retorna (folder_path, list_of_years)"""
        return self.folder_path, self.selected_years

# --- Mantendo as outras classes (CNPJSelectionDialog, DbSelectionDialog) inalteradas ---

class CNPJSelectionDialog(QDialog):
    def __init__(self, company_list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Selecione as Empresas")
        self.resize(600, 500)
        
        self.selected_cnpjs = []
        
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Marque os itens que você deseja processar:"))
        
        self.list_widget = QListWidget()
        for item in company_list:
            cnpj_or_imu = item.get('cnpj', '') 
            name = item.get('name', 'Desconhecido')
            is_selected = item.get('is_selected', True) 
            
            display_text = f"{name} ({cnpj_or_imu})"
            
            list_item = QListWidgetItem(display_text)
            list_item.setData(Qt.UserRole, cnpj_or_imu)
            list_item.setData(Qt.UserRole + 1, item.get('dir_path', '')) 
            
            list_item.setCheckState(Qt.Checked if is_selected else Qt.Unchecked) 
            self.list_widget.addItem(list_item)
            
        layout.addWidget(self.list_widget)
        
        btn_layout = QHBoxLayout()
        select_all_btn = QPushButton("Selecionar Todos")
        select_all_btn.clicked.connect(lambda: self.set_all_checked(True))
        select_none_btn = QPushButton("Desmarcar Todos")
        select_none_btn.clicked.connect(lambda: self.set_all_checked(False))
        btn_layout.addWidget(select_all_btn)
        btn_layout.addWidget(select_none_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def set_all_checked(self, checked):
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(state)

    def accept(self):
        self.selected_cnpjs = [] 
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                self.selected_cnpjs.append({
                    'cnpj': item.data(Qt.UserRole),
                    'dir_path': item.data(Qt.UserRole + 1)
                })
        super().accept()

    def get_selected_cnpjs(self):
        return self.selected_cnpjs
    
class DbSelectionDialog(QDialog):
    def __init__(self, company_list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configuração de Extração (BD)")
        self.resize(600, 600)
        
        self.selected_items = []
        self.run_dams = True
        self.run_nfse = True
        
        layout = QVBoxLayout(self)
        
        layout.addWidget(QLabel("1. Selecione as Empresas:"))
        
        self.list_widget = QListWidget()
        for item in company_list:
            display_text = f"{item.get('name')} (IMU: {item.get('imu')})"
            list_item = QListWidgetItem(display_text)
            list_item.setData(Qt.UserRole, item) 
            list_item.setCheckState(Qt.Checked) 
            self.list_widget.addItem(list_item)
            
        layout.addWidget(self.list_widget)
        
        list_btn_layout = QHBoxLayout()
        btn_all = QPushButton("Todas")
        btn_all.clicked.connect(lambda: self.set_all_list(True))
        btn_none = QPushButton("Nenhuma")
        btn_none.clicked.connect(lambda: self.set_all_list(False))
        list_btn_layout.addWidget(btn_all)
        list_btn_layout.addWidget(btn_none)
        list_btn_layout.addStretch()
        layout.addLayout(list_btn_layout)

        opt_group = QGroupBox("2. Tipo de Relatório")
        opt_layout = QHBoxLayout()
        
        self.check_dams = QCheckBox("DAMS (Pagamentos)")
        self.check_dams.setChecked(True)
        
        self.check_nfse = QCheckBox("NFSE (Notas Fiscais)")
        self.check_nfse.setChecked(True)
        
        opt_layout.addWidget(self.check_dams)
        opt_layout.addWidget(self.check_nfse)
        opt_group.setLayout(opt_layout)
        
        layout.addWidget(opt_group)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def set_all_list(self, checked):
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(state)

    def accept(self):
        self.selected_items = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                self.selected_items.append(item.data(Qt.UserRole))
        
        self.run_dams = self.check_dams.isChecked()
        self.run_nfse = self.check_nfse.isChecked()
        
        super().accept()

    def get_data(self):
        return self.selected_items, self.run_dams, self.run_nfse