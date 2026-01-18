# --- FILE: app/widgets/excel_filter.py ---

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QLineEdit, QListWidget, 
                               QListWidgetItem, QDialogButtonBox, QCheckBox, 
                               QHBoxLayout, QLabel, QPushButton)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QIcon, QAction
import pandas as pd

class ExcelFilterDialog(QDialog):
    def __init__(self, unique_values, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Filtrar")
        self.setWindowFlags(Qt.Popup) # Popup style (closes if clicked outside)
        self.setLayout(QVBoxLayout())
        self.layout().setContentsMargins(5, 5, 5, 5)

        # 1. Search Bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Pesquisar...")
        self.search_input.textChanged.connect(self.filter_list)
        self.layout().addWidget(self.search_input)

        # 2. Controls Layout (Select All + Invert)
        controls_layout = QHBoxLayout()
        
        self.check_all = QCheckBox("(Selecionar Todos os Resultados)")
        self.check_all.setChecked(True)
        # We use 'clicked' instead of stateChanged to handle logic manually
        self.check_all.clicked.connect(self.toggle_visible_items)
        
        controls_layout.addWidget(self.check_all)
        controls_layout.addStretch()
        
        # Helper: Invert Selection (Useful for "Not That" logic)
        self.invert_btn = QPushButton("Inverter")
        self.invert_btn.setToolTip("Inverte a seleção dos itens visíveis")
        self.invert_btn.setFixedWidth(60)
        self.invert_btn.clicked.connect(self.invert_visible_selection)
        controls_layout.addWidget(self.invert_btn)
        
        self.layout().addLayout(controls_layout)

        # 3. List of Values
        self.list_widget = QListWidget()
        self.items = []
        
        # Sort values, putting blanks at the top or bottom
        sorted_values = sorted(map(str, unique_values), key=lambda x: (x is None, x))
        
        for val in sorted_values:
            text_label = val if val else "(Vazio)"
            item = QListWidgetItem(text_label)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            
            # Store real value in UserRole (in case text is modified for display)
            item.setData(Qt.UserRole, val)
            
            self.list_widget.addItem(item)
            self.items.append(item)
            
        self.layout().addWidget(self.list_widget)

        # 4. Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.layout().addWidget(buttons)

    def filter_list(self, text):
        """Hides items that don't match the search text."""
        text = text.lower()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item_text = item.text().lower()
            
            is_match = text in item_text
            item.setHidden(not is_match)
        
        # Update "Check All" state based on visible items
        self.update_check_all_state()

    def toggle_visible_items(self):
        """
        Triggered when user clicks 'Select All'. 
        Only checks/unchecks items that are currently VISIBLE (searched).
        """
        state = self.check_all.checkState()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                item.setCheckState(state)

    def invert_visible_selection(self):
        """
        Inverts the check state of currently visible items.
        Great for: Search 'Curitiba' -> Invert (unchecks them) -> OK.
        """
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                current = item.checkState()
                new_state = Qt.Unchecked if current == Qt.Checked else Qt.Checked
                item.setCheckState(new_state)
        self.update_check_all_state()

    def update_check_all_state(self):
        """
        Updates the master checkbox based on the state of visible items.
        """
        all_checked = True
        any_checked = False
        visible_count = 0
        
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                visible_count += 1
                if item.checkState() == Qt.Checked:
                    any_checked = True
                else:
                    all_checked = False
        
        if visible_count == 0:
            self.check_all.setCheckState(Qt.Unchecked)
        elif all_checked:
            self.check_all.setCheckState(Qt.Checked)
        elif any_checked:
            self.check_all.setCheckState(Qt.PartiallyChecked)
        else:
            self.check_all.setCheckState(Qt.Unchecked)

    def get_selected_values(self):
        """Returns a set of strings that are checked."""
        selected = set()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                # Retrieve original value stored in UserRole
                val = item.data(Qt.UserRole)
                selected.add(val)
        return selected

# --- FilterableHeaderView remains mostly the same, ensuring it calls this Dialog ---
from PySide6.QtWidgets import QHeaderView

class FilterableHeaderView(QHeaderView):
    filter_changed = Signal(int, set) 

    def __init__(self, parent=None):
        super().__init__(Qt.Horizontal, parent)
        self.setSectionsClickable(True)
        self.setHighlightSections(True)
        self.filters = {} 
        self.df_source = None 

    def set_dataframe(self, df):
        self.df_source = df

    def paintSection(self, painter, rect, logicalIndex):
        painter.save()
        super().paintSection(painter, rect, logicalIndex)
        painter.restore()

        if logicalIndex in self.filters:
            icon_size = 14
            # Draw a funnel icon or standard 'list view' icon to indicate filter
            icon = self.style().standardIcon(self.style().StandardPixmap.SP_FileDialogListView) 
            
            # Position icon on the right
            icon_rect = rect.adjusted(rect.width() - icon_size - 6, 6, -6, -6)
            icon_rect.setWidth(icon_size)
            icon_rect.setHeight(icon_size)
            icon.paint(painter, icon_rect)

    def mousePressEvent(self, event):
        logicalIndex = self.logicalIndexAt(event.position().toPoint())
        
        # Right click OR Left click on the far right edge triggers filter
        # Let's make Right Click the easiest way
        if event.button() == Qt.RightButton and logicalIndex >= 0:
            self.show_filter_dialog(logicalIndex)
        else:
            super().mousePressEvent(event)

    def show_filter_dialog(self, col_index):
        if self.df_source is None or self.df_source.empty:
            return

        # Attempt to get column name
        try:
            # Assumes the parent table has a 'column_mapping' list
            raw_col_name = self.parent().column_mapping[col_index]
        except:
            return

        unique_vals = self.df_source[raw_col_name].fillna("").astype(str).unique()

        dialog = ExcelFilterDialog(unique_vals, self)
        
        # Position dialog right under the header section
        header_pos = self.mapToGlobal(self.sectionPosition(col_index))
        dialog.move(header_pos.x(), header_pos.y() + self.height())
        
        # Pre-select based on existing filters
        if col_index in self.filters:
            current_allowed = self.filters[col_index]
            for i in range(dialog.list_widget.count()):
                item = dialog.list_widget.item(i)
                val = item.data(Qt.UserRole)
                if val not in current_allowed:
                    item.setCheckState(Qt.Unchecked)
            dialog.update_check_all_state()

        if dialog.exec():
            selected = dialog.get_selected_values()
            
            # If everything is selected, clear the filter (optimization)
            if len(selected) == len(unique_vals):
                if col_index in self.filters:
                    del self.filters[col_index]
            else:
                self.filters[col_index] = selected
            
            self.filter_changed.emit(col_index, selected)
            self.viewport().update()