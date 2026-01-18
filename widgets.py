# --- FILE: app/widgets.py ---

from PySide6.QtWidgets import (QTableWidgetItem, QDialog, QListWidget,
                               QListWidgetItem, QDialogButtonBox, QVBoxLayout,
                               QGroupBox, QWidget, QToolButton, QSizePolicy) # Added QGroupBox, QWidget, QToolButton, QSizePolicy
from PySide6.QtCore import (Qt, QPropertyAnimation, QEasingCurve, # Added animation imports
                            QParallelAnimationGroup)

# Custom data role for sorting to avoid conflicts with display text.
SORT_ROLE = Qt.ItemDataRole.UserRole + 1

class NumericTableWidgetItem(QTableWidgetItem):
    """Custom table item for correct numeric sorting."""
    def __lt__(self, other):
        try:
            # Use data(SORT_ROLE) for sorting
            return float(self.data(SORT_ROLE)) < float(other.data(SORT_ROLE))
        except (ValueError, TypeError):
            # Fallback to default comparison if conversion fails
            return super().__lt__(other)
#

class DateTableWidgetItem(QTableWidgetItem):
    """Custom table item for correct date sorting."""
    def __lt__(self, other):
        try:
            # Use data(SORT_ROLE) which should store datetime objects
            return self.data(SORT_ROLE) < other.data(SORT_ROLE)
        except (AttributeError, TypeError):
            # Fallback to default comparison if data is not comparable
            return super().__lt__(other)
#

class ColumnSelectionDialog(QDialog):
    """A dialog to let the user select which table columns are visible."""
    def __init__(self, all_columns, selected_columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Selecionar Colunas Visíveis")
        self.setMinimumWidth(400)

        layout = QVBoxLayout(self)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection) # Allow multi-select if needed, though checkboxes handle it
        for col in all_columns:
            item = QListWidgetItem(col, self.list_widget)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable) # Make item checkable
            # Set initial check state based on currently selected columns
            item.setCheckState(Qt.CheckState.Checked if col in selected_columns else Qt.CheckState.Unchecked)
        layout.addWidget(self.list_widget)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
    #

    def get_selected_columns(self):
        """Returns a list of column names that are checked."""
        selected = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected.append(item.text())
        return selected
    #

# --- Collapsible GroupBox Widget ---
class CollapsibleGroupBox(QGroupBox):
    """A QGroupBox that can be collapsed/expanded with an animation."""
    def __init__(self, title="", parent=None):
        super().__init__(title, parent)

        # Button to toggle collapse/expand, looks like a header
        self.toggle_button = QToolButton(text=title, checkable=True, checked=False)
        self.toggle_button.setStyleSheet("QToolButton { border: none; font-weight: bold; }") # Style as needed
        self.toggle_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toggle_button.setArrowType(Qt.ArrowType.RightArrow) # Start collapsed
        self.toggle_button.pressed.connect(self.on_pressed)

        # Widget to hold the actual content (e.g., the table)
        self.content_area = QWidget()
        self.content_area.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        # Start collapsed
        self.content_area.setMaximumHeight(0)
        self.content_area.setMinimumHeight(0)

        # Layout for the content area itself
        self.content_layout = QVBoxLayout(self.content_area) # Use this to add widgets
        self.content_layout.setContentsMargins(5, 5, 5, 5) # Add some margin inside

        # Main layout for the whole CollapsibleGroupBox
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(self.toggle_button) # Header button
        main_layout.addWidget(self.content_area) # Collapsible content area

        # --- Animation Setup ---
        self.toggle_animation = QParallelAnimationGroup(self)
        # Animation for the content area's height
        self.content_animation = QPropertyAnimation(self.content_area, b"maximumHeight")

        self.toggle_animation.addAnimation(self.content_animation)

    def on_pressed(self):
        """Handles button click to animate expand/collapse."""
        checked = self.toggle_button.isChecked() # State *before* the click is fully processed
        is_expanding = not checked # If it wasn't checked, it's about to be (expanding)

        self.toggle_button.setArrowType(Qt.ArrowType.DownArrow if is_expanding else Qt.ArrowType.RightArrow)

        collapsed_height = 0
        
        # Calculate the height needed by the content layout
        content_height = self.content_layout.sizeHint().height()

        target_height = content_height if is_expanding else collapsed_height
        # A very large number effectively means "expand as much as possible"
        expand_target_height = 16777215 # Essentially QWIDGETSIZE_MAX
        collapse_target_height = 0

        self.content_animation.setDuration(300) # Animation duration in ms
        self.content_animation.setStartValue(self.content_area.maximumHeight()) # Start from current height
        # Set start and end values based on whether we are expanding or collapsing
        self.content_animation.setEndValue(expand_target_height if is_expanding else collapse_target_height)
        self.content_animation.setEasingCurve(QEasingCurve.Type.InOutQuad)

        self.toggle_animation.start() # Start the animation

    def setContentLayout(self, layout):
        """Replaces the existing content layout."""
        # Clean up the old layout and its widgets properly
        old_layout = self.content_area.layout()
        if old_layout:
             while old_layout.count():
                 item = old_layout.takeAt(0)
                 widget = item.widget()
                 if widget:
                     widget.setParent(None) # Remove widget ownership from layout
                     widget.deleteLater() # Schedule widget for deletion
             old_layout.deleteLater() # Schedule layout for deletion

        self.content_area.setLayout(layout)
        self.content_layout = layout # Update internal reference

        # Ensure it starts collapsed after setting new content
        if not self.toggle_button.isChecked():
            self.content_area.setMaximumHeight(0)
            self.content_area.setMinimumHeight(0)

    def addContentWidget(self, widget):
        """Helper to add a widget to the collapsible content area's layout."""
        self.content_layout.addWidget(widget)

    def expand(self):
        """Programmatically expands the group box if collapsed."""
        if not self.toggle_button.isChecked():
            self.toggle_button.setChecked(True) # Set checked state
            self.on_pressed() # Trigger the animation logic

    def collapse(self):
        """Programmatically collapses the group box if expanded."""
        if self.toggle_button.isChecked():
            self.toggle_button.setChecked(False) # Set unchecked state
            self.on_pressed() # Trigger the animation logic