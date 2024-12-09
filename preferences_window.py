from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QColorDialog


class PreferencesWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preferences")
        self.layout = QVBoxLayout()

        self.errorColors = {}

        for error_type, color in error_type_color_map.items():
            label = QLabel(f"{error_type}:")
            self.layout.addWidget(label)

            colorButton = QPushButton()
            colorButton.setStyleSheet(f"background-color: {color.name}")
            colorButton.clicked.connect(
                lambda _, error_type=error_type: self.setColor(error_type)
            )
            self.layout.addWidget(colorButton)

            self.errorColors[error_type] = colorButton

        self.setLayout(self.layout)
        self.setMinimumWidth(300)  # Set minimum width for the dialog

    def setColor(self, error_type):
        color = QColorDialog.getColor()
        if color.isValid():
            self.errorColors[error_type].setStyleSheet(
                f"background-color: {color.name}"
            )
            error_type_color_map[error_type] = color
