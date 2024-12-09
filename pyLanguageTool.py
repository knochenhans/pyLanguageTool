import os
import sys

from PySide6.QtCore import QSettings, Qt

from PySide6.QtWidgets import (
    QApplication,
)

templates = [
    {
        "name": "Smartcat",
        "simple": False,
        "row": 0,
        "source_col_index": 1,
        "target_col_index": 2,
    },
    {
        "name": "MemoQ",
        "simple": False,
        "row": 0,
        "source_col_index": 1,
        "target_col_index": 2,
    },
    {
        "name": "Memsouce",
        "simple": False,
        "row": 0,
        "source_col_index": 3,
        "target_col_index": 4,
    },
    {
        "name": "Target General",
        "simple": True,
    },
]

error_type_color_map = {
    "uncategorized": Qt.GlobalColor.magenta,
    "misspelling": Qt.GlobalColor.red,
    "style": Qt.GlobalColor.blue,
}


if __name__ == "__main__":
    # Set the organization name and application name for QSettings
    QApplication.setOrganizationName("Andre Jonas")
    QApplication.setApplicationName("pyLanguageTool")

    # Set the user config directory
    config_dir = QSettings().fileName()
    config_dir = os.path.dirname(config_dir)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, config_dir)

    app = QApplication(sys.argv)

    from text_editor import TextEditor

    editor = TextEditor()
    sys.exit(app.exec())
