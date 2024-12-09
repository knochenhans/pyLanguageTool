from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import QInputDialog
from colorama import Fore, Style

from file_handler import FileHandler


class FileLoaderWorker(QThread):
    fileLoaded = Signal(str)
    def __init__(self, text_editor, file_name: str):
        super().__init__()
        self.text_editor = text_editor
        self.file_name = file_name
        self.file_handler = FileHandler(text_editor)

    def run(self):
        # debugpy.debug_this_thread()
        self.text_editor.statusBar().showMessage(f"Loading {self.file_name}...")
        text = self.file_handler.load_file(self.file_name)
        self.text_editor.statusBar().showMessage(
            f"Checking {self.file_name} with LanguageTool..."
        )
        self.checkSegment(text)
        self.fileLoaded.emit(text)

    def select_column(self, column_names) -> int:
        column_name, ok = QInputDialog.getItem(
            self.text_editor,
            "Select Column",
            "Select the column to check:",
            column_names,
            0,
            False,
        )

        column_index = column_names.index(column_name)
        return column_index

    def select_table(self, table_names, tables) -> int:
        if len(tables) > 1:
            table_name, ok = QInputDialog.getItem(
                self.text_editor,
                "Select Table",
                "Select the table to check:",
                table_names,
                0,
                False,
            )
            if not ok:
                table_name = table_names[0]
        else:
            table_name = table_names[0]

        table_index = table_names.index(table_name) + 1
        return table_index

    def checkSegment(self, segment: str):
        matches = self.text_editor.language_tool.check(segment)
        for match in matches:
            print(f"{Fore.RED}Error: {match.message}{Style.RESET_ALL}")
            error_type = f"{match.ruleIssueType} - {match.category}"
            error = {
                "Error": error_type,
                "Message": match.message,
                "Replacements": match.replacements,
                "Context": match.context,
                "Sentence": match.sentence,
                "Offset": match.offset,
                "Length": match.errorLength,
            }
            self.text_editor.errors[match.offset] = error
