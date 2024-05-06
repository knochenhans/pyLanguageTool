import os
import re
import sys

import language_tool_python
from colorama import Fore, Style
from docx import Document, table
from docx.shared import Cm
from PySide6.QtCore import QSettings, Qt
from PySide6.QtGui import QAction, QFont, QTextCharFormat, QTextCursor, QTextDocument
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QMainWindow,
    QSplitter,
    QTextEdit,
)


class TextEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.recentFiles: list[str] = []
        self.initUI()

    def initUI(self):
        self.splitter = QSplitter()
        self.splitter.setOrientation(Qt.Orientation.Horizontal)
        self.setCentralWidget(self.splitter)

        self.textDisplay = QTextEdit()
        self.textDisplay.setAcceptRichText(True)
        self.textDisplay.setReadOnly(True)
        self.textDisplay.setStyleSheet("background-color: white;" "color: black;")
        self.splitter.addWidget(self.textDisplay)

        self.errorDisplay = QTextEdit()
        self.errorDisplay.setAcceptRichText(True)
        self.errorDisplay.setReadOnly(True)
        self.splitter.addWidget(self.errorDisplay)

        openAction = QAction("Open", self)
        openAction.triggered.connect(self.openFile)

        exitAction = QAction("Exit", self)
        exitAction.triggered.connect(self.close)

        menubar = self.menuBar()
        fileMenu = menubar.addMenu("File")
        fileMenu.addAction(openAction)
        fileMenu.addAction(exitAction)

        fileMenu.addSeparator()

        # Load recent files from QSettings
        recentFiles = QSettings().value("recentFiles", [])

        # Check if this is a list of strings or a single string
        if isinstance(recentFiles, str):
            recentFiles = [recentFiles]

        self.recentFiles = recentFiles
        self.recentMenu = fileMenu.addMenu("Recent Files")
        self.updateRecentFilesMenu()

        self.setWindowTitle("pyLanguageTool")

        # Maximize the window
        self.showMaximized()

        self.loadWindowPosition()

        self.show()

        self.language_tool = language_tool_python.LanguageTool("de-DE")

    def closeEvent(self, event):
        self.saveWindowPosition()

        QSettings().setValue("recentFiles", self.recentFiles)

        event.accept()

    def read_docx_tables(self, file_path):
        """
        Reads all the tables in a docx file and returns them as a list.
        """
        document = Document(file_path)
        tables = document.tables
        return tables

    def extract_table_columns(self, table, columns, num_rows=-1):
        """
        Extracts the contents of the specified columns of a table up to a specified number of rows.
        """
        extracted_columns = [[] for _ in range(len(columns))]

        if num_rows == -1:
            num_rows = len(table.rows)

        for i, row in enumerate(table.rows):
            if num_rows > -1 and i >= num_rows:
                break
            for j, col in enumerate(columns):
                extracted_columns[j].append(row.cells[col].text.strip())
            if i % 100 == 0:
                print(
                    f"{Fore.GREEN}Current row number: {i}, First column: {Fore.BLUE}{extracted_columns[0][-1]}{Style.RESET_ALL}"
                )
        return extracted_columns

    def addError(self, match: language_tool_python.Match):
        cursor = QTextCursor(self.errorDisplay.document())

        error_type = f"{match.ruleIssueType} - {match.category}"

        fields = {
            "Error": error_type,
            "Message": match.message,
            "Replacements": match.replacements,
            "Context": match.context,
            "Sentence": match.sentence,
            # "Rule ID": match.ruleId,
        }

        for field_name, field_value in fields.items():
            # Ignore replacements if there are none
            if field_name == "Replacements" and field_value != [" "]:
                continue

            format = QTextCharFormat()

            format.setForeground(Qt.GlobalColor.red)
            format.setFontWeight(QFont.Weight.Bold)
            cursor.insertText(f"{field_name}: ", format)

            format = QTextCharFormat()

            cursor.insertText(f"{field_value}\n", format)

        cursor.insertText("\n")

    def openFile(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "Open File")
        if fileName:
            offset_lengths = []

            with open(fileName, "r") as file:
                if fileName.endswith(".docx") or fileName.endswith(".doc"):
                    table1 = self.read_docx_tables(fileName)
                    num_rows = len(table1[0].rows)

                    id1, source1, target1 = self.extract_table_columns(
                        table1, [0, 1, 2]
                    )

                    text = "\n".join(target1)
                else:
                    text = file.read()

                matches = self.language_tool.check(text)
                for match in matches:
                    offset_lengths.append((match.offset, match.errorLength))
                    self.addError(match)

                formatted_text = self.format(text, offset_lengths)
                self.textDisplay.setDocument(formatted_text)
            self.addRecentFile(fileName)

    def format(self, text, offset_lengths=None):
        document = QTextDocument()
        cursor = QTextCursor(document)

        # Display the text in the QTextEdit
        # cursor.insertText(text)

        format = QTextCharFormat()

        # Loop through the text and add formatting based on the offset and length
        for character in text:

            if offset_lengths:
                for offset, length in offset_lengths:
                    if cursor.position() == offset:
                        format.setFontUnderline(True)
                        format.setUnderlineColor(Qt.GlobalColor.blue)
                        format.setUnderlineStyle(
                            QTextCharFormat.UnderlineStyle.SpellCheckUnderline
                        )
                    if cursor.position() == offset + length:
                        format = QTextCharFormat()

            cursor.insertText(character, format)

        # lines = text.splitlines()
        # for line in lines:
        #     words = re.split(r"(\W+)", line)
        #     for word in words:
        #         format = QTextCharFormat()
        #         if import_keyword in word:
        #             format.setForeground(Qt.GlobalColor.red)
        #             format.setFontWeight(QFont.Weight.Bold)
        #         cursor.insertText(word, format)
        #     cursor.insertBlock()
        return document

    def addRecentFile(self, fileName):
        if fileName in self.recentFiles:
            self.recentFiles.remove(fileName)
        self.recentFiles.insert(0, fileName)
        self.updateRecentFilesMenu()

    def openRecentFile(self, fileName):
        with open(fileName, "r") as file:
            self.textDisplay.setText(file.read())
        self.addRecentFile(fileName)

    def updateRecentFilesMenu(self):
        self.recentMenu.clear()
        for i, fileName in enumerate(self.recentFiles):
            action = QAction(f"{i+1}: {fileName}", self)

            action.triggered.connect(
                lambda _, fileName=fileName: self.openRecentFile(fileName)
            )
            self.recentMenu.addAction(action)

    def saveWindowPosition(self):
        settings = QSettings()
        settings.setValue("window/position", self.pos())
        settings.setValue("window/size", self.size())
        settings.setValue("recentFiles", self.recentFiles)

    def loadWindowPosition(self):
        settings = QSettings()
        pos = settings.value("window/position", self.pos())
        size = settings.value("window/size", self.size())
        self.move(pos)
        self.resize(size)


if __name__ == "__main__":
    # Set the organization name and application name for QSettings
    QApplication.setOrganizationName("Andre Jonas")
    QApplication.setApplicationName("pyLanguageTool")

    # Set the user config directory
    config_dir = QSettings().fileName()
    config_dir = os.path.dirname(config_dir)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, config_dir)

    app = QApplication(sys.argv)
    editor = TextEditor()
    sys.exit(app.exec())
