import io
import os
import re
import sys
import tempfile
from pathlib import Path

import aspose.words as aw
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


class TextDisplay(QTextEdit):
    def __init__(self):
        super().__init__()

    def mouseMoveEvent(self, e):
        self.anchor = self.anchorAt(e.position().toPoint())
        if self.anchor:
            QApplication.setOverrideCursor(Qt.CursorShape.PointingHandCursor)
        else:
            QApplication.setOverrideCursor(Qt.CursorShape.ArrowCursor)

    def mousePressEvent(self, e):
        self.anchor = self.anchorAt(e.position().toPoint())

        if self.anchor:
            print(f"Anchor: {self.anchor}")

        # if self.anchor:
        #     QApplication.setOverrideCursor(Qt.CursorShape.PointingHandCursor)

    def mouseReleaseEvent(self, e):
        # if self.anchor:
            # QDesktopServices.openUrl(QUrl(self.anchor))
            # QApplication.setOverrideCursor(Qt.CursorShape.ArrowCursor)
            # self.anchor = None
        pass

class TextEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.recentFiles: list[str] = []
        self.initUI()

    def initUI(self):
        self.splitter = QSplitter()
        self.splitter.setOrientation(Qt.Orientation.Horizontal)
        self.setCentralWidget(self.splitter)

        self.textDisplay = TextDisplay()
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

        self.errors: dict = {}

        self.error_type_color_map = {
            "uncategorized": Qt.GlobalColor.gray,
            "misspelling": Qt.GlobalColor.red,
            "style": Qt.GlobalColor.blue,
        }

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
                if col < len(row.cells):
                    extracted_columns[j].append(row.cells[col].text.strip())
            if i % 100 == 0:
                print(
                    f"{Fore.GREEN}Current row number: {i}, First column: {Fore.BLUE}{extracted_columns[0][-1]}{Style.RESET_ALL}"
                )
        return extracted_columns

    def printError(self, error: dict):
        cursor = QTextCursor(self.errorDisplay.document())

        for field_name, field_value in error.items():
            if field_name == "Offset" or field_name == "Length":
                continue
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
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File")
        if file_name:
            with open(file_name, "r") as file:
                if (
                    file_name.endswith(".docx")
                    or file_name.endswith(".doc")
                    or file_name.endswith(".rtf")
                ):
                    file_path = None

                    if file_name.endswith(".rtf"):
                        # Load file as bytesio
                        with open(file_name, "rb") as f:
                            data = f.read()

                        stream = io.BytesIO(data)
                        doc = aw.Document(stream)

                        # Save as docx
                        stream = io.BytesIO()
                        doc.save(stream, aw.SaveFormat.DOCX)
                        stream.seek(0)

                        file_path = stream

                    elif file_name.endswith(".docx") or file_name.endswith(".doc"):
                        file_path = file_name

                    tables = self.read_docx_tables(file_path)

                    for table in tables:
                        _, _, target1 = self.extract_table_columns(table, [0, 1, 2])

                        text = "\n".join(target1)
                else:
                    text = file.read()

                matches = self.language_tool.check(text)
                for match in matches:
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
                    self.errors[match.offset] = error
                    self.printError(error)

                formatted_text = self.format(text)
                self.textDisplay.setDocument(formatted_text)
            self.addRecentFile(file_name)

    def format(self, text):
        document = QTextDocument()
        cursor = QTextCursor(document)

        # Display the text in the QTextEdit
        # cursor.insertText(text)

        format = QTextCharFormat()

        # Get a list of all keys (offsets) in the errors dictionary
        keys = list(self.errors.keys())

        current_error = None

        # Loop through the text and add formatting based on the offset and length
        for character in text:
            offset = cursor.position()
            if offset in keys:
                error = self.errors[offset]
                format.setFontUnderline(True)

                error_type = error["Error"].split(" - ")[0]

                format.setUnderlineColor(
                    self.error_type_color_map.get(error_type, Qt.GlobalColor.black)
                )
                format.setUnderlineStyle(
                    QTextCharFormat.UnderlineStyle.SpellCheckUnderline
                )
                format.setToolTip(
                    f"Error: {error['Error']}\nMessage: {error['Message']}\nReplacements: {error['Replacements']}\nContext: {error['Context']}\nSentence: {error['Sentence']}"
                )
                format.setAnchor(True)
                format.setAnchorHref(f"#{offset}")
                current_error = error
            elif current_error:
                if offset == error["Offset"] + error["Length"]:
                    format = QTextCharFormat()
                    current_error = None

            cursor.insertText(character, format)
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
