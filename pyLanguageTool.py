import io
import os
import re
import sys
import tempfile
from pathlib import Path

import aspose.words as aw

# import debugpy
import language_tool_python
from colorama import Fore, Style
from docx import Document, table
from docx.shared import Cm
from PySide6.QtCore import QSettings, Qt, QThread, Signal
from PySide6.QtGui import (
    QAction,
    QFont,
    QTextCharFormat,
    QTextCursor,
    QTextDocument,
    QColor,
    QTextBlock,
)
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QMainWindow,
    QSplitter,
    QStyle,
    QTextEdit,
    QInputDialog,
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
        openAction.setShortcut("Ctrl+O")
        openAction.setStatusTip("Open File")
        openAction.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton)
        )
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

        self.statusBar()

        # Add an icon bar
        self.toolbar = self.addToolBar("Toolbar")
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, self.toolbar)
        self.toolbar.addAction(openAction)

        # Maximize the window
        # self.showMaximized()
        self.resize(1200, 800)

        self.loadWindowPosition()

        self.show()

        self.language_tool = language_tool_python.LanguageTool("de-DE")

        self.errors: dict = {}

        self.error_type_color_map = {
            "uncategorized": Qt.GlobalColor.magenta,
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

    def printError(self, cursor: QTextCursor, error: dict):
        # Get color based on error type
        error_type = error["Error"].split(" - ")[0]
        color = QColor(self.error_type_color_map.get(error_type, Qt.GlobalColor.yellow))

        for field_name, field_value in error.items():
            if field_name == "Offset" or field_name == "Length":
                continue
            # Ignore replacements if there are none
            if field_name == "Replacements" and field_value != [" "]:
                continue

            block_format = cursor.blockFormat()
            # block_format.setTopMargin(10)
            block_format.setBackground(color.lighter(190))
            # block_format.setForeground(Qt.GlobalColor.black)

            cursor.setBlockFormat(block_format)
            cursor.insertBlock()
            format = QTextCharFormat()

            format.setForeground(Qt.GlobalColor.black)
            format.setFontWeight(QFont.Weight.Bold)

            cursor.insertText(f"{field_name}: ", format)

            format.setFontWeight(QFont.Weight.Normal)
            # format = QTextCharFormat()
            # format.setForeground(Qt.GlobalColor.black)

            cursor.insertText(f"{field_value}\n", format)

    def handleFileLoaded(self, text: str):
        cursor = QTextCursor(self.errorDisplay.document())

        for error in self.errors.values():
            self.printError(cursor, error)

        block_format = cursor.blockFormat()
        block_format.setBackground(QColor(Qt.GlobalColor.white))
        cursor.insertBlock()
        cursor.insertText("\n")

        formatted_text = self.formatText(text)
        self.textDisplay.setDocument(formatted_text)

        self.addRecentFile(self.fileLoaderWorker.file_name)
        self.statusBar().showMessage("File loaded")

    def openFile(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File")
        if file_name:
            self.fileLoaderWorker = FileLoaderWorker(self, file_name)
            self.fileLoaderWorker.fileLoaded.connect(self.handleFileLoaded)
            self.fileLoaderWorker.start()

    def formatText(self, text):
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
                    f"{error['Error']}\n{error['Message']}\n→ {error['Replacements']}\nContext: {error['Context']}\n{error['Sentence']}"
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


class FileLoaderWorker(QThread):
    fileLoaded = Signal(str)

    def __init__(self, text_editor: TextEditor, file_name: str):
        super().__init__()
        self.text_editor = text_editor
        self.file_name = file_name

    def run(self):
        # debugpy.debug_this_thread()
        self.text_editor.statusBar().showMessage(f"Loading {self.file_name}...")
        with open(self.file_name, "r") as file:
            if (
                self.file_name.endswith(".docx")
                or self.file_name.endswith(".doc")
                or self.file_name.endswith(".rtf")
            ):
                file_path = None

                if self.file_name.endswith(".rtf"):
                    # Load file as bytesio
                    with open(self.file_name, "rb") as f:
                        data = f.read()

                    stream = io.BytesIO(data)
                    doc = aw.Document(stream)

                    # Save as docx
                    stream = io.BytesIO()
                    doc.save(stream, aw.SaveFormat.DOCX)
                    stream.seek(0)

                    file_path = stream

                elif self.file_name.endswith(".docx") or self.file_name.endswith(
                    ".doc"
                ):
                    file_path = self.file_name

                tables = self.text_editor.read_docx_tables(file_path)

                text = ""

                table_names = [f"Table {i+1}" for i in range(len(tables))]

                # Display a message box to select the table that should be checked
                table_name = self.select_table(table_names, tables)

                # Display a message box to select the column that should be checked

                column_names = [
                    f"Column {i+1}" for i in range(len(tables[0].rows[0].cells))
                ]
                
                column_name = self.select_column(column_names)

                table_index = table_names.index(table_name)
                table = tables[table_index]

                column_index = column_names.index(column_name)

                target1 = self.text_editor.extract_table_columns(table, [column_index])[
                    0
                ]

                # _, _, target1 = self.text_editor.extract_table_columns(
                #     table, [0, 1, 2]
                # )

                text = "\n".join(target1)
            else:
                text = file.read()

            self.text_editor.statusBar().showMessage(
                f"Checking {self.file_name} with LanguageTool..."
            )
            matches = self.text_editor.language_tool.check(text)
            self.text_editor.language_tool.close()
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

        self.fileLoaded.emit(text)

    def select_column(self, column_names):
        column_name, ok = QInputDialog.getItem(
                    self.text_editor,
                    "Select Column",
                    "Select the column to check:",
                    column_names,
                    0,
                    False,
                )
        
        return column_name

    def select_table(self, table_names, tables):
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

        return table_name


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
