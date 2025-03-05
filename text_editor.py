from PySide6.QtGui import (
    QTextCursor,
    QTextCharFormat,
    QColor,
    QFont,
    QAction,
    QTextDocument,
)
from PySide6.QtCore import Qt, QSettings
import language_tool_python
import re

from PySide6.QtWidgets import (
    QMainWindow,
    QSplitter,
    QTextEdit,
    QStyle,
    QMenuBar,
    QComboBox,
    QCheckBox,
    QFileDialog,
    QToolBar,
)

from file_loader_worker import FileLoaderWorker
from text_display import TextDisplay
from colorama import Fore, Style  # type: ignore


class TextEditor(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.recentFiles: list[str] = []
        self.initUI()

    def initUI(self) -> None:
        self.splitter = QSplitter()
        self.splitter.setOrientation(Qt.Orientation.Horizontal)
        self.setCentralWidget(self.splitter)

        self.textDisplay = TextDisplay()
        self.textDisplay.setAcceptRichText(True)
        self.textDisplay.setStyleSheet("background-color: white; color: black;")
        self.textDisplay.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextEditorInteraction
            | Qt.TextInteractionFlag.TextBrowserInteraction
        )
        self.splitter.addWidget(self.textDisplay)

        self.errorDisplay = QTextEdit()
        self.errorDisplay.setAcceptRichText(True)
        self.errorDisplay.setReadOnly(True)
        self.splitter.addWidget(self.errorDisplay)

        self.errorDisplay.setText("Errors will be displayed here.")

        cursor = QTextCursor(self.errorDisplay.document())
        cursor.setPosition(0)
        cursor.insertHtml("<b>")
        # cursor.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 5)
        cursor.setPosition(5)
        cursor.insertHtml("</b>")

        # Load recent files from QSettings
        recentFiles = QSettings().value("recentFiles", [])

        # Check if this is a list of strings or a single string
        if isinstance(recentFiles, str):
            self.recentFiles = [recentFiles]

        openAction = QAction("Open", self)
        openAction.setShortcut("Ctrl+O")
        openAction.setStatusTip("Open File")
        openAction.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton)
        )
        openAction.triggered.connect(self.openFile)

        exitAction = QAction("Exit", self)
        exitAction.triggered.connect(self.close)

        checkAction = QAction("Check Text", self)
        checkAction.setShortcut("F5")
        checkAction.triggered.connect(self.checkText)

        clearAction = QAction("Clear Text", self)
        clearAction.setShortcut("F6")
        clearAction.triggered.connect(self.textDisplay.clear)

        menubar = self.menuBar()
        fileMenu = menubar.addMenu("File")
        fileMenu.addAction(openAction)

        self.recentMenu = fileMenu.addMenu("Recent Files")
        self.updateRecentFilesMenu()

        fileMenu.addAction(checkAction)
        fileMenu.addAction(clearAction)
        fileMenu.addSeparator()

        preferencesAction = QAction("Preferences", self)
        preferencesAction.triggered.connect(self.openPreferences)

        fileMenu.addAction(preferencesAction)
        fileMenu.addSeparator()
        fileMenu.addAction(exitAction)

        self.setWindowTitle("pyLanguageTool")

        self.statusBar()

        # Add an icon bar
        self.toolbar = self.addToolBar("Toolbar")
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, self.toolbar)
        self.toolbar.addAction(openAction)

        # Add template dropdown
        from pyLanguageTool import templates

        self.templateComboBox = QComboBox()
        self.templateComboBox.addItems(
            [str(template["name"]) for template in templates]
        )
        self.templateComboBox.currentIndexChanged.connect(self.templateChanged)
        self.toolbar.addWidget(self.templateComboBox)

        # Add checkbox "Remove tags"
        self.removeTagsCheckBox = QCheckBox("Remove tags")
        self.toolbar.addWidget(self.removeTagsCheckBox)

        # Maximize the window
        # self.showMaximized()
        self.resize(1200, 800)

        self.loadWindowPosition()

        self.show()

        self.language_tool = language_tool_python.LanguageTool("de-DE")

        self.errors: dict = {}

        self.current_template = templates[0]

    def openPreferences(self):
        from preferences_window import PreferencesWindow

        preferencesWindow = PreferencesWindow(self)
        preferencesWindow.exec()

    def templateChanged(self, index):
        from pyLanguageTool import templates

        template_name = self.templateComboBox.currentText()
        template = next(
            (template for template in templates if template["name"] == template_name),
            None,
        )
        if template:
            self.current_template = template

    def closeEvent(self, event):
        self.saveWindowPosition()

        QSettings().setValue("recentFiles", self.recentFiles)

        event.accept()

    def checkText(self):
        self.statusBar().showMessage("Checking text with LanguageTool...")

        text = self.textDisplay.toPlainText()

        if self.removeTagsCheckBox.isChecked():
            text = re.sub(r"<.*?>", "", text)

        matches = self.language_tool.check(text)
        self.errors = {}
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
            self.errors[match.offset] = error

        cursor = QTextCursor(self.errorDisplay.document())

        for error in self.errors.values():
            self.printError(cursor, error)

        block_format = cursor.blockFormat()
        block_format.setBackground(QColor(Qt.GlobalColor.white))
        cursor.insertBlock()
        cursor.insertText("\n")

        formatted_text = self.formatText(text)
        self.textDisplay.setDocument(formatted_text)

        self.statusBar().showMessage("Text checked")

    def fileLoaded(self, text: str):
        self.checkText()

        self.addRecentFile(self.fileLoaderWorker.file_name)
        self.statusBar().showMessage("File loaded")

    def openFile(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File")
        if file_name:
            from file_loader_worker import FileLoaderWorker

            self.fileLoaderWorker = FileLoaderWorker(self, file_name)
            self.fileLoaderWorker.fileLoaded.connect(self.fileLoaded)
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
                from pyLanguageTool import error_type_color_map

                underline_color = error_type_color_map.get(
                    error_type, QColor(0, 0, 0, 128)  # semi-transparent black
                )
                format.setUnderlineColor(underline_color)
                format.setUnderlineStyle(
                    QTextCharFormat.UnderlineStyle.SpellCheckUnderline
                )
                format.setToolTip(
                    f"{error['Error']}\n{error['Message']}\n→ {error['Replacements']}\nContext: {error['Context']}\n{error['Sentence']}"
                )
                format.setAnchor(True)
                format.setAnchorHref(f"#{offset}")
                format.setBackground(QColor(underline_color).lighter(190))
                current_error = error
            elif current_error:
                if offset == error["Offset"] + error["Length"]:
                    format = QTextCharFormat()
                    current_error = None

            cursor.insertText(character, format)
        return document

    def addRecentFile(self, file_name):
        if file_name in self.recentFiles:
            self.recentFiles.remove(file_name)
        self.recentFiles.insert(0, file_name)
        self.updateRecentFilesMenu()

    def openRecentFile(self, file_name):
        if file_name:
            self.fileLoaderWorker = FileLoaderWorker(self, file_name)
            self.fileLoaderWorker.fileLoaded.connect(self.fileLoaded)
            self.fileLoaderWorker.start()
        self.addRecentFile(file_name)

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

    def printError(self, cursor: QTextCursor, error: dict):
        # Get color based on error type
        error_type = error["Error"].split(" - ")[0]
        from pyLanguageTool import error_type_color_map

        color = QColor(error_type_color_map.get(error_type, Qt.GlobalColor.yellow))

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
