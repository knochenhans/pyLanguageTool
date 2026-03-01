import re
from typing import Any, Dict

import language_tool_python  # type: ignore
from colorama import Fore, Style  # type: ignore
from PySide6.QtCore import QEvent, QObject, QSettings, Qt
from PySide6.QtGui import (
    QAction,
    QColor,
    QFont,
    QTextCharFormat,
    QTextCursor,
    QTextDocument,
)
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QMainWindow,
    QSplitter,
    QStyle,
    QTextEdit,
)

from file_loader_worker import FileLoaderWorker
from text_display import TextDisplay


class TextEditor(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.recentFiles: list[str] = []
        self.initUI()

    def initUI(self) -> None:
        self.splitter = QSplitter()
        self.splitter.setOrientation(Qt.Orientation.Horizontal)
        self.setCentralWidget(self.splitter)

        self.text_display = TextDisplay()
        self.text_display.setAcceptRichText(True)
        self.text_display.setStyleSheet("background-color: white; color: black;")
        self.text_display.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextEditorInteraction
        )
        self.splitter.addWidget(self.text_display)
        # allow handling clicks / Escape to clear selection
        self.text_display.installEventFilter(self)
        # track mouse press to distinguish click vs drag
        self._mouse_press_pos = None
        self._mouse_moved = False

        self.error_display = QTextEdit()
        self.error_display.setAcceptRichText(True)
        self.error_display.setReadOnly(True)
        self.splitter.addWidget(self.error_display)

        self.error_display.setText("Errors will be displayed here.")

        cursor = QTextCursor(self.error_display.document())
        cursor.setPosition(0)
        cursor.insertHtml("<b>")
        # cursor.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 5)
        cursor.setPosition(5)
        cursor.insertHtml("</b>")

        # Load recent files from QSettings
        recent_files = QSettings().value("recentFiles", [])

        # Check if this is a list of strings or a single string
        if isinstance(recent_files, str):
            self.recentFiles = [recent_files]

        open_action = QAction("Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.setStatusTip("Open File")
        open_action.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton)
        )
        open_action.triggered.connect(self.openFile)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)

        check_action = QAction("Check Text", self)
        check_action.setShortcut("F5")
        check_action.triggered.connect(self.checkText)

        clear_action = QAction("Clear Text", self)
        clear_action.setShortcut("F6")
        clear_action.triggered.connect(self.text_display.clear)

        open_latest_recent_file_action = QAction("Open Latest Recent File", self)
        open_latest_recent_file_action.triggered.connect(self.open_latest_recent_file)
        open_latest_recent_file_action.setStatusTip("Open the latest recent file")
        open_latest_recent_file_action.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton)
        )
        open_latest_recent_file_action.setShortcut("Ctrl+Shift+O")
        self.addAction(open_latest_recent_file_action)

        menubar = self.menuBar()
        file_menu = menubar.addMenu("File")
        file_menu.addAction(open_action)
        file_menu.addAction(open_latest_recent_file_action)

        self.recentMenu = file_menu.addMenu("Recent Files")
        self.updateRecentFilesMenu()

        file_menu.addAction(check_action)
        file_menu.addAction(clear_action)
        file_menu.addSeparator()

        preferences_action = QAction("Preferences", self)
        preferences_action.triggered.connect(self.openPreferences)

        file_menu.addAction(preferences_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)

        self.setWindowTitle("pyLanguageTool")

        self.statusBar()

        # Add an icon bar
        self.toolbar = self.addToolBar("Toolbar")
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, self.toolbar)
        self.toolbar.addAction(open_action)

        # Add template dropdown
        from pyLanguageTool import templates

        self.template_combo_box = QComboBox()
        self.template_combo_box.addItems(
            [str(template["name"]) for template in templates]
        )
        self.template_combo_box.currentIndexChanged.connect(self.templateChanged)
        self.toolbar.addWidget(self.template_combo_box)

        # Add checkbox "Remove tags"
        self.remove_tags_check_box = QCheckBox("Remove tags")
        self.toolbar.addWidget(self.remove_tags_check_box)

        # Add language selection dropdown
        self.language_combo_box = QComboBox()
        self.language_codes = {
            "German (de-DE)": "de-DE",
            "English (en-US)": "en-US",
            "English (en-GB)": "en-GB",
            "French (fr-FR)": "fr-FR",
            "Spanish (es-ES)": "es-ES",
        }
        self.language_combo_box.addItems(list(self.language_codes.keys()))
        self.language_combo_box.setCurrentText("German (de-DE)")
        self.toolbar.addWidget(self.language_combo_box)

        # Maximize the window
        # self.showMaximized()
        self.resize(1200, 800)

        self.loadWindowPosition()

        self.show()

        # Initialize with default language
        self.language_tool = language_tool_python.LanguageTool(
            self.language_codes[self.language_combo_box.currentText()]
        )

        self.errors: Dict[int, Dict[str, Any]] = {}

        self.current_template = templates[0]

    def openPreferences(self) -> None:
        from preferences_window import PreferencesWindow

        preferencesWindow = PreferencesWindow(self)
        preferencesWindow.exec()

    def templateChanged(self, index: int) -> None:
        from pyLanguageTool import templates

        template_name = self.template_combo_box.currentText()
        template = next(
            (template for template in templates if template["name"] == template_name),
            None,
        )
        if template:
            self.current_template = template

    def closeEvent(self, event: Any) -> None:
        self.saveWindowPosition()

        QSettings().setValue("recentFiles", self.recentFiles)

        event.accept()

    def checkText(self) -> None:
        self.statusBar().showMessage("Checking text with LanguageTool...")

        # Get selected language and re-initialize language_tool
        selected_language = self.language_codes[self.language_combo_box.currentText()]
        self.language_tool = language_tool_python.LanguageTool(selected_language)

        text = self.text_display.toPlainText()

        if self.remove_tags_check_box.isChecked():
            text = re.sub(r"<.*?>", "", text)

        matches = self.language_tool.check(text)
        self.errors = {}
        for match in matches:
            print(f"{Fore.RED}Error: {match.message}{Style.RESET_ALL}")
            error_type = f"{match.rule_issue_type} - {match.category}"
            error: Dict[str, Any] = {
                "Error": error_type,
                "Message": match.message,
                "Replacements": match.replacements,
                "Context": match.context,
                "Sentence": match.sentence,
                "Offset": match.offset,
                "Length": match.error_length,
            }
            self.errors[match.offset] = error

        cursor = QTextCursor(self.error_display.document())

        for error in self.errors.values():
            self.printError(cursor, error)

        block_format = cursor.blockFormat()
        block_format.setBackground(QColor(Qt.GlobalColor.white))
        cursor.insertBlock()
        cursor.insertText("\n")

        formatted_text = self.formatText(text)
        self.text_display.setDocument(formatted_text)

        self.statusBar().showMessage("Text checked")

    def fileLoaded(self, text: str) -> None:
        self.text_display.setPlainText(text)
        self.checkText()

        self.addRecentFile(self.fileLoaderWorker.file_name)
        self.statusBar().showMessage("File loaded")

    def openFile(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File")
        if file_name:
            from file_loader_worker import FileLoaderWorker

            self.fileLoaderWorker = FileLoaderWorker(self, file_name)
            self.fileLoaderWorker.fileLoaded.connect(self.fileLoaded)
            self.fileLoaderWorker.start()

    def formatText(self, text: str) -> QTextDocument:
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

    def addRecentFile(self, file_name: str) -> None:
        if file_name in self.recentFiles:
            self.recentFiles.remove(file_name)
        self.recentFiles.insert(0, file_name)
        self.updateRecentFilesMenu()

    def openRecentFile(self, file_name: str) -> None:
        if file_name:
            self.fileLoaderWorker = FileLoaderWorker(self, file_name)
            self.fileLoaderWorker.fileLoaded.connect(self.fileLoaded)
            self.fileLoaderWorker.start()
        self.addRecentFile(file_name)

    def updateRecentFilesMenu(self) -> None:
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

    def loadWindowPosition(self) -> None:
        settings = QSettings()
        pos = settings.value("window/position", self.pos())
        size = settings.value("window/size", self.size())
        self.move(pos)
        self.resize(size)

    def printError(self, cursor: QTextCursor, error: Dict[str, Any]) -> None:
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

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:  # type: ignore[override]
        # Only handle events for the text display
        if obj is self.text_display:
            if event.type() == QEvent.MouseButtonPress:
                # record press position, don't change cursor (preserve drag selection)
                try:
                    self._mouse_press_pos = event.pos()
                    self._mouse_moved = False
                except Exception:
                    self._mouse_press_pos = None
                return False
            if event.type() == QEvent.MouseMove:
                # mark moved if user is dragging
                try:
                    if self._mouse_press_pos is not None:
                        if (event.pos() - self._mouse_press_pos).manhattanLength() > 4:
                            self._mouse_moved = True
                except Exception:
                    pass
                return False
            if event.type() == QEvent.MouseButtonRelease:
                # if it was a simple click (no drag) clear selection
                try:
                    if self._mouse_press_pos is not None and not self._mouse_moved:
                        cursor = self.text_display.textCursor()
                        if cursor.hasSelection():
                            cursor.clearSelection()
                            self.text_display.setTextCursor(cursor)
                            # consume the release so no further toggling happens
                            return True
                except Exception:
                    pass
                finally:
                    self._mouse_press_pos = None
                    self._mouse_moved = False
                return False
            if event.type() == QEvent.KeyPress:
                if getattr(event, "key", lambda: None)() == Qt.Key.Key_Escape:
                    cursor = self.text_display.textCursor()
                    if cursor.hasSelection():
                        cursor.clearSelection()
                        self.text_display.setTextCursor(cursor)
                        return True
        return super().eventFilter(obj, event)

    def open_latest_recent_file(self) -> None:
        # Load first recent file from Linux desktop if available using Gtk.RecentlyUsed API
        try:
            import gi

            gi.require_version("Gtk", "3.0")
            from gi.repository import Gtk

            recently_used = Gtk.RecentManager.get_default()
            items = recently_used.get_items()
            if items:
                first_item = items[0]
                file_path = first_item.get_uri()
                if file_path.startswith("file://"):
                    file_path = file_path[7:]
                self.openRecentFile(file_path)
        except Exception as e:
            print(f"Error loading recent file from desktop: {e}")
