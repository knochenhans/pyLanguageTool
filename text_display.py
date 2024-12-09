from PySide6.QtWidgets import QTextEdit, QApplication
from PySide6.QtCore import Qt


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
