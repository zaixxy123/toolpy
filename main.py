import sys

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from theme import APP_STYLE
from utils import resource_path
from word_page import ToolPyWindow


def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path("assets/logo.ico")))
    app.setStyleSheet(APP_STYLE)

    window = ToolPyWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
