import sys

from PySide6.QtWidgets import QApplication

from theme import APP_STYLE
from word_page import ToolPyWindow


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)

    window = ToolPyWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
