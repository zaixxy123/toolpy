import sys

from PySide6.QtCore import QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from theme import APP_STYLE
from updater import UpdateManager
from utils import resource_path
from word_page import ToolPyWindow


def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(
        QIcon(resource_path("assets/logo.ico"))
    )
    app.setStyleSheet(APP_STYLE)

    window = ToolPyWindow()
    window.show()

    app.update_manager = UpdateManager()

    QTimer.singleShot(
        4000,
        lambda: app.update_manager.check_for_updates(
            window
        ),
    )

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
