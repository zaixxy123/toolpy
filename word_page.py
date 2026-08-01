from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from word_tools import (
    refresh_open_documents,
    resize_images,
    set_all_images_behind_text,
)


class ToolPyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowIcon(QIcon("assets/logo.ico"))

        self.setWindowTitle("ToolPy")
        self.resize(920, 700)
        self.setMinimumSize(760, 540)

        self.document_dropdown = QComboBox()

        self._build_ui()

        refresh_open_documents(
            self.document_dropdown,
            self,
            show_warning=False,
        )

    def _build_ui(self):
        central_widget = QWidget()

        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        main_layout.addWidget(self._create_sidebar())
        main_layout.addWidget(self._create_word_page(), 1)

        self.setCentralWidget(central_widget)

    def _create_sidebar(self):
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(215)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(12, 0, 12, 16)
        layout.setSpacing(8)

        app_title = QLabel("ToolPy")
        app_title.setObjectName("appTitle")

        layout.addWidget(app_title)
        layout.addWidget(self._nav_button("Word", checked=True))
        layout.addWidget(self._nav_button("Excel"))
        layout.addStretch()

        return sidebar

    def _nav_button(self, text, checked=False):
        button = QPushButton(text)
        button.setObjectName("navButton")
        button.setCheckable(True)
        button.setChecked(checked)
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        button.setMinimumHeight(42)
        return button

    def _create_word_page(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )

        content = QWidget()
        content.setObjectName("contentPage")

        layout = QVBoxLayout(content)
        layout.setContentsMargins(42, 36, 42, 36)
        layout.setSpacing(18)

        page_title = QLabel("Word Tools")
        page_title.setObjectName("pageTitle")

        description = QLabel(
            "Choose a document currently open in Microsoft Word, "
            "then apply a tool."
        )
        description.setObjectName("description")
        description.setWordWrap(True)

        layout.addWidget(page_title)
        layout.addWidget(description)
        layout.addSpacing(8)
        layout.addWidget(self._create_document_card())
        layout.addWidget(self._create_behind_text_card())
        layout.addWidget(self._create_resize_card())
        layout.addStretch()

        scroll_area.setWidget(content)
        return scroll_area

    def _create_document_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(145)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Open Word document")
        title.setObjectName("cardTitle")

        text = QLabel(
            "ToolPy lists documents that are currently open "
            "in Microsoft Word."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        row = QHBoxLayout()
        row.setSpacing(10)

        self.document_dropdown.addItem(
            "Choose an open Word document",
            "",
        )
        self.document_dropdown.setMinimumHeight(42)

        refresh_button = QPushButton("Refresh")
        refresh_button.setObjectName("secondaryButton")
        refresh_button.setCursor(Qt.CursorShape.PointingHandCursor)
        refresh_button.setMinimumHeight(42)
        refresh_button.setFixedWidth(110)
        refresh_button.clicked.connect(
            lambda: refresh_open_documents(
                self.document_dropdown,
                self,
                show_warning=True,
            )
        )

        row.addWidget(self.document_dropdown, 1)
        row.addWidget(refresh_button)

        layout.addWidget(title)
        layout.addWidget(text)
        layout.addLayout(row)

        return card

    def _create_behind_text_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(165)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Set pictures behind text")
        title.setObjectName("cardTitle")

        text = QLabel(
            "Converts inline pictures into movable shapes and places "
            "every picture behind the document text."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        apply_button = QPushButton("Apply")
        apply_button.setObjectName("actionButton")
        apply_button.setCursor(Qt.CursorShape.PointingHandCursor)
        apply_button.setFixedSize(135, 44)
        apply_button.clicked.connect(
            lambda: set_all_images_behind_text(
                self.document_dropdown,
                self,
            )
        )

        layout.addWidget(title)
        layout.addWidget(text)
        layout.addSpacing(8)
        layout.addWidget(
            apply_button,
            alignment=Qt.AlignmentFlag.AlignLeft,
        )

        return card

    def _create_resize_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(320)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Custom image resize")
        title.setObjectName("cardTitle")

        text = QLabel(
            "Resize all images or only the image currently selected "
            "inside Microsoft Word."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        mode_label = QLabel("Mode")
        mode_label.setObjectName("cardText")

        self.resize_mode = QComboBox()
        self.resize_mode.addItems(
            [
                "All images",
                "Selected image only",
            ]
        )
        self.resize_mode.setMinimumHeight(42)

        size_row = QHBoxLayout()
        size_row.setSpacing(12)

        width_column = QVBoxLayout()
        width_column.setSpacing(6)

        width_label = QLabel("Width (cm)")
        width_label.setObjectName("cardText")

        self.width_input = QLineEdit("18.99")
        self.width_input.setMinimumHeight(42)

        width_column.addWidget(width_label)
        width_column.addWidget(self.width_input)

        height_column = QVBoxLayout()
        height_column.setSpacing(6)

        height_label = QLabel("Height (cm)")
        height_label.setObjectName("cardText")

        self.height_input = QLineEdit("22.67")
        self.height_input.setMinimumHeight(42)

        height_column.addWidget(height_label)
        height_column.addWidget(self.height_input)

        size_row.addLayout(width_column)
        size_row.addLayout(height_column)

        apply_button = QPushButton("Apply")
        apply_button.setObjectName("actionButton")
        apply_button.setCursor(Qt.CursorShape.PointingHandCursor)
        apply_button.setFixedSize(135, 44)
        apply_button.clicked.connect(self._apply_resize)

        layout.addWidget(title)
        layout.addWidget(text)
        layout.addWidget(mode_label)
        layout.addWidget(self.resize_mode)
        layout.addLayout(size_row)
        layout.addSpacing(8)
        layout.addWidget(
            apply_button,
            alignment=Qt.AlignmentFlag.AlignLeft,
        )

        return card

    def _apply_resize(self):
        try:
            width_cm = float(self.width_input.text())
            height_cm = float(self.height_input.text())
        except ValueError:
            QMessageBox.warning(
                self,
                "Invalid Size",
                "Enter valid numbers for width and height.",
            )
            return

        if width_cm <= 0 or height_cm <= 0:
            QMessageBox.warning(
                self,
                "Invalid Size",
                "Width and height must be greater than zero.",
            )
            return

        resize_all = self.resize_mode.currentIndex() == 0

        resize_images(
            self.document_dropdown,
            self,
            width_cm,
            height_cm,
            resize_all,
        )
