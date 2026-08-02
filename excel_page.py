from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from excel_tools import (
    apply_date_format,
    apply_text_cleanup,
    refresh_open_workbooks,
    refresh_worksheets,
)
from no_wheel_combo import NoWheelComboBox


class ExcelPage(QWidget):
    def __init__(self):
        super().__init__()

        self.workbook_dropdown = NoWheelComboBox()
        self.worksheet_dropdown = NoWheelComboBox()

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

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

        title = QLabel("Excel Tools")
        title.setObjectName("pageTitle")

        description = QLabel(
            "Choose an open workbook and worksheet, "
            "then apply a tool."
        )
        description.setObjectName("description")
        description.setWordWrap(True)

        layout.addWidget(title)
        layout.addWidget(description)
        layout.addSpacing(8)
        layout.addWidget(self._create_workbook_card())
        layout.addWidget(self._create_date_format_card())
        layout.addWidget(self._create_text_cleanup_card())
        layout.addStretch()

        scroll_area.setWidget(content)
        outer_layout.addWidget(scroll_area)

        self.workbook_dropdown.currentIndexChanged.connect(
            self._workbook_changed
        )

        refresh_open_workbooks(
            self.workbook_dropdown,
            self.worksheet_dropdown,
            self,
            show_warning=False,
        )

    def _create_workbook_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(245)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Open Excel workbook")
        title.setObjectName("cardTitle")

        text = QLabel(
            "ToolPy only applies Excel tools to the workbook "
            "and worksheet selected below."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        workbook_label = QLabel("Workbook")
        workbook_label.setObjectName("cardText")

        workbook_row = QHBoxLayout()
        workbook_row.setSpacing(10)

        self.workbook_dropdown.addItem(
            "Choose an open Excel workbook",
            "",
        )
        self.workbook_dropdown.setMinimumHeight(42)

        refresh_button = QPushButton("Refresh")
        refresh_button.setObjectName("secondaryButton")
        refresh_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        refresh_button.setMinimumHeight(42)
        refresh_button.setFixedWidth(110)
        refresh_button.clicked.connect(
            lambda: refresh_open_workbooks(
                self.workbook_dropdown,
                self.worksheet_dropdown,
                self,
                show_warning=True,
            )
        )

        workbook_row.addWidget(
            self.workbook_dropdown,
            1,
        )
        workbook_row.addWidget(refresh_button)

        worksheet_label = QLabel("Worksheet")
        worksheet_label.setObjectName("cardText")

        self.worksheet_dropdown.addItem(
            "Choose a worksheet",
            "",
        )
        self.worksheet_dropdown.setMinimumHeight(42)

        layout.addWidget(title)
        layout.addWidget(text)
        layout.addWidget(workbook_label)
        layout.addLayout(workbook_row)
        layout.addWidget(worksheet_label)
        layout.addWidget(self.worksheet_dropdown)

        return card

    def _create_date_format_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(430)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Date Format")
        title.setObjectName("cardTitle")

        text = QLabel(
            "Enter the first date cell. ToolPy reads downward in "
            "that column and stops at the first blank cell."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        start_cell_label = QLabel("Start cell")
        start_cell_label.setObjectName("cardText")

        self.start_cell_input = QLineEdit("A2")
        self.start_cell_input.setPlaceholderText("Example: A2")
        self.start_cell_input.setMinimumHeight(42)
        self.start_cell_input.setMaxLength(8)

        order_label = QLabel("Date order")
        order_label.setObjectName("cardText")

        self.date_order_dropdown = NoWheelComboBox()
        self.date_order_dropdown.addItems(
            [
                "Auto Detect",
                "MM/DD/YYYY",
                "DD/MM/YYYY",
            ]
        )
        self.date_order_dropdown.setMinimumHeight(42)

        output_label = QLabel("Output format")
        output_label.setObjectName("cardText")

        self.output_format_dropdown = NoWheelComboBox()
        self.output_format_dropdown.addItems(
            [
                "MM/DD/YYYY",
                "DD/MM/YYYY",
                "Month DD, YYYY",
                "MMM DD, YYYY",
                "YYYY-MM-DD",
            ]
        )
        self.output_format_dropdown.setMinimumHeight(42)

        alignment_label = QLabel("Alignment")
        alignment_label.setObjectName("cardText")

        self.date_alignment_dropdown = NoWheelComboBox()
        self.date_alignment_dropdown.addItems(
            [
                "No Change",
                "Left",
                "Center",
                "Right",
            ]
        )
        self.date_alignment_dropdown.setMinimumHeight(42)

        apply_button = QPushButton("Apply")
        apply_button.setObjectName("actionButton")
        apply_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        apply_button.setFixedSize(135, 44)
        apply_button.clicked.connect(
            self._apply_date_format
        )

        layout.addWidget(title)
        layout.addWidget(text)
        layout.addWidget(start_cell_label)
        layout.addWidget(self.start_cell_input)
        layout.addWidget(order_label)
        layout.addWidget(self.date_order_dropdown)
        layout.addWidget(output_label)
        layout.addWidget(self.output_format_dropdown)
        layout.addWidget(alignment_label)
        layout.addWidget(self.date_alignment_dropdown)
        layout.addSpacing(8)
        layout.addWidget(
            apply_button,
            alignment=Qt.AlignmentFlag.AlignLeft,
        )

        return card

    def _create_text_cleanup_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(340)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Text Cleanup")
        title.setObjectName("cardTitle")

        text = QLabel(
            "Enter the first text cell. ToolPy cleans downward in "
            "that column and stops at the first blank cell."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        start_cell_label = QLabel("Start cell")
        start_cell_label.setObjectName("cardText")

        self.text_start_cell_input = QLineEdit("A2")
        self.text_start_cell_input.setPlaceholderText("Example: A2")
        self.text_start_cell_input.setMinimumHeight(42)
        self.text_start_cell_input.setMaxLength(8)

        action_label = QLabel("Action")
        action_label.setObjectName("cardText")

        self.text_action_dropdown = NoWheelComboBox()
        self.text_action_dropdown.addItems(
            [
                "Trim Spaces",
                "Proper Case",
                "UPPERCASE",
                "lowercase",
                "Sentence case",
            ]
        )
        self.text_action_dropdown.setMinimumHeight(42)

        alignment_label = QLabel("Alignment")
        alignment_label.setObjectName("cardText")

        self.text_alignment_dropdown = NoWheelComboBox()
        self.text_alignment_dropdown.addItems(
            [
                "No Change",
                "Left",
                "Center",
                "Right",
            ]
        )
        self.text_alignment_dropdown.setMinimumHeight(42)

        apply_button = QPushButton("Apply")
        apply_button.setObjectName("actionButton")
        apply_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        apply_button.setFixedSize(135, 44)
        apply_button.clicked.connect(
            self._apply_text_cleanup
        )

        layout.addWidget(title)
        layout.addWidget(text)
        layout.addWidget(start_cell_label)
        layout.addWidget(self.text_start_cell_input)
        layout.addWidget(action_label)
        layout.addWidget(self.text_action_dropdown)
        layout.addWidget(alignment_label)
        layout.addWidget(self.text_alignment_dropdown)
        layout.addSpacing(8)
        layout.addWidget(
            apply_button,
            alignment=Qt.AlignmentFlag.AlignLeft,
        )

        return card

    def _workbook_changed(self):
        refresh_worksheets(
            self.workbook_dropdown,
            self.worksheet_dropdown,
            self,
            show_warning=False,
        )

    def _apply_date_format(self):
        apply_date_format(
            self,
            self.workbook_dropdown,
            self.worksheet_dropdown,
            self.start_cell_input.text(),
            self.date_order_dropdown.currentText(),
            self.output_format_dropdown.currentText(),
            self.date_alignment_dropdown.currentText(),
        )

    def _apply_text_cleanup(self):
        apply_text_cleanup(
            self,
            self.workbook_dropdown,
            self.worksheet_dropdown,
            self.text_start_cell_input.text(),
            self.text_action_dropdown.currentText(),
            self.text_alignment_dropdown.currentText(),
        )
