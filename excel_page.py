from PySide6.QtCore import Qt, QTimer
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
    activate_excel_selection_target,
    apply_date_format,
    apply_quick_calculate,
    apply_text_cleanup,
    read_active_excel_cell,
    refresh_open_workbooks,
    refresh_worksheets,
)
from no_wheel_combo import NoWheelComboBox


class ExcelPage(QWidget):
    def __init__(self):
        super().__init__()

        self.workbook_dropdown = NoWheelComboBox()
        self.worksheet_dropdown = NoWheelComboBox()

        self._selection_timer = QTimer(self)
        self._selection_timer.setInterval(150)
        self._selection_timer.timeout.connect(
            self._poll_excel_selection
        )
        self._selection_context = None
        self._selection_target_input = None
        self._selection_range_mode = False
        self._selection_start_address = None
        self._selection_last_address = None

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
        layout.addWidget(self._create_quick_calculate_card())
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

        date_start_row = QHBoxLayout()
        date_start_row.setSpacing(10)

        date_use_cell_button = QPushButton("Select Cell")
        date_use_cell_button.setObjectName("secondaryButton")
        date_use_cell_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        date_use_cell_button.setMinimumHeight(42)
        date_use_cell_button.setFixedWidth(100)
        date_use_cell_button.clicked.connect(
            lambda: self._begin_cell_selection(
                self.start_cell_input,
                range_mode=False,
            )
        )

        date_start_row.addWidget(
            self.start_cell_input,
            1,
        )
        date_start_row.addWidget(
            date_use_cell_button
        )

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
        layout.addLayout(date_start_row)
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

        text_start_row = QHBoxLayout()
        text_start_row.setSpacing(10)

        text_use_cell_button = QPushButton("Select Cell")
        text_use_cell_button.setObjectName("secondaryButton")
        text_use_cell_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        text_use_cell_button.setMinimumHeight(42)
        text_use_cell_button.setFixedWidth(100)
        text_use_cell_button.clicked.connect(
            lambda: self._begin_cell_selection(
                self.text_start_cell_input,
                range_mode=False,
            )
        )

        text_start_row.addWidget(
            self.text_start_cell_input,
            1,
        )
        text_start_row.addWidget(
            text_use_cell_button
        )

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
        layout.addLayout(text_start_row)
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


    def _create_quick_calculate_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(540)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Quick Calculate")
        title.setObjectName("cardTitle")

        text = QLabel(
            "Calculate numbers downward or to the right. "
            "Auto mode stops at the first blank cell. "
            "Range mode uses an exact selected range."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        mode_label = QLabel("Mode")
        mode_label.setObjectName("cardText")

        self.calculate_mode_dropdown = NoWheelComboBox()
        self.calculate_mode_dropdown.addItems(
            [
                "Auto",
                "Range",
            ]
        )
        self.calculate_mode_dropdown.setMinimumHeight(42)
        self.calculate_mode_dropdown.currentTextChanged.connect(
            self._quick_calculate_mode_changed
        )

        start_label = QLabel("Start cell")
        start_label.setObjectName("cardText")

        self.calculate_start_input = QLineEdit("A2")
        self.calculate_start_input.setPlaceholderText(
            "Example: A2"
        )
        self.calculate_start_input.setMinimumHeight(42)
        self.calculate_start_input.setMaxLength(8)

        end_label = QLabel("End cell")
        end_label.setObjectName("cardText")
        self.calculate_end_label = end_label

        self.calculate_end_input = QLineEdit("A5")
        self.calculate_end_input.setPlaceholderText(
            "Example: A5"
        )
        self.calculate_end_input.setMinimumHeight(42)
        self.calculate_end_input.setMaxLength(8)
        self.calculate_end_input.setEnabled(False)
        self.calculate_end_label.setEnabled(False)

        self.calculate_select_cells_button = QPushButton(
            "Select Cell"
        )
        self.calculate_select_cells_button.setObjectName(
            "secondaryButton"
        )
        self.calculate_select_cells_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        self.calculate_select_cells_button.setMinimumHeight(42)
        self.calculate_select_cells_button.setFixedWidth(135)
        self.calculate_select_cells_button.clicked.connect(
            self._select_quick_calculate_cells
        )

        direction_label = QLabel("Direction")
        direction_label.setObjectName("cardText")

        self.calculate_direction_dropdown = NoWheelComboBox()
        self.calculate_direction_dropdown.addItems(
            [
                "Down",
                "Right",
            ]
        )
        self.calculate_direction_dropdown.setMinimumHeight(42)

        operation_label = QLabel("Operation")
        operation_label.setObjectName("cardText")

        self.calculate_operation_dropdown = NoWheelComboBox()
        self.calculate_operation_dropdown.addItems(
            [
                "Sum",
                "Average",
                "Count",
            ]
        )
        self.calculate_operation_dropdown.setMinimumHeight(42)

        alignment_label = QLabel("Alignment")
        alignment_label.setObjectName("cardText")

        self.calculate_alignment_dropdown = NoWheelComboBox()
        self.calculate_alignment_dropdown.addItems(
            [
                "No Change",
                "Left",
                "Center",
                "Right",
            ]
        )
        self.calculate_alignment_dropdown.setMinimumHeight(42)

        apply_button = QPushButton("Apply")
        apply_button.setObjectName("actionButton")
        apply_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        apply_button.setFixedSize(135, 44)
        apply_button.clicked.connect(
            self._apply_quick_calculate
        )

        layout.addWidget(title)
        layout.addWidget(text)

        layout.addWidget(mode_label)
        layout.addWidget(self.calculate_mode_dropdown)

        layout.addWidget(start_label)
        layout.addWidget(self.calculate_start_input)

        layout.addWidget(self.calculate_end_label)
        layout.addWidget(self.calculate_end_input)

        layout.addWidget(
            self.calculate_select_cells_button,
            alignment=Qt.AlignmentFlag.AlignLeft,
        )

        layout.addWidget(direction_label)
        layout.addWidget(self.calculate_direction_dropdown)

        layout.addWidget(operation_label)
        layout.addWidget(self.calculate_operation_dropdown)

        layout.addWidget(alignment_label)
        layout.addWidget(self.calculate_alignment_dropdown)

        layout.addSpacing(8)
        layout.addWidget(
            apply_button,
            alignment=Qt.AlignmentFlag.AlignLeft,
        )

        return card

    def _quick_calculate_mode_changed(self, mode):
        if self._selection_timer.isActive():
            self._cancel_cell_selection()

        range_enabled = mode == "Range"

        self.calculate_end_label.setEnabled(
            range_enabled
        )
        self.calculate_end_input.setEnabled(
            range_enabled
        )

        self._restore_select_button_text()

    def _select_quick_calculate_cells(self):
        range_mode = (
            self.calculate_mode_dropdown.currentText()
            == "Range"
        )

        self._begin_cell_selection(
            self.calculate_start_input,
            range_mode=range_mode,
        )

    def _begin_cell_selection(
        self,
        target_input,
        range_mode=False,
    ):
        if self._selection_timer.isActive():
            self._cancel_cell_selection()
            return

        context = activate_excel_selection_target(
            self,
            self.workbook_dropdown,
            self.worksheet_dropdown,
        )

        if context is None:
            return

        self._selection_context = context
        self._selection_target_input = target_input
        self._selection_range_mode = range_mode
        self._selection_start_address = None
        self._selection_last_address = context.get(
            "initial_address"
        )

        if range_mode:
            self.calculate_select_cells_button.setText(
                "Cancel Selection"
            )
        elif hasattr(
            self,
            "calculate_select_cells_button",
        ):
            self.calculate_select_cells_button.setText(
                "Cancel Selection"
            )

        window = self.window()
        window._suppress_floating_logo = True
        window.floating_logo.hide()
        window.hide()

        self._selection_timer.start()

    def _poll_excel_selection(self):
        if self._selection_context is None:
            self._cancel_cell_selection()
            return

        address = read_active_excel_cell(
            self._selection_context
        )

        if address is None:
            return

        if address == self._selection_last_address:
            return

        self._selection_last_address = address

        if not self._selection_range_mode:
            self._selection_target_input.setText(address)
            self._finish_cell_selection()
            return

        if self._selection_start_address is None:
            self._selection_start_address = address
            self.calculate_start_input.setText(address)
            return

        self.calculate_end_input.setText(address)
        self._finish_cell_selection()

    def _finish_cell_selection(self):
        self._selection_timer.stop()
        self._selection_context = None
        self._selection_target_input = None
        self._selection_range_mode = False
        self._selection_start_address = None
        self._selection_last_address = None

        self._restore_select_button_text()

        window = self.window()
        window.floating_logo.hide()
        window._suppress_floating_logo = False
        window.showNormal()
        window.raise_()
        window.activateWindow()

    def _cancel_cell_selection(self):
        self._selection_timer.stop()
        self._selection_context = None
        self._selection_target_input = None
        self._selection_range_mode = False
        self._selection_start_address = None
        self._selection_last_address = None

        self._restore_select_button_text()

        window = self.window()
        window.floating_logo.hide()
        window._suppress_floating_logo = False
        window.showNormal()
        window.raise_()
        window.activateWindow()

    def _restore_select_button_text(self):
        if not hasattr(
            self,
            "calculate_select_cells_button",
        ):
            return

        if (
            self.calculate_mode_dropdown.currentText()
            == "Range"
        ):
            self.calculate_select_cells_button.setText(
                "Select Cells"
            )
        else:
            self.calculate_select_cells_button.setText(
                "Select Cell"
            )

    def _apply_quick_calculate(self):
        apply_quick_calculate(
            self,
            self.workbook_dropdown,
            self.worksheet_dropdown,
            self.calculate_mode_dropdown.currentText(),
            self.calculate_direction_dropdown.currentText(),
            self.calculate_start_input.text(),
            self.calculate_end_input.text(),
            self.calculate_operation_dropdown.currentText(),
            self.calculate_alignment_dropdown.currentText(),
        )

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
