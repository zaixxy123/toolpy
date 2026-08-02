from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from clipboard_capture import ImageCaptureManager
from replacement_queue import ReplacementQueueManager
from global_hotkeys import GlobalReplacementHotkeys
from excel_page import ExcelPage
from utils import resource_path
from word_tools import (
    refresh_open_documents,
    resize_images,
    set_all_images_behind_text,
)


class ToolPyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowIcon(
            QIcon(resource_path("assets/logo.ico"))
        )
        self.setWindowTitle("ToolPy")
        self.resize(920, 700)
        self.setMinimumSize(760, 540)

        self.document_dropdown = QComboBox()

        self.capture_manager = ImageCaptureManager(self)

        self.replacement_queue_manager = ReplacementQueueManager(self)
        self.replacement_queue_manager.state_changed.connect(
            self._update_replacement_queue_state
        )
        self.replacement_queue_manager.count_changed.connect(
            self._update_replacement_queue_count
        )
        self.replacement_queue_manager.current_changed.connect(
            self._update_replacement_queue_current
        )
        self.capture_manager.count_changed.connect(
            self._update_capture_count
        )
        self.capture_manager.state_changed.connect(
            self._update_capture_state
        )

        self.global_replacement_hotkeys = GlobalReplacementHotkeys(self)
        self.global_replacement_hotkeys.finish_requested.connect(
            self._finish_replacement_queue
        )
        self.global_replacement_hotkeys.replace_requested.connect(
            self._replace_from_hotkey
        )
        self.global_replacement_hotkeys.paste_requested.connect(
            self._paste_from_hotkey
        )
        self.global_replacement_hotkeys.escape_requested.connect(
            self._escape_from_hotkey
        )
        self.global_replacement_hotkeys.activation_failed.connect(
            self._global_hotkey_failed
        )

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

        self.page_stack = QStackedWidget()
        self.page_stack.addWidget(self._create_word_page())
        self.page_stack.addWidget(ExcelPage())

        main_layout.addWidget(self._create_sidebar())
        main_layout.addWidget(self.page_stack, 1)

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

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)

        self.word_button = self._nav_button("Word", checked=True)
        self.excel_button = self._nav_button("Excel")

        self.nav_group.addButton(self.word_button)
        self.nav_group.addButton(self.excel_button)

        self.word_button.clicked.connect(
            lambda: self.page_stack.setCurrentIndex(0)
        )
        self.excel_button.clicked.connect(
            lambda: self.page_stack.setCurrentIndex(1)
        )

        layout.addWidget(app_title)
        layout.addWidget(self.word_button)
        layout.addWidget(self.excel_button)
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
        layout.addWidget(self._create_image_capture_card())
        layout.addWidget(self._create_replacement_queue_card())
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

    def _create_image_capture_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(230)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Image capture")
        title.setObjectName("cardTitle")

        text = QLabel(
            "Start recording, copy images from different apps or websites, "
            "then paste every captured image into the selected Word document."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        self.capture_state_label = QLabel("Status: Idle")
        self.capture_state_label.setObjectName("cardText")

        self.capture_count_label = QLabel("Captured: 0 images")
        self.capture_count_label.setObjectName("cardText")

        button_row = QHBoxLayout()
        button_row.setSpacing(10)

        self.start_capture_button = QPushButton("Start Capture")
        self.start_capture_button.setObjectName("actionButton")
        self.start_capture_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        self.start_capture_button.setFixedSize(145, 44)
        self.start_capture_button.clicked.connect(
            self._start_image_capture
        )

        self.paste_capture_button = QPushButton("Paste Images")
        self.paste_capture_button.setObjectName("actionButton")
        self.paste_capture_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        self.paste_capture_button.setFixedSize(145, 44)
        self.paste_capture_button.setEnabled(False)
        self.paste_capture_button.clicked.connect(
            self._paste_captured_images
        )

        self.cancel_capture_button = QPushButton("Cancel")
        self.cancel_capture_button.setObjectName("secondaryButton")
        self.cancel_capture_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        self.cancel_capture_button.setFixedSize(105, 44)
        self.cancel_capture_button.setEnabled(False)
        self.cancel_capture_button.clicked.connect(
            self._cancel_image_capture
        )

        button_row.addWidget(self.start_capture_button)
        button_row.addWidget(self.paste_capture_button)
        button_row.addWidget(self.cancel_capture_button)
        button_row.addStretch()

        layout.addWidget(title)
        layout.addWidget(text)
        layout.addWidget(self.capture_state_label)
        layout.addWidget(self.capture_count_label)
        layout.addSpacing(6)
        layout.addLayout(button_row)

        return card

    def _create_replacement_queue_card(self):
        card = QFrame()
        card.setObjectName("card")
        card.setMaximumWidth(700)
        card.setMinimumHeight(255)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(12)

        title = QLabel("Replacement queue")
        title.setObjectName("cardTitle")

        text = QLabel(
            "Capture several images in order. Finish capture, then select "
            "one Word image at a time and replace it with the next image."
        )
        text.setObjectName("cardText")
        text.setWordWrap(True)

        self.replacement_queue_state_label = QLabel("Status: Idle")
        self.replacement_queue_state_label.setObjectName("cardText")
        self.replacement_queue_count_label = QLabel("Captured: 0 images")
        self.replacement_queue_count_label.setObjectName("cardText")
        self.replacement_queue_current_label = QLabel("Current: —")
        self.replacement_queue_current_label.setObjectName("cardText")

        self.replacement_hotkey_label = QLabel(
            "Replacement mode: Off"
        )
        self.replacement_hotkey_label.setObjectName("cardText")

        button_row = QHBoxLayout()
        button_row.setSpacing(10)

        self.start_replacement_queue_button = QPushButton("Start Capture")
        self.start_replacement_queue_button.setObjectName("actionButton")
        self.start_replacement_queue_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.start_replacement_queue_button.setFixedSize(145, 44)
        self.start_replacement_queue_button.clicked.connect(self._start_replacement_queue)

        self.finish_replacement_queue_button = QPushButton("Finish Capture")
        self.finish_replacement_queue_button.setObjectName("secondaryButton")
        self.finish_replacement_queue_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.finish_replacement_queue_button.setFixedSize(145, 44)
        self.finish_replacement_queue_button.setEnabled(False)
        self.finish_replacement_queue_button.clicked.connect(self._finish_replacement_queue)

        self.clear_replacement_queue_button = QPushButton("Clear Queue")
        self.clear_replacement_queue_button.setObjectName("secondaryButton")
        self.clear_replacement_queue_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_replacement_queue_button.setFixedSize(120, 44)
        self.clear_replacement_queue_button.setEnabled(False)
        self.clear_replacement_queue_button.clicked.connect(self._clear_replacement_queue)

        button_row.addWidget(self.start_replacement_queue_button)
        button_row.addWidget(self.finish_replacement_queue_button)
        button_row.addWidget(self.clear_replacement_queue_button)
        button_row.addStretch()

        layout.addWidget(title)
        layout.addWidget(text)
        layout.addWidget(self.replacement_queue_state_label)
        layout.addWidget(self.replacement_queue_count_label)
        layout.addWidget(self.replacement_queue_current_label)
        layout.addWidget(self.replacement_hotkey_label)
        layout.addSpacing(6)
        layout.addLayout(button_row)
        return card

    def _start_replacement_queue(self):
        self.replacement_queue_manager.start_capture()
        self.start_replacement_queue_button.setEnabled(False)
        self.finish_replacement_queue_button.setEnabled(True)
        self.clear_replacement_queue_button.setEnabled(True)

        if self.global_replacement_hotkeys.activate_capture():
            self.replacement_hotkey_label.setText(
                "Capture mode: On — F finishes, Esc cancels"
            )
        else:
            self.replacement_hotkey_label.setText(
                "Capture mode: Hotkeys unavailable"
            )

    def _finish_replacement_queue(self):
        if not self.replacement_queue_manager.is_recording:
            return

        if self.replacement_queue_manager.finish_capture(self):
            self.global_replacement_hotkeys.deactivate()
            self.finish_replacement_queue_button.setEnabled(False)
            self.clear_replacement_queue_button.setEnabled(True)

            if self.global_replacement_hotkeys.activate_replacement():
                self.replacement_hotkey_label.setText(
                    "Replacement mode: On — R replaces, P pastes, Esc exits"
                )
            else:
                self.replacement_hotkey_label.setText(
                    "Replacement mode: Hotkeys unavailable"
                )

    def _replace_from_queue(self):
        if self.replacement_queue_manager.replace_selected(self):
            self._check_queue_finished()

    def _paste_from_queue(self):
        if self.replacement_queue_manager.paste_next(self):
            self._check_queue_finished()

    def _check_queue_finished(self):
        if not self.replacement_queue_manager.has_remaining:
            self.global_replacement_hotkeys.deactivate()
            self.replacement_hotkey_label.setText(
                "Replacement mode: Off — queue complete"
            )

    def _clear_replacement_queue(self):
        if not self._confirm_queue_clear():
            return

        self._force_clear_replacement_queue()

    def _force_clear_replacement_queue(self):
        self.global_replacement_hotkeys.deactivate()
        self.replacement_queue_manager.clear()
        self.replacement_hotkey_label.setText(
            "Replacement mode: Off"
        )
        self._reset_replacement_queue_buttons()

    def _confirm_queue_clear(self):
        if self.replacement_queue_manager.is_recording:
            captured = self.replacement_queue_manager.count

            if captured == 0:
                return True

            answer = QMessageBox.question(
                self,
                "Cancel Capture?",
                f"Discard {captured} captured image(s)?",
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            return answer == QMessageBox.StandardButton.Yes

        remaining = self.replacement_queue_manager.remaining

        if remaining == 0:
            return True

        answer = QMessageBox.question(
            self,
            "Exit Replacement Mode?",
            f"Discard {remaining} remaining queued image(s)?",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return answer == QMessageBox.StandardButton.Yes

    def _replace_from_hotkey(self):
        if self.global_replacement_hotkeys.mode == "replacement":
            self._replace_from_queue()

    def _paste_from_hotkey(self):
        if self.global_replacement_hotkeys.mode == "replacement":
            self._paste_from_queue()

    def _escape_from_hotkey(self):
        if self.global_replacement_hotkeys.mode in (
            "capture",
            "replacement",
        ):
            self._clear_replacement_queue()

    def _global_hotkey_failed(self, message):
        QMessageBox.warning(
            self,
            "Hotkey Unavailable",
            message,
        )

    def _update_replacement_queue_state(self, state):
        self.replacement_queue_state_label.setText(f"Status: {state}")
        if state == "Idle":
            self._reset_replacement_queue_buttons()

    def _update_replacement_queue_count(self, count):
        word = "image" if count == 1 else "images"
        self.replacement_queue_count_label.setText(f"Captured: {count} {word}")

    def _update_replacement_queue_current(self, current, total):
        if total == 0:
            self.replacement_queue_current_label.setText("Current: —")
        else:
            self.replacement_queue_current_label.setText(f"Current: {current} / {total}")

    def _reset_replacement_queue_buttons(self):
        self.start_replacement_queue_button.setEnabled(True)
        self.finish_replacement_queue_button.setEnabled(False)
        self.clear_replacement_queue_button.setEnabled(False)
        self.replacement_hotkey_label.setText(
            "Replacement mode: Off"
        )

    def _start_image_capture(self):
        self.capture_manager.start()

        self.start_capture_button.setEnabled(False)
        self.start_capture_button.setText("Recording...")
        self.paste_capture_button.setEnabled(False)
        self.cancel_capture_button.setEnabled(True)

    def _paste_captured_images(self):
        self.capture_manager.paste_into_word(
            self.document_dropdown,
            self,
        )

    def _cancel_image_capture(self):
        self.capture_manager.cancel()
        self._reset_capture_buttons()

    def _update_capture_count(self, count):
        word = "image" if count == 1 else "images"
        self.capture_count_label.setText(
            f"Captured: {count} {word}"
        )
        self.paste_capture_button.setEnabled(
            self.capture_manager.is_recording and count > 0
        )

    def _update_capture_state(self, state):
        self.capture_state_label.setText(f"Status: {state}")

        if state == "Idle":
            self._reset_capture_buttons()

    def _reset_capture_buttons(self):
        self.start_capture_button.setEnabled(True)
        self.start_capture_button.setText("Start Capture")
        self.paste_capture_button.setEnabled(False)
        self.cancel_capture_button.setEnabled(False)

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

    def closeEvent(self, event):
        self.global_replacement_hotkeys.deactivate()
        self.replacement_queue_manager.clear()
        self.capture_manager.cancel()
        super().closeEvent(event)
