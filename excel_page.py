from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget


class ExcelPage(QWidget):
    def __init__(self):
        super().__init__()

        self.setObjectName("contentPage")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(42, 36, 42, 36)
        layout.setSpacing(18)

        title = QLabel("Excel")
        title.setObjectName("pageTitle")

        description = QLabel("Coming Soon...")
        description.setObjectName("description")

        layout.addWidget(title)
        layout.addWidget(description)
        layout.addStretch()
