APP_STYLE = """
QWidget {
    background-color: #1D1C26;
    color: #FFFFFF;
    font-family: "Segoe UI";
    font-size: 14px;
}

QMainWindow {
    background-color: #18181F;
}

QFrame#sidebar {
    background-color: #1D1C26;
    border-right: 1px solid #242330;
}

QLabel#appTitle {
    color: #FFFFFF;
    font-size: 24px;
    font-weight: 700;
    padding: 20px 16px;
}

QPushButton#navButton {
    background-color: transparent;
    color: #969696;
    border: none;
    text-align: left;
    padding: 13px 18px;
    border-radius: 8px;
}

QPushButton#navButton:hover {
    background-color: #242330;
    color: #FFFFFF;
}

QPushButton#navButton:checked {
    background-color: #6C1ED2;
    color: #FFFFFF;
}

QWidget#contentPage {
    background-color: #18181F;
}

QLabel#pageTitle {
    color: #FFFFFF;
    font-size: 28px;
    font-weight: 700;
}

QLabel#description,
QLabel#cardText {
    color: #969696;
    font-size: 14px;
}

QFrame#card {
    background-color: #1D1C26;
    border: 1px solid #242330;
    border-radius: 14px;
}

QLabel#cardTitle {
    color: #FFFFFF;
    font-size: 18px;
    font-weight: 600;
}

QPushButton#secondaryButton {
    background-color: #242330;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 10px 16px;
    font-weight: 600;
}

QPushButton#secondaryButton:hover {
    background-color: #302F3E;
}

QPushButton#actionButton {
    background-color: #6C1ED2;
    color: #FFFFFF;
    border: none;
    border-radius: 9px;
    padding: 12px 20px;
    font-weight: 600;
}

QPushButton#actionButton:hover {
    background-color: #7C2CE3;
}

QPushButton#actionButton:pressed {
    background-color: #5917B4;
}

QComboBox {
    background-color: #18181F;
    color: #FFFFFF;
    border: 1px solid #242330;
    border-radius: 8px;
    padding: 10px 12px;
}

QComboBox:hover,
QComboBox:focus {
    border: 1px solid #6C1ED2;
}

QComboBox QAbstractItemView {
    background-color: #18181F;
    color: #FFFFFF;
    border: 1px solid #242330;
    selection-background-color: #6C1ED2;
}

QMessageBox {
    background-color: #1D1C26;
}

QMessageBox QLabel {
    background-color: #1D1C26;
    color: #FFFFFF;
}

QMessageBox QPushButton {
    min-width: 85px;
    background-color: #6C1ED2;
    color: #FFFFFF;
    border: none;
    border-radius: 7px;
    padding: 8px 14px;
}

QDoubleSpinBox {
    background-color: #18181F;
    color: #FFFFFF;
    border: 1px solid #242330;
    border-radius: 8px;
    padding: 8px 10px;
    min-width: 105px;
}

QDoubleSpinBox:hover,
QDoubleSpinBox:focus {
    border: 1px solid #6C1ED2;
}

QRadioButton {
    color: #FFFFFF;
    spacing: 8px;
}

QRadioButton::indicator {
    width: 16px;
    height: 16px;
}

QRadioButton::indicator:unchecked {
    border: 2px solid #969696;
    border-radius: 9px;
    background-color: #242330;
}

QRadioButton::indicator:checked {
    border: 4px solid #6C1ED2;
    border-radius: 9px;
    background-color: #E1C8FF;
}

"""
