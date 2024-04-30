"""
 Dialog to show all of the USB addresses of attached devices
"""

from PyQt6 import uic
from PyQt6.QtCore import QObject, QSettings, Qt
from PyQt6.QtGui import QAction, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QRadioButton,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenuBar,
    QMessageBox,
    QPushButton,
    QStatusBar,
    QDialog,
    QVBoxLayout,
    QDialogButtonBox,
    QWidget,
)


class AddressSelector(QDialog):

    def __init__(self, address_list: list[tuple]) -> None:
        super().__init__()

        self.radio_buttons: list[QRadioButton] = []
        self.uut_address: str = ""

        self.selector = QDialog()

        self.selector.setWindowTitle("Select UUT address")

        self.layout1 = QVBoxLayout()

        self.layout1.addWidget(QLabel("Select UUT Address"))

        selected = False  # highlight the first

        for addr in address_list:
            rb = QRadioButton(f"{addr[0]} ({addr[1]})")
            if not selected:
                rb.setChecked(True)
                selected = True
            self.radio_buttons.append(rb)
            self.layout1.addWidget(rb)

        buttons = (
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttonBox = QDialogButtonBox(buttons)

        self.buttonBox.clicked.connect(self.button_pressed)

        self.layout1.addWidget(self.buttonBox)

        self.selector.setLayout(self.layout1)

    def show(self) -> None:

        self.selector.exec()

    def button_pressed(self, button: QPushButton) -> None:

        self.uut_address = []

        if button.text() == "OK":
            # Get the selected radio button
            for rb in self.radio_buttons:
                if rb.isChecked():
                    # We added the model number, strip it off
                    addr = rb.text().split("(")
                    self.uut_address = addr[0].strip()
                    break

        self.selector.close()
