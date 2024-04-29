"""
    Individual test selector dialog
"""

from PyQt6 import uic
from PyQt6.QtCore import QObject, QSettings
from PyQt6.QtGui import QAction, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
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
)


class IndividualTestSelector(QDialog):

    def __init__(self, test_names: list):
        super().__init__()

        self.selector = QDialog()
        self.selector.setWindowTitle("Select tests to perform")

        self.layout1 = QVBoxLayout()

        self.layout1.addWidget(QLabel("Select tests to perform"))

        for name in test_names:
            cb = QCheckBox(name)
            self.layout1.addWidget(cb)

        buttons = (
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttonBox = QDialogButtonBox(buttons)

        self.layout1.addWidget(self.buttonBox)

        self.selector.setLayout(self.layout1)

    def show(self) -> None:

        self.selector.exec()
