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
)


class IndividualTestSelector(QDialog, object):

    def __init__(self, test_names: list):
        super.__init__()
