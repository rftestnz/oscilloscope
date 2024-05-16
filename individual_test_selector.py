"""
    Individual test selector dialog
"""

from PyQt6.QtWidgets import (
    QCheckBox,
    QLabel,
    QPushButton,
    QDialog,
    QVBoxLayout,
    QDialogButtonBox,
)


class IndividualTestSelector(QDialog):

    def __init__(self, test_names: list):
        super().__init__()

        self.selected_tests = []

        self.checkboxes: list[QCheckBox] = []

        self.selector = QDialog()
        self.selector.setWindowTitle("Select tests to perform")

        self.layout1 = QVBoxLayout()

        self.layout1.addWidget(QLabel("Select tests to perform"))

        for name in test_names:
            cb = QCheckBox(name)
            cb.setChecked(True)
            self.checkboxes.append(cb)
            self.layout1.addWidget(cb)

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

        self.selected_tests = []

        if button.text() == "OK":
            # Get the selected checkboxes
            for cb in self.checkboxes:
                if cb.isChecked():
                    self.selected_tests.append(cb.text())

        self.selector.close()
