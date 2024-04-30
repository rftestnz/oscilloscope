"""
Test Oscilloscopes
# DK Jan 23
"""

import os
from pprint import pformat
from typing import List
from zipfile import BadZipFile

import PySimpleGUI as sg
from PyQt6 import uic
from PyQt6.QtCore import QSettings
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
)

from drivers.excel_interface import ExcelInterface
from drivers.fluke_5700a import Fluke5700A
from drivers.keysight_scope import Keysight_Oscilloscope
from drivers.Ks3458A import Ks3458A
from drivers.Ks33250A import Ks33250A
from drivers.meatest_m142 import M142
from drivers.rf_signal_generator import RF_Signal_Generator
from drivers.scpi_id import SCPI_ID
from individual_test_selector import IndividualTestSelector
from oscilloscope_tester import TestOscilloscope
from select_uut_address import AddressSelector
from utilities import get_path

VERSION = "A.01.09"


calibrator = Fluke5700A()
ks33250 = Ks33250A()
uut = Keysight_Oscilloscope()
mxg = RF_Signal_Generator()
ks3458 = Ks3458A()
simulating: bool = False

cursor_results: List = []
test_progress: sg.ProgressBar
test_number: int = 0
current_test_text: sg.Text


class UI(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

        self.m142 = M142()
        self.fl5700 = Fluke5700A()
        self.ks33250 = Ks33250A()
        self.ks3458 = Ks3458A()
        self.uut = Keysight_Oscilloscope()

        self.calibrator = self.m142

        self.do_parallel = False

        uic.loadUi(get_path("ui\\main_window.ui"), self)  # type: ignore

        self.settings = QSettings("RFTS", "Oscilloscope")
        self.statusbar.addWidget(QLabel(f"   {VERSION}   "))  # type: ignore

        self.txt_results_file = self.findChild(QLineEdit, "txtResultsFile")
        self.txt_uut_addr = self.findChild(QLineEdit, "txtUUTAddr")

        self.cmb_calibrator = self.findChild(QComboBox, "cmbCalibrator")
        self.cmb_calibrator_gpib = self.findChild(QComboBox, "cmbCalibratorGPIB")
        self.cmb_calibrator_addr = self.findChild(QComboBox, "cmbCalibratorAddr")
        self.cmb33250_gpib = self.findChild(QComboBox, "cmb33250GPIB")
        self.cmb33250_addr = self.findChild(QComboBox, "cmb33250Addr")
        self.cmb3458_gpib = self.findChild(QComboBox, "cmb3458GPIB")
        self.cmb3458_addr = self.findChild(QComboBox, "cmb3458Addr")
        self.cmb_number_channels = self.findChild(QComboBox, "cmbNumberChannels")

        self.lbl_calibrator_connection = self.findChild(
            QLabel, "lblCalibratorConnection"
        )
        self.lbl33250_connection = self.findChild(QLabel, "lbl33250Connection")
        self.lbl3458_connection = self.findChild(QLabel, "lbl3458Connection")
        self.lbl_uut_connection = self.findChild(QLabel, "lblUUTConnection")

        self.cb_skip_rows = self.findChild(QCheckBox, "cbSkipRows")
        self.cb_simulating = self.findChild(QCheckBox, "cbSimulation")

        self.btn_browse_results = self.findChild(QPushButton, "btnBrowseResults")
        self.btn_view_results = self.findChild(QPushButton, "btnViewResults")
        self.btn_select_uut_addr = self.findChild(QPushButton, "btnSelectUUTAddr")
        self.btn_test_connections = self.findChild(QPushButton, "btnTestConnections")
        self.btn_perform_tests = self.findChild(QPushButton, "btnPerformTests")
        self.btn_hide_excel_rows = self.findChild(QPushButton, "btnHideExcelRows")

        self.initialize_controls()

        self.create_connections()

    def create_connections(self) -> None:
        self.btn_browse_results.clicked.connect(self.browse_results)
        self.btn_view_results.clicked.connect(self.view_results)
        self.btn_select_uut_addr.clicked.connect(self.select_uut_addr)
        self.btn_test_connections.clicked.connect(self.test_connections)
        self.btn_perform_tests.clicked.connect(self.perform_tests)
        self.btn_hide_excel_rows.clicked.connect(self.hide_excel_rows)

    def initialize_controls(self) -> None:
        self.txt_results_file.setText(self.settings.value("filename"))

        self.cmb_number_channels.addItems(["2", "4", "6", "8"])

        self.cmb_calibrator.addItems(["5700A/5730A", "M142"])

        for gpib in range(5):
            self.cmb_calibrator_gpib.addItem(f"GPIB{gpib}")
            self.cmb33250_gpib.addItem(f"GPIB{gpib}")
            self.cmb3458_gpib.addItem(f"GPIB{gpib}")

        for addr in range(1, 31):
            self.cmb_calibrator_addr.addItem(f"{addr}")
            self.cmb33250_addr.addItem(f"{addr}")
            self.cmb3458_addr.addItem(f"{addr}")

        self.cmb_calibrator.setCurrentIndex(
            max(0, self.cmb_calibrator.findText(self.settings.value("calibrator")))
        )
        self.cmb_calibrator_gpib.setCurrentIndex(
            max(
                0,
                self.cmb_calibrator_gpib.findText(
                    self.settings.value("calibrator gpib")
                ),
            )
        )
        self.cmb_calibrator_addr.setCurrentIndex(
            max(
                0,
                self.cmb_calibrator_addr.findText(
                    self.settings.value("calibrator addr")
                ),
            )
        )
        self.cmb33250_gpib.setCurrentIndex(
            max(
                0,
                self.cmb33250_gpib.findText(self.settings.value("33250 gpib")),
            )
        )
        self.cmb33250_addr.setCurrentIndex(
            max(
                0,
                self.cmb33250_addr.findText(self.settings.value("33250 addr")),
            )
        )

        self.cmb3458_gpib.setCurrentIndex(
            max(
                0,
                self.cmb3458_gpib.findText(self.settings.value("3458 gpib")),
            )
        )
        self.cmb3458_addr.setCurrentIndex(
            max(
                0,
                self.cmb3458_addr.findText(self.settings.value("3458 addr")),
            )
        )
        self.txt_uut_addr.setText(self.settings.value("uut addr"))

    def test_connections(self) -> None:
        connected_pix = get_path("ui\\tick.png")
        unconnected_pix = get_path("ui\\cross.png")
        simulating = self.cb_simulating.isChecked()

        # Check impedance first calls the check result sheet
        impedance_tests = self.check_impedance()
        self.cmb3458_addr.setEnabled(impedance_tests)
        self.cmb3458_gpib.setEnabled(impedance_tests)
        self.lbl3458_connection.setVisible(impedance_tests)

        if self.cmb_calibrator.currentText() == "M142":
            self.calibrator = self.m142
        else:
            self.calibrator = self.fl5700

        self.calibrator.visa_address = f"{self.cmb_calibrator_gpib.currentText()}::{self.cmb_calibrator_addr.currentText()}::INSTR"
        self.calibrator.simulating = simulating
        self.calibrator.open_connection()
        self.lbl_calibrator_connection.setPixmap(
            QPixmap(connected_pix)
            if self.calibrator.is_connected()
            else QPixmap(unconnected_pix)
        )
        self.lbl_calibrator_connection.resize(QPixmap(connected_pix).size())
        QApplication.processEvents()

        self.ks33250.visa_address = f"{self.cmb33250_gpib.currentText()}::{self.cmb33250_addr.currentText()}::INSTR"
        self.ks33250.simulating = simulating
        self.ks33250.open_connection()
        self.lbl33250_connection.setPixmap(
            QPixmap(connected_pix)
            if self.ks33250.is_connected()
            else QPixmap(unconnected_pix)
        )
        self.lbl33250_connection.resize(QPixmap(connected_pix).size())
        QApplication.processEvents()

        if impedance_tests:
            self.ks3458.visa_address = f"{self.cmb3458_gpib.currentText()}::{self.cmb3458_addr.currentText()}::INSTR"
            self.ks3458.simulating = simulating
            self.ks3458.open_connection()
            self.lbl3458_connection.setPixmap(
                QPixmap(connected_pix)
                if self.ks3458.is_connected()
                else QPixmap(unconnected_pix)
            )
            self.lbl3458_connection.resize(QPixmap(connected_pix).size())
            QApplication.processEvents()

        # The uut is more complex, as we need to load the correct driver temporarily.

        check = TestOscilloscope(
            calibrator=self.calibrator,
            ks33250=self.ks33250,
            ks3458=self.ks3458,
            uut=self.uut,
            simulating=self.cb_simulating.isChecked(),
        )

        uut_connected = check.load_uut_driver(self.txt_uut_addr.text())

        if uut_connected:
            self.cmb_number_channels.setCurrentIndex(
                self.cmb_number_channels.findText(str(check.uut.num_channels))
            )

        self.uut.visa_address = self.txt_uut_addr.text()
        self.uut.simulating = simulating

        self.lbl_uut_connection.setPixmap(
            QPixmap(connected_pix) if uut_connected else QPixmap(unconnected_pix)
        )
        self.lbl_uut_connection.resize(QPixmap(connected_pix).size())
        QApplication.processEvents()

        # Save details

        self.settings.setValue("calibrator", self.cmb_calibrator.currentText())
        self.settings.setValue(
            "calibrator gpib", self.cmb_calibrator_gpib.currentText()
        )
        self.settings.setValue(
            "calibrator addr", self.cmb_calibrator_addr.currentText()
        )
        self.settings.setValue("33250 gpib", self.cmb33250_gpib.currentText())
        self.settings.setValue("33250 addr", self.cmb33250_addr.currentText())
        self.settings.setValue("3458 gpib", self.cmb3458_gpib.currentText())
        self.settings.setValue("3458 addr", self.cmb3458_addr.currentText())
        self.settings.setValue("uut addr", self.txt_uut_addr.text())

    def browse_results(self) -> None:
        if fname := QFileDialog.getOpenFileName(
            self, "Select results file", filter="Excel Files (*.xlsx)"
        ):
            self.txt_results_file.setText(fname[0])  # type: ignore
            self.settings.setValue("filename", fname[0])

    def view_results(self) -> None:
        try:
            os.startfile(f'"{self.txt_results_file.text()}"')
        except FileNotFoundError:
            QMessageBox.critical(self, "Error", "Results file not found")

    def select_uut_addr(self) -> None:

        addresses = SCPI_ID.get_all_attached()

        visa_instruments = []

        for addr in addresses:
            if addr.startswith("USB"):
                with SCPI_ID(address=addr) as scpi:
                    model = scpi.get_id()[1]
                visa_instruments.append((addr, model))

        if not len(visa_instruments):
            QMessageBox.critical(self, "Error", "No USB Visa instruments found")
        else:
            selector = AddressSelector(visa_instruments)
            selector.show()
            if selector.uut_address:
                self.txt_uut_addr.setText(selector.uut_address)

    def perform_tests(self) -> None:
        """
        perform_tests
        This just presents list of available tests to perform
        """

        # Have to read the results sheet to find what tests are performed, then present as a series of checkboxes

        with ExcelInterface(filename=self.txt_results_file.text()) as excel:
            test_names = excel.get_test_types()

            selector = IndividualTestSelector(test_names=list(test_names))
            selector.show()

            print(selector.selected_tests)

            # Now we have the list of test names, we need to get the associated test rows

            test_steps = []

            for name in selector.selected_tests:
                rows = excel.get_test_rows(name)
                test_steps.extend(iter(rows))

            self.do_parallel = False
            if "DCV" in selector.selected_tests:
                self.do_parallel = True
            if "DCV-BAL" in selector.selected_tests:
                self.do_parallel = True
            if "CURS" in selector.selected_tests:
                self.do_parallel = True

            test_rows = sorted(test_steps)

            self.perform_oscilloscope_tests(test_rows=test_rows)

    def perform_oscilloscope_tests(self, test_rows: list) -> None:
        """
        perform_tests
        Perform the actual oscilloscope tests

        Args:
            test_rows (list): _description_

        Returns:
            _type_: _description_
        """

        tester = TestOscilloscope(
            calibrator=self.calibrator,
            ks33250=self.ks33250,
            ks3458=self.ks3458,
            uut=self.uut,
            simulating=self.cb_simulating.isChecked(),
        )

        tester.run_tests(
            filename=self.txt_results_file.text(),
            test_rows=test_rows,
            uut_address=self.txt_uut_addr.text(),
            parallel_channels=self.do_parallel,
            skip_completed=self.cb_skip_rows.isChecked(),
        )

    def result_sheet_check(self) -> bool:
        """
        result_sheet_check

        Make sure the result sheet has the correct named cells

        Returns:
            bool: _description_
        """

        with ExcelInterface(filename=self.txt_results_file.text()) as excel:
            nr = excel.get_named_cell("StartCell")
            if not nr:
                QMessageBox.critical(
                    self,
                    "Error",
                    "No cell named StartCell. Name the first cell with function data StartCell",
                )
                return False

            valid_tests = excel.get_test_types()

            QMessageBox.information(
                self,
                "Tests",
                f"The following tests are found: {pformat( list(valid_tests))}",
            )

            invalid_tests = excel.get_invalid_tests()

            if len(invalid_tests):
                QMessageBox.critical(
                    self,
                    "Untested",
                    f"The following rows will not be tested: {pformat(invalid_tests)}",
                )

            return True

    def check_impedance(self) -> bool:
        """
        check_impedance
        Check the results sheet to see if there are any impedance tests.
        If not, no requirement to check 3458A

        Returns:
            bool: _description_
        """

        with ExcelInterface(filename=self.txt_results_file.text()) as excel:
            # Repeat the part of the check from result_sheet_check, but we don't want it messaging the tests to be performed
            nr = excel.get_named_cell("StartCell")
            if not nr:
                QMessageBox.critical(
                    self,
                    "Error",
                    "No cell named StartCell. Name the first cell with function data StartCell",
                )
                return False

            valid_tests = excel.get_test_types()

        return "IMP" in valid_tests

    def hide_excel_rows(self) -> None:
        with ExcelInterface(filename=self.txt_results_file.text()) as excel:
            excel.hide_excel_rows(channel=int(self.cmb_number_channels.currentText()))


if __name__ == "__main__":
    app = QApplication([])
    window = UI()
    window.show()
    app.exec()


def update_test_progress() -> None:
    """
    update_test_progress
    Test step complted, increment progress
    """

    global test_number
    global test_progress

    test_number += 1
    test_progress.update(test_number)


def template_help() -> None:
    """
    template_help _summary_
    """

    help_text = """
Test data  begins to the right of test data, outside print area.

There are 3 types of setting data, all columns must be specified, data not required in all columns though.

The presence of the data in the function column determines if this is an automated test row.

For the first row, name the cell in the function column 'StartCell' so the software knows where to start looking.

The software will look for a column including the text 'result' or 'measured' to store the test result for that table

The only implemented test names (exactly) are:

BAL - DC Balance test, DC Volts with no signal applied
DCV - DC Voltage tests
DCV-BAL - Tek style test where positive and negative voltage are applied
POS - Vertical position test (Tektronix)
CURS - Keysight test of cursor delta
RISE - Risetime using fast pulse generator
TIME - Timebase test
TRIG - Trigger sensitivity test

Columns:
There are 3 different sets of column tyopes for the tests

DCV, DCV-BAL, POS, BAL
Function, Channel, Coupling [AC/DC/GND], Scale (1x probe), Voltage (calibrator), Offset, Bandwidth, Impedance, Invert

TIME, RISE
Function, Channel (blank for time), Timebase (ns), Impedance, Bandwidth (MHz)
Impedance use 50 if available, else blank

TRIG
Function, Channel [1-4,EXT], Scale (1x probe), Voltage (RF Source), Impedance (channel input), Frequency (MHz), Edge [Rise/Fall]

Don't mix tables with different types of tests. The above column headers are not read, just assumed
    """

    layout = [
        [
            sg.Text(
                "Ensure the following are followed in creating/modifying template",
                background_color="blue",
            )
        ],
        [sg.Multiline(default_text=help_text, size=(60, 20), disabled=True)],
        [sg.Ok(size=(12, 1))],
    ]

    window = sg.Window(
        "Template setup",
        layout=layout,
        icon=get_path("ui\\scope.ico"),
        background_color="blue",
    )

    window.read()

    window.close()


def hide_excel_rows(filename: str, channel: int) -> None:
    """
    hide_excel_rows
    Hide the template rows for sheets which have different number of channels
    Channl filter must be in column A

    Args:
        channel (int): _description_
    """

    with ExcelInterface(filename=filename) as excel:
        excel.hide_excel_rows(channel=channel)


if __name__ == "_main_":
    sg.theme("black")

    gpib_ifc_list = [f"GPIB{x}" for x in range(5)]
    gpib_addresses = list(range(1, 32))

    layout = [
        [sg.Text("Oscilloscope Test")],
        [sg.Text(f"DK Apr 23 VERSION {VERSION}")],
        [sg.Text()],
        [
            sg.Text("Create Excel sheet from template first", text_color="red"),
            sg.Text(" " * 70),
            sg.Button("Template Help", key="-TEMPLATE_HELP-"),
            sg.Button(
                "Check",
                size=(12, 1),
                key="-RESULTS_CHECK-",
                tooltip="Check results sheet for conformance",
            ),
        ],
        [
            sg.Text("Results file", size=(10, 1)),
            sg.Input(
                default_text=sg.user_settings_get_entry("-FILENAME-"),  # type: ignore
                size=(60, 1),
                key="-FILE-",
            ),
            sg.FileBrowse(
                "...",
                target="-FILE-",
                file_types=(("Excel Files", "*.xlsx"), ("All Files", "*.*")),
                initial_folder=sg.user_settings_get_entry("-RESULTS_FOLDER-"),
            ),
            sg.Button("View", size=(15, 1), key="-VIEW-", disabled=False),
        ],
        [sg.Text()],
        [
            sg.Check(
                "Simulate",
                default=sg.user_settings_get_entry("-SIMULATE-"),  # type: ignore
                key="-SIMULATE-",
                disabled=False,
            )
        ],
        [sg.Text("Instrument Config:")],
        [
            sg.Text("Main Calibrator", size=(15, 1)),
            sg.Combo(
                ["5700A", "5730A", "M-142"],
                size=(10, 1),
                default_value=sg.user_settings_get_entry("-CALIBRATOR-"),
                key="-CALIBRATOR-",
                enable_events=True,
            ),
        ],
        [
            sg.Text("Calibrator", size=(15, 1)),
            sg.Combo(
                gpib_ifc_list,
                size=(10, 1),
                key="GPIB_FLUKE_5700A",
                default_value=sg.user_settings_get_entry("-FLUKE_5700A_GPIB_IFC-"),
            ),
            sg.Combo(
                gpib_addresses,
                default_value=(
                    "4"
                    if sg.user_settings_get_entry("-CALIBRATOR-") == "M-142"
                    else "6"
                ),
                size=(6, 1),
                key="GPIB_ADDR_FLUKE_5700A",
            ),
        ],
        [
            sg.Text("33250A", size=(15, 1)),
            sg.Combo(
                gpib_ifc_list,
                size=(10, 1),
                key="GPIB_IFC_33250",
                default_value=sg.user_settings_get_entry("-33250_GPIB_IFC-"),
            ),
            sg.Combo(
                gpib_addresses,
                default_value=sg.user_settings_get_entry("-33250_GPIB_ADDR-"),
                size=(6, 1),
                key="GPIB_ADDR_33250",
            ),
        ],
        [
            sg.Text("3458A", size=(15, 1)),
            sg.Combo(
                gpib_ifc_list,
                size=(10, 1),
                key="GPIB_IFC_3458",
                default_value=sg.user_settings_get_entry("-3458_GPIB_IFC-"),
            ),
            sg.Combo(
                gpib_addresses,
                default_value=sg.user_settings_get_entry("-3458_GPIB_ADDR-"),
                size=(6, 1),
                key="GPIB_ADDR_3458",
            ),
            sg.Text("(Only for impedance measurements)"),
        ],
        [
            sg.Text("UUT", size=(15, 1)),
            sg.Input(
                default_text=sg.user_settings_get_entry("-UUT_ADDRESS-"),  # type: ignore
                size=(60, 1),
                key="-UUT_ADDRESS-",
            ),
            sg.Button("Select", size=(12, 1), key="-SELECT_ADDRESS-"),
        ],
        [
            sg.Text("Number Channels", size=(15, 1)),
            sg.Combo([2, 4, 6, 8], size=(10, 1), key="-UUT_CHANNELS-", default_value=4),
        ],
        [sg.Text()],
        [
            sg.ProgressBar(
                max_value=100, size=(35, 10), visible=False, key="-PROGRESS-"
            ),
            sg.Text("Progress", visible=False, key="-PROG_TEXT-"),
        ],
        [sg.Text(key="-CURRENT_TEST-")],
        [sg.Text()],
        [sg.Check("Skip already tested", default=True, key="-SKIP_TESTED-")],
        [
            sg.Button("Test Connections", size=(15, 1), key="-TEST_CONNECTIONS-"),
            sg.Button("Perform Tests", size=(12, 1), key="-INDIVIDUAL-"),
            sg.Button(
                "Hide Excel Rows",
                size=(12, 1),
                key="-HIDE_EXCEL_ROWS-",
                tooltip="For templates with more channels than present in UUT",
            ),
            sg.Exit(size=(12, 1)),
        ],
    ]

    window = sg.Window(
        "Oscilloscope Test",
        layout=layout,
        finalize=True,
        icon=get_path("ui\\scope.ico"),
        titlebar_icon=get_path("ui\\scope.ico"),
    )

    back_color = window["-FILE-"].BackgroundColor
    test_progress = window["-PROGRESS-"]  # type:ignore
    current_test_text = window["-CURRENT_TEST-"]  # type:ignore

    simulating = False

    while True:
        event, values = window.read(200)  # type: ignore

        if event in ["Exit", sg.WIN_CLOSED]:
            break

        # Get the default text color in case changing simulate
        txt_clr = sg.theme_element_text_color()

        # read the simulate setting
        if values["-SIMULATE-"]:
            window["-SIMULATE-"].update(text_color="red")
        else:
            window["-SIMULATE-"].update(text_color=txt_clr)

        sg.user_settings_set_entry("-SIMULATE-", values["-SIMULATE-"])
        sg.user_settings_set_entry("-CALIBRATOR-", values["-CALIBRATOR-"])
        sg.user_settings_set_entry("-FLUKE_5700A_GPIB_IFC-", values["GPIB_FLUKE_5700A"])
        sg.user_settings_set_entry("-33250_GPIB_IFC-", values["GPIB_IFC_33250"])
        sg.user_settings_set_entry("-33250_GPIB_ADDR-", values["GPIB_ADDR_33250"])
        sg.user_settings_set_entry("-3458_GPIB_IFC-", values["GPIB_IFC_3458"])
        sg.user_settings_set_entry("-3458_GPIB_ADDR-", values["GPIB_ADDR_3458"])
        sg.user_settings_set_entry("-UUT_ADDRESS-", values["-UUT_ADDRESS-"])
        sg.user_settings_set_entry("-FILENAME-", values["-FILE-"])

        sg.user_settings_save()

        simulating = values["-SIMULATE-"]

        calibrator.simulating = simulating
        calibrator_address = (
            f"{values['GPIB_FLUKE_5700A']}::{values['GPIB_ADDR_FLUKE_5700A']}::INSTR"
        )

        ks33250.simulating = simulating
        ks33250_address = (
            f"{values['GPIB_IFC_33250']}::{values['GPIB_ADDR_33250']}::INSTR"
        )

        uut.simulating = simulating
        uut.visa_address = values["-UUT_ADDRESS-"]

        ks3458.simulating = simulating
        ks3458_address = f"{values['GPIB_IFC_3458']}::{values['GPIB_ADDR_3458']}::INSTR"

        if event in ["-INDIVIDUAL-", "-TEST_CONNECTIONS-"]:
            # make sure the file is valid, we have to read it to check impedance tests

            # Common check to make sure everything is in order

            valid = True

            # reset the back colors instead of else statements after checking

            window["-FILE-"].update(background_color=back_color)

            if not values["-FILE-"]:
                window["-FILE-"].update(background_color="Red")
                valid = False

            if not os.path.isfile(values["-FILE-"]):
                sg.popup_error(
                    "Filenname does not exist",
                    background_color="blue",
                    icon=get_path("ui\\scope.ico"),
                )
                window["-FILE-"].update(
                    background_color="Red",
                )
                continue

            try:
                with ExcelInterface(filename=values["-FILE-"]) as excel:
                    if not excel.check_valid_results():
                        sg.popup_error(
                            "No Named range for StartCell in results",
                            background_color="blue",
                            icon=get_path("ui\\scope.ico"),
                        )
                        valid = False
            except BadZipFile:
                sg.popup_error(
                    "Result sheet appears corrupted. Check backups folder for most recent, or generate new from template",
                    background_color="blue",
                    icon=get_path("ui\\scope.ico"),
                )
                valid = False

            if not valid:
                continue

        if event in [
            "-INDIVIDUAL-",
        ]:
            window["-VIEW-"].update(disabled=True)  # Disable while testing

            if values["-CALIBRATOR-"] == "M-142":
                calibrator = M142(simulate=simulating)
            else:
                calibrator = Fluke5700A(simulate=simulating)
            calibrator.visa_address = calibrator_address
            ks33250.visa_address = ks33250_address
            ks3458.visa_address = ks3458_address

            if load_uut_driver(address=values["-UUT_ADDRESS-"], simulating=simulating):
                uut.visa_address = values["-UUT_ADDRESS-"]
                uut.open_connection()
                window["-UUT_CHANNELS-"].update(value=f"{uut.num_channels}")

                test_rows, do_parallel = individual_tests(filename=values["-FILE-"])
                if len(test_rows):
                    test_progress.update(0, max=len(test_rows))
                    test_progress.update(visible=True)
                    window["-PROG_TEXT-"].update(visible=True)
                    parallel = "NO"

                    if do_parallel:
                        parallel = sg.popup_yes_no(
                            "Will you connect all channels in parallel for DCV tests?",
                            title="Parallel Channels",
                            background_color="blue",
                            icon=get_path("ui\\scope.ico"),
                        )
                    run_tests(
                        filename=values["-FILE-"],
                        test_rows=test_rows,
                        parallel_channels=parallel == "Yes",
                        skip_completed=values["-SKIP_TESTED-"],
                    )

                    test_progress.update(visible=False)
                    window["-PROG_TEXT-"].update(visible=False)
                    current_test_text.update("")
                    sg.popup(
                        "Finished",
                        background_color="blue",
                        icon=get_path("ui\\scope.ico"),
                    )

            window["-VIEW-"].update(disabled=False)

        if event == "-VIEW-":
            os.startfile(f'"{values["-FILE-"]}"')

        if event == "-SELECT_ADDRESS-":
            address = select_visa_address()
            if address:
                window["-UUT_ADDRESS-"].update(address)  # type: ignore

        if event == "-TEMPLATE_HELP-":
            template_help()

        if event == "-RESULTS_CHECK-":
            results_sheet_check(filename=values["-FILE-"])

        if event == "-CALIBRATOR-":
            window["GPIB_ADDR_FLUKE_5700A"].update(
                value="4" if values["-CALIBRATOR-"] == "M-142" else "6"
            )

        if event == "-HIDE_EXCEL_ROWS-":
            hide_excel_rows(filename=values["-FILE-"], channel=values["-UUT_CHANNELS-"])
