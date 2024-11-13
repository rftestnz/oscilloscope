"""
Test Oscilloscopes
# DK Jan 23
"""

import os
from pprint import pformat
from zipfile import BadZipFile

from PyQt6 import uic
from PyQt6.QtCore import QSettings
from PyQt6.QtGui import QPixmap, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStatusBar,
    QGroupBox,
)

from pathlib import Path
from drivers.excel_interface import ExcelInterface
from drivers.fluke_5700a import Fluke5700A
from drivers.keysight_scope import Keysight_Oscilloscope
from drivers.Ks3458A import Ks3458A
from drivers.Ks33250A import Ks33250A
from drivers.meatest_m142 import M142
from drivers.scpi_id import SCPI_ID
from individual_test_selector import IndividualTestSelector
from oscilloscope_tester import TestOscilloscope
from select_uut_address import AddressSelector
from utilities import get_path

VERSION = "A.02.04"


class UI(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

        # Create all as simulated, else it will scan the instruments
        # instruments may not be present or wrong address, which has the effect of a very long startup
        self.m142 = M142(simulate=True)
        self.fl5700 = Fluke5700A(simulate=True)
        self.ks33250 = Ks33250A(simulate=True)
        self.ks3458 = Ks3458A(simulate=True)
        self.uut = Keysight_Oscilloscope(simulate=True)

        self.calibrator = self.m142

        self.tester = TestOscilloscope(
            calibrator=self.calibrator,
            ks33250=self.ks33250,
            ks3458=self.ks3458,
            uut=self.uut,
            simulating=True,
        )

        self.do_parallel = False

        uic.loadUi(get_path("ui\\main_window.ui"), self)  # type: ignore
        self.setWindowIcon(QIcon(get_path("ui\\scope.ico")))

        self.settings = QSettings("RFTS", "Oscilloscope")

        self.statusbar = self.findChild(QStatusBar, "statusbar")

        self.statusbar.addPermanentWidget(
            QLabel(f"  Version: {VERSION}   ")
        )  # shows on the right

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
        self.cb_filter_low_ranges = self.findChild(QCheckBox, "cbFilterLowRanges")

        self.btn_browse_results = self.findChild(QPushButton, "btnBrowseResults")
        self.btn_view_results = self.findChild(QPushButton, "btnViewResults")
        self.btn_select_uut_addr = self.findChild(QPushButton, "btnSelectUUTAddr")
        self.btn_test_connections = self.findChild(QPushButton, "btnTestConnections")
        self.btn_perform_tests = self.findChild(QPushButton, "btnPerformTests")
        self.btn_hide_excel_rows = self.findChild(QPushButton, "btnHideExcelRows")
        self.btn_abort = self.findChild(QPushButton, "btnAbort")

        self.progress_test = self.findChild(QProgressBar, "progressTest")

        self.group_hardware = self.findChild(QGroupBox, "groupHardware")

        self.initialize_controls()

        self.create_connections()

    def create_connections(self) -> None:
        self.btn_browse_results.clicked.connect(self.browse_results)
        self.btn_view_results.clicked.connect(self.view_results)
        self.btn_select_uut_addr.clicked.connect(self.select_uut_addr)
        self.btn_test_connections.clicked.connect(self.test_connections)
        self.btn_perform_tests.clicked.connect(self.perform_tests)
        self.btn_hide_excel_rows.clicked.connect(self.hide_excel_rows)
        self.txt_results_file.textChanged.connect(self.check_excel_button)
        self.btn_abort.clicked.connect(self.abort_test)

    def initialize_controls(self) -> None:

        self.progress_test.setVisible(False)
        self.btn_abort.setVisible(False)

        self.cb_filter_low_ranges.setChecked(False)
        self.cb_skip_rows.setChecked(False)
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

        self.check_excel_button()

    def test_connections(self) -> bool:
        """
        test_connections
        Check instruments connected

        Returns:
            bool: state of uut connection, the only one that is critical
        """
        connected_pix = get_path("ui\\tick.png")
        unconnected_pix = get_path("ui\\cross.png")
        simulating = self.cb_simulating.isChecked()

        # Check impedance first calls the check result sheet
        impedance_tests = self.check_test_required("IMP")
        self.cmb3458_addr.setEnabled(impedance_tests)
        self.cmb3458_gpib.setEnabled(impedance_tests)
        self.lbl3458_connection.setVisible(impedance_tests)

        # The 32250A is used for most scopes, but not the Tek MSO4, MSO5, and MSO6 series
        time_tests = self.check_test_required("TIME")
        self.cmb33250_addr.setEnabled(time_tests)
        self.cmb33250_gpib.setEnabled(time_tests)
        self.lbl33250_connection.setVisible(time_tests)

        if self.cmb_calibrator.currentText() == "M142":
            self.calibrator = self.m142
        else:
            self.calibrator = self.fl5700

        self.calibrator.close()
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

        if time_tests:
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

        uut_connected = check.load_uut_driver(
            self.txt_uut_addr.text(), simulating=simulating
        )

        if uut_connected[0] and not simulating:
            self.cmb_number_channels.setCurrentIndex(
                self.cmb_number_channels.findText(str(check.uut.num_channels))
            )

        self.uut = uut_connected[1]
        if self.uut:
            self.uut.simulating = simulating

        self.lbl_uut_connection.setPixmap(
            QPixmap(connected_pix) if uut_connected[0] else QPixmap(unconnected_pix)
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

        return uut_connected[0]

    def browse_results(self) -> None:
        dir = "."
        if len(self.txt_results_file.text()):
            p = Path(self.txt_results_file.text())
            dir = p.parent.__str__()
        if fname := QFileDialog.getOpenFileName(
            self, "Select results file", filter="Excel Files (*.xlsx)", directory=dir
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

        if not os.path.isfile(self.txt_results_file.text()):
            QMessageBox.critical(self, "Error", "Cannot find results file")
            return

        # Have to read the results sheet to find what tests are performed, then present as a series of checkboxes

        try:
            with ExcelInterface(filename=self.txt_results_file.text()) as excel:
                # check excel first

                if not excel.check_excel_available():
                    QMessageBox.critical(
                        self, "Error", "Unable to write to results sheet, is it open?"
                    )
                    return

                if not self.test_connections():
                    return

                if (
                    self.uut.manufacturer.startswith("TEK")
                    and not self.cb_filter_low_ranges.isChecked()
                ):
                    reply = QMessageBox.question(
                        self,
                        "Use filter",
                        "Recommend using filter for Tek scopes. Enable?",
                    )

                    if reply == QMessageBox.StandardButton.Yes:
                        self.cb_filter_low_ranges.setChecked(True)

                test_names = excel.get_test_types()

                selector = IndividualTestSelector(test_names=list(test_names))
                selector.show()

                print(selector.selected_tests)

                if not selector.selected_tests:
                    return  # cancelled

                # Now we have the list of test names, we need to get the associated test rows

                # Have to trap CURS tests, as the results come from the DCV tests. If DCV selected automatically select CURS, if CURS only display message

                add_cursor = False

                if "CURS" in list(test_names):
                    dcv = "DCV" in selector.selected_tests
                    curs = "CURS" in selector.selected_tests

                    if curs and not dcv:
                        QMessageBox.critical(
                            self, "Error", "Cursor tests rely on results from DCV"
                        )
                        return

                    if dcv and not curs:
                        add_cursor = True
                        self.cb_skip_rows.setChecked(False)

                test_steps = []

                for name in selector.selected_tests:
                    rows = excel.get_test_rows(name)
                    test_steps.extend(iter(rows))

                if add_cursor:
                    rows = excel.get_test_rows("CURS")
                    test_steps.extend(iter(rows))

                self.do_parallel = False

                if int(self.cmb_number_channels.currentText()) <= 4:
                    if "DCV" in selector.selected_tests:
                        self.do_parallel = True
                    if "DCV-BAL" in selector.selected_tests:
                        self.do_parallel = True
                    if "CURS" in selector.selected_tests:
                        self.do_parallel = True

                test_rows = sorted(test_steps)

                self.perform_oscilloscope_tests(test_rows=test_rows)
        except BadZipFile:
            QMessageBox.critical(
                self,
                "Error",
                "Results file is corrupted. Copy from previous version in backups folder (subfolder of current results folder)",
            )

    def perform_oscilloscope_tests(self, test_rows: list) -> None:
        """
        perform_tests
        Perform the actual oscilloscope tests

        Args:
            test_rows (list): _description_

        Returns:
            _type_: _description_
        """

        self.progress_test.setVisible(True)
        self.progress_test.setValue(0)
        self.btn_abort.setVisible(True)
        self.statusbar.showMessage("")

        self.set_control_state(False)

        if self.do_parallel:
            button = QMessageBox.question(
                self,
                "Connect channels in parallel",
                "Would you like to connect all channels in parallel for voltage tests",
            )
            if button == QMessageBox.StandardButton.Yes:
                self.do_parallel = True
            else:
                self.do_parallel = False

        self.tester = TestOscilloscope(
            calibrator=self.calibrator,
            ks33250=self.ks33250,
            ks3458=self.ks3458,
            uut=self.uut,
            simulating=self.cb_simulating.isChecked(),
        )

        self.tester.use_filter = self.cb_filter_low_ranges.isChecked()

        self.tester.test_progress.connect(self.update_progress)
        self.tester.current_test.connect(self.current_test_message)

        try:
            num_channels = int(self.cmb_number_channels.currentText())
        except ValueError:
            QMessageBox.critical(self, "Error", "Invalid number of channels")
            return

        self.tester.run_tests(
            filename=self.txt_results_file.text(),
            test_rows=test_rows,
            uut_address=self.txt_uut_addr.text(),
            parallel_channels=self.do_parallel,
            skip_completed=self.cb_skip_rows.isChecked(),
            num_channels=num_channels,
        )

        self.progress_test.setVisible(False)
        self.statusbar.showMessage("Finished")
        self.btn_abort.setVisible(False)

        self.set_control_state(True)

        QMessageBox.information(self, "Finished", "Completed, check results")

    def update_progress(self, progress: float) -> None:
        self.progress_test.setValue(int(progress))
        QApplication.processEvents()

    def current_test_message(self, message: str) -> None:
        self.statusbar.showMessage(message)

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

    def check_test_required(self, test_name: str) -> bool:
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

        return test_name in valid_tests

    def hide_excel_rows(self) -> None:
        with ExcelInterface(filename=self.txt_results_file.text()) as excel:
            check = excel.check_channel_rows()
            if check:
                excel.hide_excel_rows(
                    channel=int(self.cmb_number_channels.currentText())
                )

    def check_excel_button(self) -> None:
        """
        Results file has chenged, enable or disable the hide excel rows button
        """

        if os.path.isfile(self.txt_results_file.text()):
            try:
                with ExcelInterface(filename=self.txt_results_file.text()) as excel:
                    check = excel.check_channel_rows()

                    self.btn_hide_excel_rows.setEnabled(check)
            except Exception:
                self.btn_hide_excel_rows.setEnabled(False)
        else:
            self.btn_hide_excel_rows.setEnabled(False)

    def set_control_state(self, state: bool) -> None:
        """
        Many of the buttons need to be disabled during test, and reenabled after
        """

        self.btn_perform_tests.setEnabled(state)
        self.btn_view_results.setEnabled(state)
        self.btn_browse_results.setEnabled(state)
        self.btn_select_uut_addr.setEnabled(state)
        self.txt_results_file.setEnabled(state)
        self.group_hardware.setEnabled(state)
        self.cb_skip_rows.setEnabled(state)

        # Hide excel rows is different

        if not state:
            self.btn_hide_excel_rows.setEnabled(False)
        else:
            self.check_excel_button()

    def abort_test(self) -> None:
        """
        abort_test
        Abort buton pressed.
        The test is running in a subclass, so need to send the signal to that class
        """

        # Set the flag in the class, the test loops check for abort
        self.tester.abort_test = True


if __name__ == "__main__":
    app = QApplication([])
    window = UI()
    window.show()
    app.exec()
