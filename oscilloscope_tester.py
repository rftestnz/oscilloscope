"""
  Perform the main oscilloscope tests

  Tests vary by manufacturer
"""

import math
import time
from datetime import datetime
from typing import Dict, List

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import QDialog, QMessageBox, QInputDialog

from drivers.excel_interface import ExcelInterface
from drivers.fluke_5700a import Fluke5700A
from drivers.keysight_scope import DSOX_FAMILY, Keysight_Oscilloscope
from drivers.Ks3458A import Ks3458A, Ks3458A_Function
from drivers.Ks33250A import Ks33250A
from drivers.meatest_m142 import M142
from drivers.rf_signal_generator import RF_Signal_Generator
from drivers.rohde_shwarz_scope import RohdeSchwarz_Oscilloscope
from drivers.tek_scope import Tek_Acq_Mode, Tektronix_Oscilloscope

DELAY_PERIOD = 0.001  # 1 ms


class TestOscilloscope(QDialog, object):

    current_test = pyqtSignal(object)
    test_progress = pyqtSignal(object)

    def __init__(
        self,
        calibrator: Fluke5700A | M142,
        ks33250: Ks33250A,
        ks3458: Ks3458A,
        uut: Keysight_Oscilloscope | RohdeSchwarz_Oscilloscope | Tektronix_Oscilloscope,
        simulating: bool,
    ) -> None:

        super().__init__()

        self.calibrator = calibrator
        self.ks33250 = ks33250
        self.ks3458 = ks3458
        self.mxg = RF_Signal_Generator()
        self.uut = uut
        self.simulating = simulating

        self.number_tests = 0
        self.test_number = 0  # current test

        self.cursor_results: list = []

        self.abort_test = False

    def local_all(self) -> None:
        """
        local_all
        Set all instruments back to local
        """

        self.calibrator.go_to_local()
        self.ks33250.go_to_local()
        self.ks3458.go_to_local()

    def load_uut_driver(self, address: str, simulating: bool = False) -> bool:
        """
        load_uut_driver
        Use a generic driver to figure out which driver of the scope should be used
        """

        if simulating:
            self.uut = Tektronix_Oscilloscope(simulate=simulating)
            self.uut.model = "MSO5104B"
            self.uut.open_connection()
            return True

        # TODO when using the SCPI ID class, it affects all subsequent uses

        with Keysight_Oscilloscope(simulate=False) as scpi_uut:
            scpi_uut.visa_address = address
            scpi_uut.open_connection()
            manufacturer = scpi_uut.get_manufacturer()
            num_channels = scpi_uut.get_number_channels()

        if manufacturer == "":
            QMessageBox.critical(
                self, "Error", "Unable to contact UUT. Is address correct?"
            )
            return False
        elif manufacturer == "KEYSIGHT":
            self.uut = Keysight_Oscilloscope(simulate=False)

        elif manufacturer == "TEKTRONIX":
            self.uut = Tektronix_Oscilloscope(simulate=False)
            self.uut.visa_address = address
            self.uut.open_connection()
            self.num_channels = self.uut.get_number_channels()

        elif manufacturer == "ROHDE&SCHWARZ":
            self.uut = RohdeSchwarz_Oscilloscope(simulate=False)
            num_channels = 4
        else:
            QMessageBox.critical(
                self, "Error", f"No driver for {manufacturer}. Using Tektronix driver"
            )
            self.uut = Tektronix_Oscilloscope(simulate=False)

        # Make sure the address is set correctly
        self.uut.visa_address = address
        self.uut.num_channels = num_channels

        return True

    def consolidate_dcv_tests(self, test_steps: List, filename: str) -> List:
        """
        consolidate_dcv_tests
        Go through the test steps and all of the DCV and DCV-BAL tests group by channel to minimize channel swapping

        Args:
            test_steps (List): _description_
            filename (str): _description_

        Returns:
            List: _description_
        """

        sorted_steps = []

        # we only have the row number, so have to read it again

        # TODO full settings are available, use them

        # Go through first to get the DCV tests

        with ExcelInterface(filename=filename) as excel:
            for step in test_steps:
                test_name, channel = excel.get_test_name(step)
                if test_name.startswith("DCV"):
                    sorted_steps.append(
                        (channel, step)
                    )  # Append channel first to help with sorting

            new_list = sorted(sorted_steps)

            print(new_list)

            # and make a new list with just the row numbers

            sorted_steps = set()

            for step in new_list:
                sorted_steps.add(step[1])

            # And the rest of the steps

            for step in test_steps:
                test_name, channel = excel.get_test_name(step)
                if test_name.startswith("DCV"):
                    sorted_steps.add(step)

        print(list(sorted_steps))

        return list(sorted_steps)

    def run_tests(
        self,
        filename: str,
        test_rows: List,
        uut_address: str,
        parallel_channels: bool = False,
        skip_completed: bool = False,
    ) -> None:
        """
        run_tests
        Main test sequencer
        From the list of test rows, work out the test names and call the appropriate functions

        Args:
            filename (str): _description_
            test_rows (List): _description_
        """

        # TODO the test rows are generated from the list of names
        # then need to be converted back into a list of test names

        # Get the test names

        global test_number

        self.test_number = 0
        self.number_tests = len(test_rows)

        self.abort_test = False

        self.load_uut_driver(address=uut_address, simulating=self.simulating)

        with ExcelInterface(filename=filename) as excel:
            excel.backup()

            # first update the model and serial

            self.uut.open_connection()

            # If the named range doesn't exist, nothing is written
            excel.write_data(data=self.uut.model, named_range="Model")
            excel.write_data(data=self.uut.serial, named_range="Serial")

            test_names = set()

            for row in test_rows:
                settings = excel.get_volt_settings(row=row)
                test_names.add(settings.function)

            # Get all the tests. If there are cursor tests, then automatically select them if dcv selected as they cannot be done in isolation
            all_tests = excel.get_test_types()

            if "CURS" in all_tests and "DCV" in test_names:
                test_names.add(
                    "CURS"
                )  # It is a set, so doesn't matter if it was already in

            # python sets are unordered, and not deterministic. We need the set to be in a specific order for the sequencer
            # eg can't do cursor tests before dcv

            ordered_test_names = [
                name for name in excel.supported_test_names if name in test_names
            ]

            # Would like to join DCV and DCV-BAL into same test for consolidating.

            for test_name in ordered_test_names:
                testing_rows = excel.get_test_rows(test_name)
                # At the moment we only do full tests, so we can get the test rows from the excel sheet

                # TODO use functional method

                if "DCV" in test_name:
                    sorted_rows = self.consolidate_dcv_tests(
                        test_rows, filename=filename
                    )
                    if not self.test_dcv(
                        filename=filename,
                        test_rows=sorted_rows,
                        parallel_channels=parallel_channels,
                        skip_completed=skip_completed,
                    ):
                        break

                elif test_name == "POS":
                    if not self.test_position(
                        filename=filename,
                        test_rows=testing_rows,
                        parallel_channels=parallel_channels,
                    ):
                        break

                elif test_name == "BAL":
                    if not self.test_dc_balance(
                        filename=filename, test_rows=testing_rows
                    ):
                        break

                elif test_name == "CURS":
                    if not self.test_cursor(filename=filename, test_rows=testing_rows):
                        break

                elif test_name == "RISE":
                    if not self.test_risetime(
                        filename=filename, test_rows=testing_rows
                    ):
                        break

                elif test_name == "TIME":
                    if not self.test_timebase(filename=filename, row=testing_rows[0]):
                        break

                elif test_name == "TRIG":
                    if not self.test_trigger_sensitivity(
                        filename=filename, test_rows=testing_rows
                    ):
                        break

                elif test_name == "IMP":
                    if not self.test_impedance(
                        filename=filename, test_rows=testing_rows
                    ):
                        break

                elif test_name == "NOISE":
                    if not self.test_random_noise(
                        filename=filename,
                        test_rows=test_rows,
                        skip_completed=skip_completed,
                    ):
                        break

                elif test_name == "DELTAT":
                    if not self.test_delta_time(filename=filename, test_rows=test_rows):
                        break

                elif test_name == "THR":
                    if not self.test_threshold(
                        filename=filename, test_rows=testing_rows
                    ):
                        break

        self.local_all()

    def update_test_progress(self) -> None:
        """
        update_progress
        Increment the test count and emit signal for main ui to update progress bar
        """

        self.test_number += 1
        self.test_progress.emit(100 * self.test_number / self.number_tests)

    def test_connections(self, check_3458: bool) -> Dict:
        """
        test_connections
        Check all of the instruments are connected

        Args:
            check_3458 (bool): 3458 is only used for impedance, which few oscilloscopes require measurement

        Returns:
            bool: _description_
        """

        fluke_5700a_conn = self.calibrator.is_connected()
        ks33250_conn = self.ks33250.is_connected()
        uut_conn = self.uut.is_connected()
        ks3458_conn = self.ks3458.is_connected() if check_3458 else False

        return {
            "FLUKE_5700A": fluke_5700a_conn,
            "33250A": ks33250_conn,
            "DSO": uut_conn,
            "3458": ks3458_conn,
        }

    def test_dc_balance(self, filename: str, test_rows: List) -> bool:
        """
        test_dc_balance
        Test the dc balance of each channel with no signal applied

        Args:
            filename (str): _description_
            test_rows (int): _description_
        """

        # no equipment required

        self.current_test.emit("Testing: DCV Balance")

        response = QMessageBox.information(
            self,
            "Connections",
            "Remove inputs from all channels",
            buttons=QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
        )

        if response == QMessageBox.StandardButton.Cancel:
            return False

        self.uut.reset()

        self.uut.set_acquisition(32)

        self.uut.set_timebase(200e-6)

        with ExcelInterface(filename=filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False

            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                settings = excel.get_volt_settings()
                units = excel.get_units()

                if settings.function == "BAL":
                    self.uut.set_channel(
                        chan=int(settings.channel), enabled=True, only=True
                    )
                    self.uut.set_voltage_scale(
                        chan=int(settings.channel), scale=settings.scale
                    )
                    self.uut.set_voltage_offset(chan=int(settings.channel), offset=0)
                    self.uut.set_channel_coupling(
                        chan=int(settings.channel), coupling=settings.coupling
                    )

                    reading = self.uut.measure_voltage(
                        chan=int(settings.channel), delay=2
                    )

                    if units == "mV":
                        reading *= 1000

                    excel.write_result(reading, col=results_col)
                    self.update_test_progress()

        self.uut.reset()

        return True

    def test_delta_time(self, filename: str, test_rows: List) -> bool:
        """
        test_delta_time
        Test delta time function
        for Tek scope

        Args:
            filename (str): _description_
            test_rows (List): _description_

        Returns:
            bool: _description_
        """

        self.current_test.emit("Testing: Delta Time")

        connections = self.test_connections(check_3458=False)  # Always required

        if not connections["RFGEN"]:
            QMessageBox.critical(self, "Error", "Cannot find RF Generator")
            return False

        if not connections["33250A"]:
            QMessageBox.critical(self, "Error", "Cannot find 33250A")
            return False

        self.uut.open_connection()
        self.uut.reset()

        # For the moment, this is a Tek MSO5000  special test, so commands written directly here. If any more
        # are required, put into driver

        self.uut.set_acquisition(1)

        last_channel = -1
        last_generator = 0

        self.ks33250.set_output_z("50")

        with ExcelInterface(filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False
            excel.find_units_col(test_rows[0])

            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                units = excel.get_units()

                settings = excel.get_sample_rate_settings()

                if settings.channel != last_channel:
                    response = QMessageBox.information(
                        self,
                        "Connections",
                        f"Connect Sig Gen to Channel {settings.channel}",
                        buttons=QMessageBox.StandardButton.Ok
                        | QMessageBox.StandardButton.Cancel,
                    )

                    if response == QMessageBox.StandardButton.Cancel:

                        return False

                    last_generator = "MXG"

                    last_channel = settings.channel

                self.uut.set_channel(chan=settings.channel, enabled=True, only=True)
                self.uut.set_voltage_scale(chan=settings.channel, scale=settings.scale)
                self.uut.set_channel_coupling(
                    chan=settings.channel, coupling=settings.coupling
                )
                self.uut.set_channel_impedance(chan=settings.channel, impedance="50")
                self.uut.set_trigger_level(chan=settings.channel, level=0)

                self.uut.write(f"HORIZONTAL:MODE:SAMPLERATE {settings.sample_rate}")

                # Have to adjust the record length to get the right timebase setting

                self.uut.write("HOR:MODE MANUAL")
                recordlength = 10 * settings.sample_rate * settings.timebase
                self.uut.write(f"HOR:MODE:RECORDLENGTH {recordlength}")

                if settings.frequency > 250000:
                    if last_generator != "MXG":
                        response = QMessageBox.information(
                            self,
                            "Connections",
                            f"Connect Sig Gen output to channel {settings.channel}",
                            buttons=QMessageBox.StandardButton.Ok
                            | QMessageBox.StandardButton.Cancel,
                        )

                        if response == QMessageBox.StandardButton.Cancel:

                            return False

                    self.mxg.set_frequency(settings.frequency)
                    self.mxg.set_level(settings.voltage / 2.82, units="V")
                    self.mxg.set_output_state(True)

                    last_generator = "MXG"
                else:
                    if last_generator != "33250A":
                        response = QMessageBox.information(
                            self,
                            "Connections",
                            f"Connect 33250A output to channel {settings.channel}",
                            buttons=QMessageBox.StandardButton.Ok
                            | QMessageBox.StandardButton.Cancel,
                        )

                        if response == QMessageBox.StandardButton.Cancel:
                            return False

                        last_generator = "33250A"
                    self.ks33250.set_sin(
                        frequency=settings.frequency, amplitude=settings.voltage / 2.82
                    )
                    self.ks33250.enable_output(True)

                time.sleep(0.25)

                self.uut.write("MEASU:MEAS1:TYPE DELAY")
                self.uut.write(f"MEASU:MEAS1:SOURCE CH{settings.channel}")
                self.uut.write(f"MEASU:MEAS1:SOURCE2 CH{settings.channel}")
                self.uut.write("MEASU:MEAS1:DELAY:EDGE1 RISE")
                self.uut.write("MEASU:MEAS1:DELAY:EDGE2 FALL")

                self.uut.write("MEASURE:STATISTICS:MODE MEANSTDDEV")
                self.uut.write("MEASURE:STATISTICS:WEIGHTING 1000")
                self.uut.write("MEASUREMENT:STATISTICS:COUNT RESET")

                self.uut.write("MEASU:MEAS1:STATE ON")

                self.uut.write("MEASU:MEAS1:DISPLAYSTAT:ENABLE ON")

                time.sleep(10)

                try:
                    result = float(
                        self.uut.query("MEASU:MEAS1:STDDEV?").strip()
                    )  # remove LF

                    if units[0] == "p":
                        result *= 1_000_000_000_000
                    elif units[0] == "n":
                        result *= 1_000_000_000
                    elif units[0] == "u":
                        result *= 1_000_000

                    excel.write_result(result=result, col=results_col, save=True)
                except ValueError:
                    pass

                self.update_test_progress()

                self.mxg.set_output_state(False)
                self.ks33250.enable_output(False)

            excel.save_sheet()

        return True

    def test_random_noise(
        self, filename: str, test_rows: List, skip_completed: bool = False
    ) -> bool:
        """
        test_random_noise
        Test sampling random noise
        Tek scopes

        Args:
            filename (str): _description_
            test_rows (List): _description_

        Returns:
            bool: _description_
        """

        self.current_test.emit("Testing: Random noise sample acquisition")

        # No equipment required
        self.uut.open_connection()
        self.uut.reset()

        if self.uut.model.startswith("MSO5"):
            self.uut.set_sample_rate("6.25G")  # type: ignore

        self.uut.set_acquisition(16)

        with ExcelInterface(filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False

            excel.find_units_col(test_rows[0])

            response = QMessageBox.information(
                self,
                "Connections",
                "Remove inputs from all channels",
                buttons=QMessageBox.StandardButton.Ok
                | QMessageBox.StandardButton.Cancel,
            )

            if response == QMessageBox.StandardButton.Cancel:

                return False

            row_count = 0

            self.uut.set_horizontal_mode("MAN", 2000000)  # type: ignore

            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                if skip_completed:
                    if not excel.check_empty_result(col=results_col):
                        continue

                units = excel.get_units()

                settings = excel.get_volt_settings()

                # if settings.bandwidth == "250M":
                #    continue

                channel = int(settings.channel)

                self.uut.set_channel(chan=channel, enabled=True, only=True)  # type: ignore
                self.uut.set_channel_impedance(
                    chan=channel, impedance=settings.impedance  # type: ignore
                )
                self.uut.set_channel_bw_limit(chan=channel, bw_limit=settings.bandwidth)  # type: ignore

                if settings.acq_mode:
                    if settings.acq_mode.upper() == "HIRES":
                        self.uut.set_acquisition_mode(Tek_Acq_Mode.HIRES)  # type: ignore
                    elif settings.acq_mode.upper() == "SAMPLE":
                        self.uut.set_acquisition(Tek_Acq_Mode.SAMPLE)  # type: ignore
                    else:
                        self.uut.set_acquisition_mode(Tek_Acq_Mode.AVERAGE)  # type: ignore
                else:
                    self.uut.set_acquisition_mode(Tek_Acq_Mode.AVERAGE)  # type: ignore

                self.uut.set_voltage_position(
                    chan=channel, position=settings.scale * 0.34
                )  # 340 mdiv

                rnd = self.uut.measure_rms_noise(chan=settings.channel, delay=10)  # type: ignore

                self.uut.measure_clear()
                self.uut.set_voltage_position(
                    chan=channel, position=settings.scale * 0.36
                )  # 360 mdiv

                avg = self.uut.measure_rms_noise(chan=settings.channel, delay=10)  # type: ignore

                result = (rnd + avg) / 2

                if units.startswith("m"):
                    result *= 1000

                excel.write_result(result=result, col=results_col, save=True)

                self.update_test_progress()

                row_count += 1
                print(row)
                # if row_count > 5:
                #    break

            excel.save_sheet()

        return True

    def test_threshold(self, filename: str, test_rows: List) -> bool:
        """
        test_threshold
        Test digital threshold

        Args:
            filename (str): _description_
            test_rows (List): _description_

        Returns:
            bool: _description_
        """

        self.current_test.emit("Testing: Digital Threshold")

        connections = self.test_connections(
            check_3458=False
        )  # Don't need 3458 for this test

        # require self.calibrator

        if not connections["FLUKE_5700A"]:
            QMessageBox.critical(self, "Error", "Cannot find self.calibrator")
            return False

        self.uut.open_connection()
        self.uut.reset()

        self.uut.set_timebase(1e-3)

        self.uut.set_digital_channel_on(chan=0, all_channels=True)  # type: ignore

        response = QMessageBox.information(
            self,
            "Connections",
            "Connect self.calibrator output to digital IO Pods",
            buttons=QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
        )

        if response == QMessageBox.StandardButton.Cancel:
            return False

        with ExcelInterface(filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False
            excel.find_units_col(test_rows[0])
            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                settings = excel.get_threshold_settings()

                # Tests are performed in blocks of 8 channels

                # Get the starting threshold, and direction

                voltage = settings.voltage
                delta = -0.01 if settings.polarity == "NEG" else 0.01

                for _ in range(50):
                    reading = self.uut.measure_digital_channels(pod=settings.pod)  # type: ignore
                    if (
                        settings.polarity == "POS"
                        and reading == 1
                        or settings.polarity != "POS"
                        and reading == 0
                    ):
                        excel.write_result(voltage)
                        break
                    voltage += delta

        return True

    def test_impedance(self, filename: str, test_rows: List) -> bool:
        """
        test_impedance
        Test the input impedance of the channels

        Args:
            filename (str): _description_
            test_rows (List): _description_

        Returns:
            bool: _description_
        """

        self.current_test.emit("Testing: Input Impedance")

        connections = self.test_connections(check_3458=True)  # Always required

        if not connections["3458"]:
            QMessageBox.critical(self, "Error", "Cannot find 3458A")
            return False

        self.uut.open_connection()
        self.uut.reset()

        self.ks3458.open_connection()
        self.ks3458.reset()

        last_channel = -1

        # Turn off all channels but 1
        for chan in range(self.uut.num_channels):
            self.uut.set_channel(chan=chan + 1, enabled=chan == 0)

        self.uut.set_acquisition(1)

        with ExcelInterface(filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False
            excel.find_units_col(test_rows[0])
            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                settings = excel.get_volt_settings()

                channel = int(settings.channel)
                units = excel.get_units()

                if channel > self.uut.num_channels:
                    continue

                if channel != last_channel:
                    response = QMessageBox.information(
                        self,
                        "Connections",
                        f"Connect 3458A Input to self.uut Ch {channel}, and sense",
                        buttons=QMessageBox.StandardButton.Ok
                        | QMessageBox.StandardButton.Cancel,
                    )

                    if response == QMessageBox.StandardButton.Cancel:
                        return False

                    if last_channel > 0:
                        # changed channel to another, but not channel 1. reset all of the settings on the channel just measured
                        self.uut.set_voltage_scale(chan=last_channel, scale=1)
                        self.uut.set_voltage_offset(chan=last_channel, offset=0)
                        self.uut.set_channel(chan=last_channel, enabled=False)
                        self.uut.set_channel_bw_limit(chan=last_channel, bw_limit=False)
                        self.uut.set_channel(chan=channel, enabled=True)
                        self.uut.set_channel_impedance(
                            chan=last_channel, impedance="1M"
                        )  # always
                    last_channel = channel

                self.uut.set_voltage_scale(chan=channel, scale=settings.scale)
                self.uut.set_voltage_offset(chan=channel, offset=settings.offset)
                self.uut.set_channel_impedance(
                    chan=channel, impedance=settings.impedance
                )
                self.uut.set_channel_bw_limit(chan=channel, bw_limit=settings.bandwidth)

                time.sleep(0.5)

                reading = self.ks3458.measure(function=Ks3458A_Function.R4W)["Average"]  # type: ignore
                if units.lower().startswith("k"):
                    reading /= 1000
                if units.upper().startswith("M"):
                    reading /= 1_000_000

                excel.write_result(reading, col=results_col)

                self.update_test_progress()

            # Turn off all channels but 1
            for chan in range(self.uut.num_channels):
                self.uut.set_channel(chan=chan + 1, enabled=chan == 0)
                self.uut.set_channel_bw_limit(chan=chan, bw_limit=False)

            self.uut.reset()
            self.uut.close()

        return True

    def test_dcv(
        self,
        filename: str,
        test_rows: List,
        parallel_channels: bool = False,
        skip_completed: bool = False,
    ) -> bool:
        # sourcery skip: extract-method, low-code-quality
        """
        test_dcv
        Perform the basic DC V tests
        Set the self.calibrator to the voltage, allow the scope to stabilizee, then read the cursors or measurement values
        """

        self.current_test.emit("Testing: DC Voltage")

        last_channel = -1

        connections = self.test_connections(
            check_3458=False
        )  # Don't need 3458 for this test

        acquisitions = 32

        # require self.calibrator

        if not connections["FLUKE_5700A"]:
            QMessageBox.critical(self, "Error", "Cannot find self.calibrator")
            return False

        self.uut.open_connection()
        self.uut.reset()

        self.uut.set_timebase(200e-6)

        self.cursor_results = []  # save results for cursor tests

        filter_connected = False  # noqa: F841

        if parallel_channels:
            response = QMessageBox.information(
                self,
                "Connections",
                "Connect Calibrator output to all channels in parallel",
                buttons=QMessageBox.StandardButton.Ok
                | QMessageBox.StandardButton.Cancel,
            )

            if response == QMessageBox.StandardButton.Cancel:

                return False

        # Turn off all channels but 1
        for chan in range(self.uut.num_channels):
            self.uut.set_channel(chan=chan + 1, enabled=chan == 0)

        self.uut.set_acquisition(acquisitions)

        with ExcelInterface(filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False
            excel.find_units_col(test_rows[0])

            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                if skip_completed:
                    if not excel.check_empty_result(results_col):
                        continue

                settings = excel.get_volt_settings()

                units = excel.get_units()

                self.calibrator.set_voltage_dc(0)

                channel = int(settings.channel)

                if channel > self.uut.num_channels:
                    continue

                if channel != last_channel:
                    if last_channel > 0:
                        # changed channel to another, but not channel 1. reset all of the settings on the channel just measured
                        self.uut.set_voltage_scale(chan=last_channel, scale=1)
                        self.uut.set_voltage_offset(chan=last_channel, offset=0)
                        self.uut.set_channel(chan=last_channel, enabled=False)
                        self.uut.set_channel_bw_limit(chan=last_channel, bw_limit=False)
                        self.uut.set_channel(chan=channel, enabled=True)
                        self.uut.set_channel_impedance(
                            chan=last_channel, impedance="1M"
                        )  # always

                    self.uut.set_voltage_scale(chan=channel, scale=5)
                    self.uut.set_voltage_offset(chan=channel, offset=0)

                    self.uut.set_cursor_xy_source(chan=1, cursor=1)
                    self.uut.set_cursor_position(cursor="X1", pos=0)
                    if not parallel_channels:
                        response = QMessageBox.information(
                            self,
                            "Connections",
                            f"Connect Calibrator output to channel {channel}",
                            buttons=QMessageBox.StandardButton.Ok
                            | QMessageBox.StandardButton.Cancel,
                        )
                        if response == QMessageBox.StandardButton.Cancel:
                            return False
                    last_channel = channel

                self.uut.set_channel(chan=channel, enabled=True)
                self.uut.set_voltage_scale(chan=channel, scale=settings.scale)
                self.uut.set_voltage_offset(chan=channel, offset=settings.offset)

                if settings.impedance:
                    self.uut.set_channel_impedance(
                        chan=channel, impedance=settings.impedance
                    )

                if settings.bandwidth:
                    self.uut.set_channel_bw_limit(
                        chan=channel, bw_limit=settings.bandwidth
                    )
                else:
                    self.uut.set_channel_bw_limit(chan=channel, bw_limit=False)

                if settings.invert:
                    # already casted to a bool
                    self.uut.set_channel_invert(chan=channel, inverted=settings.invert)
                else:
                    self.uut.set_channel_invert(chan=channel, inverted=False)

                """
                if not filter_connected and settings.voltage < 0.1:
                    self.calibrator.standby()
                    sg.popup(
                        "Connect filter capacitor to input channel", background_color="blue"
                    )
                    filter_connected = True
                elif filter_connected and settings.voltage >= 0.1:
                    self.calibrator.standby()
                    sg.popup(
                        "Remove filter capacitor from input channel",
                        background_color="blue",
                    )
                    filter_connected = False
                """

                if self.uut.keysight or settings.function == "DCV-BAL":
                    if settings.function == "DCV-BAL":
                        # Non keysight, apply the half the voltage and the offset then do the reverse

                        self.calibrator.set_voltage_dc(settings.voltage)

                    # 0V test
                    self.calibrator.operate()

                    self.uut.set_acquisition(1)

                    if not self.simulating:
                        time.sleep(0.2)

                    self.uut.set_acquisition(acquisitions)

                    if not self.simulating:
                        time.sleep(1)

                    if settings.scale <= 0.005:
                        self.uut.set_acquisition(64)
                        time.sleep(5)  # little longer to average for sensitive scales

                    if self.uut.keysight:
                        voltage1 = self.uut.read_cursor_avg()

                    self.uut.measure_clear()
                    reading1 = self.uut.measure_voltage(chan=channel, delay=2)

                if settings.function == "DCV-BAL":
                    # still set up for the + voltage

                    self.calibrator.set_voltage_dc(-settings.voltage)
                    self.uut.set_voltage_offset(chan=channel, offset=-settings.offset)
                else:
                    self.calibrator.set_voltage_dc(settings.voltage)

                self.calibrator.operate()

                self.uut.set_acquisition(1)

                if not self.simulating:
                    time.sleep(0.2)

                self.uut.set_acquisition(acquisitions)

                if not self.simulating:
                    time.sleep(1)

                if settings.scale <= 0.005:
                    self.uut.set_acquisition(64)
                    time.sleep(5)  # little longer to average for sensitive scales

                self.uut.measure_clear()

                reading = self.uut.measure_voltage(chan=channel, delay=3)

                if self.uut.keysight and self.uut.family != DSOX_FAMILY.DSO5000:  # type: ignore
                    voltage2 = self.uut.read_cursor_avg()

                    self.cursor_results.append(
                        {
                            "chan": channel,
                            "scale": settings.scale,
                            "result": voltage2 - voltage1,  # type: ignore
                        }
                    )

                self.calibrator.standby()

                if units and units.startswith("m"):
                    reading *= 1000

                if settings.function == "DCV-BAL":
                    if units.startswith("m"):
                        reading1 *= 1000
                    diff = reading1 - reading
                    excel.write_result(diff, col=results_col)  # auto saving
                else:
                    # DCV (offset) test. 0V is measured for the cursors only
                    excel.write_result(reading, col=results_col)

                self.update_test_progress()

            self.calibrator.reset()
            self.calibrator.close()

            # Turn off all channels but 1
            for chan in range(self.uut.num_channels):
                self.uut.set_channel(chan=chan + 1, enabled=chan == 0)
                self.uut.set_channel_bw_limit(chan=chan, bw_limit=False)

            self.uut.reset()
            self.uut.close()

        return True

    def test_cursor(self, filename: str, test_rows: List) -> bool:
        """
        test_cursor
        Dual cursor test. Measure voltage with no voltage applied, apply voltage, measure again, record the difference
        Measurements are taken during the DCV test, and recalled here

        Args:
            filename (str): _description_
            test_rows (List): _description_
        """

        self.current_test.emit("Testing: Cursor position")

        # no equipment as using buffered results

        with ExcelInterface(filename) as excel:
            excel.find_units_col(test_rows[0])
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False
            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                settings = excel.get_volt_settings()

                if len(self.cursor_results):
                    for res in self.cursor_results:
                        if (
                            res["chan"] == settings.channel
                            and res["scale"] == settings.scale
                        ):
                            units = excel.get_units()
                            result = res["result"]
                            if units.startswith("m"):
                                result *= 1000
                            excel.write_result(result, save=False, col=results_col)
                            self.update_test_progress()
                            break

            excel.save_sheet()

        return True

    def test_position(
        self, filename: str, test_rows: List, parallel_channels: bool = False
    ) -> bool:
        """
        test_position
        Test vertical position

        Args:
            filename (str): _description_
            test_rows (List): _description_

        Returns:
            _type_: _description_
        """

        self.current_test.emit("Testing: DC Position")

        connections = self.test_connections(
            check_3458=False
        )  # Don't need 3458 for this test

        # require self.calibrator

        if not connections["FLUKE_5700A"]:
            QMessageBox.critical(self, "Error", "Cannot find self.calibrator")
            return False

        self.uut.reset()

        self.uut.set_timebase(200e-6)

        self.uut.set_acquisition(32)

        last_channel = -1

        with ExcelInterface(filename=filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False

            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                settings = excel.get_volt_settings()

                if settings.channel != last_channel and not parallel_channels:
                    response = QMessageBox.information(
                        self,
                        "Connections",
                        f"Connect Calibrator output to channel {settings.channel}",
                        buttons=QMessageBox.StandardButton.Ok
                        | QMessageBox.StandardButton.Cancel,
                    )

                    if response == QMessageBox.StandardButton.Cancel:
                        return False

                    last_channel = settings.channel

                self.uut.set_channel(
                    chan=int(settings.channel), enabled=True, only=True
                )
                self.uut.set_channel_bw_limit(chan=int(settings.channel), bw_limit=True)
                self.uut.set_voltage_scale(
                    chan=int(settings.channel), scale=settings.scale
                )
                pos = -4 if settings.offset > 0 else 4
                self.uut.set_voltage_position(
                    chan=int(settings.channel), position=pos
                )  # divisions
                self.uut.set_voltage_offset(
                    chan=int(settings.channel), offset=settings.offset
                )

                self.uut.set_acquisition(1)  # Too slow to adjust otherwise
                self.calibrator.set_voltage_dc(settings.voltage)

                self.calibrator.operate()

                self.uut.set_acquisition(32)

                # self.uut.measure_voltage_clear()

                # reading = self.uut.measure_voltage(chan=int(settings.channel), delay=2)
                response = QMessageBox.question(
                    parent=self,
                    title="Check cursor",
                    text="Trace within 0.2 div of center?",
                    buttons=QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No,
                )

                result = (
                    "Pass" if response == QMessageBox.StandardButton.Yes else "Fail"
                )

                self.calibrator.standby()

                excel.write_result(result=result, col=results_col)
                self.update_test_progress()

        self.calibrator.reset()
        self.calibrator.close()

        self.uut.reset()
        self.uut.close()

        return True

    # DELAY_PERIOD = 0.00099998  # 1 ms
    # DELAY_PERIOD = 0.00100002  # 1 ms
    DELAY_PERIOD = 0.001  # 1 ms

    def test_timebase(self, filename: str, row: int) -> bool:
        # sourcery skip: low-code-quality
        """
        test_timebase
        Test the timebase. Simple single row test

        Args:
            row (int): _description_
        """

        global current_test_text

        self.current_test.emit("Testing: Timebase")

        connections = self.test_connections(
            check_3458=False
        )  # Don't need 3458 for this test

        # require RF gen

        if not connections["33250A"]:
            QMessageBox.critical(self, "Error", "Cannot find 33250A Generator")
            return False

        response = QMessageBox.information(
            self,
            "Connections",
            "Connect 33250A output to Ch1",
            buttons=QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
        )

        if response == QMessageBox.StandardButton.Cancel:
            return False

        with ExcelInterface(filename=filename) as excel:
            results_col = excel.find_results_col(row)
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {row}.\nEnsure col headed with results or measured",
                )
                return False

            setting = excel.get_tb_test_settings(row=row)

            self.uut.reset()

            with ExcelInterface(filename=filename) as excel:
                excel.row = row
                self.uut.set_channel(chan=1, enabled=True)

                self.uut.set_voltage_scale(chan=1, scale=0.5)
                self.uut.set_voltage_offset(chan=1, offset=0)

                self.uut.set_acquisition(32)

                self.ks33250.set_pulse(
                    period=DELAY_PERIOD, pulse_width=200e-6, amplitude=1
                )
                self.ks33250.enable_output(True)

                self.uut.set_trigger_level(chan=1, level=0)

                if setting.timebase:
                    self.uut.set_timebase(setting.timebase / 1e9)
                else:
                    self.uut.set_timebase(10e-9)

                if self.uut.keysight:
                    time.sleep(0.1)
                    self.uut.cursors_on()
                    time.sleep(1.5)

                    ref_x = self.uut.read_cursor("X1")  # get the reference time
                    ref = self.uut.read_cursor(
                        "Y1"
                    )  # get the voltage, so delayed can be adjusted to same
                else:
                    QMessageBox.information(
                        self,
                        "Instructions",
                        "Adjust Horz position so waveform is on center graticule",
                    )

                self.uut.set_timebase_pos(DELAY_PERIOD)  # delay 1ms to next pulse

                if not self.uut.keysight:
                    valid = False
                    while not valid:
                        result = QInputDialog.getDouble(
                            self,
                            "Difference",
                            "Enter difference in div of waveform crossing from center?",
                        )

                        try:
                            val = float(result)  # type: ignore
                            valid = True
                        except ValueError:
                            valid = False

                    excel.write_result(result=val, col=results_col)  # type: ignore
                else:
                    # Keysight
                    self.uut.set_cursor_position(
                        cursor="X1", pos=DELAY_PERIOD
                    )  # 1 ms delay
                    time.sleep(1)

                    self.uut.adjust_cursor(
                        target=ref  # type: ignore
                    )  # adjust the cursor until voltage is the same as measured from the reference pulse

                    offset_x = self.uut.read_cursor("X1")

                    error = ref_x - offset_x + 0.001  # type: ignore
                    print(f"TB Error {error}")

                    excel.row = row

                    if self.uut.family != DSOX_FAMILY.DSO5000:  # type: ignore
                        code = QInputDialog.getText(
                            self,
                            "Date code",
                            "Enter date code from serial label (0 if no code)",
                        )

                        age = 10

                        try:
                            val = int(code)  # type: ignore
                            print(f"{val/100}, {datetime.now().year-2000}")
                            if val // 100 > datetime.now().year - 2000:
                                val = 0
                            age = datetime.now().year - (val / 100) - 2000

                        except ValueError:
                            val = 0

                        if not val and len(self.uut.serial) >= 10:
                            code = self.uut.serial[2:6]

                            try:
                                val = int(code)
                                # start is from 1960
                                age = datetime.now().year - (val - 4000) / 100 - 2000
                            except ValueError:
                                # Invalid. Just assume 10
                                age = 10

                        age_years = int(age + 0.5)
                        excel.write_result(age_years, save=False, col=1)

                    # results in ppm
                    ppm = error / 1e-3 * 1e6

                    excel.write_result(ppm, save=True, col=results_col)

                self.update_test_progress()

        self.ks33250.enable_output(False)
        self.ks33250.close()
        self.uut.reset()
        self.uut.close()

        return True

    def test_trigger_sensitivity(self, filename: str, test_rows: List) -> bool:
        # sourcery skip: low-code-quality
        """
        test_trigger_sensitivity

        Do a simpler test of sensitivity than the Keysight method

        No other manufacturer tests sensitivity, and it is likely a design time issue rather than
        degradation problem

        Args:
            filename (str): _description_
            test_rows (list): _description_
        """

        QMessageBox.information(
            self, "Not IMplemented", "Not yet debugged, test manually"
        )

        return True

        self.current_test.emit("Testing: Trigger sensitivity")

        connections = self.test_connections(
            check_3458=False
        )  # Don't need 3458 for this test

        # require RF gen

        if not connections["RFGEN"]:
            QMessageBox.critical(
                self,
                "Error",
                "Cannot find RF Signal Generator",
            )
            return False

        self.uut.reset()

        # Turn off all channels but 1
        for chan in range(self.uut.num_channels):
            self.uut.set_channel(chan=chan + 1, enabled=chan == 0)

        # Need to know if the self.uut has 50 Ohm input or not

        ext_termination = True

        with ExcelInterface(filename=filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False

            for row in test_rows:
                excel.row = row
                settings = excel.get_trigger_settings()
                if settings.channel == 1 and settings.impedance == 50:
                    ext_termination = False
                    break

            # now the main test loop

            last_channel = 0

            for row in test_rows:
                excel.row = row

                settings = excel.get_trigger_settings()

                feedthru_msg = (
                    "via 50 Ohm Feedthru"
                    if ext_termination or str(settings.channel).upper() == "EXT"
                    else ""
                )

                if settings.channel != last_channel:
                    response = sg.popup_ok_cancel(
                        f"Connect signal generator output to channel {settings.channel} {feedthru_msg}",
                    )
                    if response == "Cancel":
                        return False

                    last_channel = settings.channel

                self.mxg.set_frequency_MHz(settings.frequency)
                self.mxg.set_level(settings.voltage, units="mV")
                self.mxg.set_output_state(True)

                if str(settings.channel).upper() != "EXT":
                    for chan in range(1, self.uut.num_channels + 1):
                        self.uut.set_channel(
                            chan=chan, enabled=chan == settings.channel
                        )
                    self.uut.set_channel(chan=int(settings.channel), enabled=True)
                    self.uut.set_voltage_scale(chan=int(settings.channel), scale=0.5)
                    self.uut.set_voltage_offset(chan=int(settings.channel), offset=0)
                    self.uut.set_trigger_level(chan=int(settings.channel), level=0)

                else:
                    # external. use channel 1
                    self.uut.set_channel(chan=1, enabled=True)
                    self.uut.set_trigger_level(chan=0, level=0)

                period = 1 / settings.frequency / 1e6
                # Round it off to a nice value of 1, 2, 5 or multiple

                period = self.round_range(period)

                self.uut.set_timebase(period * 2)

                triggered = self.uut.check_triggered(
                    sweep_time=0.1
                )  # actual sweep time is ns

                test_result = "Pass" if triggered else "Fail"
                excel.write_result(result=test_result, save=True, col=results_col)
                self.update_test_progress()

        self.mxg.set_output_state(False)
        self.mxg.close()

        self.uut.reset()
        self.uut.close()

        return True

    def test_risetime(self, filename: str, test_rows: List) -> bool:
        """
        test_risetime
        Use fast pulse generator to test rise time of each channel

        Args:
            filename (str): _description_
            test_rows (List): _description_
        """

        self.current_test.emit("Testing: Rise time")

        # only pulse gen required

        self.uut.open_connection()
        self.uut.reset()

        with ExcelInterface(filename=filename) as excel:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                )
                return False

            for row in test_rows:
                if self.abort_test:
                    return False

                excel.row = row

                settings = excel.get_tb_test_settings()

                if settings.channel > self.uut.num_channels:
                    continue

                message = f"Connect fast pulse generator to channel {settings.channel}"

                if settings.impedance != 50:
                    message += " via 50 Ohm feedthru"

                response = QMessageBox.information(
                    self,
                    "Connections",
                    message,
                    buttons=QMessageBox.StandardButton.Ok
                    | QMessageBox.StandardButton.Cancel,
                )

                if response == QMessageBox.StandardButton.Cancel:
                    return False

                for chan in range(self.uut.num_channels):
                    self.uut.set_channel(
                        chan=chan + 1, enabled=settings.channel == chan + 1
                    )

                self.uut.set_voltage_scale(chan=settings.channel, scale=0.2)

                if settings.impedance == 50:
                    self.uut.set_channel_impedance(
                        chan=settings.channel, impedance="50"
                    )

                if settings.bandwidth:
                    self.uut.set_channel_bw_limit(
                        chan=settings.channel, bw_limit=settings.bandwidth
                    )

                self.uut.set_timebase(settings.timebase * 1e-9)
                self.uut.set_trigger_level(chan=settings.channel, level=0)

                risetime = (
                    self.uut.measure_risetime(chan=settings.channel, num_readings=1)
                    * 1e9
                )

                # save in ns

                excel.write_result(risetime, save=True, col=results_col)
                self.update_test_progress()

        self.uut.reset()
        self.uut.close()

        return True

    def round_range(self, val: float) -> float:
        """
        round_range
        Round the setting (voltage or time) to the closest 1, 2, 5 multiple

        Args:
            val (float): _description_

        Returns:
            float: _description_
        """

        # Using logs to get the multiplier

        lg = math.log10(val)
        decade = int(lg)

        # for less than 1 we want to multiply
        if val < 1:
            decade -= 1

        # normalize. Anything greater than 1 will work fine here, but less than 1 need to multiply, hence the -decade
        normalized = val * math.pow(10, -decade)
        first_digit = int(str(normalized)[0])

        if first_digit < 2:
            first_digit = 1
        elif first_digit < 5:
            first_digit = 2
        else:
            first_digit = 5

        return first_digit * math.pow(10, decade)
