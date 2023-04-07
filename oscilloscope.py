"""
Test Oscilloscopes
# DK Jan 23
"""

from typing import Dict, List, Tuple
import PySimpleGUI as sg

from drivers.fluke_5700a import Fluke5700A
from drivers.Ks33250A import Ks33250A
from drivers.meatest_m142 import M142
from drivers.keysight_scope import Keysight_Oscilloscope
from drivers.tek_scope import Tektronix_Oscilloscope
from drivers.excel_interface import ExcelInterface
from drivers.rf_signal_generator import RF_Signal_Generator
from drivers.scpi_id import SCPI_ID
from drivers.Ks3458A import Ks3458A, Ks3458A_Function
import os
import sys
import time
from pathlib import Path
from datetime import datetime
import math
from pprint import pprint, pformat
from zipfile import BadZipFile


VERSION = "A.01.00"


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


def get_path(filename: str) -> str:
    """
    get_path
    The location of the diagram is different if run through interpreter or compiled.

    Args:
        filename ([type]): [description]

    Returns:
        [type]: [description]
    """

    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS, filename).__str__()  # type: ignore
    else:
        return filename


def led_indicator(key: str | None = None, radius: float = 30) -> sg.Graph:
    """
    led_indicator _summary_

    Args:
        key (str | None, optional): _description_. Defaults to None.
        radius (float, optional): _description_. Defaults to 30.

    Returns:
        sg.Graph: _description_
    """
    return sg.Graph(
        canvas_size=(radius, radius),
        graph_bottom_left=(-radius, -radius),
        graph_top_right=(radius, radius),
        pad=(0, 0),
        key=key,
    )


def set_led(window: sg.Window, key: str, color: str) -> None:
    """
    set_led _summary_

    Args:
        window (_type_): _description_
        key (_type_): _description_
        color (_type_): _description_
    """
    graph = window[key]
    graph.erase()  # type: ignore
    graph.draw_circle((0, 0), 12, fill_color=color, line_color=color)  # type: ignore


def connections_check_form(check_3458: bool) -> None:
    """
    connections_check_form _summary_
    """

    layout = [
        [sg.Text("Checking instruments.....", key="-CHECK_MSG-", text_color="Red")],
        [sg.Text("Calibrator", size=(20, 1)), led_indicator("-FLUKE_5700A_CONN-")],
        [sg.Text("33250A", size=(20, 1)), led_indicator("-33250_CONN-")],
        [sg.Text("RF Generator", size=(20, 1)), led_indicator("-RFGEN_CONN-")],
        [
            sg.Text(
                "3458A", size=(20, 1), text_color="white" if check_3458 else "grey"
            ),
            led_indicator("-3458-"),
        ],
        [sg.Text("UUT", size=(20, 1)), led_indicator("-UUT-")],
        [sg.Text()],
        [sg.Ok(size=(14, 1)), sg.Button("Try Again", size=(14, 1))],
    ]

    window = sg.Window(
        "Oscilloscope Test",
        layout,
        finalize=True,
        icon=get_path("ui\\scope.ico"),
    )

    connected = test_connections(check_3458)
    window["-CHECK_MSG-"].update(visible=False)

    set_led(
        window,
        "-FLUKE_5700A_CONN-",
        color="green" if connected["FLUKE_5700A"] else "red",
    )
    set_led(
        window,
        "-33250_CONN-",
        color="green" if connected["33250A"] else "red",
    )
    set_led(
        window,
        "-RFGEN_CONN-",
        color="green" if connected["RFGEN"] else "red",
    )
    set_led(
        window,
        "-UUT-",
        color="green" if connected["DSO"] else "red",
    )
    if check_3458:
        set_led(window, "-3458-", color="green" if connected["3458"] else "red")
    else:
        set_led(window, "-3458-", color="grey")

    while True:
        event, values = window.read()  # type: ignore

        if event in ["Ok", sg.WIN_CLOSED]:
            break

        if event == "Try Again":
            window["-CHECK_MSG-"].update(visible=True)
            connected = test_connections(check_3458)
            window["-CHECK_MSG-"].update(visible=False)

            set_led(
                window,
                "-FLUKE_5700A_CONN-",
                color="green" if connected["FLUKE_5700A"] else "red",
            )
            set_led(
                window,
                "-33250_CONN-",
                color="green" if connected["33250A"] else "red",
            )
            set_led(
                window,
                "-RFGEN_CONN-",
                color="green" if connected["RFGEN"] else "red",
            )
            set_led(
                window,
                "-UUT-",
                color="green" if connected["DSO"] else "red",
            )

    window.close()


def test_connections(check_3458: bool) -> Dict:
    """
    Make sure all of the instruments are connected
    """

    global calibrator
    global ks33250
    global mxg
    global uut

    fluke_5700a_conn = calibrator.is_connected()
    ks33250_conn = ks33250.is_connected()
    rfgen_conn = mxg.is_connected()
    uut_conn = uut.is_connected()
    ks3458_conn = ks3458.is_connected() if check_3458 else False

    return {
        "FLUKE_5700A": fluke_5700a_conn,
        "33250A": ks33250_conn,
        "RFGEN": rfgen_conn,
        "DSO": uut_conn,
        "3458": ks3458_conn,
    }


def update_test_progress() -> None:
    """
    update_test_progress
    Test step complted, increment progress
    """

    global test_number
    global test_progress

    test_number += 1
    test_progress.update(test_number)


def run_tests(filename: str, test_rows: List, parallel_channels: bool = False) -> None:
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

    test_number = 0

    with ExcelInterface(filename=filename) as excel:
        excel.backup()

        # first update the model and serial

        uut.open_connection()

        # If the named range doesn't exist, nothing is written
        excel.write_data(data=uut.model, named_range="Model")
        excel.write_data(data=uut.serial, named_range="Serial")

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

        for test_name in ordered_test_names:
            testing_rows = excel.get_test_rows(test_name)
            # At the moment we only do full tests, so we can get the test rows form the excel sheet

            # TODO use functional method

            if "DCV" in test_name:
                if not test_dcv(
                    filename=filename,
                    test_rows=testing_rows,
                    parallel_channels=parallel_channels,
                ):
                    break

            elif test_name == "POS":
                if not test_position(
                    filename=filename,
                    test_rows=testing_rows,
                    parallel_channels=parallel_channels,
                ):
                    break

            elif test_name == "BAL":
                if not test_dc_balance(filename=filename, test_rows=testing_rows):
                    break

            elif test_name == "CURS":
                if not test_cursor(filename=filename, test_rows=testing_rows):
                    break

            elif test_name == "RISE":
                if not test_risetime(filename=filename, test_rows=testing_rows):
                    break

            elif test_name == "TIME":
                if not test_timebase(filename=filename, row=testing_rows[0]):
                    break

            elif test_name == "TRIG":
                if not test_trigger_sensitivity(
                    filename=filename, test_rows=testing_rows
                ):
                    break

            elif test_name == "IMP":
                if not test_impedance(filename=filename, test_rows=testing_rows):
                    break

            elif test_name == "NOISE":
                if not test_random_noise(filename=filename, test_rows=test_rows):
                    break

            elif test_name == "DELTAT":
                if not test_delta_time(filename=filename, test_rows=test_rows):
                    break


def test_dc_balance(filename: str, test_rows: List) -> bool:
    """
    test_dc_balance
    Test the dc balance of each channel with no signal applied

    Args:
        filename (str): _description_
        test_rows (int): _description_
    """

    global uut
    global current_test_text

    # no equipment required

    current_test_text.update("Testing: DCV Balance")

    response = sg.popup_ok_cancel(
        "Remove inputs from all channels",
        background_color="blue",
        icon=get_path("ui\\scope.ico"),  # type: ignore
    )

    if response == "Cancel":
        return False

    uut.reset()

    uut.set_acquisition(32)

    uut.set_timebase(0.001)

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )
            return False

        for row in test_rows:
            excel.row = row

            settings = excel.get_volt_settings()

            if settings.function == "BAL":
                uut.set_channel(chan=int(settings.channel), enabled=True, only=True)
                uut.set_voltage_scale(chan=int(settings.channel), scale=settings.scale)
                uut.set_voltage_offset(chan=int(settings.channel), offset=0)
                uut.set_channel_coupling(
                    chan=int(settings.channel), coupling=settings.coupling
                )

                reading = (
                    uut.measure_voltage(chan=int(settings.channel), delay=2) * 1000
                )  # mV

                excel.write_result(reading, col=results_col)
                update_test_progress()

    uut.reset()

    return True


def test_delta_time(filename: str, test_rows: List) -> bool:
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

    global uut
    global current_test_text

    current_test_text.update("Testing: Delta Time")

    connections = test_connections(check_3458=False)  # Always required

    if not connections["RFGEN"]:
        sg.popup_error(
            "Cannot find RF Generator",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )
        return False

    if not connections["33250A"]:
        sg.popup_error(
            "Cannot find 33250A",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )
        return False

    uut.open_connection()
    uut.reset()

    # For the moment, this is a Tek MSO5000  special test, so commands written directly here. If any more
    # are required, put into driver

    uut.set_acquisition(16)

    last_channel = -1

    with ExcelInterface(filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )
            return False
        for row in test_rows:
            excel.row = row

            units = excel.get_units()

            settings = excel.get_sample_rate_settings()

            if settings.channel != last_channel:
                sg.popup(
                    f"Connect RF Sig gen to Channel {settings.channel}",
                    background_color="blue",
                )
                last_channel = settings.channel

            uut.set_channel(chan=settings.channel, enabled=True, only=True)
            uut.set_voltage_scale(chan=settings.channel, scale=settings.scale)
            uut.set_channel_coupling(chan=settings.channel, coupling=settings.coupling)
            uut.set_channel_impedance(chan=settings.channel, impedance="50")
            uut.set_timebase(settings.timebase)

            uut.write(f"HORIZONTAL:MODE:SAMPLERATE {settings.sample_rate}")
            # uut.write("HORIZONTAL:MODE:RECORDLENGTH ")

            mxg.set_frequency(settings.frequency)
            mxg.set_level(settings.voltage, units="VPP")
            mxg.set_output_state(True)

            uut.write("MEASU:MEAS1:BURST")

            uut.write("MEASURE:STATISTICS:MODE MEANSTDDEV")
            uut.write("MEASURE:STATISTICS:WEIGHTING 1000")
            uut.write("MEASUREMENT:STATISTICS:COUNT RESET")

            time.sleep(10)

            result = uut.query("MEASU:MEAS1:VAL?")

            if units[0] == "p":
                result *= 1_000_000_000_000
            elif units[0] == "n":
                result *= 1_000_000_000

            excel.write_result(result=result, col=results_col)

            mxg.set_output_state(False)

    return True


def test_random_noise(filename: str, test_rows: List) -> bool:
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

    global uut
    global current_test_text

    current_test_text.update("Testing: Random noise sample acquisition")

    # No equipment required
    uut.open_connection()
    uut.reset()

    uut.set_acquisition(16)

    with ExcelInterface(filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )
            return False
        for row in test_rows:
            excel.row = row

            settings = excel.get_volt_settings()
            # Only need the channel

            uut.set_channel(chan=settings.channel, enabled=True, only=True)  # type: ignore
            uut.set_channel_impedance(
                chan=settings.channel, impedance=settings.impedance  # type: ignore
            )
            uut.set_channel_bw_limit(chan=settings.channel, bw_limit=settings.bandwidth)  # type: ignore

            rnd = uut.measure_rms_noise(chan=settings.channel)  # type: ignore

            uut.measure_clear()
            avg = uut.measure_voltage(chan=settings.channel)  # type: ignore

            result = rnd - avg

            excel.write_result(result=result, col=results_col, save=False)

        excel.save_sheet()

    return True


def test_impedance(filename: str, test_rows: List) -> bool:
    """
    test_impedance
    Test the input impedance of the channels

    Args:
        filename (str): _description_
        test_rows (List): _description_

    Returns:
        bool: _description_
    """

    global uut
    global current_test_text
    global ks3458

    current_test_text.update("Testing: Input Impedance")

    connections = test_connections(check_3458=True)  # Always required

    if not connections["3458"]:
        sg.popup_error(
            "Cannot find 3458A", background_color="blue", icon=get_path("ui\\scope.ico")
        )
        return False

    uut.open_connection()
    uut.reset()

    ks3458.open_connection()
    ks3458.reset()

    last_channel = -1

    # Turn off all channels but 1
    for chan in range(uut.num_channels):
        uut.set_channel(chan=chan + 1, enabled=chan == 0)

    uut.set_acquisition(1)

    with ExcelInterface(filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )
            return False
        for row in test_rows:
            excel.row = row

            settings = excel.get_volt_settings()

            channel = int(settings.channel)
            units = excel.get_units()

            if channel > uut.num_channels:
                continue

            if channel != last_channel:
                sg.popup(
                    f"Connect 3458A Input to UUT Ch {channel}", background_color="blue"
                )
                if last_channel > 0:
                    # changed channel to another, but not channel 1. reset all of the settings on the channel just measured
                    uut.set_voltage_scale(chan=last_channel, scale=1)
                    uut.set_voltage_offset(chan=last_channel, offset=0)
                    uut.set_channel(chan=last_channel, enabled=False)
                    uut.set_channel_bw_limit(chan=last_channel, bw_limit=False)
                    uut.set_channel(chan=channel, enabled=True)
                    uut.set_channel_impedance(
                        chan=last_channel, impedance="1M"
                    )  # always
                last_channel = channel

            uut.set_voltage_scale(chan=channel, scale=settings.scale)
            uut.set_voltage_offset(chan=channel, offset=settings.offset)
            uut.set_channel_impedance(chan=channel, impedance=settings.impedance)
            uut.set_channel_bw_limit(chan=channel, bw_limit=settings.bandwidth)

            time.sleep(0.5)

            reading = ks3458.measure(function=Ks3458A_Function.R2W)["Average"]  # type: ignore
            if units.lower().startswith("k"):
                reading /= 1000
            if units.upper().startswith("M"):
                reading /= 1_000_000

            excel.write_result(reading, col=results_col)

            update_test_progress()

        # Turn off all channels but 1
        for chan in range(uut.num_channels):
            uut.set_channel(chan=chan + 1, enabled=chan == 0)
            uut.set_channel_bw_limit(chan=chan, bw_limit=False)

        uut.reset()
        uut.close()

    return True


def test_dcv(filename: str, test_rows: List, parallel_channels: bool = False) -> bool:
    # sourcery skip: extract-method, low-code-quality
    """
    test_dcv
    Perform the basic DC V tests
    Set the calibrator to the voltage, allow the scope to stabilizee, then read the cursors or measurement values
    """

    global calibrator
    global uut
    global simulating
    global cursor_results
    global current_test_text

    current_test_text.update("Testing: DC Voltage")

    last_channel = -1

    connections = test_connections(check_3458=False)  # Don't need 3458 for this test

    # require calibrator

    if not connections["FLUKE_5700A"]:
        sg.popup_error(
            "Cannot find calibrator",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )
        return False

    uut.open_connection()
    uut.reset()

    uut.set_timebase(1e-3)

    cursor_results = []  # save results for cursor tests

    if parallel_channels:
        response = sg.popup_ok_cancel(
            "Connect calibrator output to all channels in parallel",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),  # type: ignore
        )

        if response == "Cancel":
            return False

    # Turn off all channels but 1
    for chan in range(uut.num_channels):
        uut.set_channel(chan=chan + 1, enabled=chan == 0)

    uut.set_acquisition(32)

    with ExcelInterface(filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )
            return False
        for row in test_rows:
            excel.row = row

            settings = excel.get_volt_settings()

            units = excel.get_units()

            calibrator.set_voltage_dc(0)

            channel = int(settings.channel)

            if channel > uut.num_channels:
                continue

            if channel != last_channel:
                if last_channel > 0:
                    # changed channel to another, but not channel 1. reset all of the settings on the channel just measured
                    uut.set_voltage_scale(chan=last_channel, scale=1)
                    uut.set_voltage_offset(chan=last_channel, offset=0)
                    uut.set_channel(chan=last_channel, enabled=False)
                    uut.set_channel_bw_limit(chan=last_channel, bw_limit=False)
                    uut.set_channel(chan=channel, enabled=True)
                    uut.set_channel_impedance(
                        chan=last_channel, impedance="1M"
                    )  # always

                uut.set_voltage_scale(chan=channel, scale=5)
                uut.set_voltage_offset(chan=channel, offset=0)

                uut.set_cursor_xy_source(chan=1, cursor=1)
                uut.set_cursor_position(cursor="X1", pos=0)
                if not parallel_channels:
                    response = sg.popup_ok_cancel(
                        f"Connect calibrator output to channel {channel}",
                        background_color="blue",
                        icon=get_path("ui\\scope.ico"),  # type: ignore
                    )
                    if response == "Cancel":
                        return False
                last_channel = channel

            uut.set_channel(chan=channel, enabled=True)
            uut.set_voltage_scale(chan=channel, scale=settings.scale)
            uut.set_voltage_offset(chan=channel, offset=settings.offset)

            if settings.impedance:
                uut.set_channel_impedance(chan=channel, impedance=settings.impedance)

            if settings.bandwidth:
                uut.set_channel_bw_limit(chan=channel, bw_limit=settings.bandwidth)
            else:
                uut.set_channel_bw_limit(chan=channel, bw_limit=False)

            if settings.invert:
                # already casted to a bool
                uut.set_channel_invert(chan=channel, inverted=settings.invert)
            else:
                uut.set_channel_invert(chan=channel, inverted=False)

            if settings.function == "DCV-BAL":
                # Non keysight, apply the half the voltage and the offset then do the reverse

                calibrator.set_voltage_dc(settings.voltage)

            # 0V test
            calibrator.operate()

            uut.set_acquisition(1)

            if not simulating:
                time.sleep(0.2)

            uut.set_acquisition(32)

            if not simulating:
                time.sleep(1)

            if uut.keysight:
                voltage1 = uut.read_cursor_avg()

            uut.measure_clear()
            reading1 = uut.measure_voltage(chan=channel)

            if settings.function == "DCV-BAL":
                # still set up for the + voltage

                calibrator.set_voltage_dc(-settings.voltage)
                uut.set_voltage_offset(chan=channel, offset=-settings.offset)
            else:
                calibrator.set_voltage_dc(settings.voltage)

            calibrator.operate()

            uut.set_acquisition(1)

            if not simulating:
                time.sleep(0.2)

            uut.set_acquisition(32)

            if not simulating:
                time.sleep(1)

            uut.measure_clear()

            reading = uut.measure_voltage(chan=channel)

            if uut.keysight:
                voltage2 = uut.read_cursor_avg()

            if uut.keysight:
                cursor_results.append(
                    {
                        "chan": channel,
                        "scale": settings.scale,
                        "result": voltage2 - voltage1,  # type: ignore
                    }
                )

            calibrator.standby()

            if units.startswith("m"):
                reading *= 1000
                reading1 *= 1000

            if settings.function == "DCV-BAL":
                diff = reading1 - reading
                excel.write_result(diff, col=results_col)  # auto saving
            else:
                # Keysight simple test. 0V is measured for the cursors only
                excel.write_result(reading, col=results_col)

            update_test_progress()

        calibrator.reset()
        calibrator.close()

        # Turn off all channels but 1
        for chan in range(uut.num_channels):
            uut.set_channel(chan=chan + 1, enabled=chan == 0)
            uut.set_channel_bw_limit(chan=chan, bw_limit=False)

        uut.reset()
        uut.close()

    return True


def test_cursor(filename: str, test_rows: List) -> bool:
    """
    test_cursor
    Dual cursor test. Measure voltage with no voltage applied, apply voltage, measure again, record the difference
    Measurements are taken during the DCV test, and recalled here

    Args:
        filename (str): _description_
        test_rows (List): _description_
    """
    global simulating
    global cursor_results
    global current_test_text

    current_test_text.update("Testing: Cursor position")

    # no equipment as using buffered results

    with ExcelInterface(filename) as excel:
        for row in test_rows:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                sg.popup_error(
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                    background_color="blue",
                    icon=get_path("ui\\scope.ico"),
                )
                return False
            excel.row = row

            settings = excel.get_volt_settings()

            if len(cursor_results):
                for res in cursor_results:
                    if (
                        res["chan"] == settings.channel
                        and res["scale"] == settings.scale
                    ):
                        units = excel.get_units()
                        result = res["result"]
                        if units.startswith("m"):
                            result *= 1000
                        excel.write_result(result, save=False, col=results_col)
                        update_test_progress()
                        break

        excel.save_sheet()

    return True


def test_position(
    filename: str, test_rows: List, parallel_channels: bool = False
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

    global calibrator
    global uut
    global current_test_text

    current_test_text.update("Testing: DC Position")

    connections = test_connections(check_3458=False)  # Don't need 3458 for this test

    # require calibrator

    if not connections["FLUKE_5700A"]:
        sg.popup_error(
            "Cannot find calibrator",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )
        return False

    uut.reset()

    uut.set_acquisition(32)

    last_channel = -1

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )
            return False

        for row in test_rows:
            excel.row = row

            settings = excel.get_volt_settings()

            if settings.channel != last_channel and not parallel_channels:
                response = sg.popup_ok_cancel(
                    f"Connect calibrator output to channel {settings.channel}",
                    background_color="blue",
                    icon=get_path("ui\\scope.ico"),  # type: ignore
                )
                if response == "Cancel":
                    return False

                last_channel = settings.channel

            uut.set_channel(chan=int(settings.channel), enabled=True, only=True)
            uut.set_channel_bw_limit(chan=int(settings.channel), bw_limit=True)
            uut.set_voltage_scale(chan=int(settings.channel), scale=settings.scale)
            pos = -4 if settings.offset > 0 else 4
            uut.set_voltage_position(
                chan=int(settings.channel), position=pos
            )  # divisions
            uut.set_voltage_offset(chan=int(settings.channel), offset=settings.offset)

            uut.set_acquisition(1)  # Too slow to adjust otherwise
            calibrator.set_voltage_dc(settings.voltage)

            calibrator.operate()

            uut.set_acquisition(32)

            # uut.measure_voltage_clear()

            # reading = uut.measure_voltage(chan=int(settings.channel), delay=2)

            response = sg.popup_yes_no(
                "Trace within 0.2 div of center?",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )

            result = "Pass" if response == "Yes" else "Fail"

            calibrator.standby()

            excel.write_result(result=result, col=results_col)
            update_test_progress()

    calibrator.reset()
    calibrator.close()

    uut.reset()
    uut.close()

    return True


# DELAY_PERIOD = 0.00099998  # 1 ms
# DELAY_PERIOD = 0.00100002  # 1 ms
DELAY_PERIOD = 0.001  # 1 ms


def test_timebase(filename: str, row: int) -> bool:
    # sourcery skip: low-code-quality
    """
    test_timebase
    Test the timebase. Simple single row test

    Args:
        row (int): _description_
    """

    global current_test_text

    current_test_text.update("Testing: Timebase")

    connections = test_connections(check_3458=False)  # Don't need 3458 for this test

    # require RF gen

    if not connections["33250A"]:
        sg.popup_error(
            "Cannot find 33250A Generator",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )
        return False

    response = sg.popup_ok_cancel(
        "Connect 33250A output to Ch1",
        background_color="blue",
        icon=get_path("ui\\scope.ico"),  # type: ignore
    )

    if response == "Cancel":
        return False

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(row)
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )
            return False

        setting = excel.get_tb_test_settings(row=row)

        uut.reset()

        with ExcelInterface(filename=filename) as excel:
            excel.row = row
            uut.set_channel(chan=1, enabled=True)

            uut.set_voltage_scale(chan=1, scale=0.5)
            uut.set_voltage_offset(chan=1, offset=0)

            uut.set_acquisition(32)

            ks33250.set_pulse(period=DELAY_PERIOD, pulse_width=200e-6, amplitude=1)
            ks33250.enable_output(True)

            uut.set_trigger_level(chan=1, level=0)

            if setting.timebase:
                uut.set_timebase(setting.timebase / 1e9)
            else:
                uut.set_timebase(10e-9)

            if uut.keysight:
                time.sleep(0.1)
                uut.cursors_on()
                time.sleep(1.5)

                ref_x = uut.read_cursor("X1")  # get the reference time
                ref = uut.read_cursor(
                    "Y1"
                )  # get the voltage, so delayed can be adjusted to same
            else:
                sg.popup(
                    "Adjust Horz position so waveform is on center graticule",
                    background_color="blue",
                    icon=get_path("ui\\scope.ico"),
                )

            uut.set_timebase_pos(DELAY_PERIOD)  # delay 1ms to next pulse

            if not uut.keysight:
                valid = False
                while not valid:
                    result = sg.popup_get_text(
                        "Enter difference in div of waveform crossing from center?",
                        background_color="blue",
                        icon=get_path("ui\\scope.ico"),
                    )
                    try:
                        val = float(result)  # type: ignore
                        valid = True
                    except ValueError:
                        valid = False

                excel.write_result(result=val, col=results_col)  # type: ignore
            else:
                # Keysight
                uut.set_cursor_position(cursor="X1", pos=DELAY_PERIOD)  # 1 ms delay
                time.sleep(1)

                uut.adjust_cursor(
                    target=ref  # type: ignore
                )  # adjust the cursor until voltage is the same as measured from the reference pulse

                offset_x = uut.read_cursor("X1")

                error = ref_x - offset_x + 0.001  # type: ignore
                print(f"TB Error {error}")

                code = sg.popup_get_text(
                    "Enter date code from serial label (0 if no code)",
                    background_color="blue",
                    icon=get_path("ui\\scope.ico"),
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

                if not val and len(uut.serial) >= 10:
                    code = uut.serial[2:6]

                    try:
                        val = int(code)
                        # start is from 1960
                        age = datetime.now().year - (val - 4000) / 100 - 2000
                    except ValueError:
                        # Invalid. Just assume 10
                        age = 10

                age_years = int(age + 0.5)

                # results in ppm
                ppm = error / 1e-3 * 1e6
                excel.row = row
                excel.write_result(ppm, save=False, col=results_col)
                excel.write_result(age_years, save=True, col=1)
            update_test_progress()

    ks33250.enable_output(False)
    ks33250.close()
    uut.reset()
    uut.close()

    return True


def test_trigger_sensitivity(filename: str, test_rows: List) -> bool:
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

    global mxg
    global uut
    global current_test_text

    sg.popup(
        "Not yet debugged, test manually",
        background_color="blue",
        icon=get_path("ui\\scope.ico"),
    )

    return True

    current_test_text.update("Testing: Trigger sensitivity")

    connections = test_connections(check_3458=False)  # Don't need 3458 for this test

    # require RF gen

    if not connections["RFGEN"]:
        sg.popup_error(
            "Cannot find RF Signal Generator",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )
        return False

    uut.reset()

    # Turn off all channels but 1
    for chan in range(uut.num_channels):
        uut.set_channel(chan=chan + 1, enabled=chan == 0)

    # Need to know if the UUT has 50 Ohm input or not

    ext_termination = True

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
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
                    background_color="blue",
                    icon=get_path("ui\\scope.ico"),
                )
                if response == "Cancel":
                    return False

                last_channel = settings.channel

            mxg.set_frequency_MHz(settings.frequency)
            mxg.set_level(settings.voltage, units="mV")
            mxg.set_output_state(True)

            if str(settings.channel).upper() != "EXT":
                for chan in range(1, uut.num_channels + 1):
                    uut.set_channel(chan=chan, enabled=chan == settings.channel)
                uut.set_channel(chan=int(settings.channel), enabled=True)
                uut.set_voltage_scale(chan=int(settings.channel), scale=0.5)
                uut.set_voltage_offset(chan=int(settings.channel), offset=0)
                uut.set_trigger_level(chan=int(settings.channel), level=0)

            else:
                # external. use channel 1
                uut.set_channel(chan=1, enabled=True)
                uut.set_trigger_level(chan=0, level=0)

            period = 1 / settings.frequency / 1e6
            # Round it off to a nice value of 1, 2, 5 or multiple

            period = round_range(period)

            uut.set_timebase(period * 2)

            triggered = uut.check_triggered(sweep_time=0.1)  # actual sweep time is ns

            test_result = "Pass" if triggered else "Fail"
            excel.write_result(result=test_result, save=True, col=results_col)
            update_test_progress()

    mxg.set_output_state(False)
    mxg.close()

    uut.reset()
    uut.close()

    return True


def test_risetime(filename: str, test_rows: List) -> bool:
    """
    test_risetime
    Use fast pulse generator to test rise time of each channel

    Args:
        filename (str): _description_
        test_rows (List): _description_
    """

    global current_test_text

    current_test_text.update("Testing: Rise time")

    # only pulse gen required

    uut.open_connection()
    uut.reset()

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )
            return False

        for row in test_rows:
            excel.row = row

            settings = excel.get_tb_test_settings()

            message = f"Connect fast pulse generator to channel {settings.channel}"

            if settings.impedance != 50:
                message += "via 50 Ohm feedthru"

            response = sg.popup_ok_cancel(
                message, background_color="blue", icon=get_path("ui\\scope.ico")  # type: ignore
            )
            if response == "Cancel":
                return False

            for chan in range(uut.num_channels):
                uut.set_channel(chan=chan + 1, enabled=settings.channel == chan + 1)

            uut.set_voltage_scale(chan=settings.channel, scale=0.2)

            if settings.impedance == 50:
                uut.set_channel_impedance(chan=settings.channel, impedance="50")

            if settings.bandwidth:
                uut.set_channel_bw_limit(
                    chan=settings.channel, bw_limit=settings.bandwidth
                )

            uut.set_timebase(settings.timebase * 1e-9)
            uut.set_trigger_level(chan=settings.channel, level=0)

            risetime = uut.measure_risetime(chan=settings.channel, num_readings=1) * 1e9

            # save in ns

            excel.write_result(risetime, save=True, col=results_col)
            update_test_progress()

    uut.reset()
    uut.close()

    return True


def round_range(val: float) -> float:
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


def individual_tests(filename: str) -> Tuple:
    """
    individual_tests
    show form for selecting individual tests

    Returns:
        List: _description_
    """

    with ExcelInterface(filename=filename) as excel:
        test_names = excel.get_test_types()

    layout = [
        [
            sg.Text(
                "Select tests to perform:", background_color="blue", text_color="white"
            )
        ],
        [
            [sg.Checkbox(name, key=name, background_color="blue", default=False)]
            for name in test_names
        ],
        [sg.Button("Test", size=(10, 1)), sg.Cancel(size=(10, 1))],
    ]

    window = sg.Window(
        "Test individual tests",
        layout=layout,
        finalize=True,
        background_color="blue",
        icon=get_path("ui\\scope.ico"),
    )

    event, values = window.read()  # type: ignore

    test_steps = []

    do_parallel = False

    if event not in [sg.WIN_CLOSED, "Cancel"]:
        # Now work out which are checked
        # There must be a pythonic way to get the list of values in one line, but couldn't work it out using lambdas and filters

        checked_list = [val for val in values if values[val]]

        print(checked_list)

        # we actually need the list of rows for all of the steps

        for row in checked_list:
            rows = excel.get_test_rows(row)
            test_steps.extend(iter(rows))
        print(test_steps)

        do_parallel = False
        if "DCV" in checked_list:
            do_parallel = True
        if "DCV-BAL" in checked_list:
            do_parallel = True
        if "CURS" in checked_list:
            do_parallel = True

    window.close()

    excel.close()

    return (sorted(test_steps), do_parallel)


def load_uut_driver(address: str, simulating: bool = False) -> bool:
    """
    load_uut_driver
    Use a generic driver to figure out which driver of the scope should be used
    """

    global uut

    if simulating:
        uut = Tektronix_Oscilloscope(simulate=simulating)
        uut.model = "MSO5104B"
        uut.open_connection()
        return True

    # TODO when using the SCPI ID class, it affects all subsequent uses

    with Keysight_Oscilloscope(simulate=False) as scpi_uut:
        scpi_uut.visa_address = address
        scpi_uut.open_connection()
        manufacturer = scpi_uut.get_manufacturer()
        num_channels = scpi_uut.get_number_channels()

    if manufacturer == "":
        sg.popup_error(
            "Unable to contact UUT. Is address correct?",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )
        return False
    elif manufacturer == "KEYSIGHT":
        uut = Keysight_Oscilloscope(simulate=False)

    elif manufacturer == "TEKTRONIX":
        uut = Tektronix_Oscilloscope(simulate=False)

    else:
        sg.popup_error(
            f"No driver for {manufacturer}. Using Tektronix driver",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )
        uut = Tektronix_Oscilloscope(simulate=False)

    uut.num_channels = num_channels

    return True


def select_visa_address() -> str:
    """
    select_visa_address _summary_

    Show a form with all of the found visa resource strings

    Returns:
        str: _description_
    """

    addresses = SCPI_ID.get_all_attached()

    visa_instruments = []

    for addr in addresses:
        with SCPI_ID(address=addr) as scpi:
            idn = scpi.get_id()[0]
            visa_instruments.append((addr, idn))

    radio_buttons = [
        [
            sg.Radio(
                addr,
                1,
                background_color="blue",
                default=bool(addr[0].startswith("USB")),
            )
        ]
        for addr in visa_instruments
    ]

    layout = [
        [sg.Text("Select item", background_color="blue")],
        [radio_buttons],
        [sg.Ok(size=(12, 1)), sg.Cancel(size=(12, 1))],
    ]

    window = sg.Window(
        "Addresses selection",
        layout=layout,
        icon=get_path("ui\\scope.ico"),
        background_color="blue",
    )

    event, values = window.read()  # type: ignore

    window.close()

    if event in ["Cancel", sg.WIN_CLOSED]:
        return ""

    return next((addr for index, addr in enumerate(addresses) if values[index]), "")


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


def results_sheet_check(filename: str) -> None:
    """
    results_sheet_check
    Do a check of the results sheet to make sure valid
    """

    with ExcelInterface(filename=filename) as excel:
        nr = excel.get_named_cell("StartCell")
        if not nr:
            sg.popup_error(
                "No cell named StartCell. Name the first cell with function data StartCell",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )

        valid_tests = excel.get_test_types()

        sg.popup(
            f"The following tests are found: {pformat( list(valid_tests))}",
            background_color="blue",
            icon=get_path("ui\\scope.ico"),
        )

        invalid_tests = excel.get_invalid_tests()

        if len(invalid_tests):
            sg.popup(
                f"The following rows will not be tested: {pformat(invalid_tests)}",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )


def check_impedance(filename: str) -> bool:
    """
    check_impedance
    Check the results sheet to see if there are any impedance tests.
    If not, no requirement to check 3458A

    Returns:
        bool: _description_
    """

    with ExcelInterface(filename=filename) as excel:
        nr = excel.get_named_cell("StartCell")
        if not nr:
            sg.popup_error(
                "No cell named StartCell. Name the first cell with function data StartCell",
                background_color="blue",
                icon=get_path("ui\\scope.ico"),
            )

        valid_tests = excel.get_test_types()

    return "IMP" in valid_tests


if __name__ == "__main__":
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
                default_value="6",
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
            sg.Text("RF Gen", size=(15, 1)),
            sg.Combo(
                gpib_ifc_list,
                size=(10, 1),
                key="GPIB_IFC_RFGEN",
                default_value=sg.user_settings_get_entry("-RFGEN_GPIB_IFC-"),
            ),
            sg.Combo(
                gpib_addresses,
                default_value=sg.user_settings_get_entry("-RFGEN_GPIB_ADDR-"),
                size=(6, 1),
                key="GPIB_ADDR_RFGEN",
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
        [sg.Text()],
        [
            sg.ProgressBar(
                max_value=100, size=(35, 10), visible=False, key="-PROGRESS-"
            ),
            sg.Text("Progress", visible=False, key="-PROG_TEXT-"),
        ],
        [sg.Text(key="-CURRENT_TEST-")],
        [sg.Text()],
        [
            sg.Button("Test Connections", size=(15, 1), key="-TEST_CONNECTIONS-"),
            sg.Button("Perform Tests", size=(12, 1), key="-INDIVIDUAL-"),
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
        sg.user_settings_set_entry("-RFGEN_GPIB_IFC-", values["GPIB_IFC_RFGEN"])
        sg.user_settings_set_entry("-RFGEN_GPIB_ADDR-", values["GPIB_ADDR_RFGEN"])
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

        mxg.simulating = simulating
        mxg_address = f"{values['GPIB_IFC_RFGEN']}::{values['GPIB_ADDR_RFGEN']}::INSTR"

        uut.simulating = simulating
        uut.visa_address = values["-UUT_ADDRESS-"]

        ks3458.simulating = simulating
        ks3458_address = f"{values['GPIB_IFC_3458']}::{values['GPIB_ADDR_3458']}::INSTR"

        if event in [
            "-INDIVIDUAL-",
        ]:
            # Common check to make sure everything is in order

            valid = True

            # reset the back colors instead of else statements after checking

            window["-FILE-"].update(background_color=back_color)

            if not values["-FILE-"]:
                window["-FILE-"].update(background_color="Red")
                valid = False

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

            window["-VIEW-"].update(disabled=True)  # Disable while testing

            if values["-CALIBRATOR-"] == "M-142":
                calibrator = M142(simulate=simulating)
            else:
                calibrator = Fluke5700A(simulate=simulating)
            calibrator.visa_address = calibrator_address
            ks33250.visa_address = ks33250_address
            mxg.visa_address = mxg_address
            ks3458.visa_address = ks3458_address

            if load_uut_driver(address=values["-UUT_ADDRESS-"], simulating=simulating):
                uut.visa_address = values["-UUT_ADDRESS-"]
                uut.open_connection()

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

        if event == "-TEST_CONNECTIONS-":
            if values["-CALIBRATOR-"] == "M-142":
                calibrator = M142(simulate=simulating)
            else:
                calibrator = Fluke5700A(simulate=simulating)
            calibrator.visa_address = calibrator_address
            ks33250.visa_address = ks33250_address
            mxg.visa_address = mxg_address
            ks3458.visa_address = ks3458_address

            # Check if the 3458A is needed

            impedance_tests = check_impedance(values["-FILE-"])

            window["GPIB_IFC_3458"].update(disabled=not impedance_tests)
            window["GPIB_ADDR_3458"].update(disabled=not impedance_tests)

            load_uut_driver(address=values["-UUT_ADDRESS-"], simulating=simulating)
            uut.visa_address = values["-UUT_ADDRESS-"]

            connections_check_form(impedance_tests)
            continue

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
