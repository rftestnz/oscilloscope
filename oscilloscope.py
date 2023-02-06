"""
Test DSOX Oscilloscopes
# DK Jan 23
"""

from typing import Dict, List
import PySimpleGUI as sg

from drivers.fluke_5700a import Fluke5700A
from drivers.Ks33250A import Ks33250A
from drivers.meatest_m142 import M142
from drivers.keysight_scope import Keysight_Oscilloscope
from drivers.tek_scope import Tektronix_Oscilloscope
from drivers.excel_interface import ExcelInterface
from drivers.rf_signal_generator import RF_Signal_Generator
from drivers.scpi_id import SCPI_ID
import os
import sys
import time
from pathlib import Path
from datetime import datetime
import math

VERSION = "A.00.00"


calibrator = Fluke5700A()
ks33250 = Ks33250A()
uut = Keysight_Oscilloscope()
mxg = RF_Signal_Generator()
simulating: bool = False

cursor_results: List = []


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


def connections_check_form() -> None:
    """
    connections_check_form _summary_
    """

    layout = [
        [sg.Text("Checking instruments.....", key="-CHECK_MSG-", text_color="Red")],
        [sg.Text("Calibrator", size=(20, 1)), led_indicator("-FLUKE_5700A_CONN-")],
        [sg.Text("33250A", size=(20, 1)), led_indicator("-33250_CONN-")],
        [sg.Text("RF Generator", size=(20, 1)), led_indicator("-RFGEN_CONN-")],
        [sg.Text("UUT", size=(20, 1)), led_indicator("-UUT-")],
        [sg.Text()],
        [sg.Ok(size=(14, 1)), sg.Button("Try Again", size=(14, 1))],
    ]

    window = sg.Window("DSOX Oscilloscope Test", layout, finalize=True)

    connected = test_connections()
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

    while True:
        event, values = window.read()  # type: ignore

        if event in ["Ok", sg.WIN_CLOSED]:
            break

        if event == "Try Again":
            window["-CHECK_MSG-"].update(visible=True)
            connected = test_connections()
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


def test_connections() -> Dict:
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

    return {
        "FLUKE_5700A": fluke_5700a_conn,
        "33250A": ks33250_conn,
        "RFGEN": rfgen_conn,
        "DSO": uut_conn,
    }


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

    with ExcelInterface(filename=filename) as excel:
        excel.backup()

        test_names = set()

        for row in test_rows:
            settings = excel.get_test_settings(row=row)
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
                test_dcv(
                    filename=filename,
                    test_rows=testing_rows,
                    parallel_channels=parallel_channels,
                )

            elif test_name == "POS":
                test_position(
                    filename=filename,
                    test_rows=testing_rows,
                    parallel_channels=parallel_channels,
                )

            elif test_name == "BAL":
                test_dc_balance(filename=filename, test_rows=testing_rows)

            elif test_name == "CURS":
                test_cursor(filename=filename, test_rows=testing_rows)

            elif test_name == "RISE":
                test_risetime(filename=filename, test_rows=testing_rows)

            elif test_name == "TIME":
                test_timebase(filename=filename, row=testing_rows[0])

            elif test_name == "TRIG":
                test_trigger_sensitivity(filename=filename, test_rows=testing_rows)


def test_dc_balance(filename: str, test_rows: List) -> None:
    """
    test_dc_balance
    Test the dc balance of each channel with no signal applied

    Args:
        filename (str): _description_
        test_rows (int): _description_
    """

    global uut

    # no equipment required

    sg.popup("Remove inputs from all channels", background_color="blue")

    uut.reset()

    uut.set_acquisition(32)

    uut.set_timebase(0.001)

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured"
            )
            return

        for row in test_rows:
            excel.row = row

            settings = excel.get_test_settings()

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

    uut.reset()


def test_dcv(filename: str, test_rows: List, parallel_channels: bool = False) -> None:
    # sourcery skip: extract-method
    """
    test_dcv
    Perform the basic DC V tests
    Set the calibrator to the voltage, allow the scope to stabilizee, then read the cursors or measurement values
    """

    global calibrator
    global uut
    global simulating
    global cursor_results

    last_channel = -1

    connections = test_connections()

    # require calibrator

    if not connections["FLUKE_5700A"]:
        sg.popup_error("Cannot find calibrator")
        return

    uut.reset()

    uut.set_timebase(1e-3)

    cursor_results = []  # save results for cursor tests

    if parallel_channels:
        sg.popup(
            "Connect calibrator output to all channels in parallel",
            background_color="blue",
        )

    # Turn off all channels but 1
    for chan in range(uut.num_channels):
        uut.set_channel(chan=chan + 1, enabled=chan == 0)

    uut.set_acquisition(32)

    set_impedance = False

    with ExcelInterface(filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured"
            )
            return
        for row in test_rows:
            excel.row = row

            settings = excel.get_test_settings()

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
                    sg.popup(
                        f"Connect calibrator output to channel {channel}",
                        background_color="blue",
                    )
                last_channel = channel

            uut.set_channel(chan=channel, enabled=True)
            uut.set_voltage_scale(chan=channel, scale=settings.scale)
            uut.set_voltage_offset(chan=channel, offset=settings.offset)

            if settings.impedance:
                uut.set_channel_impedance(chan=channel, impedance=settings.impedance)
                set_impedance = True

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
                        "result": voltage2 - voltage1,
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

        calibrator.reset()
        calibrator.close()

        # Turn off all channels but 1
        for chan in range(uut.num_channels):
            uut.set_channel(chan=chan + 1, enabled=chan == 0)
            uut.set_channel_bw_limit(chan=chan, bw_limit=False)

        uut.reset()
        uut.close()


def test_cursor(filename: str, test_rows: List) -> None:
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

    # no equipment as using buffered results

    with ExcelInterface(filename) as excel:
        for row in test_rows:
            results_col = excel.find_results_col(test_rows[0])
            if results_col == 0:
                sg.popup_error(
                    f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured"
                )
                return
            excel.row = row

            settings = excel.get_test_settings()

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
                        break

        excel.save_sheet()


def test_position(
    filename: str, test_rows: List, parallel_channels: bool = False
) -> None:
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

    connections = test_connections()

    # require calibrator

    if not connections["FLUKE_5700A"]:
        sg.popup_error("Cannot find calibrator")
        return

    uut.reset()

    uut.set_acquisition(32)

    last_channel = -1

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured"
            )
            return

        for row in test_rows:
            excel.row = row

            settings = excel.get_test_settings()

            if settings.channel != last_channel and not parallel_channels:
                sg.popup(
                    f"Connect calibrator output to channel {settings.channel}",
                    background_color="blue",
                )
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
                "Trace within 0.2 div of center?", background_color="blue"
            )

            result = "Pass" if response == "Yes" else "Fail"

            calibrator.standby()

            excel.write_result(result=result, col=results_col)

    calibrator.reset()
    calibrator.close()

    uut.reset()
    uut.close()


# DELAY_PERIOD = 0.00099998  # 1 ms
# DELAY_PERIOD = 0.00100002  # 1 ms
DELAY_PERIOD = 0.001  # 1 ms


def test_timebase(filename: str, row: int) -> None:
    """
    test_timebase
    Test the timebase. Simple single row test

    Args:
        row (int): _description_
    """

    connections = test_connections()

    # require RF gen

    if not connections["33250A"]:
        sg.popup_error("Cannot find 33250A Generator")
        return

    sg.popup("Connect 33250A output to Ch1", background_color="blue")

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(row)
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured"
            )
            return

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
                )

            uut.set_timebase_pos(DELAY_PERIOD)  # delay 1ms to next pulse

            if not uut.keysight:
                valid = False
                while not valid:
                    result = sg.popup_get_text(
                        "Enter difference in div of waveform crossing from center?",
                        background_color="blue",
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

                if not val:
                    # work it out from serial
                    # All the modern Agilent/Keysight have 2 character manufacturing site, then 4 digits for date

                    if len(uut.serial) >= 10:
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

    ks33250.enable_output(False)
    ks33250.close()
    uut.reset()
    uut.close()


def test_trigger_sensitivity(filename: str, test_rows: List) -> None:
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

    connections = test_connections()

    # require RF gen

    if not connections["RFGEN"]:
        sg.popup_error("Cannot find RF Signal Generator")
        return

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
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured"
            )
            return

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
                sg.popup(
                    f"Connect signal generator output to channel {settings.channel} {feedthru_msg}",
                    background_color="blue",
                )
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

    mxg.set_output_state(False)
    mxg.close()

    uut.reset()
    uut.close()


def test_risetime(filename: str, test_rows: List) -> None:
    """
    test_risetime
    Use fast pulse generator to test rise time of each channel

    Args:
        filename (str): _description_
        test_rows (List): _description_
    """

    # only pulse gen required

    uut.reset()

    with ExcelInterface(filename=filename) as excel:
        results_col = excel.find_results_col(test_rows[0])
        if results_col == 0:
            sg.popup_error(
                f"Unable to find results col from row {test_rows[0]}.\nEnsure col headed with results or measured"
            )
            return

        for row in test_rows:
            excel.row = row

            settings = excel.get_tb_test_settings()

            # TODO feedthru

            sg.popup(
                f"Connect fast pulse generator to channel {settings.channel}",
                background_color="blue",
            )

            for chan in range(uut.num_channels):
                uut.set_channel(chan=chan + 1, enabled=settings.channel == chan + 1)

            uut.set_voltage_scale(chan=settings.channel, scale=0.5)

            # TODO set impedance if 50 Ohm

            uut.set_timebase(settings.timebase * 1e-9)
            uut.set_trigger_level(chan=settings.channel, level=0)

            risetime = uut.measure_risetime(chan=settings.channel, num_readings=1) * 1e9

            # save in ns

            excel.write_result(risetime, save=True, col=results_col)

    uut.reset()
    uut.close()


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


def individual_tests(filename: str) -> List:
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
            [sg.Checkbox(name, key=name, background_color="blue", default=True)]
            for name in test_names
        ],
        [sg.Button("Test", size=(10, 1)), sg.Cancel(size=(10, 1))],
    ]

    window = sg.Window(
        "Test individual tests", layout=layout, finalize=True, background_color="blue"
    )

    event, values = window.read()  # type: ignore

    test_steps = []

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

    window.close()

    excel.close()

    return sorted(test_steps)


def load_uut_driver(address: str, simulating: bool = False) -> bool:
    """
    load_uut_driver
    Use a generic driver to figure out which driver of the scope should be used
    """

    global uut

    if simulating:
        uut = Keysight_Oscilloscope(simulate=simulating)
        uut.open_connection()
        return True

    with SCPI_ID(address=address) as scpi_uut:
        manfacturer = scpi_uut.get_manufacturer()

        if manfacturer == "":
            sg.popup_error("Unable to contact UUT. Is address correct?")
            return False
        elif manfacturer == "KEYSIGHT":
            uut = Keysight_Oscilloscope(simulate=simulating)
            uut.open_connection()
        elif manfacturer == "TEKTRONIX":
            uut = Tektronix_Oscilloscope(simulate=simulating)
            uut.open_connection()
        else:
            sg.popup_error(f"No driver for {manfacturer}. Using Tektronix driver")
            uut = Tektronix_Oscilloscope(simulate=simulating)
            uut.open_connection()

        uut.num_channels = scpi_uut.get_number_channels()

        return True


if __name__ == "__main__":
    sg.theme("black")

    gpib_ifc_list = [f"GPIB{x}" for x in range(5)]
    gpib_addresses = list(range(1, 32))

    layout = [
        [sg.Text("DSOX Oscilloscope Test")],
        [sg.Text(f"DK Jan 23 VERSION {VERSION}")],
        [sg.Text()],
        [sg.Text("Create Excel sheet from template first", text_color="red")],
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
            sg.Button("Test Connections", size=(15, 1), key="-TEST_CONNECTIONS-"),
            sg.Button("Individual Tests", size=(12, 1), key="-INDIVIDUAL-"),
            sg.Exit(size=(12, 1)),
        ],
    ]

    window = sg.Window("Oscilloscope Test", layout=layout, finalize=True)

    back_color = window["-FILE-"].BackgroundColor

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

            with ExcelInterface(filename=values["-FILE-"]) as excel:
                if not excel.check_valid_results():
                    sg.popup_error(
                        "No Named range for StartCell in results",
                        background_color="blue",
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
            calibrator.open_connection()
            ks33250.visa_address = ks33250_address
            ks33250.open_connection()
            mxg.visa_address = mxg_address
            mxg.open_connection()

            if load_uut_driver(address=values["-UUT_ADDRESS-"], simulating=simulating):
                uut.visa_address = values["-UUT_ADDRESS-"]
                uut.open_connection()

                test_rows = individual_tests(filename=values["-FILE-"])
                if len(test_rows):
                    parallel = sg.popup_yes_no(
                        "Will you connect all channels in parallel?",
                        title="Parallel Channels",
                        background_color="blue",
                    )
                    run_tests(
                        filename=values["-FILE-"],
                        test_rows=test_rows,
                        parallel_channels=parallel == "Yes",
                    )

                    sg.popup("Finished", background_color="blue")
            window["-VIEW-"].update(disabled=False)

        if event == "-TEST_CONNECTIONS-":
            if values["-CALIBRATOR-"] == "M-142":
                calibrator = M142(simulate=simulating)
            else:
                calibrator = Fluke5700A(simulate=simulating)
            calibrator.visa_address = calibrator_address
            calibrator.open_connection()
            ks33250.visa_address = ks33250_address
            ks33250.open_connection()
            mxg.visa_address = mxg_address
            mxg.open_connection()
            uut.visa_address = values["-UUT_ADDRESS-"]
            uut.open_connection()

            connections_check_form()
            continue

        if event == "-VIEW-":
            os.startfile(f'"{values["-FILE-"]}"')
