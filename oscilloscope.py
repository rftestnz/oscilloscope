"""
Test DSOX Oscilloscopes
# DK Jan 23
"""

from typing import Dict, List
import PySimpleGUI as sg

from drivers.fluke_5700a import Fluke5700A
from drivers.Ks33250A import Ks33250A
from drivers.meatest_m142 import M142
from drivers.dsox_3000 import DSOX_3000
from drivers.excel_interface import ExcelInterface
import os
import sys
import time
from pathlib import Path

VERSION = "A.00.00"


calibrator = Fluke5700A()
ks33250 = Ks33250A()
uut = DSOX_3000()
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
    global uut

    fluke_5700a_conn = calibrator.is_connected()
    ks33250_conn = ks33250.is_connected()
    uut_conn = uut.is_connected()

    return {"FLUKE_5700A": fluke_5700a_conn, "33250A": ks33250_conn, "DSO": uut_conn}


def test_dcv(filename: str, test_rows: List) -> None:
    # sourcery skip: extract-method
    """
    test_dcv
    Perform the basic DC V tests
    Set the calibrator to the voltage, allow the scope to stabilizee, then read the cursors or measurement values
    """

    global simulating
    global cursor_results

    last_channel = -1

    uut.reset()

    cursor_results = []  # save results for cursor tests

    # Turn off all channels but 1
    for chan in range(uut.num_channels):
        uut.set_channel(chan=chan + 1, enabled=chan == 0)

    uut.set_acquisition(64)

    with ExcelInterface(filename) as excel:

        for row in test_rows:
            excel.row = row

            settings = excel.get_test_settings()

            calibrator.set_voltage_dc(0)  # type: ignore

            channel = settings.channel  # type: ignore

            if channel > uut.num_channels:
                continue

            if channel != last_channel:
                if last_channel > 0:
                    uut.set_voltage_scale(chan=last_channel, scale=1)
                    uut.set_voltage_offset(chan=last_channel, offset=0)
                    uut.set_channel(chan=last_channel, enabled=False)
                    uut.set_channel(chan=channel, enabled=True)

                # uut.set_channel_bw_limit(chan=channel, bw_limit=True)
                uut.set_voltage_scale(chan=channel, scale=5)
                uut.set_voltage_offset(chan=channel, offset=0)
                uut.set_cursor_xy_source(chan=1, cursor=1)
                uut.set_cursor_position(cursor="X1", pos=0)
                sg.popup(
                    f"Connect calibrator output to channel {channel}",
                    background_color="blue",
                )
                last_channel = channel

            uut.set_channel(chan=channel, enabled=True)
            uut.set_voltage_scale(chan=channel, scale=settings.scale)  # type: ignore
            uut.set_voltage_offset(chan=channel, offset=settings.offset)  # type: ignore

            calibrator.operate()

            if not simulating:
                time.sleep(1)

            voltage1 = uut.read_cursor_avg()

            calibrator.set_voltage_dc(settings.voltage)  # type: ignore
            calibrator.operate()

            if not simulating:
                time.sleep(1)

            reading = uut.measure_voltage(chan=channel)
            units = excel.get_units()

            if units.startswith("m"):
                reading *= 1000

            voltage2 = uut.read_cursor_avg()

            cursor_results.append({"chan": channel, "scale": settings.scale, "result": voltage2 - voltage1})  # type: ignore

            calibrator.standby()

            excel.write_result(reading)  # auto saving

        # excel.save_sheet()

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

    with ExcelInterface(filename) as excel:
        for row in test_rows:
            excel.row = row

            settings = excel.get_test_settings()

            if len(cursor_results):
                for res in cursor_results:
                    if (
                        res["chan"] == settings.channel  # type: ignore
                        and res["scale"] == settings.scale  # type: ignore
                    ):
                        units = excel.get_units()
                        result = res["result"]
                        if units.startswith("m"):
                            result *= 1000
                        excel.write_result(result, save=False)
                        break

        excel.save_sheet()


def test_timebase(self, filename: str, row: int) -> None:
    """
    test_timebase
    Test the timebase. Simple single row test

    Args:
        row (int): _description_
    """

    sg.popup("Connect 33250A output to Ch1", background_color="blue")

    uut.reset()

    uut.set_channel(chan=1, enabled=True)

    uut.set_voltage_scale(chan=1, scale=0.5)
    uut.set_voltage_offset(chan=1, offset=0)

    ks33250.set_pulse(period=1e-3, pulse_width=200e-6, amplitude=1)

    uut.set_trigger_level(level=0.5, chan=1)
    uut.set_timebase(10e-9)  # TODO add setting to sheet
    time.sleep(0.1)
    ref_x = uut.read_cursor("X1")
    ref = uut.read_cursor("Y1")

    uut.set_cursor_position(cursor="X1", pos=0.001)  # 1 ms delay
    time.sleep(0.1)

    uut.adjust_cursor(target=ref)

    offset_x = uut.read_cursor("X1")

    error = ref_x - offset_x + 0.001
    print(f"TB Error {error}")

    with ExcelInterface(filename=filename) as excel:
        if uut.manufacturer == "Keysight":
            # results in ppm
            ppm = error / 1e-3 * 1e6
            excel.row = row
            excel.write_result(ppm, save=True, col=2)

    uut.reset()


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
            sg.Text("UUT", size=(15, 1)),
            sg.Input(
                default_text=sg.user_settings_get_entry("-UUT_ADDRESS-"),  # type: ignore
                size=(60, 1),
                key="-UUT_ADDRESS-",
            ),
        ],
        [sg.Text()],
        [
            sg.Button("Test Connections", size=(15, 1), key="-TEST_CONNECTIONS-"),
            sg.Button("Test DCV", size=(12, 1), key="-TEST_DCV-"),
            sg.Button("Test Timebase", size=(12, 1), key="-TEST_TB-"),
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

        if event in ["-TEST_DCV-"]:
            # Common check to make sure everything is in order

            valid = True

            # reset the back colors instead of else statements after checking

            window["-FILE-"].update(background_color=back_color)

            if not values["-FILE-"]:
                window["-FILE-"].update(background_color="Red")
                valid = False

            if not valid:
                continue

            window["-VIEW-"].update(disabled=True)  # Disable while testing

            if values["-CALIBRATOR-"] == "M-142":
                calibrator = M142(simulate=simulating)
            else:
                calibrator = Fluke5700A(simulate=simulating)
            calibrator.visa_address = calibrator_address  # type: ignore
            calibrator.open_connection()
            ks33250.visa_address = ks33250_address  # type: ignore
            ks33250.open_connection()

            uut = DSOX_3000(simulate=simulating)
            uut.visa_address = values["-UUT_ADDRESS-"]

            uut.open_connection()

            with ExcelInterface(values["-FILE-"]) as excel:
                test_rows = excel.get_test_rows("DCV")
                test_dcv(filename=values["-FILE-"], test_rows=test_rows)
                test_rows = excel.get_test_rows("CURS")
                if len(test_rows):
                    test_cursor(filename=values["-FILE-"], test_rows=test_rows)

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
            connections_check_form()
            continue

        if event == "-VIEW-":
            os.startfile(f'"{values["-FILE-"]}"')
