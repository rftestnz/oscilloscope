"""
Test DSOX Oscilloscopes
# DK Jan 23
"""

from typing import Dict, List
import PySimpleGUI as sg

from drivers.fluke_5700a import Fluke5700A
from drivers.meatest_m142 import M142
from drivers.dsox_3000 import DSOX_3000
from drivers.excel_interface import ExcelInterface
import os
import sys
import time
from pathlib import Path

VERSION = "A.00.00"


calibrator = Fluke5700A()
calibrator_address: str = "GPIB0::06::INSTR"
uut = DSOX_3000()
simulating: bool = False


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


def led_indicator(key=None, radius=30):
    return sg.Graph(
        canvas_size=(radius, radius),
        graph_bottom_left=(-radius, -radius),
        graph_top_right=(radius, radius),
        pad=(0, 0),
        key=key,
    )


def set_led(window, key, color):
    graph = window[key]
    graph.erase()
    graph.draw_circle((0, 0), 12, fill_color=color, line_color=color)


def connections_check_form() -> None:

    layout = [
        [sg.Text("Checking instruments.....", key="-CHECK_MSG-", text_color="Red")],
        [sg.Text("Calibrator", size=(20, 1)), led_indicator("-FLUKE_5700A_CONN-")],
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
        "-UUT-",
        color="green" if connected["DSO"] else "red",
    )

    while True:
        event, values = window.read()

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
    global uut

    fluke_5700a_conn = calibrator.is_connected()
    uut_conn = uut.is_connected()

    return {"FLUKE_5700A": fluke_5700a_conn, "DSO": uut_conn}


def test_dcv(filename: str, test_rows: List) -> None:
    """
    test_dcv
    Perform the basic DC V tests
    Set the calibrator to the voltage, allow the scope to stabilizee, then read the cursors or measurement values
    """

    global simulating

    # TODO read both the cursor and mean at the same time

    last_channel = -1

    # Turn off all channels but 1
    for chan in range(uut.num_channels):
        uut.set_channel(chan=chan + 1, enabled=chan == 0)

    with ExcelInterface(filename) as excel:

        for row in test_rows:
            excel.row = row

            settings = excel.get_test_settings()

            calibrator.set_voltage_dc(settings.voltage)

            channel = settings.channel

            if channel != last_channel:
                if last_channel:
                    uut.set_voltage_scale(chan=last_channel, scale=1)
                    uut.set_voltage_offset(chan=last_channel, offset=0)
                    uut.set_channel(chan=last_channel, enabled=False)
                sg.popup(
                    f"Connect calibrator output to channel {channel}",
                    background_color="blue",
                )
                last_channel = channel

            uut.set_channel(chan=channel, enabled=True)
            uut.set_voltage_scale(chan=channel, scale=settings.scale)
            uut.set_voltage_offset(chan=channel, offset=settings.offset)

            calibrator.operate()

            if not simulating:
                time.sleep(2)

            reading = uut.measure_voltage(chan=channel)

            calibrator.standby()

            excel.write_result(reading)  # auto saving

        # excel.save_sheet()

        calibrator.close()

        # Turn off all channels but 1
        for chan in range(uut.num_channels):
            uut.set_channel(chan=chan + 1, enabled=chan == 0)

        uut.close()

        sg.popup("Finished", background_color="blue")


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
                default_text=sg.user_settings_get_entry("-FILENAME-"),
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
                default=sg.user_settings_get_entry("-SIMULATE-"),
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
            sg.Text("UUT", size=(15, 1)),
            sg.Input(
                default_text=sg.user_settings_get_entry("-UUT_ADDRESS-"),
                size=(60, 1),
                key="-UUT_ADDRESS-",
            ),
        ],
        [sg.Text()],
        [
            sg.Button("Test Connections", size=(15, 1), key="-TEST_CONNECTIONS-"),
            sg.Button("Test DCV", size=(12, 1), key="-TEST_DCV-"),
            sg.Exit(size=(12, 1)),
        ],
    ]

    window = sg.Window("Oscilloscope Test", layout=layout, finalize=True)

    back_color = window["-FILE-"].BackgroundColor

    simulating = False

    while True:
        event, values = window.read(200)

        if event in ["Exit", sg.WIN_CLOSED]:
            break

        # Get the default text color in case changing simulate
        txt_clr = sg.theme_element_text_color()

        # read the simulate setting
        if values["-SIMULATE-"]:
            window["-SIMULATE-"].update(text_color="red")
        else:
            window["-SIMULATE-"].update(text_color=txt_clr)

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
            calibrator.visa_address = calibrator_address
            calibrator.open_connection()

            uut = DSOX_3000(simulate=simulating)
            uut.visa_address = values["-UUT_ADDRESS-"]

            uut.open_connection()

            with ExcelInterface(values["-FILE-"]) as excel:
                test_rows = excel.get_test_rows("DCV")

            test_dcv(filename=values["-FILE-"], test_rows=test_rows)

            window["-VIEW-"].update(disabled=False)

        sg.user_settings_set_entry("-SIMULATE-", values["-SIMULATE-"])
        sg.user_settings_set_entry("-CALIBRATOR-", values["-CALIBRATOR-"])
        sg.user_settings_set_entry("-FLUKE_5700A_GPIB_IFC-", values["GPIB_FLUKE_5700A"])
        sg.user_settings_set_entry("-UUT_ADDRESS-", values["-UUT_ADDRESS-"])
        sg.user_settings_set_entry("-FILENAME-", values["-FILE-"])

        sg.user_settings_save()

        simulating = values["-SIMULATE-"]

        calibrator.simulating = simulating
        calibrator_address = (
            f"{values['GPIB_FLUKE_5700A']}::{values['GPIB_ADDR_FLUKE_5700A']}::INSTR"
        )

        uut.simulating = simulating
        uut.visa_address = values["-UUT_ADDRESS-"]

        if event == "-TEST_CONNECTIONS-":
            if values["-CALIBRATOR-"] == "M-142":
                calibrator = M142(simulate=simulating)
            else:
                calibrator = Fluke5700A(simulate=simulating)
            calibrator.visa_address = calibrator_address
            calibrator.open_connection()
            connections_check_form()
            continue

        if event == "-VIEW-":
            os.startfile(f'"{values["-FILE-"]}"')
