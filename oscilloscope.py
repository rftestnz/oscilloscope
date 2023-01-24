#TODO basic script description
# DK Jan 23

from typing import Dict, List
import PySimpleGUI as sg

from drivers.driver_boilerplate import DRIVER_BOILERPLATE   # TODO fix the import names
from drivers.fluke_5700a import FLUKE_5700A   # TODO fix the import names
from drivers.meatest_m142 import MEATEST_M142   # TODO fix the import names
import os
import sys

VERSION = "A.00.00"

driver_boilerplate = DRIVER_BOILERPLATE()    # TODO fix the class names
driver_boilerplate_address:str = "GPIB0::00::INSTR"    #TODO set default address
fluke_5700a = FLUKE_5700A()    # TODO fix the class names
fluke_5700a_address:str = "GPIB0::00::INSTR"    #TODO set default address
meatest_m142 = MEATEST_M142()    # TODO fix the class names
meatest_m142_address:str = "GPIB0::00::INSTR"    #TODO set default address

def get_path(filename):
    """
    get_path
    The location of the diagram is different if run through interpreter or compiled.

    Args:
        filename ([type]): [description]

    Returns:
        [type]: [description]
    """

    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    else:
        return filename

def led_indicator(key=None, radius=30):
    return sg.Graph(canvas_size=(radius, radius),graph_bottom_left=(-radius, -radius),graph_top_right=(radius, radius),pad=(0, 0),key=key,)

def set_led(window, key, color):
    graph = window[key]
    graph.erase()
    graph.draw_circle((0, 0), 12, fill_color=color, line_color=color)

def connections_check_form() -> None:

    layout = [
        [sg.Text("Checking instruments.....", key="-CHECK_MSG-", text_color="Red")],
        [sg.Text("DRIVER_BOILERPLATE", size=(20, 1)), led_indicator("-DRIVER_BOILERPLATE_CONN-")],
        [sg.Text("FLUKE_5700A", size=(20, 1)), led_indicator("-FLUKE_5700A_CONN-")],
        [sg.Text("MEATEST_M142", size=(20, 1)), led_indicator("-MEATEST_M142_CONN-")],
        [sg.Ok(size=(14, 1)), sg.Button("Try Again", size=(14, 1))],
    ]

    window = sg.Window("Project name", layout, finalize=True)   # TODO Update project name

    connected = test_connections()
    window["-CHECK_MSG-"].update(visible=False)

    set_led(window, "-DRIVER_BOILERPLATE_CONN-", color="green" if connected["DRIVER_BOILERPLATE"] else "red")
    set_led(window, "-FLUKE_5700A_CONN-", color="green" if connected["FLUKE_5700A"] else "red")
    set_led(window, "-MEATEST_M142_CONN-", color="green" if connected["MEATEST_M142"] else "red")

    while True:
        event, values = window.read()

        if event in ["Ok", sg.WIN_CLOSED]:
            break

        if event == "Try Again":
            window["-CHECK_MSG-"].update(visible=True)
            connected = test_connections()
            window["-CHECK_MSG-"].update(visible=False)
            set_led(window,"-DRIVER_BOILERPLATE_CONN-",color="green" if connected["DRIVER_BOILERPLATE"] else "red",)
            set_led(window,"-FLUKE_5700A_CONN-",color="green" if connected["FLUKE_5700A"] else "red",)
            set_led(window,"-MEATEST_M142_CONN-",color="green" if connected["MEATEST_M142"] else "red",)

    window.close()

def test_connections() -> Dict:
    """
    Make sure all of the instruments are connected
    """

    global driver_boilerplate    # TODO name of instance
    global fluke_5700a    # TODO name of instance
    global meatest_m142    # TODO name of instance

    driver_boilerplate_conn = driver_boilerplate.is_connected()   # TODO name of instance
    fluke_5700a_conn = fluke_5700a.is_connected()   # TODO name of instance
    meatest_m142_conn = meatest_m142.is_connected()   # TODO name of instance

    return {
        "DRIVER_BOILERPLATE": driver_boilerplate_conn,
        "FLUKE_5700A": fluke_5700a_conn,
        "MEATEST_M142": meatest_m142_conn,
    }

if __name__ == "__main__":
    sg.theme("black")

    gpib_ifc_list = [f'GPIB{x}' for x in range(5)]
    gpib_addresses = list(range(1, 32))

    layout = [
        [sg.Text("Project description")],   #TODO Add description
        [sg.Text(f"DK Jan 23 VERSION {VERSION}")],
        [sg.Text()],
        [sg.Text("Create Excel sheet from template first", text_color="red")],        [sg.Text("Results file", size=(10, 1)), sg.Input(size=(60, 1),key="-FILE-",default_text=""),        sg.FileBrowse("...",target="-FILE-",file_types=(("Excel Files", "*.xlsx"), ("All Files", "*.*")),        initial_folder=sg.user_settings_get_entry("-RESULTS_FOLDER-"),), sg.Button("View", size=(15, 1), key="-VIEW-", disabled=False),],        [sg.Text()],
        [sg.Check("Simulate", default=sg.user_settings_get_entry("-SIMULATE-"), key="-SIMULATE-", disabled=False)],
        [sg.Text("Instrument Config:")],
        [sg.Text("DRIVER_BOILERPLATE",size=(15,1)), sg.Combo(gpib_ifc_list, size=(10, 1), key="GPIB_DRIVER_BOILERPLATE", default_value=sg.user_settings_get_entry("-DRIVER_BOILERPLATE_GPIB_IFC-")), sg.Combo(gpib_addresses, default_value="0", size=(6, 1), key="GPIB_ADDR_DRIVER_BOILERPLATE")],  #TODO set defaults
        [sg.Text("FLUKE_5700A",size=(15,1)), sg.Combo(gpib_ifc_list, size=(10, 1), key="GPIB_FLUKE_5700A", default_value=sg.user_settings_get_entry("-FLUKE_5700A_GPIB_IFC-")), sg.Combo(gpib_addresses, default_value="0", size=(6, 1), key="GPIB_ADDR_FLUKE_5700A")],  #TODO set defaults
        [sg.Text("MEATEST_M142",size=(15,1)), sg.Combo(gpib_ifc_list, size=(10, 1), key="GPIB_MEATEST_M142", default_value=sg.user_settings_get_entry("-MEATEST_M142_GPIB_IFC-")), sg.Combo(gpib_addresses, default_value="0", size=(6, 1), key="GPIB_ADDR_MEATEST_M142")],  #TODO set defaults
        [sg.Text()],
        [sg.Button("Test Connections", size=(15, 1), key="-TEST_CONNECTIONS-"),sg.Button("Ok", size=(12, 1), key="-OK-"), sg.Exit(size=(12, 1))],
    ]

    window=sg.Window("oscilloscope", layout=layout, finalize=True)    #TODO Set window title

    back_color = window["-FILE-"].BackgroundColor

    while True:
        event, values=window.read(200)

        if event in ["Exit", sg.WIN_CLOSED]:
            break

        # Get the default text color in case changing simulate
        txt_clr = sg.theme_element_text_color()

        # read the simulate setting
        if values["-SIMULATE-"]:
            window["-SIMULATE-"].update(text_color="red")
        else:
            window["-SIMULATE-"].update(text_color=txt_clr)

        if event in ["-OK-"]:  # TODO button names
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


        sg.user_settings_set_entry("-SIMULATE-", values["-SIMULATE-"])
        sg.user_settings_set_entry("-DRIVER_BOILERPLATE_GPIB_IFC-", values["GPIB_DRIVER_BOILERPLATE"])
        sg.user_settings_set_entry("-FLUKE_5700A_GPIB_IFC-", values["GPIB_FLUKE_5700A"])
        sg.user_settings_set_entry("-MEATEST_M142_GPIB_IFC-", values["GPIB_MEATEST_M142"])
        sg.user_settings_save()

        simulate = values["-SIMULATE-"]

        driver_boilerplate.simulating = simulate
        driver_boilerplate_address=f"{values['GPIB_DRIVER_BOILERPLATE']}::{values['GPIB_ADDR_DRIVER_BOILERPLATE']}::INSTR"
        fluke_5700a.simulating = simulate
        fluke_5700a_address=f"{values['GPIB_FLUKE_5700A']}::{values['GPIB_ADDR_FLUKE_5700A']}::INSTR"
        meatest_m142.simulating = simulate
        meatest_m142_address=f"{values['GPIB_MEATEST_M142']}::{values['GPIB_ADDR_MEATEST_M142']}::INSTR"

        if event == "-TEST_CONNECTIONS-":
            connections_check_form()
            continue

        if event == "-VIEW-":
            os.startfile(f'"{values["-FILE-"]}"')
