"""
# Driver for Tek DPO2000


# DK Jan 23
"""

import pyvisa
import time
from random import random
from typing import List
import numpy as np
from struct import unpack

try:
    from drivers.base_scope_driver import ScopeDriver, Scope_Simulator
except ModuleNotFoundError:
    from base_scope_driver import ScopeDriver, Scope_Simulator

VERSION = "A.00.00"


class RohdeSchwarz_Oscilloscope(ScopeDriver):
    """
     _summary_

    Returns:
        _type_: _description_
    """

    connected: bool = False
    visa_address: str = "USB0::0x0699::0x0373::C010049::INSTR"
    model = ""
    manufacturer = ""
    serial = ""
    timeout = 5000
    num_channels: int = 4
    keysight: bool = False

    def __init__(self, simulate=False):
        self.simulating = simulate
        self.rm = pyvisa.ResourceManager()

    def __enter__(self):
        return super().__enter__()

    def __exit__(self, exc_type, exc_val, exc_tb):
        return super().__exit__(exc_type, exc_val, exc_tb)

    def open_connection(self) -> bool:
        """
        open_connection _summary_

        Returns:
            bool: _description_
        """
        try:
            if self.simulating:
                self.instr = Scope_Simulator()
                self.model = "RTH1004"
                self.manufacturer = "R&S"
                self.serial = "666"
            else:
                self.instr = self.rm.open_resource(
                    self.visa_address, write_termination="\n"
                )
                self.instr.timeout = self.timeout
                # self.instr.control_ren(VI_GPIB_REN_ASSERT)  # type: ignore
                self.get_id()
            self.connected = True
        except Exception:
            self.connected = False

        return self.connected

    def initialize(self) -> None:
        """
        initialize _summary_

        Returns:
            _type_: _description_
        """
        return super().initialize()

    def close(self) -> None:
        """
        close _summary_

        Returns:
            _type_: _description_
        """
        return super().close()

    def is_connected(self) -> bool:
        """
        is_connected _summary_

        Returns:
            bool: _description_
        """
        return bool(
            self.open_connection()
            and (not self.simulating and self.model.find("RTH") >= 0 or self.simulating)
        )

    def write(self, command: str) -> None:
        """
        write _summary_

        Args:
            command (str): _description_

        Returns:
            _type_: _description_
        """

        attempts = 0

        while attempts < 3:
            try:
                self.instr.write(command)  # type: ignore
                break
            except pyvisa.VisaIOError:
                time.sleep(1)
                attempts += 1

    def read(self) -> str:
        """
        read _summary_

        Returns:
            str: _description_
        """

        attempts = 0

        ret: str = ""

        while attempts < 3:
            try:
                ret = self.instr.read()  # type: ignore
                break
            except pyvisa.VisaIOError:
                time.sleep(1)
                attempts += 1

        return ret

    def query(self, command: str) -> str:
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """

        assert command.find("?") > 0

        attempts = 0
        ret = ""

        while attempts < 3:
            try:
                ret = self.instr.query(command)  # type: ignore
                break
            except pyvisa.VisaIOError:
                time.sleep(1)
                attempts += 1

        return ret

    def read_query(self, command: str) -> float:
        """
        read_query _summary_

        Args:
            command (str): _description_

        Returns:
            float: _description_
        """

        assert command.find("?") > 0

        reply = self.query(command)

        try:
            val = float(reply)
        except ValueError:
            val = 0.0

        return val

    def reset(self) -> None:
        """
        reset _summary_
        """
        self.write("*CLS")
        self.write("*RST")
        self.write("*OPC")

    def get_id(self) -> List:
        """
        get_id _summary_

        Returns:
            List: _description_
        """
        try:
            self.instr.timeout = 2000  # type: ignore
            response = self.instr.query("*IDN?")  # type: ignore
            print(response)
            identity = response.split(",")
            if len(identity) >= 3:
                self.manufacturer = identity[0]
                self.model = identity[1]
                self.serial = identity[2]

        except pyvisa.VisaIOError:
            self.model = ""
            self.manufacturer = ""
            self.serial = ""
            response = ",,,"

        self.instr.timeout = self.timeout  # type: ignore

        return response.split(",")

    def set_channel_bw_limit(self, chan: int, bw_limit: bool | int | str) -> None:
        """
        set_channel_bw_limit
        Set bandwidth limit on or off

        Args:
            chan (int): _description_
            bw_limit (bool): _description_
        """

        if type(bw_limit) is bool:
            state = "B20" if bw_limit else "FULL"
        elif type(bw_limit) is int:
            # Assuming MHz
            state = f"B{bw_limit}"
        else:
            # str, last character should be k for kHz
            assert str(bw_limit)[-1] in {"k", "M"}, " Bandwidth kHz must end with k"
            # if the third character is H, then it is hundreds of kHz

            if str(bw_limit)[-1] == "M":
                state = f"B{str(bw_limit)[:-1]}"
            elif bw_limit == "500k":
                state = "B5HK"
            elif bw_limit == "400k":
                state = "B4HK"
            elif bw_limit == "200k":
                state = "B2HK"
            elif bw_limit == "100k":
                state = "B1HK"
            elif bw_limit == "50k":
                state = "B50K"
            elif bw_limit == "40k":
                state = "B40K"
            elif bw_limit == "20k":
                state = "B20K"
            elif bw_limit == "10k":
                state = "B10K"
            elif bw_limit == "5k":
                state = "B5K"
            elif bw_limit == "4k":
                state = "B4K"
            elif bw_limit == "2k":
                state = "B2K"
            elif bw_limit == "1K":
                state = "B1K"
            else:
                print(f"Invalid bandwidth {bw_limit}")
                state = "FULL"

        # Some use BAN and some BAND, so use the full command
        self.write(f"CHAN{chan}:BANDWIDTH {state}")
        self.write("*OPC")

    def set_channel_impedance(self, chan: int, impedance: str) -> None:
        """
        set_channel_impedance

        Although not all models can set impedance, command is available for compatibility

        Args:
            chan (int): _description_
            imedance (str): _description_
        """

        # Not supported for this scope

        pass

    def set_channel_invert(self, chan: int, inverted: bool) -> None:
        """
        set_channel_invert _summary_

        Args:
            inverted (bool): _description_
        """

        state = "INVERTED" if inverted else "NORMAL"

        self.write(f"CHAN{chan}:POL {state}")

    def set_channel(self, chan: int, enabled: bool, only: bool = False) -> None:
        """
        set_channel
        Turn display of channel on or off

        Args:
            chan (int): _description_
            enabled (bool): _description_
        """

        if only:
            for channel in range(1, self.num_channels + 1):
                state = "ON" if channel == chan else "OFF"
                self.write(f"SEL:CH{channel} {state}")

        else:
            state = "ON" if enabled else "OFF"
            self.write(f"CHAN{chan}:STATE {state}")

        self.write("*OPC")

    def set_channel_coupling(self, chan: int, coupling: str) -> None:
        """
        set_channel_coupling _summary_

        Args:
            chan (int): _description_
            coupling (str): _description_
        """

        # DCL or ACL. No GND

        if coupling.upper() != "GND":
            if not coupling.endswith("L"):
                coupling += "L"

            self.write(f"CHAN{chan}:COUP {coupling}")

    def set_voltage_scale(self, chan: int, scale: float, probe_atten: int = 1) -> None:
        """
        set_voltage_scale _summary_

        Args:
            chan (int): _description_
            scale (float): _description_
        """

        self.write(f"CHAN{chan}:PROBE V1TO1")
        self.write(f"CHAN{chan}:SCALE {scale}")

    def set_voltage_offset(self, chan: int, offset: float) -> None:
        """
        set_voltage_offset
        For Tek, offset is the center of the vertical acquisition window

        Args:
            chan (int): _description_
            offset (float): _description_
        """

        self.write(f"CHAN{chan}:OFFS {offset}")

    def set_voltage_position(self, chan: int, position: float) -> None:
        """
        set_voltage_position
        set the position, in divisions

        Args:
            chan (int): _description_
            position (float): _description_
        """

        self.write(f"CHAN{chan}:POS {position}")

    def set_timebase(self, timebase: float) -> None:
        """
        set_timebase


        Args:
            timebase (float): _description_
        """

        self.write(f"TIM:SCALE {timebase}")

    def set_timebase_pos(self, pos: float) -> None:
        """
        set_timebase_pos _summary_

        Args:
            pos (float): _description_
        """

        self.write(f"TIM:HOR:POS {pos}")

    def set_acquisition(self, num_samples: int) -> None:
        """
        set_acquisition _summary_

        Args:
            num_samples (int): _description_
        """

        self.write("ACQ:MODE AVER")
        self.write(f"ACQ:AVER:COUNT {num_samples}")

    def set_trigger_type(self, mode: str, auto_trig: bool = True) -> None:
        """
        set_trigger_mode _summary_

        Args:
            mode (str): _description_
        """

        self.write("TRIG:TYPE EDGE")
        trig_mode = "AUTO" if auto_trig else "NORMAL"
        self.write(f"TRIG:MODE {trig_mode}")

    def set_trigger_level(self, level: float, chan: int) -> None:
        """
        set_trigger_level
        Set the trigger level and source

        Args:
            level (float): _description_
            chan (int): _description_
        """

        self.write(f"TRIG:SOUR C{chan}")
        self.write(f"TRIG:LEV{chan}:VAL {level}")

    def measure_voltage(self, chan: int, delay: float = 6) -> float:
        """
        measure_voltage _summary_

        Args:
            chan (int): _description_
            delay (float, optional): Number seconds to allow measurement. Defaults to 1.

        Returns:
            float: _description_
        """

        # Only using measurement 1

        self.write(f"MEAS1:SOURCE C{chan}")
        self.write("MEAS1:TYPE MEAN")

        self.write("MEAS1:ENABLE ON")

        if not self.simulating:
            time.sleep(delay)

        return self.read_query("MEAS1:RESULT:ACTUAL?")

    def measure_clear(self) -> None:
        """
        measure_clear _summary_
        """

        self.write("MEAS1:ENABLE OFF")

    def measure_risetime(self, chan: int, num_readings: int = 1) -> float:
        """
        measure_risetime
        Use the measure function to mneasure the rise time average of n measurements

        Args:
            chan (int): _description_
            num_readings (int, optional): _description_. Defaults to 1.

        Returns:
            float: _description_
        """

        self.measure_clear()

        self.write("MEAS:TYPE RTIM")
        self.write(f"MEAS1:SOURCE C{chan}")
        self.write("MEAS1:ENABLE ON")

        if not self.simulating:
            time.sleep(2)  # allow time to measure

        # TODO check if RTH will automatically average successive readings

        total = 0
        for _ in range(num_readings):
            total += self.read_query("MEAS1:RESULT:ACTUAL?")
            time.sleep(0.1)

        return total / num_readings

    def read_cursor(self, cursor: str) -> float:
        """
        read_cursor
        Enable cursor and read

        Args:
            cursor (str): X1, X2, Y1, Y2

        Returns:
            float: _description_
        """

        # TODO check this

        self.write("CURSOR:FUNC MEAS")
        self.write("CURS:MEAS:TYPE MEAN")
        self.write("CURS:STATE ON")
        return self.read_query("CURS:MEAS1:RESULT:ACTUAL?")

    def read_cursor_avg(self) -> float:
        """
        read_cursor_avg
        Read both of the Y cursors, and return the average

        Returns:
            float: _description_
        """

        # TODO does this work?
        return self.read_cursor("Y1")

    def read_cursor_ydelta(self) -> float:
        """
        read_cursor_ydelta _summary_

        Returns:
            float: _description_
        """

        # TODO does the mode need to be set?
        return self.read_query("CURS:DELTA?")

    def set_cursor_xy_source(self, chan: int, cursor: int) -> None:
        """
        set_cursor_xy_source
        Set the X1Y1 marker source

        Args:
            chan (int): _description_
        """

        # TODO figure this out
        self.write("CURS:FUNC TRACK")
        self.write("CURS:MOD:TRACK")
        self.write(f"CURS:X{cursor}Y{cursor} CHAN{chan}")

    def set_cursor_position(self, cursor: str, pos: float) -> None:
        """
        set_cursor_position _summary_

        Args:
            cursor (str): _description_
            pos (float): _description_
        """

        pass  # not supported

    def adjust_cursor(self, target: float) -> None:
        """
        adjust_cursor
        Adjust the cursor time until target voltage met

        Args:
            target (float): _description_
        """

        current_y = self.read_query("MARK:Y1P?")
        current_x = self.read_query("MARK:X1P?")

        time_inc = self.read_query("TIM:SCAL?") / 20

        direction = +1 if current_y < target else -1

        if current_y < target:
            for _ in range(100):
                self.set_cursor_position(
                    cursor="X1", pos=current_x + time_inc * direction
                )
                current_x = self.read_query("MARK:X1P?")
                current_y = self.read_query("MARK:Y1P?")
                if ((current_y > target) and direction == 1) or (
                    (current_y < target) and direction == -1
                ):
                    break
        else:
            ...

    def cursors_on(self) -> None:
        """
        cursors_on
        Turn the markers on
        """

        self.write("CURS:STATE ON")
        self.write("*OPC")

    def check_triggered(self, sweep_time: float = 0.1) -> bool:
        """
        check_triggered
        Check the state of the trigger

        Args:
            sweep_time (float, optional): _description_. Defaults to 0.1.

        Returns:
            bool: _description_
        """

        # TODO is there a command?
        response = self.query("TRIG:STATE?").strip()

        return response in ["AUTO", "TRIG"]


if __name__ == "__main__":
    rth1004 = RohdeSchwarz_Oscilloscope(simulate=False)
    rth1004.visa_address = "USB0::0x0AAD::0x012F::1317.5000K04/101102::INSTR"

    rth1004.open_connection()

    print(f"Model {rth1004.model}")

    rth1004.reset()

    rth1004.set_channel(chan=1, enabled=True)
    rth1004.set_channel(chan=2, enabled=True)
    rth1004.set_voltage_scale(chan=1, scale=0.002)
    rth1004.set_voltage_scale(chan=2, scale=0.2)
    # rth1004.set_voltage_offset(chan=1, offset=-1.5)  # Opposite direction to Keysight
    # rth1004.set_voltage_offset(chan=2, offset=+0.5)
    rth1004.set_timebase(0.001)

    rth1004.set_acquisition(32)

    print(rth1004.measure_voltage(2))

    print(rth1004.check_triggered())

    rth1004.set_trigger_type("EDGE")

    print(rth1004.check_triggered())

    time.sleep(1)

    print(f"Measurement {rth1004.measure_voltage(chan=1)}")

    rth1004.set_trigger_level(level=0.1, chan=1)

    rth1004.set_timebase(20e-9)

    rth1004.set_cursor_xy_source(chan=1, cursor=1)
    rth1004.set_cursor_position(cursor="X1", pos=0)

    ref_x = rth1004.read_cursor("X1")
    time.sleep(0.1)
    ref = rth1004.read_cursor("Y1")
    print(ref)

    rth1004.set_timebase_pos(0.001)
    time.sleep(0.1)

    rth1004.set_cursor_position(cursor="X1", pos=0.001)
    time.sleep(0.1)

    rth1004.adjust_cursor(target=ref)

    offset_x = rth1004.read_cursor("X1")

    print(f"TB Error {ref_x-offset_x+0.001}")

    input("Set voltage source to 0V")
    y1 = rth1004.read_cursor_avg()

    print(y1)

    input("Set voltage source to 1V")
    y2 = rth1004.read_cursor_avg()
    print(y2)
    print(f"Y Delta {y2-y1}")
