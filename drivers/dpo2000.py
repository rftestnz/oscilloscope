"""
# Driver for Tek DPO2000


# DK Jan 23
"""

import pyvisa
import time
from random import random
from typing import List

VERSION = "A.00.00"


class DPO2000_Simulator:
    """
    _summary_
    """

    def close(self) -> None:
        """
        close _summary_
        """
        pass

    def write(self, command: str) -> None:
        """
        write _summary_

        Args:
            command (str): _description_
        """
        print(f"DPO2000 <- {command}")

    def read(self) -> float | str:
        """
        read _summary_

        Returns:
            float| str: _description_
        """
        return 0.5 + random()

    def query(self, command: str) -> str:
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """

        print(f"DRIVER_NAME <- {command}")

        if command == "*IDN?":
            return "Tektronix,DPO2024,MY_Simulated,B.00.00"

        return str(0.5 + random()) if command.startswith("READ") else ""


class DPO_2000:
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

    def __init__(self, simulate=False):
        self.simulating = simulate
        self.rm = pyvisa.ResourceManager()

    def __enter__(self):
        self.open_connection()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def open_connection(self) -> bool:
        """
        open_connection _summary_

        Returns:
            bool: _description_
        """
        try:
            if self.simulating:
                self.instr = DPO2000_Simulator
                self.model = "DPO3034"
                self.manufacturer = "Tektronix"
                self.serial = "666"
            else:
                self.instr = self.rm.open_resource(
                    self.visa_address, write_termination="\n"
                )
                self.instr.timeout = self.timeout
                # self.instr.control_ren(VI_GPIB_REN_ASSERT)  # type: ignore
                self.get_id()
            self.connected = True
        except Exception as ex:
            self.connected = False

        return self.connected

    def initialize(self) -> None:
        """
        initialize _summary_
        """
        self.open_connection()

    def close(self) -> None:
        """
        close _summary_
        """
        self.instr.close()  # type: ignore
        self.connected = False

    def is_connected(self) -> bool:
        """
        is_connected _summary_

        Returns:
            bool: _description_
        """
        return bool(
            self.open_connection()
            and (
                not self.simulating and self.model.find("3034") >= 0 or self.simulating
            )
        )

    def write(self, command: str) -> None:
        """
        write

        Args:
            command (str): [description]
        """

        attempts = 0

        while attempts < 3:
            try:
                self.instr.write(command)  # type: ignore
                break
            except pyvisa.VisaIOError as ex:
                time.sleep(1)
                attempts += 1

    def read(self) -> str:
        """
        read
        Read back from the unit

        Returns:
            str: [description]
        """

        attempts = 0

        ret: str = ""

        while attempts < 3:
            try:
                ret = self.instr.read()  # type: ignore
                break
            except pyvisa.VisaIOError as ex:
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

        attempts = 0
        ret = ""

        while attempts < 3:
            try:
                ret = self.instr.query(command)  # type: ignore
                break
            except pyvisa.VisaIOError as ex:
                time.sleep(1)
                attempts += 1

        return ret

    def read_query(self, command: str) -> float:
        """
        read_query
        send query command, return a float or -999 if invalid

        Args:
            command (str): _description_

        Returns:
            float: _description_
        """

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

    def get_id(self) -> List:
        """
        get_id _summary_

        Returns:
            List: _description_
        """
        try:
            self.instr.timeout = 2000  # type: ignore
            response = self.instr.query("*IDN?")  # type: ignore
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

    def set_channel(self, chan: int, enabled: bool) -> None:
        """
        set_channel
        Turn display of channel on or off

        Args:
            chan (int): _description_
            enabled (bool): _description_
        """

        state = "ON" if enabled else "OFF"

        self.write(f"SEL:CH{chan} {state}")

    def set_voltage_scale(self, chan: int, scale: float, probe_atten: int = 1) -> None:
        """
        set_voltage_scale _summary_

        Args:
            chan (int): _description_
            scale (float): _description_
        """

        self.write(f"CH{chan}:PRO:GAIN 1")
        self.write(f"CH{chan}:VOL {scale}")

    def set_voltage_offset(self, chan: int, offset: float) -> None:
        """
        set_voltage_offset _summary_

        Args:
            chan (int): _description_
            offset (float): _description_
        """

        # TODO Is it offset or pos?
        self.write(f"CH{chan}:POS {offset}")

    def set_timebase(self, timebase: float) -> None:
        """
        set_timebase _summary_

        Args:
            timebase (float): _description_
        """

        self.write(f"HOR:SCAL {timebase}")

    def set_timebase_pos(self, pos: float) -> None:
        """
        set_timebase_pos _summary_

        Args:
            pos (float): _description_
        """

        self.write(f"HOR:POS {pos}")

    def set_acquisition(self, num_samples: int) -> None:
        """
        set_acquisition _summary_

        Args:
            num_samples (int): _description_
        """

        self.write("ACQ:MODE AVE")
        self.write(f"ACQ:NUMAV {num_samples}")

    def set_trigger_mode(self, mode: str) -> None:
        """
        set_trigger_mode _summary_

        Args:
            mode (str): _description_
        """

        # self.write(f"TRIG:A:MODE {mode}")
        self.write("TRIG:SWE AUTO")

    def set_trigger_level(self, level: float, chan: int) -> None:
        """
        set_trigger_level
        Set the trigger level and source

        Args:
            level (float): _description_
            chan (int): _description_
        """

        self.write(f"TRIG:A:EDGE:SOUR CH{chan}")
        self.write(f"TRIG:A:LEV {level}")

    def measure_voltage(self, chan: int) -> float:
        """
        measure_voltage
        Return the average voltage

        Returns:
            float: _description_
        """

        self.write("MEASU:MEAS1:TYPE MEAN")

        self.write(f"MEAS:SOURCE CH{chan}")
        self.write(f"MEASU:MEAS{chan}:STATE ON")

        return self.read_query(f"MEASU:MEAS{chan}:MEAN?")

    def read_cursor(self, cursor: str) -> float:
        """

        read_cursor _summary_

        Enable cursor and read
        """

        self.write("MARK:MODE WAV")
        self.write(f"MARK:{cursor}:DISP ON")
        return self.read_query(f"MARK:{cursor}P?")

    def read_cursor_avg(self) -> float:
        """
        read_cursor_avg
        Read both of the Y cursors, and return the average

        Returns:
            float: _description_
        """

        self.write("MARK:MODE WAV")
        if self.family == DSOX_FAMILY.DSOX3000:
            self.write("MARK:Y1:DISP ON")
            self.write("MARK:Y2:DISP ON")
        y1 = self.read_query("MARK:Y1P?")
        y2 = self.read_query("MARK:Y1P?")

        return (y1 + y2) / 2

    def read_cursor_ydelta(self) -> float:
        """
        read_cursor_ydelta _summary_

        Returns:
            float: _description_
        """

        return self.read_query("MARK:YDEL?")

    def set_cursor_xy_source(self, chan: int, cursor: int) -> None:
        """
        set_cursor_xy_source
        Set the X1Y1 marker source

        Args:
            chan (int): _description_
        """

        self.write("CURS:FUNC WAV")
        self.write("CURS:MOD:TRACK")
        self.write(f"CURS:X{cursor}Y{cursor} CHAN{chan}")

    def set_cursor_position(self, cursor: str, pos: float) -> None:
        """
        set_cursor_position _summary_

        Args:
            cursor (str): _description_
            pos (float): _description_
        """

        self.write(f"MARK:{cursor}P {pos}")

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


if __name__ == "__main__":

    dpo2014 = DPO_2000()
    dpo2014.visa_address = "USB0::0x0699::0x0373::C010049::INSTR"

    dpo2014.open_connection()

    print(f"Model {dpo2014.model}")

    dpo2014.reset()

    dpo2014.set_channel(chan=1, enabled=True)
    dpo2014.set_channel(chan=2, enabled=True)
    dpo2014.set_voltage_scale(chan=1, scale=1)
    dpo2014.set_voltage_scale(chan=2, scale=0.2)
    dpo2014.set_voltage_offset(chan=1, offset=-1.5)  # Opposite direction to Keysight
    dpo2014.set_voltage_offset(chan=2, offset=+0.5)
    dpo2014.set_timebase(0.001)

    dpo2014.set_acquisition(64)

    dpo2014.set_trigger_mode("EDGE")

    time.sleep(1)

    print(f"Measurement {dpo2014.measure_voltage(chan=1)}")

    dpo2014.set_trigger_level(level=0.1, chan=1)

    dpo2014.set_timebase(20e-9)

    dpo2014.set_cursor_xy_source(chan=1, cursor=1)
    dpo2014.set_cursor_position(cursor="X1", pos=0)

    ref_x = dpo2014.read_cursor("X1")
    time.sleep(0.1)
    ref = dpo2014.read_cursor("Y1")
    print(ref)

    dpo2014.set_timebase_pos(0.001)
    time.sleep(0.1)

    dpo2014.set_cursor_position(cursor="X1", pos=0.001)
    time.sleep(0.1)

    dpo2014.adjust_cursor(target=ref)

    offset_x = dpo2014.read_cursor("X1")

    print(f"TB Error {ref_x-offset_x+0.001}")

    input("Set voltage source to 0V")
    y1 = dpo2014.read_cursor_avg()

    print(y1)

    input("Set voltage source to 1V")
    y2 = dpo2014.read_cursor_avg()
    print(y2)
    print(f"Y Delta {y2-y1}")
