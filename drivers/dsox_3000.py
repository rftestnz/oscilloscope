# Boilerplate driver code for ...


# DK Jul 22

from enum import Enum
import pyvisa
from pyvisa.constants import VI_GPIB_REN_ASSERT
from pprint import pprint
import time
from random import random
from typing import List

VERSION = "A.00.00"


class DSOX_FAMILY(Enum):
    """
    DSOX_FAMILY
    Enum to cope with minor differences in commands

    Args:
        Enum (_type_): _description_
    """

    DSOX1000 = 1
    DSOX2000 = 2
    DSOX3000 = 3


class DSOX3000_Simulator:
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
        print(f"DRIVER_NAME <- {command}")

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
            return "Keysight,DSOX3000,MY_Simulated,B.00.00"

        return str(0.5 + random()) if command.startswith("READ") else ""


class DSOX_3000:
    """
     _summary_

    Returns:
        _type_: _description_
    """

    connected: bool = False
    visa_address: str = "USB0::0x2A8D::0x1797::CN59296333::INSTR"
    model = ""
    manufacturer = ""
    serial = ""
    family = DSOX_FAMILY.DSOX3000
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
                self.instr = DSOX3000_Simulator
                self.model = "DSOX3034T"
                self.manufacturer = "Keysight"
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
                model_type = self.model.split(" ")
                if len(model_type) > 1:
                    family = model_type[1][0]
                    if family == "1":
                        self.family = DSOX_FAMILY.DSOX1000
                    else:
                        self.family = (
                            DSOX_FAMILY.DSOX2000
                            if family == "2"
                            else DSOX_FAMILY.DSOX3000
                        )
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

        self.write(f"CHAN{chan}:DISP {state}")

    def set_voltage_scale(self, chan: int, scale: float, probe: int = 1) -> None:
        """
        set_voltage_scale _summary_

        Args:
            chan (int): _description_
            scale (float): _description_
        """

        self.write(f"CHAN{chan}:PROB {probe}")  # Set before the scale
        self.write(f"CHAN{chan}:SCAL {scale}")

    def set_voltage_offset(self, chan: int, offset: float) -> None:
        """
        set_voltage_offset _summary_

        Args:
            chan (int): _description_
            offset (float): _description_
        """

        self.write(f"CHAN{chan}:OFFS {offset}")

    def set_timebase(self, timebase: float) -> None:
        """
        set_timebase _summary_

        Args:
            timebase (float): _description_
        """

        self.write(f"TIM:SCAL {timebase}")

    def set_timebase_pos(self, pos: float) -> None:
        """
        set_timebase_pos _summary_

        Args:
            pos (float): _description_
        """

        self.write(f"TIM:POS {pos}")

    def set_acquisition(self, num_samples: int) -> None:
        """
        set_acquisition _summary_

        Args:
            num_samples (int): _description_
        """

        self.write("ACQ:TYPE AVER")
        self.write(f"ACQ:COUNT {num_samples}")

    def set_trigger_mode(self, mode: str) -> None:
        """
        set_trigger_mode _summary_

        Args:
            mode (str): _description_
        """

        self.write(f"TRIG:MODE {mode}")
        self.write("TRIG:SWE AUTO")

    def set_trigger_level(self, level: float, chan: int) -> None:
        """
        set_trigger_level
        Set the trigger level and source

        Args:
            level (float): _description_
            chan (int): _description_
        """

        self.write(f"TRIG:EDGE:SOUR chan{chan}")
        self.write(f"TRIG:EDGE:LEV {level}")

    def measure_voltage(self, chan: int) -> float:
        """
        measure_voltage
        Return the average voltage

        Returns:
            float: _description_
        """

        self.write(f"MEAS:SOURCE CHAN{chan}")

        return self.read_query("MEAS:VAV?")

    def read_cursor(self, cursor: int) -> float:
        """

        set_cursor_y1 _summary_

        Enable cursor Y1, set to center of screen
        """

        self.write("MARK:MODE WAV")
        self.write(f"MARK:Y{cursor}:DISP ON")
        return self.read_query(f"MARK:Y{cursor}P?")

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


if __name__ == "__main__":

    dsox3034t = DSOX_3000()
    dsox3034t.visa_address = "USB0::0x2A8D::0x1797::CN59296333::INSTR"

    dsox3034t.open_connection()

    print(f"Model {dsox3034t.model}")

    dsox3034t.reset()

    dsox3034t.set_channel(chan=1, enabled=True)
    dsox3034t.set_channel(chan=2, enabled=True)
    dsox3034t.set_voltage_scale(chan=1, scale=1)
    dsox3034t.set_voltage_scale(chan=2, scale=0.2)
    dsox3034t.set_voltage_offset(chan=1, offset=3.5)
    dsox3034t.set_voltage_offset(chan=2, offset=-0.5)
    dsox3034t.set_timebase(0.001)

    dsox3034t.set_acquisition(64)

    dsox3034t.set_trigger_mode("EDGE")

    time.sleep(1)

    print(f"Measurement {dsox3034t.measure_voltage(chan=1)}")

    dsox3034t.set_trigger_level(level=0.5, chan=1)

    dsox3034t.set_timebase_pos(0.001)
    dsox3034t.set_timebase(20e-9)

    input("Set voltage source to 0V")
    y1 = dsox3034t.read_cursor_avg()

    print(y1)

    input("Set voltage source to 1V")
    y2 = dsox3034t.read_cursor_avg()
    print(y2)
    print(f"Y Delta {y2-y1}")
