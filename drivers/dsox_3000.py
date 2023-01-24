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


class DRIVER_NAME_Simulator:
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
            return "Keysight,DRIVER_NAME,MY_Simulated,B.00.00"

        return str(0.5 + random()) if command.startswith("READ") else ""


class DSOX_3000:
    """
     _summary_

    Returns:
        _type_: _description_
    """

    connected: bool = False
    visa_address: str = "GPIB0::26::INSTR"
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
                self.instr = DRIVER_NAME_Simulator
                self.model = "DRIVER_NAME"
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

    def set_voltage_scale(self, chan: int, scale: float) -> None:
        """
        set_voltage_scale _summary_

        Args:
            chan (int): _description_
            scale (float): _description_
        """

        self.write(f"CHAN{chan}:SCAL {scale}")

    def set_voltage_offset(self, chan: int, offset: float) -> None:
        """
        set_voltage_offset _summary_

        Args:
            chan (int): _description_
            offset (float): _description_
        """

        self.write(f"CHAN:OFFS {offset}")

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


if __name__ == "__main__":

    driver_name = DSOX_3000()
    driver_name.visa_address = "GPIB0::0::INSTR"

    driver_name.open_connection()
