"""
# Simple driver for Ks 33250A
# DK Jun 22
"""

from enum import Enum
import pyvisa
from pyvisa.constants import VI_GPIB_REN_ASSERT
from pprint import pprint
import time
from random import random
from typing import List

VERSION = "A.00.03"


class Ks3250A_Simulator:
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
        print(f"Ks33250A <- {command}")

    def read(self) -> str:
        """
        read _summary_

        Returns:
            str: _description_
        """
        return ""

    def query(self, command: str) -> str:
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """

        print(f"Ks33250A <- {command}")

        return "Keysight,33250A,MY_Simulated,B.00.00" if command == "*IDN?" else ""


class Ks33250A:
    """
     _summary_

    Returns:
        _type_: _description_
    """

    visa_address = "GPIB0::2::INSTR"
    connected = False
    model = ""
    manufacturer = ""
    serial = ""
    timeout = 5000

    def __init__(self, simulate=False):
        self.simulating = simulate
        self.rm = pyvisa.ResourceManager()
        self.open_connection()

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
                self.instr = Ks3250A_Simulator()
                self.model = "33250A"
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
        self.instr.close()
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
                not self.simulating
                and self.model.find("33250A") >= 0
                or self.simulating
            )
        )

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

    def set_sin(self, frequency: float, amplitude: float) -> None:
        """
        set_sin
        Set sine wave output

        Args:
            frequency (float): frequency in Hz
            amplitude (float): amplitude VRMS
        """
        self.instr.write(  # type: ignore
            f"FUNC SIN;FREQ {frequency};VOLT:UNIT VRMS; VOLTAGE {amplitude}"
        )

    def enable_output(self, state: bool) -> None:
        """
        enable_output _summary_

        Args:
            state (bool): _description_
        """
        op = "ON" if state else "OFF"

        self.instr.write(f"OUTP {op}")  # type: ignore


if __name__ == "__main__":

    with Ks33250A(simulate=True) as ks33250:
        ks33250.visa_address = "GPIB0::9::INSTR"
        ks33250.open_connection()

        print(ks33250.is_connected())
        print(ks33250.model)

        ks33250.set_sin(1560, 0.25)

        ks33250.enable_output(True)
