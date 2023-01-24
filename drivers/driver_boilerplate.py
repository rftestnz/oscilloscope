# Boilerplate driver code for ...


# DK Jul 22

from enum import Enum
import pyvisa
from pyvisa.constants import VI_GPIB_REN_ASSERT
from pprint import pprint
import time
from random import random

replace DRIVER_NAME with item

VERSION = "A.00.00"


class DRIVER_NAME_Simulator:
    def close(self):
        pass

    def write(self, command: str):
        print(f"DRIVER_NAME <- {command}")

    def read(self):
        return 0.5 + random()

    def query(self, command: str) -> str:

        print(f"DRIVER_NAME <- {command}")

        if command == "*IDN?":
            return "Keysight,DRIVER_NAME,MY_Simulated,B.00.00"

        if command.startswith("READ"):
            return str(0.5 + random())

        return ""


class DRIVER_NAME:

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

    def initialize(self):
        self.open_connection()

    def close(self):
        self.instr.close()
        self.connected = False

    def is_connected(self):
        return bool(
            self.open_connection()
            and (
                not self.simulating
                and self.model.find("DRIVER_NAME") >= 0
                or self.simulating
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
        self.write("*RST")

    def get_id(self):
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




if __name__ == "__main__":

    driver_name = DRIVER_NAME()
    driver_name.visa_address = "GPIB0::0::INSTR"

    driver_name.open_connection()