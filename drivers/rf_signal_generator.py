# sourcery skip: snake-case-functions
"""
# Driver for signal generator
# Target E4438C, but should work with other Keysight generators
# DK Feb 2022
"""

VERSION = "A.00.03"

from enum import Enum
import pyvisa
from pprint import pprint
import time
from typing import List


class RF_Sig_Gen_Simulator:
    """ """

    def close(self) -> None:
        pass

    def write(self, command: str) -> None:
        """
        write _summary_

        Args:
            command (str): _description_
        """
        print(f"E4438C <- {command}")

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

        print(f"E4438C <- {command}")

        return "Keysight,E4438C,MY_Simulated,B.00.00" if command == "*IDN?" else ""


class RF_Signal_Generator:
    """
     _summary_

    Returns:
        _type_: _description_
    """

    connected: bool = False
    visa_address: str = "GPIB0::18::INSTR"
    timeout: int = 5000
    simulating: bool = False

    def __init__(self, simulate: bool = False):
        self.simulating = simulate
        self.rm = pyvisa.ResourceManager()

    def __enter__(self):
        self.open_connection()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def open_connection(self) -> bool:
        """
        open_connection

        Returns:
            bool: _description_
        """
        try:
            if self.simulating:
                self.instr = RF_Sig_Gen_Simulator()
                self.model = "E4438C"
                self.manufacturer = "Agilent"
                self.serial = "0"
            else:
                self.instr = self.rm.open_resource(
                    self.visa_address, write_termination="\n"
                )
                self.instr.timeout = self.timeout
                # self.instr.control_ren(VI_GPIB_REN_ASSERT)
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
                and self.model in {"E4438C", "E8257D", "N5183A"}
                or self.simulating
            )
        )

    def write(self, command: str) -> None:
        """
        write
        E4438C is unreliable sending using the pyvisa, so buffer it

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

        ret = ""

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
        self.write("*RST")

    def get_id(self) -> List:
        """
        get_id _summary_

        Returns:
            List: _description_
        """

        try:
            self.instr.timeout = 2000  # type: ignore
            response = self.query("*IDN?")
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

    def set_frequency(self, freq: float) -> None:
        """
        set_frequency
        Set the frequency

        Args:
            freq (float): Hz
        """

        self.write(f"SOUR:FREQ {freq} Hz")

    def set_frequency_MHz(self, freq: float) -> None:
        """
        set_frequency_MHz
        Set the frequency in MHz

        Args:
            freq (float): MHz
        """

        self.write(f"SOUR:FREQ {freq} MHz")

    def set_level(self, level: float) -> None:
        """
        set_level
        Set the level

        Args:
            level (float): dBm
        """

        self.write(f"POW {level} dBm")

    def set_output_state(self, state: bool) -> None:
        """
        set_output_state
        Enable or disable the output

        Args:
            state (bool): _description_
        """

        st = "ON" if state else "OFF"

        self.write(f"OUTP:STATE {st}")

    def set_modulation_state(self, state: bool) -> None:
        """
        set_modulation_state _summary_

        Args:
            state (bool): _description_
        """

        st = "ON" if state else "OFF"

        self.write(f"OUTP:MOD:STAT {st}")


if __name__ == "__main__":

    with RF_Signal_Generator(simulate=True) as e4438c:
        e4438c.visa_address = "GPIB0::24::INSTR"
        e4438c.open_connection()

        print(e4438c.get_id())

        e4438c.reset()

        e4438c.set_frequency_MHz(100)
        e4438c.set_level(-10.5)
        e4438c.set_output_state(True)
