"""
# Simple driver for Ks 33250A
# DK Jun 22
"""

import contextlib
import pyvisa
from pyvisa import InvalidSession, VisaIOError
from typing import List
import time

VERSION = "A.00.04"


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
                self.rm = pyvisa.ResourceManager()
                self.instr = self.rm.open_resource(
                    self.visa_address, write_termination="\n"
                )
                self.instr.timeout = self.timeout
                # self.instr.control_ren(VI_GPIB_REN_ASSERT)  # type: ignore
                self.get_id()
            self.connected = True
        except Exception as ex:
            print(ex)
            print(self.visa_address)
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

    def go_to_local(self) -> None:
        """
        go_to_local
        Set back to local operation
        """

        # The Keysight unit supports the SYST:LOC command, but the Agilent doesn't appear to. At least the RFTS ones do not.

        with contextlib.suppress(InvalidSession, VisaIOError, AttributeError):
            if not self.simulating:
                self.instr.control_ren(6)  # type: ignore

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

    def write(self, command: str) -> None:
        """
        write
        Fluke 5700A is unreliable sending using the pyvisa, so buffer it

        Args:
            command (str): [description]
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
        read
        Read back from the fluke

        Returns:
            str: [description]
        """

        attempts = 0

        ret = ""

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

    def get_id(self) -> List:
        """
        get_id _summary_

        Returns:
            List: _description_
        """
        try:
            self.instr.timeout = 2000  # type: ignore
            response = self.query("*IDN?")  # type: ignore
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

        self.write("VOLT:UNIT VRMS")
        self.write(f"FUNC SIN;FREQ {frequency}; VOLTAGE {amplitude} VRMS")

    def set_output_z(self, z: str) -> None:
        """
        Set the output impedance to 50 or MAX
        """

        assert z in {"50", "MAX"}

        self.write(f"OUTPUT:LOAD {z}")

    def set_pulse(
        self, period: float, pulse_width: float, amplitude: float, offset: float = 0
    ) -> None:
        """
        set_pulse _summary_

        Args:
            period (float): period in seconds
            pulse_width (float): pulse width in seconds
            amplitude (float): amplitude in volts peak
        """

        # The unit didn't like setting everything up in one line
        self.write("FUNC PULS;")
        self.write(f"PULS:PER {period}")
        self.write(f"PULS:WIDTH {pulse_width}")
        self.write("PULS:TRAN MIN")
        self.write("VOLT:UNIT VPP")
        self.write(f"VOLT {amplitude} VPP;VOLT:OFFSET {offset}")

    def enable_output(self, state: bool) -> None:
        """
        enable_output _summary_

        Args:
            state (bool): _description_
        """
        op = "ON" if state else "OFF"

        self.write(f"OUTP {op}")  # type: ignore


if __name__ == "__main__":
    with Ks33250A(simulate=False) as ks33250:
        ks33250.visa_address = "GPIB0::2::INSTR"
        ks33250.open_connection()

        print(ks33250.is_connected())
        print(ks33250.model)

        ks33250.set_sin(1560, 0.25)

        time.sleep(2)

        ks33250.set_pulse(period=1e-3, pulse_width=200e-6, amplitude=2)

        ks33250.enable_output(True)

        ks33250.go_to_local()
