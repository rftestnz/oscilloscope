"""
Basic driver to get ID of equipment
DK Feb 23
"""

import pyvisa
from typing import List


class SCPI_ID:
    """
    _summary_
    """

    def __init__(self, address: str) -> None:
        self.rm = pyvisa.ResourceManager()
        self.visa_address = address

    def __enter__(self):
        self.open_connection()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def open_connection(self) -> None:
        """
        open_connection _summary_
        """
        self.instr = self.rm.open_resource(self.visa_address, write_termination="\n")
        self.instr.timeout = 2000

    def close(self) -> None:
        """
        close _summary_
        """

        self.rm.close()

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

        return response.split(",")

    def get_manufacturer(self) -> str:
        """
        get_manufacturer
        request the manufacturer
        keysight and agilent are returned as keysight

        Returns:
            str: _description_
        """

        self.get_id()

        manufacturer = self.manufacturer.upper()
        if "KEYSIGHT" in manufacturer or "AGILENT" in manufacturer:
            manufacturer = "KEYSIGHT"

        return manufacturer


if __name__ == "__main__":
    with SCPI_ID(address="USB0::0x0699::0x03A3::C044602::INSTR") as scpi:
        print(scpi.get_id())

        print(scpi.get_manufacturer())
