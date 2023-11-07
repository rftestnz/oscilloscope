"""
# Quick script to control 5700A
"""


import contextlib
from enum import Enum
import pyvisa
from pyvisa import InvalidSession
import time
from pprint import pprint
from typing import List
from pyvisa.constants import VI_GPIB_REN_ASSERT

VERSION = "A.00.07"


class Fluke5700AOutput(Enum):
    """
    Fluke5700AOutput _summary_

    Args:
        Enum (_type_): _description_
    """

    NORMAL = 0
    AUX = 1
    AMPLIFER = 2


class Fluke5700A_Simulator:
    """
    Simple class to divert the command write function to the stdout
    For query, return a default reply
    """

    def write(self, command: str) -> None:
        """
        write _summary_

        Args:
            command (str): _description_
        """

        print(f"5700A: {command}")

    def query(self, command: str) -> str:
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """

        print(f"5700A: {command}")
        reply = "1.02"
        if command == "*OPC?":
            time.sleep(1)
            reply = "1"
        elif command == "FAULT?":
            reply = "0"
        elif command == "ISR?":
            # Check for settled
            reply = "4096"

        elif command == "OUT?":
            reply = "1,0"
        return reply + "\n"

    def close(self) -> None:
        """
        close _summary_
        """
        pass


class Fluke5700A:
    """
     _summary_

    Returns:
        _type_: _description_
    """

    visa_address: str = "GPIB0::06::INSTR"
    connected: bool = False
    serial: str = ""
    manufacturer: str = ""
    model: str = ""
    simulating: bool = False
    boost: bool = False
    timeout = 5000
    settle_timeout: int = 15

    def __init__(self, simulate=False) -> None:
        self.simulating = simulate
        if not simulate:
            self.rm = pyvisa.ResourceManager()
            self.open_connection()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()

    def open_connection(self) -> bool:
        """
        open_connection _summary_

        Returns:
            bool: _description_
        """
        try:
            if self.simulating:
                self.instr = Fluke5700A_Simulator()
                self.model = "5700A"
                self.manufacturer = "Fluke"
                self.serial = "666"
            else:
                self.instr = self.rm.open_resource(
                    self.visa_address, write_termination="\n"
                )
                self.instr.timeout = self.timeout
                self.instr.control_ren(VI_GPIB_REN_ASSERT)  # type: ignore
                self.get_id()
            self.connected = True
        except Exception:
            self.connected = False

        return self.connected

    def close(self) -> None:
        """
        close _summary_
        """
        self.instr.close()
        self.connected = False

    def is_connected(self) -> bool:
        """
        is_connected
        Check if the unit is connected

        Returns:
            bool: _description_
        """
        return bool(
            self.open_connection()
            and (
                not self.simulating
                and self.model in {"5700A", "5720A", "5730A"}
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
            _type_: _description_
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

    def operate(self) -> None:
        """
        operate _summary_
        """
        self.write("OPER")
        self.settle()

    def standby(self) -> None:
        """
        standby _summary_
        """
        self.write("STBY")
        self.write("*OPC")

    def reset(self) -> None:
        """
        reset _summary_
        """
        self.write("*RST;*CLS;*WAI")
        time.sleep(1)

    def go_to_local(self) -> None:
        """
        go_to_local
        Set back to local operation
        """

        with contextlib.suppress(Exception):
            if not self.simulating:
                self.instr.control_ren(6)  # type: ignore

    def settle(self) -> None:
        """
        Wait until the output is settled
        """

        if self.simulating:
            return

        timeout_count = 0
        self.write("*OPC")

        while timeout_count < self.settle_timeout:
            try:
                status = self.query("ISR?")
                if int(status) & 0b0001_0000_0000_0000 > 0:  # bit 12 is settled
                    break
                time.sleep(1)
                timeout_count += 1
            except pyvisa.VisaIOError as tmo:
                if tmo.abbreviation == "VI_ERROR_TMO":
                    timeout_count += 1
                else:
                    pprint(tmo)

        # If it takes a while to settle, likely there will be interrupted query errors. Clear the queue

        self.get_faults()

    def get_faults(self) -> None:
        """
        # Empty the queue of bad commands
        """

        fault = 1

        while fault:
            try:
                fault = int(self.query("FAULT?").strip())
                if fault:
                    explanation = self.query(f"EXPLAIN? {fault}")
                    print(f"{fault}; {explanation}")
            except Exception as ex:
                # Can timeout here, so just ignore faults
                print(f"Exception during 5700A fault request {ex}")
                fault = 0
                time.sleep(3)

    def set_ext_sense(self, setting: bool) -> None:
        """
        set_ext_sense _summary_

        Args:
            setting (bool): _description_
        """

        cmd = "ON" if setting else "OFF"
        self.write(f"EXTSENSE {cmd}")

    def set_voltage_dc(self, voltage: float) -> None:
        """
        set_voltage_dc _summary_

        Args:
            voltage (float): _description_
        """
        assert abs(voltage) <= 1000
        self.boost = False
        self.write(  # type: ignore
            f"OUT {voltage} V, 0 Hz"
        )  # Write 0 Hz in case was previous AC voltage

    def set_voltage_ac(self, voltage: float, frequency: float) -> None:
        """
        set_voltage_ac _summary_

        Args:
            voltage (float): _description_
            frequency (float): _description_
        """
        assert voltage <= 1000
        self.boost = False
        # todo frequency voltage trade off assert
        self.write(f"OUT {voltage} V, {frequency} Hz")

    def set_resistance_value(self, resistance: float) -> None:
        """
        set_resistance_value _summary_

        Args:
            resistance (float): _description_
        """
        # Can only be a decade or 1.9
        assert resistance in {
            0,
            1,
            1.9,
            10,
            19,
            100,
            190,
            1000,
            1900,
            10000,
            19000,
            100000,
            190000,
            1000000,
            1900000,
            10000000,
            19000000,
            100000000,
        }
        self.boost = False
        self.write(f"OUT {resistance} OHM")

        time.sleep(1)

    def set_2w_resistance(self, resistance: float) -> None:
        """
        set_2w_resistance _summary_

        Args:
            resistance (float): _description_
        """

        self.write(
            "RCOMP OFF"
        )  # turn off first else it will error if the current resistance is too high

        self.set_resistance_value(resistance=resistance)
        self.boost = False

        if resistance <= 19000:
            self.write("RCOMP ON")
        else:
            self.write("RCOMP OFF")

    def set_4w_resistance(self, resistance: float) -> None:
        """
        set_4w_resistance
        Set up for 4 wire Ohms. Same as 2 wire, but turn on external sense, and turn off 2 wire comp

        Args:
            resistance ([type]): [description]
        """

        self.write("RCOMP OFF")
        self.set_resistance_value(resistance=resistance)
        self.set_ext_sense(True)

    def get_resistance(self) -> float:
        """
        # get the currently set resistance, from the last calibration
        # todo check in resistance mode. It will return whatever is set though
        """

        try_count = 0
        res = 0

        while try_count < 3:
            try:
                setting = self.query("OUT?")
                fields = setting.split(",")
                res = float(fields[0])
                break
            except Exception:
                try_count += 1

        return res

    def set_current_output(self, output: Fluke5700AOutput) -> None:
        """
        set_current_output _summary_

        Args:
            output (Fluke5700AOutput): _description_
        """
        if output == Fluke5700AOutput.AMPLIFER:
            self.write("CUR_POST IB5725;BOOST ON")
            self.boost = True
        else:
            self.write("CUR_POST NORMAL;BOOST OFF")
            self.boost = False

    def set_current_dc(self, current: float, boost: bool | None = None) -> None:
        """
        set_current_dc _summary_

        Args:
            current (float): _description_
            boost (bool | None, optional): _description_. Defaults to None.
        """
        assert abs(current) <= 10

        # set to 0 first, in case using other output previously
        self.write("OUT 0 A, 0 Hz")

        if abs(current) > 1 or boost:
            self.set_current_output(Fluke5700AOutput.AMPLIFER)
        else:
            self.set_current_output(Fluke5700AOutput.NORMAL)
        self.write(f"OUT {current} A, 0 Hz")

    def set_current_ac(
        self, current: float, frequency: float, boost: bool | None = None
    ) -> None:
        """
        set_current_ac _summary_

        Args:
            current (float): _description_
            frequency (float): _description_
            boost (bool | None, optional): _description_. Defaults to None.
        """
        assert current <= 10

        # set to 0 first, in case using other output previously
        self.write("OUT 0 A, 0 Hz")

        if current > 1.99 or boost:
            self.set_current_output(Fluke5700AOutput.AMPLIFER)
        else:
            self.set_current_output(Fluke5700AOutput.NORMAL)
        self.write(f"OUT {current} A, {frequency} Hz")

    def test_status(self) -> None:
        """
        test_status _summary_
        """
        self.write("OUT 1V, 20HZ;OPER")

        for _ in range(20):
            status = self.query("ISR?")  # type: ignore
            print(status, int(status) & 0b0001_0000_0000_0000)
            time.sleep(1)

        self.standby()

    def set_phase_lock(self, state: bool) -> None:
        """
        set_phase_lock _summary_

        Args:
            state (bool): _description_
        """

        cmd = "ON" if state else "OFF"

        self.write(f"PHASELCK {cmd}")  # type: ignore


if __name__ == "__main__":
    fl5700a = Fluke5700A(simulate=True)
    fl5700a.visa_address = "GPIB0::06::INSTR"
    fl5700a.open_connection()

    print(fl5700a.get_id())
    print(
        f"Manufacturer: {fl5700a.manufacturer}, Model: {fl5700a.model}, Serial {fl5700a.serial}"
    )

    fl5700a.reset()
    # fl5700a.get_faults()

    skip = True

    if not skip:
        fl5700a.test_status()

        fl5700a.set_voltage_dc(10)
        fl5700a.operate()
        time.sleep(2)
        fl5700a.standby()
        fl5700a.get_faults()

        fl5700a.set_voltage_ac(5, 1000)
        fl5700a.operate()
        time.sleep(2)
        fl5700a.standby()
        fl5700a.get_faults()

        fl5700a.set_2w_resistance(10000000)
        fl5700a.operate()
        print(fl5700a.get_resistance())
        time.sleep(2)
        fl5700a.standby()

        fl5700a.get_faults()

    print("Testing current")
    fl5700a.set_current_dc(0.1)
    fl5700a.operate()
    time.sleep(1)
    fl5700a.get_faults()
    fl5700a.set_current_dc(1)
    fl5700a.operate()
    time.sleep(1)
    fl5700a.get_faults()

    fl5700a.reset()
