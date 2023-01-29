# sourcery skip: snake-case-functions
"""
# Quick script to control Meatest M142
"""

import pyvisa
import time
from pprint import pprint
from typing import List

VERSION = "A.00.11"


class M142_Simulate:
    """
    _summary_
    """

    def write(self, command: str) -> None:
        """
        write _summary_

        Args:
            command (str): _description_
        """
        print(f"M142: {command}")

    def query(self, command: str) -> str:
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """
        print(command)
        reply = "1.02"
        if command == "*OPC?":
            # time.sleep(1)
            reply = "1"
        elif command == "FAULT?":
            reply = "0"
        elif command == "OUT?":
            reply = "1,0"
        return reply + "\n"

    def close(self) -> None:
        """
        close _summary_

        Returns:
            _type_: _description_
        """

        pass


class M142:
    """
     _summary_

    Returns:
        _type_: _description_
    """

    visa_address = "GPIB0::06::INSTR"
    connected = False
    serial = ""
    manufacturer = ""
    model = ""
    timeout = 5000
    simulating = False

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
                self.instr = M142_Simulate()
                self.model = "M-142"
                self.manufacturer = "Meatest"
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
                not self.simulating and self.model.find("M-142") >= 0 or self.simulating
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

    def write(self, command: str) -> None:
        """
        write _summary_

        Args:
            command (str): _description_
        """
        self.instr.write(command)  # type: ignore

    def operate(self) -> None:
        """
        operate _summary_
        """
        self.instr.write("OUTP ON")  # type: ignore
        self.settle()

    def standby(self) -> None:
        """
        standby _summary_
        """
        self.instr.write("OUTP OFF")  # type: ignore
        self.instr.write("*OPC")  # type: ignore

    def reset(self) -> None:
        """
        reset _summary_
        """
        self.instr.write("*CLS;*RST")  # type: ignore
        self.instr.write("FUNC DC;VOLT 0V")  # type: ignore
        self.instr.write("OUTP:ISEL HIGH")  # type: ignore  # Turn off the coil if it is on

    def settle(self) -> None:
        """
        settle _summary_
        Wait until the output is settled
        """

        timeout_count = 0

        while timeout_count < 6:  # 30 seconds
            try:
                self.instr.query("*OPC?")  # type: ignore
                break
            except pyvisa.VisaIOError as tmo:
                if tmo.abbreviation == "VI_ERROR_TMO":
                    timeout_count += 1
                else:
                    pprint(tmo)

    def set_ext_sense(self, setting: bool) -> None:
        """
        set_ext_sense _summary_

        Args:
            setting (bool): _description_
        """
        if setting:
            self.instr.write("EXTSENSE ON")  # type: ignore
        else:
            self.instr.write("EXTSENSE OFF")  # type: ignore

    def set_voltage_dc(self, voltage: float) -> None:
        """
        set_voltage_dc _summary_

        Args:
            voltage (float): _description_
        """
        assert abs(voltage) <= 1000
        self.instr.write(  # type: ignore
            f"FUNC DC;VOLT {voltage} V"
        )  # Write 0 Hz in case was previous AC voltage

    def set_voltage_ac(self, voltage: float, frequency: float) -> None:
        """
        set_voltage_ac _summary_

        Args:
            voltage (float): _description_
            frequency (float): _description_
        """
        assert voltage <= 1000
        # todo frequency voltage trade off assert
        self.instr.write(f"FUNC SIN;VOLT {voltage} V;FREQ {frequency} Hz")  # type: ignore

    def set_2W_resistance(self, resistance: float) -> None:
        """
        set_2W_resistance _summary_

        Args:
            resistance (float): _description_
        """
        self.instr.write(f"RES {resistance} OHM")  # type: ignore

    def get_resistance(self) -> float:
        """
        get_resistance _summary_
        get the currently set resistance, from the last calibration

        Returns:
            float: _description_
        """

        # todo check in resistance mode. It will return whatever is set though
        return float(self.instr.query("RES?"))  # type: ignore

    def set_2W_compensation(self, resistance: float) -> None:
        """
        set_2W_compensation _summary_

        Args:
            resistance (float): _description_
        """
        self.set_2W_resistance(resistance)

    def set_current_dc(self, current: float) -> None:
        """
        set_current_dc _summary_

        Args:
            current (float): _description_
        """
        assert current <= 30

        self.instr.write(f"FUNC DC;CURR {current} A")  # type: ignore

    def set_current_ac(self, current: float, frequency: float) -> None:
        """
        set_current_ac _summary_

        Args:
            current (float): _description_
            frequency (float): _description_
        """
        assert current <= 30

        self.instr.write(f"FUNC SIN;CURR {current} A;FREQ {frequency} Hz")  # type: ignore

    def set_temperature(
        self, temperature: float, tc_type: str = "T", ref_junction: float = 78.3
    ) -> None:
        """
        set_temperature
        Set to output a temperature configured as a thermocouple

        Args:
            temperature (float): _description_
            tc_type (str, optional): _description_. Defaults to "T".
            ref_junction (float, optional): _description_. Defaults to 78.3.
        """

        self.instr.write("TEMP:UNITS C")  # type: ignore
        self.instr.write(f"TEMP:THER:TYPE {tc_type}")  # type: ignore
        self.instr.write(f"TEMP:THERM {temperature}")  # type: ignore

    def set_power(
        self,
        power: float,
        voltage: float | None = None,
        freq: float | None = None,
        phase: int | None = None,
    ) -> None:
        """
        set_power
        Set power in Watts

        Args:
            power (float): _description_
            voltage (float, optional): volts. Defaults to None.
        """

        if voltage:
            assert voltage <= 240

        assert power <= 4800

        freq_command = ""

        if freq:
            self.instr.write("FUNC SIN")  # type: ignore
            freq_command = f"FREQ {freq} Hz"

            # Have to have 0 degrees to set power

            self.instr.write("POWER:PHASE 0 LEAD")  # type: ignore

        else:
            self.instr.write("FUNC DC")  # type: ignore

        if voltage:
            self.instr.write(f"POWER:VOLT {voltage} V")  # type: ignore

        self.instr.write(f"POWER {power} W; {freq_command}")  # type: ignore

        if freq is not None:
            if phase:
                dirn = "LAG" if phase < 0 else "LEAD"
                self.instr.write(f"POWER:PHASE {phase} {dirn}")  # type: ignore
            else:
                self.instr.write("POWER:PHASE 0 LEAD")  # type: ignore


if __name__ == "__main__":
    m142 = M142(simulate=True)
    m142.visa_address = "GPIB0::06::INSTR"
    m142.open_connection()
    m142.is_connected()

    print(m142.get_id())
    print(
        f"Manufacturer: {m142.manufacturer}, Model: {m142.model}, Serial {m142.serial}"
    )

    m142.reset()
    m142.set_voltage_dc(10)
    m142.operate()
    time.sleep(2)
    m142.standby()

    m142.set_voltage_ac(5, 1000)
    m142.operate()
    time.sleep(2)
    m142.standby()

    m142.set_2W_resistance(10000000)
    m142.operate()
    print(m142.get_resistance())
    time.sleep(2)
    m142.standby()

    m142.set_power(100, voltage=12)  # DC
    time.sleep(2)
    m142.set_power(150, freq=50)
