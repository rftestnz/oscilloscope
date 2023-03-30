"""
    _summary_

Returns:
    _type_: _description_
"""

import pyvisa
from enum import Enum
import numpy as np
from pprint import pprint
import random
from datetime import datetime
from typing import Dict, List

from pyvisa.constants import VI_GPIB_REN_ASSERT

VERSION = "A.00.06"


class Ks3458A_Simulator:
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
        print(f"3458A: {command}")

    def query(self, command: str) -> str:
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """
        command = command.upper()
        print(f"3458A: {command}")
        if "SYST:ERR?" in command:
            return "0,0"
        result = 0.95 + random.random() / 10
        return str(result)

    def read(self) -> str:
        """
        read
        Do dummy reading

        Returns:
            str: _description_
        """

        return str(0.5 + random.random())


class Ks3458A_Function(Enum):
    """
    Ks3458A_Function _summary_

    Args:
        Enum (_type_): _description_
    """

    DCV = 0
    ACV = 1
    R2W = 2
    R4W = 3
    DCI = 4
    ACI = 5
    # There are more modes like frequency, period, conduction, diode, but they are rarely used


class Ks3458A_InputR(Enum):
    """
    Ks3458A_InputR _summary_

    Args:
        Enum (_type_): _description_
    """

    M10 = 0  ## 10 M Ohm
    G10 = 0  ## 10 G Ohm


class Ks3458A_QuestionableData(Enum):
    """
    Ks3458A_QuestionableData _summary_

    Args:
        Enum (_type_): _description_
    """

    VOLTAGE_OVERLOAD = 1
    CURRENT_OVERLOAD = 2
    OHMS_OVERLOAD = 512


class Ks3458A_StandardEvent(Enum):
    """
    Ks3458A_StandardEvent _summary_

    Args:
        Enum (_type_): _description_
    """

    OPERATION_COMPLETE = 1
    QUERY_ERROR = 4
    DEVICE_ERROR = 8
    EXECUTION_ERROR = 16
    COMMAND_ERROR = 32
    POWER_ON = 128


class Ks3458A_StatusByte(Enum):
    """
    Ks3458A_StatusByte _summary_

    Args:
        Enum (_type_): _description_
    """

    QUESTIONABLE_DATA = 8
    MESSAGE_AVAILABLE = 16
    STANDARD_EVENT = 16
    REQUEST_SERVICE = 64


class Ks3458A_ACV_CONFIG(Enum):
    """
    Ks3458A_ACV_CONFIG _summary_

    Args:
        Enum (_type_): _description_
    """

    DEFAULT = 0
    BEST = 1


class Ks3458A:
    """
     _summary_

    Returns:
        _type_: _description_
    """

    visa_address = "GPIB0::23::INSTR"
    connected = False
    model = ""
    current_mode = None
    timeout = 15000
    simulating = False
    option001 = False

    def __init__(self, simulate=False) -> None:
        self.simulating = simulate
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
                self.instr = Ks3458A_Simulator()
                self.model = "3458A"

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
        self.rm.close
        self.connected = False

    def __wait_until_ready(self) -> None:
        # Wait until the unit is ready

        # doesn't support the OPC

        pass

    def is_connected(self) -> bool:
        """
        is_connected _summary_

        Returns:
            bool: _description_
        """
        return bool(
            self.open_connection()
            and (
                not self.simulating and self.model.find("3458") >= 0 or self.simulating
            )
        )

    def get_id(self) -> str:
        """
        get_id _summary_

        Returns:
            List: _description_
        """
        try:
            self.instr.timeout = 2000  # type: ignore
            self.model = self.instr.query("ID?")  # type: ignore

        except pyvisa.VisaIOError:
            self.model = ""

        self.instr.timeout = self.timeout  # type: ignore

        return self.model

    def reset(self) -> None:
        """
        reset _summary_
        """
        self.__wait_until_ready()
        self.instr.write("RESET")  # type: ignore

    def set_input_resistance(self, inputr: Ks3458A_InputR) -> None:
        """
        set_input_resistance _summary_

        Args:
            inputr (Ks3458A_InputR): _description_
        """
        # 10 GOhm only available up to 10 V
        self.__wait_until_ready()

        if inputr == Ks3458A_InputR.M10:
            self.instr.write("FIXEDZ OFF;")  # type: ignore
        else:
            self.instr.write("FIXEDZ ON;")  # type: ignore

    def set_function(self, function: Ks3458A_Function) -> None:
        """
        set_function _summary_

        Args:
            function (Ks3458A_Function): _description_
        """
        # assert int(function) <= int(Ks3458A_Function.ACI)

        self.instr.timeout = self.timeout  # type: ignore
        self.__wait_until_ready()
        if function == Ks3458A_Function.DCV:
            self.instr.write("DCV")  # type: ignore
        elif function == Ks3458A_Function.ACV:
            self.instr.write("ACV")  # type: ignore
            self.instr.timeout = self.timeout * 2  # type: ignore # longer timeout
        elif function == Ks3458A_Function.R2W:
            self.instr.write("OHM")  # type: ignore
        elif function == Ks3458A_Function.R4W:
            self.instr.write("OHMF")  # type: ignore
        elif function == Ks3458A_Function.DCI:
            self.instr.write("DCI")  # type: ignore
        elif function == Ks3458A_Function.ACI:
            self.instr.write("ACI")  # type: ignore

        # self.check_error()
        self.current_mode = function

    def configure_dc_nplc(self, nplc: int) -> None:
        """
        configure_dc_nplc
        Set the number power line cycles to integrate signal over

        Args:
            nplc (int): _description_
        """
        self.instr.write(f"NPLC {nplc}")  # type: ignore

    def configure_acv(self, cfg: Ks3458A_ACV_CONFIG) -> None:
        """
        configure_acv _summary_

        Args:
            cfg (Ks3458A_ACV_CONFIG): _description_
        """
        if cfg == Ks3458A_ACV_CONFIG.BEST:
            self.instr.write("SETACV SYNC")  # type: ignore
            self.instr.write("LFILTER ON")  # type: ignore
            self.instr.write("RES 0.002")  # type: ignore

    def measure(
        self, function: Ks3458A_Function, number_readings: int = 1
    ) -> Dict[float, float]:
        # sourcery skip: extract-method
        """
        measure _summary_

        Args:
            function (Ks3458A_Function): _description_
            number_readings (int, optional): _description_. Defaults to 1.

        Returns:
            Dict[float, float]: _description_
        """
        if function != self.current_mode:
            self.set_function(function)

        self.instr.write(f"NRDGS {number_readings}")  # type: ignore
        self.instr.write("TARM SGL")  # type: ignore

        readings = []

        try:
            for _ in range(number_readings):

                reply = self.instr.read().strip()  # type: ignore
                readings.append(float(reply))

            if self.simulating:
                # Create an array
                reply = [
                    str(0.95 + random.random() / 10) for _ in range(number_readings)
                ]
            rdgs = np.array(readings)

            result = {"Average": np.average(rdgs), "StdDev": np.std(rdgs)}

            if self.simulating:
                pprint(f"3458A: {result}")

            # put it back
            self.instr.write("NRDGS 1")  # type: ignore

            return result  # type: ignore
        except pyvisa.VisaIOError as ex:
            print("Error reading")
            # self.check_error()
            return {"Average": 0.0, "StdDev": 0.0}  # type: ignore

    def measure_sampling(
        self, sample_period: float, number_samples: int, resolution: float = 4.5
    ) -> List:
        """
        measure_sampling
        Use the built in sampling functions to take readings at specified interval

        Args:
            sample_period (_type_): _description_
            number_samples (_type_): _description_

        Returns:
            List: _description_
        """

        # standard memory is 20k, so we can only store 10240 readings in SINT mode, or 5120 in DINT mode

        assert resolution in {4.5, 5.5}, "Resolution must be 4.5 or 5.5"

        # Standard unit has 20kb, option 001 148 kb
        # The command to get options is OPT?, but none of the units RFTS have support the command

        if self.option001:
            max_readings = 75776 if resolution == 4.5 else 37888
        else:
            max_readings = 10240 if resolution == 4.5 else 5120

        assert number_samples <= max_readings, "Number readings exceeds 3458A memory"

        tmo = self.timeout

        self.instr.write("PRESET DIG")  # type: ignore
        self.instr.write("TRIG AUTO")  # type: ignore   # Suspend triggering
        self.instr.write("DSDC 10")  # type: ignore   # Direct sampling DC
        self.instr.write("APER 1.4E-6")  # type: ignore
        self.instr.write("MEM FIFO")  # type: ignore   # Have to use the memory, as using direct transfer have to be able to transfer at >134kb/s
        self.instr.write("DISP OFF")  # type: ignore

        # The SWEEP cpommand is supposed to be the easier way to take specified number of samples at interval, but could
        # not figure out how to get the readings back in python
        # It was probably to do with ability to transfer data fast enough. Turning MEM FIFO helps
        # But as the command is only a combination of NRDGS and TIMER, it doesn't matter

        self.instr.write(f"TIMER {sample_period}")  # type: ignore

        self.instr.write(f"NRDGS {number_samples}, TIMER")  # type: ignore

        # self.instr.write("DELAY 0")  # type: ignore
        self.instr.write("AZERO OFF")  # type: ignore

        if resolution == 4.5:
            format_out = "SINT"
            byte_size = 2
        else:
            format_out = "DINT"
            byte_size = 4

        # Set the measurement and output formats to double integer (4 bytes signed) or single integer
        self.instr.write(f"MFORMAT {format_out}")  # type: ignore
        self.instr.write(f"OFORMAT {format_out}")  # type: ignore

        # Now start the samples

        self.instr.write("TARM SGL")  # type: ignore

        # And read everything back

        self.instr.timeout = 2000  # type: ignore

        # DINT format is 4 bytes per reading
        raw_data = self.instr.read_bytes(number_samples * byte_size)  # type: ignore

        scale = float(self.instr.query("ISCALE?").strip())  # type: ignore

        self.instr.write("DISP ON")  # type: ignore

        # We have a binary dump, convert into voltages
        readings = []

        for index, _ in enumerate(range(number_samples)):
            bt = raw_data[index * byte_size : index * byte_size + byte_size]
            rdg = int.from_bytes(bt, "big", signed=True)
            readings.append(rdg * scale)

        self.instr.timeout = tmo  # type: ignore

        return readings


if __name__ == "__main__":
    # Simple testing of the class

    simulating = True

    ks3458 = Ks3458A(simulate=simulating)
    ks3458.visa_address = "GPIB0::23::INSTR"
    ks3458.open_connection()

    ks3458.reset()
    print(ks3458.get_id())
    ks3458.option001 = False

    ks3458.configure_dc_nplc(5)

    pprint(ks3458.measure(Ks3458A_Function.DCV, 10))

    if not simulating:
        start = datetime.now()
        samples = ks3458.measure_sampling(10e-6, number_samples=5000, resolution=5.5)
        end = datetime.now()

        print(f"Time for sampling {(end - start).microseconds/1000000}")

        with open("samples.csv", "w") as outfile:
            for rdg in samples:
                outfile.write(f"{rdg}\n")

    ks3458.reset()

    ks3458.close()
