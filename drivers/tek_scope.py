"""
# Driver for Tek DPO2000


# DK Jan 23
"""

import contextlib
import pyvisa
import time
from random import random
from typing import List
import numpy as np
from struct import unpack
from enum import Enum

try:
    from drivers.base_scope_driver import ScopeDriver, Scope_Simulator
except ModuleNotFoundError:
    from base_scope_driver import ScopeDriver, Scope_Simulator

VERSION = "A.00.01"


class Tek_Acq_Mode(Enum):
    """
    Tek_Acq_Mode
    The acquisition modes for the MSO4-5-6

    Args:
        Enum (_type_): _description_
    """

    SAMPLE = 1
    PEAK = 2
    HIRES = 3
    AVERAGE = 4
    ENVELOPE = 5


class Tektronix_Oscilloscope(ScopeDriver):
    """
     _summary_

    Returns:
        _type_: _description_
    """

    connected: bool = False
    visa_address: str = "USB0::0x0699::0x0373::C010049::INSTR"
    model = ""
    manufacturer = ""
    serial = ""
    timeout = 5000
    num_channels: int = 4
    keysight: bool = False

    def __init__(self, simulate=False):
        self.simulating = simulate
        self.rm = pyvisa.ResourceManager()

    def __enter__(self):
        return super().__enter__()

    def __exit__(self, exc_type, exc_val, exc_tb):
        return super().__exit__(exc_type, exc_val, exc_tb)

    def open_connection(self) -> bool:
        """
        open_connection _summary_

        Returns:
            bool: _description_
        """
        try:
            if self.simulating:
                self.instr = Scope_Simulator()
                self.model = "MSO5104B"
                self.manufacturer = "Tektronix"
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
        except Exception:
            self.connected = False

        return self.connected

    def initialize(self) -> None:
        """
        initialize _summary_

        Returns:
            _type_: _description_
        """
        return super().initialize()

    def close(self) -> None:
        """
        close _summary_

        Returns:
            _type_: _description_
        """

        with contextlib.suppress(AttributeError):
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
                not self.simulating
                and self.manufacturer.upper() == "TEKTRONIX"
                or self.simulating
            )
        )

    def write(self, command: str) -> None:
        """
        write _summary_

        Args:
            command (str): _description_

        Returns:
            _type_: _description_
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
        read _summary_

        Returns:
            str: _description_
        """

        attempts = 0

        ret: str = ""

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

        assert command.find("?") > 0

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

    def read_query(self, command: str) -> float:
        """
        read_query _summary_

        Args:
            command (str): _description_

        Returns:
            float: _description_
        """

        assert command.find("?") > 0

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
        self.write("*OPC")

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

            self.instr.timeout = self.timeout  # type: ignore

        except (pyvisa.VisaIOError, AttributeError):
            self.model = ""
            self.manufacturer = ""
            self.serial = ""
            response = ",,,"

        return response.split(",")

    def get_number_channels(self) -> int:
        """
        get_number_channels
        Work out how mnay channels from the model number

        Args:
            seld (_type_): _description_

        Returns:
            int: _description_
        """

        self.num_channels = 4

        if self.model.startswith("MSO"):

            if len(self.model) == 5:
                with contextlib.suppress(ValueError):
                    self.num_channels = int(self.model[-1])

            elif self.model.endswith("B"):
                with contextlib.suppress(ValueError):
                    self.num_channels = int(self.model[-2])

        # TODO Support mode models

        return self.num_channels

    def set_channel_bw_limit(self, chan: int, bw_limit: bool | int) -> None:
        """
        set_channel_bw_limit
        Set bandwidth limit on or off

        Args:
            chan (int): _description_
            bw_limit (bool): _description_
        """

        if self.model in {
            "MSO44",
            "MSO46",
            "MSO56",
            "MSO58",
            "MSO58B",
            "MSO64",
            "MSO66",
            "MSO68",
            "MSO68B",
        }:
            if type(bw_limit) is str and bw_limit.find("M") > 0:
                bw_limit = int(bw_limit.strip()[:-1]) * 1000000
            state = str(bw_limit) if bw_limit else "FULL"
        elif type(bw_limit) is bool:
            state = "TWE" if bw_limit else "FULL"
        elif bw_limit == 150:
            state = "ONEFIFTY"
        elif bw_limit == 20:
            state = "TWENTY"
        else:
            state = "FULL"

        # Some use BAN and some BAND, so use the full command
        self.write(f"CH{chan}:BANDWIDTH {state}")
        self.write("*OPC")

    def set_channel_impedance(self, chan: int, impedance: str) -> None:
        """
        set_channel_impedance

        Although not all models can set impedance, command is available for compatibility

        Args:
            chan (int): _description_
            imedance (str): _description_
        """

        if self.model.startswith("MSO"):
            if type(impedance) is str and impedance.find("k") > 0:
                imp = int(impedance.strip()[:-1]) * 1000
            elif type(impedance) is str and impedance.find("M") > 0:
                imp = int(impedance.strip()[:-1]) * 1000000
            else:
                imp = impedance

            self.write(f"CH{chan}:TER {imp}")
        else:
            imp = "FIFTY" if impedance == "50" else "MEG"
            self.write(f"CH{chan}:IMP {imp}")

    def set_channel_invert(self, chan: int, inverted: bool) -> None:
        """
        set_channel_invert _summary_

        Args:
            inverted (bool): _description_
        """

        state = "ON" if inverted else "OFF"

        self.write(f"CH{chan}:INV {state}")

    def set_channel(self, chan: int, enabled: bool, only: bool = False) -> None:
        """
        set_channel
        Turn display of channel on or off

        Args:
            chan (int): _description_
            enabled (bool): _description_
        """

        if only:
            for channel in range(1, self.num_channels + 1):
                state = "ON" if channel == chan else "OFF"
                self.write(f"SEL:CH{channel} {state}")

        else:
            state = "ON" if enabled else "OFF"
            self.write(f"SEL:CH{chan} {state}")

        self.write("*OPC")

    def set_channel_coupling(self, chan: int, coupling: str) -> None:
        """
        set_channel_coupling _summary_

        Args:
            chan (int): _description_
            coupling (str): _description_
        """

        self.write(f"CH{chan}:COUP {coupling}")

    def set_voltage_scale(self, chan: int, scale: float, probe_atten: int = 1) -> None:
        """
        set_voltage_scale _summary_

        Args:
            chan (int): _description_
            scale (float): _description_
        """

        if self.model.startswith(("TDS")):
            self.write(f"CH{chan}:PROBE 1")
        else:
            self.write(f"CH{chan}:PROBE:GAIN 1")

        self.write(f"CH{chan}:VOL {scale}")

    def set_voltage_offset(self, chan: int, offset: float) -> None:
        """
        set_voltage_offset
        For Tek, offset is the center of the vertical acquisition window

        Args:
            chan (int): _description_
            offset (float): _description_
        """

        self.write(f"CH{chan}:OFFS {offset}")

    def set_voltage_position(self, chan: int, position: float) -> None:
        """
        set_voltage_position _summary_

        Args:
            chan (int): _description_
            position (float): _description_
        """

        self.write(f"CH{chan}:POS {position}")

    def set_timebase(self, timebase: float) -> None:
        """
        set_timebase _summary_

        Args:
            timebase (float): _description_
        """

        self.write(f"HOR:SCAL {timebase}")

    def set_timebase_pos(self, pos: float) -> None:
        """
        set_timebase_pos _summary_

        Args:
            pos (float): _description_
        """

        self.write(f"HOR:DEL:TIM {pos}")

    def set_acquisition(self, num_samples: int) -> None:
        """
        set_acquisition _summary_

        Args:
            num_samples (int): _description_
        """

        self.write("ACQ:MODE AVE")
        self.write(f"ACQ:NUMAV {num_samples}")

    def set_sample_rate(self, rate: str) -> None:
        """
        set_sample_rate
        Set the sampling rate

        Args:
            rate (float): _description_
        """

        multipliers = ["k", "M", "G"]
        if any(multiplier in rate for multiplier in multipliers):
            val = float(rate[:-1])
            if rate[-1] == "k":
                val *= 1000
            elif rate[-1] == "M":
                val *= 1_000_000
            elif rate[-1] == "G":
                val *= 1e9
        else:
            val = float(rate)

        self.write(f"HORZ:MODE:SAMPLERATE {val}")

    def set_acquisition_mode(self, mode: Tek_Acq_Mode) -> None:
        """
        set_acquisition_mode
        Set the mode for the sampling

        Args:
            mode (Tek_Acq_Mode): _description_
        """

        if mode == Tek_Acq_Mode.SAMPLE:
            md = "SAMPLE"
        elif mode == Tek_Acq_Mode.PEAK:
            md = "PEAK"
        elif mode == Tek_Acq_Mode.HIRES:
            md = "HIRES"
        elif mode == Tek_Acq_Mode.ENVELOPE:
            md = "ENV"
        else:
            md = "AVERAGE"

        self.write(f"ACQ:MODE {md}")

    def limit_measurement_population(self, channel: int, pop: int) -> None:
        """
        limit_measurement_population
        Apply a population limit for statistics

        Args:
            pop (int): _description_
        """

        self.write(f"MEASU:MEAS{channel}:POPULATION:LIMIT:STATE ON")
        self.write(f"MEASU:MEAS{channel}:POPULATION:LIMIT:VAL {pop}")

    def set_trigger_type(self, mode: str, auto_trig: bool = True) -> None:
        """
        set_trigger_mode _summary_

        Args:
            mode (str): _description_
        """

        # TODO implement edge triggering
        self.write("TRIG:SWE AUTO")

    def set_trigger_level(self, level: float, chan: int) -> None:
        """
        set_trigger_level
        Set the trigger level and source

        Args:
            level (float): _description_
            chan (int): _description_
        """

        self.write(f"TRIG:A:EDGE:SOUR CH{chan}")
        self.write(f"TRIG:A:LEV {level}")

    def measure_voltage(self, chan: int, delay: float = 1.5) -> float:
        """
        measure_voltage _summary_

        Args:
            chan (int): _description_
            delay (float, optional): Number seconds to allow measurement. Defaults to 1.

        Returns:
            float: _description_
        """

        # Only using measurement 1

        self.write("MEASU:MEAS1:TYPE MEAN")

        self.write(f"MEASU:MEAS1:SOURCE CH{chan}")
        self.write("MEASU:MEAS1:STATE ON")

        if not self.simulating:
            time.sleep(delay)

        val = self.read_query("MEASU:MEAS1:VAL?")

        if val > 9e30:
            time.sleep(1)
            val = self.read_query("MEASU:MEAS1:VAL?")

        return val

    def measure_rms_noise(self, chan: int, delay: float = 2) -> float:
        """
        measure_rms_noise
        Measure the RMS noise for sample acquisition

        Args:
            chan (int): _description_
            delay (float, optional): _description_. Defaults to 2.
        """

        self.measure_clear()
        self.write("MEASU:MEAS1:TYPE RMS")

        self.write(f"MEASU:MEAS1:SOURCE CH{chan}")
        self.write("MEASU:MEAS1:STATE ON")

        # self.limit_measurement_population(channel=chan, pop=100)  # type: ignore

        if not self.simulating:
            time.sleep(delay)

        return self.read_query("MEASU:MEAS1:VAL?")

    def get_waveform(self, chan: int, delay: float) -> tuple:
        """
        measure_voltage _summary_

        Args:
            chan (int): _description_
            delay (float): _description_
        """

        try:
            self.write(f"DATA:SOURCE CH{chan}")
            self.write("DATA:WIDTH 1")
            self.write("DATA:ENC RPB")
            self.write("DATA:START 1")
            self.write("DATA:STOP 1000")

            ymult = float(self.query("WFMINPRE:YMULT?"))
            yzero = float(self.query("WFMINPRE:YZERO?"))
            yoff = float(self.query("WFMINPRE:YOFF?"))
            xincr = float(self.query("WFMINPRE:XINCR?"))
            xdelay = float(self.query("HORizontal:POSition?"))
            self.write("CURVE?")
            data = self.instr.read_raw()  # type: ignore
            headerlen = 2 + int(data[1])
            header = data[:headerlen]
            ADC_wave = data[headerlen:-1]
            ADC_wave = np.array(unpack("%sB" % len(ADC_wave), ADC_wave))
            Volts = (ADC_wave - yoff) * ymult + yzero
            Time = np.arange(0, (xincr * len(Volts)), xincr) - (
                (xincr * len(Volts)) / 2 - xdelay
            )
            return (Time, Volts)
        except IndexError:
            return (0, 0)

    def measure_clear(self) -> None:
        """
        measure_clear _summary_
        """

        self.write("MEASU:MEAS1:STATE OFF")

    def measure_risetime(self, chan: int, num_readings: int = 1) -> float:
        """
        measure_risetime
        Use the measure function to mneasure the rise time average of n measurements

        Args:
            chan (int): _description_
            num_readings (int, optional): _description_. Defaults to 1.

        Returns:
            float: _description_
        """

        # Tek cannot reste the statistics, so a hack is to change timebase

        self.measure_clear()

        if self.model in {
            "MSO44",
            "MSO46",
            "MSO56",
            "MSO58",
            "MSO58B",
            "MSO64",
            "MSO66",
            "MSO68",
            "MSO68B",
        }:
            self.write("MEASU:MEAS1:TYPE RISETIME")
        else:
            self.write("MEASU:MEAS1:TYPE RISE")
        self.write(f"MEASU:MEAS1:SOURCE CH{chan}")
        self.write("MEASU:MEAS1:STATE ON")

        if not self.simulating:
            time.sleep(2)  # allow time to measure

        temp = self.read_query("MEASU:MEAS1:VAL?")
        if temp > 9e30:
            time.sleep(2)  # some models takes much longer to get an initial reading

        total = 0
        for _ in range(num_readings):
            total += self.read_query("MEASU:MEAS1:VAL?")
            time.sleep(0.2)

        return total / num_readings

    def read_cursor(self, cursor: str) -> float:
        """
        read_cursor
        Enable cursor and read

        Args:
            cursor (str): X1, X2, Y1, Y2

        Returns:
            float: _description_
        """

        self.write("MARK:MODE WAV")
        self.write(f"MARK:{cursor}:DISP ON")
        return self.read_query(f"MARK:{cursor}P?")

    def read_cursor_avg(self) -> float:
        """
        read_cursor_avg
        Read both of the Y cursors, and return the average

        Returns:
            float: _description_
        """

        # TODO what functions do Tek support?
        self.write("MARK:MODE WAV")

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

    def set_cursor_xy_source(self, chan: int, cursor: int) -> None:
        """
        set_cursor_xy_source
        Set the X1Y1 marker source

        Args:
            chan (int): _description_
        """

        self.write("CURS:FUNC WAV")
        self.write("CURS:MOD:TRACK")
        self.write(f"CURS:X{cursor}Y{cursor} CHAN{chan}")

    def set_cursor_position(self, cursor: str, pos: float) -> None:
        """
        set_cursor_position _summary_

        Args:
            cursor (str): _description_
            pos (float): _description_
        """

        self.write(f"MARK:{cursor}P {pos}")

    def adjust_cursor(self, target: float) -> None:
        """
        adjust_cursor
        Adjust the cursor time until target voltage met

        Args:
            target (float): _description_
        """

        current_y = self.read_query("MARK:Y1P?")
        current_x = self.read_query("MARK:X1P?")

        time_inc = self.read_query("TIM:SCAL?") / 20

        direction = +1 if current_y < target else -1

        if current_y < target:
            for _ in range(100):
                self.set_cursor_position(
                    cursor="X1", pos=current_x + time_inc * direction
                )
                current_x = self.read_query("MARK:X1P?")
                current_y = self.read_query("MARK:Y1P?")
                if ((current_y > target) and direction == 1) or (
                    (current_y < target) and direction == -1
                ):
                    break
        else:
            ...

    def cursors_on(self) -> None:
        """
        cursors_on
        Turn the markers on
        """

        self.write("CURS:FUNC WAV")
        self.write("*OPC")

    def check_triggered(self, sweep_time: float = 0.1) -> bool:
        """
        check_triggered
        Check the state of the trigger

        Args:
            sweep_time (float, optional): _description_. Defaults to 0.1.

        Returns:
            bool: _description_
        """

        response = self.query("TRIG:STATE?").strip()

        return response in ["AUTO", "TRIG"]

    def set_horizontal_mode(self, mode, record_length) -> None:
        """
        set_horizontal_mode


        Args:
            mode (_type_): _description_
            record_length (_type_): _description_
        """

        self.write(f"HOR:MODE {mode.upper()}")

        if mode.upper().startswith("MAN"):
            self.write(f"HOR:MODE:RECORDLENGTH {record_length}")


if __name__ == "__main__":
    dpo2014 = Tektronix_Oscilloscope(simulate=False)
    dpo2014.visa_address = "GPIB0::3::INSTR"

    dpo2014.open_connection()

    print(f"Model {dpo2014.model}")

    dpo2014.reset()

    dpo2014.set_channel(chan=1, enabled=True)
    dpo2014.set_channel(chan=2, enabled=True)
    dpo2014.set_voltage_scale(chan=1, scale=1)
    dpo2014.set_voltage_scale(chan=2, scale=0.2)
    dpo2014.set_voltage_offset(chan=1, offset=-1.5)  # Opposite direction to Keysight
    dpo2014.set_voltage_offset(chan=2, offset=+0.5)
    dpo2014.set_timebase(0.001)

    dpo2014.set_acquisition(32)

    print(dpo2014.check_triggered())

    dpo2014.set_trigger_type("EDGE")

    print(dpo2014.check_triggered())

    time.sleep(1)

    print(f"Measurement {dpo2014.measure_voltage(chan=1)}")

    dpo2014.set_trigger_level(level=0.1, chan=1)

    dpo2014.set_timebase(20e-9)

    dpo2014.set_cursor_xy_source(chan=1, cursor=1)
    dpo2014.set_cursor_position(cursor="X1", pos=0)

    ref_x = dpo2014.read_cursor("X1")
    time.sleep(0.1)
    ref = dpo2014.read_cursor("Y1")
    print(ref)

    dpo2014.set_timebase_pos(0.001)
    time.sleep(0.1)

    dpo2014.set_cursor_position(cursor="X1", pos=0.001)
    time.sleep(0.1)

    dpo2014.adjust_cursor(target=ref)

    offset_x = dpo2014.read_cursor("X1")

    print(f"TB Error {ref_x-offset_x+0.001}")

    input("Set voltage source to 0V")
    y1 = dpo2014.read_cursor_avg()

    print(y1)

    input("Set voltage source to 1V")
    y2 = dpo2014.read_cursor_avg()
    print(y2)
    print(f"Y Delta {y2-y1}")
