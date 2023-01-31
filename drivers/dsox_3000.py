"""
# Driver for Keysight DSOX oscilloscopes


# DK Jan 23
"""

from enum import Enum
import pyvisa
import time
from random import random
from typing import List

VERSION = "A.00.00"


class DSOX_FAMILY(Enum):
    """
    DSOX_FAMILY
    Enum to cope with minor differences in commands

    Args:
        Enum (_type_): _description_
    """

    DSOX1000 = 1
    DSOX2000 = 2
    DSOX3000 = 3


class DSOX3000_Simulator:
    """
    _summary_
    """

    def close(self) -> None:  # type: ignore
        """
        close _summary_
        """
        pass

    def write(self, command: str) -> None:  # type: ignore
        """
        write _summary_

        Args:
            command (str): _description_
        """
        # sourcery skip: instance-method-first-arg-name
        print(f"DSOX3000 <- {command}")

    def read(self) -> float | str:  # type: ignore
        """
        read _summary_

        Returns:
            float| str: _description_
        """

        return 0.5 + random()

    def query(self, command: str) -> str:  # type: ignore
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """

        # sourcery skip: instance-method-first-arg-name

        print(f"DRIVER_NAME <- {command}")

        if command == "*IDN?":
            return "Keysight,DSOX3034G,MY_Simulated,B.00.00"

        return (
            str(0.5 + random())
            if command.startswith("READ") or command.startswith("MEAS")
            else ""
        )


class DSOX_3000:
    """
     _summary_

    Returns:
        _type_: _description_
    """

    connected: bool = False
    visa_address: str = "USB0::0x2A8D::0x1797::CN59296333::INSTR"
    model = ""
    manufacturer = ""
    serial = ""
    family = DSOX_FAMILY.DSOX3000
    timeout = 2000
    num_channels = 4
    keysight: bool = True

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
                self.instr = DSOX3000_Simulator()
                self.model = "DSO-X 3034T"
                self.manufacturer = "Keysight"
                self.serial = "666"
            else:
                self.instr = self.rm.open_resource(
                    self.visa_address, write_termination="\n"
                )
                self.instr.timeout = self.timeout
                # self.instr.control_ren(VI_GPIB_REN_ASSERT)  # type: ignore
                self.get_id()
                self.get_num_channels()
            self.connected = True
        except Exception:
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
                not self.simulating and self.model.find("DSO-X") >= 0 or self.simulating
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
            except pyvisa.VisaIOError:
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

    def read_query(self, command: str) -> float:
        """
        read_query
        send query command, return a float or -999 if invalid

        Args:
            command (str): _description_

        Returns:
            float: _description_
        """

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
                model_type = self.model.split(" ")
                if len(model_type) > 1:
                    family = model_type[1][0]
                    if family == "1":
                        self.family = DSOX_FAMILY.DSOX1000
                    else:
                        self.family = (
                            DSOX_FAMILY.DSOX2000
                            if family == "2"
                            else DSOX_FAMILY.DSOX3000
                        )
                self.serial = identity[2]

        except pyvisa.VisaIOError:
            self.model = ""
            self.manufacturer = ""
            self.serial = ""
            response = ",,,"

        self.instr.timeout = self.timeout  # type: ignore

        return response.split(",")

    def get_num_channels(self) -> None:
        """
        get_num_channels
        The number of channels can be obtained from the model number, but use a more universal method
        of querying channels
        """

        # use model for now until figuring out timeout

        valid = False
        model_fields = self.model.split(" ")
        if len(model_fields) > 1:
            try:
                self.num_channels = int(model_fields[1][3])
                valid = True
            except (ValueError, IndexError):
                valid = False

        if not valid:
            tmo = self.instr.timeout  # type: ignore
            self.instr.timeout = 500  # type: ignore

            # TODO it ignores the timeout, and uses 5000 ms
            for chan in range(4):
                chan_num = chan + 1

                try:
                    reply = self.query(f"CHAN{chan_num}?")
                    if not len(reply) and chan_num > 2:
                        self.num_channels = chan
                        break
                except pyvisa.VisaIOError:
                    if chan_num > 2:
                        self.num_channels = chan
                    break

            self.timeout = tmo

    def set_channel_bw_limit(self, chan: int, bw_limit: bool) -> None:
        """
        set_channel_bw_limit
        Set bandwidth limit on or off

        Args:
            chan (int): _description_
            bw_limit (bool): _description_
        """

        state = "ON" if bw_limit else "OFF"

        self.write(f"CHAN{chan}:BWL {state}")
        self.write("*OPC")

    def set_channel_impedance(self, chan: int, impedance: str) -> None:
        """
        set_channel_impedance _summary_

        Args:
            chan (int): _description_
            imedance (str): _description_
        """

        imp = "FIFTY" if impedance == "50" else "ONEMEG"

        self.write(f"CHAN{chan}:IMP {imp}")

    def set_channel(self, chan: int, enabled: bool) -> None:
        """
        set_channel
        Turn display of channel on or off

        Args:
            chan (int): _description_
            enabled (bool): _description_
        """

        state = "ON" if enabled else "OFF"

        self.write(f"CHAN{chan}:DISP {state}")
        self.write("*OPC")

    def set_voltage_scale(self, chan: int, scale: float, probe: int = 1) -> None:
        """
        set_voltage_scale _summary_

        Args:
            chan (int): _description_
            scale (float): _description_
        """

        self.write(f"CHAN{chan}:PROB {probe}")  # Set before the scale
        self.write(f"CHAN{chan}:SCAL {scale}")
        self.write("*OPC")

    def set_voltage_offset(self, chan: int, offset: float) -> None:
        """
        set_voltage_offset _summary_

        Args:
            chan (int): _description_
            offset (float): _description_
        """

        self.write(f"CHAN{chan}:OFFS {offset}")
        self.write("*OPC")

    def set_timebase(self, timebase: float) -> None:
        """
        set_timebase _summary_

        Args:
            timebase (float): _description_
        """

        self.write(f"TIM:SCAL {timebase}")
        self.write("*OPC")

    def set_timebase_pos(self, pos: float) -> None:
        """
        set_timebase_pos _summary_

        Args:
            pos (float): _description_
        """

        self.write(f"TIM:POS {pos}")
        self.write("*OPC")

    def set_acquisition(self, num_samples: int) -> None:
        """
        set_acquisition _summary_

        Args:
            num_samples (int): _description_
        """

        self.write("ACQ:TYPE AVER")
        self.write(f"ACQ:COUNT {num_samples}")
        self.write("*OPC")

    def set_trigger_mode(self, mode: str) -> None:
        """
        set_trigger_mode _summary_

        Args:
            mode (str): _description_
        """

        self.write(f"TRIG:MODE {mode}")
        self.write("TRIG:SWE AUTO")
        self.write("*OPC")

    def set_trigger_level(
        self,
        chan: int,
        level: float,
    ) -> None:
        """
        set_trigger_level
        Set the trigger level and source

        Args:
            level (float): _description_
            chan (int): 0 for ext, else channel number
        """

        source = f"CHAN{chan}" if chan else "EXT"

        self.write(f"TRIG:EDGE:SOUR {source}")
        self.write(f"TRIG:EDGE:LEV {level}")
        self.write("*OPC")

    def measure_voltage(self, chan: int) -> float:
        """
        measure_voltage
        Return the average voltage

        Returns:
            float: _description_
        """

        self.write(f"MEAS:SOURCE CHAN{chan}")

        return self.read_query("MEAS:VAV?")

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
        self.write(f"MEAS:RIS CHAN{chan}")
        self.write("*OPC")

        total = 0
        for _ in range(num_readings):
            total += self.read_query(f"MEAS:RIS? CHAN{chan}")
            time.sleep(0.1)

        return total / num_readings

    def cursors_on(self) -> None:
        """
        cursors_on
        Turn the markers on
        """

        self.write("MARK:MODE WAV")
        self.write("*OPC")

    def read_cursor(self, cursor: str) -> float:
        """

        read_cursor _summary_

        Enable cursor and read
        """

        # TODO which family supprt this command
        if self.family != DSOX_FAMILY.DSOX1000:
            self.write(f"MARK:{cursor}:DISP ON")
        self.write("*OPC")
        pos = self.read_query(f"MARK:{cursor}P?")
        if pos > 9e37:
            time.sleep(0.2)
            pos = self.read_query(f"MARK:{cursor}P?")
        return pos

    def read_cursor_avg(self) -> float:
        """
        read_cursor_avg
        Read both of the Y cursors, and return the average

        Returns:
            float: _description_
        """

        self.write("MARK:MODE WAV")
        if self.family == DSOX_FAMILY.DSOX3000:
            self.write("MARK:Y1:DISP ON")
            self.write("MARK:Y2:DISP ON")
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

        self.write("MARK:MODE WAV")
        self.write(f"MARK:X{cursor}Y{cursor} CHAN{chan}")
        self.write("*OPC")

    def set_cursor_position(self, cursor: str, pos: float) -> None:
        """
        set_cursor_position _summary_

        Args:
            cursor (str): _description_
            pos (float): _description_
        """

        self.write(f"MARK:{cursor}P {pos}")
        self.write("*OPC")

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

        diff = abs(current_y - target)

        if diff > 0.05:  # 0.1 div
            for _ in range(100):
                self.set_cursor_position(
                    cursor="X1", pos=current_x + time_inc * direction
                )
                self.write("*OPC")
                time.sleep(0.05)
                current_x = self.read_query("MARK:X1P?")
                current_y = self.read_query("MARK:Y1P?")

                diff = current_y - target

                # As the slope can be steep, detect when it has passed the crossing

                val = direction * diff
                if val > 0:
                    break
        else:
            ...

    def check_triggered(self, sweep_time: float = 0.1) -> bool:
        """
        check_triggered
        Check if the scope is triggered
        clears it first then waits for a sweep and returns the status

        Returns:
            bool: _description_
        """

        self.write("*CLS")

        time.sleep(sweep_time)

        triggered = self.query("TER?")

        return triggered == "1"


if __name__ == "__main__":

    dsox3034t = DSOX_3000(simulate=False)
    dsox3034t.visa_address = "USB0::0x2A8D::0x1797::CN59296333::INSTR"

    dsox3034t.open_connection()

    print(f"Model {dsox3034t.model}")
    print(f"Num channels {dsox3034t.num_channels}")

    dsox3034t.reset()

    dsox3034t.set_channel(chan=1, enabled=True)
    dsox3034t.set_channel(chan=2, enabled=True)
    dsox3034t.set_channel_bw_limit(chan=1, bw_limit=True)
    dsox3034t.set_voltage_scale(chan=1, scale=0.5)
    dsox3034t.set_voltage_scale(chan=2, scale=0.2)
    dsox3034t.set_voltage_offset(chan=1, offset=0)
    dsox3034t.set_voltage_offset(chan=2, offset=-0.5)
    dsox3034t.set_timebase(0.001)
    dsox3034t.set_timebase(5e-9)
    dsox3034t.set_acquisition(64)

    dsox3034t.set_trigger_mode("EDGE")

    time.sleep(1)

    print(f"Measurement {dsox3034t.measure_voltage(chan=1)}")

    print(f"Measurement {dsox3034t.measure_risetime(chan=1, num_readings=10)}")

    dsox3034t.set_trigger_level(
        chan=1,
        level=0,
    )

    dsox3034t.set_timebase(20e-9)

    dsox3034t.set_cursor_xy_source(chan=1, cursor=1)
    dsox3034t.set_cursor_position(cursor="X1", pos=0)

    ref_x = dsox3034t.read_cursor("X1")
    time.sleep(0.1)
    ref = dsox3034t.read_cursor("Y1")
    print(ref)

    dsox3034t.set_timebase_pos(0.001)
    time.sleep(0.1)

    dsox3034t.set_cursor_position(cursor="X1", pos=0.001)
    time.sleep(0.1)

    dsox3034t.adjust_cursor(target=ref)

    offset_x = dsox3034t.read_cursor("X1")

    print(f"TB Error {ref_x-offset_x+0.001}")

    input("Set voltage source to 0V")
    y1 = dsox3034t.read_cursor_avg()

    print(y1)

    input("Set voltage source to 1V")
    y2 = dsox3034t.read_cursor_avg()
    print(y2)
    print(f"Y Delta {y2-y1}")
