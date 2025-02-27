"""
Base class for scope drivers
All drivers inherot from this abstract class

DK Feb 23
"""

import abc  # Abstract Base Class
import pyvisa
import time
from typing import List
from random import random


class Scope_Simulator:
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
        print(f"SCOPE <- {command}")

    def read(self) -> float | str:
        """
        read _summary_

        Returns:
            float| str: _description_
        """
        return 0.5 + random()

    def query(self, command: str) -> str:
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """

        print(f"SCOPE <- {command}")

        if command == "*IDN?":
            return "Tektronix,DPO2024,MY_Simulated,B.00.00"

        return str(0.5 + random()) if command.startswith("READ") else ""


class ScopeDriver(metaclass=abc.ABCMeta):
    """
    ScopeDriver
    Defines all functions a derived class must implement

    Args:
        metaclass (_type_, optional): _description_. Defaults to abc.ABCMeta.
    """

    @abc.abstractmethod
    def __init__(self, simulate=False):
        self.simulating = simulate
        self.rm = pyvisa.ResourceManager()

    @abc.abstractmethod
    def __enter__(self):
        self.open_connection()
        return self

    @abc.abstractmethod
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    @abc.abstractmethod
    def open_connection(self) -> bool:
        """
        open_connection _summary_

        Returns:
            bool: _description_
        """

    def initialize(self) -> None:
        """
        initialize _summary_
        """
        self.open_connection()

    @abc.abstractmethod
    def close(self) -> None:
        """
        close _summary_
        """

    @abc.abstractmethod
    def is_connected(self) -> bool:
        """
        is_connected _summary_

        Returns:
            bool: _description_
        """

        pass

    @abc.abstractmethod
    def write(self, command: str) -> None:
        """
        write
        Wite command to instrument

        Args:
            command (str): _description_
        """

    @abc.abstractmethod
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

    @abc.abstractmethod
    def query(self, command: str) -> str:
        """
        query _summary_

        Args:
            command (str): _description_

        Returns:
            str: _description_
        """

    @abc.abstractmethod
    def read_query(self, command: str) -> float:
        """
        read_query
        send query command, return a float or -999 if invalid

        Args:
            command (str): _description_

        Returns:
            float: _description_
        """

    @abc.abstractmethod
    def get_id(self) -> List:
        """
        get_id _summary_

        Returns:
            List: _description_
        """

        pass

    @abc.abstractmethod
    def set_channel_bw_limit(self, chan: int, bw_limit: bool | int) -> None:
        """
        set_channel_bw_limit _summary_

        Args:
            chan (int): _description_
            bw_limit (bool | int): _description_
        """

        pass

    @abc.abstractmethod
    def set_channel_impedance(self, chan: int, impedance: str) -> None:
        """
        set_channel_impedance _summary_

        Args:
            chan (int): _description_
            impedance (str): _description_
        """

        pass

    @abc.abstractmethod
    def set_channel_invert(self, chan: int, inverted: bool) -> None:
        """
        set_channel_invert _summary_

        Args:
            chan (int): _description_
            inverted (bool): _description_
        """

        pass

    @abc.abstractmethod
    def set_channel(self, chan: int, enabled: bool, only: bool = False) -> None:
        """
        set_channel _summary_

        Args:
            chan (int): _description_
            enabled (bool): _description_
            only (bool, optional): _description_. Defaults to False.
        """

        pass

    @abc.abstractmethod
    def set_channel_coupling(self, chan: int, coupling: str) -> None:
        """
        set_channel_coupling _summary_

        Args:
            chan (int): _description_
            coupling (str): _description_
        """

        pass

    @abc.abstractmethod
    def set_voltage_scale(self, chan: int, scale: float, probe: int = 1) -> None:
        """
        set_voltage_scale _summary_

        Args:
            chan (int): _description_
            scale (float): _description_
            probe (int, optional): _description_. Defaults to 1.
        """

        pass

    @abc.abstractmethod
    def set_voltage_offset(self, chan: int, offset: float) -> None:
        """
        set_voltage_offset _summary_

        Args:
            chan (int): _description_
            offset (float): _description_
        """

        pass

    @abc.abstractmethod
    def set_voltage_position(self, chan: int, position: float) -> None:
        """
        set_voltage_position _summary_

        Args:
            chan (int): _description_
            position (float): _description_
        """

        pass

    @abc.abstractmethod
    def set_timebase(self, timebase: float) -> None:
        """
        set_timebase _summary_

        Args:
            timebase (float): _description_
        """

        pass

    @abc.abstractmethod
    def set_timebase_pos(self, pos: float) -> None:
        """
        set_timebase_pos _summary_

        Args:
            pos (float): _description_
        """

        pass

    @abc.abstractmethod
    def set_acquisition(self, num_samples: int) -> None:
        """
        set_acquisition _summary_

        Args:
            num_samples (int): _description_
        """

        pass

    @abc.abstractmethod
    def set_trigger_type(self, mode: str) -> None:
        """
        set_trigger_mode _summary_

        Args:
            mode (str): _description_
        """

        pass

    @abc.abstractmethod
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

    @abc.abstractmethod
    def measure_voltage(self, chan: int, delay: float = 2) -> float:
        """
        measure_voltage
        Return the average voltage

        Returns:
            float: _description_
        """

        pass

    @abc.abstractmethod
    def measure_clear(self) -> None:
        """
        measure_voltage_clear _summary_
        """

        pass

    @abc.abstractmethod
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

        pass

    @abc.abstractmethod
    def check_triggered(self, sweep_time: float = 0.1) -> bool:
        """
        check_triggered
        Check if the scope is triggered
        clears it first then waits for a sweep and returns the status

        Returns:
            bool: _description_
        """

        pass
