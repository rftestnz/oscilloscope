"""
Generic test class for driver.
Running ensures all methods have been implemented
"""

from base_scope_driver import ScopeDriver
from typing import List


class TestScope(ScopeDriver):
    def __init__(self) -> None:
        super().__init__()

    def __enter__(self):
        return super().__enter__()

    def __exit__(self, exc_type, exc_val, exc_tb):
        return super().__exit__(exc_type, exc_val, exc_tb)

    def close(self) -> None:
        return super().close()

    def is_connected(self) -> bool:
        return super().is_connected()

    def open_connection(self) -> bool:
        return super().open_connection()

    def write(self, command: str) -> None:

        return super().write(command)

    def read(self) -> str:
        return super().read()

    def query(self, command: str) -> str:
        return super().query(command)

    def read_query(self, command: str) -> float:
        return super().read_query(command)

    def get_id(self) -> List:
        return super().get_id()

    def set_channel_bw_limit(self, chan: int, bw_limit: bool | int) -> None:
        return super().set_channel_bw_limit(chan, bw_limit)

    def set_channel_impedance(self, chan: int, impedance: str) -> None:
        return super().set_channel_impedance(chan, impedance)

    def set_channel_coupling(self, chan: int, coupling: str) -> None:
        return super().set_channel_coupling(chan, coupling)

    def set_channel(self, chan: int, enabled: bool, only: bool = False) -> None:
        return super().set_channel(chan, enabled, only)

    def set_channel_invert(self, chan: int, inverted: bool) -> None:
        return super().set_channel_invert(chan, inverted)

    def set_voltage_scale(self, chan: int, scale: float, probe: int = 1) -> None:
        return super().set_voltage_scale(chan, scale, probe)

    def set_voltage_offset(self, chan: int, offset: float) -> None:
        return super().set_voltage_offset(chan, offset)

    def set_voltage_position(self, chan: int, position: float) -> None:
        return super().set_voltage_position(chan, position)

    def set_timebase(self, timebase: float) -> None:
        return super().set_timebase(timebase)

    def set_timebase_pos(self, pos: float) -> None:
        return super().set_timebase_pos(pos)

    def set_acquisition(self, num_samples: int) -> None:
        return super().set_acquisition(num_samples)

    def set_trigger_type(self, mode: str) -> None:
        return super().set_trigger_type(mode)

    def set_trigger_level(self, chan: int, level: float) -> None:
        return super().set_trigger_level(chan, level)

    def measure_voltage(self, chan: int, delay: float = 2) -> float:
        return super().measure_voltage(chan, delay)

    def measure_risetime(self, chan: int, num_readings: int = 1) -> float:
        return super().measure_risetime(chan, num_readings)

    def measure_clear(self) -> None:
        return super().measure_clear()

    def check_triggered(self, sweep_time: float = 0.1) -> bool:
        return super().check_triggered(sweep_time)


if __name__ == "__main__":

    # Checks all abstract methods have been overwritten

    try:
        driver_test = TestScope()
    except TypeError as ex:
        message = ex.args.__str__()
        search_string = "abstract method"
        if search_index := message.find(search_string):
            search_index += len(search_string)
            if message[search_index] == "s":
                search_index += 1
            end_index = message.find('"', search_index)
            missing = message[search_index:end_index].split(",")
            print(f"Missing functions in class {missing}")
        else:
            print(f"Methods missing: {ex.args}")
