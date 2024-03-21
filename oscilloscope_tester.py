"""
  Perform the main oscilloscope tests

  Tests vary by manufacturer
"""

from drivers.excel_interface import ExcelInterface
from drivers.fluke_5700a import Fluke5700A
from drivers.keysight_scope import DSOX_FAMILY, Keysight_Oscilloscope
from drivers.Ks3458A import Ks3458A, Ks3458A_Function
from drivers.Ks33250A import Ks33250A
from drivers.meatest_m142 import M142
from drivers.rf_signal_generator import RF_Signal_Generator
from drivers.rohde_shwarz_scope import RohdeSchwarz_Oscilloscope
from drivers.scpi_id import SCPI_ID
from drivers.tek_scope import Tek_Acq_Mode, Tektronix_Oscilloscope


class TestOscilloscope:

    def __init__(
        self,
        calibrator: Fluke5700A | M142,
        ks33250: Ks33250A,
        ks3458: Ks3458A,
        uut: Keysight_Oscilloscope | RohdeSchwarz_Oscilloscope | Tektronix_Oscilloscope,
        simulating: bool,
    ) -> None:
        self.calibrator = calibrator
        self.ks33250 = ks33250
        self.ks3458 = ks3458
        self.uut = uut
        self.simulating = simulating

    def local_all(self) -> None:
        """
        local_all
        Set all instruments back to local
        """

        self.calibrator.go_to_local()
        self.ks33250.go_to_local()
        self.ks3458.go_to_local()
