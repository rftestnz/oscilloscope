"""
    Generic utilities
"""

import os
import sys


def get_path(filename: str) -> str:
    """
    get_path
    The location of the diagram is different if run through interpreter or compiled.

    Args:
        filename ([str]): filename with relative path

    Returns:
        [str]: Full path to the file
    """
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)  # type: ignore
    else:
        return filename
