"""Section-related public enumerations, numbered as python-docx numbers them."""

from enum import IntEnum


class WD_ORIENTATION(IntEnum):
    """Page orientation, which ``Document.update_section`` takes as
    ``orientation="portrait"`` or ``"landscape"``."""

    PORTRAIT = 0
    LANDSCAPE = 1


class WD_SECTION_START(IntEnum):
    """How a section starts, which ``Document.add_section`` takes."""

    CONTINUOUS = 0
    NEW_COLUMN = 1
    NEW_PAGE = 2
    EVEN_PAGE = 3
    ODD_PAGE = 4


WD_ORIENT = WD_ORIENTATION
WD_SECTION = WD_SECTION_START

__all__ = ["WD_ORIENT", "WD_ORIENTATION", "WD_SECTION", "WD_SECTION_START"]
