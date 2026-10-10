"""Table-related public enumerations."""

from enum import IntEnum


class WD_TABLE_ALIGNMENT(IntEnum):
    """Table horizontal alignment."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2


class WD_CELL_VERTICAL_ALIGNMENT(IntEnum):
    """Vertical alignment within a table cell."""

    TOP = 0
    CENTER = 1
    BOTTOM = 3


class WD_ROW_HEIGHT_RULE(IntEnum):
    """How a table row height is applied."""

    AT_LEAST = 1
    EXACTLY = 2


WD_ALIGN_VERTICAL = WD_CELL_VERTICAL_ALIGNMENT

__all__ = [
    "WD_ALIGN_VERTICAL",
    "WD_TABLE_ALIGNMENT",
    "WD_CELL_VERTICAL_ALIGNMENT",
    "WD_ROW_HEIGHT_RULE",
]
