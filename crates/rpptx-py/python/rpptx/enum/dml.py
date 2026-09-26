"""DrawingML enumerations, mirroring python-pptx `pptx.enum.dml`."""

from enum import IntEnum


class MSO_FILL_TYPE(IntEnum):
    """The kind of a fill, as reported by `FillFormat.type`."""

    BACKGROUND = 5
    GRADIENT = 3
    GROUP = 101
    PATTERNED = 2
    PICTURE = 6
    SOLID = 1
    TEXTURED = 4


MSO_FILL = MSO_FILL_TYPE


__all__ = ["MSO_FILL", "MSO_FILL_TYPE"]
