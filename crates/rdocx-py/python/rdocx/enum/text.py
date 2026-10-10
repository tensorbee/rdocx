"""Text-related public enumerations."""

from enum import IntEnum


class WD_ALIGN_PARAGRAPH(IntEnum):
    """Paragraph horizontal alignment."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2
    JUSTIFY = 3


class WD_UNDERLINE(IntEnum):
    """Underline styles supported by the rdocx run facade."""

    NONE = 0
    SINGLE = 1
    WORDS = 2
    DOUBLE = 3
    DOTTED = 4
    THICK = 6
    DASH = 7
    DOT_DASH = 9
    DOT_DOT_DASH = 10
    WAVY = 11


class WD_BREAK(IntEnum):
    """Break types for ``Run.add_break``, with python-docx's values."""

    LINE = 6
    PAGE = 7
    COLUMN = 8
    LINE_CLEAR_LEFT = 9
    LINE_CLEAR_RIGHT = 10
    LINE_CLEAR_ALL = 11
    TEXT_WRAPPING = 11


WD_BREAK_TYPE = WD_BREAK


__all__ = ["WD_ALIGN_PARAGRAPH", "WD_BREAK", "WD_BREAK_TYPE", "WD_UNDERLINE"]
