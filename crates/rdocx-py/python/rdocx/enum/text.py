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


class WD_TAB_ALIGNMENT(IntEnum):
    """Tab stop alignments supported by the rdocx paragraph facade."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2
    DECIMAL = 3


class WD_TAB_LEADER(IntEnum):
    """Tab leader characters supported by the rdocx paragraph facade."""

    SPACES = 0
    DOTS = 1
    DASHES = 2
    LINES = 3


__all__ = ["WD_ALIGN_PARAGRAPH", "WD_TAB_ALIGNMENT", "WD_TAB_LEADER", "WD_UNDERLINE"]
