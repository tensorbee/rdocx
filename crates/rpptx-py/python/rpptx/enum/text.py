"""Text enumerations, with python-pptx 1.0.2 member values."""

from enum import IntEnum


class MSO_AUTO_SIZE(IntEnum):
    """How a text frame fits its text."""

    NONE = 0
    SHAPE_TO_FIT_TEXT = 1
    TEXT_TO_FIT_SHAPE = 2


class MSO_VERTICAL_ANCHOR(IntEnum):
    """Vertical alignment of text in a text frame.

    `JUSTIFY` and `DISTRIBUTE` are the `just` and `dist` anchors of the
    schema, which python-pptx does not model.
    """

    TOP = 1
    MIDDLE = 3
    BOTTOM = 4
    JUSTIFY = 6
    DISTRIBUTE = 7


MSO_ANCHOR = MSO_VERTICAL_ANCHOR


class PP_PARAGRAPH_ALIGNMENT(IntEnum):
    """Horizontal alignment of a paragraph."""

    LEFT = 1
    CENTER = 2
    RIGHT = 3
    JUSTIFY = 4
    DISTRIBUTE = 5
    THAI_DISTRIBUTE = 6
    JUSTIFY_LOW = 7


PP_ALIGN = PP_PARAGRAPH_ALIGNMENT


class MSO_TEXT_UNDERLINE_TYPE(IntEnum):
    """Underline style of a run."""

    NONE = 0
    WORDS = 1
    SINGLE_LINE = 2
    DOUBLE_LINE = 3
    HEAVY_LINE = 4
    DOTTED_LINE = 5
    DOTTED_HEAVY_LINE = 6
    DASH_LINE = 7
    DASH_HEAVY_LINE = 8
    DASH_LONG_LINE = 9
    DASH_LONG_HEAVY_LINE = 10
    DOT_DASH_LINE = 11
    DOT_DASH_HEAVY_LINE = 12
    DOT_DOT_DASH_LINE = 13
    DOT_DOT_DASH_HEAVY_LINE = 14
    WAVY_LINE = 15
    WAVY_HEAVY_LINE = 16
    WAVY_DOUBLE_LINE = 17


MSO_UNDERLINE = MSO_TEXT_UNDERLINE_TYPE

__all__ = [
    "MSO_ANCHOR",
    "MSO_AUTO_SIZE",
    "MSO_TEXT_UNDERLINE_TYPE",
    "MSO_UNDERLINE",
    "MSO_VERTICAL_ANCHOR",
    "PP_ALIGN",
    "PP_PARAGRAPH_ALIGNMENT",
]
