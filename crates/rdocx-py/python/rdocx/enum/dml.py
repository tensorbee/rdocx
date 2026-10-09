"""DrawingML colour enumerations, mirroring python-docx `docx.enum.dml`."""

from enum import IntEnum


class MSO_COLOR_TYPE(IntEnum):
    """The kind of a run colour, as reported by `ColorFormat.type`."""

    RGB = 1
    THEME = 2
    AUTO = 101


MSO_COLOR = MSO_COLOR_TYPE


class MSO_THEME_COLOR_INDEX(IntEnum):
    """A theme colour, as `ColorFormat.theme_color` reads and writes it.

    Values follow python-docx. Each member writes the `w:themeColor` token
    Word uses, such as `accent1` for `ACCENT_1`.
    """

    NOT_THEME_COLOR = 0
    ACCENT_1 = 5
    ACCENT_2 = 6
    ACCENT_3 = 7
    ACCENT_4 = 8
    ACCENT_5 = 9
    ACCENT_6 = 10
    BACKGROUND_1 = 14
    BACKGROUND_2 = 16
    DARK_1 = 1
    DARK_2 = 3
    FOLLOWED_HYPERLINK = 12
    HYPERLINK = 11
    LIGHT_1 = 2
    LIGHT_2 = 4
    TEXT_1 = 13
    TEXT_2 = 15


MSO_THEME_COLOR = MSO_THEME_COLOR_INDEX


__all__ = [
    "MSO_COLOR",
    "MSO_COLOR_TYPE",
    "MSO_THEME_COLOR",
    "MSO_THEME_COLOR_INDEX",
]
