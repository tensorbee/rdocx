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


class MSO_LINE_DASH_STYLE(IntEnum):
    """The preset dash of a line, as `LineFormat.dash_style` reads and writes it.

    The first nine members carry the python-pptx 1.0.2 values and write the
    same `a:prstDash` value python-pptx writes. python-pptx cannot read
    `dot`, `sysDashDot` or `sysDashDotDot`, which PowerPoint for Mac's
    scripting interface writes, so rpptx adds `DOT`, `SYSTEM_DASH_DOT` and
    `SYSTEM_DASH_DOT_DOT` to read every preset back. `SYSTEM_DASH_DOT` takes
    the value that scripting interface gives `sysDashDot`, the other two take
    values outside `MsoLineDashStyle`.
    `DASH_STYLE_MIXED` is never read and cannot be written.
    """

    DASH = 4
    DASH_DOT = 5
    DASH_DOT_DOT = 6
    LONG_DASH = 7
    LONG_DASH_DOT = 8
    ROUND_DOT = 3
    SOLID = 1
    SQUARE_DOT = 2
    DASH_STYLE_MIXED = -2
    SYSTEM_DASH_DOT = 12
    DOT = 13
    SYSTEM_DASH_DOT_DOT = 14


MSO_LINE = MSO_LINE_DASH_STYLE


class MSO_ARROWHEAD_STYLE(IntEnum):
    """The shape of a line end, the `type` of `a:headEnd` or `a:tailEnd`.

    Values follow `MsoArrowheadStyle`. `OPEN` is the `arrow` end.
    """

    NONE = 1
    TRIANGLE = 2
    OPEN = 3
    STEALTH = 4
    DIAMOND = 5
    OVAL = 6


class MSO_ARROWHEAD_WIDTH(IntEnum):
    """The width of a line end, its `w` of `sm`, `med` or `lg`."""

    NARROW = 1
    MEDIUM = 2
    WIDE = 3


class MSO_ARROWHEAD_LENGTH(IntEnum):
    """The length of a line end, its `len` of `sm`, `med` or `lg`."""

    SHORT = 1
    MEDIUM = 2
    LONG = 3


class MSO_PATTERN_TYPE(IntEnum):
    """The preset of a pattern fill, as `FillFormat.pattern` reads and writes it."""

    PERCENT_5 = 1
    PERCENT_10 = 2
    PERCENT_20 = 3
    PERCENT_25 = 4
    PERCENT_30 = 5
    PERCENT_40 = 6
    PERCENT_50 = 7
    PERCENT_60 = 8
    PERCENT_70 = 9
    PERCENT_75 = 10
    PERCENT_80 = 11
    PERCENT_90 = 12
    DARK_HORIZONTAL = 13
    DARK_VERTICAL = 14
    DARK_DOWNWARD_DIAGONAL = 15
    DARK_UPWARD_DIAGONAL = 16
    SMALL_CHECKER_BOARD = 17
    TRELLIS = 18
    LIGHT_HORIZONTAL = 19
    LIGHT_VERTICAL = 20
    LIGHT_DOWNWARD_DIAGONAL = 21
    LIGHT_UPWARD_DIAGONAL = 22
    SMALL_GRID = 23
    DOTTED_DIAMOND = 24
    WIDE_DOWNWARD_DIAGONAL = 25
    WIDE_UPWARD_DIAGONAL = 26
    DASHED_UPWARD_DIAGONAL = 27
    DASHED_DOWNWARD_DIAGONAL = 28
    NARROW_VERTICAL = 29
    NARROW_HORIZONTAL = 30
    DASHED_VERTICAL = 31
    DASHED_HORIZONTAL = 32
    LARGE_CONFETTI = 33
    LARGE_GRID = 34
    HORIZONTAL_BRICK = 35
    LARGE_CHECKER_BOARD = 36
    SMALL_CONFETTI = 37
    ZIG_ZAG = 38
    SOLID_DIAMOND = 39
    DIAGONAL_BRICK = 40
    OUTLINED_DIAMOND = 41
    PLAID = 42
    SPHERE = 43
    WEAVE = 44
    DOTTED_GRID = 45
    DIVOT = 46
    SHINGLE = 47
    WAVE = 48
    HORIZONTAL = 49
    VERTICAL = 50
    CROSS = 51
    DOWNWARD_DIAGONAL = 52
    UPWARD_DIAGONAL = 53
    DIAGONAL_CROSS = 54
    MIXED = -2


MSO_PATTERN = MSO_PATTERN_TYPE


__all__ = [
    "MSO_ARROWHEAD_LENGTH",
    "MSO_ARROWHEAD_STYLE",
    "MSO_ARROWHEAD_WIDTH",
    "MSO_FILL",
    "MSO_FILL_TYPE",
    "MSO_LINE",
    "MSO_LINE_DASH_STYLE",
    "MSO_PATTERN",
    "MSO_PATTERN_TYPE",
]
