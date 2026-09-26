"""Shape enumerations, mirroring python-pptx `pptx.enum.shapes`."""

from enum import IntEnum


class MSO_SHAPE(IntEnum):
    """Preset shapes `add_shape` can author, with python-pptx values.

    python-pptx also defines `UP_ARROW`, which is absent here because the
    ECMA preset table rpptx generates from lists `upDownArrow` twice and no
    `upArrow`.
    """

    ACTION_BUTTON_BACK_OR_PREVIOUS = 129
    ACTION_BUTTON_BEGINNING = 131
    ACTION_BUTTON_CUSTOM = 125
    ACTION_BUTTON_DOCUMENT = 134
    ACTION_BUTTON_END = 132
    ACTION_BUTTON_FORWARD_OR_NEXT = 130
    ACTION_BUTTON_HELP = 127
    ACTION_BUTTON_HOME = 126
    ACTION_BUTTON_INFORMATION = 128
    ACTION_BUTTON_MOVIE = 136
    ACTION_BUTTON_RETURN = 133
    ACTION_BUTTON_SOUND = 135
    ARC = 25
    BALLOON = 137
    BENT_ARROW = 41
    BENT_UP_ARROW = 44
    BEVEL = 15
    BLOCK_ARC = 20
    CAN = 13
    CHART_PLUS = 182
    CHART_STAR = 181
    CHART_X = 180
    CHEVRON = 52
    CHORD = 161
    CIRCULAR_ARROW = 60
    CLOUD = 179
    CLOUD_CALLOUT = 108
    CORNER = 162
    CORNER_TABS = 169
    CROSS = 11
    CUBE = 14
    CURVED_DOWN_ARROW = 48
    CURVED_DOWN_RIBBON = 100
    CURVED_LEFT_ARROW = 46
    CURVED_RIGHT_ARROW = 45
    CURVED_UP_ARROW = 47
    CURVED_UP_RIBBON = 99
    DECAGON = 144
    DIAGONAL_STRIPE = 141
    DIAMOND = 4
    DODECAGON = 146
    DONUT = 18
    DOUBLE_BRACE = 27
    DOUBLE_BRACKET = 26
    DOUBLE_WAVE = 104
    DOWN_ARROW = 36
    DOWN_ARROW_CALLOUT = 56
    DOWN_RIBBON = 98
    EXPLOSION1 = 89
    EXPLOSION2 = 90
    FLOWCHART_ALTERNATE_PROCESS = 62
    FLOWCHART_CARD = 75
    FLOWCHART_COLLATE = 79
    FLOWCHART_CONNECTOR = 73
    FLOWCHART_DATA = 64
    FLOWCHART_DECISION = 63
    FLOWCHART_DELAY = 84
    FLOWCHART_DIRECT_ACCESS_STORAGE = 87
    FLOWCHART_DISPLAY = 88
    FLOWCHART_DOCUMENT = 67
    FLOWCHART_EXTRACT = 81
    FLOWCHART_INTERNAL_STORAGE = 66
    FLOWCHART_MAGNETIC_DISK = 86
    FLOWCHART_MANUAL_INPUT = 71
    FLOWCHART_MANUAL_OPERATION = 72
    FLOWCHART_MERGE = 82
    FLOWCHART_MULTIDOCUMENT = 68
    FLOWCHART_OFFLINE_STORAGE = 139
    FLOWCHART_OFFPAGE_CONNECTOR = 74
    FLOWCHART_OR = 78
    FLOWCHART_PREDEFINED_PROCESS = 65
    FLOWCHART_PREPARATION = 70
    FLOWCHART_PROCESS = 61
    FLOWCHART_PUNCHED_TAPE = 76
    FLOWCHART_SEQUENTIAL_ACCESS_STORAGE = 85
    FLOWCHART_SORT = 80
    FLOWCHART_STORED_DATA = 83
    FLOWCHART_SUMMING_JUNCTION = 77
    FLOWCHART_TERMINATOR = 69
    FOLDED_CORNER = 16
    FRAME = 158
    FUNNEL = 174
    GEAR_6 = 172
    GEAR_9 = 173
    HALF_FRAME = 159
    HEART = 21
    HEPTAGON = 145
    HEXAGON = 10
    HORIZONTAL_SCROLL = 102
    ISOSCELES_TRIANGLE = 7
    LEFT_ARROW = 34
    LEFT_ARROW_CALLOUT = 54
    LEFT_BRACE = 31
    LEFT_BRACKET = 29
    LEFT_CIRCULAR_ARROW = 176
    LEFT_RIGHT_ARROW = 37
    LEFT_RIGHT_ARROW_CALLOUT = 57
    LEFT_RIGHT_CIRCULAR_ARROW = 177
    LEFT_RIGHT_RIBBON = 140
    LEFT_RIGHT_UP_ARROW = 40
    LEFT_UP_ARROW = 43
    LIGHTNING_BOLT = 22
    LINE_CALLOUT_1 = 109
    LINE_CALLOUT_1_ACCENT_BAR = 113
    LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR = 121
    LINE_CALLOUT_1_NO_BORDER = 117
    LINE_CALLOUT_2 = 110
    LINE_CALLOUT_2_ACCENT_BAR = 114
    LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR = 122
    LINE_CALLOUT_2_NO_BORDER = 118
    LINE_CALLOUT_3 = 111
    LINE_CALLOUT_3_ACCENT_BAR = 115
    LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR = 123
    LINE_CALLOUT_3_NO_BORDER = 119
    LINE_CALLOUT_4 = 112
    LINE_CALLOUT_4_ACCENT_BAR = 116
    LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR = 124
    LINE_CALLOUT_4_NO_BORDER = 120
    LINE_INVERSE = 183
    MATH_DIVIDE = 166
    MATH_EQUAL = 167
    MATH_MINUS = 164
    MATH_MULTIPLY = 165
    MATH_NOT_EQUAL = 168
    MATH_PLUS = 163
    MOON = 24
    NON_ISOSCELES_TRAPEZOID = 143
    NOTCHED_RIGHT_ARROW = 50
    NO_SYMBOL = 19
    OCTAGON = 6
    OVAL = 9
    OVAL_CALLOUT = 107
    PARALLELOGRAM = 2
    PENTAGON = 51
    PIE = 142
    PIE_WEDGE = 175
    PLAQUE = 28
    PLAQUE_TABS = 171
    QUAD_ARROW = 39
    QUAD_ARROW_CALLOUT = 59
    RECTANGLE = 1
    RECTANGULAR_CALLOUT = 105
    REGULAR_PENTAGON = 12
    RIGHT_ARROW = 33
    RIGHT_ARROW_CALLOUT = 53
    RIGHT_BRACE = 32
    RIGHT_BRACKET = 30
    RIGHT_TRIANGLE = 8
    ROUNDED_RECTANGLE = 5
    ROUNDED_RECTANGULAR_CALLOUT = 106
    ROUND_1_RECTANGLE = 151
    ROUND_2_DIAG_RECTANGLE = 153
    ROUND_2_SAME_RECTANGLE = 152
    SMILEY_FACE = 17
    SNIP_1_RECTANGLE = 155
    SNIP_2_DIAG_RECTANGLE = 157
    SNIP_2_SAME_RECTANGLE = 156
    SNIP_ROUND_RECTANGLE = 154
    SQUARE_TABS = 170
    STAR_10_POINT = 149
    STAR_12_POINT = 150
    STAR_16_POINT = 94
    STAR_24_POINT = 95
    STAR_32_POINT = 96
    STAR_4_POINT = 91
    STAR_5_POINT = 92
    STAR_6_POINT = 147
    STAR_7_POINT = 148
    STAR_8_POINT = 93
    STRIPED_RIGHT_ARROW = 49
    SUN = 23
    SWOOSH_ARROW = 178
    TEAR = 160
    TRAPEZOID = 3
    UP_ARROW_CALLOUT = 55
    UP_DOWN_ARROW = 38
    UP_DOWN_ARROW_CALLOUT = 58
    UP_RIBBON = 97
    U_TURN_ARROW = 42
    VERTICAL_SCROLL = 101
    WAVE = 103

    @property
    def xml_value(self) -> str:
        """The DrawingML preset name, such as `roundRect`."""
        return _PRESETS[self.value]


_PRESETS: dict[int, str] = {
    129: "actionButtonBackPrevious",
    131: "actionButtonBeginning",
    125: "actionButtonBlank",
    134: "actionButtonDocument",
    132: "actionButtonEnd",
    130: "actionButtonForwardNext",
    127: "actionButtonHelp",
    126: "actionButtonHome",
    128: "actionButtonInformation",
    136: "actionButtonMovie",
    133: "actionButtonReturn",
    135: "actionButtonSound",
    25: "arc",
    137: "wedgeRoundRectCallout",
    41: "bentArrow",
    44: "bentUpArrow",
    15: "bevel",
    20: "blockArc",
    13: "can",
    182: "chartPlus",
    181: "chartStar",
    180: "chartX",
    52: "chevron",
    161: "chord",
    60: "circularArrow",
    179: "cloud",
    108: "cloudCallout",
    162: "corner",
    169: "cornerTabs",
    11: "plus",
    14: "cube",
    48: "curvedDownArrow",
    100: "ellipseRibbon",
    46: "curvedLeftArrow",
    45: "curvedRightArrow",
    47: "curvedUpArrow",
    99: "ellipseRibbon2",
    144: "decagon",
    141: "diagStripe",
    4: "diamond",
    146: "dodecagon",
    18: "donut",
    27: "bracePair",
    26: "bracketPair",
    104: "doubleWave",
    36: "downArrow",
    56: "downArrowCallout",
    98: "ribbon",
    89: "irregularSeal1",
    90: "irregularSeal2",
    62: "flowChartAlternateProcess",
    75: "flowChartPunchedCard",
    79: "flowChartCollate",
    73: "flowChartConnector",
    64: "flowChartInputOutput",
    63: "flowChartDecision",
    84: "flowChartDelay",
    87: "flowChartMagneticDrum",
    88: "flowChartDisplay",
    67: "flowChartDocument",
    81: "flowChartExtract",
    66: "flowChartInternalStorage",
    86: "flowChartMagneticDisk",
    71: "flowChartManualInput",
    72: "flowChartManualOperation",
    82: "flowChartMerge",
    68: "flowChartMultidocument",
    139: "flowChartOfflineStorage",
    74: "flowChartOffpageConnector",
    78: "flowChartOr",
    65: "flowChartPredefinedProcess",
    70: "flowChartPreparation",
    61: "flowChartProcess",
    76: "flowChartPunchedTape",
    85: "flowChartMagneticTape",
    80: "flowChartSort",
    83: "flowChartOnlineStorage",
    77: "flowChartSummingJunction",
    69: "flowChartTerminator",
    16: "foldedCorner",
    158: "frame",
    174: "funnel",
    172: "gear6",
    173: "gear9",
    159: "halfFrame",
    21: "heart",
    145: "heptagon",
    10: "hexagon",
    102: "horizontalScroll",
    7: "triangle",
    34: "leftArrow",
    54: "leftArrowCallout",
    31: "leftBrace",
    29: "leftBracket",
    176: "leftCircularArrow",
    37: "leftRightArrow",
    57: "leftRightArrowCallout",
    177: "leftRightCircularArrow",
    140: "leftRightRibbon",
    40: "leftRightUpArrow",
    43: "leftUpArrow",
    22: "lightningBolt",
    109: "borderCallout1",
    113: "accentCallout1",
    121: "accentBorderCallout1",
    117: "callout1",
    110: "borderCallout2",
    114: "accentCallout2",
    122: "accentBorderCallout2",
    118: "callout2",
    111: "borderCallout3",
    115: "accentCallout3",
    123: "accentBorderCallout3",
    119: "callout3",
    112: "borderCallout3",
    116: "accentCallout3",
    124: "accentBorderCallout3",
    120: "callout3",
    183: "lineInv",
    166: "mathDivide",
    167: "mathEqual",
    164: "mathMinus",
    165: "mathMultiply",
    168: "mathNotEqual",
    163: "mathPlus",
    24: "moon",
    143: "nonIsoscelesTrapezoid",
    50: "notchedRightArrow",
    19: "noSmoking",
    6: "octagon",
    9: "ellipse",
    107: "wedgeEllipseCallout",
    2: "parallelogram",
    51: "homePlate",
    142: "pie",
    175: "pieWedge",
    28: "plaque",
    171: "plaqueTabs",
    39: "quadArrow",
    59: "quadArrowCallout",
    1: "rect",
    105: "wedgeRectCallout",
    12: "pentagon",
    33: "rightArrow",
    53: "rightArrowCallout",
    32: "rightBrace",
    30: "rightBracket",
    8: "rtTriangle",
    5: "roundRect",
    106: "wedgeRoundRectCallout",
    151: "round1Rect",
    153: "round2DiagRect",
    152: "round2SameRect",
    17: "smileyFace",
    155: "snip1Rect",
    157: "snip2DiagRect",
    156: "snip2SameRect",
    154: "snipRoundRect",
    170: "squareTabs",
    149: "star10",
    150: "star12",
    94: "star16",
    95: "star24",
    96: "star32",
    91: "star4",
    92: "star5",
    147: "star6",
    148: "star7",
    93: "star8",
    49: "stripedRightArrow",
    23: "sun",
    178: "swooshArrow",
    160: "teardrop",
    3: "trapezoid",
    55: "upArrowCallout",
    38: "upDownArrow",
    58: "upDownArrowCallout",
    97: "ribbon2",
    42: "uturnArrow",
    101: "verticalScroll",
    103: "wave",
}

MSO_AUTO_SHAPE_TYPE = MSO_SHAPE


class MSO_SHAPE_TYPE(IntEnum):
    """The kind of a shape, as reported by `Shape.shape_type`."""

    AUTO_SHAPE = 1
    CALLOUT = 2
    CANVAS = 20
    CHART = 3
    COMMENT = 4
    DIAGRAM = 21
    EMBEDDED_OLE_OBJECT = 7
    FORM_CONTROL = 8
    FREEFORM = 5
    GROUP = 6
    IGX_GRAPHIC = 24
    INK = 22
    INK_COMMENT = 23
    LINE = 9
    LINKED_OLE_OBJECT = 10
    LINKED_PICTURE = 11
    MEDIA = 16
    OLE_CONTROL_OBJECT = 12
    PICTURE = 13
    PLACEHOLDER = 14
    SCRIPT_ANCHOR = 18
    TABLE = 19
    TEXT_BOX = 17
    TEXT_EFFECT = 15
    WEB_VIDEO = 26
    MIXED = -2


class MSO_CONNECTOR_TYPE(IntEnum):
    """Connector geometries `add_connector` can author."""

    CURVE = 3
    ELBOW = 2
    STRAIGHT = 1
    MIXED = -2

    @property
    def xml_value(self) -> str:
        """The DrawingML preset name, empty for `MIXED`."""
        return _CONNECTORS[self.value]


_CONNECTORS: dict[int, str] = {
    3: "curvedConnector3",
    2: "bentConnector3",
    1: "line",
    -2: "",
}

MSO_CONNECTOR = MSO_CONNECTOR_TYPE


__all__ = [
    "MSO_AUTO_SHAPE_TYPE",
    "MSO_CONNECTOR",
    "MSO_CONNECTOR_TYPE",
    "MSO_SHAPE",
    "MSO_SHAPE_TYPE",
]
