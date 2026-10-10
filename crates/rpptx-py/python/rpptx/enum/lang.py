"""Language enumeration, mirroring python-pptx `pptx.enum.lang`."""

from enum import IntEnum


class MSO_LANGUAGE_ID(IntEnum):
    """A language, as `Font.language_id` reads and writes it.

    Each member carries the python-pptx 1.0.2 value and stands for the tag in
    `_TAGS`, which `Font.language` reads and writes as a string. Unlike
    python-pptx, GAELIC_SCOTLAND writes `gd-GB` rather than `en-US`, and
    NONE, MIXED, and NO_PROOFING name no tag.
    """

    NONE = 0
    AFRIKAANS = 1078
    ALBANIAN = 1052
    AMHARIC = 1118
    ARABIC = 1025
    ARABIC_ALGERIA = 5121
    ARABIC_BAHRAIN = 15361
    ARABIC_EGYPT = 3073
    ARABIC_IRAQ = 2049
    ARABIC_JORDAN = 11265
    ARABIC_KUWAIT = 13313
    ARABIC_LEBANON = 12289
    ARABIC_LIBYA = 4097
    ARABIC_MOROCCO = 6145
    ARABIC_OMAN = 8193
    ARABIC_QATAR = 16385
    ARABIC_SYRIA = 10241
    ARABIC_TUNISIA = 7169
    ARABIC_UAE = 14337
    ARABIC_YEMEN = 9217
    ARMENIAN = 1067
    ASSAMESE = 1101
    AZERI_CYRILLIC = 2092
    AZERI_LATIN = 1068
    BASQUE = 1069
    BELGIAN_DUTCH = 2067
    BELGIAN_FRENCH = 2060
    BENGALI = 1093
    BOSNIAN = 4122
    BOSNIAN_BOSNIA_HERZEGOVINA_CYRILLIC = 8218
    BOSNIAN_BOSNIA_HERZEGOVINA_LATIN = 5146
    BRAZILIAN_PORTUGUESE = 1046
    BULGARIAN = 1026
    BURMESE = 1109
    BYELORUSSIAN = 1059
    CATALAN = 1027
    CHEROKEE = 1116
    CHINESE_HONG_KONG_SAR = 3076
    CHINESE_MACAO_SAR = 5124
    CHINESE_SINGAPORE = 4100
    CROATIAN = 1050
    CZECH = 1029
    DANISH = 1030
    DIVEHI = 1125
    DUTCH = 1043
    EDO = 1126
    ENGLISH_AUS = 3081
    ENGLISH_BELIZE = 10249
    ENGLISH_CANADIAN = 4105
    ENGLISH_CARIBBEAN = 9225
    ENGLISH_INDONESIA = 14345
    ENGLISH_IRELAND = 6153
    ENGLISH_JAMAICA = 8201
    ENGLISH_NEW_ZEALAND = 5129
    ENGLISH_PHILIPPINES = 13321
    ENGLISH_SOUTH_AFRICA = 7177
    ENGLISH_TRINIDAD_TOBAGO = 11273
    ENGLISH_UK = 2057
    ENGLISH_US = 1033
    ENGLISH_ZIMBABWE = 12297
    ESTONIAN = 1061
    FAEROESE = 1080
    FARSI = 1065
    FILIPINO = 1124
    FINNISH = 1035
    FRANCH_CONGO_DRC = 9228
    FRENCH = 1036
    FRENCH_CAMEROON = 11276
    FRENCH_CANADIAN = 3084
    FRENCH_COTED_IVOIRE = 12300
    FRENCH_HAITI = 15372
    FRENCH_LUXEMBOURG = 5132
    FRENCH_MALI = 13324
    FRENCH_MONACO = 6156
    FRENCH_MOROCCO = 14348
    FRENCH_REUNION = 8204
    FRENCH_SENEGAL = 10252
    FRENCH_WEST_INDIES = 7180
    FRISIAN_NETHERLANDS = 1122
    FULFULDE = 1127
    GAELIC_IRELAND = 2108
    GAELIC_SCOTLAND = 1084
    GALICIAN = 1110
    GEORGIAN = 1079
    GERMAN = 1031
    GERMAN_AUSTRIA = 3079
    GERMAN_LIECHTENSTEIN = 5127
    GERMAN_LUXEMBOURG = 4103
    GREEK = 1032
    GUARANI = 1140
    GUJARATI = 1095
    HAUSA = 1128
    HAWAIIAN = 1141
    HEBREW = 1037
    HINDI = 1081
    HUNGARIAN = 1038
    IBIBIO = 1129
    ICELANDIC = 1039
    IGBO = 1136
    INDONESIAN = 1057
    INUKTITUT = 1117
    ITALIAN = 1040
    JAPANESE = 1041
    KANNADA = 1099
    KANURI = 1137
    KASHMIRI = 1120
    KASHMIRI_DEVANAGARI = 2144
    KAZAKH = 1087
    KHMER = 1107
    KIRGHIZ = 1088
    KONKANI = 1111
    KOREAN = 1042
    LAO = 1108
    LATIN = 1142
    LATVIAN = 1062
    LITHUANIAN = 1063
    MACEDONINAN_FYROM = 1071
    MALAY_BRUNEI_DARUSSALAM = 2110
    MALAYALAM = 1100
    MALAYSIAN = 1086
    MALTESE = 1082
    MANIPURI = 1112
    MAORI = 1153
    MARATHI = 1102
    MEXICAN_SPANISH = 2058
    MONGOLIAN = 1104
    NEPALI = 1121
    NO_PROOFING = 1024
    NORWEGIAN_BOKMOL = 1044
    NORWEGIAN_NYNORSK = 2068
    ORIYA = 1096
    OROMO = 1138
    PASHTO = 1123
    POLISH = 1045
    PORTUGUESE = 2070
    PUNJABI = 1094
    QUECHUA_BOLIVIA = 1131
    QUECHUA_ECUADOR = 2155
    QUECHUA_PERU = 3179
    RHAETO_ROMANIC = 1047
    ROMANIAN = 1048
    ROMANIAN_MOLDOVA = 2072
    RUSSIAN = 1049
    RUSSIAN_MOLDOVA = 2073
    SAMI_LAPPISH = 1083
    SANSKRIT = 1103
    SEPEDI = 1132
    SERBIAN_BOSNIA_HERZEGOVINA_CYRILLIC = 7194
    SERBIAN_BOSNIA_HERZEGOVINA_LATIN = 6170
    SERBIAN_CYRILLIC = 3098
    SERBIAN_LATIN = 2074
    SESOTHO = 1072
    SIMPLIFIED_CHINESE = 2052
    SINDHI = 1113
    SINDHI_PAKISTAN = 2137
    SINHALESE = 1115
    SLOVAK = 1051
    SLOVENIAN = 1060
    SOMALI = 1143
    SORBIAN = 1070
    SPANISH = 1034
    SPANISH_ARGENTINA = 11274
    SPANISH_BOLIVIA = 16394
    SPANISH_CHILE = 13322
    SPANISH_COLOMBIA = 9226
    SPANISH_COSTA_RICA = 5130
    SPANISH_DOMINICAN_REPUBLIC = 7178
    SPANISH_ECUADOR = 12298
    SPANISH_EL_SALVADOR = 17418
    SPANISH_GUATEMALA = 4106
    SPANISH_HONDURAS = 18442
    SPANISH_MODERN_SORT = 3082
    SPANISH_NICARAGUA = 19466
    SPANISH_PANAMA = 6154
    SPANISH_PARAGUAY = 15370
    SPANISH_PERU = 10250
    SPANISH_PUERTO_RICO = 20490
    SPANISH_URUGUAY = 14346
    SPANISH_VENEZUELA = 8202
    SWAHILI = 1089
    SWEDISH = 1053
    SWEDISH_FINLAND = 2077
    SWISS_FRENCH = 4108
    SWISS_GERMAN = 2055
    SWISS_ITALIAN = 2064
    SYRIAC = 1114
    TAJIK = 1064
    TAMAZIGHT = 1119
    TAMAZIGHT_LATIN = 2143
    TAMIL = 1097
    TATAR = 1092
    TELUGU = 1098
    THAI = 1054
    TIBETAN = 1105
    TIGRIGNA_ERITREA = 2163
    TIGRIGNA_ETHIOPIC = 1139
    TRADITIONAL_CHINESE = 1028
    TSONGA = 1073
    TSWANA = 1074
    TURKISH = 1055
    TURKMEN = 1090
    UKRAINIAN = 1058
    URDU = 1056
    UZBEK_CYRILLIC = 2115
    UZBEK_LATIN = 1091
    VENDA = 1075
    VIETNAMESE = 1066
    WELSH = 1106
    XHOSA = 1076
    YI = 1144
    YIDDISH = 1085
    YORUBA = 1130
    ZULU = 1077
    MIXED = -2


_TAGS: dict[int, str] = {
    1078: "af-ZA",
    1052: "sq-AL",
    1118: "am-ET",
    1025: "ar-SA",
    5121: "ar-DZ",
    15361: "ar-BH",
    3073: "ar-EG",
    2049: "ar-IQ",
    11265: "ar-JO",
    13313: "ar-KW",
    12289: "ar-LB",
    4097: "ar-LY",
    6145: "ar-MA",
    8193: "ar-OM",
    16385: "ar-QA",
    10241: "ar-SY",
    7169: "ar-TN",
    14337: "ar-AE",
    9217: "ar-YE",
    1067: "hy-AM",
    1101: "as-IN",
    2092: "az-AZ",
    1068: "az-Latn-AZ",
    1069: "eu-ES",
    2067: "nl-BE",
    2060: "fr-BE",
    1093: "bn-IN",
    4122: "hr-BA",
    8218: "bs-BA",
    5146: "bs-Latn-BA",
    1046: "pt-BR",
    1026: "bg-BG",
    1109: "my-MM",
    1059: "be-BY",
    1027: "ca-ES",
    1116: "chr-US",
    3076: "zh-HK",
    5124: "zh-MO",
    4100: "zh-SG",
    1050: "hr-HR",
    1029: "cs-CZ",
    1030: "da-DK",
    1125: "div-MV",
    1043: "nl-NL",
    1126: "bin-NG",
    3081: "en-AU",
    10249: "en-BZ",
    4105: "en-CA",
    9225: "en-CB",
    14345: "en-ID",
    6153: "en-IE",
    8201: "en-JA",
    5129: "en-NZ",
    13321: "en-PH",
    7177: "en-ZA",
    11273: "en-TT",
    2057: "en-GB",
    1033: "en-US",
    12297: "en-ZW",
    1061: "et-EE",
    1080: "fo-FO",
    1065: "fa-IR",
    1124: "fil-PH",
    1035: "fi-FI",
    9228: "fr-CD",
    1036: "fr-FR",
    11276: "fr-CM",
    3084: "fr-CA",
    12300: "fr-CI",
    15372: "fr-HT",
    5132: "fr-LU",
    13324: "fr-ML",
    6156: "fr-MC",
    14348: "fr-MA",
    8204: "fr-RE",
    10252: "fr-SN",
    7180: "fr-WINDIES",
    1122: "fy-NL",
    1127: "ff-NG",
    2108: "ga-IE",
    1084: "gd-GB",
    1110: "gl-ES",
    1079: "ka-GE",
    1031: "de-DE",
    3079: "de-AT",
    5127: "de-LI",
    4103: "de-LU",
    1032: "el-GR",
    1140: "gn-PY",
    1095: "gu-IN",
    1128: "ha-NG",
    1141: "haw-US",
    1037: "he-IL",
    1081: "hi-IN",
    1038: "hu-HU",
    1129: "ibb-NG",
    1039: "is-IS",
    1136: "ig-NG",
    1057: "id-ID",
    1117: "iu-Cans-CA",
    1040: "it-IT",
    1041: "ja-JP",
    1099: "kn-IN",
    1137: "kr-NG",
    1120: "ks-Arab",
    2144: "ks-Deva",
    1087: "kk-KZ",
    1107: "kh-KH",
    1088: "ky-KG",
    1111: "kok-IN",
    1042: "ko-KR",
    1108: "lo-LA",
    1142: "la-Latn",
    1062: "lv-LV",
    1063: "lt-LT",
    1071: "mk-MK",
    2110: "ms-BN",
    1100: "ml-IN",
    1086: "ms-MY",
    1082: "mt-MT",
    1112: "mni-IN",
    1153: "mi-NZ",
    1102: "mr-IN",
    2058: "es-MX",
    1104: "mn-MN",
    1121: "ne-NP",
    1044: "nb-NO",
    2068: "nn-NO",
    1096: "or-IN",
    1138: "om-Ethi-ET",
    1123: "ps-AF",
    1045: "pl-PL",
    2070: "pt-PT",
    1094: "pa-IN",
    1131: "quz-BO",
    2155: "quz-EC",
    3179: "quz-PE",
    1047: "rm-CH",
    1048: "ro-RO",
    2072: "ro-MO",
    1049: "ru-RU",
    2073: "ru-MO",
    1083: "se-NO",
    1103: "sa-IN",
    1132: "ns-ZA",
    7194: "sr-BA",
    6170: "sr-Latn-BA",
    3098: "sr-SP",
    2074: "sr-Latn-CS",
    1072: "st-ZA",
    2052: "zh-CN",
    1113: "sd-Deva-IN",
    2137: "sd-Arab-PK",
    1115: "si-LK",
    1051: "sk-SK",
    1060: "sl-SI",
    1143: "so-SO",
    1070: "wen-DE",
    1034: "es-ES_tradnl",
    11274: "es-AR",
    16394: "es-BO",
    13322: "es-CL",
    9226: "es-CO",
    5130: "es-CR",
    7178: "es-DO",
    12298: "es-EC",
    17418: "es-SV",
    4106: "es-GT",
    18442: "es-HN",
    3082: "es-ES",
    19466: "es-NI",
    6154: "es-PA",
    15370: "es-PY",
    10250: "es-PE",
    20490: "es-PR",
    14346: "es-UR",
    8202: "es-VE",
    1089: "sw-KE",
    1053: "sv-SE",
    2077: "sv-FI",
    4108: "fr-CH",
    2055: "de-CH",
    2064: "it-CH",
    1114: "syr-SY",
    1064: "tg-TJ",
    1119: "tzm-Arab-MA",
    2143: "tmz-DZ",
    1097: "ta-IN",
    1092: "tt-RU",
    1098: "te-IN",
    1054: "th-TH",
    1105: "bo-CN",
    2163: "ti-ER",
    1139: "ti-ET",
    1028: "zh-TW",
    1073: "ts-ZA",
    1074: "tn-ZA",
    1055: "tr-TR",
    1090: "tk-TM",
    1058: "uk-UA",
    1056: "ur-PK",
    2115: "uz-UZ",
    1091: "uz-Latn-UZ",
    1075: "ve-ZA",
    1066: "vi-VN",
    1106: "cy-GB",
    1076: "xh-ZA",
    1144: "ii-CN",
    1085: "yi-Hebr",
    1130: "yo-NG",
    1077: "zu-ZA",
}

_MEMBERS: dict[str, MSO_LANGUAGE_ID] = {}
for _value, _tag in _TAGS.items():
    _MEMBERS.setdefault(_tag.lower(), MSO_LANGUAGE_ID(_value))


__all__ = ["MSO_LANGUAGE_ID"]
