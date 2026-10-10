"""Unit and color values shared by the public rdocx API."""

from collections.abc import Callable
from typing import Any, SupportsIndex, overload


def _truncating_division(value: int, divisor: int) -> int:
    """Divide integers with truncation toward zero."""

    quotient = abs(value) // divisor
    return -quotient if value < 0 else quotient


class Length(int):
    """An integer length measured in English Metric Units (EMU)."""

    _EMUS_PER_INCH = 914400
    _EMUS_PER_CM = 360000
    _EMUS_PER_MM = 36000
    _EMUS_PER_PT = 12700
    _EMUS_PER_TWIP = 635

    @property
    def inches(self) -> float:
        """Return this length in inches."""

        return int(self) / self._EMUS_PER_INCH

    @property
    def cm(self) -> float:
        """Return this length in centimetres."""

        return int(self) / self._EMUS_PER_CM

    @property
    def mm(self) -> float:
        """Return this length in millimetres."""

        return int(self) / self._EMUS_PER_MM

    @property
    def pt(self) -> float:
        """Return this length in points."""

        return int(self) / self._EMUS_PER_PT

    @property
    def emu(self) -> int:
        """Return this length in English Metric Units."""

        return int(self)

    @property
    def twips(self) -> int:
        """Return this length in twips, truncated toward zero."""

        return _truncating_division(int(self), self._EMUS_PER_TWIP)


class Inches(Length):
    """A length constructed from inches."""

    def __new__(cls, inches: float) -> "Inches":
        return super().__new__(cls, int(inches * cls._EMUS_PER_INCH))


class Cm(Length):
    """A length constructed from centimetres."""

    def __new__(cls, cm: float) -> "Cm":
        return super().__new__(cls, int(cm * cls._EMUS_PER_CM))


class Mm(Length):
    """A length constructed from millimetres."""

    def __new__(cls, mm: float) -> "Mm":
        return super().__new__(cls, int(mm * cls._EMUS_PER_MM))


class Pt(Length):
    """A length constructed from points."""

    def __new__(cls, points: float) -> "Pt":
        return super().__new__(cls, int(points * cls._EMUS_PER_PT))


class Emu(Length):
    """A length constructed directly from English Metric Units."""

    def __new__(cls, emu: float) -> "Emu":
        return super().__new__(cls, int(emu))


class RGBColor(tuple[int, int, int]):
    """An immutable red, green, blue color triple."""

    __slots__ = ()

    def __new__(cls, r: int, g: int, b: int) -> "RGBColor":
        channels = (r, g, b)
        if any(not isinstance(channel, int) or not 0 <= channel <= 255 for channel in channels):
            raise ValueError("RGB color channels must be integers from 0 to 255")
        return super().__new__(cls, channels)

    @classmethod
    def from_string(cls, value: str) -> "RGBColor":
        """Create a color from a six-character hexadecimal string."""

        if len(value) != 6:
            raise ValueError("RGB color strings must contain exactly six hexadecimal digits")
        try:
            return cls(*(int(value[offset : offset + 2], 16) for offset in (0, 2, 4)))
        except ValueError as error:
            raise ValueError("RGB color strings must contain only hexadecimal digits") from error

    def __str__(self) -> str:
        return "%02X%02X%02X" % self

    def __setattr__(self, name: str, value: object) -> None:
        # Without this, `font.color.rgb = RGBColor(...)` on a run that already
        # has a color would set an attribute on a copy and change nothing.
        raise AttributeError(
            f"RGBColor is immutable, so setting {name} would change nothing: "
            "assign a new color instead, for example font.color = RGBColor(0xC0, 0x00, 0x00)"
        )


class _Styles(tuple[Any, ...]):
    """``Document.styles``: the style snapshots in a tuple that is also
    indexed by style name or ID, as python-docx's ``styles['Normal']``."""

    __slots__ = ()

    def _find(self, key: str) -> Any:
        lowered = key.lower()
        checks: tuple[Callable[[Any], bool], ...] = (
            lambda style: style.style_id == key,
            lambda style: style.name == key,
            lambda style: (style.name or "").lower() == lowered,
        )
        for matches in checks:
            for style in self:
                if matches(style):
                    return style
        raise KeyError(f"no style with name or ID '{key}'")

    @overload
    def __getitem__(self, key: SupportsIndex) -> Any: ...
    @overload
    def __getitem__(self, key: slice) -> tuple[Any, ...]: ...
    @overload
    def __getitem__(self, key: str) -> Any: ...
    def __getitem__(self, key: "SupportsIndex | slice | str") -> Any:
        if isinstance(key, str):
            return self._find(key)
        return super().__getitem__(key)

    def __contains__(self, key: object) -> bool:
        if isinstance(key, str):
            try:
                self._find(key)
            except KeyError:
                return False
            return True
        return super().__contains__(key)


__all__ = ["Length", "Inches", "Cm", "Mm", "Pt", "Emu", "RGBColor"]
