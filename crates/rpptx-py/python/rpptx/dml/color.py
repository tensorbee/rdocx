"""Colour values used by the public rpptx API."""


class RGBColor(tuple[int, int, int]):
    """An immutable red, green, blue colour triple."""

    def __new__(cls, r: int, g: int, b: int) -> "RGBColor":
        channels = (r, g, b)
        if any(not isinstance(channel, int) or not 0 <= channel <= 255 for channel in channels):
            raise ValueError("RGBColor() takes three integer values 0-255")
        return super().__new__(cls, channels)

    @classmethod
    def from_string(cls, rgb_hex_str: str) -> "RGBColor":
        """Create a colour from a six-digit hexadecimal string such as `3C2F80`."""

        if len(rgb_hex_str) != 6:
            raise ValueError("RGB colour strings must contain exactly six hexadecimal digits")
        try:
            return cls(*(int(rgb_hex_str[offset : offset + 2], 16) for offset in (0, 2, 4)))
        except ValueError as error:
            raise ValueError("RGB colour strings must contain only hexadecimal digits") from error

    def __str__(self) -> str:
        return "%02X%02X%02X" % self


__all__ = ["RGBColor"]
