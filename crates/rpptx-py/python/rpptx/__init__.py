"""Python bindings for rpptx."""

from .dml.color import RGBColor
from .enum.shapes import MSO_SHAPE
from .enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, MSO_UNDERLINE, PP_ALIGN
from .util import Inches, Length, Pt


class RpptxError(Exception):
    """Base class for errors raised by rpptx."""


class PackageError(RpptxError):
    """An OPC package, file, or presentation operation failed."""


class XmlError(RpptxError):
    """PresentationML could not be parsed or serialized."""


class StaleElementError(RpptxError):
    """A held content handle was invalidated by structural mutation."""


class ReplacementCountError(RpptxError):
    """A counted replacement matched a different number of times than expected."""

    def __init__(self, message: str, expected: int, found: int) -> None:
        super().__init__(message, expected, found)
        self.expected = expected
        self.found = found

    def __str__(self) -> str:
        return str(self.args[0])


from ._rpptx import Comment, CommentAuthor, CommentReply, Presentation
from ._rpptx import BoundingBox, TextFrameLayout, TextLineLayout, ValidationIssue
from ._rpptx import AutofitResult

__all__ = [
    "AutofitResult",
    "BoundingBox",
    "Inches",
    "Length",
    "MSO_ANCHOR",
    "MSO_AUTO_SIZE",
    "MSO_SHAPE",
    "MSO_UNDERLINE",
    "PP_ALIGN",
    "PackageError",
    "Comment",
    "CommentAuthor",
    "CommentReply",
    "Presentation",
    "Pt",
    "RGBColor",
    "ReplacementCountError",
    "RpptxError",
    "StaleElementError",
    "TextFrameLayout",
    "TextLineLayout",
    "ValidationIssue",
    "XmlError",
]
