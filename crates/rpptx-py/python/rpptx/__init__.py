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


from ._rpptx import Comment, CommentAuthor, CommentReply, Presentation

__all__ = [
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
    "RpptxError",
    "StaleElementError",
    "XmlError",
]
