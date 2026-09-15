"""Python bindings for rdocx."""

from .enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from .enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from .shared import Cm, Emu, Inches, Length, Mm, Pt, RGBColor


class RdocxError(Exception):
    """Base class for errors raised by rdocx."""


class PackageError(RdocxError):
    """An OPC package, file, or document-part operation failed."""


class XmlError(RdocxError):
    """WordprocessingML could not be parsed or serialized."""


class StaleElementError(RdocxError):
    """A held content handle was invalidated by structural mutation."""


class LayoutError(RdocxError):
    """Document layout or rendering failed."""


from ._rdocx import (
    BoundingBox,
    Cell,
    CellCollection,
    CellParagraphCollection,
    Comment,
    ComparisonDiagnostic,
    Document,
    Font,
    HeaderFooterVariant,
    Hyperlink,
    LayoutFragment,
    LayoutPage,
    Paragraph,
    ParagraphCollection,
    ParagraphFormat,
    Revision,
    Row,
    RowCollection,
    Run,
    RunCollection,
    RunPosition,
    RunRange,
    Section,
    Story,
    StoryItem,
    Style,
    Table,
    TableCollection,
    TocRebuildReport,
)

__all__ = [
    "BoundingBox",
    "Cm",
    "Cell",
    "CellCollection",
    "CellParagraphCollection",
    "Comment",
    "ComparisonDiagnostic",
    "Document",
    "Emu",
    "Font",
    "HeaderFooterVariant",
    "Hyperlink",
    "Inches",
    "LayoutError",
    "Length",
    "LayoutFragment",
    "LayoutPage",
    "Mm",
    "PackageError",
    "Paragraph",
    "ParagraphCollection",
    "ParagraphFormat",
    "Pt",
    "RGBColor",
    "RdocxError",
    "Revision",
    "Row",
    "RowCollection",
    "Run",
    "RunCollection",
    "RunPosition",
    "RunRange",
    "Section",
    "StaleElementError",
    "Story",
    "StoryItem",
    "Style",
    "Table",
    "TableCollection",
    "TocRebuildReport",
    "WD_ALIGN_PARAGRAPH",
    "WD_CELL_VERTICAL_ALIGNMENT",
    "WD_TABLE_ALIGNMENT",
    "WD_UNDERLINE",
    "XmlError",
]
