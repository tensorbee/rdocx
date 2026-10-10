"""Python bindings for rdocx."""

from .enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from .enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_UNDERLINE
from .shared import Cm, Emu, Inches, Length, Mm, Pt, RGBColor, Twips


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


class ReplacementCountError(RdocxError):
    """A counted replacement matched a different number of times than expected.

    ``index`` is the position of the failing pair in a ``Document.replace_all``
    batch, and ``None`` for ``Document.try_replace_text``.
    """

    def __init__(
        self, message: str, expected: int, found: int, index: "int | None" = None
    ) -> None:
        super().__init__(message, expected, found, index)
        self.expected = expected
        self.found = found
        self.index = index

    def __str__(self) -> str:
        return str(self.args[0])


from ._rdocx import (
    Bookmark,
    BoundingBox,
    Cell,
    CellCollection,
    CellParagraphCollection,
    Comment,
    ComparisonDiagnostic,
    ContentFragment,
    CoreProperties,
    Document,
    Font,
    HeaderFooter,
    HeaderFooterCell,
    HeaderFooterRow,
    HeaderFooterTable,
    HeaderFooterVariant,
    Hyperlink,
    LayoutFragment,
    LayoutBackedFieldUpdateReport,
    LayoutPage,
    ListLevel,
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
    Settings,
    Story,
    StoryItem,
    StoryRunPosition,
    StoryRunRange,
    Style,
    SvgDiagnostic,
    SvgRenderResult,
    Table,
    TableCollection,
    TocRebuildReport,
)

__all__ = [
    "Bookmark",
    "BoundingBox",
    "Cm",
    "Cell",
    "CellCollection",
    "CellParagraphCollection",
    "Comment",
    "ComparisonDiagnostic",
    "ContentFragment",
    "CoreProperties",
    "Document",
    "Emu",
    "Font",
    "HeaderFooter",
    "HeaderFooterCell",
    "HeaderFooterRow",
    "HeaderFooterTable",
    "HeaderFooterVariant",
    "Hyperlink",
    "Inches",
    "LayoutError",
    "Length",
    "LayoutFragment",
    "LayoutBackedFieldUpdateReport",
    "LayoutPage",
    "ListLevel",
    "Mm",
    "PackageError",
    "Paragraph",
    "ParagraphCollection",
    "ParagraphFormat",
    "Pt",
    "RGBColor",
    "RdocxError",
    "ReplacementCountError",
    "Revision",
    "Row",
    "RowCollection",
    "Run",
    "RunCollection",
    "RunPosition",
    "RunRange",
    "Section",
    "Settings",
    "StaleElementError",
    "Story",
    "StoryItem",
    "StoryRunPosition",
    "StoryRunRange",
    "Style",
    "SvgDiagnostic",
    "SvgRenderResult",
    "Table",
    "TableCollection",
    "TocRebuildReport",
    "Twips",
    "WD_ALIGN_PARAGRAPH",
    "WD_BREAK",
    "WD_CELL_VERTICAL_ALIGNMENT",
    "WD_ROW_HEIGHT_RULE",
    "WD_TABLE_ALIGNMENT",
    "WD_UNDERLINE",
    "XmlError",
]
