from pathlib import Path
from typing import TYPE_CHECKING, assert_type

from rdocx import (
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
    Inches,
    LayoutFragment,
    LayoutPage,
    RGBColor,
    Paragraph,
    ParagraphCollection,
    ParagraphFormat,
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


def exercise_rdocx_types(path: Path) -> None:
    document = Document(path)
    opened: Document = Document.open(path)
    loaded: Document = Document.from_bytes(b"")
    paragraph: Paragraph = document.add_paragraph("typed")
    run: Run = paragraph.add_run(" run")
    font: Font = run.font
    font.bold = True
    font.size = Inches(1)
    color = RGBColor(1, 2, 3)
    assert_type(color[0], int)
    channels: tuple[int, int, int] = color
    paragraph_format: ParagraphFormat = paragraph.paragraph_format
    paragraph_format.keep_together = None
    paragraphs: ParagraphCollection = document.paragraphs
    first: Paragraph = paragraphs[0]
    sliced: list[Paragraph] = paragraphs[:]
    for item in paragraphs:
        item.text
    table: Table = document.add_table(1, 1)
    row: Row = table.rows[0]
    cell: Cell = row.cells[0]
    cell.text = first.text
    package_bytes: bytes = loaded.to_bytes()
    pdf_bytes: bytes = opened.to_pdf()
    pages: list[bytes] = opened.render_all_pages()
    maybe_page: bytes | None = opened.render_page_to_png(0)
    document.save(path)
    document.remove_content(0)
    position = RunPosition(body_index=0, run_index=0)
    range_ = RunRange(start=position, end=RunPosition(body_index=0, run_index=1))
    comment_id: int = document.add_comment(
        range_, author="Ada", text="review", initials=None
    )
    reply_id: int = document.reply_to(comment_id, author="Grace", text="done")
    comments: tuple[Comment, ...] = document.comments
    sections: tuple[Section, ...] = document.sections
    styles: tuple[Style, ...] = document.styles
    stories: tuple[Story, ...] = document.stories
    story_items: tuple[StoryItem, ...] = document.story_items
    variants: tuple[HeaderFooterVariant, ...] = document.header_footer_variants
    hyperlinks: tuple[Hyperlink, ...] = document.hyperlinks
    resolved: bool = document.resolve_comment(comment_id)
    removed: bool = document.remove_comment(reply_id)
    diagnostics: tuple[ComparisonDiagnostic, ...] = document.compare(
        opened, author="Ada", timestamp="2026-09-14T09:00:00Z"
    )
    fragments: tuple[LayoutFragment, ...] = document.layout()
    maybe_layout_page: LayoutPage | None = document.layout_page(0)
    report: TocRebuildReport = document.rebuild_toc()
    document.set_header("Header")
    document.set_footer("Footer")
    document.set_story_text(story_items[0], "edited")
    document.add_hyperlink_to_story(stories[0], "home", "https://example.com/")
    link_run: Run = first.add_hyperlink("docs", "https://example.com/docs")
    assert_type(story_items[0].xml, bytes)
    if fragments:
        bounds: BoundingBox = fragments[0].bounds
        assert_type(bounds.width, float)
    assert_type(comments[0].date, str | None)
    assert_type(sections[0].page_width, int | None)
    assert_type(styles[0].style_type, str)
    assert_type(stories[0].owner_index, int)
    assert_type(story_items[0].index_path, tuple[int, ...])
    assert_type(variants[0].story, Story | None)
    assert_type(hyperlinks[0].url, str | None)
    assert_type(report.entry_count, int)
    package_bytes, pdf_bytes, pages, maybe_page, sliced, channels


if TYPE_CHECKING:
    Cell()  # type: ignore[call-arg]
    CellCollection()  # type: ignore[call-arg]
    CellParagraphCollection()  # type: ignore[call-arg]
    Font()  # type: ignore[call-arg]
    Paragraph()  # type: ignore[call-arg]
    ParagraphCollection()  # type: ignore[call-arg]
    ParagraphFormat()  # type: ignore[call-arg]
    Row()  # type: ignore[call-arg]
    RowCollection()  # type: ignore[call-arg]
    Run()  # type: ignore[call-arg]
    RunCollection()  # type: ignore[call-arg]
    Table()  # type: ignore[call-arg]
    TableCollection()  # type: ignore[call-arg]
