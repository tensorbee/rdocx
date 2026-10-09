from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Literal, assert_type

from rdocx import (
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
    HeaderFooterVariant,
    Hyperlink,
    Inches,
    Length,
    LayoutFragment,
    LayoutBackedFieldUpdateReport,
    LayoutPage,
    ListLevel,
    RGBColor,
    ReplacementCountError,
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
    StoryRunPosition,
    StoryRunRange,
    Style,
    SvgDiagnostic,
    SvgRenderResult,
    Table,
    TableCollection,
    TocRebuildReport,
    WD_ROW_HEIGHT_RULE,
)


def exercise_rdocx_types(path: Path) -> None:
    document = Document(path)
    # Comment-aware removal keeps the existing bool surface.
    assert_type(document.remove_content(1000), bool)
    opened: Document = Document.open(path)
    loaded: Document = Document.from_bytes(b"")
    paragraph: Paragraph = document.add_paragraph("typed")
    run: Run = paragraph.add_run(" run")
    paragraph.text = None
    paragraph.text = "retyped\tline"
    assert_type(paragraph.text, str)
    paragraph.style = "Heading1"
    paragraph.numbering = (1, 2)
    assert_type(paragraph.style, str)
    assert_type(paragraph.numbering, tuple[int, int])
    run.style_id = "Strong"
    assert_type(run.style_id, str)
    font: Font = run.font
    font.bold = True
    font.size = Inches(1)
    font.highlight = "yellow"
    font.shading = "FFFF00"
    color = RGBColor(1, 2, 3)
    assert_type(color[0], int)
    channels: tuple[int, int, int] = color
    paragraph_format: ParagraphFormat = paragraph.paragraph_format
    paragraph_format.keep_together = None
    split_boundary: int = document.split_run(0, 0, 1)
    handle_boundary: int = document.split_run(paragraph, 0, 1)
    document.paragraphs[0].runs[0].remove()
    paragraphs: ParagraphCollection = document.paragraphs
    first: Paragraph = paragraphs[0]
    sliced: list[Paragraph] = paragraphs[:]
    for item in paragraphs:
        item.text
    table: Table = document.add_table(1, 1)
    row: Row = table.rows[0]
    cloned_row: Row = table.clone_row(0)
    table.remove_row(0)
    cell: Cell = row.cells[0]
    cell.text = first.text
    assert_type(table.border("top"), tuple[str, int | None, str | None] | None)
    table.set_borders("single", size=4, color="000000")
    table.set_border("insideV", "dashed", size=8, color="FF0000")
    assert_type(
        table.cell_margins,
        tuple[Length | None, Length | None, Length | None, Length | None] | None,
    )
    table.set_cell_margins(top=0, right=Inches(0.1), bottom=0, left=Inches(0.1))
    assert_type(table.grid_widths, tuple[Length, ...])
    table.grid_widths = [Inches(1)]
    table.set_column_width(0, Inches(2))
    assert_type(row.height, Length | None)
    assert_type(row.height_rule, WD_ROW_HEIGHT_RULE | None)
    assert_type(row.cant_split, bool | None)
    assert_type(row.is_header, bool | None)
    row.height = Inches(0.5)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    row.cant_split = True
    row.is_header = None
    assert_type(cell.shading, str | None)
    assert_type(cell.border("bottom"), tuple[str, int | None, str | None] | None)
    assert_type(
        cell.margins,
        tuple[Length | None, Length | None, Length | None, Length | None] | None,
    )
    cell.shading = "D9D9D9"
    cell.set_border("bottom", "double", size=6, color="auto")
    cell.set_margins(top=0, right=0, bottom=0, left=0)
    inserted_table: Table = document.insert_table(0, 2, 2)
    inserted_table.set_cell_grid_span(0, 0, 2)
    inserted_table.set_cell_grid_span(0, -1, None)
    inserted_table.set_cell_vertical_merge(0, 0, "restart")
    inserted_table.set_cell_vertical_merge(1, 0, None)
    assert_type(cell.grid_span, int)
    assert_type(cell.vertical_merge, Literal["restart", "continue"] | None)
    package_bytes: bytes = loaded.to_bytes()
    pdf_bytes: bytes = opened.to_pdf()
    tracked_pdf: bytes = opened.to_pdf(revision_view="tracked")
    tracked_pages: list[bytes] | bytes = opened.render_pages(revision_view="tracked")
    pages: list[bytes] = opened.render_all_pages()
    maybe_page: bytes | None = opened.render_page_to_png(0)
    font_pdf: bytes = opened.to_pdf(fonts=[("Carlito", b"font")], font_dir=path)
    archival_pdf: bytes = opened.to_pdfa_deterministic("pdfa-3b")
    svg_page: SvgRenderResult | None = opened.render_page_to_svg(0)
    if svg_page is not None:
        assert_type(svg_page.svg, str)
        svg_diagnostics: tuple[SvgDiagnostic, ...] = svg_page.diagnostics
    document.save(path)
    document.remove_content(0)
    position = RunPosition(body_index=0, run_index=0)
    range_ = RunRange(start=position, end=RunPosition(body_index=0, run_index=1))
    comment_id: int = document.add_comment(
        range_,
        author="Ada",
        text="review",
        initials=None,
        date="2026-09-16T10:15:30Z",
    )
    text_comment_id: int = document.add_comment_on_text(
        "review", author="Ada", text="here", occurrence=0, initials=None, date=None
    )
    reply_id: int = document.reply_to(
        comment_id,
        author="Grace",
        text="done",
        date="2026-09-16T11:00:00+01:00",
    )
    bookmark_id: int = document.add_bookmark("target", range_)
    bookmarks: tuple[Bookmark, ...] = document.bookmarks
    assert_type(bookmarks[0].direct_range, RunRange | None)
    run.add_tab()
    run.add_field("PAGEREF target \\h", "1")
    document.insert_toc(0, max_level=2)
    comments: tuple[Comment, ...] = document.comments
    sections: tuple[Section, ...] = document.sections
    updated_section: Section = document.update_section(
        0,
        orientation="landscape",
        margin_top=Inches(0.5),
        column_count=2,
        column_spacing=Inches(0.25),
        different_first_page=True,
        break_type="continuous",
    )
    document.insert_section(1)
    document.remove_section(1)
    styles: tuple[Style, ...] = document.styles
    note: Style = document.add_style("Note", "paragraph", based_on="Normal")
    document.add_style(
        "Boxed Note",
        style_id="BoxedNote",
        next_style="Normal",
        font_name="Arial",
        font_size=Inches(0.25),
        bold=True,
        italic=None,
        color=RGBColor(0x11, 0x22, 0x33),
        space_before=Inches(0.1),
        space_after=None,
        left_indent=Inches(0.5),
        right_indent=None,
        first_line_indent=-Inches(0.25),
    )
    list_level = ListLevel(format="decimal", text="%1.", start=1, left_indent=Inches(0.5))
    assert_type(list_level.format, str)
    assert_type(list_level.text, str | None)
    assert_type(list_level.hanging_indent, int | None)
    definition_id: int = document.add_numbering_definition([list_level, ListLevel()])
    num_id: int = document.add_numbering_instance(definition_id)
    document.link_style_to_numbering(note.style_id, num_id, 0)
    document.set_default_style("Normal")
    style_removed: bool = document.remove_style("BoxedNote")
    stories: tuple[Story, ...] = document.stories
    image_data: bytes | None = document.image_data("rId1")
    document.replace_image("rId1", b"image")
    document.replace_image_for_story(stories[0], "rId1", b"image")
    resized: int = document.set_picture_size("rId1", width=Inches(1), height=Inches(1))
    story_items: tuple[StoryItem, ...] = document.story_items
    inserted_picture: StoryItem = document.add_picture(
        b"png", "image.png", Inches(1), Inches(1), after=story_items[0]
    )
    story_position = StoryRunPosition(item=story_items[0], run_index=0)
    handle_position = StoryRunPosition(paragraph=document.paragraphs[0], run_index=0)
    story_range = StoryRunRange(start=story_position, end=handle_position)
    assert_type(document.move_comment(comment_id, story_range), None)
    assert_type(document.move_comment_to_text(comment_id, "target", occurrence=0), None)
    story_comment_id: int = document.add_comment(
        story_range, author="Ada", text="story review"
    )
    direct_body_index: int | None = story_items[0].direct_body_index if story_items else None
    variants: tuple[HeaderFooterVariant, ...] = document.header_footer_variants
    hyperlinks: tuple[Hyperlink, ...] = document.hyperlinks
    document.set_hyperlink_url(hyperlinks[0], "https://example.org/new")
    document.remove_hyperlink(hyperlinks[0])
    resolved: bool = document.resolve_comment(comment_id)
    removed: bool = document.remove_comment(reply_id)
    diagnostics: tuple[ComparisonDiagnostic, ...] = document.compare(
        opened, author="Ada", timestamp="2026-09-14T09:00:00Z"
    )
    optioned: tuple[ComparisonDiagnostic, ...] = document.compare(
        opened,
        "Ada",
        "2026-09-14T09:00:00Z",
        granularity="word",
        ignore_formatting=True,
        ignore_whitespace=True,
        ignore_fields=True,
        ignore_comments=True,
        ignored_stories=("header", "text_box"),
    )
    fragments: tuple[LayoutFragment, ...] = document.layout()
    maybe_layout_page: LayoutPage | None = document.layout_page(0)
    report: TocRebuildReport = document.rebuild_toc()
    core: CoreProperties = document.core_properties
    core.title = "Typed title"
    core.author = None
    core.created = datetime(2026, 9, 29, tzinfo=timezone.utc)
    core.last_printed = None
    core.revision = 2
    assert_type(core.title, str)
    assert_type(core.comments, str)
    assert_type(core.modified, datetime | None)
    assert_type(core.revision, int)
    update_fields_on_open: bool | None = document.update_fields_on_open
    document.update_fields_on_open = True
    document.update_fields_on_open = None
    content_index: int = document.find_content_index(first)
    inserted: Paragraph = document.insert_paragraph(content_index, "inserted")
    fragment: ContentFragment = document.pop_content(content_index)
    fragment_kind: str = fragment.kind
    document.insert_content(content_index, fragment)
    document.clone_content(inserted, content_index)
    document.move_content(table, content_index)
    replacement_count: int = document.try_replace_text("old", "new")
    regex_count: int = document.replace_all_regex([("old", "new")])
    expected_count: int = document.try_replace_text("old", "new", expect=1)
    batch_counts: tuple[int, ...] = document.replace_all([("a", "b", 1), ("c", "d")])
    count_error = ReplacementCountError("message", 1, 2, index=0)
    assert_type(count_error.index, int | None)
    assert_type(count_error.found, int)
    revisions: tuple[Revision, ...] = document.revisions
    accepted: int = document.accept_all()
    dated: int = document.reject_revisions_in_date_range(
        start="2026-01-01T00:00:00Z", end="2026-12-31T00:00:00Z"
    )
    replaced: int = document.try_replace_text("{{name}}", "Ada")
    matched: int = document.replace_all_regex([(r"\d", "#")])
    updated: int = document.update_fields(
        file_name="report.docx", merge_fields={"Name": "Ada"}
    )
    assert_type(revisions[0].timestamp, str | None)
    assert_type(revisions[0].story, Story | None)
    document.set_header("Header")
    document.set_footer("Footer")
    document.set_story_text(story_items[0], "edited")
    document.add_hyperlink_to_story(stories[0], "home", "https://example.com/")
    section_footer: Story = document.create_section_story(0, "footer", "default")
    linked_footer: Story = document.link_section_story(0, "footer", "first", section_footer)
    unlinked_footer: Story = document.unlink_section_story(0, "footer", "first")
    document.insert_content(section_footer, document.pop_content(story_items[0]))
    document.insert_content(story_items[0], fragment)
    link_run: Run = first.add_hyperlink("docs", "https://example.com/docs")
    assert_type(story_items[0].xml, bytes)
    compatible_story_item = StoryItem(
        story=stories[0],
        kind="paragraph",
        index_path=(0,),
        text=None,
        xml=None,
    )
    assert_type(compatible_story_item.xml, bytes)
    text_content_index: int = document.find_content_index("typed")
    text_content_indices: tuple[int, ...] = document.find_content_indices("typed")
    page_fields_updated: int = document.update_page_fields()
    layout_field_report: LayoutBackedFieldUpdateReport = (
        document.update_layout_backed_fields()
    )
    if fragments:
        bounds: BoundingBox = fragments[0].bounds
        assert_type(bounds.width, float)
    assert_type(comments[0].date, str | None)
    assert_type(comments[0].anchor_text, str | None)
    assert_type(comments[0].anchor, StoryRunRange | None)
    assert_type(sections[0].page_width, int | None)
    assert_type(styles[0].style_type, str)
    assert_type(stories[0].owner_index, int)
    assert_type(story_items[0].index_path, tuple[int, ...])
    assert_type(story_items[0].revision, int)
    assert_type(variants[0].story, Story | None)
    assert_type(hyperlinks[0].url, str | None)
    assert_type(report.entry_count, int)
    assert_type(report.diagnostics, tuple[str, ...])
    assert_type(report.diagnostic_count, int)
    (
        package_bytes,
        pdf_bytes,
        pages,
        maybe_page,
        sliced,
        channels,
        cloned_row,
        update_fields_on_open,
        replacement_count,
        regex_count,
        image_data,
        fragment_kind,
        inserted_picture,
        story_comment_id,
        updated_section,
        style_removed,
    )
    accepted, dated, replaced, matched, updated


if TYPE_CHECKING:
    Cell()  # type: ignore[call-arg]
    CellCollection()  # type: ignore[call-arg]
    CellParagraphCollection()  # type: ignore[call-arg]
    CoreProperties()  # type: ignore[call-arg]
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
    Document().compare(Document(), "Ada", "2026-09-14T09:00:00Z", granularity="words")  # type: ignore[arg-type]
    Document().compare(Document(), "Ada", "2026-09-14T09:00:00Z", ignored_stories="header")  # type: ignore[arg-type]


def scoped_replacement_signatures_cover_paragraph_cell_and_story_item(document: Document, paragraph: Paragraph, cell: Cell, item: StoryItem) -> None:
    assert_type(paragraph.replace_text("old", "new", expect=1), int)
    assert_type(cell.replace_text(old="old", new="new", expect=1), int)
    assert_type(document.replace_text_at(item, "old", "new", expect=1), int)

def whole_story_setter_signatures(document: Document) -> None:
    assert_type(document.set_header(text="header"), None)
    assert_type(document.set_footer(text="footer"), None)


def issue_303_formatting_signatures(document: Document, paragraph: Paragraph) -> None:
    from rdocx import Pt, TabStop, WD_TAB_ALIGNMENT, WD_TAB_LEADER

    font = paragraph.runs[0].font
    font.superscript = True
    font.small_caps = None
    font.character_spacing = Pt(1)
    font.east_asian_name = "MS Mincho"
    assert_type(font.subscript, bool | None)
    assert_type(paragraph.runs[0].font.character_spacing, Length | None)
    tab_stops = paragraph.paragraph_format.tab_stops
    assert_type(tab_stops.add_tab_stop(Inches(1), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS), TabStop)
    assert_type(tab_stops[0].alignment, WD_TAB_ALIGNMENT | None)
    paragraph.paragraph_format.set_border("bottom", size=6, color="4472C4")
    assert_type(paragraph.paragraph_format.border("bottom"), tuple[str, int | None, str | None] | None)
    paragraph.paragraph_format.outline_level = 1
    document.add_style("Callout", tab_stops=[(Inches(1), WD_TAB_ALIGNMENT.LEFT)], borders={"left": ("single", 4, "auto")})
    assert_type(document.add_numbered_list_item("item", restart=True), Paragraph)
    assert_type(document.restart_numbering(paragraph, start=3), int)
    assert_type(document.default_font_size, Length | None)
    assert_type(ListLevel.checklist(checked=True), ListLevel)
