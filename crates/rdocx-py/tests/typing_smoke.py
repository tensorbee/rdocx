from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Literal, assert_type

from rdocx.enum.dml import MSO_COLOR_TYPE

from rdocx import (
    Bookmark,
    BoundingBox,
    Cell,
    CellCollection,
    CellParagraphCollection,
    Comment,
    ComparisonDiagnostic,
    ContentFragment,
    ContentLocation,
    ContentControlTypePayload,
    CoreProperties,
    Document,
    HtmlDiagnostic,
    Font,
    HeaderFooterVariant,
    Hyperlink,
    Inches,
    Length,
    Pt,
    Emu,
    Twips,
    LayoutFragment,
    FieldEvaluation,
    FieldsUpdateReport,
    LayoutBackedFieldUpdateReport,
    LayoutPage,
    ListLevel,
    RGBColor,
    ReplacementCountError,
    Paragraph,
    Picture,
    PictureOptions,
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
    assert_type(table.indent, Length | None)
    table.indent = Inches(-0.25)
    table.indent = None
    row: Row = table.rows[0]
    cloned_row: Row = table.clone_row(0)
    table.remove_row(0)
    cell: Cell = row.cells[0]
    cell.text = first.text
    assert_type(table.border("top"), tuple[str, int | None, RGBColor | Literal["auto"] | None] | None)
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
    assert_type(cell.shading, RGBColor | Literal["auto"] | None)
    assert_type(cell.border("bottom"), tuple[str, int | None, RGBColor | Literal["auto"] | None] | None)
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
    legal_pdf: bytes = opened.to_pdf(revision_view="tracked", revision_style="legal")
    legal_pages: list[bytes] | bytes = opened.render_pages(
        revision_view="tracked", revision_style="legal"
    )
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
    document.update_section(0, margin_top=Twips(1440), margin_left=Emu(635), gutter=0)
    document.section(0).margin_top = 1440
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.space_after = 635
    run.font.character_spacing = Emu(-635)
    paragraph.paragraph_format.line_spacing = 2
    paragraph.paragraph_format.line_spacing = 1.75
    paragraph.paragraph_format.line_spacing = Pt(18)
    assert_type(document.paragraphs[0].paragraph_format.line_spacing, Length | float | None)
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
    resized: int = document.set_picture_sizes_for_relationship("rId1", width=Inches(1), height=Inches(1))
    story_items: tuple[StoryItem, ...] = document.story_items
    inserted_picture: Picture = document.add_picture(
        b"png", "image.png", Inches(1), Inches(1), after=story_items[0]
    )
    assert_type(document.pictures(), tuple[Picture, ...])
    assert_type(document.set_picture_size(inserted_picture, Inches(1), Inches(1)), None)
    assert_type(inserted_picture.bytes, bytes)
    assert_type(inserted_picture.title, str | None)
    assert_type(inserted_picture.decorative, bool)
    options = PictureOptions(title="owned", decorative=True, crop=(0.1, 0.0, 0.1, 0.0))
    assert_type(document.paragraphs[0].runs[0].add_picture(path, width=Inches(1), height=Inches(1), options=options), Picture)
    assert_type(document.tables[0].cell(0, 0).add_picture(b"png", width=Inches(1), height=Inches(1)), Picture)
    assert_type(document.paragraphs[0].add_internal_hyperlink("jump", "Target", tooltip="tip"), Run)
    assert_type(document.add_heading_hyperlink(document.paragraphs[0], document.paragraphs[1], "jump"), Run)
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
    assert_type(comments[0].anchor_issue, str | None)
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


def colour_formats() -> None:
    from rdocx import Document, MSO_THEME_COLOR, RGBColor
    from rdocx._rdocx import ColorFormat
    document = Document()
    font = document.add_paragraph("").add_run("colour").font
    assert_type(font.color, ColorFormat)
    assert_type(font.color.rgb, RGBColor | None)
    assert_type(font.color.theme_color, str | None)
    font.color = "#123456"
    font.color = (1, 2, 3)
    font.color.rgb = RGBColor(1, 2, 3)
    font.color.rgb = None
    font.color.theme_color = MSO_THEME_COLOR.ACCENT_1


def fx195_caller_font_render_signatures(document: Document, directory: Path) -> None:
    from rdocx import SvgRenderResult
    fonts: list[tuple[str, bytes]] = [("caller", b"font")]
    dirs: list[str | Path] = [directory, str(directory)]
    assert_type(document.to_pdf(fonts=fonts, font_dirs=dirs), bytes)
    assert_type(document.render_page_to_png(0, fonts=fonts, font_dirs=dirs), bytes | None)
    assert_type(document.render_page_to_svg(0, fonts=fonts, font_dirs=dirs), SvgRenderResult | None)
    assert_type(document.render_all_pages(fonts=fonts, font_dirs=dirs), list[bytes])
    assert_type(document.render_pages(pages=None, fonts=fonts, font_dirs=dirs), list[bytes] | bytes)
    assert_type(document.layout(fonts=fonts, font_dirs=dirs), tuple[LayoutFragment, ...])
    assert_type(document.layout_page(0, fonts=fonts, font_dirs=dirs), LayoutPage | None)
    assert_type(document.to_pdf(font_dir=directory), bytes)


def fx200_checked_list_factory_signatures(document: Document) -> None:
    assert_type(document.add_bullet_list_item("item"), Paragraph)
    assert_type(document.add_bullet_list_item("item", level=8), Paragraph)
    assert_type(document.add_numbered_list_item("item"), Paragraph)
    assert_type(document.add_numbered_list_item("item", level=8), Paragraph)


def fx203_owned_document_operations(document: Document, paragraph: Paragraph, path: Path) -> None:
    from rdocx import (
        ContentControl, CustomFileTime, CustomProperty, DocumentFragment,
        ContentLocation, Equation, EquationConversionResult, EquationInsertResult,
        FragmentConflictPolicy, OdtExportResult, RtfExportResult, EpubExportResult,
        TextReport, ValidationReport, validate_bytes, validate_file,
    )
    assert_type(document.render_template({"name": "Ada"}, grammar_version=1), int)
    assert_type(document.render_template({"name": "Ada"}, grammar_version=None), int)
    assert_type(document.text_report(), TextReport)
    assert_type(document.validate(), ValidationReport)
    assert_type(validate_bytes(document.to_bytes()), ValidationReport)
    assert_type(validate_file(path), ValidationReport)
    assert_type(document.to_odt(), OdtExportResult)
    assert_type(document.to_rtf(), RtfExportResult)
    assert_type(document.to_epub(), EpubExportResult)
    assert_type(document.word_count(), int)
    assert_type(document.character_count(include_spaces=False), int)
    assert_type(document.render_template({"name": "Ada", "flag": True, "items": [1, None]}), int)
    document.settings.track_revisions = None
    document.app_properties.company = "Example"
    document.custom_properties.set("date", CustomFileTime("2026-10-12T12:00:00Z"))
    assert_type(document.custom_properties.get("date"), CustomProperty | None)
    assert_type(document.custom_properties.delete("absent"), bool)
    document.custom_properties["Client"] = "Ada"
    assert_type("Client" in document.custom_properties, bool)
    del document.custom_properties["Client"]
    assert_type(document.paragraphs[0].runs[0].font.color.type, MSO_COLOR_TYPE | None)
    color_kind: MSO_COLOR_TYPE = MSO_COLOR_TYPE.RGB
    assert_type(color_kind, MSO_COLOR_TYPE)

    policy = FragmentConflictPolicy.reuse_equivalent().with_style_reuse(False)
    story = document.stories[0]
    start = ContentLocation(story, "paragraph", [0])
    end = ContentLocation.end(story)
    fragment = DocumentFragment.from_range(document, start, end)
    assert_type(document.import_fragment(end, fragment, policy=policy), None)
    assert_type(document.insert_document(Document(), at=0, policy=policy), None)
    assert_type(document.content_controls(), tuple[ContentControl, ...])
    assert_type(document.set_content_control_value_by_tag("check", "true"), int)
    assert_type(paragraph.equations(), tuple[Equation, ...])
    inserted = paragraph.add_equation("x+1", format="latex")
    assert_type(inserted, EquationInsertResult)
    assert_type(inserted.equation.to_latex(), EquationConversionResult)


def fx204_shared_operation_types(document: Document) -> None:
    assert_type(document.replace_all_regex_expected([("x", "y", 1)]), tuple[int, ...])
    assert_type(document.fill_content_control_tags([("tag", "value")]), tuple[int, ...])
    fields_report = document.update_all_fields()
    assert_type(fields_report, FieldsUpdateReport)
    assert_type(fields_report.ordinary_updated, int)
    assert_type(fields_report.layout, LayoutBackedFieldUpdateReport)
    assert_type(fields_report.evaluations, tuple[FieldEvaluation, ...])
    assert_type(fields_report.total_updated, int)
    assert_type(fields_report.diagnostics, tuple[str, ...])
    assert_type(fields_report.saved_cache_scope, str)


def html_markdown_imports(source: Path, template: Document, paragraph: Paragraph, image: bytes) -> None:
    created, notes = Document.from_html("<p>Body</p>", images=(("face.png", image, "face.png"),))
    assert_type(created, Document)
    assert_type(notes, tuple[HtmlDiagnostic, ...])
    assert_type(Document.open_html(source), tuple[Document, tuple[HtmlDiagnostic, ...]])
    assert_type(Document.from_markdown("# Title", template=template), tuple[Document, tuple[HtmlDiagnostic, ...]])
    assert_type(Document.open_markdown(source, template=source), tuple[Document, tuple[HtmlDiagnostic, ...]])
    assert_type(template.insert_html("<p>New</p>", after=paragraph), tuple[HtmlDiagnostic, ...])


def check_fx220_font_embedding_types() -> None:
    from rdocx import Document, FontEmbeddingLicense, FontDefinition, EmbeddedFont
    from typing import assert_type
    value = Document()
    license = FontEmbeddingLicense(authorized=True, identity="OFL 1.1")
    assert_type(license.authorized, bool)
    assert_type(license.identity, str)
    value.embed_font("typed face", b"font", license=license)
    assert_type(value.fonts, tuple[FontDefinition, ...])
    assert_type(value.remove_embedded_font("typed face"), EmbeddedFont | None)
    assert_type(value.remove_font("typed face"), FontDefinition | None)
    value.set_font(FontDefinition("typed face", family="swiss"))
    value.embed_system_fonts = None
    assert_type(value.embed_true_type_fonts, bool | None)
    value.embed_true_type_fonts = True
    value.save_subset_fonts = False


def fx220_licensed_embedding_recipe(font_path: Path, output_path: Path) -> None:
    from rdocx import FontEmbeddingLicense, FontDefinition
    value = Document()
    license = FontEmbeddingLicense(authorized=True, identity="SIL Open Font License 1.1")
    value.embed_font("My Face", font_path, "regular", license=license)
    value.save(output_path)
    assert_type(value.fonts, tuple[FontDefinition, ...])
    value.embed_true_type_fonts = False
    value.remove_embedded_font("My Face", "regular")


def fx215_chart_typing() -> None:
    from rdocx import CategoryChartData, XyChartData, XL_CHART_TYPE
    data = CategoryChartData()
    data.categories = ["East", "West"]
    series = data.add_series("Sales", [11.0, 23.0])
    assert_type(series.values, tuple[float, ...])
    xy = XyChartData()
    xy.add_series("A").add_data_point(1.0, 2.0)
    document = Document()
    paragraph = document.add_chart(XL_CHART_TYPE.LINE, data, Inches(4), Inches(2))
    assert_type(paragraph, Paragraph)
    assert_type(document.chart_legend(paragraph), dict[str, bool | str])
    document.set_chart_legend(paragraph, visible=True, position="b", overlay=False)


def fx211a_checked_word_element_types(document: Document) -> None:
    paragraph = document.paragraphs[0]
    assert_type(paragraph.element, Paragraph)
    assert_type(paragraph.element.xml, str)
    assert_type(paragraph.element.replace_xml(paragraph.element.xml), None)
    run = paragraph.runs[0]
    assert_type(run.element, Run)
    assert_type(run.element.xml, str)
    assert_type(run.element.replace_xml(run.element.xml), None)
    table = document.tables[0]
    assert_type(table.element, Table)
    assert_type(table.element.xml, str)
    assert_type(table.element.replace_xml(table.element.xml), None)
    cell = table.cell(0, 0)
    assert_type(cell.element, Cell)
    assert_type(cell.element.xml, str)
    assert_type(cell.element.replace_xml(cell.element.xml), None)
    section = document.section(0)
    assert_type(section.element.xml, str)
    assert_type(section.element.replace_xml(section.element.xml), None)


def fx241_form_types(document: Document) -> None:
    from rdocx import LegacyFormField
    assert_type(document.bind_content_controls({"client": "Ada"}), int)
    assert_type(document.legacy_form_fields(), tuple[LegacyFormField, ...])
    field = document.set_legacy_form_field_value(None, "value", source_part="/word/document.xml", ordinal=0)
    assert_type(field, LegacyFormField)
    assert_type(field.source_part, str)
    assert_type(field.ordinal, int)
    assert_type(field.name, str | None)
    assert_type(field.enabled, bool)
    assert_type(field.calculate_on_exit, bool)
    assert_type(field.type, Literal["text", "checkbox", "dropdown"])
    assert_type(field.value, str | bool | int)
    assert_type(field.choices, tuple[str, ...])


def fx240_owned_signature_reports(owner: Document, private_key: bytes, certificate: bytes) -> None:
    report = owner.sign(private_key, certificate)
    part: str = report["signature_part"]
    valid: bool = report["cryptographically_valid"]
    complete: bool = report["coverage_complete"]
    parts: list[str] = report["covered_parts"]
    if report["signer"] is not None:
        subject: str = report["signer"]["subject"]
        issuer: str = report["signer"]["issuer"]
        serial: str = report["signer"]["serial_number"]
        print(subject, issuer, serial)
    for relationship in report["covered_relationships"]:
        print(relationship["source_part"], relationship["relationship_id"])
    for issue in report["issues"]:
        details: dict[str, str] = issue["details"]
        print(issue["kind"], details)
    for checked in owner.verify_signatures():
        print(checked["signature_part"], checked["cryptographically_valid"])
    print(part, valid, complete, parts)


if TYPE_CHECKING:
    import io
    from rdocx.enum.text import WD_COLOR_INDEX, WD_BREAK
    from rdocx.enum.section import WD_SECTION, WD_SECTION_START, WD_ORIENT, WD_ORIENTATION
    _section_start: WD_SECTION_START = WD_SECTION.NEW_PAGE
    _orientation: WD_ORIENTATION = WD_ORIENT.LANDSCAPE
    assert_type(_section_start, WD_SECTION_START)
    assert_type(_orientation, WD_ORIENTATION)
    _parity_document = Document()
    _parity_paragraph = _parity_document.add_paragraph(style="Normal")
    assert_type(_parity_document.add_heading("Heading", 1), Paragraph)
    _parity_run = _parity_paragraph.add_run("run")
    _parity_run.bold = True
    _parity_run.italic = None
    _parity_run.underline = False
    _parity_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    _parity_run.add_break(WD_BREAK.PAGE)
    _parity_document.save(io.BytesIO())
    _parity_document = Document.open(io.BytesIO())
    assert_type(_parity_document.styles, tuple[Style, ...])


def s953_typed_authoring_contract(owner: Document, location: ContentLocation) -> None:
    payload = ContentControlTypePayload.date("yyyy-MM-dd")
    assert_type(payload, ContentControlTypePayload)
    owner.set_content_control_type(location, payload)
    assert_type(owner.add_custom_xml_store(b"<root/>"), str)
    assert_type(owner.add_custom_xml_store(b"<root/>", ("urn:example",)), str)
    assert_type(owner.remove_content_range(0, 0), int)
