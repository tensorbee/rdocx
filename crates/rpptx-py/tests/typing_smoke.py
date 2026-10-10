import io
from pathlib import Path
from typing import TYPE_CHECKING, Callable

from rpptx import (
    BoundingBox,
    Comment,
    CommentAuthor,
    CommentReply,
    Inches,
    Length,
    MSO_ANCHOR,
    MSO_AUTO_SIZE,
    MSO_SHAPE,
    MSO_UNDERLINE,
    PP_ALIGN,
    Presentation,
    Pt,
    ReplacementCountError,
    RGBColor,
    TextFrameLayout,
    TextLineLayout,
    ValidationIssue,
)
from rpptx._rpptx import (
    AdjustmentCollection,
    Background,
    Cell,
    ColorFormat,
    Column,
    ColumnCollection,
    CoreProperties,
    FillFormat,
    GradientStop,
    GradientStops,
    MediaInfo,
    Font,
    Hyperlink,
    Image,
    LineEndFormat,
    LineFormat,
    Paragraph,
    ParagraphCollection,
    PlaceholderCollection,
    Row,
    RowCollection,
    Run,
    RunCollection,
    Section,
    ShadowFormat,
    Shape,
    ShapeClickAction,
    ShapeCollection,
    ShapeHyperlink,
    Slide,
    SlideCollection,
    SlideLayout,
    SlideLayoutCollection,
    Table,
    TextFrame,
)
from rpptx.enum.dml import (
    MSO_ARROWHEAD_LENGTH,
    MSO_ARROWHEAD_STYLE,
    MSO_ARROWHEAD_WIDTH,
    MSO_FILL_TYPE,
    MSO_LINE_DASH_STYLE,
    MSO_PATTERN_TYPE,
)
from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE_TYPE


def exercise_rpptx_types(path: Path) -> None:
    presentation = Presentation(path)
    layout: SlideLayout = presentation.slide_layouts[0]
    slide: Slide = presentation.slides.add_slide(layout)
    shape: Shape = slide.shapes.add_textbox(
        Inches(1), Inches(1), Inches(4), Inches(2)
    )
    shape.text = "typed"
    shape_left: Length | None = shape.left
    shape_top: Length | None = shape.top
    shape_width: Length | None = shape.width
    shape_height: Length | None = shape.height
    shape_id: int | None = shape.shape_id
    shape_name: str | None = shape.name
    autofit: str | None = shape.text_frame.autofit
    paragraph: Paragraph = presentation.slides[0].shapes[-1].text_frame.paragraphs[0]
    paragraph.level = 1
    paragraph.font.bold = True
    paragraph.font.size = Pt(12)
    returned_size: Length | None = (
        presentation.slides[0].shapes[-1].text_frame.paragraphs[0].font.size
    )
    run: Run = paragraph.runs[0]
    run.text = "updated"
    run_font_name: str | None = run.font.name
    run_font_size: Length | None = run.font.size
    run_font_color: str | None = run.font.color
    frame: TextFrame = presentation.slides[0].shapes[-1].text_frame
    frame.margin_left = Inches(0.1)
    frame.margin_right = None
    frame.margin_top = Pt(2)
    frame.margin_bottom = 0
    margins: tuple[Length | None, ...] = (
        frame.margin_left,
        frame.margin_right,
        frame.margin_top,
        frame.margin_bottom,
    )
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    anchor: MSO_ANCHOR | None = frame.vertical_anchor
    frame.word_wrap = True
    word_wrap: bool | None = frame.word_wrap
    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    auto_size: MSO_AUTO_SIZE | None = frame.auto_size
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    alignment: PP_ALIGN | None = paragraph.alignment
    paragraph.line_spacing = 1.5
    paragraph.line_spacing = Pt(18)
    line_spacing: float | Length | None = paragraph.line_spacing
    paragraph.space_before = Pt(6)
    paragraph.space_after = 0.5
    spacing: tuple[Length | float | None, ...] = (
        paragraph.space_before,
        paragraph.space_after,
    )
    paragraph.left_indent = Inches(0.5)
    paragraph.right_indent = None
    paragraph.first_line_indent = -Inches(0.25)
    indents: tuple[Length | None, ...] = (
        paragraph.left_indent,
        paragraph.right_indent,
        paragraph.first_line_indent,
    )
    paragraph.bullet = "-"
    paragraph.bullet = False
    bullet: str | bool | None = paragraph.bullet
    added: Run = paragraph.add_run("added")
    added = frame.paragraphs[0].add_run()
    font: Font = added.font
    font.name = "Arial"
    font.color = RGBColor(0x12, 0x34, 0x56)
    font.color = "123456"
    font.color = None
    font.italic = True
    font.underline = True
    font.underline = MSO_UNDERLINE.DOUBLE_LINE
    underline: bool | MSO_UNDERLINE | None = font.underline
    font.strike = False
    font.all_caps = None
    font_flags: tuple[bool | None, ...] = (font.italic, font.strike, font.all_caps)
    broad_shape_factory: Callable[[int, int, int, int, int], Shape] = (
        presentation.slides[0].shapes.add_shape  # type: ignore[assignment]
    )
    length: Length = Pt(12)
    points: float = length.pt
    emu: int = length.emu
    shape = presentation.slides[0].shapes.add_shape(
        MSO_SHAPE.CHEVRON, Inches(1), Inches(1), Inches(2), Inches(1)
    )
    shape = presentation.slides[0].shapes.add_table(
        1, 1, Inches(1), Inches(1), Inches(4), Inches(1)
    )
    table: Table = shape.table
    column: Column = table.columns[0]
    column.width = Inches(2)
    cell: Cell = table.cell(0, 0)
    cell.text = paragraph.text
    shapes: list[Shape] = presentation.slides[0].shapes[:]
    for current_slide in presentation.slides:
        for current_shape in current_slide.shapes:
            current_shape.has_text_frame
    package_bytes: bytes = presentation.to_bytes()
    reopened: Presentation = Presentation.from_bytes(package_bytes)
    replaced: int = presentation.try_replace_text("typed", "checked")
    replacement_counts: tuple[int, int] = (0, 0)
    try:
        replaced = presentation.try_replace_text("checked", "typed", expect=2)
    except ReplacementCountError as error:
        replacement_counts = (error.expected, error.found)
    slide_replaced: int = presentation.slides[0].try_replace_text(
        "typed", "checked", expect=None, notes=False
    )
    frame_replaced: int = (
        presentation.slides[0].shapes[0].text_frame.try_replace_text("typed", "checked", expect=0)
    )
    issues: tuple[ValidationIssue, ...] = presentation.validate()
    issue_lines: list[tuple[str, str]] = [(issue.kind, issue.message) for issue in issues]
    pdf_bytes: bytes = presentation.to_pdf()
    slide_png: bytes | None = presentation.render_slide_to_png(0)
    slide_pngs: list[bytes] = presentation.render_all_slides()
    notes_pdf: bytes = presentation.to_notes_pdf()
    notes_pngs: list[bytes] = presentation.render_all_notes()
    notes_text: str | None = presentation.slides[0].notes_text
    presentation.slides[0].notes_text = "Updated speaker note"
    presentation.add_comment_author(
        id="{11111111-1111-1111-1111-111111111111}",
        name="Ada",
        user_id="ada@example.com",
        provider_id="local",
    )
    authors: tuple[CommentAuthor, ...] = presentation.comment_authors
    current_slide = presentation.slides[0]
    current_slide.add_comment(
        id="{22222222-2222-2222-2222-222222222222}",
        author_id=authors[0].id,
        created="2026-09-14T10:30:00Z",
        text="Review",
    )
    comments: tuple[Comment, ...] = presentation.slides[0].comments
    reply: CommentReply = comments[0].replies[0]
    presentation.slides[0].resolve_comment(comments[0].id)
    presentation.slides[0].remove_comment(reply.id)
    slide_width: Length | None = presentation.slide_width
    presentation.slide_height = Inches(6)
    current_slide = presentation.slides[0]
    slide_layout: SlideLayout = current_slide.slide_layout
    layout_index: int = presentation.slide_layouts.index(slide_layout)
    same_layout: bool = slide_layout == presentation.slide_layouts[0]
    current_slide.slide_layout = presentation.slide_layouts[1]
    current_slide = presentation.slides[0]
    hidden: bool = current_slide.hidden
    current_slide.hidden = True
    background: Background = current_slide.background
    background_fill: FillFormat = background.fill
    follows_master: bool = current_slide.follow_master_background
    current_slide.follow_master_background = False
    shape = current_slide.shapes.add_shape(
        "roundRect", Inches(1), Inches(1), Inches(2), Inches(1)
    )
    shape.left = Inches(2)
    shape.top = Inches(2)
    shape.width = Inches(3)
    shape.height = Inches(1)
    geometry: tuple[Length, Length, Length, Length] | None = shape.effective_geometry()
    shape.name = "Typed"
    shape.rotation = 15.0
    rotation: float = shape.rotation
    shape_type: MSO_SHAPE_TYPE | None = shape.shape_type
    shape.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
    shape.auto_shape_type = "chevron"
    auto_shape_type: MSO_SHAPE | None = shape.auto_shape_type
    member: MSO_SHAPE = MSO_SHAPE.from_xml("roundRect")
    adjustments: AdjustmentCollection = shape.adjustments
    adjustments[0] = 0.25
    first_adjustment: float = adjustments[0]
    all_adjustments: list[float] = list(adjustments)
    fill: FillFormat = shape.fill
    fill.solid()
    fill_type: MSO_FILL_TYPE | None = fill.type
    fore_color: ColorFormat = fill.fore_color
    fore_color.rgb = RGBColor(0x12, 0x34, 0x56)
    rgb: RGBColor | None = fore_color.rgb
    line: LineFormat = shape.line
    line.width = Pt(1)
    line.width = None
    line_width: Length = line.width
    line.color.rgb = RGBColor.from_string("FF0000")
    line_fill: FillFormat = line.fill
    line.dash_style = MSO_LINE_DASH_STYLE.DASH
    line.dash_style = None
    dash_style: MSO_LINE_DASH_STYLE | None = line.dash_style
    tail_end: LineEndFormat = line.tail_end
    tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE
    tail_end.width = MSO_ARROWHEAD_WIDTH.WIDE
    tail_end.length = None
    end_type: MSO_ARROWHEAD_STYLE | None = line.head_end.type
    end_width: MSO_ARROWHEAD_WIDTH | None = tail_end.width
    end_length: MSO_ARROWHEAD_LENGTH | None = tail_end.length
    shadow: ShadowFormat = shape.shadow
    shadow.inherit = False
    inherits: bool = shadow.inherit
    shadow.visible = True
    visible: bool = shadow.visible
    shadow.color.rgb = RGBColor(0, 0, 0)
    shadow.alpha = 0.4
    alpha: float | None = shadow.alpha
    shadow.blur_radius = Pt(4)
    blur_radius: Length | None = shadow.blur_radius
    shadow.distance = Pt(3)
    distance: Length | None = shadow.distance
    shadow.direction = 45.0
    direction: float | None = shadow.direction
    shadow.align = "tl"
    align: str | None = shadow.align
    shadow.rotate_with_shape = False
    rotate_with_shape: bool | None = shadow.rotate_with_shape
    shape_xml: bytes = shape.xml
    connector: Shape = presentation.slides[0].shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, 0, 0, Inches(1), Inches(1)
    )
    connector.theme_effect_index = 0
    theme_effect_index: int | None = connector.theme_effect_index
    group: Shape = presentation.slides[0].shapes.add_group_shape()
    picture = presentation.slides[0].shapes.add_picture(
        io.BytesIO(b""), 0, 0
    )
    image: Image = picture.image
    blob: bytes = image.blob
    picture.replace_image(b"")
    picture.replace_image(path)
    presentation.slides[0].shapes.remove(group)
    duplicated: Slide = presentation.slides.duplicate(presentation.slides[0])
    presentation.slides.move(0, -1)
    imported: Slide = presentation.slides.import_slide(
        presentation.slides[0], layout=presentation.slide_layouts[0], index=0
    )
    presentation.slides.remove(presentation.slides[0])
    presentation.save(path)
    (
        package_bytes,
        pdf_bytes,
        slide_png,
        slide_pngs,
        notes_pdf,
        notes_pngs,
        notes_text,
        authors,
        comments,
        reply,
        shapes,
        returned_size,
        points,
        emu,
        broad_shape_factory,
        shape_left,
        shape_top,
        shape_width,
        shape_height,
        shape_id,
        shape_name,
        autofit,
        run_font_name,
        run_font_size,
        run_font_color,
        margins,
        anchor,
        word_wrap,
        auto_size,
        alignment,
        line_spacing,
        spacing,
        indents,
        bullet,
        underline,
        font_flags,
        slide_width,
        layout_index,
        same_layout,
        hidden,
        background_fill,
        follows_master,
        rotation,
        shape_type,
        auto_shape_type,
        member,
        first_adjustment,
        all_adjustments,
        fill_type,
        rgb,
        line_width,
        line_fill,
        shape_xml,
        connector,
        theme_effect_index,
        blob,
        image.content_type,
        image.ext,
        duplicated,
        reopened,
        replaced,
        replacement_counts,
        slide_replaced,
        frame_replaced,
        issue_lines,
    )


def exercise_rpptx_table_types(table: Table) -> None:
    origin: Cell = table.cell(0, 0)
    origin.merge(table.cell(1, 1))
    spans: tuple[bool, bool, int, int] = (
        origin.is_merge_origin,
        origin.is_spanned,
        origin.span_height,
        origin.span_width,
    )
    origin.split()
    cell_fill: FillFormat = origin.fill
    cell_fill.solid()
    origin.margin_left = Inches(0.1)
    origin.margin_right = None
    cell_margins: tuple[Length | None, ...] = (
        origin.margin_left,
        origin.margin_right,
        origin.margin_top,
        origin.margin_bottom,
    )
    rows: RowCollection = table.rows
    row: Row = rows[0]
    row.height = Inches(1)
    row_heights: list[Length] = [current.height for current in rows]
    borders: tuple[LineFormat, ...] = (
        origin.border_left,
        origin.border_right,
        origin.border_top,
        origin.border_bottom,
    )
    borders[0].width = Pt(1)
    (spans, cell_margins, row_heights, rows[:])


def exercise_rpptx_table_structure_types(table: Table) -> None:
    appended: Row = table.rows.add_row()
    inserted: Row = table.rows.add_row(0)
    table.rows.remove(table.rows[-1])
    column: Column = table.columns.add_column()
    table.columns.add_column(index=-1)
    table.columns.remove(table.columns[0])
    (appended, inserted, column)


def exercise_rpptx_picture_crop_types(picture: Shape) -> None:
    picture.crop_left = 0.25
    picture.crop_top = 0
    crop: tuple[float, float, float, float] = (
        picture.crop_left,
        picture.crop_top,
        picture.crop_right,
        picture.crop_bottom,
    )
    picture.crop_right = crop[0]
    picture.crop_bottom = -0.1


def exercise_rpptx_z_order_types(slide: Slide) -> None:
    slide.shapes.move(0, -1)


def exercise_rpptx_group_population_types(group: Shape) -> None:
    members: ShapeCollection = group.shapes
    textbox: Shape = members.add_textbox(0, 0, Inches(1), Inches(1))
    nested: Shape = group.shapes.add_group_shape()
    nested.shapes.add_shape("rect", 0, 0, Inches(1), Inches(1))
    group.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, Inches(1), Inches(1))
    group.shapes.add_table(1, 2, 0, 0, Inches(2), Inches(1))
    group.shapes.add_picture(io.BytesIO(b""), 0, 0)


def exercise_rpptx_hyperlink_types(run: Run) -> None:
    hyperlink: Hyperlink = run.hyperlink
    hyperlink.address = "https://example.com"
    address: str | None = hyperlink.address
    hyperlink.address = None
    (address,)


def exercise_rpptx_click_action_types(shape: Shape, slide: Slide) -> None:
    action: ShapeClickAction = shape.click_action
    link: ShapeHyperlink = action.hyperlink
    link.address = "https://example.com"
    link.address = None
    action.target_slide = slide
    target: Slide | None = action.target_slide
    action.target_slide = None
    same: bool = target == slide
    (same,)


def exercise_rpptx_text_layout_types(presentation: Presentation) -> None:
    frames: tuple[TextFrameLayout, ...] = presentation.text_layout()
    narrower: tuple[TextFrameLayout, ...] = presentation.text_layout(width_factor=0.95)
    for frame in frames + narrower:
        slide_index: int = frame.slide_index
        shape_id: int | None = frame.shape_id
        name: str | None = frame.name
        autofit: str = frame.autofit
        box: BoundingBox = frame.frame
        usable: BoundingBox = frame.usable
        box_edges: tuple[float, float, float, float] = (
            usable.x,
            usable.y,
            usable.width,
            usable.height,
        )
        font_scale: float = frame.font_scale
        height: float = frame.height
        overflow: bool = frame.overflow
        lines: tuple[TextLineLayout, ...] = frame.lines
        for line in lines:
            paragraph_index: int = line.paragraph_index
            text: str = line.text
            bounds: BoundingBox = line.bounds
            baseline: float = line.baseline
            font_size: float = line.font_size
            (paragraph_index, text, bounds, baseline, font_size)
        (
            slide_index,
            shape_id,
            name,
            autofit,
            box,
            box_edges,
            font_scale,
            height,
            overflow,
        )


def exercise_rpptx_issue_308_types(presentation: Presentation, path: Path) -> None:
    properties: CoreProperties = presentation.core_properties
    properties.title = "Deck"
    properties.author = None
    title: str = properties.title
    revision: int = properties.revision
    sections: tuple[Section, ...] = presentation.sections
    presentation.set_sections([("Intro", [0])])
    section_indices: list[int] = sections[0].slide_indices
    jpegs: list[bytes] = presentation.render_slides(format="jpeg", quality=80, slides=[0])
    tiff: bytes = presentation.render_slides(format="tiff")
    pdfa: bytes = presentation.to_pdfa("pdfa-3b")
    handout: bytes = presentation.to_handout_pdf(3)
    handouts: list[bytes] = presentation.render_all_handouts(6, dpi=72.0)
    odp, odp_diagnostics = presentation.to_odp()
    converted, read_diagnostics = Presentation.from_odp(odp)
    saved_diagnostics: list[tuple[str, str]] = presentation.save_odp(path)

    slide = presentation.slides[0]
    slide_id: int = slide.slide_id
    movie: Shape = slide.shapes.add_movie(
        io.BytesIO(b""), 0, 0, 10, 10, poster_frame_image=None, mime_type="video/mp4"
    )
    media: tuple[MediaInfo, ...] = slide.media
    kind: str = media[0].kind
    clip: bytes | None = slide.extract_media(media[0].shape_id)
    slide.replace_media(media[0].shape_id, io.BytesIO(b""), mime_type="video/mp4")
    slide.remove_media(media[0].shape_id)

    shape = slide.shapes[0]
    fill: FillFormat = shape.fill
    fill.gradient()
    fill.gradient_angle = 45.0
    fill.gradient_path = "circle"
    stops: GradientStops = fill.gradient_stops
    stop: GradientStop = stops.append(0.5)
    position: float = stop.position
    stop.color.rgb = RGBColor(1, 2, 3)
    stop.color.alpha = 0.5
    alpha: float | None = stop.color.alpha
    del stops[0]
    fill.patterned()
    fill.pattern = MSO_PATTERN_TYPE.DIAGONAL_CROSS
    pattern: MSO_PATTERN_TYPE | None = fill.pattern
    back: ColorFormat = fill.back_color
    fill.picture(io.BytesIO(b""))
    slide.background.fill.picture(path)

    paragraph = shape.text_frame.paragraphs[0]
    paragraph.auto_number = "arabicPeriod"
    scheme: str | None = paragraph.auto_number
    paragraph.auto_number_start = 3
    paragraph.bullet_color = "#123456"
    paragraph.bullet_color = (1, 2, 3)
    bullet_color: RGBColor | None = paragraph.bullet_color
    paragraph.bullet_size = 0.8
    paragraph.bullet_font = "Arial"
    font = paragraph.runs[0].font
    font.baseline = 0.3
    font.spacing = Pt(1)
    font.language = "en-US"
    font.east_asian_name = "Yu Gothic"
    font.complex_script_name = None
    baseline: float | None = font.baseline
    spacing: Length | None = font.spacing

    table = shape.table
    table.first_row = False
    table.horz_banding = True
    table.style_id = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    style_id: str | None = table.style_id
    banding: bool = table.vert_banding

    _ = (
        title, revision, section_indices, jpegs, tiff, pdfa, handout, handouts,
        odp_diagnostics, converted, read_diagnostics, saved_diagnostics, slide_id, movie,
        kind, clip, position, alpha, pattern, back, scheme, bullet_color, baseline, spacing,
        style_id, banding,
    )


if TYPE_CHECKING:
    BoundingBox()  # type: ignore[call-arg]
    AdjustmentCollection()  # type: ignore[call-arg]
    Background()  # type: ignore[call-arg]
    Cell()  # type: ignore[call-arg]
    ColorFormat()  # type: ignore[call-arg]
    CoreProperties()  # type: ignore[call-arg]
    GradientStop()  # type: ignore[call-arg]
    GradientStops()  # type: ignore[call-arg]
    MediaInfo()  # type: ignore[call-arg]
    Section()  # type: ignore[call-arg]
    FillFormat()  # type: ignore[call-arg]
    Image()  # type: ignore[call-arg]
    LineFormat()  # type: ignore[call-arg]
    LineEndFormat()  # type: ignore[call-arg]
    Comment()  # type: ignore[call-arg]
    CommentAuthor()  # type: ignore[call-arg]
    CommentReply()  # type: ignore[call-arg]
    Column()  # type: ignore[call-arg]
    ColumnCollection()  # type: ignore[call-arg]
    Font()  # type: ignore[call-arg]
    Hyperlink()  # type: ignore[call-arg]
    Paragraph()  # type: ignore[call-arg]
    ParagraphCollection()  # type: ignore[call-arg]
    PlaceholderCollection()  # type: ignore[call-arg]
    Row()  # type: ignore[call-arg]
    RowCollection()  # type: ignore[call-arg]
    Run()  # type: ignore[call-arg]
    RunCollection()  # type: ignore[call-arg]
    ShadowFormat()  # type: ignore[call-arg]
    Shape()  # type: ignore[call-arg]
    ShapeClickAction()  # type: ignore[call-arg]
    ShapeCollection()  # type: ignore[call-arg]
    ShapeHyperlink()  # type: ignore[call-arg]
    Slide()  # type: ignore[call-arg]
    SlideCollection()  # type: ignore[call-arg]
    SlideLayout()  # type: ignore[call-arg]
    SlideLayoutCollection()  # type: ignore[call-arg]
    Table()  # type: ignore[call-arg]
    TextFrame()  # type: ignore[call-arg]
    TextFrameLayout()  # type: ignore[call-arg]
    TextLineLayout()  # type: ignore[call-arg]
    ValidationIssue()  # type: ignore[call-arg]
