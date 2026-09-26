import io
from pathlib import Path
from typing import TYPE_CHECKING, Callable

from rpptx import (
    Comment,
    CommentAuthor,
    CommentReply,
    Inches,
    Length,
    MSO_SHAPE,
    Presentation,
    Pt,
    RGBColor,
)
from rpptx._rpptx import (
    AdjustmentCollection,
    Background,
    Cell,
    ColorFormat,
    Column,
    ColumnCollection,
    FillFormat,
    Font,
    Image,
    LineFormat,
    Paragraph,
    ParagraphCollection,
    PlaceholderCollection,
    Run,
    RunCollection,
    Shape,
    ShapeCollection,
    Slide,
    SlideCollection,
    SlideLayout,
    SlideLayoutCollection,
    Table,
    TextFrame,
)
from rpptx.enum.dml import MSO_FILL_TYPE
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
    slide_width: Length | None = presentation.slide_width
    presentation.slide_height = Inches(6)
    current_slide = presentation.slides[0]
    slide_layout: SlideLayout = current_slide.slide_layout
    layout_index: int = presentation.slide_layouts.index(slide_layout)
    same_layout: bool = slide_layout == presentation.slide_layouts[0]
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
    shape.name = "Typed"
    shape.rotation = 15.0
    rotation: float = shape.rotation
    shape_type: MSO_SHAPE_TYPE | None = shape.shape_type
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
    shape_xml: bytes = shape.xml
    connector: Shape = presentation.slides[0].shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, 0, 0, Inches(1), Inches(1)
    )
    group: Shape = presentation.slides[0].shapes.add_group_shape()
    picture = presentation.slides[0].shapes.add_picture(
        io.BytesIO(b""), 0, 0
    )
    image: Image = picture.image
    blob: bytes = image.blob
    picture.replace_image(b"")
    picture.replace_image(path)
    presentation.slides[0].shapes.remove(group)
    presentation.slides.move(0, -1)
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
        slide_width,
        layout_index,
        same_layout,
        hidden,
        background_fill,
        follows_master,
        rotation,
        shape_type,
        first_adjustment,
        all_adjustments,
        fill_type,
        rgb,
        line_width,
        line_fill,
        shape_xml,
        connector,
        blob,
        image.content_type,
        image.ext,
    )


if TYPE_CHECKING:
    AdjustmentCollection()  # type: ignore[call-arg]
    Background()  # type: ignore[call-arg]
    Cell()  # type: ignore[call-arg]
    ColorFormat()  # type: ignore[call-arg]
    FillFormat()  # type: ignore[call-arg]
    Image()  # type: ignore[call-arg]
    LineFormat()  # type: ignore[call-arg]
    Comment()  # type: ignore[call-arg]
    CommentAuthor()  # type: ignore[call-arg]
    CommentReply()  # type: ignore[call-arg]
    Column()  # type: ignore[call-arg]
    ColumnCollection()  # type: ignore[call-arg]
    Font()  # type: ignore[call-arg]
    Paragraph()  # type: ignore[call-arg]
    ParagraphCollection()  # type: ignore[call-arg]
    PlaceholderCollection()  # type: ignore[call-arg]
    Run()  # type: ignore[call-arg]
    RunCollection()  # type: ignore[call-arg]
    Shape()  # type: ignore[call-arg]
    ShapeCollection()  # type: ignore[call-arg]
    Slide()  # type: ignore[call-arg]
    SlideCollection()  # type: ignore[call-arg]
    SlideLayout()  # type: ignore[call-arg]
    SlideLayoutCollection()  # type: ignore[call-arg]
    Table()  # type: ignore[call-arg]
    TextFrame()  # type: ignore[call-arg]
