from pathlib import Path
from typing import TYPE_CHECKING, Callable

from rpptx import (
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
    RGBColor,
)
from rpptx._rpptx import (
    Cell,
    Column,
    ColumnCollection,
    Font,
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
    )


if TYPE_CHECKING:
    Cell()  # type: ignore[call-arg]
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
