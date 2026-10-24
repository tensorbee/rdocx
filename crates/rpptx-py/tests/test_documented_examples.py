import importlib.metadata
import io
import math
import re
import struct
import sys
import threading
import zipfile
import zlib

import pytest

from rpptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR


def _png_chunk(kind, data):
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))


def _tiny_png(red=0x20, green=0x80, blue=0xE0):
    signature = b"\x89PNG\r\n\x1a\n"
    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    pixels = zlib.compress(bytes((0, red, green, blue)))
    return (
        signature
        + _png_chunk(b"IHDR", header)
        + _png_chunk(b"IDAT", pixels)
        + _png_chunk(b"IEND", b"")
    )


def _write_tiny_png(path):
    path.write_bytes(_tiny_png())


def test_python_pptx_getting_started_examples_run_with_global_revision_refetches(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    _write_tiny_png(tmp_path / "monty-truth.png")

    # Hello World
    from rpptx import Presentation

    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "Hello, World!"
    subtitle = prs.slides[0].placeholders[1]
    subtitle.text = "python-pptx was here!"

    prs.save("test.pptx")

    # Bullet slide
    from rpptx import Presentation

    prs = Presentation()
    bullet_slide_layout = prs.slide_layouts[1]

    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes

    title_shape = shapes.title
    body_shape = shapes.placeholders[1]

    title_shape.text = "Adding a Bullet Slide"

    body_shape = prs.slides[0].shapes.placeholders[1]
    tf = body_shape.text_frame
    tf.text = "Find the bullet slide layout"

    tf = prs.slides[0].shapes.placeholders[1].text_frame
    p = tf.add_paragraph()
    p.text = "Use _TextFrame.text for first bullet"
    p = prs.slides[0].shapes.placeholders[1].text_frame.paragraphs[1]
    p.level = 1

    tf = prs.slides[0].shapes.placeholders[1].text_frame
    p = tf.add_paragraph()
    p.text = "Use _TextFrame.add_paragraph() for subsequent bullets"
    p = prs.slides[0].shapes.placeholders[1].text_frame.paragraphs[2]
    p.level = 2

    prs.save("test.pptx")

    # add_textbox()
    from rpptx import Presentation
    from rpptx.util import Inches, Pt

    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)

    left = top = width = height = Inches(1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame

    tf.text = "This is text inside a textbox"

    tf = prs.slides[0].shapes[-1].text_frame
    p = tf.add_paragraph()
    p.text = "This is a second paragraph that's bold"
    p = prs.slides[0].shapes[-1].text_frame.paragraphs[1]
    p.font.bold = True

    tf = prs.slides[0].shapes[-1].text_frame
    p = tf.add_paragraph()
    p.text = "This is a third paragraph that's big"
    p = prs.slides[0].shapes[-1].text_frame.paragraphs[2]
    p.font.size = Pt(40)

    prs.save("test.pptx")

    # add_picture()
    from rpptx import Presentation
    from rpptx.util import Inches

    img_path = "monty-truth.png"

    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)

    left = top = Inches(1)
    pic = slide.shapes.add_picture(img_path, left, top)

    slide = prs.slides[0]
    left = Inches(5)
    height = Inches(5.5)
    pic = slide.shapes.add_picture(img_path, left, top, height=height)

    prs.save("test.pptx")

    # add_shape()
    from rpptx import Presentation
    from rpptx.enum.shapes import MSO_SHAPE
    from rpptx.util import Inches

    prs = Presentation()
    title_only_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_only_slide_layout)
    shapes = slide.shapes

    shapes.title.text = "Adding an AutoShape"
    shapes = prs.slides[0].shapes

    left = Inches(0.93)
    top = Inches(3.0)
    width = Inches(1.75)
    height = Inches(1.0)

    shape = shapes.add_shape(MSO_SHAPE.PENTAGON, left, top, width, height)
    shape.text = "Step 1"

    left = left + width - Inches(0.4)
    width = Inches(2.0)

    for n in range(2, 6):
        shapes = prs.slides[0].shapes
        shape = shapes.add_shape(MSO_SHAPE.CHEVRON, left, top, width, height)
        shape.text = "Step %d" % n
        left = left + width - Inches(0.4)

    prs.save("test.pptx")

    # add_table()
    from rpptx import Presentation
    from rpptx.util import Inches

    prs = Presentation()
    title_only_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_only_slide_layout)
    shapes = slide.shapes

    shapes.title.text = "Adding a Table"
    shapes = prs.slides[0].shapes

    rows = cols = 2
    left = top = Inches(2.0)
    width = Inches(6.0)
    height = Inches(0.8)

    table = shapes.add_table(rows, cols, left, top, width, height).table

    table.columns[0].width = Inches(2.0)
    table.columns[1].width = Inches(4.0)

    table.cell(0, 0).text = "Foo"
    table.cell(0, 1).text = "Bar"

    table.cell(1, 0).text = "Baz"
    table.cell(1, 1).text = "Qux"

    prs.save("test.pptx")

    # Extract all text from slides
    from rpptx import Presentation

    path_to_presentation = "test.pptx"
    prs = Presentation(path_to_presentation)

    text_runs = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text_runs.append(run.text)

    assert text_runs == ["Adding a Table"]
    assert (tmp_path / "test.pptx").is_file()


def test_lazy_collections_and_stale_handles_are_loud():
    import rpptx

    prs = rpptx.Presentation()
    first = prs.slides.add_slide(prs.slide_layouts[6])
    held = first.shapes.add_textbox(0, 0, 100, 100)
    prs.slides.add_slide(prs.slide_layouts[6])

    assert len(prs.slides) == 2
    assert list(prs.slides[:1])[0].shapes[-1].text == ""
    assert held.text == ""
    prs.slides[0].shapes.remove(held)
    with pytest.raises(rpptx.StaleElementError, match=r"First invalidated by ShapeCollection.remove "):
        _ = held.text


def test_omitted_placeholder_index_resolves_as_default_zero():
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.placeholders[0].text = "default zero"
    assert prs.slides[0].shapes.title.text == "default zero"


def _assert_stale_after_exactly_one_bump(rpptx, operation):
    with pytest.raises(rpptx.StaleElementError) as raised:
        operation()
    revisions = re.search(
        r"revision (\d+).*revision (\d+)", str(raised.value)
    )
    assert revisions is not None
    captured, current = map(int, revisions.groups())
    assert current == captured + 1


def _assert_exact_stale(rpptx, operation, kind, captured, current, recovery, cause):
    expected = (
        f"stale {kind} handle was created at document revision {captured}, but the "
        f"document is now at revision {current} (a structural change "
        f"invalidated it). First invalidated by {cause} at committed revision {current}. {recovery}"
    )
    with pytest.raises(rpptx.StaleElementError) as raised:
        operation()
    assert str(raised.value) == expected


def test_structural_append_retires_slide_inventory_and_preserves_existing_owners():
    import rpptx

    prs = rpptx.Presentation()
    layouts = prs.slide_layouts
    layout = layouts[6]
    slides = prs.slides
    slide = slides.add_slide(layout)
    _assert_exact_stale(rpptx, lambda: len(slides), "slide collection", 0, 1,
                        "Re-fetch it with prs.slides.", "SlideCollection.add_slide")
    assert len(layouts) > 6
    assert layout.name == prs.slide_layouts[6].name
    shapes = slide.shapes
    shape = shapes.add_textbox(0, 0, 100, 100)
    assert len(slide.shapes) == len(shapes) == 1
    assert shape.text == ""


def test_stale_shape_text_and_table_paths_report_exact_recovery():
    import rpptx

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    shape = prs.slides[0].shapes.add_textbox(0, 0, 100, 100)
    shape.text = "before"
    prs.slides[0].shapes.add_table(1, 1, 0, 0, 100, 100)

    slide = prs.slides[0]
    shapes = slide.shapes
    placeholders = slide.placeholders
    shape = shapes[0]
    frame = shape.text_frame
    paragraphs = frame.paragraphs
    paragraph = paragraphs[0]
    runs = paragraph.runs
    run = runs[0]
    font = paragraph.font
    table = shapes[1].table
    columns = table.columns
    column = columns[0]
    cell = table.cell(0, 0)
    prs.slides.remove(prs.slides[0])

    cases = (
        (lambda: slide.shapes, "slide", "Re-fetch it with prs.slides[0]."),
        (lambda: len(shapes), "shape collection", "Re-fetch it with prs.slides[0].shapes."),
        (lambda: placeholders[0], "placeholder collection", "Re-fetch it with prs.slides[0].placeholders."),
        (lambda: shape.text, "shape", "Re-fetch it with prs.slides[0].shapes[0]."),
        (lambda: frame.text, "text frame", "Re-fetch it with prs.slides[0].shapes[0].text_frame."),
        (lambda: len(paragraphs), "paragraph collection", "Re-fetch it with prs.slides[0].shapes[0].text_frame.paragraphs."),
        (lambda: paragraph.text, "paragraph", "Re-fetch it with prs.slides[0].shapes[0].text_frame.paragraphs[0]."),
        (lambda: len(runs), "run collection", "Re-fetch it with prs.slides[0].shapes[0].text_frame.paragraphs[0].runs."),
        (lambda: run.text, "run", "Re-fetch it with prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]."),
        (lambda: font.bold, "font", "Re-fetch it with prs.slides[0].shapes[0].text_frame.paragraphs[0].font."),
        (lambda: table.columns, "table", "Re-fetch it with prs.slides[0].shapes[1].table."),
        (lambda: len(columns), "column collection", "Re-fetch it with prs.slides[0].shapes[1].table.columns."),
        (lambda: column.width, "column", "Re-fetch it with prs.slides[0].shapes[1].table.columns[0]."),
        (lambda: cell.text, "cell", "Re-fetch it with prs.slides[0].shapes[1].table.cell(0, 0)."),
    )
    for operation, kind, recovery in cases:
        _assert_exact_stale(rpptx, operation, kind, 4, 5, recovery, "SlideCollection.remove")


def test_whole_text_replacement_stales_descendant_handles_once():
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 100, 100)
    shape.text = "before"
    shape = prs.slides[0].shapes[-1]
    frame = shape.text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    shape.text = "after"
    for operation in (
        lambda: paragraph.text,
        lambda: run.text,
    ):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
    assert shape.text == frame.text == prs.slides[0].shapes[-1].text == "after"

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_textbox(0, 0, 100, 100).text_frame
    frame.text = "before"
    frame = prs.slides[0].shapes[-1].text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    frame.text = "after"
    for operation in (lambda: paragraph.text, lambda: run.text):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
    assert frame.text == prs.slides[0].shapes[-1].text_frame.text == "after"

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_textbox(0, 0, 100, 100).text_frame
    frame.text = "before"
    frame = prs.slides[0].shapes[-1].text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    paragraph.text = "after"
    assert paragraph.text == "after"
    _assert_exact_stale(
        rpptx, lambda: run.text, "run", 3, 4,
        "Re-fetch it with prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].",
        "Paragraph.text",
    )
    assert prs.slides[0].shapes[-1].text_frame.paragraphs[0].text == "after"


def test_run_text_replaces_in_place_and_keeps_every_handle_live():
    import rpptx

    prs = _textbox_presentation(rpptx)
    prs.slides[0].shapes[0].text_frame.paragraphs[0].add_run(" world")
    slides = prs.slides
    slide = slides[0]
    shapes = slide.shapes
    shape = shapes[0]
    left = shape.left
    frame = shape.text_frame
    paragraph = frame.paragraphs[0]
    runs = paragraph.runs
    first, second = runs[0], runs[1]
    font = first.font

    for value in ("HELLO", "line\nfeed", "vertical\vtab", "tab\tstop", ""):
        first.text = value
        stored = value.replace("\v", "_x000B_")
        assert first.text == stored
        assert second.text == " world"
        assert paragraph.text == frame.text == shape.text == stored + " world"
        assert len(runs) == 2
        assert shape.left == left
        assert len(slides) == len(shapes) == len(slide.shapes) == 1
        assert font.bold is None

    second.text = "world"
    font.bold = True
    assert (first.text, second.text, first.font.bold) == ("", "world", True)
    fresh = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs
    assert [run.text for run in fresh] == ["", "world"]

    paragraph.text = "whole"
    assert shape.left == left
    assert len(slide.shapes) == 1
    for operation in (lambda: second.text, lambda: font.bold):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
    assert prs.slides[0].shapes[0].text == "whole"


def test_nested_group_paths_report_exact_recovery(tmp_path):
    if importlib.util.find_spec("pptx") is None:
        pytest.skip("python-pptx oracle is installed only for the differential gate")

    import rpptx
    from pptx import Presentation as OraclePresentation

    source = tmp_path / "nested-groups.pptx"
    oracle = OraclePresentation()
    slide = oracle.slides.add_slide(oracle.slide_layouts[6])
    outer = slide.shapes.add_group_shape()
    inner = outer.shapes.add_group_shape()
    textbox = inner.shapes.add_textbox(0, 0, 100, 100)
    textbox.text = "nested"
    oracle.save(source)

    prs = rpptx.Presentation(source)
    inner = prs.slides[0].shapes[0].shapes[0]
    nested_shapes = inner.shapes
    shape = nested_shapes[0]
    frame = shape.text_frame
    paragraphs = frame.paragraphs
    paragraph = paragraphs[0]
    runs = paragraph.runs
    run = runs[0]
    font = paragraph.font
    prs.slides.remove(prs.slides[0])

    cases = (
        (lambda: len(nested_shapes), "shape collection", "Re-fetch it with prs.slides[0].shapes[0].shapes[0].shapes."),
        (lambda: shape.text, "shape", "Re-fetch it with prs.slides[0].shapes[0].shapes[0].shapes[0]."),
        (lambda: frame.text, "text frame", "Re-fetch it with prs.slides[0].shapes[0].shapes[0].shapes[0].text_frame."),
        (lambda: len(paragraphs), "paragraph collection", "Re-fetch it with prs.slides[0].shapes[0].shapes[0].shapes[0].text_frame.paragraphs."),
        (lambda: paragraph.text, "paragraph", "Re-fetch it with prs.slides[0].shapes[0].shapes[0].shapes[0].text_frame.paragraphs[0]."),
        (lambda: len(runs), "run collection", "Re-fetch it with prs.slides[0].shapes[0].shapes[0].shapes[0].text_frame.paragraphs[0].runs."),
        (lambda: run.text, "run", "Re-fetch it with prs.slides[0].shapes[0].shapes[0].shapes[0].text_frame.paragraphs[0].runs[0]."),
        (lambda: font.bold, "font", "Re-fetch it with prs.slides[0].shapes[0].shapes[0].shapes[0].text_frame.paragraphs[0].font."),
    )
    for operation, kind, recovery in cases:
        _assert_exact_stale(rpptx, operation, kind, 0, 1, recovery, "SlideCollection.remove")


def test_unit_constructors_truncate_fractional_values_toward_zero():
    import rpptx
    from rpptx.enum.shapes import MSO_SHAPE
    from rpptx.util import Inches, Length, Pt

    for constructor, factor in ((Inches, 914_400.0), (Pt, 12_700.0)):
        assert constructor(1.75 / factor) == 1
        assert constructor(-1.75 / factor) == -1
    assert Length(914_400).inches == 1.0
    assert Length(12_700).pt == 1.0
    assert Length(-1).emu == -1
    assert MSO_SHAPE.PENTAGON == 51
    assert MSO_SHAPE.CHEVRON == 52
    assert issubclass(rpptx.PackageError, rpptx.RpptxError)
    assert issubclass(rpptx.XmlError, rpptx.RpptxError)
    assert issubclass(rpptx.StaleElementError, rpptx.RpptxError)


def test_missing_package_raises_the_named_package_error(tmp_path):
    import rpptx

    with pytest.raises(rpptx.PackageError):
        rpptx.Presentation(tmp_path / "missing.pptx")


def _author_documented_decks(Presentation, Inches, Pt, MSO_SHAPE, root, image_path):
    root.mkdir()
    paths = {name: root / f"{name}.pptx" for name in (
        "hello", "bullet", "textbox", "picture", "shapes", "table"
    )}

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Hello, World!"
    subtitle = prs.slides[0].placeholders[1]
    subtitle.text = "python-pptx was here!"
    prs.save(paths["hello"])

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    shapes = slide.shapes
    shapes.title.text = "Adding a Bullet Slide"
    shapes = prs.slides[0].shapes
    frame = shapes.placeholders[1].text_frame
    frame.text = "Find the bullet slide layout"
    frame = prs.slides[0].shapes.placeholders[1].text_frame
    paragraph = frame.add_paragraph()
    paragraph.text = "Use _TextFrame.text for first bullet"
    paragraph = prs.slides[0].shapes.placeholders[1].text_frame.paragraphs[1]
    paragraph.level = 1
    frame = prs.slides[0].shapes.placeholders[1].text_frame
    paragraph = frame.add_paragraph()
    paragraph.text = "Use _TextFrame.add_paragraph() for subsequent bullets"
    paragraph = prs.slides[0].shapes.placeholders[1].text_frame.paragraphs[2]
    paragraph.level = 2
    prs.save(paths["bullet"])

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(1), Inches(1)).text_frame
    frame.text = "This is text inside a textbox"
    frame = prs.slides[0].shapes[-1].text_frame
    paragraph = frame.add_paragraph()
    paragraph.text = "This is a second paragraph that's bold"
    paragraph = prs.slides[0].shapes[-1].text_frame.paragraphs[1]
    paragraph.font.bold = True
    frame = prs.slides[0].shapes[-1].text_frame
    paragraph = frame.add_paragraph()
    paragraph.text = "This is a third paragraph that's big"
    paragraph = prs.slides[0].shapes[-1].text_frame.paragraphs[2]
    paragraph.font.size = Pt(40)
    prs.save(paths["textbox"])

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(str(image_path), Inches(1), Inches(1))
    slide = prs.slides[0]
    slide.shapes.add_picture(
        str(image_path), Inches(5), Inches(1), height=Inches(5.5)
    )
    prs.save(paths["picture"])

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    shapes = slide.shapes
    shapes.title.text = "Adding an AutoShape"
    shapes = prs.slides[0].shapes
    left = Inches(0.93)
    top = Inches(3.0)
    width = Inches(1.75)
    height = Inches(1.0)
    shape = shapes.add_shape(MSO_SHAPE.PENTAGON, left, top, width, height)
    shape.text = "Step 1"
    left = left + width - Inches(0.4)
    width = Inches(2.0)
    for number in range(2, 6):
        shapes = prs.slides[0].shapes
        shape = shapes.add_shape(MSO_SHAPE.CHEVRON, left, top, width, height)
        shape.text = f"Step {number}"
        left = left + width - Inches(0.4)
    prs.save(paths["shapes"])

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    shapes = slide.shapes
    shapes.title.text = "Adding a Table"
    shapes = prs.slides[0].shapes
    table = shapes.add_table(
        2, 2, Inches(2), Inches(2), Inches(6), Inches(0.8)
    ).table
    table.columns[0].width = Inches(2)
    table.columns[1].width = Inches(4)
    for row, values in enumerate((("Foo", "Bar"), ("Baz", "Qux"))):
        for column, value in enumerate(values):
            table.cell(row, column).text = value
    prs.save(paths["table"])
    return paths


def _normalized_table(table):
    columns = len(table.columns)
    rows = []
    row = 0
    while True:
        try:
            rows.append(tuple(table.cell(row, column).text for column in range(columns)))
        except IndexError:
            break
        row += 1
    return tuple(int(column.width) for column in table.columns), tuple(rows)


def _normalized_presentation(prs):
    slides = []
    for slide in prs.slides:
        shapes = []
        for shape in slide.shapes:
            paragraphs = ()
            if shape.has_text_frame:
                paragraphs = tuple(
                    (
                        paragraph.text,
                        paragraph.level,
                        paragraph.font.bold,
                        int(paragraph.font.size) if paragraph.font.size is not None else None,
                        tuple(run.text for run in paragraph.runs),
                    )
                    for paragraph in shape.text_frame.paragraphs
                )
            shapes.append(
                (
                    shape.has_text_frame,
                    paragraphs,
                    shape.has_table,
                    _normalized_table(shape.table) if shape.has_table else None,
                )
            )
        slides.append(tuple(shapes))
    return tuple(slides)


def _normalized_text_extraction(prs):
    return tuple(
        run.text
        for slide in prs.slides
        for shape in slide.shapes
        if shape.has_text_frame
        for paragraph in shape.text_frame.paragraphs
        for run in paragraph.runs
    )


def _normalized_documented_records(Presentation, paths):
    records = {
        name: _normalized_presentation(Presentation(path))
        for name, path in paths.items()
    }
    records["extract"] = _normalized_text_extraction(Presentation(paths["table"]))
    return records


def _oracle_writer_contract(Presentation, paths):
    def auto_shape_type(shape):
        if int(shape.shape_type) != 1:
            return None
        try:
            return int(shape.auto_shape_type)
        except (AttributeError, TypeError, ValueError):
            return None

    records = {}
    for name, path in paths.items():
        prs = Presentation(path)
        records[name] = tuple(
            tuple(
                (
                    int(shape.shape_type),
                    (int(shape.left), int(shape.top), int(shape.width), int(shape.height)),
                    auto_shape_type(shape),
                    int(shape.placeholder_format.idx) if shape.is_placeholder else None,
                )
                for shape in slide.shapes
            )
            for slide in prs.slides
        )
    return records


def test_pinned_python_pptx_bidirectional_seven_example_records(tmp_path):
    if importlib.util.find_spec("pptx") is None:
        pytest.skip("python-pptx oracle is installed only for the differential gate")
    assert importlib.metadata.version("python-pptx") == "1.0.2"

    import rpptx
    from rpptx.enum.shapes import MSO_SHAPE as RpptxShape
    from rpptx.util import Inches as RpptxInches
    from rpptx.util import Pt as RpptxPt
    from pptx import Presentation as OraclePresentation
    from pptx.enum.shapes import MSO_SHAPE as OracleShape
    from pptx.util import Inches as OracleInches
    from pptx.util import Pt as OraclePt

    image_path = tmp_path / "monty-truth.png"
    _write_tiny_png(image_path)
    # Compare writers against one explicit independent template.
    template = io.BytesIO()
    OraclePresentation().save(template)
    template_bytes = template.getvalue()
    authored = (
        _author_documented_decks(
            lambda: rpptx.Presentation.from_bytes(template_bytes),
            RpptxInches,
            RpptxPt,
            RpptxShape,
            tmp_path / "rpptx-authored",
            image_path,
        ),
        _author_documented_decks(
            lambda: OraclePresentation(io.BytesIO(template_bytes)),
            OracleInches,
            OraclePt,
            OracleShape,
            tmp_path / "python-pptx-authored",
            image_path,
        ),
    )
    normalized = []
    writer_contracts = []
    for paths in authored:
        rpptx_records = _normalized_documented_records(rpptx.Presentation, paths)
        oracle_records = _normalized_documented_records(OraclePresentation, paths)
        assert rpptx_records == oracle_records
        normalized.append(rpptx_records)
        writer_contracts.append(_oracle_writer_contract(OraclePresentation, paths))
    assert normalized[0] == normalized[1]
    assert writer_contracts[0] == writer_contracts[1]
    assert normalized[0]["extract"] == ("Adding a Table",)


def _add_speaker_notes(source, target, text):
    from pptx import Presentation as OraclePresentation
    from pptx.enum.shapes import PP_PLACEHOLDER
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.oxml import parse_xml

    assert importlib.metadata.version("python-pptx") == "1.0.2"
    oracle = OraclePresentation(source)
    notes = oracle.slides[0].notes_slide
    kinds = [shape.placeholder_format.type for shape in notes.placeholders]
    for kind in (PP_PLACEHOLDER.SLIDE_IMAGE, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.SLIDE_NUMBER):
        assert kinds.count(kind) == 1
    local_masters = [rel for rel in notes.part.rels.values() if rel.reltype == RT.NOTES_MASTER]
    global_masters = [rel for rel in oracle.part.rels.values() if rel.reltype == RT.NOTES_MASTER]
    assert len(local_masters) == len(global_masters) == 1
    assert not local_masters[0].is_external and not global_masters[0].is_external
    assert local_masters[0].target_part.partname == global_masters[0].target_part.partname
    # python-pptx 1.0.2 relates its notes master without registering its XML id.
    # Build a complete independent fixture before testing the consumer.
    if not oracle._element.xpath("./p:notesMasterIdLst"):
        registration = parse_xml(
            '<p:notesMasterIdLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:notesMasterId/></p:notesMasterIdLst>'
        )
        registration[0].set(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
            global_masters[0].rId,
        )
        prior = oracle._element.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}sldMasterIdLst"
        )
        oracle._element.insert(0 if prior is None else oracle._element.index(prior) + 1, registration)
    assert oracle._element.xpath("./p:notesMasterIdLst/p:notesMasterId/@r:id") == [global_masters[0].rId]
    notes.notes_text_frame.text = text
    run = notes.notes_text_frame.paragraphs[0].runs[0]
    run.font.bold = True
    run._r.get_or_add_rPr().append(parse_xml(
        '<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:ext uri="{6D487B31-4C56-4F45-AED3-4C741AC43E77}">'
        '<x:payload xmlns:x="urn:rdocx:test"/></a:ext></a:extLst>'
    ))
    oracle.save(target)


def test_presentation_render_comments_and_notes_match_native_snapshots(tmp_path):
    import rpptx

    author_id = "{11111111-1111-1111-1111-111111111111}"
    comment_id = "{22222222-2222-2222-2222-222222222222}"
    reply_id = "{33333333-3333-3333-3333-333333333333}"
    second_reply_id = "{44444444-4444-4444-4444-444444444444}"
    second_comment_id = "{55555555-5555-5555-5555-555555555555}"
    created = "2026-09-14T10:30:00Z"

    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    source = tmp_path / "source.pptx"
    with_notes = tmp_path / "with-notes.pptx"
    presentation.save(source)
    _add_speaker_notes(source, with_notes, "F-X094e speaker note")
    presentation = rpptx.Presentation(with_notes)
    slide = presentation.slides[0]

    pdf = presentation.to_pdf()
    slide_png = presentation.render_slide_to_png(0, dpi=72.0)
    assert pdf.startswith(b"%PDF-")
    assert slide_png is not None
    assert slide_png.startswith(b"\x89PNG\r\n\x1a\n")
    assert presentation.render_all_slides(dpi=72.0) == [slide_png]
    assert presentation.to_notes_pdf().startswith(b"%PDF-")
    assert presentation.render_all_notes(dpi=72.0)[0].startswith(
        b"\x89PNG\r\n\x1a\n"
    )
    assert slide.notes_text == "F-X094e speaker note"

    with pytest.raises(rpptx.RpptxError, match="GUID"):
        presentation.add_comment_author(
            id="not-a-guid",
            name="Invalid",
            user_id="invalid@example.com",
            provider_id="local",
        )
    assert slide.notes_text == "F-X094e speaker note"
    with pytest.raises(rpptx.RpptxError, match="unknown comment author"):
        slide.add_comment(
            id=comment_id,
            author_id="{99999999-9999-9999-9999-999999999999}",
            created=created,
            text="Must remain atomic",
        )
    assert slide.notes_text == "F-X094e speaker note"

    presentation.add_comment_author(
        id=author_id,
        name="Ada Lovelace",
        initials="AL",
        user_id="ada@example.com",
        provider_id="local",
    )
    author = presentation.comment_authors[0]
    assert (
        author.id,
        author.name,
        author.initials,
        author.user_id,
        author.provider_id,
    ) == (author_id, "Ada Lovelace", "AL", "ada@example.com", "local")
    with pytest.raises(AttributeError):
        author.name = "mutated"

    slide = presentation.slides[0]
    slide.add_comment(
        id=comment_id,
        author_id=author_id,
        created=created,
        text="Review this slide",
    )
    slide = presentation.slides[0]
    slide.reply_to_comment(
        comment_id,
        id=reply_id,
        author_id=author_id,
        created=created,
        text="First reply",
    )
    slide = presentation.slides[0]
    slide.reply_to_comment(
        comment_id,
        id=second_reply_id,
        author_id=author_id,
        created=created,
        text="Second reply",
    )
    slide = presentation.slides[0]
    slide.add_comment(
        id=second_comment_id,
        author_id=author_id,
        created=created,
        text="Move me first",
    )
    slide = presentation.slides[0]
    slide.move_comment(1, 0)
    slide = presentation.slides[0]
    slide.move_reply(comment_id, 1, 0)

    comments = presentation.slides[0].comments
    assert tuple(comment.id for comment in comments) == (second_comment_id, comment_id)
    comment = comments[1]
    assert (comment.id, comment.author_id, comment.status, comment.created, comment.text) == (
        comment_id,
        author_id,
        None,
        created,
        "Review this slide",
    )
    assert tuple(reply.id for reply in comment.replies) == (second_reply_id, reply_id)
    assert comment.replies[0].text == "Second reply"
    with pytest.raises(AttributeError):
        comment.text = "mutated"

    output = tmp_path / "comments.pptx"
    presentation.save(output)
    reopened = rpptx.Presentation(output)
    assert reopened.comment_authors == presentation.comment_authors
    assert reopened.slides[0].comments == presentation.slides[0].comments


def _text_layout_deck(tmp_path, width, text, body_properties=None):
    import rpptx

    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = slide.shapes.add_textbox(rpptx.Inches(1), rpptx.Inches(1), width, rpptx.Pt(36))
    shape.text = text
    if body_properties is None:
        return presentation
    source = tmp_path / "text-layout-source.pptx"
    target = tmp_path / "text-layout.pptx"
    presentation.save(source)
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    slide_xml = parts["ppt/slides/slide1.xml"]
    assert slide_xml.count(b"<a:bodyPr/>") == 1
    parts["ppt/slides/slide1.xml"] = slide_xml.replace(b"<a:bodyPr/>", body_properties)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return rpptx.Presentation(target)


def test_text_layout_reports_the_renderer_fit_at_full_and_reduced_width(tmp_path):
    import rpptx

    assert rpptx.Presentation().text_layout() == ()
    presentation = _text_layout_deck(tmp_path, rpptx.Inches(4), "Fits in one line")
    shape = presentation.slides[0].shapes[0]
    (frame,) = presentation.text_layout()
    assert isinstance(frame, rpptx.TextFrameLayout)
    assert (frame.slide_index, frame.shape_id, frame.name) == (0, shape.shape_id, shape.name)
    assert (frame.autofit, frame.font_scale, frame.overflow) == ("none", 1.0, False)
    assert isinstance(frame.frame, rpptx.BoundingBox)
    assert (frame.frame.x, frame.frame.y, frame.frame.width, frame.frame.height) == (
        72.0,
        72.0,
        288.0,
        36.0,
    )
    assert (frame.usable.x, frame.usable.y) == pytest.approx((79.2, 75.6))
    assert (frame.usable.width, frame.usable.height) == pytest.approx((273.6, 28.8))
    (line,) = frame.lines
    assert isinstance(line, rpptx.TextLineLayout)
    assert (line.paragraph_index, line.text, line.font_size) == (0, "Fits in one line", 18.0)
    assert line.bounds.x == pytest.approx(79.2)
    assert frame.usable.y <= line.bounds.y < line.baseline < line.bounds.y + line.bounds.height
    assert frame.height == pytest.approx(line.bounds.height)
    assert presentation.text_layout() == (frame,)
    with pytest.raises(AttributeError):
        frame.overflow = True  # type: ignore[misc]

    presentation.slides[0].shapes[0].text_frame.add_paragraph().text = "One line too many"
    (two_lines,) = presentation.text_layout()
    assert [(line.paragraph_index, line.text) for line in two_lines.lines] == [
        (0, "Fits in one line"),
        (1, "One line too many"),
    ]
    assert two_lines.overflow
    assert two_lines.height > two_lines.usable.height

    measured = _text_layout_deck(tmp_path, rpptx.Inches(8), "Fits only at full width")
    natural = measured.text_layout()[0].lines[0].bounds.width
    tight = _text_layout_deck(
        tmp_path, math.ceil((natural / 0.975 + 14.4) * 12_700), "Fits only at full width"
    )
    (full,) = tight.text_layout()
    (narrower,) = tight.text_layout(width_factor=0.95)
    assert (len(full.lines), full.overflow) == (1, False)
    assert (len(narrower.lines), narrower.overflow) == (2, True)
    assert narrower.usable.width == pytest.approx(full.usable.width * 0.95)
    assert "".join(line.text for line in narrower.lines) == "Fits only at full width"
    with pytest.raises(rpptx.RpptxError, match="width factor"):
        tight.text_layout(width_factor=0.0)
    with pytest.raises(TypeError):
        tight.text_layout(0.95)  # type: ignore[misc]

    scaled = _text_layout_deck(
        tmp_path,
        rpptx.Inches(4),
        "Stored scale",
        b'<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:normAutofit fontScale="50000"/></a:bodyPr>',
    )
    (scaled_frame,) = scaled.text_layout()
    assert (scaled_frame.autofit, scaled_frame.font_scale) == ("normal", 0.5)
    assert (scaled_frame.usable.x, scaled_frame.usable.height) == (72.0, 36.0)
    assert scaled_frame.lines[0].font_size == 9.0
    grown = _text_layout_deck(
        tmp_path, rpptx.Inches(4), "Stored extent", b"<a:bodyPr><a:spAutoFit/></a:bodyPr>"
    )
    assert grown.text_layout()[0].autofit == "shape"

    assert _assert_rpptx_releases_gil(lambda: tight.text_layout()) == (full,)


def test_notes_mutation_preserves_text_through_save_and_reopen(tmp_path):
    import rpptx

    presentation = rpptx.Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    source = tmp_path / "source.pptx"
    with_notes = tmp_path / "with-notes.pptx"
    output = tmp_path / "updated-notes.pptx"
    presentation.save(source)
    _add_speaker_notes(source, with_notes, "Draft speaker note")

    presentation = rpptx.Presentation(with_notes)
    held = presentation.slides[0]
    held.notes_text = "Final speaker note"
    assert presentation.slides[0].notes_text == "Final speaker note"
    assert held.notes_text == presentation.slides[0].notes_text

    presentation.save(output)
    assert rpptx.Presentation(output).slides[0].notes_text == "Final speaker note"
    with zipfile.ZipFile(output) as archive:
        notes_xml = archive.read("ppt/notesSlides/notesSlide1.xml").decode()
    import xml.etree.ElementTree as ET
    ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    source_notes = ET.fromstring(_package_parts(with_notes.read_bytes())["ppt/notesSlides/notesSlide1.xml"])
    expected_body = next(ph.attrib for ph in source_notes.iter(ns + "ph") if ph.get("type") == "body")
    actual_body = next(ph.attrib for ph in ET.fromstring(notes_xml).iter(ns + "ph") if ph.get("type") == "body")
    assert actual_body == expected_body
    assert 'b="1"' in notes_xml
    assert "x:payload" in notes_xml

    without_notes = rpptx.Presentation()
    held = without_notes.slides.add_slide(without_notes.slide_layouts[6])
    assert held.notes_text is None
    held.notes_text = "Created speaker note"
    assert without_notes.slides[0].notes_text == "Created speaker note"
    assert held.notes_text == without_notes.slides[0].notes_text


def test_python_round_three_authoring_and_inspection_is_typed_and_lossless(tmp_path):
    import rpptx

    source = tmp_path / "round-three-source.pptx"
    rich = tmp_path / "round-three-rich.pptx"
    with_notes = tmp_path / "round-three-notes.pptx"
    output = tmp_path / "round-three-output.pptx"
    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = slide.shapes.add_textbox(
        rpptx.Inches(1), rpptx.Inches(2), rpptx.Inches(3), rpptx.Inches(1)
    )
    shape.text = "formatted"
    presentation.save(source)

    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    slide_xml = parts["ppt/slides/slide1.xml"]
    slide_xml = slide_xml.replace(b"<a:bodyPr/>", b"<a:bodyPr><a:spAutoFit/></a:bodyPr>", 1)
    slide_xml = slide_xml.replace(
        b"<a:r>",
        b'<a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="123456"/>'
        b'</a:solidFill><a:latin typeface="Aptos"/><a:extLst><a:ext uri="round-three">'
        b'<x:payload xmlns:x="urn:rdocx:test"/></a:ext></a:extLst></a:rPr>',
        1,
    )
    parts["ppt/slides/slide1.xml"] = slide_xml
    with zipfile.ZipFile(rich, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    _add_speaker_notes(rich, with_notes, "draft note")

    presentation = rpptx.Presentation(with_notes)
    shape = presentation.slides[0].shapes[0]
    assert (shape.left, shape.top, shape.width, shape.height) == (
        rpptx.Inches(1),
        rpptx.Inches(2),
        rpptx.Inches(3),
        rpptx.Inches(1),
    )
    assert shape.shape_id is not None
    assert shape.name is not None
    assert shape.text_frame.autofit == "shape"
    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.font.name == "Aptos"
    assert run.font.size == rpptx.Pt(18)
    assert str(run.font.color.rgb) == "123456"
    run.text = "updated"
    presentation.slides[0].notes_text = "final note"
    presentation.save(output)

    reopened = rpptx.Presentation(output)
    assert reopened.slides[0].shapes[0].text == "updated"
    assert reopened.slides[0].notes_text == "final note"
    with zipfile.ZipFile(output) as archive:
        preserved = archive.read("ppt/slides/slide1.xml")
    assert b"urn:rdocx:test" in preserved
    assert b'sz="1800"' in preserved
    assert b"updated" in preserved


def _assert_rpptx_releases_gil(operation):
    gate = threading.Lock()
    gate.acquire()
    ready = threading.Event()
    progressed = threading.Event()

    def wait_for_detached_call():
        ready.set()
        gate.acquire()
        progressed.set()

    old_switch_interval = sys.getswitchinterval()
    sys.setswitchinterval(60.0)
    worker = threading.Thread(target=wait_for_detached_call)
    try:
        worker.start()
        assert ready.wait(timeout=5.0)
        gate.release()
        result = None
        for _ in range(64):
            result = operation()
            if progressed.is_set():
                break
        progressed_during_call = progressed.is_set()
    finally:
        if not worker.is_alive() and gate.locked():
            gate.release()
        worker.join(timeout=5.0)
        if gate.locked():
            gate.release()
        sys.setswitchinterval(old_switch_interval)

    assert not worker.is_alive()
    assert progressed_during_call, "Python worker made no progress during native call"
    assert result is not None
    return result


def test_slide_and_notes_rendering_release_the_gil():
    import rpptx

    presentation = rpptx.Presentation()
    for _ in range(3):
        presentation.slides.add_slide(presentation.slide_layouts[6])

    slides = _assert_rpptx_releases_gil(
        lambda: presentation.render_all_slides(dpi=72.0)
    )
    notes = _assert_rpptx_releases_gil(
        lambda: presentation.render_all_notes(dpi=72.0)
    )

    assert len(slides) == 3
    assert len(notes) == 3


def _textbox_presentation(rpptx):
    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = slide.shapes.add_textbox(
        rpptx.Inches(1), rpptx.Inches(1), rpptx.Inches(4), rpptx.Inches(2)
    )
    shape.text = "hello"
    return presentation


def _text_body_xml(path):
    with zipfile.ZipFile(path) as archive:
        slide_xml = archive.read("ppt/slides/slide1.xml").decode()
    start = re.search(r"<p:txBody(?=[\s>])[^>]*>", slide_xml)
    assert start is not None, "expected the qualified PresentationML text-body owner"
    return slide_xml[start.start() : slide_xml.index("</p:txBody>", start.end())]


def _replace_in_slide(source, target, old, new):
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    slide_xml = parts["ppt/slides/slide1.xml"].decode()
    assert old in slide_xml
    parts["ppt/slides/slide1.xml"] = slide_xml.replace(old, new, 1).encode()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def test_text_frame_margins_anchor_wrap_and_auto_size_round_trip_and_clear(tmp_path):
    import rpptx
    from rpptx import MSO_ANCHOR, MSO_AUTO_SIZE

    presentation = _textbox_presentation(rpptx)
    frame = presentation.slides[0].shapes[0].text_frame
    assert (
        frame.margin_left,
        frame.margin_right,
        frame.margin_top,
        frame.margin_bottom,
        frame.vertical_anchor,
        frame.word_wrap,
        frame.auto_size,
    ) == (None, None, None, None, None, None, None)

    frame.margin_left = rpptx.Inches(0.25)
    frame.margin_right = 0
    frame.margin_top = rpptx.Pt(3)
    frame.margin_bottom = rpptx.Pt(4)
    frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    frame.word_wrap = False
    frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    assert isinstance(frame.margin_left, rpptx.Length)
    output = tmp_path / "frame.pptx"
    presentation.save(output)

    import xml.etree.ElementTree as ET
    p_ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    body = ET.fromstring(_package_parts(output.read_bytes())["ppt/slides/slide1.xml"]).find(".//" + p_ns + "txBody")
    assert body is not None and body.tag == p_ns + "txBody"
    assert [child.tag for child in body] == [a_ns + "bodyPr", a_ns + "lstStyle", a_ns + "p"]
    properties = body.find(a_ns + "bodyPr")
    assert properties.attrib == {"lIns": "228600", "tIns": "38100", "rIns": "0", "bIns": "50800", "anchor": "b", "wrap": "none"}
    assert [child.tag for child in properties] == [a_ns + "spAutoFit"]
    presentation = rpptx.Presentation(output)
    frame = presentation.slides[0].shapes[0].text_frame
    assert (
        frame.margin_left,
        frame.margin_right,
        frame.margin_top,
        frame.margin_bottom,
    ) == (rpptx.Inches(0.25), 0, rpptx.Pt(3), rpptx.Pt(4))
    assert frame.vertical_anchor is MSO_ANCHOR.BOTTOM
    assert frame.word_wrap is False
    assert frame.auto_size is MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    assert frame.autofit == "shape"

    for name in (
        "margin_left",
        "margin_right",
        "margin_top",
        "margin_bottom",
        "vertical_anchor",
        "word_wrap",
        "auto_size",
    ):
        setattr(frame, name, None)
    presentation.save(output)
    body = ET.fromstring(_package_parts(output.read_bytes())["ppt/slides/slide1.xml"]).find(".//" + p_ns + "txBody")
    properties = body.find(a_ns + "bodyPr")
    assert properties is not None and properties.attrib == {} and list(properties) == []


def test_text_frame_margins_read_universal_measures_and_the_justified_anchors(tmp_path):
    import rpptx
    from rpptx import MSO_ANCHOR

    source = tmp_path / "source.pptx"
    measured = tmp_path / "measured.pptx"
    _textbox_presentation(rpptx).save(source)
    _replace_in_slide(
        source,
        measured,
        "<a:bodyPr/>",
        '<a:bodyPr lIns="0.1in" rIns="1pc" tIns="2mm" bIns="3pt" anchor="dist"/>',
    )

    frame = rpptx.Presentation(measured).slides[0].shapes[0].text_frame
    assert (
        frame.margin_left,
        frame.margin_right,
        frame.margin_top,
        frame.margin_bottom,
    ) == (91440, 152400, 72000, 38100)
    assert frame.vertical_anchor is MSO_ANCHOR.DISTRIBUTE


def test_paragraph_alignment_spacing_and_indents_round_trip_and_clear(tmp_path):
    import rpptx
    from rpptx import PP_ALIGN

    presentation = _textbox_presentation(rpptx)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    names = (
        "alignment",
        "line_spacing",
        "space_before",
        "space_after",
        "left_indent",
        "right_indent",
        "first_line_indent",
    )
    assert [getattr(paragraph, name) for name in names] == [None] * len(names)

    paragraph.alignment = PP_ALIGN.JUSTIFY
    paragraph.line_spacing = 1.25
    paragraph.space_before = rpptx.Pt(6)
    paragraph.space_after = 0.5
    paragraph.left_indent = rpptx.Inches(0.5)
    paragraph.right_indent = rpptx.Inches(0.25)
    paragraph.first_line_indent = -rpptx.Inches(0.25)
    output = tmp_path / "paragraph.pptx"
    presentation.save(output)

    assert (
        '<a:pPr marL="457200" marR="228600" indent="-228600" algn="just">'
        '<a:lnSpc><a:spcPct val="125000"/></a:lnSpc>'
        '<a:spcBef><a:spcPts val="600"/></a:spcBef>'
        '<a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr>'
    ) in _text_body_xml(output)
    presentation = rpptx.Presentation(output)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    assert paragraph.alignment is PP_ALIGN.JUSTIFY
    assert paragraph.line_spacing == 1.25 and isinstance(paragraph.line_spacing, float)
    assert paragraph.space_before == rpptx.Pt(6)
    assert isinstance(paragraph.space_before, rpptx.Length)
    assert paragraph.space_after == 0.5 and isinstance(paragraph.space_after, float)
    assert paragraph.left_indent == rpptx.Inches(0.5)
    assert paragraph.right_indent == rpptx.Inches(0.25)
    assert paragraph.first_line_indent == -rpptx.Inches(0.25)

    paragraph.line_spacing = rpptx.Pt(18)
    paragraph.space_after = rpptx.Pt(3)
    assert paragraph.line_spacing == rpptx.Pt(18)
    assert isinstance(paragraph.line_spacing, rpptx.Length)
    assert paragraph.space_after == rpptx.Pt(3)

    for name in names:
        setattr(paragraph, name, None)
    presentation.save(output)
    assert "<a:pPr/>" in _text_body_xml(output)


def test_a_deck_with_line_spacing_above_two_lines_opens_and_reads_back(tmp_path):
    import rpptx

    source = tmp_path / "source.pptx"
    spaced = tmp_path / "spaced.pptx"
    _textbox_presentation(rpptx).save(source)
    _replace_in_slide(
        source,
        spaced,
        "<a:p><a:r>",
        '<a:p><a:pPr><a:lnSpc><a:spcPct val="250000"/></a:lnSpc>'
        '<a:spcBef><a:spcPct val="13200000"/></a:spcBef></a:pPr><a:r>',
    )

    paragraph = rpptx.Presentation(spaced).slides[0].shapes[0].text_frame.paragraphs[0]
    assert paragraph.line_spacing == 2.5
    assert paragraph.space_before == 132.0


def test_run_font_setters_round_trip_and_none_clears_each_attribute(tmp_path):
    import rpptx
    from rpptx import MSO_UNDERLINE, RGBColor

    presentation = _textbox_presentation(rpptx)
    font = presentation.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
    names = ("italic", "underline", "strike", "all_caps", "name", "color", "size", "bold")
    assert [font.color.rgb if name == "color" else getattr(font, name) for name in names] == [None] * len(names)

    font.italic = True
    font.underline = True
    font.strike = True
    font.all_caps = True
    font.name = "Georgia"
    font.color = RGBColor(0x12, 0x34, 0x56)
    font.size = rpptx.Pt(20)
    font.bold = False
    output = tmp_path / "font.pptx"
    presentation.save(output)

    assert (
        '<a:rPr sz="2000" b="0" i="1" cap="all" u="sng" strike="sngStrike">'
        '<a:solidFill><a:srgbClr val="123456"/></a:solidFill>'
        '<a:latin typeface="Georgia"/></a:rPr>'
    ) in _text_body_xml(output)
    presentation = rpptx.Presentation(output)
    font = presentation.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
    assert [font.color.rgb if name == "color" else getattr(font, name) for name in names] == [
        True,
        True,
        True,
        True,
        "Georgia",
        RGBColor(0x12, 0x34, 0x56),
        rpptx.Pt(20),
        False,
    ]

    for value, expected, written in (
        (MSO_UNDERLINE.DOTTED_LINE, MSO_UNDERLINE.DOTTED_LINE, 'u="dotted"'),
        (MSO_UNDERLINE.WORDS, MSO_UNDERLINE.WORDS, 'u="words"'),
        (MSO_UNDERLINE.SINGLE_LINE, True, 'u="sng"'),
        (MSO_UNDERLINE.NONE, False, 'u="none"'),
        (False, False, 'u="none"'),
    ):
        font.underline = value
        assert font.underline is expected
        presentation.save(output)
        assert written in _text_body_xml(output)

    for name in names:
        setattr(font, name, None)
    presentation.save(output)
    assert "<a:r><a:rPr/><a:t>hello</a:t></a:r>" in _text_body_xml(output)


def test_font_color_takes_triples_or_hex_and_keeps_transforms_it_does_not_name(tmp_path):
    import rpptx

    source = tmp_path / "source.pptx"
    colored = tmp_path / "colored.pptx"
    output = tmp_path / "output.pptx"
    presentation = _textbox_presentation(rpptx)
    presentation.slides[0].shapes[0].text_frame.paragraphs[0].add_run(" theme")
    presentation.save(source)
    _replace_in_slide(
        source,
        colored,
        "<a:r><a:t>hello</a:t></a:r>",
        '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/>'
        "</a:srgbClr></a:solidFill></a:rPr><a:t>hello</a:t></a:r>",
    )
    _replace_in_slide(
        colored,
        colored,
        "<a:r><a:t",
        '<a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/>'
        "</a:schemeClr></a:solidFill></a:rPr><a:t",
    )

    presentation = rpptx.Presentation(colored)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    alpha, theme = paragraph.runs[0].font, paragraph.runs[1].font
    assert str(alpha.color.rgb) == "FF0000"
    assert theme.color.rgb is None
    alpha.color = "00ff00"
    theme.color = (1, 2, 3)
    assert (str(alpha.color.rgb), str(theme.color.rgb)) == ("00FF00", "010203")
    assert paragraph.font.color.rgb is None
    paragraph.font.color = rpptx.RGBColor.from_string("ABCDEF")
    assert str(paragraph.font.color.rgb) == "ABCDEF"
    presentation.save(output)

    body = _text_body_xml(output)
    assert '<a:srgbClr val="00FF00"><a:alpha val="50000"/></a:srgbClr>' in body
    assert '<a:solidFill><a:srgbClr val="010203"/></a:solidFill>' in body
    assert "schemeClr" not in body
    assert '<a:defRPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:defRPr>' in body

    alpha.color = None
    presentation.save(output)
    assert "00FF00" not in _text_body_xml(output)


def test_fx206_small_caps_checked_modes_noop_and_same_owner_reopen(tmp_path):
    import rpptx

    presentation = _textbox_presentation(rpptx)
    owner = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    font = owner.runs[0].font
    first = tmp_path / "caps-first.pptx"
    second = tmp_path / "caps-second.pptx"
    presentation.save(first)
    before = _text_body_xml(first)
    assert font.small_caps is None
    assert font.all_caps is None
    font.small_caps = None
    presentation.save(second)
    assert _text_body_xml(second) == before
    for name, value, token in [("small_caps", True, "small"), ("all_caps", True, "all"), ("small_caps", False, "none"), ("small_caps", True, "small")]:
        setattr(font, name, value)
        assert owner.text == "hello", "formatting preserves the actual owner and logical source text"
        presentation.save(first)
        xml = _text_body_xml(first)
        assert xml.count("cap=") == 1 and f'cap="{token}"' in xml
        reopened = rpptx.Presentation(first)
        fresh = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
        assert fresh.small_caps is (True if token == "small" else False if token == "none" else None)
        assert fresh.all_caps is (True if token == "all" else False if token == "none" else None)
        reopened.save(second)
        assert _text_body_xml(second) == xml


def test_font_strike_caps_and_name_keep_the_variants_they_cannot_name(tmp_path):
    import rpptx

    source = tmp_path / "source.pptx"
    variants = tmp_path / "variants.pptx"
    output = tmp_path / "output.pptx"
    _textbox_presentation(rpptx).save(source)
    _replace_in_slide(
        source,
        variants,
        "<a:r><a:t>",
        '<a:r><a:rPr cap="small" strike="dblStrike"><a:latin typeface="Aptos" '
        'panose="020B0004020202020204" charset="0"/></a:rPr><a:t>',
    )

    presentation = rpptx.Presentation(variants)
    font = presentation.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
    assert font.strike is True
    assert font.all_caps is None
    font.strike = font.strike
    font.all_caps = None
    font.name = "Georgia"
    presentation.save(output)
    body = _text_body_xml(output)
    assert 'strike="dblStrike"' in body
    assert 'cap="small"' in body
    assert '<a:latin typeface="Georgia" panose="020B0004020202020204" charset="0"/>' in body

    font.strike = False
    font.all_caps = True
    presentation.save(output)
    body = _text_body_xml(output)
    assert 'strike="noStrike"' in body
    assert body.count("cap=") == 1 and 'cap="all"' in body
    reopened = rpptx.Presentation(output).slides[0].shapes[0].text_frame.paragraphs[0]
    assert reopened.runs[0].font.all_caps is True


def test_paragraph_bullet_reads_each_choice_and_a_character_replaces_a_picture_bullet(tmp_path):
    import rpptx

    presentation = _textbox_presentation(rpptx)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    assert paragraph.bullet is None
    paragraph.bullet = "•"
    assert paragraph.bullet == "•"
    paragraph.bullet = False
    assert paragraph.bullet is False
    source = tmp_path / "source.pptx"
    presentation.save(source)
    assert '<a:pPr><a:buNone/></a:pPr>' in _text_body_xml(source)
    paragraph.bullet = None
    assert paragraph.bullet is None

    styled = tmp_path / "styled.pptx"
    numbered = tmp_path / "numbered.pptx"
    picture = tmp_path / "picture.pptx"
    output = tmp_path / "output.pptx"
    _replace_in_slide(
        source,
        styled,
        "<a:buNone/>",
        '<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="80000"/>'
        '<a:buFont typeface="Wingdings"/><a:buChar char="§"/>',
    )
    _replace_in_slide(source, numbered, "<a:buNone/>", '<a:buAutoNum type="arabicPeriod"/>')
    _replace_in_slide(
        source,
        picture,
        "<a:buNone/>",
        '<a:buFontTx/><a:buBlip><a:blip r:embed="rId9"/></a:buBlip>',
    )

    presentation = rpptx.Presentation(styled)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    assert paragraph.bullet == "§"
    paragraph.bullet = "-"
    presentation.save(output)
    assert (
        '<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="80000"/>'
        '<a:buFont typeface="Wingdings"/><a:buChar char="-"/>'
    ) in _text_body_xml(output)
    paragraph.bullet = None
    presentation.save(output)
    assert "<a:bu" not in _text_body_xml(output)

    assert rpptx.Presentation(numbered).slides[0].shapes[0].text_frame.paragraphs[0].bullet is True
    presentation = rpptx.Presentation(picture)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    assert paragraph.bullet is True
    paragraph.bullet = "-"
    presentation.save(output)
    body = _text_body_xml(output)
    assert "buBlip" not in body
    assert '<a:buFontTx/><a:buChar char="-"/>' in body


def test_add_run_returns_a_live_run_and_keeps_the_paragraph_handle():
    import rpptx

    presentation = _textbox_presentation(rpptx)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    run = paragraph.add_run(" world")
    assert run.text == " world"
    run.font.bold = True
    assert run.font.bold is True
    assert paragraph.text == "hello world"

    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    assert paragraph.text == "hello world"
    assert [current.font.bold for current in paragraph.runs] == [None, True]
    empty = paragraph.add_run()
    assert empty.text == ""
    assert len(presentation.slides[0].shapes[0].text_frame.paragraphs[0].runs) == 3


def test_invalid_text_property_values_raise_before_changing_the_package():
    import rpptx

    presentation = _textbox_presentation(rpptx)
    frame = presentation.slides[0].shapes[0].text_frame
    paragraph = frame.paragraphs[0]
    font = paragraph.runs[0].font
    before = presentation.to_bytes()
    for error, operation in (
        # 1800 EMU is 14 hundredths of a point, which the schema rejects.
        (ValueError, lambda: setattr(font, "size", 1800)),
        (ValueError, lambda: setattr(font, "size", rpptx.Pt(4001))),
        (ValueError, lambda: setattr(font, "color", "GG0000")),
        (ValueError, lambda: setattr(font, "color", "FF00")),
        (ValueError, lambda: setattr(font, "color", (256, 0, 0))),
        (TypeError, lambda: setattr(font, "color", 3.5)),
        (ValueError, lambda: setattr(font, "name", "")),
        (ValueError, lambda: setattr(font, "underline", 18)),
        (ValueError, lambda: setattr(paragraph, "alignment", 0)),
        (TypeError, lambda: setattr(paragraph, "alignment", "ctr")),
        (ValueError, lambda: setattr(paragraph, "line_spacing", 132.5)),
        (ValueError, lambda: setattr(paragraph, "line_spacing", -1)),
        (ValueError, lambda: setattr(paragraph, "space_before", rpptx.Pt(1585))),
        (ValueError, lambda: setattr(paragraph, "space_after", -0.5)),
        (ValueError, lambda: setattr(paragraph, "left_indent", -1)),
        (ValueError, lambda: setattr(paragraph, "first_line_indent", 51_206_401)),
        (ValueError, lambda: setattr(paragraph, "bullet", True)),
        (ValueError, lambda: setattr(paragraph, "bullet", "")),
        (ValueError, lambda: setattr(frame, "margin_left", 2**31)),
        (ValueError, lambda: setattr(frame, "vertical_anchor", 2)),
        (ValueError, lambda: setattr(frame, "auto_size", 3)),
    ):
        with pytest.raises(error):
            operation()
        assert presentation.to_bytes() == before


def test_clearing_absent_text_properties_inserts_no_empty_property_elements():
    import rpptx

    presentation = _textbox_presentation(rpptx)
    frame = presentation.slides[0].shapes[0].text_frame
    paragraph = frame.paragraphs[0]
    before = presentation.to_bytes()
    for target, names in (
        (frame, ("margin_left", "vertical_anchor", "word_wrap", "auto_size")),
        (paragraph, ("alignment", "line_spacing", "space_after", "left_indent", "bullet")),
        (paragraph.font, ("bold", "italic", "underline", "name", "color", "size")),
        (paragraph.runs[0].font, ("bold", "strike", "all_caps", "name", "color", "size")),
    ):
        for name in names:
            setattr(target, name, None)
    assert presentation.to_bytes() == before


def test_text_properties_agree_with_python_pptx_in_both_directions(tmp_path):
    if importlib.util.find_spec("pptx") is None:
        pytest.skip("python-pptx oracle is installed only for the differential gate")
    assert importlib.metadata.version("python-pptx") == "1.0.2"

    import rpptx
    from pptx import Presentation as OraclePresentation
    from pptx.dml.color import RGBColor as OracleRGBColor
    from pptx.enum.text import MSO_ANCHOR as OracleAnchor
    from pptx.enum.text import MSO_AUTO_SIZE as OracleAutoSize
    from pptx.enum.text import MSO_UNDERLINE as OracleUnderline
    from pptx.enum.text import PP_ALIGN as OracleAlign
    from pptx.util import Inches as OracleInches
    from pptx.util import Pt as OraclePt

    oracle_path = tmp_path / "python-pptx.pptx"
    oracle = OraclePresentation()
    slide = oracle.slides.add_slide(oracle.slide_layouts[6])
    frame = slide.shapes.add_textbox(
        OracleInches(1), OracleInches(1), OracleInches(4), OracleInches(2)
    ).text_frame
    frame.margin_left = OracleInches(0.3)
    frame.margin_bottom = OraclePt(2)
    frame.vertical_anchor = OracleAnchor.MIDDLE
    frame.word_wrap = True
    frame.auto_size = OracleAutoSize.TEXT_TO_FIT_SHAPE
    paragraph = frame.paragraphs[0]
    paragraph.alignment = OracleAlign.RIGHT
    paragraph.line_spacing = 2.5
    paragraph.space_before = OraclePt(12)
    paragraph.space_after = OraclePt(3)
    run = paragraph.add_run()
    run.text = "oracle"
    run.font.italic = True
    run.font.underline = OracleUnderline.DOUBLE_LINE
    run.font.name = "Georgia"
    run.font.color.rgb = OracleRGBColor(0x11, 0x22, 0x33)
    run.font.size = OraclePt(24)
    paragraph.add_run().font.underline = True
    oracle.save(oracle_path)

    frame = rpptx.Presentation(oracle_path).slides[0].shapes[0].text_frame
    assert (frame.margin_left, frame.margin_right, frame.margin_bottom) == (
        rpptx.Inches(0.3),
        None,
        rpptx.Pt(2),
    )
    assert frame.vertical_anchor is rpptx.MSO_ANCHOR.MIDDLE
    assert frame.word_wrap is True
    assert frame.auto_size is rpptx.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    paragraph = frame.paragraphs[0]
    assert paragraph.alignment is rpptx.PP_ALIGN.RIGHT
    assert paragraph.line_spacing == 2.5
    assert (paragraph.space_before, paragraph.space_after) == (rpptx.Pt(12), rpptx.Pt(3))
    font = paragraph.runs[0].font
    assert (font.italic, font.underline, font.name, str(font.color.rgb), font.size) == (
        True,
        rpptx.MSO_UNDERLINE.DOUBLE_LINE,
        "Georgia",
        "112233",
        rpptx.Pt(24),
    )
    assert paragraph.runs[1].font.underline is True

    rpptx_path = tmp_path / "rpptx.pptx"
    presentation = _textbox_presentation(rpptx)
    frame = presentation.slides[0].shapes[0].text_frame
    frame.margin_top = rpptx.Pt(5)
    frame.vertical_anchor = rpptx.MSO_ANCHOR.BOTTOM
    frame.word_wrap = False
    frame.auto_size = rpptx.MSO_AUTO_SIZE.NONE
    paragraph = frame.paragraphs[0]
    paragraph.alignment = rpptx.PP_ALIGN.CENTER
    paragraph.line_spacing = rpptx.Pt(20)
    paragraph.space_before = rpptx.Pt(4)
    font = paragraph.runs[0].font
    font.bold = True
    font.italic = False
    font.underline = rpptx.MSO_UNDERLINE.WAVY_LINE
    font.name = "Arial"
    font.color = rpptx.RGBColor(0xAA, 0xBB, 0xCC)
    font.size = rpptx.Pt(28)
    presentation.save(rpptx_path)

    frame = OraclePresentation(rpptx_path).slides[0].shapes[0].text_frame
    assert frame.margin_top == OraclePt(5)
    assert frame.vertical_anchor == OracleAnchor.BOTTOM
    assert frame.word_wrap is False
    assert frame.auto_size == OracleAutoSize.NONE
    paragraph = frame.paragraphs[0]
    assert paragraph.alignment == OracleAlign.CENTER
    assert paragraph.line_spacing == OraclePt(20)
    assert paragraph.space_before == OraclePt(4)
    font = paragraph.runs[0].font
    assert (font.bold, font.italic, font.underline, font.name, font.size) == (
        True,
        False,
        OracleUnderline.WAVY_LINE,
        "Arial",
        OraclePt(28),
    )
    assert font.color.rgb == OracleRGBColor(0xAA, 0xBB, 0xCC)


def _slide_hyperlink_targets(data):
    import xml.etree.ElementTree as ElementTree

    relationships = _package_parts(data)["ppt/slides/_rels/slide1.xml.rels"]
    return sorted(
        relationship.get("Target")
        for relationship in ElementTree.fromstring(relationships)
        if relationship.get("Type").endswith("/hyperlink")
    )


def test_run_hyperlink_address_reads_writes_and_prunes_like_python_pptx(tmp_path):
    import rpptx

    prs = _textbox_presentation(rpptx)
    prs.slides[0].shapes[0].text_frame.paragraphs[0].add_run(" world")
    first, second = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs
    link = first.hyperlink
    assert link.address is None
    before = prs.to_bytes()
    link.address = None
    link.address = ""
    assert prs.to_bytes() == before
    link.address = "https://example.com/a?x=1&y=2"
    second.hyperlink.address = "https://example.com/a?x=1&y=2"
    assert (first.hyperlink.address, second.text) == ("https://example.com/a?x=1&y=2", " world")
    assert _slide_hyperlink_targets(prs.to_bytes()) == ["https://example.com/a?x=1&y=2"]
    for _ in range(3):
        link.address = "https://example.com/b"
        link.address = "https://example.com/c"
    assert _slide_hyperlink_targets(prs.to_bytes()) == [
        "https://example.com/a?x=1&y=2",
        "https://example.com/c",
    ]
    second.hyperlink.address = None
    assert (second.hyperlink.address, second.font.name) == (None, None)
    assert _slide_hyperlink_targets(prs.to_bytes()) == ["https://example.com/c"]
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="control characters"):
        link.address = "https://example.com/\nnext"
    assert prs.to_bytes() == before
    assert b"https://example.com/c" in prs.to_pdf()
    held = first.hyperlink
    prs.slides.add_slide(prs.slide_layouts[6])
    assert held.address == first.hyperlink.address
    output = tmp_path / "hyperlinks.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    runs = pptx.Presentation(output).slides[0].shapes[0].text_frame.paragraphs[0].runs
    assert [run.hyperlink.address for run in runs] == ["https://example.com/c", None]

    def build(deck):
        frame = deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_textbox(0, 0, 100, 100)
        paragraph = frame.text_frame.paragraphs[0]
        for text in ("one", "two"):
            run = paragraph.add_run()
            run.text = text
            run.hyperlink.address = f"https://example.com/{text}"

    source = _python_pptx_deck(tmp_path / "python-pptx-links.pptx", build)
    prs = rpptx.Presentation(source)
    runs = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs
    assert [run.hyperlink.address for run in runs] == [
        "https://example.com/one",
        "https://example.com/two",
    ]
    runs[0].hyperlink.address = "https://example.com/two"
    runs[1].hyperlink.address = None
    retargeted = tmp_path / "python-pptx-links-out.pptx"
    prs.save(retargeted)
    assert _slide_hyperlink_targets(retargeted.read_bytes()) == ["https://example.com/two"]
    runs = pptx.Presentation(retargeted).slides[0].shapes[0].text_frame.paragraphs[0].runs
    assert [run.hyperlink.address for run in runs] == ["https://example.com/two", None]


def _slide_jump_targets(data):
    import xml.etree.ElementTree as ElementTree

    relationships = _package_parts(data)["ppt/slides/_rels/slide1.xml.rels"]
    return sorted(
        relationship.get("Target")
        for relationship in ElementTree.fromstring(relationships)
        if relationship.get("Type").endswith("/slide")
    )


def test_shape_click_action_links_and_jumps_like_python_pptx(tmp_path):
    import rpptx
    from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE

    prs = _textbox_presentation(rpptx)
    for _ in range(2):
        prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, 100, 100)
    prs.slides[0].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 100, 100)
    prs.slides[0].shapes.add_group_shape()
    shapes = prs.slides[0].shapes
    box, rect, line, group = shapes
    action = box.click_action
    assert (action.hyperlink.address, action.target_slide) == (None, None)
    before = prs.to_bytes()
    action.hyperlink.address = None
    action.hyperlink.address = ""
    action.target_slide = None
    assert prs.to_bytes() == before

    address = "https://example.com/a?x=1&y=2"
    for shape in (box, rect, line, group):
        shape.click_action.hyperlink.address = address
    assert [shape.click_action.hyperlink.address for shape in shapes] == [address] * 4
    assert _slide_hyperlink_targets(prs.to_bytes()) == [address]
    rect.click_action.target_slide = prs.slides[2]
    line.click_action.target_slide = prs.slides[2]
    assert rect.click_action.target_slide == prs.slides[2]
    assert rect.click_action.target_slide != prs.slides[1]
    assert rect.click_action.hyperlink.address == "slide3.xml"
    assert _slide_jump_targets(prs.to_bytes()) == ["slide3.xml"]
    for shape in (box, group):
        shape.click_action.hyperlink.address = None
    assert _slide_hyperlink_targets(prs.to_bytes()) == []
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="control characters"):
        rect.click_action.hyperlink.address = "https://example.com/\nnext"
    other = rpptx.Presentation()
    other.slides.add_slide(other.slide_layouts[6])
    with pytest.raises(ValueError, match="not in this presentation"):
        rect.click_action.target_slide = other.slides[0]
    assert prs.to_bytes() == before
    held = rect.click_action
    prs.slides.add_slide(prs.slide_layouts[6])
    assert held.target_slide == rect.click_action.target_slide
    output = tmp_path / "click-actions.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    deck = pptx.Presentation(output)
    first, second, third = list(deck.slides[0].shapes)[:3]
    assert [shape.click_action.hyperlink.address for shape in (first, second, third)] == [
        None,
        "slide3.xml",
        "slide3.xml",
    ]
    assert second.click_action.target_slide == deck.slides[2]
    assert first.click_action.target_slide is None

    def build(deck):
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        target = deck.slides.add_slide(deck.slide_layouts[6])
        slide.shapes.add_textbox(0, 0, 100, 100).click_action.hyperlink.address = (
            "https://example.com/one"
        )
        slide.shapes.add_textbox(0, 0, 100, 100).click_action.target_slide = target

    source = _python_pptx_deck(tmp_path / "python-pptx-click.pptx", build)
    prs = rpptx.Presentation(source)
    first, second = prs.slides[0].shapes
    assert first.click_action.hyperlink.address == "https://example.com/one"
    assert first.click_action.target_slide is None
    assert second.click_action.target_slide == prs.slides[1]
    first.click_action.target_slide = prs.slides[1]
    second.click_action.hyperlink.address = "https://example.com/two"
    retargeted = tmp_path / "python-pptx-click-out.pptx"
    prs.save(retargeted)
    assert _slide_hyperlink_targets(retargeted.read_bytes()) == ["https://example.com/two"]
    assert _slide_jump_targets(retargeted.read_bytes()) == ["slide2.xml"]
    deck = pptx.Presentation(retargeted)
    first, second = deck.slides[0].shapes
    assert first.click_action.target_slide == deck.slides[1]
    assert second.click_action.hyperlink.address == "https://example.com/two"

    prs.slides.remove(prs.slides[1])
    first, second = prs.slides[0].shapes
    assert (first.click_action.target_slide, first.click_action.hyperlink.address) == (None, None)
    assert second.click_action.hyperlink.address == "https://example.com/two"
    removed = tmp_path / "python-pptx-click-removed.pptx"
    prs.save(removed)
    assert _slide_jump_targets(removed.read_bytes()) == []
    first = pptx.Presentation(removed).slides[0].shapes[0]
    assert first.click_action.action == pptx.enum.action.PP_ACTION.NONE


def test_text_enums_match_python_pptx_member_values_and_xml_tokens():
    if importlib.util.find_spec("pptx") is None:
        pytest.skip("python-pptx oracle is installed only for the differential gate")
    assert importlib.metadata.version("python-pptx") == "1.0.2"

    import rpptx
    from pptx.enum import text as oracle
    from rpptx.enum import text

    presentation = _textbox_presentation(rpptx)

    def written(attribute):
        body = presentation.to_bytes()
        with zipfile.ZipFile(io.BytesIO(body)) as archive:
            slide_xml = archive.read("ppt/slides/slide1.xml").decode()
        return re.search(f' {attribute}="([^"]+)"', slide_xml).group(1)

    for ours, theirs, attribute, assign in (
        (
            text.MSO_TEXT_UNDERLINE_TYPE,
            oracle.MSO_TEXT_UNDERLINE_TYPE,
            "u",
            lambda value: setattr(
                presentation.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font,
                "underline",
                value,
            ),
        ),
        (
            text.PP_PARAGRAPH_ALIGNMENT,
            oracle.PP_PARAGRAPH_ALIGNMENT,
            "algn",
            lambda value: setattr(
                presentation.slides[0].shapes[0].text_frame.paragraphs[0], "alignment", value
            ),
        ),
        (
            text.MSO_VERTICAL_ANCHOR,
            oracle.MSO_VERTICAL_ANCHOR,
            "anchor",
            lambda value: setattr(
                presentation.slides[0].shapes[0].text_frame, "vertical_anchor", value
            ),
        ),
        (text.MSO_AUTO_SIZE, oracle.MSO_AUTO_SIZE, None, None),
    ):
        members = {member.name: member for member in theirs if member.name != "MIXED"}
        extensions = {"JUSTIFY", "DISTRIBUTE"} if ours is text.MSO_VERTICAL_ANCHOR else set()
        assert {member.name for member in ours} == set(members) | extensions
        for name, member in members.items():
            assert int(ours[name]) == int(member)
            if assign is not None:
                assign(ours[name])
                assert written(attribute) == member.xml_value
def _python_pptx_deck(path, build):
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    deck = pptx.Presentation()
    build(deck)
    deck.save(path)
    return path


def _package_parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _package_bytes(parts):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


def _template_jpeg():
    import rpptx

    return _package_parts(rpptx.Presentation().to_bytes())["docProps/thumbnail.jpeg"]


def test_slide_size_reads_and_writes_keep_the_other_dimension(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    assert (prs.slide_width, prs.slide_height) == (12_192_000, 6_858_000)
    assert isinstance(prs.slide_width, rpptx.Length)
    held = prs.slides
    prs.slide_width = rpptx.Inches(10)
    assert (prs.slide_width, prs.slide_height) == (rpptx.Inches(10), 6_858_000)
    assert len(held) == 0
    before = prs.to_bytes()
    with pytest.raises(ValueError, match=r"^slide_height must be between 1 and 56 inches$"):
        prs.slide_height = 0
    assert prs.to_bytes() == before
    path = tmp_path / "sized.pptx"
    prs.save(path)
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(path)
    assert (oracle.slide_width, oracle.slide_height) == (rpptx.Inches(10), 6_858_000)

    parts = _package_parts(prs.to_bytes())
    parts["ppt/presentation.xml"] = re.sub(
        rb"<p:sldSz [^>]*/>", b"", parts["ppt/presentation.xml"]
    )
    unsized = tmp_path / "unsized.pptx"
    unsized.write_bytes(_package_bytes(parts))
    prs = rpptx.Presentation(unsized)
    assert (prs.slide_width, prs.slide_height) == (None, None)
    prs.slide_height = rpptx.Inches(5)
    assert (prs.slide_width, prs.slide_height) == (12_192_000, rpptx.Inches(5))


def test_slide_layout_hidden_and_background_round_trip_through_python_pptx(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor
    from rpptx.enum.dml import MSO_FILL_TYPE

    source = _python_pptx_deck(
        tmp_path / "slides.pptx",
        lambda deck: [deck.slides.add_slide(deck.slide_layouts[index]) for index in (5, 1)],
    )
    prs = rpptx.Presentation(source)
    first, second = prs.slides[0], prs.slides[1]
    assert first.slide_layout == prs.slide_layouts[5]
    assert first.slide_layout != prs.slide_layouts[1]
    assert first.slide_layout.name == "Title Only"
    assert prs.slide_layouts.index(second.slide_layout) == 1
    with pytest.raises(ValueError, match="layout not in this SlideLayouts collection"):
        prs.slide_layouts.index(rpptx.Presentation().slide_layouts[1])

    assert first.hidden is False
    first.hidden = True
    assert prs.slides[0].hidden is True

    background = first.background
    assert first.follow_master_background is True
    assert background.fill.type is None
    assert first.follow_master_background is True
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)
    assert background.fill.type == MSO_FILL_TYPE.SOLID
    assert first.follow_master_background is False
    inherited_paint = prs.render_slide_to_png(1, dpi=72)
    second.follow_master_background = False
    # Detaching preserves the actual inherited reference, not a fabricated no-fill.
    assert second.background.fill.type is None
    assert prs.render_slide_to_png(1, dpi=72) == inherited_paint
    output = tmp_path / "slides-out.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output)
    assert oracle.slides[0]._element.get("show") == "0"
    assert oracle.slides[0].follow_master_background is False
    assert oracle.slides[0].background.fill.fore_color.rgb == pptx.dml.color.RGBColor(0x12, 0x34, 0x56)
    assert oracle.slides[1].follow_master_background is False
    assert oracle.slides[1]._element.xpath("./p:cSld/p:bg/p:bgRef/@idx") == ["1001"]
    assert oracle.slides[1]._element.xpath("./p:cSld/p:bg/p:bgRef/a:schemeClr/@val") == ["bg1"]

    prs.slides[0].follow_master_background = True
    prs.slides[0].hidden = False
    prs.save(output)
    oracle = pptx.Presentation(output)
    assert oracle.slides[0].follow_master_background is True
    assert oracle.slides[0]._element.get("show") in (None, "1")


def test_slide_layout_assignment_retargets_the_slide_and_keeps_unplaced_placeholders(tmp_path):
    import rpptx

    source = _python_pptx_deck(
        tmp_path / "layouts.pptx", lambda deck: deck.slides.add_slide(deck.slide_layouts[1])
    )
    prs = rpptx.Presentation(source)
    for index, text in enumerate(("Title", "Body")):
        prs.slides[0].shapes[index].text = text
    slide = prs.slides[0]
    body_geometry = slide.shapes[1].effective_geometry()
    slide.slide_layout = prs.slide_layouts[5]
    assert slide.slide_layout == prs.slide_layouts[5]
    slide = prs.slides[0]
    assert slide.slide_layout == prs.slide_layouts[5]
    title, body = slide.shapes
    assert title.left is None
    assert (body.left, body.top, body.width, body.height) == body_geometry
    assert [frame.shape_id for frame in prs.text_layout()] == [title.shape_id, body.shape_id]

    prs.slides[0].slide_layout = prs.slide_layouts[0]
    assert prs.slides[0].shapes[0].effective_geometry() == (685800, 2130425, 7772400, 1470025)
    with pytest.raises(ValueError, match="slide layout is not in this presentation"):
        prs.slides[0].slide_layout = rpptx.Presentation().slide_layouts[1]
    with pytest.raises(TypeError):
        prs.slides[0].slide_layout = 1
    output = tmp_path / "relaid.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0]
    assert oracle.slide_layout.name == "Title Slide"
    title, body = oracle.shapes
    assert (title.left, title.top, title.width, title.height) == (685800, 2130425, 7772400, 1470025)
    assert (body.left, body.top, body.width, body.height) == body_geometry


def test_shape_geometry_name_and_rotation_setters_match_python_pptx(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(1, 2, 3, 4)
    held = shape
    shape.left, shape.top = rpptx.Inches(1), rpptx.Inches(2)
    shape.width, shape.height = rpptx.Inches(3), rpptx.Inches(4)
    shape.name = "Renamed box"
    shape.rotation = 45.5
    assert (held.left, held.top, held.width, held.height) == (
        rpptx.Inches(1),
        rpptx.Inches(2),
        rpptx.Inches(3),
        rpptx.Inches(4),
    )
    assert (held.name, held.rotation) == ("Renamed box", 45.5)
    for value, expected in ((-90, 270.0), (360, 0.0), (720.25, 0.25)):
        shape.rotation = value
        assert shape.rotation == expected
    shape.rotation = 30
    with pytest.raises(ValueError):
        shape.width = -1
    with pytest.raises(ValueError):
        shape.rotation = float("nan")

    group = prs.slides[0].shapes.add_group_shape()
    assert (group.left, group.width, group.rotation) == (None, None, 0.0)
    group.width = 100
    assert (group.left, group.width, group.height) == (None, 100, 0)
    output = tmp_path / "geometry.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes[0]
    assert (oracle.left, oracle.top, oracle.width, oracle.height) == (
        rpptx.Inches(1),
        rpptx.Inches(2),
        rpptx.Inches(3),
        rpptx.Inches(4),
    )
    assert (oracle.name, oracle.rotation) == ("Renamed box", 30.0)


def test_placeholder_effective_geometry_matches_python_pptx_and_one_setter_keeps_the_rest(tmp_path):
    import rpptx

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    source = _python_pptx_deck(
        tmp_path / "placeholders.pptx",
        lambda deck: [deck.slides.add_slide(deck.slide_layouts[index]) for index in (0, 1, 5, 8)],
    )
    oracle = pptx.Presentation(source)
    prs = rpptx.Presentation(source)
    assert len(prs.slides) == len(oracle.slides)
    for slide, oracle_slide in zip(prs.slides, oracle.slides):
        assert len(slide.shapes) == len(oracle_slide.shapes)
        for shape, expected in zip(slide.shapes, oracle_slide.shapes):
            assert (shape.left, shape.top, shape.width, shape.height) == (None,) * 4
            geometry = shape.effective_geometry()
            assert geometry == (expected.left, expected.top, expected.width, expected.height)
            assert all(isinstance(value, rpptx.Length) for value in geometry)
    assert prs.slides[2].shapes[0].effective_geometry() == (457200, 274638, 8229600, 1143000)

    prs.slides[0].shapes[0].text = "Rotated title"
    prs.slides[2].shapes[0].text = "Moved title"
    title = prs.slides[2].shapes[0]
    title.left = title.effective_geometry()[0] + rpptx.Inches(0.1)
    moved = (457200 + 91440, 274638, 8229600, 1143000)
    assert (title.left, title.top, title.width, title.height) == moved
    assert title.effective_geometry() == moved
    body = prs.slides[1].shapes[1]
    body.height = 1000000
    assert (body.left, body.top, body.width, body.height) == (457200, 1600200, 8229600, 1000000)
    rotated = prs.slides[0].shapes[0]
    rotated_geometry = rotated.effective_geometry()
    rotated.rotation = 30.0
    assert rotated.rotation == 30.0
    assert rotated.effective_geometry() == rotated_geometry
    assert (rotated.left, rotated.top, rotated.width, rotated.height) == rotated_geometry
    assert [(frame.slide_index, frame.shape_id) for frame in prs.text_layout()] == [
        (0, rotated.shape_id),
        (2, title.shape_id),
    ]
    output = tmp_path / "moved.pptx"
    prs.save(output)
    reread = pptx.Presentation(output)
    oracle_title = reread.slides[2].shapes[0]
    assert (oracle_title.left, oracle_title.top, oracle_title.width, oracle_title.height) == moved
    oracle_body = reread.slides[1].shapes[1]
    assert (oracle_body.left, oracle_body.width, oracle_body.height) == (457200, 8229600, 1000000)
    oracle_rotated = reread.slides[0].shapes[0]
    assert (
        oracle_rotated.left,
        oracle_rotated.top,
        oracle_rotated.width,
        oracle_rotated.height,
    ) == rotated_geometry
    assert oracle_rotated.rotation == 30.0


def test_shape_type_reports_the_python_pptx_member_for_every_shape_kind(tmp_path):
    import rpptx
    from rpptx.enum.shapes import MSO_SHAPE_TYPE

    def build(deck):
        pptx = pytest.importorskip("pptx")
        slide = deck.slides.add_slide(deck.slide_layouts[1])
        slide.shapes.add_textbox(0, 0, 10, 10)
        slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, 0, 0, 10, 10)
        slide.shapes.add_picture(io.BytesIO(_tiny_png()), 0, 0)
        slide.shapes.add_table(1, 1, 0, 0, 10, 10)
        slide.shapes.add_connector(pptx.enum.shapes.MSO_CONNECTOR.ELBOW, 0, 0, 10, 10)
        slide.shapes.add_group_shape()
        builder = slide.shapes.build_freeform(0, 0)
        builder.add_line_segments([(10, 10), (0, 10)])
        builder.convert_to_shape()

    source = _python_pptx_deck(tmp_path / "kinds.pptx", build)
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    expected = [int(shape.shape_type) for shape in pptx.Presentation(source).slides[0].shapes]
    shapes = rpptx.Presentation(source).slides[0].shapes
    actual = [shape.shape_type for shape in shapes]
    assert actual == expected
    assert all(isinstance(value, MSO_SHAPE_TYPE) for value in actual)
    assert actual[0] == MSO_SHAPE_TYPE.PLACEHOLDER
    assert actual[-1] == MSO_SHAPE_TYPE.FREEFORM


def test_auto_shape_type_reads_like_python_pptx_and_replaces_the_preset(tmp_path):
    import rpptx
    from rpptx.enum.shapes import MSO_SHAPE

    def build(deck):
        pptx = pytest.importorskip("pptx")
        slide = deck.slides.add_slide(deck.slide_layouts[1])
        slide.shapes.add_textbox(0, 0, 10, 10)
        rounded = slide.shapes.add_shape(
            pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, 0, 0, 100, 50
        )
        rounded.text = "Keep me"
        rounded.fill.solid()
        rounded.fill.fore_color.rgb = pptx.dml.color.RGBColor(0x11, 0x22, 0x33)
        rounded.adjustments[0] = 0.3
        slide.shapes.add_picture(io.BytesIO(_tiny_png()), 0, 0)
        slide.shapes.add_table(1, 1, 0, 0, 10, 10)
        slide.shapes.add_connector(pptx.enum.shapes.MSO_CONNECTOR.ELBOW, 0, 0, 10, 10)
        slide.shapes.add_group_shape()
        builder = slide.shapes.build_freeform(0, 0)
        builder.add_line_segments([(10, 10), (0, 10)])
        builder.convert_to_shape()

    def read(shape):
        try:
            value = shape.auto_shape_type
        except (AttributeError, ValueError):
            return "not an auto shape"
        return None if value is None else (value.name, int(value))

    source = _python_pptx_deck(tmp_path / "auto-shapes.pptx", build)
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    expected = [read(shape) for shape in pptx.Presentation(source).slides[0].shapes]
    prs = rpptx.Presentation(source)
    shapes = prs.slides[0].shapes
    assert [read(shape) for shape in shapes] == expected
    assert expected[3:5] == [("ROUNDED_RECTANGLE", 5), ("RECTANGLE", 1)]
    assert shapes[3].auto_shape_type is MSO_SHAPE.ROUNDED_RECTANGLE
    with pytest.raises(ValueError, match="shape is not an auto shape"):
        _ = shapes[-1].auto_shape_type
    assert MSO_SHAPE.from_xml("wedgeRoundRectCallout") is MSO_SHAPE.BALLOON
    assert MSO_SHAPE.from_xml("upArrow") is MSO_SHAPE.UP_ARROW

    rounded = shapes[3]
    identity = (rounded.shape_id, rounded.name, rounded.left, rounded.width)
    assert list(rounded.adjustments) == [0.3]
    rounded.auto_shape_type = MSO_SHAPE.RECTANGLE
    assert rounded.auto_shape_type is MSO_SHAPE.RECTANGLE
    assert list(rounded.adjustments) == []
    rounded.auto_shape_type = MSO_SHAPE.OVAL
    shapes[-1].auto_shape_type = "roundRect"
    shapes[4].auto_shape_type = MSO_SHAPE.OVAL
    assert list(shapes[-1].adjustments) == [0.16667]
    for index in (2, 5, 6, 7):
        with pytest.raises(ValueError, match="shape is not an auto shape"):
            shapes[index].auto_shape_type = MSO_SHAPE.RECTANGLE
    rounded.auto_shape_type = 35
    assert rounded.auto_shape_type is MSO_SHAPE.UP_ARROW
    rounded.auto_shape_type = "upArrow"
    assert rounded.auto_shape_type is MSO_SHAPE.UP_ARROW
    rounded.auto_shape_type = MSO_SHAPE.OVAL

    target = tmp_path / "changed.pptx"
    prs.save(target)
    oracle = pptx.Presentation(target).slides[0].shapes
    assert [read(shape) for shape in oracle] == [
        *expected[:3],
        ("OVAL", 9),
        ("OVAL", 9),
        *expected[5:8],
        ("ROUNDED_RECTANGLE", 5),
    ]
    changed = oracle[3]
    assert (changed.shape_id, changed.name, changed.left, changed.width) == identity
    assert changed.text_frame.text == "Keep me"
    assert str(changed.fill.fore_color.rgb) == "112233"
    assert list(changed.adjustments) == []
    assert [shape.shape_id for shape in oracle] == [
        shape.shape_id for shape in pptx.Presentation(source).slides[0].shapes
    ]


def test_fill_and_line_formats_write_what_python_pptx_reads(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor
    from rpptx.enum.dml import MSO_FILL, MSO_FILL_TYPE

    assert MSO_FILL is MSO_FILL_TYPE
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape("rect", 0, 0, 100, 100)
    fill = shape.fill
    assert fill.type is None
    with pytest.raises(TypeError, match="fill type _NoneFill has no foreground color"):
        _ = fill.fore_color
    fill.solid()
    assert fill.type == MSO_FILL_TYPE.SOLID
    assert fill.fore_color.rgb is None
    fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
    fill.solid()
    assert shape.fill.fore_color.rgb == RGBColor(0xFF, 0x00, 0x00)
    assert str(shape.fill.fore_color.rgb) == "FF0000"
    with pytest.raises(TypeError, match="FillFormat.fore_color.rgb"):
        fill.fore_color.rgb = (1, 2)

    line = shape.line
    assert (line.width, line.color.rgb, line.fill.type) == (0, None, None)
    line.color.rgb = RGBColor(0x00, 0x80, 0x00)
    assert line.fill.type == MSO_FILL_TYPE.SOLID
    line.width = rpptx.Pt(2)
    assert line.width == rpptx.Pt(2)
    with pytest.raises(ValueError):
        line.width = 20_116_801
    line.width = None
    assert line.width == 0
    line.width = rpptx.Pt(2)

    plain = prs.slides[0].shapes.add_shape("ellipse", 0, 0, 10, 10)
    _ = plain.line.color.rgb, plain.line.width, plain.fill.type
    plain.fill.background()
    assert plain.fill.type == MSO_FILL_TYPE.BACKGROUND
    with pytest.raises(ValueError, match="shape has no fill"):
        _ = prs.slides[0].shapes.add_group_shape().fill
    with pytest.raises(ValueError, match="shape has no line"):
        _ = prs.slides[0].shapes.add_table(1, 1, 0, 0, 10, 10).line

    held = prs.slides[0].shapes[0].fill
    prs.slides.add_slide(prs.slide_layouts[6])
    assert held.type == prs.slides[0].shapes[0].fill.type
    output = tmp_path / "fills.pptx"
    prs.save(output)
    parts = _package_parts(output.read_bytes())
    slide_xml = parts["ppt/slides/slide1.xml"].decode()
    assert len(re.findall(r"<a:ln[ />]", slide_xml)) == 1
    assert slide_xml.count("<p:style>") == 2

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes
    assert oracle[0].fill.type == pptx.enum.dml.MSO_FILL.SOLID
    assert oracle[0].fill.fore_color.rgb == pptx.dml.color.RGBColor(0xFF, 0x00, 0x00)
    assert oracle[0].line.color.rgb == pptx.dml.color.RGBColor(0x00, 0x80, 0x00)
    assert oracle[0].line.width == rpptx.Pt(2)
    assert oracle[1].fill.type == pptx.enum.dml.MSO_FILL.BACKGROUND


def _cell_records(table, rows, columns):
    return [
        (cell.text, cell.is_merge_origin, cell.is_spanned, cell.span_height, cell.span_width)
        for cell in (table.cell(row, column) for row in range(rows) for column in range(columns))
    ]


def _python_pptx_merged_table(pptx, split):
    deck = pptx.Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    table = slide.shapes.add_table(3, 3, 0, 0, 5_486_400, 2_743_200).table
    for row in range(3):
        for column in range(3):
            table.cell(row, column).text = f"{row}{column}"
    table.cell(0, 0).merge(table.cell(1, 1))
    if split:
        table.cell(0, 0).split()
    return table


def test_table_cells_merge_split_fill_and_margins_like_python_pptx(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor
    from rpptx.enum.dml import MSO_FILL_TYPE

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_table(3, 3, 0, 0, 5_486_400, 2_743_200)
    prs.slides[0].shapes.add_table(1, 2, 0, 3_000_000, 100, 100)
    table = prs.slides[0].shapes[0].table
    for row in range(3):
        for column in range(3):
            table.cell(row, column).text = f"{row}{column}"
    origin, held = table.cell(0, 0), table.cell(2, 2)
    origin.merge(table.cell(1, 1))
    assert (origin.is_merge_origin, origin.is_spanned, origin.span_height, origin.span_width) == (
        True,
        False,
        2,
        2,
    )
    assert (table.cell(1, 0).is_merge_origin, table.cell(1, 0).is_spanned) == (False, True)
    assert (held.text, held.is_merge_origin, held.span_width) == ("22", False, 1)
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="merged cells"):
        table.cell(1, 1).merge(table.cell(2, 2))
    with pytest.raises(ValueError, match="other_cell from different table"):
        held.merge(prs.slides[0].shapes[1].table.cell(0, 0))
    with pytest.raises(rpptx.RpptxError, match="merge-origin"):
        held.split()
    assert prs.to_bytes() == before
    merged = tmp_path / "merged.pptx"
    prs.save(merged)
    origin.split()
    assert not any(record[1] or record[2] for record in _cell_records(table, 3, 3))

    cell, other = table.cell(2, 0), table.cell(2, 1)
    assert cell.fill.type is None
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)
    other.fill.background()
    assert (cell.fill.type, other.fill.type) == (MSO_FILL_TYPE.SOLID, MSO_FILL_TYPE.BACKGROUND)
    assert (cell.margin_left, cell.margin_right, cell.margin_top, cell.margin_bottom) == (None,) * 4
    cell.margin_left = rpptx.Inches(0.25)
    cell.margin_top = rpptx.Pt(3)
    assert (cell.margin_left, cell.margin_top) == (rpptx.Inches(0.25), rpptx.Pt(3))
    assert isinstance(cell.margin_left, rpptx.Length)
    before = prs.to_bytes()
    cell.margin_right = None
    with pytest.raises(ValueError, match="32-bit"):
        cell.margin_bottom = 2**31
    assert prs.to_bytes() == before
    cell.margin_top = None
    assert cell.margin_top is None
    held_fill = cell.fill
    prs.slides.add_slide(prs.slide_layouts[6])
    assert held_fill.type == cell.fill.type
    split = tmp_path / "split.pptx"
    prs.save(split)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    for path, was_split in ((merged, False), (split, True)):
        expected = _cell_records(_python_pptx_merged_table(pptx, was_split), 3, 3)
        assert _cell_records(pptx.Presentation(path).slides[0].shapes[0].table, 3, 3) == expected
        assert _cell_records(rpptx.Presentation(path).slides[0].shapes[0].table, 3, 3) == expected
    oracle = pptx.Presentation(split).slides[0].shapes[0].table
    assert oracle.cell(2, 0).fill.fore_color.rgb == pptx.dml.color.RGBColor(0x12, 0x34, 0x56)
    assert oracle.cell(2, 1).fill.type == pptx.enum.dml.MSO_FILL.BACKGROUND
    assert oracle.cell(2, 0).margin_left == rpptx.Inches(0.25)
    assert (oracle.cell(2, 0).margin_right, oracle.cell(2, 0).margin_top) == (91_440, 45_720)

    def build(deck):
        written = deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_table(1, 1, 0, 0, 100, 100)
        written = written.table.cell(0, 0)
        written.margin_bottom = rpptx.Pt(4)
        written.fill.solid()
        written.fill.fore_color.rgb = pptx.dml.color.RGBColor(0xAB, 0xCD, 0xEF)

    source = _python_pptx_deck(tmp_path / "python-pptx-cells.pptx", build)
    read = rpptx.Presentation(source).slides[0].shapes[0].table.cell(0, 0)
    assert (read.margin_left, read.margin_bottom) == (None, rpptx.Pt(4))
    assert read.fill.fore_color.rgb == RGBColor(0xAB, 0xCD, 0xEF)


def test_table_row_heights_and_cell_borders_write_what_python_pptx_reads(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor
    from rpptx.enum.dml import MSO_ARROWHEAD_STYLE, MSO_FILL_TYPE, MSO_LINE_DASH_STYLE

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_table(3, 2, 0, 0, rpptx.Inches(4), rpptx.Inches(3))
    shape = prs.slides[0].shapes[0]
    rows = shape.table.rows
    assert len(rows) == 3
    assert [row.height for row in rows] == [rpptx.Inches(1)] * 3
    assert isinstance(rows[0].height, rpptx.Length)
    held = rows[-1]
    rows[1].height = rpptx.Inches(1.5)
    assert (held.height, shape.height) == (rpptx.Inches(1), rpptx.Inches(3.5))
    assert [row.height for row in rows[1:]] == [rpptx.Inches(1.5), rpptx.Inches(1)]
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="row height must be positive"):
        rows[0].height = 0
    with pytest.raises(IndexError):
        _ = rows[3]

    cell = shape.table.cell(0, 0)
    left = cell.border_left
    assert (left.width, left.color.rgb, left.fill.type) == (0, None, None)
    assert prs.to_bytes() == before
    left.width = rpptx.Pt(2)
    left.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    left.dash_style = MSO_LINE_DASH_STYLE.DASH
    left.tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE
    assert left.dash_style is MSO_LINE_DASH_STYLE.DASH
    assert left.tail_end.type is MSO_ARROWHEAD_STYLE.TRIANGLE
    cell.border_bottom.fill.solid()
    cell.border_bottom.fill.fore_color.rgb = RGBColor(0x00, 0x80, 0x00)
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(0x11, 0x22, 0x33)
    assert (cell.border_left.width, cell.border_left.color.rgb) == (
        rpptx.Pt(2),
        RGBColor(0xFF, 0x00, 0x00),
    )
    assert (cell.border_top.fill.type, cell.border_bottom.fill.type) == (None, MSO_FILL_TYPE.SOLID)
    held_border = cell.border_right
    prs.slides.add_slide(prs.slide_layouts[6])
    assert held_border.width == cell.border_right.width
    assert held.height == prs.slides[0].shapes[0].table.rows[2].height
    output = tmp_path / "rows-borders.pptx"
    prs.save(output)
    xml = _package_parts(output.read_bytes())["ppt/slides/slide1.xml"].decode()
    properties = xml[xml.index("<a:tcPr") :]
    assert (
        properties.index("<a:lnL")
        < properties.index("<a:lnB")
        < properties.index('<a:srgbClr val="112233"/>')
    )

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    from pptx.oxml.ns import qn

    oracle = pptx.Presentation(output).slides[0].shapes[0]
    assert [row.height for row in oracle.table.rows] == [
        rpptx.Inches(1),
        rpptx.Inches(1.5),
        rpptx.Inches(1),
    ]
    assert oracle.height == rpptx.Inches(3.5)
    border = oracle.table.cell(0, 0)._tc.tcPr.find(qn("a:lnL"))
    assert border.get("w") == str(rpptx.Pt(2))
    assert border.find(qn("a:solidFill"))[0].get("val") == "FF0000"

    def build(deck):
        table = deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_table(2, 1, 0, 0, 100, 100)
        table.table.rows[0].height = rpptx.Pt(30)

    source = _python_pptx_deck(tmp_path / "python-pptx-rows.pptx", build)
    assert rpptx.Presentation(source).slides[0].shapes[0].table.rows[0].height == rpptx.Pt(30)


def _python_pptx_options_table(deck):
    """Build the slide 4 table of the rdocx-skills deck fixture."""
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    rows = (
        ("", "A. Repair", "B. Refurbish", "C. Replace deck"),
        ("Cost (EUR)", "46,000", "212,000", "690,000"),
        ("Closure", "2 weeks", "7 weeks", "5 months"),
        ("Next major work", "2031", "2040", "2055"),
    )
    slide = deck.slides.add_slide(deck.slide_layouts[5])
    slide.shapes.title.text = "Three options"
    frame = slide.shapes.add_table(4, 4, Inches(0.5), Inches(1.5), Inches(9), Inches(2.4))
    for row_index, row in enumerate(rows):
        for column_index, value in enumerate(row):
            cell = frame.table.cell(row_index, column_index)
            cell.text = value
            for run in cell.text_frame.paragraphs[0].runs:
                run.font.size = Pt(15)
            if row_index == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0x3E, 0x4A, 0x59)


def test_table_rows_and_columns_are_added_and_removed_like_python_pptx_add_tr(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor

    source = _python_pptx_deck(tmp_path / "options.pptx", _python_pptx_options_table)
    prs = rpptx.Presentation(source)
    held_cell = prs.slides[0].shapes[1].table.cell(3, 0)
    rows = prs.slides[0].shapes[1].table.rows
    row = rows.add_row()
    assert type(row).__name__ == "Row"
    assert row.height == rpptx.Inches(0.6)
    _assert_stale_after_exactly_one_bump(rpptx, lambda: held_cell.text)
    assert len(rows) == 5
    stable = prs.to_bytes()
    with pytest.raises(rpptx.StaleElementError, match=r"First invalidated by RowCollection.add_row "):
        held_cell.text = "unsafe"
    assert prs.to_bytes() == stable
    shape = prs.slides[0].shapes[1]
    assert (shape.width, shape.height) == (rpptx.Inches(9), rpptx.Inches(3))
    table = shape.table
    assert [len(table.rows), len(table.columns)] == [5, 4]
    assert [table.cell(4, column).text for column in range(4)] == [""] * 4
    table.cell(4, 0).text = "Risk"

    top = table.rows.add_row(0)
    table = prs.slides[0].shapes[1].table
    assert (top.height, table.cell(1, 1).text) == (rpptx.Inches(0.6), "A. Repair")
    assert table.cell(0, 1).fill.fore_color.rgb == RGBColor(0x3E, 0x4A, 0x59)
    table.rows.remove(table.rows[0])

    column = prs.slides[0].shapes[1].table.columns.add_column(-1)
    assert column.width == rpptx.Inches(2.25)
    column.width = rpptx.Inches(1)
    shape = prs.slides[0].shapes[1]
    assert (shape.width, shape.height) == (rpptx.Inches(10), rpptx.Inches(3))
    table = shape.table
    assert [table.cell(1, index).text for index in range(5)] == [
        "Cost (EUR)",
        "46,000",
        "212,000",
        "",
        "690,000",
    ]

    before = prs.to_bytes()
    other = rpptx.Presentation(source).slides[0].shapes[1].table
    with pytest.raises(IndexError, match="row index out of range"):
        table.rows.add_row(6)
    with pytest.raises(IndexError, match="column index out of range"):
        table.columns.add_column(-6)
    with pytest.raises(ValueError, match="row is not in this collection"):
        table.rows.remove(other.rows[0])
    with pytest.raises(ValueError, match="column is not in this collection"):
        table.columns.remove(other.columns[0])
    with pytest.raises(TypeError):
        table.rows.remove(table.columns[0])
    assert prs.to_bytes() == before
    held_row = table.rows[0]
    added_row = table.rows.add_row()
    with pytest.raises(rpptx.StaleElementError, match=r"table\.rows\[0\]\."):
        prs.slides[0].shapes[1].table.rows.remove(held_row)
    table.rows.remove(added_row)
    output = tmp_path / "rows-columns.pptx"
    prs.save(output)
    xml = _package_parts(output.read_bytes())["ppt/slides/slide1.xml"].decode()
    appended = xml[xml.rindex("<a:tr ") : xml.index("</a:tbl>")]
    assert appended.count('<a:endParaRPr sz="1500"/>') == 5
    assert appended.startswith('<a:tr h="548640"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>')
    assert '<a:r><a:rPr sz="1500"/><a:t>Risk</a:t></a:r>' in appended
    assert appended.index("Risk") < appended.index("<a:tcPr/>") < appended.index("</a:tc>")

    def measured(deck):
        _python_pptx_options_table(deck)
        deck.slides[0].shapes[1].height = rpptx.Inches(3)

    measured_source = _python_pptx_deck(tmp_path / "measured.pptx", measured)
    measured_deck = rpptx.Presentation(measured_source)
    measured_deck.slides[0].shapes[1].table.rows.add_row()
    assert measured_deck.slides[0].shapes[1].height == rpptx.Inches(3.6)
    measured_table = measured_deck.slides[0].shapes[1].table
    measured_table.rows.remove(measured_table.rows[0])
    assert measured_deck.slides[0].shapes[1].height == rpptx.Inches(3)

    single = rpptx.Presentation()
    single.slides.add_slide(single.slide_layouts[6]).shapes.add_table(1, 2, 0, 0, 200, 100)
    table = single.slides[0].shapes[0].table
    with pytest.raises(rpptx.RpptxError, match="at least one row"):
        table.rows.remove(table.rows[0])
    table.columns.remove(table.columns[1])
    table = single.slides[0].shapes[0].table
    with pytest.raises(rpptx.RpptxError, match="at least one column"):
        table.columns.remove(table.columns[0])
    assert single.slides[0].shapes[0].width == 100

    merged = rpptx.Presentation()
    merged.slides.add_slide(merged.slide_layouts[6]).shapes.add_table(3, 3, 0, 0, 300, 300)
    table = merged.slides[0].shapes[0].table
    table.cell(0, 0).merge(table.cell(1, 1))
    table.rows.add_row(1)
    table = merged.slides[0].shapes[0].table
    table.columns.add_column(1)
    table = merged.slides[0].shapes[0].table
    origin = table.cell(0, 0)
    assert (origin.span_height, origin.span_width) == (3, 3)
    table.rows.remove(table.rows[0])
    table = merged.slides[0].shapes[0].table
    assert (table.cell(0, 0).span_height, table.cell(0, 0).is_merge_origin) == (2, True)
    merged_output = tmp_path / "merged-rows.pptx"
    merged.save(merged_output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes[1]
    assert (len(oracle.table.rows), len(oracle.table.columns)) == (5, 5)
    assert oracle.height == sum(row.height for row in oracle.table.rows) == rpptx.Inches(3)
    assert oracle.width == sum(column.width for column in oracle.table.columns)
    assert [cell.text for cell in oracle.table.rows[4].cells] == ["Risk", "", "", "", ""]
    assert oracle.table.cell(4, 0).text_frame.paragraphs[0].runs[0].font.size == pptx.util.Pt(15)
    assert oracle.table.cell(4, 1)._tc.txBody.p_lst[0].endParaRPr.get("sz") == "1500"
    oracle = pptx.Presentation(merged_output).slides[0].shapes[0].table
    assert (oracle.cell(0, 0).span_height, oracle.cell(0, 0).span_width) == (2, 3)
    assert oracle.cell(1, 2).is_spanned and not oracle.cell(2, 0).is_spanned

def test_line_dash_style_and_ends_write_what_python_pptx_reads(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor
    from rpptx.enum.dml import (
        MSO_ARROWHEAD_LENGTH,
        MSO_ARROWHEAD_STYLE,
        MSO_ARROWHEAD_WIDTH,
        MSO_LINE,
        MSO_LINE_DASH_STYLE,
    )
    from rpptx.enum.shapes import MSO_CONNECTOR

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    assert MSO_LINE is MSO_LINE_DASH_STYLE
    for member in pptx.enum.dml.MSO_LINE:
        assert MSO_LINE[member.name] == member.value

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    plain = slide.shapes.add_shape("rect", 0, 0, 100, 100)
    plain.line.dash_style = None
    plain.line.tail_end.type = None
    plain.line.head_end.width = None
    assert not re.search(rb"<a:ln[ />]", plain.xml)

    prs.slides[0].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, rpptx.Inches(2), 0)
    connector = prs.slides[0].shapes[1]
    line = connector.line
    line.color.rgb = RGBColor(0x00, 0x00, 0x00)
    line.width = rpptx.Pt(2)
    tail = line.tail_end
    assert line.dash_style is None
    assert (tail.type, tail.width, tail.length) == (None, None, None)
    line.dash_style = MSO_LINE.DASH_DOT_DOT
    tail.type = MSO_ARROWHEAD_STYLE.TRIANGLE
    tail.width = MSO_ARROWHEAD_WIDTH.WIDE
    tail.length = MSO_ARROWHEAD_LENGTH.LONG
    line.head_end.type = MSO_ARROWHEAD_STYLE.OVAL
    assert (
        b'<a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        b'<a:prstDash val="lgDashDotDot"/><a:headEnd type="oval"/>'
        b'<a:tailEnd type="triangle" w="lg" len="lg"/></a:ln>'
    ) in connector.xml
    assert line.dash_style is MSO_LINE.DASH_DOT_DOT
    assert (tail.type, tail.width, tail.length) == (
        MSO_ARROWHEAD_STYLE.TRIANGLE,
        MSO_ARROWHEAD_WIDTH.WIDE,
        MSO_ARROWHEAD_LENGTH.LONG,
    )
    with pytest.raises(ValueError, match="other than DASH_STYLE_MIXED"):
        line.dash_style = MSO_LINE.DASH_STYLE_MIXED
    with pytest.raises(ValueError, match="MSO_ARROWHEAD_WIDTH"):
        tail.width = 7

    for member, value in (
        (MSO_LINE.SQUARE_DOT, "sysDash"),
        (MSO_LINE.ROUND_DOT, "sysDot"),
        (MSO_LINE.DOT, "dot"),
        (MSO_LINE.SYSTEM_DASH_DOT, "sysDashDot"),
        (MSO_LINE.SYSTEM_DASH_DOT_DOT, "sysDashDotDot"),
    ):
        line.dash_style = member
        assert f'<a:prstDash val="{value}"/>'.encode() in connector.xml
        assert line.dash_style is member
    line.dash_style = MSO_LINE.LONG_DASH

    output = tmp_path / "lines.pptx"
    prs.save(output)
    oracle = pptx.Presentation(output).slides[0].shapes[1]
    assert oracle.line.dash_style == pptx.enum.dml.MSO_LINE.LONG_DASH
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    oracle_ln = oracle._element.spPr.find(a + "ln")
    assert dict(oracle_ln.find(a + "tailEnd").attrib) == {"type": "triangle", "w": "lg", "len": "lg"}
    assert dict(oracle_ln.find(a + "headEnd").attrib) == {"type": "oval"}

    # PowerPoint scripting writes type="none" and keeps the size when an arrow is removed.
    tail.type = MSO_ARROWHEAD_STYLE.NONE
    assert b'<a:tailEnd type="none" w="lg" len="lg"/>' in connector.xml
    tail.type = None
    tail.width = None
    assert b'<a:tailEnd len="lg"/>' in connector.xml
    tail.length = None
    line.head_end.type = None
    line.dash_style = None
    assert b"<a:prstDash" not in connector.xml
    assert b"End" not in connector.xml
    assert b'<a:ln w="25400"><a:solidFill>' in connector.xml

    held = line.tail_end
    prs.slides.add_slide(prs.slide_layouts[6])
    assert held.type == line.tail_end.type

    # A custom dash reads as None, and a preset or None replaces it.
    deck = pptx.Presentation()
    shape = deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_shape(
        pptx.enum.shapes.MSO_SHAPE.RECTANGLE, 0, 0, 10, 10
    )
    shape.line.width = rpptx.Pt(1)
    shape.line._get_or_add_ln().append(
        pptx.oxml.parse_xml(
            '<a:custDash xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:ds d="100000" sp="50000"/></a:custDash>'
        )
    )
    source = tmp_path / "custom.pptx"
    deck.save(source)
    custom = rpptx.Presentation(source).slides[0].shapes[0]
    assert custom.line.dash_style is None
    custom.line.dash_style = None
    assert b"custDash" not in custom.xml
    assert custom.line.width == rpptx.Pt(1)


def test_colour_edits_keep_python_pptx_brightness_transforms(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor

    def build(deck):
        pptx = pytest.importorskip("pptx")
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        shape = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RECTANGLE, 0, 0, 10, 10)
        shape.fill.solid()
        shape.fill.fore_color.rgb = pptx.dml.color.RGBColor(0xFF, 0x00, 0x00)
        shape.fill.fore_color.brightness = -0.25

    source = _python_pptx_deck(tmp_path / "brightness.pptx", build)
    prs = rpptx.Presentation(source)
    prs.slides[0].shapes[0].fill.fore_color.rgb = RGBColor(0x00, 0x00, 0xFF)
    output = tmp_path / "brightness-out.pptx"
    prs.save(output)
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    color = pptx.Presentation(output).slides[0].shapes[0].fill.fore_color
    assert color.rgb == pptx.dml.color.RGBColor(0x00, 0x00, 0xFF)
    assert color.brightness == pytest.approx(-0.25)


def test_shadow_reads_a_producer_outer_shadow_and_keeps_python_pptx_inherit(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor

    drawingml = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

    def build(deck):
        pptx = pytest.importorskip("pptx")
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import qn

        slide = deck.slides.add_slide(deck.slide_layouts[6])
        card = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, 0, 0, 10, 10)
        card.shadow.inherit = False
        effects = card._element.spPr.find(qn("a:effectLst"))
        effects.append(parse_xml(f'<a:glow {drawingml} rad="63500"><a:srgbClr val="FF0000"/></a:glow>'))
        effects.append(parse_xml(
            f'<a:outerShdw {drawingml} blurRad="152400" dist="76200" dir="5400000" algn="t" '
            'rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw>'
        ))
        slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RECTANGLE, 0, 0, 10, 10)

    source = _python_pptx_deck(tmp_path / "producer.pptx", build)
    prs = rpptx.Presentation(source)
    card, plain = prs.slides[0].shapes[0].shadow, prs.slides[0].shapes[1].shadow
    assert (card.inherit, card.visible, card.color.rgb, card.alpha) == (
        False, True, RGBColor(0, 0, 0), pytest.approx(0.6)
    )
    assert (card.blur_radius, card.distance, card.direction) == (152400, 76200, 90.0)
    assert (card.align, card.rotate_with_shape) == ("t", False)
    assert (plain.inherit, plain.visible, plain.color.rgb, plain.alpha, plain.blur_radius) == (
        True, False, None, None, None
    )
    assert (plain.distance, plain.direction, plain.align, plain.rotate_with_shape) == (
        None, None, None, None
    )

    plain.inherit = False
    assert (plain.inherit, plain.visible) == (False, False)
    output = tmp_path / "inherit.pptx"
    prs.save(output)
    slide_xml = _package_parts(output.read_bytes())["ppt/slides/slide1.xml"].decode()
    assert '</a:prstGeom><a:effectLst/></p:spPr>' in slide_xml
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    assert [shape.shadow.inherit for shape in pptx.Presentation(output).slides[0].shapes] == [
        False, False
    ]

    card.visible = False
    plain.inherit = True
    plain.visible = False
    assert (card.inherit, card.visible, card.alpha, plain.inherit) == (False, False, None, True)
    prs.save(output)
    slide_xml = _package_parts(output.read_bytes())["ppt/slides/slide1.xml"].decode()
    assert '<a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow></a:effectLst>' in slide_xml
    assert slide_xml.count("effectLst") == 2
    assert [shape.shadow.inherit for shape in pptx.Presentation(output).slides[0].shapes] == [
        False, True
    ]

    card.visible = True
    assert (card.color.rgb, card.alpha) == (None, 0.4)
    card.alpha = 1.0
    prs.save(output)
    slide_xml = _package_parts(output.read_bytes())["ppt/slides/slide1.xml"].decode()
    assert (
        '<a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow><a:outerShdw '
        'blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr '
        'val="black"/></a:outerShdw></a:effectLst>'
    ) in slide_xml


def test_shadow_colour_rpptx_does_not_model_reads_none_and_is_replaced_whole(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor

    source = tmp_path / "scrgb.pptx"
    output = tmp_path / "scrgb-out.pptx"
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_shape("rect", 0, 0, 100, 100)
    prs.slides[0].shapes[0].shadow.distance = 12700
    prs.save(source)
    _replace_in_slide(
        source,
        source,
        '<a:prstClr val="black"><a:alpha val="40000"/></a:prstClr>',
        '<a:scrgbClr r="0" g="0" b="0"><a:alpha val="50000"/></a:scrgbClr>',
    )

    shadow = rpptx.Presentation(source).slides[0].shapes[0].shadow
    assert (shadow.visible, shadow.color.rgb, shadow.alpha, shadow.distance) == (
        True, None, None, 12700
    )
    with pytest.raises(ValueError, match="a:scrgbClr or a:hslClr"):
        shadow.alpha = 0.5
    shadow.color.rgb = RGBColor(0x10, 0x20, 0x30)
    assert (shadow.color.rgb, shadow.alpha) == (RGBColor(0x10, 0x20, 0x30), 1.0)
    shadow.alpha = 0.5

    presentation = rpptx.Presentation(source)
    presentation.slides[0].shapes[0].shadow.color.rgb = RGBColor(0x10, 0x20, 0x30)
    presentation.save(output)
    slide_xml = _package_parts(output.read_bytes())["ppt/slides/slide1.xml"].decode()
    assert "scrgbClr" not in slide_xml
    assert (
        '<a:outerShdw blurRad="50800" dist="12700" dir="2700000" algn="tl" rotWithShape="0">'
        '<a:srgbClr val="102030"/></a:outerShdw>'
    ) in slide_xml


def test_shadow_parameters_write_the_outer_shadow_on_every_kind_python_pptx_shadows(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor
    from rpptx.enum.shapes import MSO_CONNECTOR

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_shape("rect", 0, 0, 100, 100)
    prs.slides[0].shapes.add_textbox(0, 0, 100, 100)
    prs.slides[0].shapes.add_picture(io.BytesIO(_tiny_png()), 0, 0)
    prs.slides[0].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 100, 100)
    prs.slides[0].shapes.add_group_shape()
    shapes = list(prs.slides[0].shapes)
    for shape in shapes:
        shadow = shape.shadow
        shadow.color.rgb = RGBColor(0x12, 0x34, 0x56)
        shadow.alpha = 0.25
        shadow.blur_radius = rpptx.Pt(6)
        shadow.distance = rpptx.Pt(4)
        shadow.direction = -45.0
        shadow.align = "ctr"
        shadow.rotate_with_shape = True
        assert (shadow.inherit, shadow.visible, shadow.color.rgb) == (
            False, True, RGBColor(0x12, 0x34, 0x56)
        )
        assert (shadow.alpha, shadow.blur_radius, shadow.distance) == (0.25, 76200, 50800)
        assert (shadow.direction, shadow.align, shadow.rotate_with_shape) == (315.0, "ctr", True)

    with pytest.raises(ValueError, match="shadow alpha must be between 0.0 and 1.0"):
        shapes[0].shadow.alpha = 1.5
    with pytest.raises(ValueError, match="shadow align must be one of tl, t, tr"):
        shapes[0].shadow.align = "middle"
    with pytest.raises(ValueError, match="shadow blur_radius must be between 0"):
        shapes[0].shadow.blur_radius = -1
    with pytest.raises(ValueError, match="shadow direction must be a finite number"):
        shapes[0].shadow.direction = float("nan")
    with pytest.raises(TypeError, match="ShadowFormat.color.rgb"):
        shapes[0].shadow.color.rgb = (1, 2)
    held = shapes[0].shadow
    held_color = held.color
    prs.slides.add_slide(prs.slide_layouts[6])
    assert held.visible == shapes[0].shadow.visible
    assert held_color.rgb == shapes[0].shadow.color.rgb
    with pytest.raises(NotImplementedError, match="GraphicFrame"):
        _ = prs.slides[1].shapes.add_table(1, 1, 0, 0, 10, 10).shadow

    output = tmp_path / "shadows.pptx"
    prs.save(output)
    slide_xml = _package_parts(output.read_bytes())["ppt/slides/slide1.xml"].decode()
    written = (
        '<a:effectLst><a:outerShdw blurRad="76200" dist="50800" dir="18900000" algn="ctr" '
        'rotWithShape="1"><a:srgbClr val="123456"><a:alpha val="25000"/></a:srgbClr>'
        "</a:outerShdw></a:effectLst>"
    )
    assert slide_xml.count(written + "</p:spPr>") == 4
    assert slide_xml.count(written + "</p:grpSpPr>") == 1
    assert rpptx.Presentation(output).slides[0].shapes[4].shadow.direction == 315.0

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes
    assert [shape.shadow.inherit for shape in list(oracle)[:5]] == [False] * 5


def test_pictures_accept_bytes_and_file_objects_and_replace_their_image(tmp_path):
    import rpptx

    red, blue, jpeg = _tiny_png(0xFF, 0, 0), _tiny_png(0, 0, 0xFF), _template_jpeg()
    image_path = tmp_path / "red.png"
    image_path.write_bytes(red)
    stream = io.BytesIO(red)
    stream.read()
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(image_path, 0, 0)
    prs.slides[0].shapes.add_picture(stream, 0, 0, width=rpptx.Inches(1))
    prs.slides[0].shapes.add_picture(jpeg, 0, 0)
    prs.slides[0].shapes.add_textbox(0, 0, 10, 10)
    shapes = prs.slides[0].shapes
    assert [shape.image.blob for shape in shapes[:3]] == [red, red, jpeg]
    image = shapes[2].image
    assert (image.content_type, image.ext) == ("image/jpeg", "jpg")
    assert (shapes[0].image.content_type, shapes[0].image.ext) == ("image/png", "png")
    assert shapes[1].width == rpptx.Inches(1)
    with pytest.raises(ValueError, match="shape is not a picture"):
        _ = shapes[3].image

    held = shapes[0]
    held.replace_image(blue)
    assert held.image.blob == blue
    assert shapes[1].image.blob == red
    shapes[1].replace_image(io.BytesIO(jpeg))
    assert shapes[1].image.blob == jpeg
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="not a supported image"):
        held.replace_image(b"not an image")
    assert prs.to_bytes() == before
    with pytest.raises(rpptx.RpptxError, match="unsupported image bytes"):
        prs.slides[0].shapes.add_picture(b"\x89PNG", 0, 0)
    output = tmp_path / "pictures.pptx"
    prs.save(output)
    media = [name for name in _package_parts(output.read_bytes()) if name.startswith("ppt/media/")]
    assert len(media) == 2

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes
    assert [oracle[index].image.blob for index in range(3)] == [blue, jpeg, jpeg]


def test_picture_crop_matches_python_pptx_in_both_directions(tmp_path):
    import rpptx

    header = struct.pack(">IIBBBBB", 2, 1, 8, 2, 0, 0, 0)
    red_then_blue = (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", header)
        + _png_chunk(b"IDAT", zlib.compress(bytes((0, 0xFF, 0, 0, 0, 0, 0xFF))))
        + _png_chunk(b"IEND", b"")
    )
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(red_then_blue, 0, 0, rpptx.Inches(2), rpptx.Inches(1))
    prs.slides[0].shapes.add_textbox(0, 0, 10, 10)
    picture = prs.slides[0].shapes[0]
    assert (picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom) == (0.0,) * 4
    before = prs.to_bytes()
    picture.crop_right = 0.0
    assert prs.to_bytes() == before
    uncropped = prs.render_slide_to_png(0, dpi=36.0)
    held = prs.slides[0].shapes[0]
    picture.crop_left = 0.5
    picture.crop_bottom = -0.125
    assert (held.crop_left, held.crop_bottom) == (0.5, -0.125)
    assert prs.render_slide_to_png(0, dpi=36.0) != uncropped
    cropped = prs.to_bytes()
    for bad in (float("nan"), float("inf"), 21474.83648 + 1):
        with pytest.raises(ValueError, match="crop must be a finite fraction"):
            picture.crop_top = bad
    textbox = prs.slides[0].shapes[1]
    with pytest.raises(ValueError, match="shape is not a picture"):
        _ = textbox.crop_left
    with pytest.raises(ValueError, match="shape is not a picture"):
        textbox.crop_left = 0.1
    assert prs.to_bytes() == cropped
    output = tmp_path / "cropped.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes[0]
    assert (oracle.crop_left, oracle.crop_top, oracle.crop_right, oracle.crop_bottom) == (
        0.5,
        0.0,
        0.0,
        -0.125,
    )

    def build(deck):
        written = deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_picture(
            io.BytesIO(_tiny_png()), 0, 0
        )
        written.crop_top = 0.2
        written.crop_right = 1 / 3

    source = _python_pptx_deck(tmp_path / "python-pptx-crop.pptx", build)
    read = rpptx.Presentation(source).slides[0].shapes[0]
    assert (read.crop_left, read.crop_top, read.crop_right) == (0.0, 0.2, 0.33333)


def test_shapes_move_changes_z_order_and_preserves_owner_handles(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    for name in ("back", "middle", "front"):
        prs.slides[0].shapes.add_shape("rect", 0, 0, 100, 100).name = name
    shapes = prs.slides[0].shapes
    held = shapes[0]
    shapes.move(0, -1)
    assert [shape.name for shape in prs.slides[0].shapes] == ["middle", "front", "back"]
    assert held.name == "back"
    assert len(shapes) == 3
    prs.slides[0].shapes.move(-1, 1)
    assert [shape.name for shape in prs.slides[0].shapes] == ["middle", "back", "front"]
    before = prs.to_bytes()
    with pytest.raises(IndexError, match="shape index out of range"):
        prs.slides[0].shapes.move(0, 3)
    group = prs.slides[0].shapes.add_group_shape()
    with pytest.raises(ValueError, match="nested shape collections cannot remove or move shapes"):
        group.shapes.move(0, 0)
    prs.slides[0].shapes.remove(prs.slides[0].shapes[3])
    assert prs.to_bytes() == before
    output = tmp_path / "z-order.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes
    assert [shape.name for shape in oracle] == ["middle", "back", "front"]


def test_group_shapes_add_members_and_the_group_fits_them(tmp_path):
    import rpptx
    from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE

    emu = 914_400
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    group = prs.slides[0].shapes.add_group_shape()
    shapes = group.shapes
    textbox = shapes.add_textbox(emu, emu, emu, emu)
    assert group.name is not None
    assert len(shapes) == 1
    textbox.text = "inside"
    group = prs.slides[0].shapes[0]
    assert (group.left, group.top, group.width, group.height) == (emu, emu, emu, emu)

    prs.slides[0].shapes[0].shapes.add_shape(MSO_SHAPE.OVAL, 2 * emu, emu // 2, emu, emu)
    prs.slides[0].shapes[0].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, emu, emu)
    prs.slides[0].shapes[0].shapes.add_table(1, 2, emu, 2 * emu, emu, emu)
    prs.slides[0].shapes[0].shapes.add_picture(io.BytesIO(_tiny_png()), emu, emu, emu, emu)
    inner = prs.slides[0].shapes[0].shapes.add_group_shape()
    assert (inner.left, inner.width, len(inner.shapes)) == (0, 0, 0)
    nested = inner.shapes.add_textbox(3 * emu, 3 * emu, emu, emu)
    assert nested.left == 3 * emu
    group = prs.slides[0].shapes[0]
    assert [shape.shape_type for shape in group.shapes] == [
        MSO_SHAPE_TYPE.TEXT_BOX,
        MSO_SHAPE_TYPE.AUTO_SHAPE,
        MSO_SHAPE_TYPE.LINE,
        MSO_SHAPE_TYPE.TABLE,
        MSO_SHAPE_TYPE.PICTURE,
        MSO_SHAPE_TYPE.GROUP,
    ]
    assert (group.left, group.top, group.width, group.height) == (0, 0, 4 * emu, 4 * emu)
    inner = group.shapes[5]
    assert (inner.left, inner.top, inner.width, inner.height) == (3 * emu, 3 * emu, emu, emu)
    ids = [shape.shape_id for shape in group.shapes] + [group.shape_id, inner.shapes[0].shape_id]
    assert len(set(ids)) == len(ids) == 8
    output = tmp_path / "group.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes[0]
    assert (oracle.left, oracle.top, oracle.width, oracle.height) == (0, 0, 4 * emu, 4 * emu)
    assert [int(shape.shape_type) for shape in oracle.shapes] == [17, 1, 9, 19, 13, 6]
    assert oracle.shapes[0].text_frame.text == "inside"
    assert oracle.shapes[5].shapes[0].left == 3 * emu


def test_a_new_group_holds_one_text_box_that_its_extents_cover(tmp_path):
    import rpptx

    emu = 914_400
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    g = prs.slides[0].shapes.add_group_shape()
    g.shapes.add_textbox(emu, emu, emu, emu)
    output = tmp_path / "one-member.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    group = pptx.Presentation(output).slides[0].shapes[0]
    assert [int(shape.shape_type) for shape in group.shapes] == [17]
    member = group.shapes[0]
    assert (group.left, group.top, group.width, group.height) == (
        member.left,
        member.top,
        member.width,
        member.height,
    )
    xfrm = group._element.grpSpPr.xfrm
    assert (xfrm.chOff.x, xfrm.chOff.y, xfrm.chExt.cx, xfrm.chExt.cy) == (emu, emu, emu, emu)


def test_only_groups_accept_members_and_nested_collections_do_not_remove():
    import rpptx

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_group_shape()
    prs.slides[0].shapes[0].shapes.add_textbox(0, 0, 10, 10)
    textbox = prs.slides[0].shapes.add_textbox(0, 0, 10, 10)
    before = prs.to_bytes()
    with pytest.raises(ValueError, match="shape is not a group"):
        textbox.shapes.add_textbox(0, 0, 10, 10)
    with pytest.raises(ValueError, match="shape is not a group"):
        textbox.shapes.add_picture("missing.png", 0, 0)
    with pytest.raises(rpptx.RpptxError):
        prs.slides[0].shapes[0].shapes.add_shape("notAPreset", 0, 0, 10, 10)
    group = prs.slides[0].shapes[0]
    with pytest.raises(ValueError, match="nested shape collections cannot remove or move shapes"):
        group.shapes.remove(group.shapes[0])
    assert prs.to_bytes() == before
    assert len(prs.slides[0].shapes[0].shapes) == 1


def test_a_group_in_an_alternate_content_fallback_does_not_take_new_shapes(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_group_shape()
    prs.slides[0].shapes[0].shapes.add_textbox(0, 0, 100, 100)
    parts = _package_parts(prs.to_bytes())
    slide = parts["ppt/slides/slide1.xml"].decode()
    start = slide.index("<p:grpSp>")
    end = slide.index("</p:grpSp>") + len("</p:grpSp>")
    group = slide[start:end]
    parts["ppt/slides/slide1.xml"] = (
        slide[:start]
        + '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        + ' xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
        + f'<mc:Choice Requires="p14">{group}</mc:Choice><mc:Fallback>{group}</mc:Fallback>'
        + "</mc:AlternateContent>"
        + slide[end:]
    ).encode()
    source = tmp_path / "fallback-group.pptx"
    source.write_bytes(_package_bytes(parts))
    prs = rpptx.Presentation(source)
    fallback_group = prs.slides[0].shapes[0].shapes[0]
    assert int(fallback_group.shape_type) == 6
    before = prs.to_bytes()
    with pytest.raises(ValueError, match="mc:AlternateContent fallback cannot take new shapes"):
        fallback_group.shapes.add_textbox(0, 0, 10, 10)
    assert prs.to_bytes() == before
    assert int(fallback_group.shape_type) == 6
    prs.slides[0].shapes.add_textbox(0, 0, 10, 10)
    with pytest.raises(rpptx.StaleElementError, match=r"First invalidated by ShapeCollection.add_textbox "):
        _ = fallback_group.shape_type
    fresh_fallback = prs.slides[0].shapes[0].shapes[0]
    prs.slides[0].shapes[1].text = "ordinary edit"
    with pytest.raises(rpptx.StaleElementError, match=r"First invalidated by Shape.text "):
        _ = fresh_fallback.shape_type


def test_an_empty_nested_group_lets_python_pptx_refit_its_parent(tmp_path):
    import rpptx

    emu = 914_400
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_group_shape()
    prs.slides[0].shapes[0].shapes.add_textbox(emu, emu, emu, emu)
    inner = prs.slides[0].shapes[0].shapes.add_group_shape()
    assert (inner.left, inner.top, inner.width, inner.height) == (0, 0, 0, 0)
    outer = prs.slides[0].shapes[0]
    assert (outer.left, outer.top, outer.width, outer.height) == (emu, emu, emu, emu)
    output = tmp_path / "empty-nested.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output)
    group = oracle.slides[0].shapes[0]
    group.shapes.add_textbox(2 * emu, 2 * emu, emu, emu)
    # python-pptx counts the empty group's zero box, which rpptx leaves out.
    assert (group.left, group.top, group.width, group.height) == (0, 0, 3 * emu, 3 * emu)
    oracle.save(output)
    assert len(rpptx.Presentation(output).slides[0].shapes[0].shapes) == 3


def test_shapes_and_slides_are_removed_and_reordered_with_stale_handles(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    for index in range(3):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(0, 0, 10, 10).text = f"slide {index}"
    held = prs.slides[0]
    prs.slides.move(0, -1)
    assert [slide.shapes[0].text for slide in prs.slides] == ["slide 1", "slide 2", "slide 0"]
    assert held.shapes[0].text == "slide 0"
    with pytest.raises(IndexError):
        prs.slides.move(0, 3)
    prs.slides.remove(prs.slides[1])
    assert [slide.shapes[0].text for slide in prs.slides] == ["slide 1", "slide 0"]
    other = rpptx.Presentation()
    other_slide = other.slides.add_slide(other.slide_layouts[6])
    with pytest.raises(ValueError, match="slide is not in this collection"):
        prs.slides.remove(other_slide)

    shapes = prs.slides[0].shapes
    picture = shapes.add_picture(io.BytesIO(_tiny_png()), 0, 0)
    group = prs.slides[0].shapes.add_group_shape()
    assert len(prs.slides[0].shapes) == 3
    shapes = prs.slides[0].shapes
    selected = shapes[0]
    selected_id, selected_xml = selected.shape_id, selected.xml
    before_foreign_removal = prs.to_bytes()
    with pytest.raises(ValueError, match=r"^shape owner is not in this presentation collection$"):
        prs.slides[1].shapes.remove(shapes[0])
    assert prs.to_bytes() == before_foreign_removal
    assert selected.shape_id == shapes[0].shape_id == selected_id
    assert selected.xml == selected_xml
    shapes.remove(picture)
    with pytest.raises(rpptx.StaleElementError, match=r"First invalidated by ShapeCollection.remove "):
        _ = picture.name
    assert group.name is not None
    assert [shape.shape_type for shape in prs.slides[0].shapes] == [17, 6]
    output = tmp_path / "removed.pptx"
    prs.save(output)
    parts = _package_parts(output.read_bytes())
    assert not [name for name in parts if name.startswith("ppt/media/")]
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output)
    assert [len(slide.shapes) for slide in oracle.slides] == [2, 1]


def test_slide_duplicate_inserts_after_the_source_with_its_notes(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    for index in range(3):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(0, 0, 10, 10).text = f"slide {index}"
    prs.slides[1].notes_text = "note 1"
    source = prs.slides[1]
    held = prs.slides[2].shapes[0]

    duplicate = prs.slides.duplicate(source)
    assert [slide.shapes[0].text for slide in prs.slides] == [
        "slide 0",
        "slide 1",
        "slide 1",
        "slide 2",
    ]
    assert [slide.notes_text for slide in prs.slides] == [None, "note 1", "note 1", None]
    assert (duplicate.shapes[0].text, duplicate.notes_text) == ("slide 1", "note 1")
    assert held.text == "slide 2"
    assert source.notes_text == "note 1"
    duplicate.notes_text = "copy note"
    assert [slide.notes_text for slide in prs.slides] == [None, "note 1", "copy note", None]

    other = rpptx.Presentation()
    other_slide = other.slides.add_slide(other.slide_layouts[6])
    with pytest.raises(ValueError, match="slide is not in this collection"):
        prs.slides.duplicate(other_slide)

    prs.add_comment_author(
        id="{11111111-1111-1111-1111-111111111111}",
        name="Ada Lovelace",
        user_id="ada@example.com",
        provider_id="local",
    )
    prs.slides[0].add_comment(
        id="{22222222-2222-2222-2222-222222222222}",
        author_id="{11111111-1111-1111-1111-111111111111}",
        created="2026-09-14T10:30:00Z",
        text="Blocks duplication",
    )
    commented = prs.slides[0]
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="modern comments"):
        prs.slides.duplicate(commented)
    assert prs.to_bytes() == before
    assert (len(prs.slides), commented.shapes[0].text) == (4, "slide 0")

    output = tmp_path / "duplicated.pptx"
    prs.save(output)
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output)
    assert [slide.shapes[0].text for slide in oracle.slides] == [
        "slide 0",
        "slide 1",
        "slide 1",
        "slide 2",
    ]
    assert [
        slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else None
        for slide in oracle.slides
    ] == [None, "note 1", "copy note", None]


def test_presentation_from_bytes_opens_like_a_path_and_rejects_other_bytes(tmp_path):
    import rpptx

    prs = _textbox_presentation(rpptx)
    prs.slides[0].notes_text = "note"
    path = tmp_path / "source.pptx"
    prs.save(path)

    reopened = rpptx.Presentation.from_bytes(path.read_bytes())
    assert len(reopened.slides) == 1
    assert (reopened.slides[0].shapes[0].text, reopened.slides[0].notes_text) == (
        "hello",
        "note",
    )
    assert reopened.to_bytes() == rpptx.Presentation(path).to_bytes()
    with pytest.raises(rpptx.PackageError):
        rpptx.Presentation.from_bytes(b"not a package")


def test_try_replace_text_checks_the_expected_count_before_publishing():
    import pickle

    import rpptx

    prs = _textbox_presentation(rpptx)
    prs.slides[0].shapes[0].text = "NAME and NAME"
    prs.slides[0].shapes.add_table(1, 1, 0, 0, 100, 100).table.cell(0, 0).text = "NAME"
    prs.slides[0].notes_text = "Notes for NAME"
    held = prs.slides[0].shapes[0]
    before = prs.to_bytes()

    with pytest.raises(rpptx.ReplacementCountError) as raised:
        prs.try_replace_text("NAME", "Ada", expect=3)
    error = raised.value
    assert isinstance(error, rpptx.RpptxError)
    assert str(error) == 'expected 3 replacement(s) of "NAME", found 4'
    assert (error.expected, error.found) == (3, 4)
    # Worker pools pickle exceptions, so the counts must survive the trip.
    copied = pickle.loads(pickle.dumps(error))
    assert (str(copied), copied.expected, copied.found) == (str(error), 3, 4)
    assert prs.to_bytes() == before
    assert held.text == "NAME and NAME"

    assert prs.try_replace_text("MISSING", "x") == 0
    assert prs.try_replace_text("MISSING", "x", expect=0) == 0
    with pytest.raises(rpptx.ReplacementCountError, match="found 0"):
        prs.try_replace_text("MISSING", "x", expect=1)
    with pytest.raises(rpptx.RpptxError, match="placeholder must not be empty"):
        prs.try_replace_text("", "x")
    assert prs.to_bytes() == before
    assert held.text == "NAME and NAME"

    assert prs.try_replace_text("NAME", "Ada", expect=4) == 4
    assert held.text == "Ada and Ada"
    shapes = prs.slides[0].shapes
    assert (shapes[0].text, shapes[1].table.cell(0, 0).text) == ("Ada and Ada", "Ada")
    assert prs.slides[0].notes_text == "Notes for Ada"
    held = prs.slides[0].shapes[0]
    assert prs.try_replace_text("Ada", "Grace") == 4
    assert held.text == prs.slides[0].shapes[0].text == "Grace and Grace"


def test_slide_resolve_and_remove_comment_match_the_cli_operations():
    import rpptx

    author_id = "{11111111-1111-1111-1111-111111111111}"
    comment_id = "{22222222-2222-2222-2222-222222222222}"
    reply_id = "{33333333-3333-3333-3333-333333333333}"
    second_id = "{55555555-5555-5555-5555-555555555555}"
    unknown_id = "{99999999-9999-9999-9999-999999999999}"
    created = "2026-09-14T10:30:00Z"
    prs = rpptx.Presentation()
    for _ in range(2):
        prs.slides.add_slide(prs.slide_layouts[6])
    prs.add_comment_author(
        id=author_id, name="Ada Lovelace", user_id="ada@example.com", provider_id="local"
    )
    prs.slides[0].add_comment(id=comment_id, author_id=author_id, created=created, text="Thread")
    prs.slides[0].reply_to_comment(
        comment_id, id=reply_id, author_id=author_id, created=created, text="Reply"
    )
    prs.slides[0].add_comment(id=second_id, author_id=author_id, created=created, text="Second")

    slide = prs.slides[0]
    before = prs.to_bytes()
    for operation, comment in (
        (slide.resolve_comment, unknown_id),
        (slide.resolve_comment, reply_id),
        (slide.remove_comment, unknown_id),
        (prs.slides[1].remove_comment, comment_id),
    ):
        with pytest.raises(rpptx.RpptxError, match="unknown comment id"):
            operation(comment)
    assert prs.to_bytes() == before
    assert [comment.status for comment in slide.comments] == [None, None]

    slide.resolve_comment(comment_id)
    assert slide.comments[0].status == "resolved"
    thread, second = prs.slides[0].comments
    assert (thread.id, thread.status, second.status) == (comment_id, "resolved", None)
    assert [(reply.id, reply.status) for reply in thread.replies] == [(reply_id, None)]

    prs.slides[0].remove_comment(reply_id)
    assert prs.slides[0].comments[0].replies == ()
    prs.slides[0].remove_comment(comment_id)
    assert [comment.id for comment in prs.slides[0].comments] == [second_id]
    reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
    assert reopened.slides[0].comments == prs.slides[0].comments

    # The facade keeps the emptied comments part, so duplicate still refuses.
    prs.slides[0].remove_comment(second_id)
    assert prs.slides[0].comments == ()
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="modern comments"):
        prs.slides.duplicate(prs.slides[0])
    assert (prs.to_bytes(), len(prs.slides)) == (before, 2)


def test_validate_returns_the_issues_the_cli_prints_as_frozen_snapshots(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    assert prs.validate() == ()
    prs.slides.add_slide(prs.slide_layouts[6])
    for text in ("first", "second"):
        prs.slides[0].shapes.add_textbox(0, 0, 10, 10).text = text
    assert prs.validate() == ()
    first, second = (shape.shape_id for shape in prs.slides[0].shapes)
    source = tmp_path / "source.pptx"
    broken = tmp_path / "broken.pptx"
    prs.save(source)
    _replace_in_slide(source, broken, f'id="{second}"', f'id="{first}"')

    broken_deck = rpptx.Presentation(broken)
    (issue,) = broken_deck.validate()
    assert (issue.kind, issue.message) == (
        "duplicate_shape_id",
        f"DuplicateShapeId {{ slide: 0, id: {first} }}",
    )
    assert broken_deck.validate() == (issue,)
    with pytest.raises(AttributeError):
        issue.kind = "other"


def test_try_replace_text_scoped_to_a_slide_or_a_text_frame(tmp_path):
    import rpptx

    prs = _textbox_presentation(rpptx)
    prs.slides[0].shapes[0].text = "Same text"
    table = prs.slides[0].shapes.add_table(1, 1, 0, 0, 100, 100).table
    table.cell(0, 0).text = "Same text"
    prs.slides[0].shapes.add_textbox(0, 0, 100, 100).text = "Same text"
    prs.slides[0].notes_text = "Same text in the notes"
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[1].shapes.add_textbox(0, 0, 100, 100).text = "Same text"
    prs.slides[1].notes_text = "Same text in the notes"
    before = prs.to_bytes()

    with pytest.raises(rpptx.ReplacementCountError) as raised:
        prs.try_replace_text("Same text", "Other text", expect=1)
    assert str(raised.value) == 'expected 1 replacement(s) of "Same text", found 6'
    held = prs.slides[0]
    with pytest.raises(rpptx.ReplacementCountError) as raised:
        held.try_replace_text("Same text", "Other text", expect=1)
    assert str(raised.value) == 'expected 1 replacement(s) of "Same text", found 4'
    assert (raised.value.expected, raised.value.found) == (1, 4)
    with pytest.raises(rpptx.ReplacementCountError, match="found 3"):
        held.try_replace_text("Same text", "Other text", expect=1, notes=False)
    with pytest.raises(rpptx.RpptxError, match="placeholder must not be empty"):
        held.try_replace_text("", "x")
    assert held.try_replace_text("MISSING", "x") == 0
    assert held.try_replace_text("MISSING", "x", expect=0) == 0
    assert prs.to_bytes() == before

    assert held.try_replace_text("Same text", "Other text", expect=3, notes=False) == 3
    assert held.shapes[0].text == "Other text"
    first = prs.slides[0]
    shapes = first.shapes
    assert shapes[0].text == "Other text"
    assert shapes[1].table.cell(0, 0).text == "Other text"
    assert shapes[2].text == "Other text"
    assert first.notes_text == "Same text in the notes"
    assert first.try_replace_text("Same text", "Other text") == 1
    assert prs.slides[0].notes_text == "Other text in the notes"
    assert prs.slides[1].shapes[0].text == "Same text"
    assert prs.slides[1].notes_text == "Same text in the notes"

    frame = prs.slides[1].shapes[0].text_frame
    frame.text = "Same text and Same text"
    frame = prs.slides[1].shapes[0].text_frame
    snapshot = prs.to_bytes()
    with pytest.raises(rpptx.ReplacementCountError) as raised:
        frame.try_replace_text("Same text", "Other text", expect=1)
    assert str(raised.value) == 'expected 1 replacement(s) of "Same text", found 2'
    with pytest.raises(rpptx.RpptxError, match="placeholder must not be empty"):
        frame.try_replace_text("", "x")
    assert frame.try_replace_text("MISSING", "x") == 0
    assert prs.to_bytes() == snapshot
    assert frame.try_replace_text("Same text", "Other text", expect=2) == 2
    assert frame.text == "Other text and Other text"
    assert prs.slides[1].shapes[0].text == "Other text and Other text"
    assert prs.slides[1].notes_text == "Same text in the notes"
    assert prs.slides[1].shapes[0].text_frame.try_replace_text("Other", "New") == 2

    if importlib.util.find_spec("pptx") is None:
        return
    from pptx import Presentation as OraclePresentation

    # rpptx cannot add group children, so python-pptx builds a group whose
    # text box splits the match across a bold run and a plain one.
    source = tmp_path / "grouped.pptx"
    oracle = OraclePresentation()
    slide = oracle.slides.add_slide(oracle.slide_layouts[6])
    group = slide.shapes.add_group_shape()
    paragraph = group.shapes.add_textbox(0, 0, 100, 100).text_frame.paragraphs[0]
    bold = paragraph.add_run()
    bold.text = "Same te"
    bold.font.bold = True
    paragraph.add_run().text = "xt in a group"
    slide.shapes.add_textbox(0, 0, 100, 100).text_frame.text = "Same text"
    oracle.save(source)

    prs = rpptx.Presentation(source)
    frame = prs.slides[0].shapes[0].shapes[0].text_frame
    with pytest.raises(rpptx.ReplacementCountError, match="found 1"):
        frame.try_replace_text("Same text", "Other", expect=2)
    assert frame.try_replace_text("Same text", "Other", expect=1) == 1
    runs = prs.slides[0].shapes[0].shapes[0].text_frame.paragraphs[0].runs
    assert [run.text for run in runs] == ["Other", " in a group"]
    assert runs[0].font.bold is True
    assert prs.slides[0].shapes[1].text == "Same text"
    prs.slides[0].shapes[0].shapes[0].text_frame.text = "Same text in a group"
    assert prs.slides[0].try_replace_text("Same text", "Other", expect=2) == 2
    assert prs.slides[0].shapes[0].shapes[0].text == "Other in a group"


def test_import_slide_carries_pictures_links_notes_and_background_from_another_deck(tmp_path):
    import rpptx

    red, blue = _tiny_png(0xDD, 0x20, 0x20), _tiny_png(0x20, 0x20, 0xDD)

    def build_source(deck):
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import XL_CHART_TYPE
        from pptx.oxml import parse_xml
        from pptx.util import Inches

        slide = deck.slides.add_slide(deck.slide_layouts[1])
        slide.shapes.title.text = "Imported title"
        slide.placeholders[1].text = "Body bullet"
        slide.shapes.add_picture(io.BytesIO(red), Inches(6), Inches(4))
        run = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(4), Inches(1)).text_frame.paragraphs[0].add_run()
        run.text = "Visit"
        run.hyperlink.address = "https://example.com/import"
        slide.notes_slide.notes_text_frame.text = "Speaker notes"
        _, background = slide.part.get_or_add_image_part(io.BytesIO(blue))
        slide._element.cSld.insert(0, parse_xml(
            '<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<p:bgPr><a:blipFill><a:blip r:embed="{background}"/></a:blipFill><a:effectLst/></p:bgPr></p:bg>'
        ))
        chart = deck.slides.add_slide(deck.slide_layouts[5])
        data = CategoryChartData()
        data.categories = ["a", "b"]
        data.add_series("s", (1, 2))
        chart.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 0, Inches(4), Inches(3), data)

    def build_target(deck):
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        slide.shapes.add_picture(io.BytesIO(red), 0, 0)

    source = rpptx.Presentation(_python_pptx_deck(tmp_path / "source.pptx", build_source))
    prs = rpptx.Presentation(_python_pptx_deck(tmp_path / "target.pptx", build_target))
    before = prs.to_bytes()
    source_before = source.to_bytes()
    chart_target = rpptx.Presentation()
    existing = chart_target.slides.add_slide(chart_target.slide_layouts[6])
    existing.shapes.add_textbox(0, 0, 914400, 914400).text = "Retained destination"
    existing_xml = existing.element.xml
    existing_parts = _package_parts(chart_target.to_bytes())
    chart_slide = chart_target.slides.import_slide(source.slides[1])
    chart_shapes = [shape for shape in chart_slide.shapes if shape.has_chart]
    assert len(chart_shapes) == 1
    chart = chart_shapes[0].chart
    assert chart.categories == ("a", "b")
    assert chart.series[0].name == "s" and chart.series[0].values == [1.0, 2.0]
    chart_output = tmp_path / "imported-chart.pptx"
    chart_target.save(chart_output)
    source_workbooks = sorted(
        value for name, value in _package_parts(source_before).items()
        if name.startswith("ppt/embeddings/") and name.endswith(".xlsx")
    )
    imported_workbooks = sorted(
        value for name, value in _package_parts(chart_output.read_bytes()).items()
        if name.startswith("ppt/embeddings/") and name.endswith(".xlsx")
    )
    assert len(source_workbooks) == 1 and imported_workbooks == source_workbooks
    imported_parts = _package_parts(chart_output.read_bytes())
    assert existing.element.xml == existing_xml
    assert imported_parts["ppt/slides/slide1.xml"] == existing_parts["ppt/slides/slide1.xml"]
    assert imported_parts["ppt/slides/_rels/slide1.xml.rels"] == existing_parts["ppt/slides/_rels/slide1.xml.rels"]
    reopened = rpptx.Presentation(chart_output)
    assert reopened.slides[0].shapes[0].text == "Retained destination"
    reopened_charts = [shape for shape in reopened.slides[1].shapes if shape.has_chart]
    assert len(reopened_charts) == 1
    reopened_chart = reopened_charts[0].chart
    assert reopened_chart.categories == ("a", "b")
    assert reopened_chart.series[0].values == [1.0, 2.0]
    assert source.to_bytes() == source_before
    assert prs.to_bytes() == before
    with pytest.raises(ValueError, match="layout is not a layout of this presentation"):
        prs.slides.import_slide(source.slides[0], layout=source.slide_layouts[1])
    with pytest.raises(IndexError):
        prs.slides.import_slide(source.slides[0], index=2)

    held = prs.slides[0]
    imported = prs.slides.import_slide(source.slides[0], index=0)
    assert len(held.shapes) == len(prs.slides[1].shapes)
    assert [shape.name for shape in imported.shapes] == [
        "Title 1", "Content Placeholder 2", "Picture 3", "TextBox 4",
    ]
    assert imported.shapes[2].image.blob == red
    assert imported.notes_text == "Speaker notes"
    again = prs.slides.import_slide(prs.slides[0], layout=prs.slide_layouts[6], index=-1)
    assert again.shapes[0].text == "Imported title"
    same = prs.slides.import_slide(prs.slides[0])
    assert same.slide_layout == prs.slide_layouts[1]
    output = tmp_path / "imported.pptx"
    prs.save(output)
    media = [name for name in _package_parts(output.read_bytes()) if name.startswith("ppt/media/")]
    assert len(media) == 2

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output)
    assert [slide.slide_layout.name for slide in oracle.slides] == [
        "Title and Content", "Blank", "Blank", "Title and Content",
    ]
    slide = oracle.slides[0]
    assert [shape.text_frame.text for shape in slide.shapes if shape.has_text_frame] == [
        "Imported title", "Body bullet", "Visit",
    ]
    assert slide.shapes[2].image.blob == red
    assert slide.shapes[3].text_frame.paragraphs[0].runs[0].hyperlink.address == "https://example.com/import"
    assert slide.notes_slide.notes_text_frame.text == "Speaker notes"
    blip = slide._element.cSld.bg.xpath(".//a:blip")[0]
    embed = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
    assert slide.part.related_part(embed).blob == blue



def test_add_shape_writes_the_python_pptx_theme_style_and_add_textbox_none(tmp_path):
    import rpptx
    import xml.etree.ElementTree as ET

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    assert pptx.__version__ == "1.0.2"

    def styles(deck):
        output = tmp_path / "styles.pptx"
        deck.save(output)
        slide_xml = _package_parts(output.read_bytes())["ppt/slides/slide1.xml"]
        root = ET.fromstring(slide_xml)
        p_ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

        def tree(element):
            return (
                element.tag,
                tuple(sorted(element.attrib.items())),
                tuple(tree(child) for child in element),
            )

        return [
            tree(style)
            for shape in root.iter(f"{p_ns}sp")
            if (style := shape.find(f"{p_ns}style")) is not None
        ]

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_shape(1, 914400, 914400, 2743200, 1828800)
    prs.slides[0].shapes.add_textbox(0, 0, 100, 100)
    oracle = pptx.Presentation()
    oracle.slides.add_slide(oracle.slide_layouts[6])
    oracle.slides[0].shapes.add_shape(1, 914400, 914400, 2743200, 1828800)
    oracle.slides[0].shapes.add_textbox(0, 0, 100, 100)

    assert len(styles(prs)) == 1
    assert styles(prs) == styles(oracle)
    shape = prs.slides[0].shapes[0]
    assert (shape.fill.type, shape.line.color.rgb) == (None, None)


def test_add_shape_accepts_preset_names_and_every_mso_shape_member(tmp_path):
    import rpptx
    from rpptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR, MSO_CONNECTOR_TYPE, MSO_SHAPE, MSO_SHAPE_TYPE

    assert MSO_AUTO_SHAPE_TYPE is MSO_SHAPE and MSO_CONNECTOR is MSO_CONNECTOR_TYPE
    assert len(MSO_SHAPE) == 182
    assert (MSO_SHAPE.PENTAGON, MSO_SHAPE.CHEVRON) == (51, 52)
    assert MSO_SHAPE(5) is MSO_SHAPE.ROUNDED_RECTANGLE
    assert MSO_SHAPE.ROUNDED_RECTANGLE.xml_value == "roundRect"
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    for member in MSO_SHAPE:
        prs.slides[0].shapes.add_shape(member, 0, 0, 10, 10)
    prs.slides[0].shapes.add_shape("roundRect", 0, 0, 10, 10)
    prs.slides[0].shapes.add_shape(1, 0, 0, 10, 10)
    with pytest.raises(ValueError, match="unsupported MSO_SHAPE value"):
        prs.slides[0].shapes.add_shape(9999, 0, 0, 10, 10)
    with pytest.raises(rpptx.RpptxError):
        prs.slides[0].shapes.add_shape("notAPreset", 0, 0, 10, 10)
    shapes = prs.slides[0].shapes
    shapes.add_connector(MSO_CONNECTOR.ELBOW, 100, 200, 10, 20)
    prs.slides[0].shapes.add_connector(1, 0, 0, 50, 50)
    with pytest.raises(ValueError, match="unsupported MSO_CONNECTOR value"):
        prs.slides[0].shapes.add_connector(MSO_CONNECTOR.MIXED, 0, 0, 1, 1)
    group = prs.slides[0].shapes.add_group_shape()
    assert (group.shape_type, len(group.shapes)) == (MSO_SHAPE_TYPE.GROUP, 0)
    output = tmp_path / "presets.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle_members = {member.name: member for member in pptx.enum.shapes.MSO_AUTO_SHAPE_TYPE}
    assert set(oracle_members) - {member.name for member in MSO_SHAPE} == set()
    for member in MSO_SHAPE:
        assert (oracle_members[member.name].value, oracle_members[member.name].xml_value) == (
            member.value,
            member.xml_value,
        )
    oracle = pptx.Presentation(output).slides[0].shapes
    def preset(shape):
        return shape._element.xpath("./p:spPr/a:prstGeom/@prst")[0]

    presets = [preset(shape) for shape in list(oracle)[: len(MSO_SHAPE) + 2]]
    assert presets == [member.xml_value for member in MSO_SHAPE] + ["roundRect", "rect"]
    elbow, straight = oracle[len(MSO_SHAPE) + 2], oracle[len(MSO_SHAPE) + 3]
    assert elbow.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.LINE
    assert (elbow.begin_x, elbow.begin_y, elbow.end_x, elbow.end_y) == (100, 200, 10, 20)
    assert preset(straight) == "line"


def test_theme_effect_index_zero_drops_the_connector_theme_shadow(tmp_path):
    import rpptx
    from rpptx.enum.shapes import MSO_CONNECTOR

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    arrow = prs.slides[0].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 914400, 0)
    assert arrow.theme_effect_index == 1
    arrow.theme_effect_index = 0
    assert arrow.theme_effect_index == 0
    assert b'<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' in arrow.xml
    with pytest.raises(OverflowError):
        arrow.theme_effect_index = -1
    with pytest.raises(TypeError):
        arrow.theme_effect_index = None
    box = prs.slides[0].shapes.add_textbox(0, 0, 10, 10)
    assert box.theme_effect_index is None
    with pytest.raises(rpptx.RpptxError, match="no p:style"):
        box.theme_effect_index = 0
    output = tmp_path / "no-shadow.pptx"
    prs.save(output)
    assert rpptx.Presentation(output).slides[0].shapes[0].theme_effect_index == 0

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes[0]
    assert oracle.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.LINE
    assert oracle._element.xpath("./p:style/a:effectRef/@idx") == ["0"]


def test_notes_text_creates_the_notes_slide_on_a_python_pptx_deck(tmp_path):
    import rpptx

    source = _python_pptx_deck(
        tmp_path / "no-notes.pptx",
        lambda deck: [deck.slides.add_slide(deck.slide_layouts[6]) for _ in range(2)],
    )
    assert not [name for name in _package_parts(source.read_bytes()) if "notes" in name]
    prs = rpptx.Presentation(source)
    prs.slides[1].notes_text = "Second slide note"
    assert [slide.notes_text for slide in prs.slides] == [None, "Second slide note"]
    output = tmp_path / "notes.pptx"
    prs.save(output)
    assert len(prs.render_all_notes(dpi=36.0)) == 2

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output)
    assert oracle.slides[0].has_notes_slide is False
    notes_slide = oracle.slides[1].notes_slide
    assert notes_slide.notes_text_frame.text == "Second slide note"
    assert [shape.name for shape in notes_slide.placeholders] == [
        "Slide Image Placeholder 1",
        "Notes Placeholder 2",
        "Slide Number Placeholder 3",
    ]


def test_line_feeds_in_assigned_text_make_paragraphs_as_python_pptx_does(tmp_path):
    assert importlib.metadata.version("python-pptx") == "1.0.2"
    import rpptx
    from rpptx.util import Inches

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    prs.slides[0].shapes.add_table(1, 1, Inches(1), Inches(4), Inches(4), Inches(1))
    prs.slides[0].shapes[0].text_frame.text = "first line\nsecond line\vsoft"
    prs.slides[0].shapes[0].text_frame.add_paragraph().text = "para\nbreak\vtab"
    prs.slides[0].shapes[0].text_frame.add_paragraph().add_run("run\nliteral")
    prs.slides[0].shapes[1].table.cell(0, 0).text = "cell\none"
    prs.slides[0].notes_text = "note\nnext"

    frame = prs.slides[0].shapes[0].text_frame
    assert [paragraph.text for paragraph in frame.paragraphs] == [
        "first line",
        "second line\vsoft",
        "para\vbreak\vtab",
        "run\nliteral",
    ]
    assert prs.slides[0].shapes[1].table.cell(0, 0).text == "cell\none"
    assert prs.slides[0].notes_text == "note\nnext"
    lines = [line.text for line in prs.text_layout()[0].lines]
    assert lines == ["first line", "second line", "soft", "para", "break", "tab", "run", "literal"]
    assert prs.render_slide_to_png(0, dpi=36.0)

    prs.slides[0].shapes[0].text = "shape\ntext"
    paragraphs = prs.slides[0].shapes[0].text_frame.paragraphs
    assert [paragraph.text for paragraph in paragraphs] == ["shape", "text"]
    output = tmp_path / "line-feeds.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output)
    assert [paragraph.text for paragraph in oracle.slides[0].shapes[0].text_frame.paragraphs] == [
        "shape",
        "text",
    ]
    assert [
        paragraph.text for paragraph in oracle.slides[0].shapes[1].table.cell(0, 0).text_frame.paragraphs
    ] == ["cell", "one"]
    expected = pptx.Presentation()
    textbox = expected.slides.add_slide(expected.slide_layouts[6]).shapes.add_textbox(0, 0, 100, 100)
    textbox.text_frame.text = "first line\nsecond line\vsoft"
    assert [paragraph.text for paragraph in textbox.text_frame.paragraphs] == [
        "first line",
        "second line\vsoft",
    ]


def test_shape_xml_is_a_self_contained_element():
    import xml.etree.ElementTree as ElementTree

    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 10, 10).text = "xml"
    data = prs.slides[0].shapes[0].xml
    assert isinstance(data, bytes)
    element = ElementTree.fromstring(data)
    assert element.tag == "{http://schemas.openxmlformats.org/presentationml/2006/main}sp"
    assert b"<a:t>xml</a:t>" in data


def test_adjustments_match_python_pptx_defaults_reads_and_writes(tmp_path):
    import rpptx
    from rpptx.enum.shapes import MSO_SHAPE

    def build(deck):
        pptx = pytest.importorskip("pptx")
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        for member in MSO_SHAPE:
            slide.shapes.add_shape(pptx.enum.shapes.MSO_AUTO_SHAPE_TYPE[member.name], 0, 0, 10, 10)
        arrow = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RIGHT_ARROW, 0, 0, 10, 10)
        arrow.adjustments[1] = 0.25

    source = _python_pptx_deck(tmp_path / "adjustments.pptx", build)
    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    expected = [list(shape.adjustments) for shape in pptx.Presentation(source).slides[0].shapes]
    prs = rpptx.Presentation(source)
    actual = [list(shape.adjustments) for shape in prs.slides[0].shapes]
    differences = {
        member.name: (ours, theirs)
        for member, ours, theirs in zip(MSO_SHAPE, actual, expected)
        if ours != theirs
    }
    assert differences == {
        "FOLDED_CORNER": ([0.16667], []),
        "UP_DOWN_ARROW": ([0.5, 0.5], [0.5, 0.5, 0.5, 0.5]),
    }
    assert actual[-1] == expected[-1] == [0.5, 0.25]

    shapes = prs.slides[0].shapes
    rounded = shapes[list(MSO_SHAPE).index(MSO_SHAPE.ROUNDED_RECTANGLE)]
    adjustments = rounded.adjustments
    assert (len(adjustments), adjustments[0], adjustments[-1]) == (1, 0.16667, 0.16667)
    adjustments[0] = 0.3
    assert rounded.adjustments[0] == 0.3
    with pytest.raises(ValueError, match="adjustment value must be numeric"):
        adjustments[0] = "wide"
    with pytest.raises(IndexError):
        _ = adjustments[1]
    textbox = prs.slides[0].shapes.add_textbox(0, 0, 10, 10)
    assert len(textbox.adjustments) == 0
    with pytest.raises(ValueError, match="shape has no adjustments"):
        _ = prs.slides[0].shapes.add_group_shape().adjustments
    output = tmp_path / "adjustments-out.pptx"
    prs.save(output)
    oracle = pptx.Presentation(output).slides[0].shapes
    assert oracle[list(MSO_SHAPE).index(MSO_SHAPE.ROUNDED_RECTANGLE)].adjustments[0] == 0.3


def test_issue_169_replace_text_alias_is_counted_and_atomic():
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 100_000, 100_000).text = "Original"
    before = prs.to_bytes()

    with pytest.raises(rpptx.ReplacementCountError):
        prs.replace_text("Original", "Updated", expect=2)
    assert prs.to_bytes() == before

    assert prs.replace_text("Original", "Updated", expect=1) == 1
    reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
    assert reopened.slides[0].shapes[0].text == "Updated"


def test_issue_169_effective_geometry_materializes_missing_partners():
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    assert title.left is None
    resolved = title.effective_geometry()
    assert resolved is not None
    left, top, width, height = (int(value) for value in resolved)
    assert width > 0 and height > 0

    title.left = left + 100
    materialized = prs.slides[0].shapes.title
    assert tuple(int(value) for value in (materialized.left, materialized.top,
                                          materialized.width, materialized.height)) == (
        left + 100, top, width, height
    )


def test_issue_169_shape_click_hyperlink_round_trips_and_python_pptx_reads_it(tmp_path):
    import pptx
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 100_000, 100_000)
    shape.text = "Open"
    shape = prs.slides[0].shapes[0]
    assert shape.click_action.hyperlink.address is None
    shape.click_action.hyperlink.address = "https://example.com/first"
    output = tmp_path / "shape-link.pptx"
    prs.save(output)
    assert pptx.Presentation(output).slides[0].shapes[0].click_action.hyperlink.address == (
        "https://example.com/first"
    )
    reopened = rpptx.Presentation(output)
    link = reopened.slides[0].shapes[0].click_action.hyperlink
    assert link.address == "https://example.com/first"
    link.address = "https://example.com/second"
    output = tmp_path / "shape-link-retargeted.pptx"
    reopened.save(output)
    assert pptx.Presentation(output).slides[0].shapes[0].click_action.hyperlink.address == (
        "https://example.com/second"
    )
    link.address = None
    assert reopened.slides[0].shapes[0].click_action.hyperlink.address is None


def test_issue_169_modern_comments_anchor_to_shape_and_text_range(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 100_000, 100_000)
    shape.text = "Hello world"
    shape_id = prs.slides[0].shapes[0].shape_id
    author_id = "{11111111-1111-1111-1111-111111111111}"
    prs.add_comment_author(
        id=author_id, name="Ada", user_id="ada@example.test", provider_id="local"
    )
    kwargs = dict(author_id=author_id, created="2026-09-30T10:00:00Z", text="Review")
    slide = prs.slides[0]
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="text range exceeds"):
        slide.add_comment(id="{22222222-2222-2222-2222-222222222222}",
                          shape_id=shape_id, text_start=10, text_length=4, **kwargs)
    assert prs.to_bytes() == before
    slide.add_comment(id="{22222222-2222-2222-2222-222222222222}",
                      shape_id=shape_id, **kwargs)
    prs.slides[0].add_comment(id="{33333333-3333-3333-3333-333333333333}",
                              shape_id=shape_id, text_start=6, text_length=5, **kwargs)
    output = tmp_path / "anchored-comments.pptx"
    prs.save(output)
    with zipfile.ZipFile(output) as archive:
        comments = archive.read("ppt/comments/comment1.xml")
    assert b"<ac:deMkLst" in comments
    assert b"<ac:txMkLst" in comments
    assert f'<ac:spMk id="{shape_id}"/>'.encode() in comments
    assert b'<ac:txMk cp="6" len="5"><ac:context len="11" hash="' in comments
    reopened = rpptx.Presentation(output)
    assert [comment.id for comment in reopened.slides[0].comments] == [
        "{22222222-2222-2222-2222-222222222222}",
        "{33333333-3333-3333-3333-333333333333}",
    ]


def test_issue_169_complete_production_deck_chain(tmp_path):
    """Issue 169's operations must coexist in one saved, readable deck."""
    import os
    import posixpath
    import shutil
    import subprocess
    import xml.etree.ElementTree as ET

    import pptx
    import rpptx
    from PIL import Image
    from rpptx.enum.shapes import MSO_SHAPE

    assert importlib.metadata.version("python-pptx") == "1.0.2"
    emu = rpptx.Inches(1)
    source = rpptx.Presentation()
    source.slides.add_slide(source.slide_layouts[6])
    source.slides[0].shapes.add_textbox(emu, emu, 3 * emu, emu).text = "Imported caption"
    source.slides[0].notes_text = "Imported speaker note"
    source.slides[0].shapes.add_picture(_tiny_png(), 5 * emu, emu, emu, emu)

    deck = rpptx.Presentation()
    deck.slides.add_slide(deck.slide_layouts[0])
    title = deck.slides[0].shapes.title
    assert title.left is None and title.effective_geometry()[2] > 0
    title.left = int(title.effective_geometry()[0]) + 100
    deck.slides[0].shapes.title.text = "Production checklist"
    deck.slides[0].placeholders[1].text = "Issue 169"
    deck.slides[0].notes_text = "Opening note"
    deck.slides.add_slide(deck.slide_layouts[6])
    deck.slides[1].shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, emu, emu, 3 * emu, emu)
    shape = deck.slides[1].shapes[0]
    shape.name = "Front card"
    shape.text = "Replace this label"
    shape = deck.slides[1].shapes[0]
    shape.click_action.hyperlink.address = "https://example.com/checklist"
    deck.slides[1].shapes.add_picture(_tiny_png(0xE0, 0x30, 0x30), emu, 3 * emu, 2 * emu, emu)
    deck.slides[1].shapes[1].crop_left = 0.25
    deck.slides[1].shapes.add_table(2, 2, 5 * emu, emu, 3 * emu, 2 * emu)
    table = deck.slides[1].shapes[2].table
    for row, values in enumerate((("Item", "State"), ("crop", "ready"))):
        for column, value in enumerate(values):
            table.cell(row, column).text = value
    table.columns[0].width = emu
    table.rows[1].height = emu
    table.cell(1, 0).text = "image"
    deck.slides[1].shapes.add_group_shape()
    deck.slides[1].shapes[3].shapes.add_textbox(5 * emu, 4 * emu, 2 * emu, emu).text = "Grouped"
    deck.slides[1].shapes.move(0, 2)
    assert [shape.name for shape in deck.slides[1].shapes][2] == "Front card"
    author_id = "{11111111-1111-1111-1111-111111111111}"
    deck.add_comment_author(id=author_id, name="Ada", user_id="ada@example.test", provider_id="local")
    card_id = deck.slides[1].shapes[2].shape_id
    deck.slides[1].add_comment(
        id="{22222222-2222-2222-2222-222222222222}", author_id=author_id,
        created="2026-09-30T10:00:00Z", text="Review label",
        shape_id=card_id, text_start=0, text_length=7,
    )
    before = deck.to_bytes()
    with pytest.raises(rpptx.ReplacementCountError):
        deck.slides[1].try_replace_text("Replace this label", "Reviewed label", expect=2)
    assert deck.to_bytes() == before
    assert deck.slides[1].try_replace_text("Replace this label", "Reviewed label", expect=1) == 1
    assert deck.slides[1].shapes[2].text == "Reviewed label"
    imported = deck.slides.import_slide(source.slides[0], index=1)
    assert imported.notes_text == "Imported speaker note"
    assert imported.shapes[1].image.blob == _tiny_png()
    assert deck.validate() == ()

    # The Python Table binding has no style selector. Set the built-in GUID
    # through the documented ZIP/XML package fallback, then reopen in rpptx.
    style_id = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    parts = _package_parts(deck.to_bytes())
    unchanged_parts = parts.copy()
    styled_part = next(
        name for name, data in parts.items()
        if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        and b"Reviewed label" in data
    )
    slide_xml = parts[styled_part]
    original = b'<a:tblPr firstRow="1" bandRow="1"/>'
    assert slide_xml.count(original) == 1
    parts[styled_part] = slide_xml.replace(
        original,
        b'<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>'
        + style_id.encode() + b'</a:tableStyleId></a:tblPr>',
    )
    assert {name: data for name, data in parts.items() if name != styled_part} == {
        name: data for name, data in unchanged_parts.items() if name != styled_part
    }
    output = tmp_path / "issue-169-production.pptx"
    output.write_bytes(_package_bytes(parts))
    assert _package_parts(output.read_bytes()) == parts
    reopened = rpptx.Presentation(output)
    assert reopened.validate() == ()
    assert [slide.slide_layout.name for slide in reopened.slides] == [
        "Title Slide", "Blank", "Blank",
    ]
    assert reopened.slides[0].shapes.title.text == "Production checklist"
    assert reopened.slides[1].notes_text == "Imported speaker note"
    assert reopened.slides[2].shapes[2].text == "Reviewed label"
    assert reopened.slides[2].shapes[2].click_action.hyperlink.address == "https://example.com/checklist"
    assert reopened.slides[2].shapes[0].crop_left == 0.25
    assert reopened.slides[2].shapes[1].table.cell(1, 0).text == "image"
    assert reopened.slides[2].shapes[3].shapes[0].text == "Grouped"
    assert len(reopened.slides[2].comments) == 1
    assert reopened.render_slide_to_png(2, dpi=72.0).startswith(b"\x89PNG")

    # Every internal package target must resolve, including notes, media and comments.
    with zipfile.ZipFile(output) as archive:
        names = set(archive.namelist())
        assert "ppt/comments/comment1.xml" in names
        assert any(name.startswith("ppt/media/") for name in names)
        for name in names:
            if not name.endswith(".rels"):
                continue
            base = "" if name == "_rels/.rels" else name.rsplit("/_rels/", 1)[0]
            for rel in ET.fromstring(archive.read(name)):
                if rel.get("TargetMode") == "External":
                    continue
                target = posixpath.normpath(posixpath.join(base, rel.attrib["Target"].lstrip("/")))
                if rel.attrib["Target"].startswith("/"):
                    target = rel.attrib["Target"].lstrip("/")
                assert target in names, (name, target)

    oracle = pptx.Presentation(output)
    assert len(oracle.slides) == 3
    assert oracle.slides[0].shapes.title.text == "Production checklist"
    assert oracle.slides[1].notes_slide.notes_text_frame.text == "Imported speaker note"
    shapes = oracle.slides[2].shapes
    assert [shape.name for shape in shapes][2] == "Front card"
    assert shapes[0].crop_left == 0.25
    assert shapes[1].table.cell(1, 0).text == "image"
    assert shapes[2].click_action.hyperlink.address == "https://example.com/checklist"
    assert shapes[3].shapes[0].text == "Grouped"
    assert style_id.encode() in parts[styled_part]

    soffice = os.environ.get("RPPTX_PINNED_SOFFICE") or shutil.which("soffice")
    if soffice:
        assert "LibreOffice 26.2.5.2" in subprocess.check_output([soffice, "--version"], text=True)
        pdf_dir = tmp_path / "lo"
        pdf_dir.mkdir()
        subprocess.run([
            soffice, "-env:UserInstallation=" + (tmp_path / "lo-profile").as_uri(),
            "--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(output),
        ], check=True, capture_output=True, text=True)
        pdf = pdf_dir / (output.stem + ".pdf")
        assert pdf.is_file()
        subprocess.run([
            "pdftoppm", "-f", "3", "-l", "3", "-r", "72", "-png", "-singlefile",
            str(pdf), str(pdf_dir / "slide3"),
        ], check=True, capture_output=True)
        native = Image.open(io.BytesIO(reopened.render_slide_to_png(2, dpi=72.0))).convert("RGB")
        native.save(pdf_dir / "native-slide3.png")
        viewer = Image.open(pdf_dir / "slide3.png").convert("RGB")
        assert abs(native.width - viewer.width) <= 1 and native.height == viewer.height
        viewer = viewer.crop((0, 0, native.width, native.height))
        assert native.getpixel((390, 95)) == viewer.getpixel((390, 95)) == (0x4F, 0x81, 0xBD)
        native_bytes, viewer_bytes = native.tobytes(), viewer.tobytes()
        errors = [abs(left - right) for left, right in zip(native_bytes, viewer_bytes)]
        pixels = native.width * native.height
        close = sum(max(errors[index:index + 3]) <= 24 for index in range(0, len(errors), 3))
        mean_error = sum(errors) / len(errors)
        assert close / pixels >= 0.97 and mean_error <= 2.0, (close / pixels, mean_error)


def test_issue_217_complete_deck_chain(tmp_path):
    """Issue 217's six authoring items survive one package and both readers."""
    import hashlib
    import os
    import posixpath
    import shutil
    import subprocess
    import xml.etree.ElementTree as ET
    from pathlib import Path

    import pptx
    import rpptx
    from PIL import Image
    from rpptx.dml.color import RGBColor
    from rpptx.enum.dml import MSO_ARROWHEAD_LENGTH, MSO_ARROWHEAD_STYLE, MSO_ARROWHEAD_WIDTH
    from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE

    assert importlib.metadata.version("python-pptx") == "1.0.2"
    unit = rpptx.Inches(1)
    source = rpptx.Presentation()
    source.slides.add_slide(source.slide_layouts[6])
    source.slides[0].shapes.add_textbox(unit, unit, 3 * unit, unit).text = "Imported TOKEN"
    source.slides[0].shapes.add_picture(_tiny_png(0x22, 0x88, 0xCC), 5 * unit, unit, unit, unit)
    source.slides[0].notes_text = "Imported notes"

    deck = rpptx.Presentation()
    deck.slides.add_slide(deck.slide_layouts[6])
    deck.slides[0].shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, unit, unit, 2 * unit, unit)
    card = deck.slides[0].shapes[0]
    card.text = "Card TOKEN"
    card = deck.slides[0].shapes[0]
    # Compare shadow paint without unrelated endpoint-gradient interpolation.
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0x4F, 0x81, 0xBD)
    card.shadow.color.rgb = RGBColor(0x12, 0x34, 0x56)
    card.shadow.alpha = 0.5
    card.shadow.blur_radius = rpptx.Pt(6)
    card.shadow.distance = rpptx.Pt(4)
    card.shadow.direction = 45.0
    deck.slides[0].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 4 * unit, unit, 7 * unit, unit)
    connector = deck.slides[0].shapes[1]
    connector.theme_effect_index = 0
    connector.line.tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE
    connector.line.tail_end.width = MSO_ARROWHEAD_WIDTH.WIDE
    connector.line.tail_end.length = MSO_ARROWHEAD_LENGTH.LONG
    deck.slides[0].shapes.add_shape(MSO_SHAPE.RECTANGLE, unit, 3 * unit, 2 * unit, unit)
    deck.slides[0].shapes[2].auto_shape_type = MSO_SHAPE.CHEVRON
    deck.slides[0].shapes[2].text = "Preset"
    # Isolate preset geometry from inherited endpoint-gradient interpolation.
    deck.slides[0].shapes[2].fill.solid()
    deck.slides[0].shapes[2].fill.fore_color.rgb = RGBColor(0x4F, 0x81, 0xBD)
    imported = deck.slides.import_slide(source.slides[0], index=1)
    assert imported.notes_text == "Imported notes"
    assert imported.shapes[1].image.blob == _tiny_png(0x22, 0x88, 0xCC)

    before = deck.to_bytes()
    with pytest.raises(rpptx.ReplacementCountError):
        deck.slides[0].try_replace_text("TOKEN", "ready", expect=2)
    assert deck.to_bytes() == before
    assert deck.slides[0].shapes[0].text_frame.try_replace_text("TOKEN", "ready", expect=1) == 1
    assert deck.slides[1].try_replace_text("TOKEN", "ready", expect=1, notes=False) == 1
    output = tmp_path / "issue-217-chain.pptx"
    deck.save(output)
    parts = _package_parts(output.read_bytes())
    media_parts = [data for name, data in parts.items() if name.startswith("ppt/media/")]
    assert [hashlib.sha256(data).hexdigest() for data in media_parts] == [
        hashlib.sha256(_tiny_png(0x22, 0x88, 0xCC)).hexdigest()
    ]
    with zipfile.ZipFile(output) as archive:
        names = set(archive.namelist())
        assert any(name.startswith("ppt/media/") for name in names)
        assert any(name.startswith("ppt/notesSlides/") for name in names)
        for name in names:
            if not name.endswith(".rels"):
                continue
            base = "" if name == "_rels/.rels" else name.rsplit("/_rels/", 1)[0]
            for rel in ET.fromstring(archive.read(name)):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.attrib["Target"]
                resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
                assert resolved in names, (name, resolved)

    reopened = rpptx.Presentation(output)
    assert reopened.validate() == ()
    assert len(reopened.slides) == 2
    assert reopened.slides[0].shapes[0].shadow.color.rgb == RGBColor(0x12, 0x34, 0x56)
    assert reopened.slides[0].shapes[1].theme_effect_index == 0
    assert reopened.slides[0].shapes[1].line.tail_end.type == MSO_ARROWHEAD_STYLE.TRIANGLE
    assert reopened.slides[0].shapes[2].auto_shape_type == MSO_SHAPE.CHEVRON
    assert reopened.slides[0].shapes[0].text == "Card ready"
    assert reopened.slides[1].shapes[0].text == "Imported ready"
    assert reopened.slides[1].notes_text == "Imported notes"
    assert reopened.slides[1].shapes[1].image.blob == _tiny_png(0x22, 0x88, 0xCC)

    oracle = pptx.Presentation(output)
    assert len(oracle.slides) == 2
    card, connector, preset = oracle.slides[0].shapes
    assert card.shadow.inherit is False and card.text == "Card ready"
    assert card._element.xpath("./p:spPr/a:effectLst/a:outerShdw/@blurRad") == ["76200"]
    assert connector._element.xpath("./p:style/a:effectRef/@idx") == ["0"]
    assert connector._element.xpath("./p:spPr/a:ln/a:tailEnd/@type") == ["triangle"]
    assert preset.auto_shape_type == pptx.enum.shapes.MSO_SHAPE.CHEVRON
    assert oracle.slides[1].shapes[0].text == "Imported ready"
    assert oracle.slides[1].shapes[1].image.blob == _tiny_png(0x22, 0x88, 0xCC)
    assert oracle.slides[1].notes_slide.notes_text_frame.text == "Imported notes"

    root = Path(__file__).resolve().parents[3]
    subprocess.run(["cargo", "run", "--quiet", "-p", "rpptx-cli", "--", "validate", str(output)],
                   cwd=root, check=True, capture_output=True, text=True)
    soffice = os.environ.get("RPPTX_PINNED_SOFFICE") or shutil.which("soffice")
    if not soffice:
        pytest.skip("pinned LibreOffice viewer oracle is unavailable")
    assert "LibreOffice 26.2.5.2" in subprocess.check_output([soffice, "--version"], text=True)
    pdf_dir = tmp_path / "lo"
    pdf_dir.mkdir()
    subprocess.run([
        soffice, "-env:UserInstallation=" + (tmp_path / "lo-profile").as_uri(),
        "--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(output),
    ], check=True, capture_output=True, text=True)
    pdf = pdf_dir / "issue-217-chain.pdf"
    assert pdf.is_file()
    # At 72 DPI these windows isolate the six checklist items. The card
    # covers both shadow and frame replacement, the connector covers its
    # effect reference and line end, and slide two covers import and slide
    # replacement. Text antialiasing accounts for the 24-level allowance.
    windows = (
        (("outer shadow", (60, 60, 240, 170)),
         ("connector theme effect", (275, 55, 525, 165)),
         ("line end", (275, 55, 525, 165)),
         ("preset geometry", (60, 200, 245, 310)),
         ("text-frame replacement", (60, 60, 240, 170))),
        (("cross-deck import", (350, 60, 450, 170)),
         ("slide replacement", (60, 60, 300, 170))),
    )
    for index, slide_windows in enumerate(windows):
        subprocess.run([
            "pdftoppm", "-f", str(index + 1), "-l", str(index + 1), "-r", "72",
            "-png", "-singlefile", str(pdf), str(pdf_dir / f"slide{index + 1}"),
        ], check=True, capture_output=True)
        native = Image.open(io.BytesIO(reopened.render_slide_to_png(index, dpi=72.0))).convert("RGB")
        viewer = Image.open(pdf_dir / f"slide{index + 1}.png").convert("RGB")
        assert abs(native.width - viewer.width) <= 1 and native.height == viewer.height
        viewer = viewer.crop((0, 0, native.width, native.height))
        for label, box in slide_windows:
            native_bytes = native.crop(box).tobytes()
            viewer_bytes = viewer.crop(box).tobytes()
            errors = [abs(left - right) for left, right in zip(native_bytes, viewer_bytes)]
            close = sum(max(errors[offset:offset + 3]) <= 24 for offset in range(0, len(errors), 3))
            assert close / (len(errors) / 3) >= 0.97, label
            assert sum(errors) / len(errors) <= 3.0, label


def test_issue_158_deck_fixture_acceptance(tmp_path):
    """The reporter's exported deck must survive its complete edit workflow."""
    import hashlib
    import os
    import posixpath
    import shutil
    import subprocess
    import urllib.request
    import xml.etree.ElementTree as ET
    from pathlib import Path

    import pptx
    import rpptx
    from PIL import Image
    from rpptx.dml.color import RGBColor
    from rpptx.enum.dml import MSO_ARROWHEAD_LENGTH, MSO_ARROWHEAD_STYLE, MSO_ARROWHEAD_WIDTH
    from rpptx.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE

    root = Path(__file__).resolve().parents[3]
    fixture = Path(os.environ.get(
        "RPPTX_ISSUE_158_FIXTURE", root / "corpus/issue-158/fixture-deck.pptx"
    ))
    if not fixture.is_file() and "RPPTX_ISSUE_158_FIXTURE" not in os.environ:
        fixture.parent.mkdir(parents=True, exist_ok=True)
        request = urllib.request.Request(
            "https://github.com/user-attachments/files/32701103/fixture-deck.pptx",
            headers={"User-Agent": "rdocx-issue-158-acceptance/1"},
        )
        with urllib.request.urlopen(request, timeout=60) as response:
            fixture.write_bytes(response.read())
    source_bytes = fixture.read_bytes()
    assert hashlib.sha256(source_bytes).hexdigest() == (
        "8b703c862792470d3732c6eea07d280d3023525f8653cee4c12d9fc9d14c464a"
    )
    assert importlib.metadata.version("python-pptx") == "1.0.2"
    source_oracle = pptx.Presentation(fixture)
    assert len(source_oracle.slides) == 7
    assert (source_oracle.slide_width, source_oracle.slide_height) == (9144000, 6858000)

    # Add two preservation sentinels to the reporter's actual package. Neither
    # feature appears in the source deck, but both must survive the same edits.
    parts = _package_parts(source_bytes)
    slide7 = parts["ppt/slides/slide7.xml"]
    # This sentinel checks omitted-angle source preservation and orientation.
    # Three non-mirrored stops use piecewise-linear interpolation in both
    # viewers. PowerPoint endpoint gamma has a separate native paint control.
    gradient = (
        b'<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst>'
        b'<a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>'
        b'<a:gs pos="50000"><a:srgbClr val="F0F5FA"/></a:gs>'
        b'<a:gs pos="100000"><a:srgbClr val="E0EAF4"/></a:gs>'
        b'</a:gsLst><a:lin scaled="0"/></a:gradFill><a:effectLst/>'
        b'</p:bgPr></p:bg>'
    )
    body_property = (
        b'<a:bodyPr rot="5400000" vertOverflow="clip" horzOverflow="clip" '
        b'numCol="2" spcCol="91440" rtlCol="1" fromWordArt="1" '
        b'anchorCtr="0" forceAA="1" upright="1" compatLnSpc="0">'
        b'<a:prstTxWarp prst="textNoShape"><a:avLst/>'
        b'</a:prstTxWarp></a:bodyPr>'
    )
    assert slide7.count(b"<p:cSld><p:spTree>") == 1
    assert b"<a:bodyPr/>" in slide7
    slide7 = slide7.replace(b"<p:cSld><p:spTree>", b"<p:cSld>" + gradient + b"<p:spTree>", 1)
    parts["ppt/slides/slide7.xml"] = slide7.replace(b"<a:bodyPr/>", body_property, 1)
    source = tmp_path / "issue-158-source.pptx"
    source.write_bytes(_package_bytes(parts))

    deck = rpptx.Presentation(source)
    assert len(deck.slides) == 7 and deck.validate() == ()
    unit = rpptx.Inches(1)

    def text_shape():
        return next(shape for shape in deck.slides[1].shapes
                    if shape.has_text_frame and shape.text.strip())

    run = text_shape().text_frame.paragraphs[0].runs[0]
    original_font = (run.font.name, run.font.size, run.font.bold)
    run.text += " edited"
    assert (text_shape().text_frame.paragraphs[0].runs[0].font.name,
            text_shape().text_frame.paragraphs[0].runs[0].font.size,
            text_shape().text_frame.paragraphs[0].runs[0].font.bold) == original_font
    text_shape().text_frame.paragraphs[0].space_after = 6 * 12700
    text_shape().text_frame.paragraphs[0].line_spacing = 0.9
    text_shape().text_frame.margin_left = 0
    text_shape().text_frame.word_wrap = True
    assert text_shape().text_frame.paragraphs[0].space_after == 6 * 12700
    original_left, _, original_width, _ = text_shape().effective_geometry()
    text_shape().left = original_left + unit // 10
    text_shape().width = original_width - unit // 10
    assert (text_shape().left, text_shape().width) == (
        original_left + unit // 10, original_width - unit // 10
    )

    shape = deck.slides[1].shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, unit, unit, 2 * unit, unit // 2
    )
    card_id = shape.shape_id
    shape.fill.solid()
    deck.slides[1].shapes[-1].fill.fore_color.rgb = RGBColor(0xDD, 0xEE, 0xFF)
    deck.slides[1].shapes[-1].line.width = 12700
    deck.slides[1].shapes[-1].line.color.rgb = RGBColor(0, 0, 0)
    connector = deck.slides[1].shapes.add_connector(
        MSO_CONNECTOR_TYPE.STRAIGHT, unit, unit, 2 * unit, 2 * unit
    )
    connector_id = connector.shape_id
    picture = Image.new("RGB", (300, 150), (200, 60, 60))
    picture_bytes = io.BytesIO()
    picture.save(picture_bytes, format="PNG")
    image_data = picture_bytes.getvalue()
    deck.slides[1].shapes.add_picture(io.BytesIO(image_data), unit, 3 * unit, width=unit)
    picture_id = deck.slides[1].shapes[-1].shape_id
    next(shape for shape in deck.slides[1].shapes if shape.shape_id == picture_id).replace_image(
        io.BytesIO(image_data)
    )

    deck.slides[2].notes_text = "Edited note."
    deck.slides[3].hidden = True
    deck.slides.move(1, 2)
    deck.slides.add_slide(deck.slide_layouts[1])
    deck.slides.remove(deck.slides[-1])
    assert len(deck.slides) == 7
    deck.slides.duplicate(deck.slides[1])
    assert len(deck.slides) == 8
    assert deck.replace_text("edited", "EDITED") >= 1
    card = next(shape for shape in deck.slides[3].shapes if shape.shape_id == card_id)
    card.shadow.color.rgb = RGBColor(0x12, 0x34, 0x56)
    card.shadow.alpha = 0.5
    card.shadow.blur_radius = rpptx.Pt(6)
    card.shadow.distance = rpptx.Pt(4)
    card.shadow.direction = 45.0
    connector = next(shape for shape in deck.slides[3].shapes
                     if shape.shape_id == connector_id)
    connector.theme_effect_index = 0
    connector.line.tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE
    connector.line.tail_end.width = MSO_ARROWHEAD_WIDTH.WIDE
    connector.line.tail_end.length = MSO_ARROWHEAD_LENGTH.LONG
    preset = deck.slides[3].shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 5 * unit, 5 * unit, 2 * unit, unit // 2
    )
    preset_id = preset.shape_id
    next(shape for shape in deck.slides[3].shapes
         if shape.shape_id == preset_id).auto_shape_type = MSO_SHAPE.CHEVRON
    next(shape for shape in deck.slides[3].shapes
         if shape.shape_id == card_id).text = "Fixture TOKEN\nSecond line\vsoft break"
    assert next(shape for shape in deck.slides[3].shapes if shape.shape_id == card_id
                ).text_frame.try_replace_text("TOKEN", "ready", expect=1) == 1
    import_deck = rpptx.Presentation()
    imported_source = import_deck.slides.add_slide(import_deck.slide_layouts[6])
    imported_source.shapes.add_textbox(unit, unit, 3 * unit, unit).text = "Imported TOKEN"
    import_deck.slides[0].shapes.add_picture(
        io.BytesIO(_tiny_png(0x22, 0x88, 0xCC)), 5 * unit, unit, unit, unit
    )
    import_deck.slides[0].notes_text = "Imported fixture note"
    deck.slides.import_slide(import_deck.slides[0], index=8)
    assert deck.slides[8].try_replace_text("TOKEN", "ready", expect=1, notes=False) == 1
    assert len(deck.slides) == 9
    author_id = "{11111111-1111-1111-1111-111111111158}"
    comment_id = "{22222222-2222-2222-2222-222222222158}"
    deck.add_comment_author(
        id=author_id, name="Fixture reviewer", user_id="reviewer@example.test",
        provider_id="local",
    )
    deck.slides[1].add_comment(
        id=comment_id, author_id=author_id, created="2026-10-02T10:00:00Z",
        text="Reviewed fixture workflow",
    )
    deck.slides[1].resolve_comment(comment_id)
    assert deck.slides[1].comments[0].status == "resolved"
    assert isinstance(deck.text_layout(), tuple)

    output = tmp_path / "issue-158-edited.pptx"
    deck.save(output)
    subprocess.run(
        ["cargo", "run", "--quiet", "-p", "rpptx-cli", "--", "validate", str(output)],
        cwd=root, check=True, capture_output=True, text=True,
    )
    saved = _package_parts(output.read_bytes())
    assert gradient in saved["ppt/slides/slide7.xml"]
    assert body_property in saved["ppt/slides/slide7.xml"]
    assert [hashlib.sha256(data).hexdigest() for name, data in parts.items()
            if name.startswith("ppt/media/")] == [
        hashlib.sha256(saved[name]).hexdigest() for name in parts
        if name.startswith("ppt/media/")
    ]
    with zipfile.ZipFile(output) as archive:
        names = set(archive.namelist())
        for name in names:
            if not name.endswith(".rels"):
                continue
            base = "" if name == "_rels/.rels" else name.rsplit("/_rels/", 1)[0]
            for relationship in ET.fromstring(archive.read(name)):
                if relationship.get("TargetMode") == "External":
                    continue
                target = relationship.attrib["Target"]
                resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
                    posixpath.join(base, target)
                )
                assert resolved in names, (name, resolved)

    reopened = rpptx.Presentation(output)
    oracle = pptx.Presentation(output)
    assert reopened.validate() == ()
    assert len(reopened.slides) == len(oracle.slides) == 9
    assert [slide.notes_text for slide in reopened.slides] == [
        slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else None
        for slide in oracle.slides
    ]
    assert reopened.slides[1].comments[0].status == "resolved"
    assert sum(slide.hidden for slide in reopened.slides) == 1
    assert reopened.slides[3].shapes[-1].auto_shape_type == MSO_SHAPE.CHEVRON
    assert next(shape for shape in reopened.slides[3].shapes
                if shape.shape_id == card_id).shadow.color.rgb == RGBColor(0x12, 0x34, 0x56)
    assert next(shape for shape in reopened.slides[3].shapes
                if shape.shape_id == connector_id).line.tail_end.type == MSO_ARROWHEAD_STYLE.TRIANGLE
    assert reopened.slides[8].notes_text == "Imported fixture note"
    assert reopened.slides[8].shapes[1].image.blob == _tiny_png(0x22, 0x88, 0xCC)
    oracle_card = next(shape for shape in oracle.slides[3].shapes
                       if shape.shape_id == card_id)
    oracle_connector = next(shape for shape in oracle.slides[3].shapes
                            if shape.shape_id == connector_id)
    assert oracle_card.text == "Fixture ready\nSecond line\vsoft break"
    assert oracle_card._element.xpath("./p:spPr/a:effectLst/a:outerShdw/@blurRad") == ["76200"]
    assert oracle_connector._element.xpath("./p:style/a:effectRef/@idx") == ["0"]
    assert oracle_connector._element.xpath("./p:spPr/a:ln/a:tailEnd/@type") == ["triangle"]
    assert oracle.slides[3].shapes[-1].auto_shape_type == pptx.enum.shapes.MSO_SHAPE.CHEVRON
    assert oracle.slides[8].notes_slide.notes_text_frame.text == "Imported fixture note"
    assert any(shape.image.blob == image_data for slide in oracle.slides
               for shape in slide.shapes if shape.shape_type == 13)
    assert any("EDITED" in shape.text for slide in oracle.slides for shape in slide.shapes
               if shape.has_text_frame)

    soffice = os.environ.get("RPPTX_PINNED_SOFFICE") or shutil.which("soffice")
    assert soffice, "pinned LibreOffice viewer oracle is required for Issue 158"
    assert "LibreOffice 26.2.5.2" in subprocess.check_output([soffice, "--version"], text=True)
    viewer_dir = tmp_path / "viewer"
    viewer_dir.mkdir()
    subprocess.run([
        soffice, "-env:UserInstallation=" + (tmp_path / "lo-profile").as_uri(),
        "--headless", "--convert-to", "pdf", "--outdir", str(viewer_dir), str(output),
    ], check=True, capture_output=True, text=True)
    pdf = viewer_dir / "issue-158-edited.pdf"
    assert pdf.is_file()
    # Impress omits the hidden slide from its PDF. These windows cover the
    # imported source images, the edited slide, and the omitted-angle gradient.
    # Picture resampling permits 90 percent close pixels and 4.5 mean error.
    # The other windows permit text antialiasing within 3 RGB levels on average.
    windows = (
        (2, 3, "source images", (430, 100, 600, 465), 0.90, 4.5),
        (3, 4, "edited slide", (0, 0, 720, 540), 0.96, 3.0),
        (7, 7, "gradient background", (0, 350, 720, 540), 0.99, 1.0),
    )
    for slide_index, page_number, label, box, min_close, max_mean in windows:
        subprocess.run([
            "pdftoppm", "-f", str(page_number), "-l", str(page_number),
            "-r", "72", "-png", "-singlefile", str(pdf),
            str(viewer_dir / f"slide{slide_index + 1}"),
        ], check=True, capture_output=True)
        native = Image.open(io.BytesIO(reopened.render_slide_to_png(slide_index, dpi=72.0))).convert("RGB")
        viewer = Image.open(viewer_dir / f"slide{slide_index + 1}.png").convert("RGB")
        assert abs(native.width - viewer.width) <= 1 and native.height == viewer.height
        native_bytes = native.crop(box).tobytes()
        viewer_bytes = viewer.crop(box).tobytes()
        errors = [abs(left - right) for left, right in zip(native_bytes, viewer_bytes)]
        close = sum(max(errors[offset:offset + 3]) <= 24 for offset in range(0, len(errors), 3))
        assert close / (len(errors) / 3) >= min_close, label
        assert sum(errors) / len(errors) <= max_mean, label


def test_table_cells_expose_the_same_editable_text_frame_surface_as_shapes():
    from rpptx import Presentation
    from rpptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(2, 2, 0, 0, Inches(4), Inches(2)).table
    cell = table.cell(0, 0)
    frame = cell.text_frame
    frame.text = "Cell text"
    assert cell.text == "Cell text"
    assert len(cell.text_frame.paragraphs) == 1


def _f_x194_table():
    from rpptx import Presentation, Inches
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    return prs, slide.shapes.add_table(2, 2, 0, 0, Inches(4), Inches(2)).table


def test_cell_text_frame_structural_edits_stale_every_earlier_nested_handle():
    from rpptx import StaleElementError
    prs, table = _f_x194_table()
    cell = table.cell(0, 0)
    cell.text = "seed"
    frame = cell.text_frame
    paragraphs = frame.paragraphs
    paragraph = paragraphs[0]
    runs = paragraph.runs
    run = runs[0]
    font = run.font
    other = table.cell(1, 1).text_frame
    cell.text = "changed"
    for operation in (lambda: len(paragraphs), lambda: paragraph.text,
                      lambda: len(runs), lambda: run.text, lambda: font.bold):
        with pytest.raises(StaleElementError, match=r"First invalidated by Cell.text "):
            operation()
    assert cell.text == frame.text == "changed"
    assert other.text == ""
    assert len(table.rows) == 2
    frame = cell.text_frame
    new_paragraph = frame.add_paragraph()
    assert new_paragraph.text == ""
    assert frame.text == cell.text
    fresh_frame = cell.text_frame
    retired_runs = new_paragraph.runs
    new_run = new_paragraph.add_run("new run")
    assert new_run.text == "new run"
    assert new_paragraph.text == "new run"
    with pytest.raises(StaleElementError, match=r"First invalidated by Paragraph.add_run at committed revision"):
        len(retired_runs)
    assert fresh_frame.paragraphs[1].text == "new run"
    assert fresh_frame.text == cell.text
    paragraph = cell.text_frame.paragraphs[0]
    held_run = paragraph.runs[0]
    paragraph.text = "replacement"
    assert paragraph.text == "replacement"
    _assert_stale_after_exactly_one_bump(__import__("rpptx"), lambda: held_run.text)
    with pytest.raises(StaleElementError, match=r"First invalidated by Paragraph.text at committed revision"):
        _ = held_run.text
    assert cell.text.startswith("replacement")
    assert prs.slides[0].shapes[0].table.cell(0, 0).text == cell.text


def test_cell_text_topology_failures_and_no_match_preserve_source_and_handles():
    from rpptx import ReplacementCountError
    prs, table = _f_x194_table()
    cell = table.cell(0, 0)
    cell.text = "needle"
    frame = cell.text_frame
    paragraph = frame.paragraphs[0]
    before = prs.to_bytes()
    assert frame.try_replace_text("absent", "other", expect=0) == 0
    with pytest.raises(ReplacementCountError):
        frame.try_replace_text("needle", "other", expect=2)
    with pytest.raises(Exception, match="empty"):
        frame.try_replace_text("", "other")
    with pytest.raises(Exception):
        cell.merge(cell)
    with pytest.raises(Exception):
        cell.split()
    assert prs.to_bytes() == before
    assert paragraph.text == "needle"
    assert frame.text == "needle"
    assert frame.try_replace_text("needle", "other", expect=1) == 1
    from rpptx import StaleElementError
    with pytest.raises(StaleElementError):
        _ = paragraph.text
    assert cell.text == "other"


def test_cell_text_frame_edits_preserve_merge_origin_and_continuation_ownership():
    from rpptx import StaleElementError
    prs, table = _f_x194_table()
    origin, continuation = table.cell(0, 0), table.cell(0, 1)
    origin.text = "origin"
    continuation.text = "continuation"
    old = continuation.text_frame
    origin.merge(continuation)
    assert old.text == continuation.text
    assert continuation.is_spanned
    continuation.text = "explicit continuation"
    assert continuation.text_frame.text == "explicit continuation"
    assert "explicit continuation" not in origin.text_frame.text
    old = origin.text_frame
    origin.split()
    assert old.text == origin.text
    assert not continuation.is_spanned
    assert table.cell(1, 1).text == ""
    assert len(prs.slides[0].shapes[0].table.rows) == 2


def test_cell_text_value_edits_and_unrelated_structure_retain_handles():
    from rpptx import StaleElementError, Pt
    from rpptx.enum.text import MSO_ANCHOR
    prs, table = _f_x194_table()
    cell = table.cell(0, 0)
    cell.text = "value"
    frame = cell.text_frame
    p = frame.paragraphs[0]
    run = p.runs[0]
    font = run.font
    run.text = "changed"
    font.bold = True
    font.size = Pt(18)
    p.space_after = Pt(4)
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.word_wrap = True
    assert frame.text == p.text == run.text == "changed"
    assert font.bold is True
    assert frame.vertical_anchor == MSO_ANCHOR.MIDDLE
    prs.slides.add_slide(prs.slide_layouts[6])
    assert cell.text == frame.text == p.text == run.text == "changed"
    assert font.bold is True
    prs.slides[0].shapes.remove(prs.slides[0].shapes[0])
    for operation in (lambda: cell.text, lambda: frame.text, lambda: p.text,
                      lambda: run.text, lambda: font.bold):
        with pytest.raises(StaleElementError, match=r"First invalidated by ShapeCollection.remove "):
            operation()


def test_cell_hyperlink_assignment_uses_slide_and_refuses_invalid_input_atomically():
    import rpptx

    prs, table = _f_x194_table()
    cell = table.cell(0, 0)
    cell.text = "link"
    hyperlink = cell.text_frame.paragraphs[0].runs[0].hyperlink
    assert hyperlink.address is None
    before = prs.to_bytes()
    for value in (None, ""):
        hyperlink.address = value
        assert prs.to_bytes() == before
        assert hyperlink.address is None
    hyperlink.address = "https://example.com"
    assert hyperlink.address == "https://example.com"
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="control characters"):
        hyperlink.address = "https://example.com/\ninvalid"
    assert prs.to_bytes() == before
    reopened = rpptx.Presentation.from_bytes(before)
    assert (
        reopened.slides[0].shapes[0].table.cell(0, 0)
        .text_frame.paragraphs[0].runs[0].hyperlink.address
        == "https://example.com"
    )


def test_cell_text_frame_getter_and_missing_body_failures_do_not_materialize_source():
    from rpptx import Presentation, ReplacementCountError
    prs, _ = _f_x194_table()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(prs.to_bytes())) as source, zipfile.ZipFile(out, "w") as target:
        for member in source.infolist():
            data = source.read(member.filename)
            if member.filename == "ppt/slides/slide1.xml":
                data, count = re.subn(rb"<a:txBody\b[^>]*>.*?</a:txBody>", b"", data, count=1, flags=re.S)
                assert count == 1
            target.writestr(member, data)
    prs = Presentation.from_bytes(out.getvalue())
    cell = prs.slides[0].shapes[0].table.cell(0, 0)
    before = prs.to_bytes()
    frame = cell.text_frame
    assert frame.autofit is None
    assert frame.try_replace_text("absent", "replacement", expect=0) == 0
    with pytest.raises(ReplacementCountError):
        frame.try_replace_text("absent", "replacement", expect=1)
    with pytest.raises(Exception, match="empty"):
        frame.try_replace_text("", "replacement")
    assert prs.to_bytes() == before
    cell.text = "materialized by edit"
    assert cell.text_frame.text == "materialized by edit"


def test_cell_text_frame_edits_reach_pdf_and_raster_output(tmp_path):
    import subprocess
    from PIL import Image
    from rpptx import Presentation, Pt
    from rpptx.enum.text import PP_ALIGN, MSO_ANCHOR
    prs, table = _f_x194_table()
    cell = table.cell(0, 0)
    cell.text = "F-X194 first"
    p = cell.text_frame.paragraphs[0]
    run = p.runs[0]
    run.font.name = "Liberation Sans"
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color = "CC0000"
    p.alignment = PP_ALIGN.CENTER
    p.line_spacing = 1.2
    p.space_before = Pt(2)
    p.space_after = Pt(3)
    frame = cell.text_frame
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.word_wrap = True
    p2 = frame.add_paragraph()
    r2 = p2.add_run("F-X194 second")
    r2.font.name = "Liberation Sans"
    r2.font.size = Pt(18)
    reopened = Presentation.from_bytes(prs.to_bytes())
    cell = reopened.slides[0].shapes[0].table.cell(0, 0)
    assert cell.text == "F-X194 first\nF-X194 second"
    p = cell.text_frame.paragraphs[0]
    assert p.runs[0].font.size == Pt(20)
    assert p.runs[0].font.bold is True
    assert str(p.runs[0].font.color.rgb) == "CC0000"
    assert p.alignment == PP_ALIGN.CENTER
    assert p.space_before == Pt(2) and p.space_after == Pt(3)
    assert cell.text_frame.word_wrap is True
    pdf = tmp_path / "cell.pdf"
    pdf.write_bytes(reopened.to_pdf())
    extracted = subprocess.check_output(["pdftotext", str(pdf), "-"], text=True)
    assert "F-X194 first" in extracted and "F-X194 second" in extracted
    png = reopened.render_slide_to_png(0, dpi=72)
    image = Image.open(io.BytesIO(png)).convert("RGB")
    assert any(red > green + 40 and red > blue + 40 for red, green, blue in image.getdata())


def test_grouped_table_cell_frame_resolves_and_imported_hyperlink_reads():
    import xml.etree.ElementTree as ET
    from rpptx import Presentation, Pt
    prs, table = _f_x194_table()
    table.cell(0, 0).text = "Grouped link"
    pns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    ans = "http://schemas.openxmlformats.org/drawingml/2006/main"
    rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    # Keep the fixture on the existing canonical-prefix parser contract.
    ET.register_namespace("p", pns)
    ET.register_namespace("a", ans)
    ET.register_namespace("r", rns)
    ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/relationships")
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(prs.to_bytes())) as source, zipfile.ZipFile(out, "w") as target:
        for member in source.infolist():
            data = source.read(member.filename)
            if member.filename == "ppt/slides/slide1.xml":
                root = ET.fromstring(data)
                tree = root.find(".//{" + pns + "}spTree")
                table_node = tree.find("{" + pns + "}graphicFrame")
                tree.remove(table_node)
                group = ET.SubElement(tree, "{" + pns + "}grpSp")
                nv = ET.SubElement(group, "{" + pns + "}nvGrpSpPr")
                ET.SubElement(nv, "{" + pns + "}cNvPr", id="10", name="Cell owner group")
                ET.SubElement(nv, "{" + pns + "}cNvGrpSpPr")
                ET.SubElement(nv, "{" + pns + "}nvPr")
                props = ET.SubElement(group, "{" + pns + "}grpSpPr")
                xfrm = ET.SubElement(props, "{" + ans + "}xfrm")
                for tag, attrs in (("off", {"x":"0", "y":"0"}), ("ext", {"cx":"3657600", "cy":"1828800"}),
                                   ("chOff", {"x":"0", "y":"0"}), ("chExt", {"cx":"3657600", "cy":"1828800"})):
                    ET.SubElement(xfrm, "{" + ans + "}" + tag, attrs)
                group.append(table_node)
                run = table_node.find(".//{" + ans + "}r")
                rp = run.find("{" + ans + "}rPr")
                if rp is None:
                    rp = ET.Element("{" + ans + "}rPr")
                    run.insert(0, rp)
                ET.SubElement(rp, "{" + ans + "}hlinkClick", {"{" + rns + "}id":"rIdCellLink"})
                data = ET.tostring(root)
            elif member.filename == "ppt/slides/_rels/slide1.xml.rels":
                root = ET.fromstring(data)
                ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                              Id="rIdCellLink", Type=rns + "/hyperlink", Target="https://example.com/grouped", TargetMode="External")
                data = ET.tostring(root)
            target.writestr(member, data)
    prs = Presentation.from_bytes(out.getvalue())
    table = prs.slides[0].shapes[0].shapes[0].table
    cell = table.cell(0, 0)
    frame = cell.text_frame
    run = frame.paragraphs[0].runs[0]
    assert run.hyperlink.address == "https://example.com/grouped"
    run.font.size = Pt(17)
    run.text = "Grouped edited"
    assert frame.text == "Grouped edited"
    reopened = Presentation.from_bytes(prs.to_bytes())
    cell = reopened.slides[0].shapes[0].shapes[0].table.cell(0, 0)
    assert cell.text == "Grouped edited"
    assert cell.text_frame.paragraphs[0].runs[0].font.size == Pt(17)


def _f_x197_ppt_colour_surface(prs, surface):
    from rpptx.enum.shapes import MSO_CONNECTOR
    slide = prs.slides[0]
    if surface.startswith("font"):
        if "cell" in surface:
            cell = slide.shapes.add_table(1, 1, 0, 0, 100, 100).table.cell(0, 0)
            frame = cell.text_frame
            frame.text = "colour"
            frame = cell.text_frame
        else:
            frame = slide.shapes[0].text_frame
        paragraph = frame.paragraphs[0]
        font = paragraph.font if "default" in surface else paragraph.runs[0].font
        colour = font.color
        if surface.endswith("direct"):
            return lambda value: setattr(font, "color", value), colour
        return lambda value: setattr(colour, "rgb", value), colour
    if surface == "background-fill":
        fill = slide.background.fill
        fill.solid()
        colour = fill.fore_color
    elif surface.startswith("cell"):
        cell = slide.shapes.add_table(1, 1, 0, 0, 100, 100).table.cell(0, 0)
        if surface == "cell-fill":
            cell.fill.solid()
            colour = cell.fill.fore_color
        else:
            colour = getattr(cell, "border_" + surface.removeprefix("cell-")).color
    else:
        kind, property_name = surface.split("-")
        if kind == "picture":
            shape = slide.shapes.add_picture(io.BytesIO(_tiny_png()), 0, 0)
        elif kind == "connector":
            shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 100, 100)
        elif kind == "group":
            shape = slide.shapes.add_group_shape()
        else:
            shape = slide.shapes[0]
        if property_name == "fill":
            shape.fill.solid()
            colour = shape.fill.fore_color
        elif property_name == "line":
            colour = shape.line.color
        else:
            colour = shape.shadow.color
    return lambda value: setattr(colour, "rgb", value), colour


_F_X197_PPT_COLOUR_SURFACES = [
    "font-run-direct", "font-run-rgb", "font-default-direct", "font-default-rgb",
    "font-cell-direct", "font-cell-rgb", "font-cell-default-direct", "font-cell-default-rgb",
    "shape-fill", "picture-fill", "connector-fill", "cell-fill", "background-fill",
    "shape-line", "picture-line", "connector-line", "cell-top", "cell-bottom", "cell-left", "cell-right",
    "shape-shadow", "picture-shadow", "connector-shadow", "group-shadow",
]


@pytest.mark.parametrize("surface", _F_X197_PPT_COLOUR_SURFACES)
@pytest.mark.parametrize("form", ["rgb", "triple", "hex", "hash"])
def test_every_powerpoint_colour_setter_accepts_rgb_and_optional_hash_hex(surface, form):
    import rpptx
    from rpptx.dml.color import RGBColor
    from xml.etree import ElementTree as ET
    prs = _textbox_presentation(rpptx)
    expected = RGBColor(0xA1, 0xB2, 0xC3)
    assert RGBColor.from_string("#a1b2c3") == expected
    assert RGBColor.from_string("a1b2c3") == expected
    value = {"rgb": expected, "triple": tuple(expected), "hex": "a1b2c3", "hash": "#a1b2c3"}[form]
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    setter(value)
    assert colour.rgb == expected and isinstance(colour.rgb, RGBColor)
    saved = prs.to_bytes()
    setter(value)
    assert prs.to_bytes() == saved
    reopened = rpptx.Presentation.from_bytes(saved)
    with zipfile.ZipFile(io.BytesIO(reopened.to_bytes())) as package:
        root = ET.fromstring(package.read("ppt/slides/slide1.xml"))
        assert any(node.get("val") == "A1B2C3" for node in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr"))


@pytest.mark.parametrize("surface", _F_X197_PPT_COLOUR_SURFACES)
@pytest.mark.parametrize("value,error_type", [(object(), TypeError), (7, TypeError), ((1, 2), TypeError), ((1.5, 2, 3), TypeError), ((10**100, 2, 3), ValueError), ((-1, 2, 3), ValueError), ("##123456", ValueError), ("１２３４５６", ValueError), ("12345\ud800", ValueError), (" 123456", ValueError)])
def test_powerpoint_colour_errors_are_atomic_on_every_existing_owner(surface, value, error_type):
    import rpptx
    prs = _textbox_presentation(rpptx)
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    before = prs.to_bytes()
    prior = colour.rgb
    with pytest.raises(error_type) as caught:
        setter(value)
    parameter = "Font.color" + ("" if surface.endswith("direct") else ".rgb") if surface.startswith("font") else "ShadowFormat.color.rgb" if surface.endswith("shadow") else "LineFormat.color.rgb" if surface.endswith("line") or surface in ("cell-top", "cell-bottom", "cell-left", "cell-right") else "FillFormat.fore_color.rgb"
    assert parameter in str(caught.value)
    if error_type is TypeError:
        assert type(value).__name__ in str(caught.value)
    assert "RGBColor" in str(caught.value) or "hex" in str(caught.value) or "channels" in str(caught.value)
    assert prs.to_bytes() == before
    assert colour.rgb == prior


def test_powerpoint_font_colour_format_preserves_theme_and_transform_sources():
    import rpptx
    prs = _textbox_presentation(rpptx)
    colour = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color
    assert colour.rgb is None
    colour.theme_color = "accent1"
    colour.brightness = 0.25
    assert colour.type is MSO_COLOR_TYPE.SCHEME
    assert colour.theme_color is MSO_THEME_COLOR.ACCENT_1
    assert colour.rgb is None
    reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
    actual = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color
    assert actual.theme_color is MSO_THEME_COLOR.ACCENT_1
    assert actual.brightness == 0.25


@pytest.mark.parametrize("surface", ["font-run-rgb", "font-default-rgb", "shape-fill", "shape-line"])
@pytest.mark.parametrize("value,modulation,offset", [(-1.0, 0, None), (-0.4, 60000, None), (0.0, None, None), (0.4, 60000, 40000), (1.0, 0, 100000), (0.123456, 87654, 12346)])
def test_powerpoint_brightness_matches_pinned_finite_oracle_and_preserves_non_luminance(surface, value, modulation, offset):
    import rpptx
    import re
    from xml.etree import ElementTree as ET
    prs = _textbox_presentation(rpptx)
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    colour.theme_color = "accent1"
    parts = _package_parts(prs.to_bytes())
    xml = parts["ppt/slides/slide1.xml"]
    # The explicit accent1 below belongs to the selected colour owner.
    xml, count = re.subn(rb'<a:schemeClr val="accent1"\s*/>', b'<a:schemeClr val="accent1"><a:alpha val="43000"/><a:tint val="12000"/><a:lumMod val="70000"/><a:lumOff val="30000"/></a:schemeClr>', xml)
    assert count == 1
    parts["ppt/slides/slide1.xml"] = xml
    prs = rpptx.Presentation.from_bytes(_package_bytes(parts))
    if surface.startswith("font"):
        paragraph = prs.slides[0].shapes[0].text_frame.paragraphs[0]
        colour = (paragraph.font if "default" in surface else paragraph.runs[0].font).color
    else:
        colour = prs.slides[0].shapes[0].fill.fore_color if surface == "shape-fill" else prs.slides[0].shapes[0].line.color
    colour.brightness = value
    expected = offset / 100000 if offset is not None else modulation / 100000 - 1 if modulation is not None else 0.0
    assert colour.brightness == pytest.approx(expected)
    reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
    root = ET.fromstring(_package_parts(reopened.to_bytes())["ppt/slides/slide1.xml"])
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    node = next(node for node in root.iter(ns + "schemeClr") if node.get("val") == "accent1")
    actual = {child.tag.removeprefix(ns): int(child.get("val")) for child in node}
    wanted = {"alpha": 43000, "tint": 12000}
    if modulation is not None:
        wanted["lumMod"] = modulation
    if offset is not None:
        wanted["lumOff"] = offset
    assert actual == wanted


@pytest.mark.parametrize("surface", ["font-run-rgb", "font-default-rgb", "shape-fill", "shape-line", "shape-shadow"])
@pytest.mark.parametrize("value,error_type", [(True, TypeError), (None, TypeError), ("0.4", TypeError), (float("nan"), ValueError), (float("inf"), ValueError), (-1.1, ValueError), (1.1, ValueError)])
def test_powerpoint_brightness_strict_finite_policy_errors_are_atomic(surface, value, error_type):
    import rpptx
    prs = _textbox_presentation(rpptx)
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    colour.theme_color = "accent1"
    before = prs.to_bytes()
    with pytest.raises(error_type, match="brightness"):
        colour.brightness = value
    assert prs.to_bytes() == before
    assert colour.theme_color is MSO_THEME_COLOR.ACCENT_1


@pytest.mark.parametrize("surface", ["font-run-rgb", "font-default-rgb", "font-cell-rgb", "cell-fill", "cell-top", "shape-shadow"])
def test_powerpoint_colour_views_reject_stale_owners_without_mutating_source(surface):
    import rpptx
    prs = _textbox_presentation(rpptx)
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    setter("123456")
    prs.slides[0].shapes.remove(prs.slides[0].shapes[0])
    if surface in {"font-cell-rgb", "cell-fill", "cell-top"}:
        # The first deletion removes the unrelated textbox, not the table owner.
        assert str(colour.rgb) == "123456"
        colour.rgb = "ABCDEF"
        assert str(colour.rgb) == "ABCDEF"
        prs.slides[0].shapes.remove(prs.slides[0].shapes[0])
    before = prs.to_bytes()
    with pytest.raises(rpptx.StaleElementError):
        _ = colour.rgb
    with pytest.raises(rpptx.StaleElementError):
        colour.rgb = "ABCDEF"
    assert prs.to_bytes() == before


_F_X197_RAW_COLOUR_CHILDREN = (
    b'<a:futureTransform val="12345"/>',
    b'<fx:colourMetadata xmlns:fx="urn:f-x197:opaque" marker="  retain &amp; exact  "><fx:child/>opaque<!--keep--></fx:colourMetadata>',
)


@pytest.mark.parametrize("surface", ["font-run-rgb", "font-default-rgb", "shape-fill", "shape-line", "shape-shadow"])
@pytest.mark.parametrize("source_kind,replacement", [("rgb", "rgb"), ("theme", "theme"), ("theme", "rgb"), ("rgb", "theme")])
def test_powerpoint_colour_replacement_preserves_family_specific_transforms_and_opaque_children(surface, source_kind, replacement):
    import rpptx
    import re
    from xml.etree import ElementTree as ET
    prs = _textbox_presentation(rpptx)
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    if source_kind == "rgb":
        setter("123456")
        tag, original = "srgbClr", "123456"
    else:
        colour.theme_color = "accent1"
        tag, original = "schemeClr", "accent1"
    parts = _package_parts(prs.to_bytes())
    xml = parts["ppt/slides/slide1.xml"]
    # Replace exactly the chosen owner's explicit colour, preserving all siblings.
    pattern = rb'<a:' + tag.encode() + rb' val="' + original.encode() + rb'"(?:\s*/>|>.*?</a:' + tag.encode() + rb'>)'
    fixture = '<a:' + tag + ' val="' + original + '"><a:alpha val="43000"/><a:tint val="12000"/><a:gamma/>' + b''.join(_F_X197_RAW_COLOUR_CHILDREN).decode() + '</a:' + tag + '>'
    xml, count = re.subn(pattern, fixture.encode(), xml, flags=re.DOTALL)
    assert count == 1
    parts["ppt/slides/slide1.xml"] = xml
    prs = rpptx.Presentation.from_bytes(_package_bytes(parts))
    shape = prs.slides[0].shapes[0]
    if surface.startswith("font"):
        paragraph = shape.text_frame.paragraphs[0]
        colour = (paragraph.font if "default" in surface else paragraph.runs[0].font).color
    elif surface == "shape-fill":
        colour = shape.fill.fore_color
    elif surface == "shape-line":
        colour = shape.line.color
    else:
        colour = shape.shadow.color
    before = prs.to_bytes()
    assert colour.type is (MSO_COLOR_TYPE.RGB if source_kind == "rgb" else MSO_COLOR_TYPE.SCHEME)
    _ = colour.rgb, colour.theme_color, colour.brightness
    assert prs.to_bytes() == before
    for raw in _F_X197_RAW_COLOUR_CHILDREN:
        assert _package_parts(before)["ppt/slides/slide1.xml"].count(raw) == 1
    colour.brightness = 0.25
    brightness_saved = prs.to_bytes()
    brightness_reopened = rpptx.Presentation.from_bytes(brightness_saved)
    brightness_xml = _package_parts(brightness_reopened.to_bytes())["ppt/slides/slide1.xml"]
    for raw in _F_X197_RAW_COLOUR_CHILDREN:
        assert brightness_xml.count(raw) == 1
    assert colour.brightness == 0.25
    if replacement == "rgb":
        colour.rgb = "ABCDEF"
        target_tag, target_value = "srgbClr", "ABCDEF"
    else:
        colour.theme_color = "accent2"
        target_tag, target_value = "schemeClr", "accent2"
    reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
    actual_parts = _package_parts(reopened.to_bytes())
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    root = ET.fromstring(actual_parts["ppt/slides/slide1.xml"])
    node = next(node for node in root.iter(ns + target_tag) if node.get("val") == target_value)
    actual = {child.tag.removeprefix(ns): child.get("val") for child in node if child.tag in {ns + name for name in ("alpha", "tint", "gamma", "lumMod", "lumOff")}}
    if source_kind == replacement:
        assert actual == {"alpha": "43000", "tint": "12000", "gamma": None, "lumMod": "75000", "lumOff": "25000"}
        for raw in _F_X197_RAW_COLOUR_CHILDREN:
            assert actual_parts["ppt/slides/slide1.xml"].count(raw) == 1
        assert b"<a:gamma/>" in actual_parts["ppt/slides/slide1.xml"]
    elif surface == "shape-shadow":
        assert actual == {"alpha": "43000"}
    else:
        assert actual == {}
    if source_kind != replacement:
        for raw in _F_X197_RAW_COLOUR_CHILDREN:
            assert raw not in actual_parts["ppt/slides/slide1.xml"]
    assert actual_parts["ppt/slides/_rels/slide1.xml.rels"] == parts["ppt/slides/_rels/slide1.xml.rels"]


@pytest.mark.parametrize("surface", ["font-run-rgb", "font-default-rgb", "font-cell-rgb", "font-cell-default-rgb"])
def test_powerpoint_unset_font_colour_reads_and_clears_do_not_create_properties(surface):
    import rpptx
    prs = _textbox_presentation(rpptx)
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    before = prs.to_bytes()
    assert (colour.rgb, colour.type, colour.brightness) == (None, None, None)
    with pytest.raises(AttributeError, match="no direct color"):
        _ = colour.theme_color
    assert prs.to_bytes() == before
    colour.rgb = None
    assert prs.to_bytes() == before
    colour.theme_color = None
    assert prs.to_bytes() == before
    with pytest.raises(ValueError, match="explicit colour"):
        colour.brightness = 0.4
    assert prs.to_bytes() == before


@pytest.mark.parametrize("surface", ["shape-fill", "shape-line", "shape-shadow", "cell-fill", "cell-top"])
def test_powerpoint_non_font_colours_retain_none_refusal_and_source(surface):
    import rpptx
    prs = _textbox_presentation(rpptx)
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    before = prs.to_bytes()
    with pytest.raises(TypeError):
        colour.rgb = None
    with pytest.raises(TypeError):
        colour.theme_color = None
    assert prs.to_bytes() == before


def test_powerpoint_nested_font_colour_reaches_saved_rendered_pixels():
    import rpptx
    from PIL import Image
    def render(value):
        prs = _textbox_presentation(rpptx)
        font = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
        font.name = "Liberation Sans"
        font.size = rpptx.Pt(28)
        font.color.rgb = value
        return rpptx.Presentation.from_bytes(prs.to_bytes()).render_slide_to_png(0, dpi=72)
    red = render("#FF0000")
    assert red == render("ff0000")
    blue = render("0000FF")
    assert red != blue
    assert any(r > g + 40 and r > b + 40 for r, g, b in Image.open(io.BytesIO(red)).convert("RGB").getdata())
    assert any(b > g + 40 and b > r + 40 for r, g, b in Image.open(io.BytesIO(blue)).convert("RGB").getdata())


def test_powerpoint_theme_and_brightness_reach_deterministic_pdf_and_raster_pixels(tmp_path):
    import shutil
    import subprocess
    from xml.etree import ElementTree as ET
    from PIL import Image
    import rpptx
    executable = shutil.which("pdftoppm")
    assert executable is not None
    version = subprocess.run([executable, "-v"], capture_output=True, text=True, timeout=60)
    assert version.returncode == 0 and "26.01.0" in version.stdout + version.stderr
    namespace = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    def render(name, rgb=None, theme=None, brightness=0.0):
        prs = _textbox_presentation(rpptx)
        font = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
        font.name = "Liberation Sans"
        font.size = rpptx.Pt(28)
        if theme is None:
            font.color.rgb = rgb
        else:
            font.color.theme_color = theme
            font.color.brightness = brightness
        parts = _package_parts(prs.to_bytes())
        tree = ET.fromstring(parts["ppt/theme/theme1.xml"])
        for token, value in (("accent1", "FF0000"), ("accent2", "0000FF")):
            colour = tree.find(".//" + namespace + token + "/" + namespace + "srgbClr")
            assert colour is not None
            colour.set("val", value)
        parts["ppt/theme/theme1.xml"] = ET.tostring(tree)
        reopened = rpptx.Presentation.from_bytes(_package_bytes(parts))
        if theme is not None:
            colour = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color
            assert colour.rgb is None and colour.theme_color is {"accent1": MSO_THEME_COLOR.ACCENT_1, "accent2": MSO_THEME_COLOR.ACCENT_2}[theme]
            assert colour.brightness == pytest.approx(brightness)
        png = Image.open(io.BytesIO(reopened.render_slide_to_png(0, dpi=72))).convert("RGB")
        pdf = reopened.to_pdf()
        assert pdf.startswith(b"%PDF")
        source = tmp_path / (name + ".pdf")
        source.write_bytes(pdf)
        prefix = tmp_path / name
        result = subprocess.run([executable, "-f", "1", "-singlefile", "-r", "72", "-png", str(source), str(prefix)], capture_output=True, text=True, timeout=60)
        assert result.returncode == 0, result.stdout + result.stderr
        raster = Image.open(prefix.with_suffix(".png")).convert("RGB")
        return png, raster
    red = render("literal-red", rgb="FF0000")
    blue = render("literal-blue", rgb="0000FF")
    theme_red = render("theme-red", theme="accent1")
    theme_blue = render("theme-blue", theme="accent2")
    lighter = render("theme-lighter", theme="accent1", brightness=0.4)
    darker = render("theme-darker", theme="accent1", brightness=-0.4)
    for index in (0, 1):
        assert theme_red[index].size == red[index].size == lighter[index].size == darker[index].size
        assert theme_red[index].tobytes() == red[index].tobytes()
        assert theme_blue[index].tobytes() == blue[index].tobytes()
        baseline = list(red[index].getdata())
        solid = [position for position, pixel in enumerate(baseline) if pixel == (255, 0, 0)]
        assert len(solid) > 5
        light_pixels = list(lighter[index].getdata())
        dark_pixels = list(darker[index].getdata())
        for position in solid:
            r, g, b = light_pixels[position]
            assert r > 220 and g > 40 and b > 40
            r, g, b = dark_pixels[position]
            assert 0 < r < 220 and g < 10 and b < 10
        assert any(b > r + 40 and b > g + 40 for r, g, b in blue[index].getdata())


@pytest.mark.parametrize("code,token", list(enumerate(["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink", "tx1", "bg1", "tx2", "bg2"], start=1)))
def test_powerpoint_supported_theme_enum_values_round_trip_without_invented_rgb(code, token):
    import rpptx
    from rpptx.enum.dml import MSO_THEME_COLOR
    prs = _textbox_presentation(rpptx)
    colour = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color
    colour.theme_color = MSO_THEME_COLOR(code)
    assert colour.rgb is None and colour.type is MSO_COLOR_TYPE.SCHEME
    assert colour.theme_color is MSO_THEME_COLOR(code)
    assert ('val="' + token + '"').encode() in _package_parts(prs.to_bytes())["ppt/slides/slide1.xml"]
    reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
    assert reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color.theme_color is MSO_THEME_COLOR(code)


def test_powerpoint_fill_foreground_prerequisite_and_shadow_default_opacity_remain_intact():
    import rpptx
    prs = _textbox_presentation(rpptx)
    shape = prs.slides[0].shapes[0]
    before = prs.to_bytes()
    with pytest.raises(TypeError, match="foreground"):
        _ = shape.fill.fore_color
    assert prs.to_bytes() == before
    shadow = shape.shadow
    shadow.color.rgb = "123456"
    assert shadow.alpha == 0.4
    shadow.alpha = 0.25
    shadow.color.theme_color = "accent1"
    assert shadow.alpha == 0.25
    shadow.color.rgb = "654321"
    assert shadow.alpha == 0.25


@pytest.mark.parametrize("surface", ["font-run-rgb", "font-default-rgb", "shape-fill", "shape-line", "shape-shadow", "cell-fill", "cell-top"])
@pytest.mark.parametrize("value,error_type", [("unknown", ValueError), (0, ValueError), (-2, ValueError), (17, ValueError), (True, TypeError), (object(), TypeError)])
def test_powerpoint_invalid_theme_selection_is_atomic_on_all_colour_owner_kinds(surface, value, error_type):
    import rpptx
    prs = _textbox_presentation(rpptx)
    setter, colour = _f_x197_ppt_colour_surface(prs, surface)
    setter("123456")
    before = prs.to_bytes()
    with pytest.raises(error_type, match="theme_color"):
        colour.theme_color = value
    assert prs.to_bytes() == before
    assert str(colour.rgb) == "123456"


@pytest.mark.parametrize("kind,source,type_token", [("srgbClr", '<a:srgbClr val="123456"/>', "rgb"), ("schemeClr", '<a:schemeClr val="accent1"/>', "theme"), ("prstClr", '<a:prstClr val="red"/>', "preset"), ("sysClr", '<a:sysClr val="windowText" lastClr="000000"/>', "system")])
def test_powerpoint_imported_colour_kind_reads_do_not_invent_rgb_or_materialize_source(kind, source, type_token):
    import rpptx
    import re
    prs = _textbox_presentation(rpptx)
    font = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
    font.color = "123456"
    parts = _package_parts(prs.to_bytes())
    parts["ppt/slides/slide1.xml"], count = re.subn(rb'<a:srgbClr val="123456"\s*/>', source.encode(), parts["ppt/slides/slide1.xml"])
    assert count == 1
    prs = rpptx.Presentation.from_bytes(_package_bytes(parts))
    colour = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color
    before = prs.to_bytes()
    assert colour.type is {"rgb": MSO_COLOR_TYPE.RGB, "theme": MSO_COLOR_TYPE.SCHEME, "preset": MSO_COLOR_TYPE.PRESET, "system": MSO_COLOR_TYPE.SYSTEM}[type_token]
    assert colour.theme_color is (MSO_THEME_COLOR.ACCENT_1 if kind == "schemeClr" else MSO_THEME_COLOR.NOT_THEME_COLOR)
    if kind == "srgbClr":
        assert str(colour.rgb) == "123456"
    else:
        assert colour.rgb is None
    assert prs.to_bytes() == before


def test_powerpoint_cell_colour_view_obeys_the_existing_cell_text_generation():
    import rpptx
    prs, table = _f_x194_table()
    cell = table.cell(0, 0)
    cell.text = "first"
    colour = cell.text_frame.paragraphs[0].runs[0].font.color
    colour.rgb = "123456"
    cell.text = "second"
    before = prs.to_bytes()
    with pytest.raises(rpptx.StaleElementError, match=r"First invalidated by Cell.text "):
        _ = colour.rgb
    with pytest.raises(rpptx.StaleElementError, match=r"First invalidated by Cell.text "):
        colour.theme_color = "accent1"
    assert prs.to_bytes() == before
    fresh = cell.text_frame.paragraphs[0].runs[0]
    assert fresh.text == "second"
    assert str(fresh.font.color.rgb) == "123456"
    assert prs.to_bytes() == before
    reopened = rpptx.Presentation.from_bytes(before)
    actual = reopened.slides[0].shapes[0].table.cell(0, 0).text_frame.paragraphs[0].runs[0]
    assert actual.text == "second"
    assert str(actual.font.color.rgb) == "123456"


def test_fx206_ancestors_and_unrelated_cell_text_survive_local_topology():
    from rpptx import Presentation, StaleElementError

    presentation = Presentation()
    layout = presentation.slide_layouts[6]
    slide = presentation.slides.add_slide(layout)
    table_shape = slide.shapes.add_table(2, 2, 0, 0, 1000, 1000)
    table = table_shape.table
    cell = table.cell(0, 0)
    other = table.cell(1, 1)
    cell.text = "old"
    other.text = "other"
    frame = cell.text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    font = run.font
    other_frame = other.text_frame
    cell.text = "new"
    assert cell.text == frame.text == "new"
    assert other_frame.text == "other"
    assert len(table.rows) == 2
    assert table_shape.has_table
    assert len(slide.shapes) == 1
    assert layout.name is not None
    for get_value in (lambda: paragraph.text, lambda: run.text, lambda: font.bold):
        with pytest.raises(StaleElementError, match=r"First invalidated by Cell.text "):
            get_value()


def test_fx206_shape_and_slide_moves_preserve_current_owner_lookup():
    from rpptx import Presentation

    presentation = Presentation()
    first = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = first.shapes.add_textbox(0, 0, 100, 100)
    shape.text = "held"
    other = first.shapes.add_textbox(0, 0, 100, 100)
    other.text = "other"
    first.shapes.move(0, 1)
    assert shape.text == "held"
    assert other.text == "other"
    second = presentation.slides.add_slide(presentation.slide_layouts[6])
    second.shapes.add_textbox(0, 0, 100, 100).text = "second"
    presentation.slides.move(0, 1)
    assert shape.text == "held"
    shape.text = "after move"
    assert presentation.slides[1].shapes[1].text == "after move"
    assert second.shapes[0].text == "second"


def test_fx206_removed_reused_shape_id_never_resurrects_a_handle():
    from rpptx import Presentation, StaleElementError

    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    removed = slide.shapes.add_textbox(0, 0, 100, 100)
    old_id = removed.shape_id
    slide.shapes.remove(removed)
    replacement = slide.shapes.add_textbox(0, 0, 100, 100)
    assert replacement.shape_id == old_id
    replacement.text = "replacement"
    with pytest.raises(StaleElementError) as raised:
        _ = removed.text
    assert "First invalidated by ShapeCollection.remove " in str(raised.value)
    assert "Shape.text" not in str(raised.value)
    assert replacement.text == "replacement"


def test_fx206_grid_change_retires_cells_but_preserves_table_and_other_shapes():
    from rpptx import Presentation, StaleElementError

    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    table_shape = slide.shapes.add_table(2, 2, 0, 0, 1000, 1000)
    table = table_shape.table
    other = slide.shapes.add_textbox(0, 0, 100, 100)
    other.text = "outside"
    held = table.cell(0, 0)
    held_rows = table.rows
    added = held_rows.add_row()
    assert len(table.rows) == 3
    assert len(table.columns) == 2
    assert [table.cell(2, index).text for index in range(2)] == ["", ""]
    assert added.height == table.rows[2].height
    assert other.text == "outside"
    assert table_shape.has_table
    with pytest.raises(StaleElementError, match=r"First invalidated by RowCollection.add_row "):
        _ = held.text
    assert len(held_rows) == 3


def test_fx206_failed_edits_and_noops_preserve_owner_and_descendant_states():
    from rpptx import Presentation, ReplacementCountError

    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 100, 100)
    shape.text = "needle"
    frame = shape.text_frame
    paragraph = frame.paragraphs[0]
    before = presentation.to_bytes()
    with pytest.raises(ReplacementCountError):
        frame.try_replace_text("needle", "changed", expect=2)
    assert presentation.to_bytes() == before
    assert paragraph.text == "needle"
    assert frame.try_replace_text("missing", "changed", expect=0) == 0
    slide.shapes.move(0, 0)
    presentation.slides.move(0, 0)
    assert presentation.to_bytes() == before
    assert paragraph.text == "needle"


    # The same displayed string is not a no-op when a match spans formatted runs.
    from rpptx import StaleElementError
    import xml.etree.ElementTree as ET

    for route, cause in (("slide", "Slide.try_replace_text"),
                         ("presentation", "Presentation.try_replace_text")):
        deck = Presentation()
        held_slide = deck.slides.add_slide(deck.slide_layouts[6])
        held_shape = held_slide.shapes.add_textbox(0, 0, 100, 100)
        held_frame = held_shape.text_frame
        first = held_frame.paragraphs[0].add_run()
        first.text = "A"
        first.font.bold = True
        second = held_frame.paragraphs[0].add_run()
        second.text = "B"
        second.font.italic = True
        # Setup additions retire descendants, so capture only final owner facts.
        held_paragraph = held_frame.paragraphs[0]
        old_first, old_second = held_paragraph.runs
        old_font = old_second.font
        unrelated_slide = deck.slides.add_slide(deck.slide_layouts[6])
        unrelated_shape = unrelated_slide.shapes.add_textbox(0, 0, 100, 100)
        unrelated_shape.text = "outside"
        unrelated_run = unrelated_shape.text_frame.paragraphs[0].runs[0]
        replace = held_slide.try_replace_text if route == "slide" else deck.try_replace_text
        original = deck.to_bytes()
        with pytest.raises(ReplacementCountError):
            replace("AB", "AB", expect=2)
        assert deck.to_bytes() == original
        assert [old_first.text, old_second.text] == ["A", "B"]
        assert replace("missing", "AB", expect=0) == 0
        assert deck.to_bytes() == original
        assert old_font.italic is True
        assert replace("AB", "AB", expect=1) == 1
        assert held_slide.shapes[0].shape_id == held_shape.shape_id
        assert held_frame.text == held_shape.text == "AB"
        assert unrelated_run.text == "outside"
        for value in (lambda: held_paragraph.text, lambda: old_first.text,
                      lambda: old_second.text, lambda: old_font.italic):
            with pytest.raises(StaleElementError, match="First invalidated by " + cause + " "):
                value()
        fresh = held_frame.paragraphs[0].runs
        assert [run.text for run in fresh] == ["AB", ""]
        assert fresh[0].font.bold is True and fresh[1].font.italic is True
        saved = deck.to_bytes()
        assert saved != original
        parts = _package_parts(saved)
        root = ET.fromstring(parts["ppt/slides/slide1.xml"])
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        runs = root.findall(".//a:p/a:r", ns)
        assert [run.find("a:t", ns).text or "" for run in runs] == ["AB", ""]
        assert runs[0].find("a:rPr", ns).get("b") == "1"
        assert runs[1].find("a:rPr", ns).get("i") == "1"
        reopened = Presentation.from_bytes(saved)
        reopened_runs = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs
        assert [run.text for run in reopened_runs] == ["AB", ""]
        assert reopened_runs[0].font.bold is True and reopened_runs[1].font.italic is True
        assert reopened.slides[1].shapes[0].text == "outside"


def test_fx206_duplicate_imported_ids_remain_readable_and_structural_refusal_is_atomic(tmp_path):
    from rpptx import Presentation

    original = Presentation()
    slide = original.slides.add_slide(original.slide_layouts[6])
    first = slide.shapes.add_textbox(0, 0, 100, 100)
    first.text = "first"
    second = slide.shapes.add_textbox(0, 0, 100, 100)
    second.text = "second"
    source = tmp_path / "unique.pptx"
    imported = tmp_path / "duplicate.pptx"
    original.save(source)
    _replace_in_slide(source, imported, f'id="{second.shape_id}"', f'id="{first.shape_id}"')
    presentation = Presentation(imported)
    held_slide = presentation.slides[0]
    held_first, held_second = held_slide.shapes[0], held_slide.shapes[1]
    assert held_first.shape_id == held_second.shape_id
    assert (held_first.text, held_second.text) == ("first", "second")
    before = presentation.to_bytes()
    with pytest.raises(ValueError, match="ambiguous shape id"):
        held_slide.shapes.move(0, 1)
    assert presentation.to_bytes() == before
    assert (held_first.text, held_second.text) == ("first", "second")
    with pytest.raises(ValueError, match="ambiguous shape id"):
        held_slide.shapes.add_textbox(0, 0, 100, 100)
    assert presentation.to_bytes() == before
    assert (held_first.text, held_second.text) == ("first", "second")


@pytest.mark.parametrize("axis", ["rows", "columns"])
@pytest.mark.parametrize("reorder", ["shape", "slide"])
def test_fx206_held_table_collection_uses_current_parent_identity(axis, reorder, tmp_path):
    from rpptx import Presentation, StaleElementError

    presentation = Presentation()
    first = presentation.slides.add_slide(presentation.slide_layouts[6])
    a = first.shapes.add_table(2, 2, 0, 0, 1000, 1000).table
    second = first if reorder == "shape" else presentation.slides.add_slide(presentation.slide_layouts[6])
    b = second.shapes.add_table(2, 2, 0, 0, 1000, 1000).table
    a.cell(0, 0).text = "A"
    b.cell(0, 0).text = "B"
    held_collection = getattr(a, axis)
    held_cell = a.cell(0, 0)
    if reorder == "shape":
        first.shapes.move(0, 1)
        fresh_a = presentation.slides[0].shapes[1].table
        fresh_b = presentation.slides[0].shapes[0].table
    else:
        presentation.slides.move(0, 1)
        fresh_a = presentation.slides[1].shapes[0].table
        fresh_b = presentation.slides[0].shapes[0].table
    before = presentation.to_bytes()
    with pytest.raises(ValueError, match="not in this collection"):
        held_collection.remove(getattr(fresh_b, axis)[1])
    assert presentation.to_bytes() == before
    assert held_cell.text == "A"
    assert len(getattr(a, axis)) == len(getattr(b, axis)) == 2
    held_collection.remove(getattr(fresh_a, axis)[1])
    assert len(getattr(a, axis)) == 1
    assert len(getattr(b, axis)) == 2
    assert a.cell(0, 0).text == "A"
    assert b.cell(0, 0).text == "B"
    with pytest.raises(StaleElementError, match="First invalidated by .*Collection.remove "):
        _ = held_cell.text
    target = tmp_path / f"parent-{axis}-{reorder}.pptx"
    presentation.save(target)
    reopened = Presentation(target)
    actual_a = reopened.slides[0 if reorder == "shape" else 1].shapes[1 if reorder == "shape" else 0].table
    actual_b = reopened.slides[0].shapes[0].table
    assert len(getattr(actual_a, axis)) == 1
    assert len(getattr(actual_b, axis)) == 2
    assert (actual_a.cell(0, 0).text, actual_b.cell(0, 0).text) == ("A", "B")


def test_fx206_held_hyperlink_resolves_current_slide_and_prunes_only_its_relationships(tmp_path):
    import xml.etree.ElementTree as ET
    from rpptx import Presentation, RpptxError

    presentation = Presentation()
    slides = [presentation.slides.add_slide(presentation.slide_layouts[6]) for _ in range(2)]
    shapes = [slide.shapes.add_textbox(0, 0, 1000, 1000) for slide in slides]
    for shape, marker in zip(shapes, ("A", "B")):
        shape.text = marker
    links = [shape.text_frame.paragraphs[0].runs[0].hyperlink for shape in shapes]
    links[0].address = "https://example.test/a"
    links[1].address = "https://example.test/b"
    assert shapes[0].shape_id == shapes[1].shape_id
    parts = _package_parts(presentation.to_bytes())
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    rel_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    ids = [ET.fromstring(parts[f"ppt/slides/slide{i}.xml"]).find(".//a:hlinkClick", ns).attrib[rel_attr] for i in (1, 2)]
    assert ids[0] == ids[1]
    presentation.slides.move(0, 1)
    assert (links[0].address, links[1].address) == ("https://example.test/a", "https://example.test/b")
    links[0].address = "https://example.test/a2"
    assert links[1].address == "https://example.test/b"
    parts = _package_parts(presentation.to_bytes())
    assert b"https://example.test/a2" in parts["ppt/slides/_rels/slide1.xml.rels"]
    assert b'Target="https://example.test/a"' not in parts["ppt/slides/_rels/slide1.xml.rels"]
    assert b"https://example.test/b" in parts["ppt/slides/_rels/slide2.xml.rels"]
    inserted = presentation.slides.import_slide(slides[1], index=0)
    assert inserted.shapes[0].text == "B"
    assert (links[0].address, links[1].address) == ("https://example.test/a2", "https://example.test/b")
    before = presentation.to_bytes()
    with pytest.raises(RpptxError, match="control characters"):
        links[0].address = "https://example.test/\ninvalid"
    assert presentation.to_bytes() == before
    links[0].address = None
    assert links[0].address is None
    assert links[1].address == "https://example.test/b"
    assert inserted.shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.address == "https://example.test/b"
    parts = _package_parts(presentation.to_bytes())
    assert b"https://example.test/a2" not in parts["ppt/slides/_rels/slide1.xml.rels"]
    assert b"https://example.test/b" in parts["ppt/slides/_rels/slide2.xml.rels"]
    target = tmp_path / "moved-links.pptx"
    presentation.save(target)
    reopened = Presentation(target)
    assert [(slide.shapes[0].text, slide.shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.address) for slide in reopened.slides] == [("B", "https://example.test/b"), ("B", "https://example.test/b"), ("A", None)]


@pytest.mark.parametrize("foreign_state", ["live", "grid", "removed"])
def test_fx206_merge_rejects_foreign_identity_before_foreign_staleness(foreign_state):
    from rpptx import Presentation, StaleElementError

    receiver = Presentation()
    slide = receiver.slides.add_slide(receiver.slide_layouts[6])
    table = slide.shapes.add_table(2, 2, 0, 0, 1000, 1000).table
    cell = table.cell(0, 0)
    cell.text = "receiver"
    run = cell.text_frame.paragraphs[0].runs[0]
    foreign = Presentation()
    foreign_slide = foreign.slides.add_slide(foreign.slide_layouts[6])
    foreign_shape = foreign_slide.shapes.add_table(2, 2, 0, 0, 1000, 1000)
    foreign_table = foreign_shape.table
    foreign_cell = foreign_table.cell(0, 0)
    foreign_cell.text = "foreign"
    if foreign_state == "grid":
        foreign_table.rows.add_row()
    elif foreign_state == "removed":
        foreign_slide.shapes.remove(foreign_shape)
    if foreign_state == "live":
        assert foreign_cell.text == "foreign"
    else:
        with pytest.raises(StaleElementError):
            _ = foreign_cell.text
    before = receiver.to_bytes()
    with pytest.raises(ValueError, match="other_cell from different table"):
        cell.merge(foreign_cell)
    assert receiver.to_bytes() == before
    assert cell.text == run.text == "receiver"
    assert len(table.rows) == len(table.columns) == 2

# Prospective additions to the existing installed-test carrier. Not executed.
def test_fx206_group_ungroup_preserves_held_children_and_retires_only_group(tmp_path):
    from rpptx import Presentation, StaleElementError
    p=Presentation();slide=p.slides.add_slide(p.slide_layouts[6])
    a=slide.shapes.add_textbox(100,200,1000,800);a.text="A"
    b=slide.shapes.add_textbox(1500,900,700,600);b.text="B"
    paragraph=a.text_frame.paragraphs[0];run=paragraph.runs[0]
    ids=(a.shape_id,b.shape_id)
    group=slide.shapes.group([a,b])
    assert [s.shape_id for s in group.shapes]==list(ids)
    assert (a.text,b.text,run.text)==("A","B","A")
    assert (a.left,a.top,b.left,b.top)==(100,200,1500,900)
    returned=group.ungroup()
    assert [s.shape_id for s in returned]==list(ids)
    assert (a.text,b.text,run.text)==("A","B","A")
    with __import__('pytest').raises(StaleElementError,match="First invalidated by Shape.ungroup "):
        _=group.shape_id
    target=tmp_path/'grouped.pptx';p.save(target)
    actual=Presentation(target)
    assert [s.shape_id for s in actual.slides[0].shapes]==list(ids)
    assert [s.text for s in actual.slides[0].shapes]==["A","B"]


def test_fx206_noncontiguous_group_and_foreign_selection_are_atomic():
    import pytest
    from rpptx import Presentation, RpptxError
    p=Presentation();slide=p.slides.add_slide(p.slide_layouts[6])
    shapes=[slide.shapes.add_textbox(x,0,80,40) for x in [0,100,200]]
    foreign=Presentation();other=foreign.slides.add_slide(foreign.slide_layouts[6]).shapes.add_textbox(0,0,80,40)
    before=p.to_bytes()
    with pytest.raises(RpptxError, match=r"^group shapes failed: group selection must be contiguous$") as failure:
        slide.shapes.group([shapes[0],shapes[2]])
    assert type(failure.value) is RpptxError
    assert p.to_bytes()==before
    with pytest.raises(ValueError, match=r"^shape owner is not in this presentation collection$"):
        slide.shapes.align([shapes[0],other],"left")
    assert p.to_bytes()==before
    assert [s.left for s in shapes]==[0,100,200]


def test_fx206_signed_distribution_keeps_handles_and_noop_is_exact():
    from rpptx import Presentation
    p=Presentation();slide=p.slides.add_slide(p.slide_layouts[6])
    shapes=[slide.shapes.add_textbox(x,0,w,20) for x,w in [(0,100),(50,120),(80,100)]]
    for s in shapes:s.text="live"
    held=[s.text_frame.paragraphs[0].runs[0] for s in shapes]
    slide.shapes.distribute(shapes,"horizontal")
    assert [s.left for s in shapes]==[0,30,80]
    assert [r.text for r in held]==["live"]*3
    before=p.to_bytes();slide.shapes.distribute(shapes,"horizontal")
    assert p.to_bytes()==before
    assert [r.text for r in held]==["live"]*3

# Source-only prospective installed controls, no imports or execution here.
def test_fx206_glue_relocates_with_grouped_target_and_detaches_only_deleted_end(tmp_path):
    from rpptx import Presentation
    from rpptx.enum.shapes import MSO_CONNECTOR_TYPE
    p=Presentation();slide=p.slides.add_slide(p.slide_layouts[6])
    target=slide.shapes.add_shape('rect',100,200,100,80)
    sibling=slide.shapes.add_textbox(1000,1000,100,80)
    end=slide.shapes.add_shape('rect',900,700,100,80)
    c=slide.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT,150,200,950,700)
    c.begin_connect(target,0);c.end_connect(end,0)
    held=target.shape_id;other=end.shape_id
    group=slide.shapes.group([target,sibling])
    assert target.shape_id==held
    target.left=300;target.top=400
    assert (c.begin_x,c.begin_y,c.end_x,c.end_y)==(350,400,950,700)
    before=p.to_bytes();c.begin_connect(target,0);assert p.to_bytes()==before
    c.begin_x=123
    assert (c.begin_x,c.begin_y,c.end_x,c.end_y)==(123,400,950,700)
    assert b'stCxn' not in c.xml and b'endCxn' in c.xml
    slide.shapes.remove(group)
    assert (c.begin_x,c.begin_y,c.end_x,c.end_y)==(123,400,950,700)
    assert end.shape_id==other
    path=tmp_path/'connectors.pptx';p.save(path);q=Presentation(path)
    saved=q.slides[0].shapes[-1]
    assert (saved.begin_x,saved.begin_y,saved.end_x,saved.end_y)==(123,400,950,700)
    assert b'stCxn' not in saved.xml and b'endCxn' in saved.xml


def test_fx206_foreign_and_cross_slide_glue_refusal_precedes_read_and_is_atomic():
    import pytest
    from rpptx import Presentation,StaleElementError
    from rpptx.enum.shapes import MSO_CONNECTOR_TYPE
    p=Presentation();s=p.slides.add_slide(p.slide_layouts[6]);t=p.slides.add_slide(p.slide_layouts[6])
    c=s.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT,0,0,100,100)
    other=t.shapes.add_shape('rect',0,0,100,100)
    f=Presentation();fs=f.slides.add_slide(f.slide_layouts[6]);foreign=fs.shapes.add_shape('rect',0,0,100,100);fs.shapes.remove(foreign)
    before=p.to_bytes()
    with pytest.raises(ValueError,match='another presentation'):c.begin_connect(foreign,0)
    assert p.to_bytes()==before
    with pytest.raises(ValueError,match='another slide'):c.end_connect(other,0)
    assert p.to_bytes()==before
    local=s.shapes.add_shape('rect',200,300,100,80);s.shapes.remove(local);before=p.to_bytes()
    with pytest.raises(StaleElementError):c.begin_connect(local,0)
    assert p.to_bytes()==before


def test_fx206_failed_site_and_borrowed_table_resize_preserve_all_owners():
    import pytest
    from rpptx import Presentation, RpptxError
    from rpptx.enum.shapes import MSO_CONNECTOR_TYPE
    p=Presentation();s=p.slides.add_slide(p.slide_layouts[6]);target=s.shapes.add_shape('rect',100,200,100,80)
    target.text='kept';run=target.text_frame.paragraphs[0].runs[0]
    c=s.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT,0,0,100,100);before=p.to_bytes()
    with pytest.raises(
        RpptxError,
        match=rf"^connect connector failed: shape id {target.shape_id} has [0-9]+ connection sites, 0 to [0-9]+, not 99$",
    ) as failure:
        c.begin_connect(target,99)
    assert type(failure.value) is RpptxError
    assert p.to_bytes()==before and run.text=='kept'
    table=s.shapes.add_table(1,1,1000,2000,300,400);cell=table.table.cell(0,0);cell.text='table'
    table.table.columns[0].width=500
    assert table.width==500 and cell.text=='table'


def test_fx206_arrangement_slide_reference_and_group_owner_relocation_are_atomic():
    import pytest
    from rpptx import Presentation
    p=Presentation();s=p.slides.add_slide(p.slide_layouts[6])
    a=s.shapes.add_textbox(100,200,100,80);b=s.shapes.add_textbox(500,700,100,80)
    a.text='A';held=a.text_frame.paragraphs[0].runs[0]
    group=s.shapes.group([a,b]);before=p.to_bytes()
    with pytest.raises(ValueError):s.shapes.align([group,a],'left')
    assert p.to_bytes()==before and held.text=='A'
    group.shapes.align([a,b],'left')
    assert a.left==b.left and held.text=='A'
    children=group.ungroup()
    s.shapes.align([a],'center','slide')
    assert a.left==(p.slide_width-a.width)//2 and children[0].shape_id==a.shape_id
    s.shapes.distribute([a],'horizontal','slide')
    assert a.left==(p.slide_width-a.width)//2 and held.text=='A'


# Imported fixtures use real generated packages and legal typed group shells.
# These controls are prospective. No Python import/test has been run in preparation.
def _fx206_replace_slide_source(deck_bytes, rewrite):
    import io
    import zipfile
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(deck_bytes)) as original:
        with zipfile.ZipFile(output, "w") as patched:
            for item in original.infolist():
                body = original.read(item.filename)
                if item.filename == "ppt/slides/slide1.xml":
                    body = rewrite(body.decode("utf-8")).encode("utf-8")
                patched.writestr(item, body)
    return output.getvalue()


def _fx206_actual_slide_shapes(xml, count):
    import re
    members = re.findall(r"<p:sp(?:\s[^>]*)?>.*?</p:sp>", xml, re.DOTALL)
    assert len(members) == count, "generated fixture changed its source shape inventory"
    return members


def _fx206_group_source(members, *, id=50, scale=10000, prefix="", extra=""):
    return (
        f'<p:grpSp{prefix}><p:nvGrpSpPr><p:cNvPr id="{id}"></p:cNvPr>'
        '<p:cNvGrpSpPr></p:cNvGrpSpPr><p:nvPr></p:nvPr></p:nvGrpSpPr>'
        f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{scale}" cy="10000"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="10000" cy="10000"/>'
        f'</a:xfrm></p:grpSpPr>{members}{extra}</p:grpSp>'
    )


def test_fx206_imported_interleaved_raw_member_group_roundtrip_keeps_order_and_handles(tmp_path):
    import io
    import zipfile
    from rpptx import Presentation
    p = Presentation()
    slide = p.slides.add_slide(p.slide_layouts[6])
    slide.shapes.add_textbox(100, 200, 100, 80).text = "A"
    slide.shapes.add_textbox(500, 200, 100, 80).text = "B"
    raw = '<f:payload keep="&#65;"><f:value>  unchanged  </f:value></f:payload>'
    def insert_between(xml):
        second = _fx206_actual_slide_shapes(xml, 2)[1]
        return xml.replace('<p:sld ', '<p:sld xmlns:f="urn:outer" ', 1).replace(second, raw + second, 1)
    patched = _fx206_replace_slide_source(p.to_bytes(), insert_between)
    q = Presentation.from_bytes(patched)
    a, b = q.slides[0].shapes[0], q.slides[0].shapes[1]
    identities = [a.shape_id, b.shape_id]
    group = q.slides[0].shapes.group([a, b])
    assert [s.shape_id for s in group.shapes] == identities
    assert (a.text, b.text) == ("A", "B")
    source = group.xml.decode()
    assert source.index('>A<') < source.index(raw) < source.index('>B<')
    path = tmp_path / "raw-grouped.pptx"
    q.save(path)
    r = Presentation(path)
    r.slides[0].shapes[0].ungroup()
    target = tmp_path / "raw-ungrouped.pptx"
    r.save(target)
    reopened = Presentation(target)
    assert [s.shape_id for s in reopened.slides[0].shapes] == identities
    assert [s.text for s in reopened.slides[0].shapes] == ["A", "B"]
    with zipfile.ZipFile(target) as archive:
        source = archive.read("ppt/slides/slide1.xml").decode()
    assert source.index('>A<') < source.index(raw) < source.index('>B<')


def test_fx206_imported_identity_raw_ungroup_and_nested_namespace_roundtrip(tmp_path):
    import io
    import zipfile
    import xml.etree.ElementTree as ET
    from rpptx import Presentation, StaleElementError
    for shadow in [False, True]:
        p = Presentation()
        slide = p.slides.add_slide(p.slide_layouts[6])
        member = slide.shapes.add_textbox(100, 200, 100, 80)
        member.text = "retained"
        def insert_outer(xml):
            shape = _fx206_actual_slide_shapes(xml, 1)[0]
            inner_prefix = (' xmlns:f="urn:inner"' if shadow else '') + ' f:keep="value"'
            inner = _fx206_group_source(shape + '<f:payload keep="&#65;"/>', id=51, prefix=inner_prefix)
            body = '<f:before/>' + inner + '<f:after>  unchanged  </f:after>'
            outer = _fx206_group_source(body, prefix=' xmlns:f="urn:ancestor"')
            return xml.replace('<p:sld ', '<p:sld xmlns:f="urn:outer" ', 1).replace(shape, outer, 1)
        patched = _fx206_replace_slide_source(p.to_bytes(), insert_outer)
        q = Presentation.from_bytes(patched)
        outer_shape = q.slides[0].shapes[0]
        outer_children = outer_shape.shapes
        nested = outer_shape.shapes[0]
        held = nested.shapes[0]
        outer_shape.ungroup()
        assert nested.shape_id == 51 and held.text == "retained"
        for retired in (lambda: outer_shape.xml, lambda: len(outer_children)):
            with pytest.raises(StaleElementError, match=r"First invalidated by Shape.ungroup "):
                retired()
        target = tmp_path / f"promoted-{shadow}.pptx"
        q.save(target)
        r = Presentation(target)
        assert r.slides[0].shapes[0].shapes[0].text == "retained"
        with zipfile.ZipFile(target) as archive:
            xml = archive.read("ppt/slides/slide1.xml")
        root = ET.fromstring(xml)
        ns = "urn:inner" if shadow else "urn:ancestor"
        nested_root = root.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}grpSp')
        assert nested_root.attrib[f"{{{ns}}}keep"] == "value"
        payload = root.find(f".//{{{ns}}}payload")
        assert payload is not None and payload.attrib["keep"] == "A"
        tree = root.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree')
        assert [child.tag for child in list(tree)[2:]] == [
            "{urn:ancestor}before",
            "{http://schemas.openxmlformats.org/presentationml/2006/main}grpSp",
            "{urn:ancestor}after",
        ]
        assert root.find('.//{urn:ancestor}after').text == "  unchanged  "
        nested_children = nested.shapes
        held_paragraph = held.text_frame.paragraphs[0]
        q.slides[0].shapes.remove(nested)
        for retired in (lambda: nested.xml, lambda: len(nested_children),
                        lambda: held.text, lambda: held_paragraph.text):
            with pytest.raises(StaleElementError, match=r"First invalidated by ShapeCollection.remove "):
                retired()
        remaining = _package_parts(q.to_bytes())["ppt/slides/slide1.xml"]
        assert b'<f:before xmlns:f="urn:ancestor"/>' in remaining
        assert b'<f:after xmlns:f="urn:ancestor">  unchanged  </f:after>' in remaining
        remaining_tree = ET.fromstring(remaining).find(
            './/{http://schemas.openxmlformats.org/presentationml/2006/main}spTree'
        )
        assert [child.tag for child in list(remaining_tree)[2:]] == [
            "{urn:ancestor}before", "{urn:ancestor}after",
        ]


def test_fx206_imported_opaque_composition_and_group_owner_refusals_are_atomic():
    import io
    import pytest
    from rpptx import Presentation, RpptxError
    for scale, raw, extra in [
        (20000, '<f:payload/>', ''),
        (10000, '', '<p:extLst><p:ext uri="unknown"><f:payload keep="&#65;"/></p:ext></p:extLst>'),
    ]:
        p = Presentation()
        slide = p.slides.add_slide(p.slide_layouts[6])
        member = slide.shapes.add_textbox(100, 200, 100, 80)
        member.text = "kept"
        def insert_group(xml):
            shape = _fx206_actual_slide_shapes(xml, 1)[0]
            group_source = _fx206_group_source(shape + raw, scale=scale, extra=extra)
            return xml.replace('<p:sld ', '<p:sld xmlns:f="urn:outer" ', 1).replace(shape, group_source, 1)
        patched = _fx206_replace_slide_source(p.to_bytes(), insert_group)
        q = Presentation.from_bytes(patched)
        group = q.slides[0].shapes[0]
        held = group.shapes[0].text_frame.paragraphs[0].runs[0]
        before = q.to_bytes()
        with pytest.raises(RpptxError, match="unmodeled member|group-owned"):
            group.ungroup()
        assert q.to_bytes() == before
        assert group.shape_id == 50 and held.text == "kept"


@pytest.mark.parametrize("owner_kind", ["run", "paragraph-default", "table-cell"])
def test_fx206_highlight_live_authoring_noop_atomic_rgb_and_reopen(tmp_path, owner_kind):
    import rpptx

    presentation = _textbox_presentation(rpptx)
    owner = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    if owner_kind == "table-cell":
        table = presentation.slides[0].shapes.add_table(1, 1, 0, 0, rpptx.Inches(4), rpptx.Inches(2)).table
        table.cell(0, 0).text = "hello"
        owner = table.cell(0, 0).text_frame.paragraphs[0]
    font = owner.font if owner_kind == "paragraph-default" else owner.runs[0].font
    def body_xml(path):
        with zipfile.ZipFile(path) as archive:
            return archive.read("ppt/slides/slide1.xml").decode()
    first = tmp_path / "highlight-first.pptx"
    second = tmp_path / "highlight-second.pptx"
    presentation.save(first)
    before = body_xml(first)
    live = font.highlight
    assert live.rgb is None
    font.highlight = None
    presentation.save(second)
    assert body_xml(second) == before
    font.highlight = "12AB34"
    assert str(live.rgb) == "12AB34"
    assert owner.text == "hello"
    presentation.save(first)
    written = body_xml(first)
    import xml.etree.ElementTree as ET
    highlights = ET.fromstring(written).findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}highlight")
    assert len(highlights) == 1
    colour = highlights[0].find("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
    assert colour is not None and colour.attrib["val"] == "12AB34"
    assert 'val="12AB34"' in written
    for malformed in ("GG0000", (256, 0, 0), ("1", 0, 0)):
        with pytest.raises((TypeError, ValueError)):
            font.highlight = malformed
        presentation.save(second)
        assert body_xml(second) == written
        assert str(live.rgb) == "12AB34"
    # Legacy integer triples retain Python bool-as-integer channel compatibility.
    font.highlight = (True, False, True)
    assert str(live.rgb) == "010001"
    font.highlight = "12AB34"
    presentation.save(second)
    assert body_xml(second) == written
    reopened = rpptx.Presentation(first)
    fresh_owner = (reopened.slides[0].shapes[1].table.cell(0, 0).text_frame.paragraphs[0]
                   if owner_kind == "table-cell" else reopened.slides[0].shapes[0].text_frame.paragraphs[0])
    fresh = fresh_owner.font if owner_kind == "paragraph-default" else fresh_owner.runs[0].font
    assert str(fresh.highlight.rgb) == "12AB34"
    reopened.save(second)
    assert body_xml(second) == written
    font.highlight = None
    assert live.rgb is None
    presentation.save(second)
    assert "<a:highlight" not in body_xml(second)
    assert owner.text == "hello"


def test_fx206_highlight_alpha_foreign_payload_theme_and_atomic_clear(tmp_path):
    import rpptx

    source = tmp_path / "highlight-source.pptx"
    imported = tmp_path / "highlight-import.pptx"
    first = tmp_path / "highlight-output.pptx"
    second = tmp_path / "highlight-repeat.pptx"
    _textbox_presentation(rpptx).save(source)
    _replace_in_slide(source, imported, "<a:r><a:t>hello</a:t></a:r>",
        '<a:r><a:rPr><a:highlight xmlns:x="urn:fx206" x:owner="opaque"><a:srgbClr val="FF0000">'
        '<a:alpha val="50000"/><x:alpha val="123"/></a:srgbClr><!--keep--></a:highlight></a:rPr><a:t>hello</a:t></a:r>')
    presentation = rpptx.Presentation(imported)
    owner = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    font = owner.runs[0].font
    assert str(font.highlight.rgb) == "FF0000"
    presentation.save(first)
    before = _text_body_xml(first)
    font.highlight = "00FF00"
    presentation.save(first)
    written = _text_body_xml(first)
    assert 'val="00FF00"' in written and '<a:alpha val="50000"/>' in written
    assert '<x:alpha val="123"/>' in written and 'x:owner="opaque"' in written and '<!--keep-->' in written
    with pytest.raises(ValueError):
        font.highlight = None
    presentation.save(second)
    assert _text_body_xml(second) == written
    assert owner.text == "hello"
    reopened = rpptx.Presentation(first)
    reopened.save(second)
    assert _text_body_xml(second) == written
    # A fresh canonical owner exercises existing X197 theme/brightness authoring.
    fresh = _textbox_presentation(rpptx)
    live = fresh.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.highlight
    live.theme_color = "accent1"
    live.brightness = 0.25
    fresh.save(second)
    themed = _text_body_xml(second)
    assert 'schemeClr val="accent1"' in themed
    assert 'lumMod val="75000"' in themed and 'lumOff val="25000"' in themed
    assert before != written


@pytest.mark.parametrize("transform", [
    '<a:alpha val="50000" x:keep="&#65;"/>',
    '<a:alpha val="50000" x:keep="&#65;"> </a:alpha>',
])
def test_fx206_highlight_transform_attributes_refuse_destructive_public_edits(tmp_path, transform):
    import rpptx

    source = tmp_path / "transform-source.pptx"
    imported = tmp_path / "transform-import.pptx"
    saved = tmp_path / "transform-saved.pptx"
    repeat = tmp_path / "transform-repeat.pptx"
    _textbox_presentation(rpptx).save(source)
    _replace_in_slide(source, imported, "<a:r><a:t>hello</a:t></a:r>",
        '<a:r><a:rPr><a:highlight xmlns:x="urn:fx206"><a:srgbClr val="FF0000">'
        + transform + '</a:srgbClr></a:highlight></a:rPr><a:t>hello</a:t></a:r>')
    presentation = rpptx.Presentation(imported)
    owner = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    font = owner.runs[0].font
    live = font.highlight
    assert str(live.rgb) == "FF0000"
    presentation.save(saved)
    before = _text_body_xml(saved)
    assert transform in before
    for operation in ("rgb", "theme", "brightness", "clear"):
        with pytest.raises(ValueError, match="opaque transform attributes"):
            if operation == "rgb":
                font.highlight = "00FF00"
            elif operation == "theme":
                live.theme_color = "accent1"
            elif operation == "brightness":
                live.brightness = 0.25
            else:
                font.highlight = None
        presentation.save(repeat)
        assert _text_body_xml(repeat) == before
        assert str(live.rgb) == "FF0000" and owner.text == "hello"
    reopened = rpptx.Presentation(saved)
    fresh = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.highlight
    assert str(fresh.rgb) == "FF0000"
    reopened.save(repeat)
    assert _text_body_xml(repeat) == before

def _fx195_checksum(data):
    padded = data + b'\x00' * (-len(data) % 4)
    return sum(struct.unpack('>' + str(len(padded) // 4) + 'I', padded)) & 4294967295

def _fx195_renamed_font(source, family):
    data = bytearray(source)
    count = struct.unpack_from('>H', data, 4)[0]
    tables = {}
    for i in range(count):
        entry = 12 + 16 * i
        tag = bytes(data[entry:entry + 4])
        offset, length = struct.unpack_from('>II', data, entry + 8)
        assert offset + length <= len(data)
        tables[tag] = (entry, offset, length)
    _, off, length = tables[b'name']
    table = data[off:off + length]
    _, records, storage = struct.unpack_from('>HHH', table, 0)
    changes = 0
    for i in range(records):
        platform, encoding, language, ident, size, offset = struct.unpack_from('>HHHHHH', table, 6 + 12 * i)
        start = storage + offset
        assert start + size <= len(table)
        if platform not in (0, 1, 3):
            continue
        codec = 'mac_roman' if platform == 1 else 'utf-16-be'
        value = bytes(table[start:start + size]).decode(codec)
        edited = value.replace('Carlito', family).replace('Caladea', family)
        if edited != value:
            encoded = edited.encode(codec)
            assert len(encoded) == size
            table[start:start + size] = encoded
            changes += 1
    assert changes > 0
    data[off:off + length] = table
    _, head, headlen = tables[b'head']
    assert headlen >= 12
    data[head + 8:head + 12] = b'\x00' * 4
    for tag, (entry, offset, length) in tables.items():
        struct.pack_into('>I', data, entry + 4, _fx195_checksum(bytes(data[offset:offset + length])))
    struct.pack_into('>I', data, head + 8, 2981146554 - _fx195_checksum(bytes(data)) & 4294967295)
    assert _fx195_checksum(bytes(data)) == 2981146554
    for tag, (entry, offset, length) in tables.items():
        if tag not in (b'name', b'head'):
            assert data[offset:offset + length] == source[offset:offset + length]
    assert family.encode('utf-16-be') in table
    return bytes(data)

def _fx195_font_inputs():
    from pathlib import Path
    directory = Path(__file__).resolve().parents[2] / 'oxml-layout' / 'fonts'
    family = 'FX195CX'
    first = _fx195_renamed_font((directory / 'Carlito-Regular.ttf').read_bytes(), family)
    second = _fx195_renamed_font((directory / 'Caladea-Regular.ttf').read_bytes(), family)
    assert first != second
    return (family, first, second)

def _fx195_presentation(family):
    import rpptx as module
    import xml.etree.ElementTree as ET
    name = family
    obj = module.Presentation()
    slide = obj.slides.add_slide(obj.slide_layouts[6])
    shape = slide.shapes.add_textbox(914400, 914400, 5486400, 914400)
    shape.text = 'WWWW iiiii caller FX195 controlled geometry'
    run = obj.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    run.font.name = name
    obj.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.size = 20 * 12700
    obj.slides[0].notes_text = 'FX195 NOTES caller controlled geometry'
    raw = obj.to_bytes()
    parts = {}
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    notes = [n for n in parts if re.fullmatch('ppt/notesSlides/notesSlide\\d+\\.xml', n)]
    assert len(notes) == 1
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    root = ET.fromstring(parts[notes[0]])
    edited = 0
    for r in root.iter(ns + 'r'):
        text = r.find(ns + 't')
        if text is None or 'FX195 NOTES' not in (text.text or ''):
            continue
        prop = r.find(ns + 'rPr')
        if prop is None:
            prop = ET.Element(ns + 'rPr')
            r.insert(0, prop)
        latin = prop.find(ns + 'latin')
        if latin is None:
            latin = ET.SubElement(prop, ns + 'latin')
        latin.set('typeface', name)
        edited += 1
    assert edited == 1
    parts[notes[0]] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as z:
        for n, b in parts.items():
            z.writestr(n, b)
    obj = module.Presentation.from_bytes(buf.getvalue())
    assert 'FX195 NOTES' in obj.slides[0].notes_text
    assert name.encode() in parts[notes[0]]
    assert obj.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.name == name
    return obj

@pytest.mark.parametrize('method', ['to_pdf', 'render_slide_to_png', 'render_all_slides', 'text_layout', 'to_notes_pdf', 'render_all_notes'])
def test_fx195_installed_ppt_caller_selection_and_ordered_inputs(method, tmp_path):
    family, first, second = _fx195_font_inputs()
    presentation = _fx195_presentation(family)
    before = presentation.to_bytes()

    def invoke(kwargs):
        if method == 'render_slide_to_png':
            return presentation.render_slide_to_png(0, dpi=72, **kwargs)
        if method == 'render_all_slides':
            return presentation.render_all_slides(dpi=72, **kwargs)
        if method == 'render_all_notes':
            return presentation.render_all_notes(dpi=72, **kwargs)
        return getattr(presentation, method)(**kwargs)
    first_args, second_args = ({'fonts': [(family, first)]}, {'fonts': [(family, second)]})
    reference, alternative = (invoke(first_args), invoke(second_args))
    assert reference != alternative, f'{method} must use actual physical caller face in slide/notes paint or metrics'
    if method in ('to_pdf', 'to_notes_pdf'):
        assert b'FX195CX' in reference and reference.startswith(b'%PDF')
    if method == 'text_layout':
        assert reference and reference[0].lines
        assert reference[0].lines[0].bounds.width != alternative[0].lines[0].bounds.width
    first_dir, second_dir, empty = (tmp_path / 'first', tmp_path / 'second', tmp_path / 'empty')
    for directory in (first_dir, second_dir, empty):
        directory.mkdir()
    (first_dir / (family + '.ttf')).write_bytes(first)
    (second_dir / (family + '.ttf')).write_bytes(second)
    assert invoke({'font_dir': first_dir}) == reference
    assert invoke({'font_dirs': [second_dir, first_dir]}) == reference
    assert invoke({'font_dirs': [first_dir, second_dir]}) == alternative
    assert invoke({'font_dirs': [second_dir], **first_args}) == reference
    assert invoke(first_args) == reference
    ordinary = invoke({})
    assert invoke({'fonts': []}) == ordinary
    assert invoke({'font_dir': empty}) == ordinary
    assert invoke({'fonts': [], 'font_dirs': [empty]}) == ordinary
    assert invoke({'font_dirs': []}) == ordinary
    assert invoke({'font_dir': empty}) == ordinary
    with pytest.raises(ValueError, match='font_dir.*font_dirs'):
        invoke({'font_dir': first_dir, 'font_dirs': []})
    with pytest.raises(TypeError, match='font_dirs'):
        invoke({'font_dirs': str(first_dir)})
    with pytest.raises(FileNotFoundError, match='font directory.*missing'):
        invoke({'font_dirs': [tmp_path / 'missing']})
    with pytest.raises(NotADirectoryError, match='font directory'):
        invoke({'font_dirs': [first_dir / (family + '.ttf')]})
    bad = tmp_path / 'bad'
    bad.mkdir()
    (bad / 'broken.ttf').write_bytes(b'not a supported font')
    with pytest.raises(Exception, match='font file.*broken.ttf'):
        invoke({'font_dirs': [bad]})
    denied = tmp_path / 'denied'
    denied.mkdir()
    denied.chmod(0)
    try:
        from rpptx import RpptxError
        with pytest.raises((PermissionError, RpptxError), match=re.escape(str(denied))) as failure:
            invoke({'font_dirs': [denied]})
        assert 'permission denied' in str(failure.value).lower()
    finally:
        denied.chmod(448)
    assert presentation.to_bytes() == before

def _fx195_four_i_ink_width(source, points):
    tables = {}
    for offset in range(12, 12 + struct.unpack_from('>H', source, 4)[0] * 16, 16):
        tables[source[offset:offset + 4]] = struct.unpack_from('>II', source, offset + 8)
    cmap, _ = tables[b'cmap']
    glyph = None
    for record in range(struct.unpack_from('>H', source, cmap + 2)[0]):
        platform, encoding, relative = struct.unpack_from('>HHI', source, cmap + 4 + 8 * record)
        subtable = cmap + relative
        if platform not in (0, 3) or struct.unpack_from('>H', source, subtable)[0] != 4:
            continue
        segments = struct.unpack_from('>H', source, subtable + 6)[0] // 2
        for segment in range(segments):
            end = struct.unpack_from('>H', source, subtable + 14 + 2 * segment)[0]
            start = struct.unpack_from('>H', source, subtable + 16 + 2 * segments + 2 * segment)[0]
            if start <= ord('i') <= end:
                delta = struct.unpack_from('>h', source, subtable + 16 + 4 * segments + 2 * segment)[0]
                range_at = subtable + 16 + 6 * segments + 2 * segment
                relative = struct.unpack_from('>H', source, range_at)[0]
                glyph = ord('i') if relative == 0 else struct.unpack_from('>H', source, range_at + relative + 2 * (ord('i') - start))[0]
                glyph = (glyph + delta) % 65536
                break
        if glyph is not None:
            break
    assert glyph is not None and glyph > 0
    head, _ = tables[b'head']
    units = struct.unpack_from('>H', source, head + 18)[0]
    hhea, _ = tables[b'hhea']
    metrics = struct.unpack_from('>H', source, hhea + 34)[0]
    hmtx, _ = tables[b'hmtx']
    advance = struct.unpack_from('>H', source, hmtx + 4 * min(glyph, metrics - 1))[0]
    loca, _ = tables[b'loca']
    if struct.unpack_from('>h', source, head + 50)[0] == 0:
        glyph_offset = struct.unpack_from('>H', source, loca + 2 * glyph)[0] * 2
    else:
        glyph_offset = struct.unpack_from('>I', source, loca + 4 * glyph)[0]
    glyf, _ = tables[b'glyf']
    x_min = struct.unpack_from('>h', source, glyf + glyph_offset + 2)[0]
    x_max = struct.unpack_from('>h', source, glyf + glyph_offset + 6)[0]
    assert x_max > x_min and units > 0
    return (3 * advance + x_max - x_min) * points / units

def _fx195_painted_width(png):
    from PIL import Image
    image = Image.open(io.BytesIO(png)).convert('RGB')
    mask = image.convert('L').point(lambda value: 255 if value < 200 else 0)
    bounds = mask.getbbox()
    assert bounds is not None, 'caller text must actually paint'
    return bounds[2] - bounds[0]

def _fx195_collection(first, second):
    collection = bytearray(b'ttcf\x00\x01\x00\x00\x00\x00\x00\x02' + b'\x00' * 8)
    for index, source in enumerate((first, second)):
        offset = len(collection)
        struct.pack_into('>I', collection, 12 + 4 * index, offset)
        face = bytearray(source)
        count = struct.unpack_from('>H', face, 4)[0]
        for table in range(count):
            at = 12 + table * 16 + 8
            original = struct.unpack_from('>I', face, at)[0]
            struct.pack_into('>I', face, at, original + offset)
        collection.extend(face)
        collection.extend(b'\x00' * (-len(collection) % 4))
    return bytes(collection)

def test_fx195_ppt_caller_png_matches_independent_font_ink_and_collection(tmp_path):
    from pathlib import Path
    from rpptx import Presentation
    family, first, second = _fx195_font_inputs()
    presentation = Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    presentation.slides[0].shapes.add_textbox(914400, 914400, 5486400, 914400).text = 'iiii'
    presentation.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.name = family
    presentation.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.size = 36 * 12700
    before = presentation.to_bytes()
    paintings = []
    for source in (first, second):
        png = presentation.render_slide_to_png(0, dpi=72, fonts=[(family, source)])
        assert abs(_fx195_painted_width(png) - _fx195_four_i_ink_width(source, 36)) <= 2, 'selected physical font must paint the independent metric width'
        paintings.append(png)
    assert paintings[0] != paintings[1]
    ordinary = (Path(__file__).resolve().parents[2] / 'oxml-layout' / 'fonts' / 'Carlito-Regular.ttf').read_bytes()
    directory = tmp_path / 'collection'
    directory.mkdir()
    (directory / 'filename-alias.ttc').write_bytes(_fx195_collection(ordinary, second))
    assert presentation.render_slide_to_png(0, dpi=72, font_dirs=[directory]) == paintings[1]
    assert presentation.to_bytes() == before




def _fx195_variable_svg_geometry(svg, weight, variable):
    import xml.etree.ElementTree as ET
    tree = ET.fromstring(svg)
    for node in tree.iter():
        if not node.tag.endswith('}text'):
            continue
        assert 'stroke' not in node.attrib and 'transform' not in node.attrib
        style = node.attrib.pop('style', None)
        if variable:
            assert style is not None and style.startswith('font-synthesis:none;font-variation-settings:')
            selectors = dict((tag,float(value)) for tag,value in re.findall(r"'([A-Za-z0-9]{4})' ([-+0-9.]+)",style))
            assert selectors.get('wght') == weight
            assert set(selectors) <= {'wght','wdth'}
            if 'wdth' in selectors: assert selectors['wdth'] == 100
        else:
            assert style is None
    # Font definitions are representation differences. Actual and independent
    # static source bytes are authenticated separately before this comparison.
    # Keep every other element, attribute, text, position and child order.
    for parent in tree.iter():
        for child in list(parent):
            if child.tag.endswith('}style'):
                assert child.text and '@font-face{' in child.text
                parent.remove(child)
    return ET.tostring(tree,encoding='unicode')


def _fx195_save_variable_svg_artifact(actual, expected, weight, directory):
    import base64
    import hashlib
    import json
    import os
    from pathlib import Path
    import xml.etree.ElementTree as ET
    authority = json.loads((Path(os.environ['F_X195_VARIABLE_ORACLES'])/'receipt.json').read_text())
    row = next(row for row in authority['controls'] if row['script']=='Thai' and row['weight']==weight)
    node = next(node for node in ET.fromstring(actual).iter() if node.tag.endswith('}text') and ''.join(node.itertext())==row['text'])
    family = node.get('font-family')
    assert "'wght' " + str(weight) in node.get('style',''), 'selected coordinates must be emitted on this real owner'
    sources = re.findall(r"@font-face\{font-family:'([^']+)';src:url\('data:[^;]+;base64,([^']+)'\)",actual)
    font = next(base64.b64decode(value) for name,value in sources if name==family)
    assert hashlib.sha256(font).hexdigest()==row['source_sha256']
    expected_sources = re.findall(r"@font-face\{font-family:'([^']+)';src:url\('data:[^;]+;base64,([^']+)'\)",expected)
    assert len(expected_sources)==1
    assert hashlib.sha256(base64.b64decode(expected_sources[0][1])).hexdigest()==row['output_sha256']
    a,e=directory/'variable-installed.svg',directory/'variable-independent-static.svg'
    a.write_text(actual); e.write_text(expected)
    root=ET.fromstring(actual)
    dimensions=[float(root.get(key).removesuffix('pt')) for key in ('width','height')]
    assert dimensions==[round(value) for value in dimensions], '72dpi companion requires integer page dimensions'
    artifact={'case':'variable-'+str(weight),'svg':str(a),'svgSHA':hashlib.sha256(a.read_bytes()).hexdigest(),'expectedSVG':str(e),'expectedSVGSHA':hashlib.sha256(e.read_bytes()).hexdigest(),'fontFamily':family,'fontSHA':row['source_sha256'],'fontIndex':0,'oracleFont':row['output'],'oracleFontSHA':row['output_sha256'],'width':int(dimensions[0]),'height':int(dimensions[1]),'threshold':.99,'selector':"'wght' "+str(weight)}
    # The emitted SVG remains unchanged. The serial native stage qualifies its
    # selected-instance font resolver separately because usvg ignores CSS axes.
    (directory/'variable-installed-svg-artifact.json').write_text(json.dumps(artifact,indent=2)+'\n')
# Candidate additions for each existing installed test carrier, not a new module.
# Root must bind the supplied retained-oracle directory and source receipts.

def _fx195_style_oracles():
    import hashlib
    import json
    from pathlib import Path
    import os
    directory = Path(os.environ['F_X195_VARIABLE_ORACLES'])
    receipt = json.loads((directory / 'receipt.json').read_text())
    assert receipt['tools']['hb-shape'] == 'hb-shape (HarfBuzz) 14.4.0'
    rows = [row for row in receipt['controls'] if row['script'] == 'Thai']
    assert {row['weight'] for row in rows} == {400, 700}
    selected = {}
    for row in rows:
        variable = Path(row['source']).read_bytes()
        static = Path(row['output']).read_bytes()
        assert hashlib.sha256(variable).hexdigest() == row['source_sha256']
        assert hashlib.sha256(static).hexdigest() == row['output_sha256']
        assert row['source_instance_positioning_matches_static'] is True
        assert row['source_static_svg_path_topology_matches'] is True
        assert row['source_static_svg_path_max_coordinate_delta_font_units'] <= 0.008
        assert sum(glyph['ax'] for glyph in row['shape']) > 0
        selected[row['weight']] = variable, static, row['text']
    assert selected[400][0] == selected[700][0]
    assert selected[400][1] != selected[700][1]
    return selected


def _fx195_style_observation(value, method, directory, label):
    # Full pixel canvases are compared, not registered/cropped masks that could
    # hide changed advances, page origins, selected page order or clipping.
    import subprocess
    from PIL import Image
    if method in ('to_pdf', 'to_notes_pdf'):
        source = directory / (label + '.pdf')
        source.write_bytes(value)
        target = directory / label
        result = subprocess.run(
            [_fx195_pinned_poppler(), '-r', '72', '-png', str(source), str(target)],
            capture_output=True, timeout=120, check=False,
        )
        assert result.returncode == 0, result.stderr.decode(errors='replace')
        images = sorted(directory.glob(label + '-*.png'), key=lambda p: int(p.stem.rsplit('-', 1)[1]))
        assert images
        return [(image.size, image.tobytes()) for path in images for image in [Image.open(path).convert('RGBA')]]
    if method in ('render_page_to_png', 'render_slide_to_png'):
        image = Image.open(io.BytesIO(value)).convert('RGBA')
        assert image.convert('RGB').convert('L').getextrema()[0] < 200
        return image.size, image.tobytes()
    if method in ('render_all_pages', 'render_pages', 'render_all_slides', 'render_all_notes'):
        assert value
        return [_fx195_style_observation(png, 'render_page_to_png', directory, label) for png in value]
    if method == 'render_page_to_svg':
        # This is positioning/searchable-text evidence, not installed SVG paint.
        # Raw caller font data can differ between the independently instanced
        # static reference and variable source. Preserve every geometry/style
        # attribute while ignoring only the actual embedded font data URI.
        assert 'data:font/' in value
        return re.sub(r'data:font/[^;]+;base64,[A-Za-z0-9+/=]+', 'FONT_BYTES', value)
    if method in ('layout', 'layout_page'):
        return value
    assert method == 'text_layout'
    assert value and value[0].lines
    return [(line.bounds.x, line.bounds.y, line.bounds.width, line.bounds.height) for item in value for line in item.lines]


def _fx195_variable_pdf_source_equivalence(actual_pdf, static_pdf, wrong_pdf, weight, parts, directory, notes):
    # Per-owner paint is qualified only after exact source/Tf/CID/geometry proof.
    import hashlib
    import json
    import math
    import os
    from pathlib import Path
    from PIL import Image
    authority=json.loads((Path(os.environ['F_X195_VARIABLE_ORACLES'])/'receipt.json').read_text())
    row=next(r for r in authority['controls'] if r['script']=='Thai' and r['weight']==weight)
    source=Path(row['output']).read_bytes()
    assert hashlib.sha256(source).hexdigest()==row['output_sha256']
    glyph_ids=list(dict.fromkeys(g['g'] for g in row['shape']))
    images = [_fx195_pdf_style_images(pdf,directory,label) for pdf,label in [(actual_pdf,'variable-qualified'),(static_pdf,'static-qualified'),(wrong_pdf,'wrong-qualified')]]
    assert all(len(pages)==1 for pages in images)
    actual,expected,wrong = [pages[0] for pages in images]
    assert actual.size==expected.size==wrong.size
    source_case={'case':'variable-weight','expected':source,'files':[source],'axis':None,'text':row['text']}
    fixture=io.BytesIO()
    with zipfile.ZipFile(fixture,'w') as package:
        for name,data in parts.items():package.writestr(name,data)
    anchors=_fx195_style_fixture_owners(fixture.getvalue(),source_case,actual.size,notes,points=24)
    actual_proof=_fx195_variable_pdf_reachable_proof(actual_pdf,anchors,row,source)
    static_proof=_fx195_variable_pdf_reachable_proof(static_pdf,anchors,row,source)
    assert actual_proof['size']==static_proof['size']==actual.size
    assert actual_proof['reachableStreams']==static_proof['reachableStreams'], 'all reachable drawing commands must stay exact'
    records = {gid:_fx195_source_glyph(source,'',glyph_index=gid) for gid in glyph_ids}
    extrema=[]; pen=0
    for shaped in row['shape']:
        record=records[shaped['g']]; x0,y0,x1,y1=record['bounds']
        extrema.extend([(pen+shaped['dx']+x0,shaped['dy']+y0),(pen+shaped['dx']+x1,shaped['dy']+y1)])
        pen+=shaped['ax']
    assert pen==row['total_x_advance']
    rois=[]; observations=[]
    for anchor in anchors:
        x,y=anchor['origin'];factor=24*anchor['scale']/row['units_per_em']
        roi=(math.floor(x+min(v[0] for v in extrema)*factor)-2,math.floor(y-max(v[1] for v in extrema)*factor)-2,math.ceil(x+max(v[0] for v in extrema)*factor)+2,math.ceil(y-min(v[1] for v in extrema)*factor)+2)
        assert 0<=roi[0]<roi[2]<=actual.width and 0<=roi[1]<roi[3]<=actual.height
        score=_fx195_mask_ssim(actual,expected,roi)
        weight_score=_fx195_mask_ssim(actual,wrong,roi)
        displaced=Image.new(expected.mode,expected.size,'white')
        displaced.paste(expected,(3,0))
        geometry_score=_fx195_mask_ssim(actual,displaced,roi)
        assert score>=.99, (anchor,score)
        assert weight_score<.99 and geometry_score<.99, (anchor,weight_score,geometry_score)
        rois.append(roi); observations.append({'owner':anchor.get('owner'),'roi':roi,'positive':score,'wrongWeight':weight_score,'wrongOrigin':geometry_score})
    # No global white-page score. All pixels outside independent source bounds
    # remain exactly equal, retaining clipping and unrelated owner evidence.
    outside_actual=actual.copy();outside_expected=expected.copy()
    from PIL import ImageDraw
    for image in (outside_actual,outside_expected):
        draw=ImageDraw.Draw(image)
        for x0,y0,x1,y1 in rois:draw.rectangle((x0,y0,x1-1,y1-1),fill='white')
    assert outside_actual.tobytes()==outside_expected.tobytes()
    (directory/'variable-pdf-source-equivalence.json').write_text(json.dumps({'weight':weight,'originalStaticSHA':row['output_sha256'],'subsetGeometryAndAdvanceExact':True,'owners':observations,'outsideSourceROIsExact':True},indent=2)+'\n')


def _fx195_variable_pdf_reachable_proof(pdf, owners, row, source):
    # Existing qualified raw PDF syntax, no new PDF-reader dependency.
    # Each reachable shown CID is tied to its scoped Tf font and expected HB
    # source glyph at independently declared OOXML placement.
    import math
    import hashlib
    import zlib
    objects = {}
    for match in re.finditer(rb'(?m)(\d+)\s+0\s+obj\b(.*?)\bendobj\b',pdf,re.S):
        oid,body=int(match[1]),match[2]
        stream=re.search(rb'\bstream\r?\n(.*?)\r?\nendstream\b',body,re.S)
        header=body[:stream.start()] if stream else body
        data=stream[1] if stream else None
        if data is not None and b'/FlateDecode' in header:data=zlib.decompress(data)
        assert oid not in objects;objects[oid]=(header,data)
    pages=[oid for oid,(header,_) in objects.items() if re.search(rb'/Type\s*/Page\b',header)]
    assert len(pages)==1
    identity=(1.,0.,0.,1.,0.,0.)
    def compose(a,b):
        return (a[0]*b[0]+a[2]*b[1],a[1]*b[0]+a[3]*b[1],a[0]*b[2]+a[2]*b[3],a[1]*b[2]+a[3]*b[3],a[0]*b[4]+a[2]*b[5]+a[4],a[1]*b[4]+a[3]*b[5]+a[5])
    def dictionary_after(header,key):
        found=re.search(rb'/'+key+rb'\s*(\d+\s+0\s+R|<<)',header)
        if found is None:return None
        if found[1]!=b'<<':return objects[int(found[1].split()[0])][0]
        start=found.end()-2;depth=0
        for token in re.finditer(rb'<<|>>',header[start:]):
            depth+=1 if token[0]==b'<<' else -1
            if depth==0:return header[start:start+token.end()]
        raise AssertionError('unclosed resource dictionary')
    def resources(oid,inherited=b''):
        header=objects[oid][0];own=dictionary_after(header,b'Resources')
        if own is not None:return own
        parent=re.search(rb'/Parent\s+(\d+)\s+0\s+R',header)
        return resources(int(parent[1]),inherited) if parent else inherited
    def literal(token):
        if token.startswith(b'<'):
            return bytes.fromhex(token[1:-1].decode())
        assert token.startswith(b'(') and token.endswith(b')')
        data=token[1:-1];out=bytearray();i=0
        escapes={ord('n'):10,ord('r'):13,ord('t'):9,ord('b'):8,ord('f'):12}
        while i<len(data):
            value=data[i];i+=1
            if value==92:
                assert i<len(data);value=data[i];i+=1
                if 48<=value<=55:
                    digits=bytes([value])
                    while i<len(data) and len(digits)<3 and 48<=data[i]<=55:digits+=bytes([data[i]]);i+=1
                    value=int(digits,8)
                else:value=escapes.get(value,value)
            out.append(value)
        return bytes(out)
    glyph_ids=list(dict.fromkeys(g['g'] for g in row['shape']))
    cid_map={gid:i+1 for i,gid in enumerate(glyph_ids)}
    expected_records={gid:_fx195_source_glyph(source,'',glyph_index=gid) for gid in glyph_ids}
    points=24;upem=row['units_per_em'];expected={};pen=0
    for owner in owners:expected[owner['owner']]=[]
    for shaped in row['shape']:
        for owner in owners:
            factor=points*owner['scale']/upem
            expected[owner['owner']].append({'cid':cid_map[shaped['g']],'sourceGID':shaped['g'],'x':owner['origin'][0]+(pen+shaped['dx'])*factor,'y':owner['origin'][1]-shaped['dy']*factor})
        pen+=shaped['ax']
    assert pen==row['total_x_advance']
    seen={name:[] for name in expected};font_cache={};reachable=[]
    def font_data(oid):
        if oid in font_cache:return font_cache[oid]
        font=objects[oid][0]
        descendant=re.search(rb'/DescendantFonts\s*\[\s*(\d+)\s+0\s+R\s*\]',font);assert descendant
        cid=objects[int(descendant[1])][0];assert re.search(rb'/CIDToGIDMap\s*/Identity\b',cid)
        descriptor=re.search(rb'/FontDescriptor\s+(\d+)\s+0\s+R',cid);assert descriptor
        stream_ref=re.search(rb'/FontFile2\s+(\d+)\s+0\s+R',objects[int(descriptor[1])][0]);assert stream_ref
        sfnt=objects[int(stream_ref[1])][1];assert sfnt is not None
        width_match=re.search(rb'/W\s*\[\s*1\s*\[([^]]+)\]\s*\]',cid);assert width_match
        widths=[float(v) for v in width_match[1].split()];assert len(widths)==len(glyph_ids)
        tables={sfnt[o:o+4]:struct.unpack_from('>II',sfnt,o+8) for o in range(12,12+16*struct.unpack_from('>H',sfnt,4)[0],16)}
        assert struct.unpack_from('>H',sfnt,tables[b'maxp'][0]+4)[0]==len(glyph_ids)+1
        for subset_gid,gid in enumerate(glyph_ids,1):
            actual=_fx195_source_glyph(sfnt,'',glyph_index=subset_gid);reference=expected_records[gid]
            assert actual['upem']==reference['upem']==upem==1000
            assert actual['advance']==reference['advance']==widths[subset_gid-1]
            assert actual['bounds']==reference['bounds'] and actual['instruction_bytes']==reference['instruction_bytes']==0
            path=lambda record:[_fx195_source_path_commands(c,lambda point:point[:2]) for c in record['contours']]
            assert path(actual)==path(reference)
        font_cache[oid]=(sfnt,widths);return font_cache[oid]
    tokens=re.compile(rb'\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]*>|/[^\s<>\[\](){}%]+|[-+]?(?:\d+\.\d*|\.\d+|\d+)|[A-Za-z][A-Za-z0-9*\'\"]*|[\[\]]')
    def execute(oid,ctx,res,path):
        assert len(path)<16 and oid not in path
        data=objects[oid][1];assert data is not None
        reachable.append((oid,data));ctx=dict(ctx);stack=[];args=[];in_text=False;cursor=0.
        for token in tokens.findall(data):
            if token[:1] in b'/([<' or re.fullmatch(rb'[-+]?(?:\d+\.\d*|\.\d+|\d+)',token) or token==b']':args.append(token);continue
            op=token
            if op==b'q':stack.append(dict(ctx))
            elif op==b'Q':assert stack;ctx=stack.pop()
            elif op==b'cm':assert len(args)==6;ctx['ctm']=compose(ctx['ctm'],tuple(float(v) for v in args))
            elif op==b'BT':in_text=True;ctx['tm']=identity;cursor=0.
            elif op==b'ET':in_text=False
            elif op==b'Tf':
                assert len(args)==2;ctx['size']=float(args[1]);ctx['font']=args[0]
            elif op==b'Tm':assert len(args)==6;ctx['tm']=tuple(float(v) for v in args);cursor=0.
            elif op in (b'Td',b'TD'):
                assert len(args)==2;ctx['tm']=compose(ctx['tm'],(1.,0.,0.,1.,float(args[0]),float(args[1])));cursor=0.
            elif op==b'T*':assert ctx.get('size')!=points, 'unqualified selected line-break positioning'
            elif op in (b'Tj',b'TJ') and in_text and ctx.get('size')==points:
                fonts=dictionary_after(res,b'Font');assert fonts is not None
                found=re.search(re.escape(ctx['font'])+rb'\s+(\d+)\s+0\s+R',fonts);assert found
                font_oid=int(found[1]);sfnt,widths=font_data(font_oid)
                for value in args:
                    if value in (b'[',b']'):continue
                    if not value.startswith((b'(',b'<')):
                        cursor-=float(value)*points/1000;continue
                    encoded=literal(value);assert len(encoded)%2==0
                    for cid in struct.unpack('>'+str(len(encoded)//2)+'H',encoded):
                        assert 1<=cid<=len(widths)
                        local=compose(ctx['ctm'],ctx['tm']);canvas=compose((1.,0.,0.,-1.,0.,float(ctx['pageHeight'])),local)
                        x=canvas[0]*cursor+canvas[4];y=canvas[1]*cursor+canvas[5]
                        assert all(math.isfinite(v) for v in (x,y))
                        matches=[]
                        for owner in owners:
                            consumed=len(seen[owner['owner']]);sequence=expected[owner['owner']]
                            if consumed==len(sequence):continue
                            target=sequence[consumed]
                            if abs(x-target['x'])<=.00005 and abs(y-target['y'])<=.00005 and cid==target['cid']:
                                assert abs(canvas[0]-owner['scale'])<=.00001 and abs(canvas[1])<=.00001 and abs(canvas[2])<=.00001 and abs(canvas[3]+owner['scale'])<=.00001
                                matches.append(owner['owner'])
                        assert len(matches)==1, ('reachable CID must match exactly one independent source owner',cid,x,y)
                        seen[matches[0]].append({'cid':cid,'fontObject':font_oid,'fontSHA':hashlib.sha256(sfnt).hexdigest(),'position':(x,y)})
                        cursor+=widths[cid-1]*points/1000
            elif op==b'Do':
                assert len(args)==1;xobjects=dictionary_after(res,b'XObject');assert xobjects is not None
                found=re.search(re.escape(args[0])+rb'\s+(\d+)\s+0\s+R',xobjects);assert found
                child=int(found[1]);header=objects[child][0]
                if re.search(rb'/Subtype\s*/Form\b',header):
                    state=dict(ctx);matrix=re.search(rb'/Matrix\s*\[([^]]+)\]',header)
                    if matrix:state['ctm']=compose(state['ctm'],tuple(float(v) for v in matrix[1].split()))
                    execute(child,state,resources(child,res),path+(oid,))
                else:assert re.search(rb'/Subtype\s*/Image\b',header)
            args=[]
        assert not stack and not in_text
    page=pages[0];header=objects[page][0]
    box=re.search(rb'/MediaBox\s*\[([^]]+)\]',header);assert box
    box=tuple(float(v) for v in box[1].split());assert box[:2]==(0.,0.)
    content=re.search(rb'/Contents\s*(\d+\s+0\s+R|\[[^]]+\])',header);assert content
    refs=[int(v) for v in re.findall(rb'(\d+)\s+0\s+R',content[1])];assert len(refs)==1
    execute(refs[0],{'ctm':identity,'size':None,'tm':identity,'pageHeight':box[3]},resources(page),())
    assert all(len(seen[name])==len(row['shape']) for name in expected), 'all and only source owners must paint the complete HB glyph sequence'
    return {'size':box[2:],'seen':seen,'reachableStreams':reachable,'fontPayloads':font_cache}


@pytest.mark.parametrize('weight', [400, 700])
@pytest.mark.parametrize('method', ['to_pdf', 'render_slide_to_png', 'render_all_slides', 'text_layout', 'to_notes_pdf', 'render_all_notes'])
def test_fx195_installed_ppt_variable_instances_match_independent_static(method, weight, tmp_path):
    # _fx195_presentation creates a genuine notes part. Set the actual slide
    # and physical notes runs to the same source text/style before measuring.
    import xml.etree.ElementTree as ET
    variable, static, text = _fx195_style_oracles()[weight]
    family = 'FX195CallerThai'
    presentation = _fx195_presentation(family)
    with zipfile.ZipFile(io.BytesIO(presentation.to_bytes())) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    counts = {'slide': 0, 'notes': 0}
    for name in list(parts):
        kind = 'slide' if re.fullmatch(r'ppt/slides/slide\d+\.xml', name) else 'notes' if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml', name) else None
        if kind is None:
            continue
        root = ET.fromstring(parts[name])
        for run in root.iter(ns + 'r'):
            properties = run.find(ns + 'rPr')
            latin = properties.find(ns + 'latin') if properties is not None else None
            if latin is None or latin.get('typeface') != family:
                continue
            properties.set('b', '1' if weight == 700 else '0')
            properties.set('sz', '2400')
            run.find(ns + 't').text = text
            counts[kind] += 1
        parts[name] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    assert counts == {'slide': 1, 'notes': 1}, 'both real caller font owners must exist'
    package = io.BytesIO()
    with zipfile.ZipFile(package, 'w') as z:
        for name, data in parts.items():
            z.writestr(name, data)
    import rpptx
    presentation = rpptx.Presentation.from_bytes(package.getvalue())
    before = presentation.to_bytes()

    def invoke(data):
        kwargs = {'fonts': [(family, data)]}
        if method == 'render_slide_to_png':
            return presentation.render_slide_to_png(0, dpi=72, **kwargs)
        if method == 'render_all_slides':
            return presentation.render_all_slides(dpi=72, **kwargs)
        if method == 'render_all_notes':
            return presentation.render_all_notes(dpi=72, **kwargs)
        return getattr(presentation, method)(**kwargs)

    static_output, actual_output = invoke(static), invoke(variable)
    expected = _fx195_style_observation(static_output, method, tmp_path, 'static')
    actual = _fx195_style_observation(actual_output, method, tmp_path, 'variable')
    wrong_output = invoke(_fx195_style_oracles()[1100 - weight][1])
    if method in ('to_pdf', 'to_notes_pdf'):
        _fx195_variable_pdf_source_equivalence(actual_output,static_output,wrong_output,weight,parts,tmp_path,notes=method=='to_notes_pdf')
    else:
        assert actual == expected
    assert _fx195_style_observation(invoke(variable), method, tmp_path, 'repeat') == actual
    wrong_weight = _fx195_style_observation(wrong_output, method, tmp_path, 'wrong-weight')
    assert actual != wrong_weight, 'each selected entrypoint must distinguish wrong physical coordinates'
    assert presentation.to_bytes() == before

# Append to the existing installed carriers only. No production code.

_FX195_STYLE_CASES = ('genuine-italic', 'ital-axis', 'slnt-axis', 'synthetic-bold', 'synthetic-italic', 'synthetic-combined')


def _fx195_style_case(case):
    import hashlib
    import json
    import os
    from pathlib import Path
    font_dir = Path(__file__).resolve().parents[2] / 'oxml-layout' / 'fonts'
    if case == 'genuine-italic':
        regular = (font_dir / 'Carlito-Regular.ttf').read_bytes()
        italic = (font_dir / 'Carlito-Italic.ttf').read_bytes()
        # The runtime admission binds both bundled source hashes and LICENSE-Carlito.
        guards = json.loads(Path(os.environ['F_X195_STYLE_FONT_GUARDS']).read_text())
        for name, data in [('Carlito-Regular.ttf', regular), ('Carlito-Italic.ttf', italic)]:
            assert hashlib.sha256(data).hexdigest() == guards[name]
        assert regular != italic
        return {'case': case, 'family': 'FX195StyleCaller', 'text': 'iiii', 'files': [regular, italic], 'ordinary': [regular], 'expected': italic, 'bold': False, 'italic': True, 'axis': None, 'synthesis': False}
    axis_dir = Path(os.environ['F_X195_AXIS_ORACLES'])
    receipt = json.loads((axis_dir / 'receipt.json').read_text())
    key = 'ital' if case == 'ital-axis' else 'slnt' if case == 'slnt-axis' else 'none'
    control = next(row for row in receipt['controls'] if row['axis'] == key)
    data = (axis_dir / (key + '.ttf')).read_bytes()
    assert hashlib.sha256(data).hexdigest() == control['input_sha256']
    # Authenticated complete R18 tuple protocol, including phantom points.
    directory = {data[o:o+4]: struct.unpack_from('>II', data, o+8) for o in range(12,12+16*struct.unpack_from('>H',data,4)[0],16)}
    if key != 'none':
        gv = directory[b'gvar'][0]
        assert struct.unpack_from('>III',data,gv+20) == (0,0,29)
        assert struct.unpack_from('>HHHH',data,gv+32) == (1,10,19,0xA000)
        assert data[gv+42:gv+44] == b'\x00\x47'
        assert list(struct.unpack_from('>8h',data,gv+44)) == [0,0,140,140,0,0,0,0]
        assert data[gv+60] == 0x87
    else:
        assert b'fvar' not in directory and b'gvar' not in directory
    assert control['exit_code'] == 0
    measured = json.loads(control['stdout'])
    assert measured == [{'g': 1, 'cl': 0, 'dx': 0, 'dy': 0, 'ax': 600, 'ay': 0}]
    if key != 'none':
        static = (axis_dir / (key + '-static.ttf')).read_bytes()
        assert hashlib.sha256(static).hexdigest() == control['static']['sha256']
        assert control['static']['glyph1_points'] == [[50, 0], [450, 0], [590, 700], [190, 700]]
        assert control['static']['fvar_absent'] is True
    return {'case': case, 'family': 'FX195StyleCaller', 'text': 'A', 'files': [data], 'ordinary': [data], 'expected': None, 'bold': case in ('synthetic-bold', 'synthetic-combined'), 'italic': case != 'synthetic-bold', 'axis': key if key != 'none' else None, 'synthesis': key == 'none'}


def _fx195_i_advance(source, points):
    tables = {}
    for offset in range(12, 12 + struct.unpack_from('>H', source, 4)[0] * 16, 16):
        tables[source[offset:offset + 4]] = struct.unpack_from('>II', source, offset + 8)
    cmap, _ = tables[b'cmap']
    glyph = None
    for record in range(struct.unpack_from('>H', source, cmap + 2)[0]):
        platform, encoding, relative = struct.unpack_from('>HHI', source, cmap + 4 + 8 * record)
        subtable = cmap + relative
        if platform not in (0, 3) or struct.unpack_from('>H', source, subtable)[0] != 4:
            continue
        segments = struct.unpack_from('>H', source, subtable + 6)[0] // 2
        for segment in range(segments):
            end = struct.unpack_from('>H', source, subtable + 14 + 2 * segment)[0]
            start = struct.unpack_from('>H', source, subtable + 16 + 2 * segments + 2 * segment)[0]
            if start <= ord('i') <= end:
                delta = struct.unpack_from('>h', source, subtable + 16 + 4 * segments + 2 * segment)[0]
                range_at = subtable + 16 + 6 * segments + 2 * segment
                relative = struct.unpack_from('>H', source, range_at)[0]
                glyph = ord('i') if relative == 0 else struct.unpack_from('>H', source, range_at + relative + 2 * (ord('i') - start))[0]
                glyph = (glyph + delta) % 65536
                break
        if glyph is not None:
            break
    assert glyph is not None and glyph > 0
    head, _ = tables[b'head']
    units = struct.unpack_from('>H', source, head + 18)[0]
    hhea, _ = tables[b'hhea']
    metrics = struct.unpack_from('>H', source, hhea + 34)[0]
    hmtx, _ = tables[b'hmtx']
    advance = struct.unpack_from('>H', source, hmtx + 4 * min(glyph, metrics - 1))[0]
    loca, _ = tables[b'loca']
    if struct.unpack_from('>h', source, head + 50)[0] == 0:
        glyph_offset = struct.unpack_from('>H', source, loca + 2 * glyph)[0] * 2
    else:
        glyph_offset = struct.unpack_from('>I', source, loca + 4 * glyph)[0]
    glyf, _ = tables[b'glyf']
    x_min = struct.unpack_from('>h', source, glyf + glyph_offset + 2)[0]
    x_max = struct.unpack_from('>h', source, glyf + glyph_offset + 6)[0]
    assert x_max > x_min and units > 0
    return advance * points / units


def _fx195_pdf_style_anchors(pdf, page_size, case, expected_preview_scale=None):
    # Parse reachable content only. Compose PDF q/Q and cm, text matrices and
    # Form XObject matrices before converting bottom-up PDF to canvas points.
    # No actual bitmap bounds, registration or first-match selection is used.
    import math
    import zlib
    objects = {}
    for match in re.finditer(rb'(?m)(\d+)\s+0\s+obj\b(.*?)\bendobj\b', pdf, re.S):
        oid, body = int(match[1]), match[2]
        stream = re.search(rb'\bstream\r?\n(.*?)\r?\nendstream\b', body, re.S)
        header = body[:stream.start()] if stream else body
        data = stream[1] if stream else None
        if data is not None and b'/FlateDecode' in header:
            data = zlib.decompress(data)
        assert oid not in objects
        objects[oid] = (header, data)
    pages = [oid for oid,(header,_) in objects.items() if re.search(rb'/Type\s*/Page\b',header)]
    assert len(pages) == 1, 'single physical page fixture is required'
    identity = (1.,0.,0.,1.,0.,0.)
    def compose(a,b):
        return (a[0]*b[0]+a[2]*b[1], a[1]*b[0]+a[3]*b[1], a[0]*b[2]+a[2]*b[3], a[1]*b[2]+a[3]*b[3], a[0]*b[4]+a[2]*b[5]+a[4], a[1]*b[4]+a[3]*b[5]+a[5])
    def dictionary_after(header,key):
        found = re.search(rb'/'+key+rb'\s*(\d+\s+0\s+R|<<)',header)
        if found is None: return None
        value=found[1]
        if value != b'<<': return objects[int(value.split()[0])][0]
        start=found.end()-2; depth=0
        for token in re.finditer(rb'<<|>>',header[start:]):
            depth += 1 if token[0]==b'<<' else -1
            if depth==0: return header[start:start+token.end()]
        raise AssertionError('unclosed actual PDF resource dictionary')
    def resources(oid, inherited=b''):
        header=objects[oid][0]
        own=dictionary_after(header,b'Resources')
        if own is not None: return own
        parent=re.search(rb'/Parent\s+(\d+)\s+0\s+R',header)
        return resources(int(parent[1]),inherited) if parent else inherited
    tokens=re.compile(rb'\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]*>|/[^\s<>\[\](){}%]+|[-+]?(?:\d+\.\d*|\.\d+|\d+)|[A-Za-z][A-Za-z0-9*\'\"]*|[\[\]]')
    anchors=[]
    def execute(oid,ctx,res,path):
        assert len(path)<16 and oid not in path, 'cyclic or unsupported deep form fixture'
        header,data=objects[oid]
        assert data is not None
        ctx=dict(ctx); ctx['history']=list(ctx['history']); stack=[]; args=[]; in_text=False; captured=False
        for token in tokens.findall(data):
            if token[:1] in b'/([<' or re.fullmatch(rb'[-+]?(?:\d+\.\d*|\.\d+|\d+)',token) or token==b']':
                args.append(token); continue
            op=token
            if op==b'q': stack.append(dict(ctx,history=list(ctx['history'])))
            elif op==b'Q': assert stack; ctx=stack.pop()
            elif op==b'cm':
                assert len(args)==6
                matrix=tuple(float(v) for v in args)
                assert all(math.isfinite(v) for v in matrix)
                ctx['ctm']=compose(ctx['ctm'],matrix);ctx['history'].append(matrix)
            elif op==b'BT': in_text=True;captured=False;ctx['tm']=identity
            elif op==b'ET': in_text=False
            elif op==b'Tf': assert len(args)==2;ctx['size']=float(args[1])
            elif op==b'Tm': assert len(args)==6;ctx['tm']=tuple(float(v) for v in args)
            elif op in (b'Td',b'TD',b'T*'):
                assert ctx.get('size')!=48, 'fixture text positioning needs a separately qualified adapter'
            elif op in (b'Tj',b'TJ') and in_text and ctx.get('size')==48 and not captured:
                assert args, '48pt owner must actually show text'
                captured=True; tm=ctx['tm']
                assert tm[:4]==(1.,0.,0.,-1.), 'nominal fixture requires upright local text matrix'
                placement=ctx['ctm']
                if case['synthesis'] and case['italic']:
                    style=(1.,0.,-math.tan(math.radians(12)),1.,tm[5]*math.tan(math.radians(12)),0.)
                    assert ctx['history'] and all(abs(a-b)<=.00002 for a,b in zip(ctx['history'][-1],style)), 'actual synthetic shear must equal independent style policy'
                    placement=identity
                    for matrix in ctx['history'][:-1]: placement=compose(placement,matrix)
                canvas=compose((1.,0.,0.,-1.,0.,float(page_size[1])),placement)
                assert abs(canvas[1])<.00001 and abs(canvas[2])<.00001 and canvas[0]>0 and abs(canvas[0]-canvas[3])<.00001, 'source fixture requires uniform nonrotated external placement'
                origin=(canvas[0]*tm[4]+canvas[2]*tm[5]+canvas[4],canvas[1]*tm[4]+canvas[3]*tm[5]+canvas[5])
                assert all(math.isfinite(v) for v in origin)
                anchors.append({'origin':origin,'scale':canvas[0],'matrix':canvas,'path':path+(oid,)})
            elif op==b'Do':
                assert len(args)==1 and args[0].startswith(b'/')
                xobjects=dictionary_after(res,b'XObject')
                assert xobjects is not None
                target=re.search(re.escape(args[0])+rb'\s+(\d+)\s+0\s+R',xobjects)
                assert target is not None
                child=int(target[1]);child_header=objects[child][0]
                if re.search(rb'/Subtype\s*/Form\b',child_header):
                    state=dict(ctx,history=list(ctx['history']))
                    matrix=re.search(rb'/Matrix\s*\[([^]]+)\]',child_header)
                    if matrix:
                        transform=tuple(float(v) for v in matrix[1].split());assert len(transform)==6
                        state['ctm']=compose(state['ctm'],transform);state['history'].append(transform)
                    execute(child,state,resources(child,res),path+(oid,))
                else:
                    assert re.search(rb'/Subtype\s*/Image\b',child_header), 'unknown XObject fixture'
            args=[]
        assert not stack and not in_text
    header=objects[pages[0]][0]
    content=re.search(rb'/Contents\s*(\d+\s+0\s+R|\[[^]]+\])',header)
    assert content is not None
    refs=[int(v) for v in re.findall(rb'(\d+)\s+0\s+R',content[1])]
    assert refs
    # This producer uses one content stream. Independent resources and forms
    # are supported, while multiple streams need state carried across streams.
    assert len(refs)==1, 'multiple content-stream fixture needs qualified continuous state'
    execute(refs[0],{'ctm':identity,'history':[],'size':None,'tm':identity},resources(pages[0]),())
    if expected_preview_scale is None:
        assert len(anchors)==1 and abs(anchors[0]['scale']-1)<.00001
        anchors[0]['owner']='body'
    else:
        assert len(anchors)==2, 'real notes body and real slide preview must both paint'
        body=[a for a in anchors if abs(a['scale']-1)<.00001]
        preview=[a for a in anchors if abs(a['scale']-expected_preview_scale)<.00001]
        assert len(body)==len(preview)==1 and body[0] is not preview[0]
        assert 0<expected_preview_scale<1
        body[0]['owner']='notes-body';preview[0]['owner']='slide-preview'
    return anchors


def _fx195_notes_preview_scale(parts):
    import xml.etree.ElementTree as ET
    p='{http://schemas.openxmlformats.org/presentationml/2006/main}'
    a='{http://schemas.openxmlformats.org/drawingml/2006/main}'
    presentation=ET.fromstring(parts['ppt/presentation.xml'])
    size=presentation.find(p+'sldSz');assert size is not None
    slide_width,slide_height=float(size.get('cx')),float(size.get('cy'))
    groups=[]
    for names in (r'ppt/notesSlides/notesSlide\d+\.xml',r'ppt/notesMasters/notesMaster\d+\.xml'):
        found=[]
        for name,data in parts.items():
            if not re.fullmatch(names,name):continue
            for shape in ET.fromstring(data).iter(p+'sp'):
                ph=shape.find(p+'nvSpPr/'+p+'nvPr/'+p+'ph')
                if ph is None or ph.get('type')!='sldImg':continue
                ext=shape.find(p+'spPr/'+a+'xfrm/'+a+'ext')
                if ext is not None:found.append((float(ext.get('cx')),float(ext.get('cy'))))
        if found:groups=found;break
    assert len(groups)==1, 'one actual source-established preview rectangle is required'
    return min(groups[0][0]/slide_width,groups[0][1]/slide_height)


def _fx195_pinned_poppler():
    # Both actual product PDF and independently authored reference use this binary.
    import hashlib
    import subprocess
    import os
    from pathlib import Path
    executable = Path(os.environ['F_X195_POPPLER'])
    assert executable.is_file()
    if hashlib.sha256(executable.read_bytes()).hexdigest() != 'd813ee83025433b811b5b7e982dc061751a1ea6cf03773107c398d2a597b70d9':
        import runpy
        installer = Path(__file__).resolve().parents[3] / 'scripts/install_pinned_poppler.py'
        runpy.run_path(str(installer))['authenticated_tool'](
            Path(os.environ['F_X195_POPPLER_BUILD_RECEIPT']),
            os.environ['F_X195_POPPLER_BUILD_RECEIPT_SHA'], executable,
        )
    version = subprocess.run([str(executable), '-v'], capture_output=True, timeout=10, check=False)
    assert version.returncode == 0 and version.stderr.startswith(b'pdftoppm version 26.01.0\n')
    return str(executable)


def _fx195_independent_style_paint(case, page_size, origin, font_scale=1., *, negative=None):
    # Independently authored PDF from original/static physical font bytes.
    # No product writer, shaper, subsetter, image registration or mask-derived ROI.
    import math
    import os
    import subprocess
    import tempfile
    from pathlib import Path
    from PIL import Image
    source = case['expected'] if case['case'] == 'genuine-italic' else case['files'][0]
    if case['axis']:
        source = (Path(os.environ['F_X195_AXIS_ORACLES']) / (case['axis']+'-static.ttf')).read_bytes()
    bold = case['bold']
    shear = math.tan(math.radians(12)) if case['synthesis'] and case['italic'] else 0.
    if negative == 'style':
        if case['case'] == 'genuine-italic':
            source = case['ordinary'][0]
        elif case['axis']:
            source = case['files'][0]
        else:
            bold, shear = False, 0.
    points = 48 * font_scale
    x, y = origin
    if negative == 'origin':
        x += 3.
    if negative == 'shear':
        assert case['synthesis'] and case['italic']
        shear = math.tan(math.radians(6))
    def u16(offset):
        return struct.unpack_from('>H', source, offset)[0]
    tables = {source[o:o+4]:struct.unpack_from('>II',source,o+8) for o in range(12,12+16*u16(4),16)}
    cmap = tables[b'cmap'][0]
    char = ord(case['text'][0])
    glyph = None
    for record in range(u16(cmap+2)):
        platform, encoding, relative = struct.unpack_from('>HHI',source,cmap+4+record*8)
        sub = cmap+relative
        if platform not in (0,3) or u16(sub) != 4:
            continue
        count = u16(sub+6)//2
        for segment in range(count):
            end,start = u16(sub+14+2*segment),u16(sub+16+2*count+2*segment)
            if start <= char <= end:
                delta = struct.unpack_from('>h',source,sub+16+4*count+2*segment)[0]
                at = sub+16+6*count+2*segment
                offset = u16(at)
                raw = char if offset == 0 else u16(at+offset+2*(char-start))
                glyph = (raw+delta)%65536
                break
        if glyph is not None:
            break
    assert glyph is not None and glyph > 0
    upem = u16(tables[b'head'][0]+18)
    metric_count = u16(tables[b'hhea'][0]+34)
    advance = u16(tables[b'hmtx'][0]+4*min(glyph,metric_count-1))*points/upem
    long_loca = struct.unpack_from('>h',source,tables[b'head'][0]+50)[0] == 1
    glyph_offset = struct.unpack_from('>I' if long_loca else '>H',source,tables[b'loca'][0]+glyph*(4 if long_loca else 2))[0]
    if not long_loca:
        glyph_offset *= 2
    bounds = struct.unpack_from('>hhhh',source,tables[b'glyf'][0]+glyph_offset+2)
    radius = points/48 if bold else 0.
    corners = [(gx*points/upem+shear*gy*points/upem,gy*points/upem) for gx in (bounds[0],bounds[2]) for gy in (bounds[1],bounds[3])]
    left = math.floor(x+min(px for px,py in corners)-radius*(1+abs(shear)))-3
    right = math.ceil(x+(len(case['text'])-1)*advance+max(px for px,py in corners)+radius*(1+abs(shear)))+3
    top = math.floor(y-max(py for px,py in corners)-radius)-3
    bottom = math.ceil(y-min(py for px,py in corners)+radius)+3
    assert 0 <= left < right <= page_size[0] and 0 <= top < bottom <= page_size[1]
    content = f'q\n1 0 0 -1 0 {page_size[1]} cm\n0 0 0 rg\n'
    if shear:
        content += f'1 0 {-shear} 1 {shear*y} 0 cm\n'
    if bold:
        content += f'0 0 0 RG\n{points/24} w\n1 J\n1 j\n2 Tr\n'
    encoded = (glyph.to_bytes(2,'big')*len(case['text'])).hex()
    content += f'BT\n/F0 {points} Tf\n1 0 0 -1 {x} {y} Tm\n<{encoded}> Tj\nET\nQ\n'
    raw = content.encode('ascii')
    width = advance*1000/points
    box = ' '.join(str(v*1000/upem) for v in struct.unpack_from('>hhhh',source,tables[b'head'][0]+36))
    objects = [
        b'<< /Type /Catalog /Pages 2 0 R >>',
        b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        f'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_size[0]} {page_size[1]}] /Resources << /Font << /F0 5 0 R >> >> /Contents 4 0 R >>'.encode(),
        f'<< /Length {len(raw)} >>\nstream\n'.encode()+raw+b'endstream',
        b'<< /Type /Font /Subtype /Type0 /BaseFont /IndependentSource /Encoding /Identity-H /DescendantFonts [6 0 R] >>',
        f'<< /Type /Font /Subtype /CIDFontType2 /BaseFont /IndependentSource /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 7 0 R /CIDToGIDMap /Identity /DW {width} /W [{glyph} [{width}]] >>'.encode(),
        f'<< /Type /FontDescriptor /FontName /IndependentSource /Flags 4 /FontBBox [{box}] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile2 8 0 R >>'.encode(),
        f'<< /Length {len(source)} /Length1 {len(source)} >>\nstream\n'.encode()+source+b'\nendstream',
    ]
    pdf, offsets = b'%PDF-1.7\n', []
    for index,obj in enumerate(objects,1):
        offsets.append(len(pdf))
        pdf += f'{index} 0 obj\n'.encode()+obj+b'\nendobj\n'
    xref = len(pdf)
    pdf += f'xref\n0 {len(objects)+1}\n0000000000 65535 f \n'.encode()
    pdf += b''.join(f'{offset:010} 00000 n \n'.encode() for offset in offsets)
    pdf += f'trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n'.encode()
    with tempfile.TemporaryDirectory(prefix='fx195-independent-vector-') as temp:
        path = Path(temp)/'source.pdf'
        path.write_bytes(pdf)
        version = subprocess.run([_fx195_pinned_poppler(),'-v'],capture_output=True,timeout=10,check=False)
        assert version.returncode == 0 and b'pdftoppm version 26.01.0' in version.stderr, 'qualified Poppler26.01.0 is required'
        result = subprocess.run([_fx195_pinned_poppler(),'-r','72','-png',str(path),str(Path(temp)/'source')],capture_output=True,timeout=120,check=False)
        assert result.returncode == 0, result.stderr.decode(errors='replace')
        image = Image.open(Path(temp)/'source-1.png').convert('RGBA')
        image.load()
    return image,(left,top,right,bottom),len(case['text'])*advance




def _fx195_mask_ssim(actual, expected, roi):
    # The ROI comes from independent source bounds and the emitted text anchor,
    # never from image registration or an actual-mask bounding box.
    a = list(actual.convert('L').crop(roi).getdata())
    b = list(expected.convert('L').crop(roi).getdata())
    assert len(a) == len(b) and len(a) > 1
    assert min(a) < 200 and min(b) < 200
    ma,mb = sum(a)/len(a),sum(b)/len(b)
    va = sum((v-ma)**2 for v in a)/(len(a)-1)
    vb = sum((v-mb)**2 for v in b)/(len(b)-1)
    cov = sum((x-ma)*(y-mb) for x,y in zip(a,b))/(len(a)-1)
    return ((2*ma*mb+2.55**2)*(2*cov+7.65**2))/((ma*ma+mb*mb+2.55**2)*(va+vb+7.65**2))


def _fx195_pdf_style_images(pdf, directory, label):
    import subprocess
    from PIL import Image
    path = directory / (label+'.pdf')
    path.write_bytes(pdf)
    result = subprocess.run([_fx195_pinned_poppler(),'-r','72','-png',str(path),str(directory/label)], capture_output=True, timeout=120, check=False)
    assert result.returncode == 0, result.stderr.decode(errors='replace')
    files = sorted(directory.glob(label+'-*.png'), key=lambda p:int(p.stem.rsplit('-',1)[1]))
    assert len(files) == 1, 'single actual styled owner fixture must produce one physical page'
    return [Image.open(path).convert('RGBA') for path in files]


def _fx195_svg_style_artifact(svg, case, expected, cubic, roi, origin, advance, directory):
    import base64
    import hashlib
    import json
    import xml.etree.ElementTree as ET
    import os
    from pathlib import Path
    tree = ET.fromstring(svg)
    texts = [node for node in tree.iter() if node.tag.endswith('}text') and ''.join(node.itertext()) == case['text']]
    assert len(texts) == 1
    node = texts[0]
    xs = [float(x) for x in node.get('x','').split()]
    assert xs and abs(xs[0]-origin[0]) <= 0.000001
    assert abs(float(node.get('y'))-origin[1]) <= 0.000001
    if len(xs) > 1:
        step = advance / len(case['text'])
        assert len(xs) == len(case['text'])
        assert all(abs(value-(origin[0]+i*step)) <= 0.000001 for i,value in enumerate(xs))
    else:
        assert len(case['text']) == 1
    assert node.get('textLength') is None, 'these nominal scalar fixtures need no identity textLength oracle workaround'
    family = node.get('font-family')
    matches = re.findall(r"@font-face\{font-family:'([^']+)';src:url\('data:[^;]+;base64,([^']+)'\)", svg)
    source = [base64.b64decode(value) for name,value in matches if name == family]
    assert len(source) == 1 and source[0] in case['files'], 'actual embedded font must be the physical requested source'
    svg_path = directory/(case['case']+'-installed.svg')
    expected_path = directory/(case['case']+'-independent.png')
    svg_path.write_text(svg)
    expected.save(expected_path)
    assert cubic.size == expected.size
    cubic_path = directory/(case['case']+'-independent-cubic.png')
    cubic.save(cubic_path)
    static_path = str(Path(os.environ['F_X195_AXIS_ORACLES']) / (case['axis']+'-static.ttf')) if case['axis'] else None
    static_sha = hashlib.sha256(Path(static_path).read_bytes()).hexdigest() if static_path else None
    artifact = {'oracleFont':static_path,'oracleFontSHA':static_sha,'case':case['case'],'svg':str(svg_path),'svgSHA':hashlib.sha256(svg_path.read_bytes()).hexdigest(),'expected':str(expected_path),'expectedSHA':hashlib.sha256(expected_path.read_bytes()).hexdigest(),'expectedQuadratic':str(expected_path),'expectedQuadraticSHA':hashlib.sha256(expected_path.read_bytes()).hexdigest(),'expectedCubic':str(cubic_path),'expectedCubicSHA':hashlib.sha256(cubic_path.read_bytes()).hexdigest(),'fontFamily':family,'fontSHA':hashlib.sha256(source[0]).hexdigest(),'fontIndex':0,'roi':list(roi),'origin':list(origin),'advance':advance,'width':expected.width,'height':expected.height,'axis':case['axis'],'smallStaticControl':case['axis'] is not None,'threshold':0.99}
    # Serial ROOT native companion receives this immutable receipt after the
    # installed stage. There is deliberately no Cargo fallback here.
    receipt = directory/(case['case']+'-installed-svg-artifact.json')
    receipt.write_text(json.dumps(artifact,indent=2)+'\n')


_FX195_REFERENCE_RENDERER_SOURCE = 'use std::{env, fs};\nuse resvg::{usvg, tiny_skia};\n\nfn count_paths(group: &usvg::Group) -> Result<usize, String> {\n    let mut count = 0;\n    for node in group.children() {\n        match node {\n            usvg::Node::Group(child) => count += count_paths(child)?,\n            usvg::Node::Path(_) => count += 1,\n            usvg::Node::Text(_) | usvg::Node::Image(_) => {\n                return Err("reference must contain paths only".into());\n            }\n        }\n    }\n    Ok(count)\n}\n\nfn main() -> Result<(), String> {\n    let args: Vec<String> = env::args().collect();\n    if args.len() != 3 { return Err("renderer INPUT.svg OUTPUT.png".into()); }\n    let bytes = fs::read(&args[1]).map_err(|e| e.to_string())?;\n    if bytes.len() > 1_000_000 { return Err("reference SVG exceeds admitted bound".into()); }\n    let options = usvg::Options::default(); // Empty font database, no system-font load.\n    let tree = usvg::Tree::from_data(&bytes, &options).map_err(|e| e.to_string())?;\n    let paths = count_paths(tree.root())?;\n    if paths == 0 || paths > 16 { return Err("unexpected path-only owner count".into()); }\n    let size = tree.size().to_int_size();\n    let mut image = tiny_skia::Pixmap::new(size.width(), size.height())\n        .ok_or_else(|| "invalid physical canvas".to_owned())?;\n    image.fill(tiny_skia::Color::WHITE);\n    resvg::render(&tree, tiny_skia::Transform::identity(), &mut image.as_mut());\n    image.save_png(&args[2]).map_err(|e| e.to_string())?;\n    println!("paths={paths} width={} height={} font_database=empty", size.width(), size.height());\n    Ok(())\n}\n'

def _fx195_source_glyph(source,char, *, glyph_index=None):
 import struct,hashlib
 def u16(at):return struct.unpack_from('>H',source,at)[0]
 tables={source[o:o+4]:struct.unpack_from('>II',source,o+8) for o in range(12,12+16*u16(4),16)}
 assert source[:4]==b'\x00\x01\x00\x00'
 assert all(o+n<=len(source) for o,n in tables.values())
 g=glyph_index
 if g is None:
  cmap=tables[b'cmap'][0];g=None
  for i in range(u16(cmap+2)):
   platform,encoding,rel=struct.unpack_from('>HHI',source,cmap+4+8*i);at=cmap+rel
   if platform not in (0,3) or u16(at)!=4:continue
   count=u16(at+6)//2
   for j in range(count):
    hi=u16(at+14+2*j);lo=u16(at+16+2*count+2*j)
    if lo<=ord(char)<=hi:
     delta=struct.unpack_from('>h',source,at+16+4*count+2*j)[0];r=at+16+6*count+2*j;off=u16(r)
     if off==0:g=(ord(char)+delta)%65536
     else:
      raw=u16(r+off+2*(ord(char)-lo));g=(raw+delta)%65536 if raw else 0
     break
   if g is not None:break
 assert g and g<u16(tables[b'maxp'][0]+4)
 head=tables[b'head'][0];upem=u16(head+18);long=struct.unpack_from('>h',source,head+50)[0];assert long in (0,1)
 def loc(k):return struct.unpack_from('>I' if long else '>H',source,tables[b'loca'][0]+k*(4 if long else 2))[0]*(1 if long else 2)
 def decode(index,depth=0):
  assert depth<8
  first,last=loc(index),loc(index+1);assert 0<=first<last<=tables[b'glyf'][1]
  raw=source[tables[b'glyf'][0]+first:tables[b'glyf'][0]+last];n=struct.unpack_from('>h',raw)[0];bounds=struct.unpack_from('>hhhh',raw,2)
  if n<0:
   at=10;contours=[];components=[]
   while True:
    flags,child=struct.unpack_from('>HH',raw,at);at+=4
    assert flags&2,'bounded composite requires explicit XY offsets, not point attachment'
    args=struct.unpack_from('>hh' if flags&1 else '>bb',raw,at);at+=4 if flags&1 else 2
    xx,xy,yx,yy=1.,0.,0.,1.
    if flags&8:xx=yy=struct.unpack_from('>h',raw,at)[0]/16384;at+=2
    elif flags&64:xx,yy=[v/16384 for v in struct.unpack_from('>hh',raw,at)];at+=4
    elif flags&128:xx,xy,yx,yy=[v/16384 for v in struct.unpack_from('>hhhh',raw,at)];at+=8
    assert not flags&4,'ROUND_XY_TO_GRID offset rule requires separate qualification'
    assert not flags&2048 or (xx,xy,yx,yy)==(1.,0.,0.,1.),'scaled offset requires separate qualification'
    child_contours,_,_,_=decode(child,depth+1)
    contours.extend([[(xx*x+xy*y+args[0],yx*x+yy*y+args[1],on) for x,y,on in contour] for contour in child_contours])
    components.append({'glyph':child,'flags':flags,'offset':args,'matrix':[xx,xy,yx,yy]})
    if not flags&32:break
   instructions=struct.unpack_from('>H',raw,at)[0] if flags&256 else 0
   if flags&256:at+=2+instructions
   assert at<=len(raw)
   return contours,bounds,instructions,raw
  assert n>0
  ends=list(struct.unpack_from('>'+str(n)+'H',raw,10));assert ends==sorted(set(ends))
  at=10+2*n;instructions=struct.unpack_from('>H',raw,at)[0];at+=2+instructions;flags=[]
  while len(flags)<ends[-1]+1:
   f=raw[at];at+=1;flags.append(f)
   if f&8:repeat=raw[at];at+=1;flags.extend([f]*repeat)
  assert len(flags)==ends[-1]+1 and not any(f&128 for f in flags)
  coords=[]
  for short,same in [(2,16),(4,32)]:
   vals=[];v=0
   for f in flags:
    if f&short:
     d=raw[at];at+=1;v+=d if f&same else -d
    elif not f&same:v+=struct.unpack_from('>h',raw,at)[0];at+=2
    vals.append(v)
   coords.append(vals)
  assert at<=len(raw)
  points=[(x,y,bool(f&1)) for x,y,f in zip(*coords,flags)];contours=[];start=0
  for end in ends:contours.append(points[start:end+1]);start=end+1
  return contours,bounds,instructions,raw
 contours,bounds,instructions,raw=decode(g)
 metric_count=u16(tables[b'hhea'][0]+34);advance=u16(tables[b'hmtx'][0]+4*min(g,metric_count-1))
 return {'source_sha256':hashlib.sha256(source).hexdigest(),'glyph':g,'upem':upem,'advance':advance,'bounds':bounds,'instruction_bytes':instructions,'glyph_bytes_sha256':hashlib.sha256(raw).hexdigest(),'contours':contours}


def _fx195_source_path_commands(points,transform,cubic=False):
 mid=lambda a,b:((a[0]+b[0])/2,(a[1]+b[1])/2,True)
 if points[0][2]:start=points[0];seq=points[1:]
 elif points[-1][2]:start=points[-1];seq=points[:-1]
 else:start=mid(points[-1],points[0]);seq=points[:]
 seq=seq+[start];current=start;result=[]
 def values(*pts):return ' '.join(f'{v:.12f}' for p in pts for v in transform(p))
 result.append('M '+values(start));i=0
 while i<len(seq):
  p=seq[i]
  if p[2]:result.append('L '+values(p));current=p;i+=1
  else:
   nxt=seq[i+1];end=nxt if nxt[2] else mid(p,nxt)
   if cubic:
    a=(current[0]+2*(p[0]-current[0])/3,current[1]+2*(p[1]-current[1])/3);b=(end[0]+2*(p[0]-end[0])/3,end[1]+2*(p[1]-end[1])/3)
    result.append('C '+values(a,b,end))
   else:result.append('Q '+values(p,end))
   current=end;i+=2 if nxt[2] else 1
 result.append('Z');return ' '.join(result)



def _fx195_independent_path_paint(case, page_size, origin, font_scale=1., *, negative=None, cubic=False):
    # Authored original/static source contours, no production contour/shaper/decoder.
    import hashlib, json, math, os, subprocess, tempfile
    import xml.etree.ElementTree as ET
    from pathlib import Path
    from PIL import Image
    source = case['expected'] if case['case'] == 'genuine-italic' else case['files'][0]
    if case['axis']:
        source = (Path(os.environ['F_X195_AXIS_ORACLES']) / (case['axis']+'-static.ttf')).read_bytes()
    bold = case['bold']
    shear = math.tan(math.radians(12)) if case['synthesis'] and case['italic'] else 0.
    if negative == 'style':
        if case['case'] == 'genuine-italic': source = case['ordinary'][0]
        elif case['axis']: source = case['files'][0]
        else: bold, shear = False, 0.
    points = 48 * font_scale
    record = _fx195_source_glyph(source, case['text'][0])
    assert len(set(case['text'])) == 1
    factor = points / record['upem']
    # Preserve the independently read metric and physical unit-conversion order.
    advance = record['advance'] * points / record['upem']
    x,y = origin
    if negative == 'origin': x += 3.
    if negative == 'shear':
        assert case['synthesis'] and case['italic']
        shear = math.tan(math.radians(6))
    radius = points/48 if bold else 0.
    bounds = record['bounds']
    corners = [(gx*factor+shear*gy*factor,gy*factor) for gx in (bounds[0],bounds[2]) for gy in (bounds[1],bounds[3])]
    roi = (math.floor(x+min(px for px,py in corners)-radius*(1+abs(shear)))-3,
           math.floor(y-max(py for px,py in corners)-radius)-3,
           math.ceil(x+(len(case['text'])-1)*advance+max(px for px,py in corners)+radius*(1+abs(shear)))+3,
           math.ceil(y-min(py for px,py in corners)+radius)+3)
    assert 0 <= roi[0] < roi[2] <= page_size[0] and 0 <= roi[1] < roi[3] <= page_size[1]
    ns = 'http://www.w3.org/2000/svg'
    ET.register_namespace('', ns)
    tree = ET.Element('{'+ns+'}svg', {'width':str(page_size[0]),'height':str(page_size[1]),'viewBox':f'0 0 {page_size[0]} {page_size[1]}'})
    for index in range(len(case['text'])):
        data = ' '.join(_fx195_source_path_commands(contour, lambda p:(x+index*advance+p[0]*factor,y-p[1]*factor),cubic) for contour in record['contours'])
        attrs = {'d':data,'fill':'#000000','fill-rule':'nonzero'}
        if bold: attrs.update({'stroke':'#000000','stroke-width':f'{points/24:.12f}','stroke-linecap':'round','stroke-linejoin':'round','paint-order':'fill stroke'})
        if shear: attrs['transform'] = f'matrix(1 0 {-shear:.12f} 1 {shear*y:.12f} 0)'
        ET.SubElement(tree,'{'+ns+'}path',attrs)
    raw = ET.tostring(tree,encoding='utf-8',xml_declaration=True)
    assert all(node.tag=='{'+ns+'}path' for node in tree)
    assert not any(token in raw for token in [b'<text',b'<image',b'<use',b'<defs',b'font',b'href',b'url(',b'<filter'])
    qualification_path = Path(os.environ['F_X195_PATH_REFERENCE_QUALIFICATION'])
    assert hashlib.sha256(qualification_path.read_bytes()).hexdigest() == os.environ['F_X195_PATH_REFERENCE_QUALIFICATION_SHA']
    q = json.loads(qualification_path.read_text())
    assert q['actual_compiler_exit'] == 0 and q['actual_transitive_abi_qualified'] is True
    assert q['renderer_source_sha256'] == hashlib.sha256(_FX195_REFERENCE_RENDERER_SOURCE.encode()).hexdigest()
    assert q['rustc_verbose_version'].startswith('rustc 1.97.1 ') and 'host: aarch64-apple-darwin\n' in q['rustc_verbose_version']
    assert q['actual_transitive_versions'] == {'resvg':'0.48.1','usvg':'0.48.1','tiny-skia':'0.12.0'}
    compile_receipt = Path(q['actualCompileReceipt'])
    assert hashlib.sha256(compile_receipt.read_bytes()).hexdigest() == q['actualCompileReceiptSHA']
    compiled = json.loads(compile_receipt.read_text())
    assert compiled['actualExit'] == 0 and compiled['guardsBefore'] is True and compiled['guardsAfter'] is True
    painter = Path(q['renderer_binary'])
    assert hashlib.sha256(painter.read_bytes()).hexdigest() == q['renderer_binary_sha256'] == compiled['binarySHA']
    assert hashlib.sha256(Path(q['qualified_resvg_rlib']).read_bytes()).hexdigest() == q['qualified_resvg_rlib_sha256']
    with tempfile.TemporaryDirectory(prefix='fx195-original-path-') as temp:
        input_path,output_path = Path(temp)/'original.svg',Path(temp)/'original.png'
        input_path.write_bytes(raw)
        result = subprocess.run([str(painter),str(input_path),str(output_path)],capture_output=True,timeout=120,check=False)
        assert result.returncode == 0, result.stderr.decode(errors='replace')
        image = Image.open(output_path).convert('RGBA')
        image.load()
    record.update({'source_sha256':hashlib.sha256(source).hexdigest(),'svg_sha256':hashlib.sha256(raw).hexdigest(),
                   'origin':list(origin),'scale':font_scale,'roi':list(roi),'negative':negative,'cubic':cubic,
                   'qualification':str(qualification_path),'qualificationSHA':os.environ['F_X195_PATH_REFERENCE_QUALIFICATION_SHA'],
                   'rendererSHA':q['renderer_binary_sha256'],'stdout':result.stdout.decode(),'actualExit':result.returncode})
    return image,roi,len(case['text'])*advance,record


def _fx195_style_fixture_owners(before, case, page_size, notes=False, points=48):
    # Bounded fixture placement from original OOXML and original physical metrics.
    import os, math, struct, zipfile, io
    import xml.etree.ElementTree as ET
    from pathlib import Path
    source = case['expected'] if case['case']=='genuine-italic' else case['files'][0]
    if case['axis']: source = (Path(os.environ['F_X195_AXIS_ORACLES'])/(case['axis']+'-static.ttf')).read_bytes()
    tables = {source[o:o+4]:struct.unpack_from('>II',source,o+8) for o in range(12,12+16*struct.unpack_from('>H',source,4)[0],16)}
    upem = struct.unpack_from('>H',source,tables[b'head'][0]+18)[0]
    with zipfile.ZipFile(io.BytesIO(before)) as z: parts = {name:z.read(name) for name in z.namelist()}
    if 'word/document.xml' in parts:
        w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        tree=ET.fromstring(parts['word/document.xml']);section=tree.find('.//'+w+'sectPr');assert section is not None
        margin=section.find(w+'pgMar');size=section.find(w+'pgSz');assert margin is not None and size is not None
        assert tuple(round(int(size.get(w+key))/20) for key in ['w','h']) == tuple(page_size)
        assert ''.join(n.text or '' for n in tree.iter(w+'t')) == case['text']
        os2=tables[b'OS/2'][0];flags=struct.unpack_from('>H',source,os2+62)[0]
        ascent=struct.unpack_from('>h',source,os2+68)[0] if flags&128 else struct.unpack_from('>H',source,os2+74)[0]
        return [{'owner':'body','origin':(int(margin.get(w+'left'))/20,int(margin.get(w+'top'))/20+points*ascent/upem),'scale':1.}]
    p='{http://schemas.openxmlformats.org/presentationml/2006/main}';a='{http://schemas.openxmlformats.org/drawingml/2006/main}'
    slide_names=[name for name in parts if re.fullmatch(r'ppt/slides/slide\d+\.xml',name)];assert len(slide_names)==1
    slide=ET.fromstring(parts[slide_names[0]]);shapes=[shape for shape in slide.iter(p+'sp') if ''.join(n.text or '' for n in shape.iter(a+'t'))==case['text']];assert len(shapes)==1
    presentation=ET.fromstring(parts['ppt/presentation.xml']);size=presentation.find(p+'sldSz');assert size is not None
    slide_size=tuple(float(size.get(key))/12700 for key in ['cx','cy'])
    asc,desc=struct.unpack_from('>hh',source,tables[b'hhea'][0]+4);height=points*1.2;descent=-desc*points/upem
    baseline=max(height*min(asc/(asc-desc),.75),height-descent)
    def bounds(shape):
        xf=shape.find(p+'spPr/'+a+'xfrm');assert xf is not None and xf.get('rot','0')=='0' and xf.get('flipH','0') in ('0','false') and xf.get('flipV','0') in ('0','false')
        off,ext=xf.find(a+'off'),xf.find(a+'ext');assert off is not None and ext is not None
        return tuple(float(off.get(key))/12700 for key in ['x','y'])+tuple(float(ext.get(key))/12700 for key in ['cx','cy'])
    def origin(shape):
        x,y,_,_=bounds(shape);bp=shape.find(p+'txBody/'+a+'bodyPr');assert bp is not None
        assert bp.get('anchor','t')=='t' and bp.get('vert','horz')=='horz'
        return x+float(bp.get('lIns','91440'))/12700,y+float(bp.get('tIns','45720'))/12700+baseline
    body={'owner':'body','origin':origin(shapes[0]),'scale':1.}
    if not notes:
        assert tuple(round(v) for v in slide_size)==tuple(page_size)
        return [body]
    notes_names=[name for name in parts if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',name)];masters=[name for name in parts if re.fullmatch(r'ppt/notesMasters/notesMaster\d+\.xml',name)];assert len(notes_names)==len(masters)==1
    note=ET.fromstring(parts[notes_names[0]]);master=ET.fromstring(parts[masters[0]])
    def kind(shape):
        ph=shape.find(p+'nvSpPr/'+p+'nvPr/'+p+'ph');return ph.get('type') if ph is not None else None
    note_shapes=[shape for shape in note.iter(p+'sp') if ''.join(n.text or '' for n in shape.iter(a+'t'))==case['text']];assert len(note_shapes)==1 and kind(note_shapes[0])=='body'
    nshape=note_shapes[0]
    if nshape.find(p+'spPr/'+a+'xfrm') is None:
        owned=[shape for shape in master.iter(p+'sp') if kind(shape)=='body'];assert len(owned)==1
        # Position is inherited, bodyPr insets remain from the current source owner.
        inherited=ET.fromstring(ET.tostring(nshape));sppr=inherited.find(p+'spPr');assert sppr is not None
        sppr.append(ET.fromstring(ET.tostring(owned[0].find(p+'spPr/'+a+'xfrm'))));nshape=inherited
    notes_size=presentation.find(p+'notesSz');assert notes_size is not None and tuple(round(float(notes_size.get(key))/12700) for key in ['cx','cy'])==tuple(page_size)
    images=[shape for shape in note.iter(p+'sp') if kind(shape)=='sldImg' and shape.find(p+'spPr/'+a+'xfrm') is not None]
    if not images:images=[shape for shape in master.iter(p+'sp') if kind(shape)=='sldImg']
    assert len(images)==1
    x,y,width,height=bounds(images[0]);scale=min(width/slide_size[0],height/slide_size[1])
    preview_origin=(x+(width-slide_size[0]*scale)/2+body['origin'][0]*scale,y+(height-slide_size[1]*scale)/2+body['origin'][1]*scale)
    return [{'owner':'notes-body','origin':origin(nshape),'scale':1.},{'owner':'slide-preview','origin':preview_origin,'scale':scale}]


def _fx195_assert_installed_style(case, method, invoke, companion_pdf, companion_png, before, source_bytes, directory, expected_preview_scale=None):
    from PIL import Image
    import hashlib,json
    pdf = companion_pdf()
    pdf_image = _fx195_pdf_style_images(pdf,directory,'position-reference')[0]
    actual_owners = _fx195_pdf_style_anchors(pdf,pdf_image.size,case,expected_preview_scale)
    owners = _fx195_style_fixture_owners(before,case,pdf_image.size,expected_preview_scale is not None)
    assert len(actual_owners)==len(owners)
    for owner in owners:
        actual = [item for item in actual_owners if item['owner']==owner['owner']]
        assert len(actual)==1
        assert abs(actual[0]['scale']-owner['scale']) <= .00001
        assert all(abs(x-y)<=.00002 for x,y in zip(actual[0]['origin'],owner['origin'])), 'actual anchor must match independently declared fixture placement'
    controls=[]
    for owner in owners:
        pdf_expected,roi,advance = _fx195_independent_style_paint(case,pdf_image.size,owner['origin'],owner['scale'])
        assert _fx195_mask_ssim(pdf_image,pdf_expected,roi) >= .99, f"PDF {owner['owner']} must paint independently measured original font-text"
        expected,path_roi,path_advance,proof = _fx195_independent_path_paint(case,pdf_image.size,owner['origin'],owner['scale'])
        cubic,cubic_roi,cubic_advance,cubic_proof = _fx195_independent_path_paint(case,pdf_image.size,owner['origin'],owner['scale'],cubic=True)
        assert path_roi==cubic_roi==roi and path_advance==cubic_advance==advance
        controls.append((owner,pdf_expected,expected,cubic,roi,advance))
        (directory/(owner['owner']+'-independent-source-paths.json')).write_text(json.dumps({'quadratic':proof,'cubic':cubic_proof},indent=2)+'\n')
    (directory/'style-position-owners.json').write_text(json.dumps({
        'case':case['case'],'method':method,'pdfSHA':hashlib.sha256(pdf).hexdigest(),'physicalPageSize':list(pdf_image.size),
        'owners':owners,'actualOwners':actual_owners,'independentPointSize':48,'expectedPreviewScale':expected_preview_scale,
        'sourceFixtureSHA':hashlib.sha256(before).hexdigest(),'originalPhysicalFontSHAs':[hashlib.sha256(data).hexdigest() for data in case['files']],
        'threshold':.99,'adapter':'original OOXML fixture and physical metrics, independently checked reachable PDF owner matrices',
    },indent=2)+'\n')
    png_image = Image.open(io.BytesIO(companion_png())).convert('RGBA')
    for owner,pdf_expected,expected,cubic,roi,advance in controls:
        assert png_image.size==expected.size==cubic.size
        assert _fx195_mask_ssim(png_image,expected,roi)>=.99, f"PNG quadratic {owner['owner']}"
        assert _fx195_mask_ssim(png_image,cubic,roi)>=.99, f"PNG cubic {owner['owner']}"
    owner,pdf_expected,expected,cubic,roi,advance = next(control for control in controls if control[0]['owner']!='slide-preview')
    origin=owner['origin'];value=invoke();is_pdf=method in ('to_pdf','to_notes_pdf')
    if is_pdf: images=_fx195_pdf_style_images(value,directory,'actual-method')
    elif method in ('render_page_to_png','render_slide_to_png'):images=[Image.open(io.BytesIO(value)).convert('RGBA')]
    elif method in ('render_all_pages','render_pages','render_all_slides','render_all_notes'):
        assert len(value)==1;images=[Image.open(io.BytesIO(value[0])).convert('RGBA')]
    elif method=='render_page_to_svg':
        _fx195_svg_style_artifact(value,case,expected,cubic,roi,origin,advance,directory);images=[]
    elif method=='text_layout':
        assert len(value)==1 and len(value[0].lines)==1
        assert abs(value[0].lines[0].bounds.width-advance)<=.00001;images=[]
    elif method=='layout':
        assert len(value)==1 and value[0].body_index==0 and value[0].physical_page==1
        assert value[0].bounds.x>=0 and value[0].bounds.height>=33.6;images=[]
    else:
        assert method=='layout_page' and value is not None
        assert value.page_number==1 and value.width==expected.width and value.height==expected.height;images=[]
    for image in images:
        for owner,pdf_expected,expected,cubic,roi,advance in controls:
            assert image.size==expected.size
            if is_pdf: assert _fx195_mask_ssim(image,pdf_expected,roi)>=.99, f"selected PDF {owner['owner']}"
            else:
                assert _fx195_mask_ssim(image,expected,roi)>=.99, f"selected PNG quadratic {owner['owner']}"
                assert _fx195_mask_ssim(image,cubic,roi)>=.99, f"selected PNG cubic {owner['owner']}"
    assert _fx195_style_observation(invoke(),method,directory,'repeat')==_fx195_style_observation(value,method,directory,'first')
    assert source_bytes()==before
    for owner,pdf_expected,expected,cubic,roi,advance in controls:
        wrong = dict(case,expected=case['ordinary'][0]) if case['case']=='genuine-italic' else dict(case,axis=None,bold=False,italic=False,synthesis=True)
        wrong_pdf,_,_ = _fx195_independent_style_paint(wrong,expected.size,owner['origin'],owner['scale'])
        wrong_path,_,_,_ = _fx195_independent_path_paint(wrong,expected.size,owner['origin'],owner['scale'])
        assert _fx195_mask_ssim(wrong_pdf,pdf_expected,roi)<.99, f"independent PDF wrong geometry {owner['owner']}"
        assert _fx195_mask_ssim(wrong_path,expected,roi)<.99, f"independent path wrong geometry {owner['owner']}"
        for negative in ('style','origin')+(('shear',) if case['synthesis'] and case['italic'] else ()):
            wrong_pdf,_,_ = _fx195_independent_style_paint(case,expected.size,owner['origin'],owner['scale'],negative=negative)
            wrong_path,_,_,_ = _fx195_independent_path_paint(case,expected.size,owner['origin'],owner['scale'],negative=negative)
            assert _fx195_mask_ssim(wrong_pdf,pdf_expected,roi)<.99, f"independent PDF {negative} discriminator {owner['owner']}"
            assert _fx195_mask_ssim(wrong_path,expected,roi)<.99, f"independent path {negative} discriminator {owner['owner']}"



@pytest.mark.parametrize('style_case', _FX195_STYLE_CASES)
@pytest.mark.parametrize('method', ['to_pdf','render_slide_to_png','render_all_slides','text_layout','to_notes_pdf','render_all_notes'])
def test_fx195_installed_ppt_styles_match_independent_paint(method, style_case, tmp_path):
    import xml.etree.ElementTree as ET
    import rpptx
    case = _fx195_style_case(style_case)
    presentation = _fx195_presentation(case['family'])
    with zipfile.ZipFile(io.BytesIO(presentation.to_bytes())) as package:
        parts = {name:package.read(name) for name in package.namelist()}
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    counts = {'slide':0,'notes':0}
    for name in list(parts):
        kind = 'slide' if re.fullmatch(r'ppt/slides/slide\d+\.xml',name) else 'notes' if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',name) else None
        if kind is None: continue
        tree = ET.fromstring(parts[name])
        for run in tree.iter(ns+'r'):
            prop = run.find(ns+'rPr')
            latin = prop.find(ns+'latin') if prop is not None else None
            if latin is None or latin.get('typeface') != case['family']: continue
            prop.set('sz','4800')
            prop.set('b','1' if case['bold'] else '0')
            prop.set('i','1' if case['italic'] else '0')
            run.find(ns+'t').text = case['text']
            counts[kind]+=1
        parts[name]=ET.tostring(tree,encoding='utf-8',xml_declaration=True)
    assert counts == {'slide':1,'notes':1}
    output = io.BytesIO()
    with zipfile.ZipFile(output,'w') as z:
        for name,data in parts.items(): z.writestr(name,data)
    presentation = rpptx.Presentation.from_bytes(output.getvalue())
    before = presentation.to_bytes()
    kwargs = {'fonts':[(case['family'],data) for data in case['files']]}
    notes = method in ('to_notes_pdf','render_all_notes')
    preview_scale = _fx195_notes_preview_scale(parts) if notes else None
    def invoke():
        if method == 'render_slide_to_png': return presentation.render_slide_to_png(0,dpi=72,**kwargs)
        if method == 'render_all_slides': return presentation.render_all_slides(dpi=72,**kwargs)
        if method == 'render_all_notes': return presentation.render_all_notes(dpi=72,**kwargs)
        return getattr(presentation,method)(**kwargs)
    pdf = lambda:presentation.to_notes_pdf(**kwargs) if notes else presentation.to_pdf(**kwargs)
    png = lambda:presentation.render_all_notes(dpi=72,**kwargs)[0] if notes else presentation.render_slide_to_png(0,dpi=72,**kwargs)
    _fx195_assert_installed_style(case,method,invoke,pdf,png,before,presentation.to_bytes,tmp_path,preview_scale)



@pytest.mark.parametrize("reason", ["index", "placeholder", "concrete", "gradient_placeholder", "gradient_concrete", "line_gradient_placeholder", "line_gradient_concrete"])
@pytest.mark.parametrize("placed", [False, True])
def test_fx206_installed_member_recovery_paints_siblings_and_preserves_source(reason, placed, tmp_path):
    import io
    import zipfile
    import rpptx
    from PIL import Image
    seed=rpptx.Presentation()
    seed.slides.add_slide(seed.slide_layouts[6])
    def member(ident, width, properties, style="", positioned=True):
        transform=(f'<a:xfrm><a:off x="127000" y="127000"/><a:ext cx="{width}" cy="254000"/></a:xfrm>' if positioned else '')
        return f'<p:sp><p:nvSpPr><p:cNvPr id="{ident}" name="owner{ident}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>{transform}<a:prstGeom prst="rect"/>{properties}<f:opaque xmlns:f="urn:f206:member" literal="&#x41;"/></p:spPr>{style}</p:sp>'
    properties,style={
        'index':('', '<p:style><a:lnRef idx="0"/><a:fillRef idx="999"/><a:effectRef idx="0"/><a:fontRef idx="minor"/></p:style>'),
        'placeholder':('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>',''),
        'concrete':('<a:solidFill><a:prstClr val="futureColour"/></a:solidFill>',''),
        'gradient_placeholder':('<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>',''),
        'gradient_concrete':('<a:gradFill><a:gsLst><a:gs pos="0"><a:prstClr val="futureColour"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>',''),
        'line_gradient_placeholder':('<a:ln><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></a:ln>',''),
        'line_gradient_concrete':('<a:ln><a:gradFill><a:gsLst><a:gs pos="0"><a:prstClr val="futureColour"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></a:ln>',''),
    }[reason]
    children=member(1,254000,'<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')+member(2,254000,properties,style,placed)+member(3,50800,'<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>')
    xml=f'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" showMasterSp="0"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="10" name="root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:grpSp><p:nvGrpSpPr><p:cNvPr id="11" name="parent"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="63500" y="63500"/><a:ext cx="2540000" cy="2540000"/><a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="1270000"/></a:xfrm></p:grpSpPr>{children}</p:grpSp></p:spTree></p:cSld></p:sld>'
    buf=io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(seed.to_bytes())) as original, zipfile.ZipFile(buf,'w') as edited:
        for info in original.infolist():
            edited.writestr(info,xml.encode() if info.filename=='ppt/slides/slide1.xml' else original.read(info.filename))
    deck=rpptx.Presentation.from_bytes(buf.getvalue());before=deck.to_bytes()
    png=deck.render_slide_to_png(0,dpi=72)
    image=Image.open(io.BytesIO(png)).convert('RGB')
    assert image.getpixel((28,40))==(0,255,0)
    assert image.getpixel((40,40))==(255,0,0)
    if placed:
        red,green,blue=image.getpixel((64,40))
        assert red<200 and green<80 and blue<80, 'failed member outline at independent scaled edge'
    assert deck.to_pdf().startswith(b'%PDF-')
    assert deck.to_bytes()==before
    saved=tmp_path/'recovery.pptx';deck.save(saved)
    reopened=rpptx.Presentation(saved)
    assert reopened.render_slide_to_png(0,dpi=72)==png
    with zipfile.ZipFile(io.BytesIO(reopened.to_bytes())) as a, zipfile.ZipFile(io.BytesIO(before)) as b:
        assert a.read('ppt/slides/slide1.xml')==b.read('ppt/slides/slide1.xml')
    with pytest.raises(rpptx.RpptxError):
        rpptx.Presentation.from_bytes(b'not a package')


def test_fx206_internal_run_target_surface_base_positive_owner():
    import rpptx
    prs = _textbox_presentation(rpptx)
    run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    assert run.text and run.hyperlink.address is None
    before = prs.to_bytes()
    assert before.startswith(b"PK")
    assert hasattr(type(run.hyperlink), "target_slide"), "existing valid owned run must expose checked slide-target capability"
    assert prs.to_bytes() == before


def test_fx206_internal_run_slide_targets_prune_and_preserve_live_owners(tmp_path):
    import xml.etree.ElementTree as ET
    import rpptx
    prs = _textbox_presentation(rpptx)
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides.add_slide(prs.slide_layouts[6])
    frame = prs.slides[0].shapes[0].text_frame
    frame.paragraphs[0].add_run(" shared")
    frame = prs.slides[0].shapes[0].text_frame
    paragraph = frame.paragraphs[0]
    first, second = paragraph.runs
    link = first.hyperlink
    font = first.font
    shape = prs.slides[0].shapes[0]
    shape_id = shape.shape_id
    before = prs.to_bytes()
    assert link.target_slide is None
    assert prs.to_bytes() == before
    link.target_slide = None
    assert prs.to_bytes() == before
    shape.click_action.target_slide = prs.slides[1]
    first.hyperlink.target_slide = prs.slides[1]
    second.hyperlink.target_slide = prs.slides[1]
    assert link.target_slide == prs.slides[1]
    assert link.address == "slide2.xml"
    assert paragraph.text == frame.text and first.text and font.bold is None
    shared = prs.to_bytes()
    link.target_slide = prs.slides[1]
    assert prs.to_bytes() == shared
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    parts = _package_parts(shared)
    slide_xml = ET.fromstring(parts["ppt/slides/slide1.xml"])
    clicks = slide_xml.findall(".//a:hlinkClick", ns)
    assert len(clicks) == 3
    ids = {click.get("{" + ns["r"] + "}id") for click in clicks}
    assert len(ids) == 1 and None not in ids
    assert all(click.get("action") == "ppaction://hlinksldjump" for click in clicks)
    shared_id = next(iter(ids))
    link.target_slide = prs.slides[2]
    for invalid in (object(), 1, "slide"):
        before = prs.to_bytes()
        with pytest.raises(TypeError):
            link.target_slide = invalid
        assert prs.to_bytes() == before and link.target_slide == prs.slides[2]
    other = rpptx.Presentation()
    foreign = other.slides.add_slide(other.slide_layouts[6])
    before = prs.to_bytes()
    with pytest.raises(ValueError, match="not in this presentation"):
        link.target_slide = foreign
    assert prs.to_bytes() == before
    second.hyperlink.target_slide = None
    shape.click_action.target_slide = None
    rels = ET.fromstring(_package_parts(prs.to_bytes())["ppt/slides/_rels/slide1.xml.rels"])
    assert all(rel.get("Id") != shared_id for rel in rels)
    link.address = "https://example.test/run"
    assert link.target_slide is None and link.address == "https://example.test/run"
    link.target_slide = prs.slides[0]
    assert link.target_slide == prs.slides[0]
    assert first.text and shape.shape_id == shape_id and font.bold is None
    saved = prs.to_bytes()
    reopened = rpptx.Presentation.from_bytes(saved)
    fresh = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].hyperlink
    assert fresh.target_slide == reopened.slides[0] and fresh.address == "slide1.xml"
    path = tmp_path / "internal-run-links.pptx"
    reopened.save(path)
    normal = rpptx.Presentation(path)
    assert normal.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.target_slide == normal.slides[0]
    link.target_slide = prs.slides[2]
    prs.slides.remove(prs.slides[2])
    fresh = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].hyperlink
    assert fresh.target_slide is None and fresh.address is None
    stale = fresh
    prs.slides[0].shapes[0].text_frame.text = "replacement"
    before = prs.to_bytes()
    with pytest.raises(rpptx.StaleElementError, match="TextFrame.text"):
        _ = stale.target_slide
    with pytest.raises(rpptx.StaleElementError, match="TextFrame.text"):
        stale.target_slide = None
    assert prs.to_bytes() == before


def test_fx206_cell_internal_run_link_refuses_before_relationship_mutation():
    import rpptx
    prs, table = _f_x194_table()
    cell = table.cell(0, 0)
    cell.text = "cell link"
    hyperlink = cell.text_frame.paragraphs[0].runs[0].hyperlink
    before = prs.to_bytes()
    for value in (prs.slides[0], None):
        with pytest.raises(ValueError, match="cell hyperlink assignment"):
            hyperlink.target_slide = value
        assert prs.to_bytes() == before
        assert hyperlink.address is None



def test_fx206_installed_imported_run_alias_and_foreign_a_round_trip(tmp_path):
    import re
    import xml.etree.ElementTree as ET
    import rpptx
    prs = _textbox_presentation(rpptx)
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides.add_slide(prs.slide_layouts[6])
    owner = prs.slides[0].shapes[0]
    owner_id = owner.shape_id
    link = owner.text_frame.paragraphs[0].runs[0].hyperlink
    link.target_slide = prs.slides[1]
    parts = _package_parts(prs.to_bytes())
    drawing = "http://schemas.openxmlformats.org/drawingml/2006/main"
    relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    root = ET.fromstring(parts["ppt/slides/slide1.xml"])
    clicks = list(root.iter("{" + drawing + "}hlinkClick"))
    assert len(clicks) == 1
    relation_id = clicks[0].get("{" + relationships + "}id")
    assert relation_id
    xml = parts["ppt/slides/slide1.xml"].decode()
    matches = list(re.finditer(r"<a:rPr(?:\s[^>]*)?>.*?</a:rPr>", xml))
    assert len(matches) == 1, "one explicit owned run-property source fixture"
    imported = (f'<q:rPr xmlns:q="{drawing}" xmlns:rel="{relationships}" '
                'xmlns:a=\'urn:foreign-a\' xmlns:r=\'urn:foreign-r\' xmlns:f="urn:payload" f:keep=\'&#65;\'>'
                '<a:hlinkClick r:id=\'foreign\'><a:payload f:keep=\'&#66;\'/></a:hlinkClick>'
                f'<q:hlinkClick rel:id="{relation_id}" r:id=\'opaque\' action="ppaction://hlinksldjump" tooltip=\'&#67;\'>'
                '<a:payload f:keep=\'&#68;\'/></q:hlinkClick></q:rPr>')
    parts["ppt/slides/slide1.xml"] = (xml[:matches[0].start()] + imported + xml[matches[0].end():]).encode()
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as package:
        for name, data in parts.items():
            package.writestr(name, data)
    deck = rpptx.Presentation.from_bytes(stream.getvalue())
    shape = deck.slides[0].shapes[0]
    assert shape.shape_id == owner_id
    run = shape.text_frame.paragraphs[0].runs[0]
    hyperlink = run.hyperlink
    assert run.text == "hello" and hyperlink.target_slide == deck.slides[1]
    before = deck.to_bytes()
    hyperlink.target_slide = deck.slides[1]
    assert deck.to_bytes() == before
    hyperlink.target_slide = deck.slides[2]
    assert hyperlink.target_slide == deck.slides[2] and run.text == "hello"
    saved = deck.to_bytes()
    xml = _package_parts(saved)["ppt/slides/slide1.xml"].decode()
    for token in ["<q:rPr", "xmlns:a='urn:foreign-a'", "xmlns:r='urn:foreign-r'", "f:keep='&#65;'",
                  "<a:hlinkClick r:id='foreign'><a:payload f:keep='&#66;'/></a:hlinkClick>",
                  "r:id='opaque'", "tooltip='&#67;'", "<a:payload f:keep='&#68;'/>"]:
        assert token in xml
    reopened = rpptx.Presentation.from_bytes(saved)
    assert reopened.slides[0].shapes[0].shape_id == owner_id
    assert reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.target_slide == reopened.slides[2]
    assert reopened.to_bytes() == saved
    path = tmp_path / "imported-alias-links.pptx"
    reopened.save(path)
    normal = rpptx.Presentation(path)
    assert normal.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.target_slide == normal.slides[2]
    hyperlink.address = "https://example.test/aliased"
    assert hyperlink.address == "https://example.test/aliased" and hyperlink.target_slide is None
    hyperlink.target_slide = None
    cleared = deck.to_bytes()
    xml = _package_parts(cleared)["ppt/slides/slide1.xml"].decode()
    assert "<a:hlinkClick r:id='foreign'><a:payload f:keep='&#66;'/></a:hlinkClick>" in xml
    assert hyperlink.address is None and hyperlink.target_slide is None
    assert rpptx.Presentation.from_bytes(cleared).to_bytes() == cleared

def test_fx206_cell_anchor_modes_round_trip_keep_physical_text_handles(tmp_path):
    import xml.etree.ElementTree as ET
    import rpptx
    from rpptx.enum.text import MSO_ANCHOR
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table_shape = slide.shapes.add_table(1, 1, 0, 0, 1000000, 1000000)
    cell = table_shape.table.cell(0, 0)
    cell.text = "anchor text"
    cell = prs.slides[0].shapes[0].table.cell(0, 0)
    frame = cell.text_frame
    frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    font = run.font
    initial_bold = font.bold
    before = prs.to_bytes()
    assert cell.vertical_anchor is None
    assert prs.to_bytes() == before
    for value, token in [(MSO_ANCHOR.TOP, "t"), (MSO_ANCHOR.MIDDLE, "ctr"),
                         (MSO_ANCHOR.BOTTOM, "b"), (None, None)]:
        cell.vertical_anchor = value
        assert cell.vertical_anchor == value
        assert frame.vertical_anchor == MSO_ANCHOR.BOTTOM
        assert run.text == "anchor text" and font.bold == initial_bold
        assert paragraph.text == "anchor text"
        saved = prs.to_bytes()
        cell.vertical_anchor = value
        assert prs.to_bytes() == saved
        with zipfile.ZipFile(io.BytesIO(saved)) as archive:
            xml = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        assert xml.find(".//a:tcPr", ns).get("anchor") == token
        assert xml.find(".//a:bodyPr", ns).get("anchor") == "b"
        reopened = rpptx.Presentation.from_bytes(saved)
        fresh = reopened.slides[0].shapes[0].table.cell(0, 0)
        assert fresh.vertical_anchor == value
        assert fresh.text_frame.vertical_anchor == MSO_ANCHOR.BOTTOM
        assert fresh.text == "anchor text"
        path = tmp_path / "cell-anchor.pptx"
        reopened.save(path)
        assert rpptx.Presentation(path).slides[0].shapes[0].table.cell(0, 0).vertical_anchor == value
    for invalid, kind in [(999, ValueError), ("bottom", TypeError), (object(), TypeError)]:
        before = prs.to_bytes()
        with pytest.raises(kind):
            cell.vertical_anchor = invalid
        assert prs.to_bytes() == before
        assert run.text == "anchor text" and font.bold == initial_bold
    prs.slides[0].shapes.remove(prs.slides[0].shapes[0])
    before = prs.to_bytes()
    with pytest.raises(rpptx.StaleElementError, match="ShapeCollection.remove"):
        _ = cell.vertical_anchor
    with pytest.raises(rpptx.StaleElementError, match="ShapeCollection.remove"):
        cell.vertical_anchor = MSO_ANCHOR.TOP
    assert prs.to_bytes() == before


def test_fx206_cell_anchor_clear_retains_foreign_tcpr_payload():
    import rpptx
    from rpptx.enum.text import MSO_ANCHOR
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_table(1, 1, 0, 0, 1000000, 1000000)
    source = prs.to_bytes()
    edited = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source)) as archive, zipfile.ZipFile(edited, "w") as output:
        for info in archive.infolist():
            payload = archive.read(info.filename)
            if info.filename == "ppt/slides/slide1.xml":
                assert payload.count(b"<a:tcPr") == 1
                match = re.search(rb"<a:tcPr([^>]*)/>", payload)
                assert match is not None
                replacement = b"<a:tcPr" + match.group(1) + b" xmlns:x='urn:foreign' x:keep='&#65;&#x42;' x:anchor='foreign' anchor='ctr'><x:opaque xmlns:a='urn:foreign'><a:node/></x:opaque></a:tcPr>"
                payload = payload[:match.start()] + replacement + payload[match.end():]
            output.writestr(info, payload)
    reopened = rpptx.Presentation.from_bytes(edited.getvalue())
    cell = reopened.slides[0].shapes[0].table.cell(0, 0)
    assert cell.vertical_anchor == MSO_ANCHOR.MIDDLE
    for value in [MSO_ANCHOR.BOTTOM, None, MSO_ANCHOR.TOP]:
        cell.vertical_anchor = value
        saved = reopened.to_bytes()
        with zipfile.ZipFile(io.BytesIO(saved)) as archive:
            xml = archive.read("ppt/slides/slide1.xml")
        assert b"x:keep='&#65;&#x42;'" in xml
        assert b"x:anchor='foreign'" in xml
        assert b"<x:opaque xmlns:a='urn:foreign'><a:node/></x:opaque>" in xml
        assert rpptx.Presentation.from_bytes(saved).slides[0].shapes[0].table.cell(0, 0).vertical_anchor == value
        cell.vertical_anchor = value
        assert reopened.to_bytes() == saved



def test_fx206_every_created_layout_placeholder_name_matches_pinned_reference(tmp_path):
    from pathlib import Path
    import importlib.metadata
    import pptx
    import rpptx
    assert importlib.metadata.version("python-pptx") == "1.0.2"
    source = Path(__file__).resolve().parents[2] / "rpptx" / "assets" / "default.pptx"
    payload = source.read_bytes()
    expected = pptx.Presentation(io.BytesIO(payload))
    actual = rpptx.Presentation.from_bytes(payload)
    assert len(actual.slide_layouts) == len(expected.slide_layouts) == 11
    held_layouts = actual.slide_layouts
    for index in range(11):
        expected_slide = expected.slides.add_slide(expected.slide_layouts[index])
        actual_slide = actual.slides.add_slide(held_layouts[index])
        names = [shape.name for shape in actual_slide.shapes]
        assert names == [shape.name for shape in expected_slide.shapes], expected.slide_layouts[index].name
        assert len(set(names)) == len(names)
        assert len(set(shape.shape_id for shape in actual_slide.shapes)) == len(names)
    destination = tmp_path / "all-layout-names.pptx"
    actual.save(destination)
    reopened = rpptx.Presentation(destination)
    assert [[shape.name for shape in slide.shapes] for slide in reopened.slides] == [[shape.name for shape in slide.shapes] for slide in actual.slides]


@pytest.mark.parametrize("size,crop_axis", [((120, 60), "horizontal"), ((60, 120), "vertical")])
def test_fx206_picture_placeholder_fill_owner_crop_and_atomicity(size, crop_axis):
    import io
    import zipfile
    from PIL import Image
    from lxml import etree
    import rpptx

    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[8])
    with zipfile.ZipFile(io.BytesIO(presentation.to_bytes())) as archive:
        root = etree.fromstring(archive.read("ppt/slides/slide1.xml"))
        namespaces = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
        picture_ids = root.xpath("//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='pic']/p:nvSpPr/p:cNvPr/@id", namespaces=namespaces)
    assert len(picture_ids) == 1
    placeholder = next(shape for shape in slide.shapes if shape.shape_id == int(picture_ids[0]))
    assert placeholder.effective_geometry() is not None
    assert callable(getattr(placeholder, "insert_picture", None)), "picture placeholder insertion is absent"
    before = presentation.to_bytes()
    with pytest.raises(rpptx.RpptxError):
        placeholder.insert_picture(b"not-an-image")
    assert presentation.to_bytes() == before
    original_id, original_name = placeholder.shape_id, placeholder.name
    original_index = next(i for i, shape in enumerate(slide.shapes) if shape.shape_id == original_id)
    geometry = placeholder.effective_geometry()
    assert geometry is not None
    sibling = slide.shapes[0]
    stream = io.BytesIO()
    colours = [(190, 20, 30), (20, 190, 30), (20, 30, 190), (190, 130, 20), (130, 20, 190), (20, 130, 190), (90, 20, 30), (20, 90, 30), (20, 30, 90)]
    image = Image.new("RGB", size)
    for row in range(3):
        for column in range(3):
            image.paste(colours[row * 3 + column], (column * size[0] // 3, row * size[1] // 3, (column + 1) * size[0] // 3, (row + 1) * size[1] // 3))
    image.save(stream, format="PNG")
    picture = placeholder.insert_picture(stream.getvalue())
    assert picture.shape_id == original_id and picture.name == original_name
    assert picture.effective_geometry() == geometry
    assert slide.shapes[original_index].shape_id == original_id
    assert sibling.shape_id != original_id
    with pytest.raises(rpptx.StaleElementError, match="Shape.insert_picture"):
        _ = placeholder.shape_id
    if crop_axis == "horizontal":
        assert picture.crop_left == picture.crop_right and picture.crop_left > 0
        assert picture.crop_top == picture.crop_bottom == 0
    else:
        assert picture.crop_top == picture.crop_bottom and picture.crop_top > 0
        assert picture.crop_left == picture.crop_right == 0
    saved = presentation.to_bytes()
    assert saved != before
    with zipfile.ZipFile(io.BytesIO(saved)) as archive:
        root = etree.fromstring(archive.read("ppt/slides/slide1.xml"))
        ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "a": "http://schemas.openxmlformats.org/drawingml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
        owners = root.xpath("//p:pic[p:nvPicPr/p:cNvPr/@id=$id]", namespaces=ns, id=str(original_id))
        assert len(owners) == 1
        assert owners[0].xpath("p:nvPicPr/p:nvPr/p:ph/@type", namespaces=ns) == ["pic"]
        assert len(owners[0].xpath("p:blipFill/a:blip/@r:embed", namespaces=ns)) == 1
        assert owners[0].xpath("p:spPr/a:prstGeom/@prst", namespaces=ns) == ["rect"]
        assert owners[0].xpath("p:spPr/a:custGeom", namespaces=ns) == []
        assert [etree.QName(child).localname for child in owners[0].xpath("p:spPr/*[self::a:xfrm or self::a:prstGeom]", namespaces=ns)] == ["xfrm", "prstGeom"]
    reopened = rpptx.Presentation.from_bytes(saved)
    again = reopened.slides[0].shapes[original_index]
    assert again.shape_id == original_id and again.name == original_name
    actual_png = presentation.render_slide_to_png(0, dpi=72.0)
    assert actual_png is not None
    actual = Image.open(io.BytesIO(actual_png)).convert("RGB")
    left, top, width, height = (float(int(value)) / 12700.0 for value in geometry)
    # Independent centered cover mapping, using only original pixel dimensions
    # and the inherited native rectangle. No product crop values feed this oracle.
    source_aspect = size[0] / size[1]
    box_aspect = width / height
    crop_x = (1.0 - box_aspect / source_aspect) / 2.0 if source_aspect > box_aspect else 0.0
    crop_y = (1.0 - source_aspect / box_aspect) / 2.0 if source_aspect < box_aspect else 0.0
    differs_from_uncropped = False
    for fx, fy in [(0.30, 0.25), (0.70, 0.25), (0.30, 0.75), (0.70, 0.75)]:
        sx = (crop_x + fx * (1 - 2 * crop_x)) * size[0]
        sy = (crop_y + fy * (1 - 2 * crop_y)) * size[1]
        colour = image.getpixel((int(sx), int(sy)))
        uncropped = image.getpixel((int(fx * size[0]), int(fy * size[1])))
        differs_from_uncropped |= colour != uncropped
        pixel = actual.getpixel((round(left + width * fx), round(top + height * fy)))
        assert max(abs(a - b) for a, b in zip(pixel, colour)) <= 1
    assert differs_from_uncropped, "the painter control must distinguish cover crop from stretch"
    assert reopened.render_slide_to_png(0, dpi=72.0) == actual_png
    stable = presentation.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="not an unfilled picture placeholder"):
        picture.insert_picture(b"not-an-image")
    assert presentation.to_bytes() == stable


@pytest.mark.parametrize("extent_y", [2 * 914400, 3 * 914400 // 2], ids=["uniform-2x", "nonuniform-2x-1p5x"])
def test_fx206_issue323_added_group_connector_paints_combined_affine_and_keeps_owners(extent_y, tmp_path):
    import io
    import math
    import re
    from PIL import Image
    from lxml import etree
    import rpptx
    from rpptx.dml.color import RGBColor
    from rpptx.enum.shapes import MSO_CONNECTOR

    unit = 914400
    deck = rpptx.Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    rectangle = slide.shapes.add_shape("rect", 0, 0, 2 * unit, unit)
    rectangle.fill.solid()
    rectangle.fill.fore_color.rgb = RGBColor(180, 30, 40)
    rectangle.line.width = 0
    rectangle_id = rectangle.shape_id
    parts = _package_parts(deck.to_bytes())
    text = parts["ppt/slides/slide1.xml"].decode()
    body = _fx206_actual_slide_shapes(text, 1)[0]
    opaque = '<a:extLst><a:ext uri="urn:fx206:preserved"><f:payload xmlns:f="urn:fx206:opaque" keep="&#65;"> untouched </f:payload></a:ext></a:extLst>'
    assert body.count("</p:spPr>") == 1
    body = body.replace("</p:spPr>", opaque + "</p:spPr>")
    transform = (
        '<a:xfrm rot="1800000" flipH="1"><a:off x="2286000" y="2286000"/>'
        f'<a:ext cx="3657600" cy="{extent_y}"/><a:chOff x="0" y="0"/>'
        '<a:chExt cx="1828800" cy="914400"/></a:xfrm>'
    )
    group_xml = _fx206_group_source(body)
    group_xml, changed = re.subn(r"<a:xfrm>.*?</a:xfrm>", transform, group_xml, count=1, flags=re.DOTALL)
    assert changed == 1
    patched = _fx206_replace_slide_source(deck.to_bytes(), lambda xml: xml.replace(_fx206_actual_slide_shapes(xml, 1)[0], group_xml, 1))
    actual = rpptx.Presentation.from_bytes(patched)
    group = actual.slides[0].shapes[0]
    held = group.shapes[0]
    assert held.shape_id == rectangle_id and (held.width, held.height) == (2 * unit, unit)
    assert group.rotation == 30 and group.flip_h is True
    # Exercise the reported public operation, rather than importing a pre-existing connector.
    begin, end = (3 * unit // 10, unit // 5), (8 * unit // 5, 7 * unit // 10)
    connector = group.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, *begin, *end)
    connector.line.color.rgb = RGBColor(0, 0, 0)
    connector.line.width = rpptx.Pt(4)
    connector.theme_effect_index = 0
    assert group.shape_id == 50 and held.shape_id == rectangle_id
    assert [shape.shape_id for shape in group.shapes] == [rectangle_id, connector.shape_id]
    assert (connector.begin_x, connector.begin_y, connector.end_x, connector.end_y) == (*begin, *end)
    source = actual.to_bytes()
    xml = _package_parts(source)["ppt/slides/slide1.xml"]
    assert opaque.encode() in xml
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    root = etree.fromstring(xml)
    xfrm = root.xpath("//p:grpSp[p:nvGrpSpPr/p:cNvPr/@id='50']/p:grpSpPr/a:xfrm", namespaces=ns)
    assert len(xfrm) == 1
    assert xfrm[0].get("rot") == "1800000" and xfrm[0].get("flipH") == "1"
    assert xfrm[0].find("a:off", ns).attrib == {"x": "2286000", "y": "2286000"}
    assert xfrm[0].find("a:ext", ns).attrib == {"cx": "3657600", "cy": str(extent_y)}
    assert xfrm[0].find("a:chExt", ns).attrib == {"cx": "1828800", "cy": "914400"}
    # Source-controlled mathematical oracle: scale child coordinates, reflect about
    # the group frame center, then rotate clockwise, as the approved contract says.
    # No product-resolved coordinates or emitted crop/transform values feed this map.
    cosine, sine = math.cos(math.pi / 6), math.sin(math.pi / 6)
    def mapped(point, wrong_order=False):
        dx = point[0] * 2 - 2 * unit
        dy = point[1] * extent_y / unit - extent_y / 2
        if wrong_order:
            x, y = -(cosine * dx - sine * dy), sine * dx + cosine * dy
        else:
            x, y = cosine * -dx - sine * dy, sine * -dx + cosine * dy
        return ((2286000 + 2 * unit + x) / 12700, (2286000 + extent_y / 2 + y) / 12700)
    painted = actual.render_slide_to_png(0, dpi=72.0)
    image = Image.open(io.BytesIO(painted)).convert("RGB")
    def samples(point):
        x, y = map(round, point)
        assert 2 <= x < image.width - 2 and 2 <= y < image.height - 2
        return [image.getpixel((x + dx, y + dy)) for dx in (-1, 0, 1) for dy in (-1, 0, 1)]
    for fraction in (0.2, 0.4, 0.6, 0.8):
        point = tuple(a + fraction * (b - a) for a, b in zip(begin, end))
        assert any(max(pixel) <= 20 for pixel in samples(mapped(point))), "Issue323 grouped connector must paint at flip-before-rotate position"
    for point in ((unit // 5, 4 * unit // 5), (9 * unit // 5, unit // 5)):
        assert any(max(abs(a - b) for a, b in zip(pixel, (180, 30, 40))) <= 1 for pixel in samples(mapped(point))), "Issue323 resolvable rectangle sibling must remain painted"
    assert any(not any(max(pixel) <= 20 for pixel in samples(mapped(tuple(a + fraction * (b - a) for a, b in zip(begin, end)), True))) for fraction in (0.2, 0.4, 0.6, 0.8)), "negative control must distinguish reversed transform order"
    assert actual.to_bytes() == source and held.shape_id == rectangle_id
    target = tmp_path / ("grouped-connector-" + str(extent_y) + ".pptx")
    actual.save(target)
    reopened = rpptx.Presentation(target)
    assert reopened.render_slide_to_png(0, dpi=72.0) == painted
    assert opaque.encode() in _package_parts(reopened.to_bytes())["ppt/slides/slide1.xml"]
    assert [shape.shape_id for shape in reopened.slides[0].shapes[0].shapes] == [rectangle_id, connector.shape_id]

def test_fx210_ppt_companions_and_frame_diagnostics_share_actual_output():
    from pathlib import Path
    from rpptx import Presentation

    presentation = Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    for index, family in enumerate(("Aptos", "Calibri")):
        shape = presentation.slides[0].shapes.add_textbox(0, index * 914400, 4000000, 914400)
        shape.text = "Actual contributing text"
        shape = presentation.slides[0].shapes[index]
        shape.text_frame.paragraphs[0].runs[0].font.name = family
    source = (Path(__file__).resolve().parents[2] / "oxml-layout/fonts/Carlito-Regular.ttf").read_bytes()
    kwargs = {"fonts": [("Aptos", source)]}
    pdf, facts = presentation.to_pdf_with_diagnostics(**kwargs)
    assert pdf == presentation.to_pdf(**kwargs)
    assert isinstance(facts, tuple) and all(isinstance(value, str) for value in facts)
    assert sum("Font 'Aptos' is approximate" in value for value in facts) == 1
    png, png_facts = presentation.render_slide_to_png_with_diagnostics(0, 72.0, **kwargs)
    assert png == presentation.render_slide_to_png(0, 72.0, **kwargs)
    assert png_facts == facts
    all_pngs, all_facts = presentation.render_all_slides_with_diagnostics(72.0, **kwargs)
    assert all_pngs == presentation.render_all_slides(72.0, **kwargs)
    assert all_facts == facts
    frames = presentation.text_layout(**kwargs)
    assert len(frames) == 2
    assert any("Aptos" in value and "approximate" in value for value in frames[0].diagnostics)
    assert not any("Aptos" in value for value in frames[1].diagnostics)
    assert presentation.text_layout(**kwargs) == frames

    notes = _fx195_presentation("Aptos")
    notes.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.name = "Calibri"
    assert not any("Aptos" in value for value in notes.to_pdf_with_diagnostics(**kwargs)[1])
    before_notes = notes.to_bytes()
    notes_pdf, notes_facts = notes.to_notes_pdf_with_diagnostics(**kwargs)
    assert notes_pdf == notes.to_notes_pdf(**kwargs)
    assert isinstance(notes_facts, tuple)
    assert sum("Font 'Aptos' is approximate" in value for value in notes_facts) == 1
    notes_png, png_facts = notes.render_all_notes_with_diagnostics(72.0, **kwargs)
    assert notes_png == notes.render_all_notes(72.0, **kwargs)
    assert png_facts == notes_facts
    assert notes_png and all(page.startswith(b"\x89PNG") for page in notes_png)
    assert notes.to_notes_pdf_with_diagnostics(**kwargs) == (notes_pdf, notes_facts)
    assert notes.to_bytes() == before_notes


def test_fx206_picture_placeholder_opaque_empty_body_refuses_without_package_mutation():
    import io
    import rpptx
    from PIL import Image
    from lxml import etree

    deck = rpptx.Presentation()
    deck.slides.add_slide(deck.slide_layouts[8])
    parts = _package_parts(deck.to_bytes())
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    root = etree.fromstring(parts["ppt/slides/slide1.xml"])
    owners = root.xpath("//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='pic']", namespaces=ns)
    assert len(owners) == 1
    ident = int(owners[0].find("p:nvSpPr/p:cNvPr", ns).get("id"))
    # Only replace this existing empty owner body, leaving all other members exact.
    source = parts["ppt/slides/slide1.xml"].decode()
    shape = next(s for s in _fx206_actual_slide_shapes(source, len(root.xpath("//p:sp", namespaces=ns)))
                 if f'id="{ident}"' in s and 'type="pic"' in s)
    import re
    body = '<p:txBody xmlns:x="urn:producer" x:keep=\'&#65;\'><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
    replacement, changed = re.subn(r"<p:txBody\b[^>]*>.*?</p:txBody>", body, shape, count=1, flags=re.DOTALL)
    assert changed == 1
    parts["ppt/slides/slide1.xml"] = source.replace(shape, replacement, 1).encode()
    actual = rpptx.Presentation.from_bytes(_package_bytes(parts))
    before = actual.to_bytes()
    before_parts = _package_parts(before)
    assert b"x:keep='&#65;'" in before_parts["ppt/slides/slide1.xml"]
    assert rpptx.Presentation.from_bytes(before).to_bytes() == before
    slide = actual.slides[0]
    owner = next(s for s in slide.shapes if s.shape_id == ident)
    ids = [s.shape_id for s in slide.shapes]
    image = io.BytesIO()
    Image.new("RGB", (2, 2), (10, 20, 30)).save(image, format="PNG")
    with pytest.raises(rpptx.RpptxError, match="picture placeholder text body contains source that a picture cannot retain"):
        owner.insert_picture(image.getvalue())
    assert actual.to_bytes() == before
    assert _package_parts(actual.to_bytes()) == before_parts
    assert [s.shape_id for s in slide.shapes] == ids
    assert owner.shape_id == ident


# F-X207 Base controls use existing public construction and measurement first.
# The missing explicit operation is the deciding Base assertion, not save-time
# mutation. The persisted-result assertions remain mandatory after adoption.
def _fx207_fixture(tmp_path, text, mode):
    from pathlib import Path
    import rpptx

    font_path = Path(__file__).resolve().parents[2] / "oxml-layout/fonts/Carlito-Regular.ttf"
    font = font_path.read_bytes()
    assert font[:4] == b"\x00\x01\x00\x00"
    presentation = _text_layout_deck(tmp_path, rpptx.Inches(4), text)
    shape = presentation.slides[0].shapes[0]
    frame = shape.text_frame
    frame.word_wrap = True
    frame.auto_size = mode
    for paragraph in frame.paragraphs:
        paragraph.font.name = "Carlito"
        paragraph.font.size = rpptx.Pt(20)
        for run in paragraph.runs:
            run.font.name = "Carlito"
            run.font.size = rpptx.Pt(20)
    return presentation, [("Carlito", font)]


def _fx207_slide_xml(presentation):
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(io.BytesIO(presentation.to_bytes())) as archive:
        return ET.fromstring(archive.read("ppt/slides/slide1.xml"))


def _fx207_shape_xml(presentation, shape_id):
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    for shape in _fx207_slide_xml(presentation).findall(".//p:sp", ns):
        if shape.find("p:nvSpPr/p:cNvPr", ns).get("id") == str(shape_id):
            return shape
    raise AssertionError("F-X207 source-owned shape is absent")


def _fx207_update(result, shape_id):
    updates = [u for u in result.updates if u.owner_part == "/ppt/slides/slide1.xml"
               and u.shape_id == shape_id and u.cell is None]
    assert len(updates) == 1
    return updates[0]


def test_fx207_explicit_normal_refresh_persists_owned_scale_pair(tmp_path):
    import rpptx

    presentation, fonts = _fx207_fixture(
        tmp_path, "First line\nSecond line\nThird line",
        rpptx.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
    )
    shape = presentation.slides[0].shapes[0]
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    before = presentation.to_bytes()
    normal = _fx207_shape_xml(presentation, shape.shape_id).find(".//a:normAutofit", ns)
    assert normal is not None and normal.attrib == {}
    (measured,) = presentation.text_layout(fonts=fonts)
    assert measured.autofit == "normal" and measured.font_scale < 1.0
    assert presentation.to_bytes() == before
    refresh = getattr(presentation, "refresh_autofit", None)
    assert callable(refresh), "F-X207 missing explicit normal refresh persistence"
    result = refresh(fonts=fonts)
    update = _fx207_update(result, shape.shape_id)
    assert update.mode == "normal" and update.changed and result.changed
    assert update.status == "updated"
    pair = (update.font_scale, update.line_spacing_reduction)
    assert pair in ((100000, 0), (100000, 10000), (92500, 0), (92500, 10000),
                    (92500, 20000), (85000, 20000), (77500, 20000),
                    (70000, 20000), (62500, 20000), (55000, 20000),
                    (47500, 20000), (40000, 20000), (32500, 20000), (25000, 20000))
    normal = _fx207_shape_xml(presentation, shape.shape_id).find(".//a:normAutofit", ns)
    assert normal.attrib == {"fontScale": str(pair[0]), "lnSpcReduction": str(pair[1])}
    saved = presentation.to_bytes()
    assert saved != before
    assert not presentation.text_layout(fonts=fonts)[0].overflow
    assert presentation.to_bytes() == saved
    assert not refresh(fonts=fonts).changed
    assert presentation.to_bytes() == saved
    reopened = rpptx.Presentation.from_bytes(saved)
    assert _fx207_shape_xml(reopened, shape.shape_id).find(".//a:normAutofit", ns).attrib == normal.attrib


def test_fx207_explicit_shape_refresh_persists_changed_bounds(tmp_path):
    import rpptx

    presentation, fonts = _fx207_fixture(
        tmp_path, "First line\nSecond line\nThird line",
        rpptx.MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT,
    )
    shape = presentation.slides[0].shapes[0]
    old = (shape.left, shape.top, shape.width, shape.height)
    before = presentation.to_bytes()
    assert old[3] == rpptx.Pt(36)
    (measured,) = presentation.text_layout(fonts=fonts)
    assert measured.autofit == "shape" and measured.overflow
    assert (shape.left, shape.top, shape.width, shape.height) == old
    assert presentation.to_bytes() == before
    refresh = getattr(presentation, "refresh_autofit", None)
    assert callable(refresh), "F-X207 missing explicit shape refresh persistence"
    result = refresh(fonts=fonts)
    update = _fx207_update(result, shape.shape_id)
    current = presentation.slides[0].shapes[0]
    new = (current.left, current.top, current.width, current.height)
    assert update.mode == "shape" and update.status == "updated"
    assert update.old_bounds == old and update.new_bounds == new
    assert new[:3] == old[:3] and new[3] > old[3]
    assert update.changed and result.changed
    saved = presentation.to_bytes()
    assert not presentation.text_layout(fonts=fonts)[0].overflow
    assert not refresh(fonts=fonts).changed
    assert presentation.to_bytes() == saved
    reopened = rpptx.Presentation.from_bytes(saved).slides[0].shapes[0]
    assert (reopened.left, reopened.top, reopened.width, reopened.height) == new


def test_fx207_fit_text_persists_uniform_size_without_autofit(tmp_path):
    import rpptx

    presentation, fonts = _fx207_fixture(
        tmp_path, "First line\nSecond line\nThird line",
        rpptx.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
    )
    frame = presentation.slides[0].shapes[0].text_frame
    original = frame.text
    before = presentation.to_bytes()
    fit = getattr(frame, "fit_text", None)
    assert callable(fit), "F-X207 missing explicit checked TextFrame.fit_text"
    update = fit(font_family=None, max_size=12, fonts=fonts)
    assert isinstance(update.diagnostics, tuple)
    assert update.mode == "none" and update.status == "updated" and update.changed
    assert isinstance(update.point_size, int) and 1 <= update.point_size <= 12
    current = presentation.slides[0].shapes[0].text_frame
    assert current.text == original and current.word_wrap is True
    assert current.auto_size == rpptx.MSO_AUTO_SIZE.NONE
    for paragraph in current.paragraphs:
        for run in paragraph.runs:
            assert run.font.name == "Carlito"
            assert run.font.size == rpptx.Pt(update.point_size)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    source = _fx207_shape_xml(presentation, update.shape_id)
    assert len(source.findall(".//a:noAutofit", ns)) == 1
    assert not source.findall(".//a:normAutofit", ns)
    assert {r.get("sz") for r in source.findall(".//a:rPr", ns)} == {str(update.point_size * 100)}
    saved = presentation.to_bytes()
    assert saved != before
    assert not presentation.text_layout(fonts=fonts)[0].overflow
    reopened = rpptx.Presentation.from_bytes(saved)
    assert reopened.slides[0].shapes[0].text_frame.text == original


def test_fx207_refresh_accounts_for_flagged_empty_frames_and_is_explicit(tmp_path):
    import rpptx

    presentation, fonts = _fx207_fixture(
        tmp_path, "Short text", rpptx.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
    )
    slide = presentation.slides[0]
    first = slide.shapes[0]
    first_id = first.shape_id
    empty = slide.shapes.add_textbox(rpptx.Inches(1), rpptx.Inches(2), rpptx.Inches(4), rpptx.Pt(36))
    empty.text_frame.auto_size = rpptx.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    before = presentation.to_bytes()
    empty_before = _fx207_shape_xml(presentation, empty.shape_id)
    import xml.etree.ElementTree as ET
    empty_bytes = ET.tostring(empty_before)
    assert len(presentation.text_layout(fonts=fonts)) == 1
    assert presentation.to_bytes() == before
    refresh = getattr(presentation, "refresh_autofit", None)
    assert callable(refresh), "F-X207 missing explicit complete flagged-owner refresh"
    result = refresh(fonts=fonts)
    # The stock deck also has flagged master owners. Account for this slide
    # without excluding the complete physical-owner inventory.
    slide_updates = [u for u in result.updates if u.owner_part == "/ppt/slides/slide1.xml"]
    assert len(slide_updates) == 2
    assert [u.shape_id for u in slide_updates] == [first_id, empty.shape_id]
    nonempty = _fx207_update(result, first_id)
    assert (nonempty.font_scale, nonempty.line_spacing_reduction) == (100000, 0)
    empty_update = _fx207_update(result, empty.shape_id)
    assert empty_update.status == "empty" and not empty_update.changed
    assert empty_update.font_scale is None and empty_update.point_size is None
    assert ET.tostring(_fx207_shape_xml(presentation, empty.shape_id)) == empty_bytes
    saved = presentation.to_bytes()
    assert not refresh(fonts=fonts).changed
    assert presentation.to_bytes() == saved


@pytest.mark.parametrize("bad_size", [True, False, 0, -1, 1.5, 4001, 2**100])
def test_fx207_fit_size_validation_is_strict_and_atomic(tmp_path, bad_size):
    import rpptx

    presentation, fonts = _fx207_fixture(tmp_path, "Strict size", rpptx.MSO_AUTO_SIZE.NONE)
    frame = presentation.slides[0].shapes[0].text_frame
    before = presentation.to_bytes()
    expected = TypeError if isinstance(bad_size, bool) or isinstance(bad_size, float) else ValueError
    with pytest.raises(expected, match="max_size"):
        frame.fit_text(max_size=bad_size, fonts=fonts)
    assert presentation.to_bytes() == before
    assert frame.text == "Strict size"


def test_fx207_snapshots_and_existing_handles_remain_owned_and_read_only(tmp_path):
    import rpptx

    presentation, fonts = _fx207_fixture(
        tmp_path, "First line\nSecond line\nThird line", rpptx.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
    )
    slide = presentation.slides[0]
    shape = slide.shapes[0]
    frame = shape.text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    result = presentation.refresh_autofit(fonts=fonts)
    update = _fx207_update(result, shape.shape_id)
    assert isinstance(result.updates, tuple) and isinstance(result.diagnostics, tuple)
    assert isinstance(update.diagnostics, tuple)
    old_pair = (update.font_scale, update.line_spacing_reduction)
    for owner, field, value in [(result, "changed", False), (update, "font_scale", 1),
                                (update, "old_bounds", (0, 0, 1, 1))]:
        with pytest.raises(AttributeError):
            setattr(owner, field, value)
    assert frame.text == "First line\nSecond line\nThird line"
    assert paragraph.text == "First line" and run.text == "First line"
    fitted = frame.fit_text(max_size=12, fonts=fonts)
    assert fitted.mode == "none" and run.font.size == rpptx.Pt(fitted.point_size)
    assert shape.shape_id == update.shape_id and slide.shapes[0].shape_id == shape.shape_id
    assert (update.font_scale, update.line_spacing_reduction) == old_pair
    assert not frame.fit_text(max_size=12, fonts=fonts).changed
    assert frame.text == "First line\nSecond line\nThird line"


def test_fx207_invalid_font_directory_and_late_unfit_preserve_the_complete_deck(tmp_path):
    import rpptx

    presentation, fonts = _fx207_fixture(
        tmp_path, "First line\nSecond line\nThird line", rpptx.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
    )
    before = presentation.to_bytes()
    missing = tmp_path / "missing-caller-directory"
    with pytest.raises(OSError, match="missing-caller-directory"):
        presentation.refresh_autofit(fonts=fonts, font_dirs=[missing])
    assert presentation.to_bytes() == before
    late = presentation.slides[0].shapes.add_textbox(
        rpptx.Inches(1), rpptx.Inches(3), rpptx.Inches(4), rpptx.Pt(36),
    )
    late.text_frame.text = "Unfit line\n" * 500
    late = presentation.slides[0].shapes[1]
    late.text_frame.auto_size = rpptx.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    before = presentation.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="slide1.xml.*floor"):
        presentation.refresh_autofit(fonts=fonts)
    assert presentation.to_bytes() == before
    assert presentation.slides[0].shapes[0].text_frame.text == "First line\nSecond line\nThird line"


def test_fx208_deck_footer_public_base(tmp_path):
    from rpptx import Presentation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[0])
    prs.slides.add_slide(prs.slide_layouts[1])
    path = tmp_path / "footer-positive.pptx"
    prs.save(path)
    assert len(Presentation(path).slides) == 2
    prs.set_header_footer(slide_number=True, footer="F208 owned footer", date="fixed date")
    assert prs.slides[1].header_footer.footer == "F208 owned footer"
    assert prs.slides[0].header_footer.footer is None
    hf = prs.slides[1].header_footer
    assert hf.source == "slide" and hf.policy_source == "layout"
    assert hf.footer_policy is True and hf.footer_present
    before = prs.to_bytes()
    prs.set_header_footer(slide_number=True, footer="F208 owned footer", date="fixed date")
    assert prs.to_bytes() == before
    reopened = Presentation.from_bytes(before)
    assert reopened.slides[1].header_footer.date == "fixed date"
    from rpptx import RpptxError
    with pytest.raises(RpptxError):
        prs.set_header_footer(footer="invalid\x00")
    assert prs.to_bytes() == before
    prs.set_header_footer(date="auto", date_cache="same", hide_on_title=False)
    hf = prs.slides[1].header_footer
    assert hf.date_mode == "automatic" and hf.date_cache == "same"
    hf.date = "same"
    assert hf.date_mode == "fixed" and hf.date_cache is None
    assert Presentation.from_bytes(prs.to_bytes()).slides[1].header_footer.date_mode == "fixed"
    before = prs.to_bytes()
    hf.date = "same"
    assert prs.to_bytes() == before


def test_fx208_slide_footer_public_base():
    from rpptx import Presentation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[1])
    assert len(prs.slides[0].shapes) >= 2
    prs.slides[0].header_footer.footer = "F208 local"
    assert prs.slides[0].header_footer.footer == "F208 local"
    hf = prs.slides[0].header_footer
    hf.date = "auto"
    hf.date_format = "datetime13"
    hf.slide_number = True
    assert hf.footer == "F208 local" and hf.date_format == "datetime13"
    before = prs.to_bytes()
    from rpptx import RpptxError
    with pytest.raises(RpptxError):
        hf.date_format = "datetime14"
    assert prs.to_bytes() == before
    hf.footer = ""
    assert hf.footer == "" and hf.footer_present
    hf.footer = None
    assert hf.footer is None and not hf.footer_present


def test_fx208_transition_public_base():
    from rpptx import Presentation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[1])
    assert len(prs.slides) == 1
    prs.slides[0].transition.type = "split"
    prs.slides[0].transition.direction = "horizontal-in"
    assert prs.slides[0].transition.type == "split"
    transition = prs.slides[0].transition
    directions = {
        "fade": [None], "cut": [None],
        "push": [None, "left", "up", "right", "down"],
        "wipe": [None, "left", "up", "right", "down"],
        "cover": [None, "left", "up", "right", "down", "left-up", "right-up", "left-down", "right-down"],
        "uncover": [None, "left", "up", "right", "down", "left-up", "right-up", "left-down", "right-down"],
        "split": [None, "horizontal-in", "horizontal-out", "vertical-in", "vertical-out"],
        "zoom": [None, "in", "out"],
    }
    for kind, allowed in directions.items():
        transition.type = kind
        for direction in allowed:
            transition.direction = direction
            transition.duration = 1.2349
            transition.advance_after = 0.0009
            transition.advance_on_click = False
            assert transition.duration == 1.234
            assert transition.advance_after == 0.0
            assert transition.advance_on_click is False
            reopened = Presentation.from_bytes(prs.to_bytes()).slides[0].transition
            assert reopened.type == kind and reopened.direction == direction
            assert reopened.duration == 1.234 and reopened.advance_after == 0.0
            before = prs.to_bytes()
            transition.duration = 1.2349
            assert prs.to_bytes() == before
    before = prs.to_bytes()
    for invalid in [float("nan"), float("inf"), -0.01]:
        with pytest.raises(ValueError):
            transition.duration = invalid
        assert prs.to_bytes() == before
    with pytest.raises(ValueError):
        transition.type = "unknown"
    assert prs.to_bytes() == before
    transition.type = None
    assert transition.type is None and transition.duration is None
    assert transition.advance_after is None and transition.advance_on_click is None


def test_fx208_master_theme_public_base(tmp_path):
    from rpptx import Presentation
    prs = Presentation()
    assert len(prs.slide_layouts) >= 2
    master = prs.slide_masters[0]
    assert len(master.theme.colors) == 12
    assert master.theme.fonts.major.latin
    assert set(master.theme.colors) == {"dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"}
    assert len(master.slide_layouts) == len(prs.slide_layouts)
    theme = master.theme
    before = prs.to_bytes()
    theme.name = "changed"
    assert theme.name == master.theme.name == "changed"
    after = prs.to_bytes()
    original_parts, changed_parts = _package_parts(before), _package_parts(after)
    assert original_parts.keys() == changed_parts.keys()
    assert [name for name in original_parts if original_parts[name] != changed_parts[name]] == ["ppt/theme/theme1.xml"]
    with pytest.raises(TypeError):
        theme.name = 17
    assert prs.to_bytes() == after and theme.name == "changed"
    output = tmp_path / "named-theme.pptx"
    prs.save(output)
    assert Presentation(output).slide_masters[0].theme.name == "changed"
    prs.slides.add_slide(prs.slide_layouts[1])
    assert theme.name == master.theme.name


def test_fx208_placeholder_metadata_public_base():
    from rpptx import Presentation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[0])
    title = prs.slides[0].shapes.title
    assert title is not None
    assert title.is_placeholder
    assert title.placeholder_format.idx == 0
    assert int(title.placeholder_format.type) == 3
    from rpptx import PP_PLACEHOLDER
    assert title.placeholder_format.type == PP_PLACEHOLDER.CENTER_TITLE
    metadata = title.placeholder_format
    title.text = "New title"
    assert metadata.idx == 0
    assert metadata.type == PP_PLACEHOLDER.CENTER_TITLE


def test_fx208_field_authoring_public_base():
    from rpptx import Presentation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    shape = prs.slides[0].shapes.add_textbox(914400, 914400, 3657600, 914400)
    shape.text_frame.paragraphs[0].add_run("Page ")
    paragraph = prs.slides[0].shapes[-1].text_frame.paragraphs[0]
    assert paragraph.text == "Page "
    paragraph.add_field("slidenum", "cached")
    assert "cached" in prs.slides[0].shapes[-1].text_frame.paragraphs[0].text
    import io
    import zipfile
    import xml.etree.ElementTree as ET
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    for kind, cache in [("datetime1", None), ("datetime13", "")]:
        prs.slides[0].shapes[-1].text_frame.paragraphs[0].add_field(kind, cache)
    with zipfile.ZipFile(io.BytesIO(prs.to_bytes())) as package:
        fields = ET.fromstring(package.read("ppt/slides/slide1.xml")).findall(".//a:fld", ns)
    assert len(fields) == 3 and len({field.get("id") for field in fields}) == 3
    assert fields[1].find("a:t", ns) is None
    assert fields[2].find("a:t", ns) is not None
    before = prs.to_bytes()
    from rpptx import RpptxError
    with pytest.raises(RpptxError):
        prs.slides[0].shapes[-1].text_frame.paragraphs[0].add_field("unknown")
    assert prs.to_bytes() == before
    assert Presentation.from_bytes(before).to_bytes() == before


def test_fx208_number_start_public_base(tmp_path):
    from rpptx import Presentation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[1])
    assert len(prs.slides) == 1
    prs.slide_number_start = 0
    path = tmp_path / "start-zero.pptx"
    prs.save(path)
    assert Presentation(path).slide_number_start == 0
    before = prs.to_bytes()
    prs.slide_number_start = 0
    assert prs.to_bytes() == before
    with pytest.raises((OverflowError, ValueError)):
        prs.slide_number_start = -1
    assert prs.to_bytes() == before
    prs.slide_number_start = None
    assert Presentation.from_bytes(prs.to_bytes()).slide_number_start is None


def _fx205_text_owner():
    from rpptx import Presentation
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_textbox(0, 0, 1270000, 1270000).text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.add_run("F205 parity")
    assert run.text == "F205 parity"
    return prs, slide, frame.paragraphs[0], run


@pytest.mark.parametrize("name,value", [("baseline", .25), ("spacing", 127), ("language", "ja-JP"), ("east_asian_name", "Arial"), ("complex_script_name", "Arial")])
def test_fx205_extended_font_public_base(name, value):
    from rpptx import Presentation
    prs, _, _, run = _fx205_text_owner()
    font = run.font
    assert getattr(font, name) is None
    setattr(font, name, value)
    assert getattr(font, name) == value
    reopened = Presentation.from_bytes(prs.to_bytes())
    actual = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
    assert getattr(actual, name) == value
    setattr(font, name, None)
    assert getattr(font, name) is None


@pytest.mark.parametrize("name,value", [("auto_number", "arabicPeriod"), ("auto_number_start", 3), ("bullet_color", "336699"), ("bullet_font", "Arial"), ("bullet_size", 1.25)])
def test_fx205_bullet_public_base(name, value):
    prs, _, paragraph, _ = _fx205_text_owner()
    getattr(paragraph, name)
    if name == "auto_number_start":
        paragraph.auto_number = "arabicPeriod"
    setattr(paragraph, name, value)
    actual = getattr(paragraph, name)
    if name == "bullet_color":
        assert str(actual) == value
    else:
        assert actual == value
    before = prs.to_bytes()
    setattr(paragraph, name, value)
    assert prs.to_bytes() == before


@pytest.mark.parametrize("choice", ["gradient", "patterned", "picture", "transparency"])
def test_fx205_fill_public_base(tmp_path, choice):
    prs, slide, _, _ = _fx205_text_owner()
    fill = slide.shapes[0].fill
    if choice == "picture":
        image = tmp_path / "owned.png"
        _write_tiny_png(image)
        fill.picture(image)
    elif choice == "transparency":
        fill.solid()
        fill.fore_color.rgb = "336699"
        fill.transparency = .25
        assert fill.transparency == .25
    elif choice == "gradient":
        fill.gradient()
        assert len(fill.gradient_stops) >= 2
        fill.gradient_angle = 45.0
        assert fill.gradient_angle == 45.0
    else:
        fill.patterned()
        assert fill.type is not None
    assert prs.to_bytes()


@pytest.mark.parametrize("name", ["first_row", "last_row", "first_col", "last_col", "horz_banding", "vert_banding", "style_id"])
def test_fx205_table_style_public_base(name):
    from rpptx import Presentation
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(2, 2, 0, 0, 1270000, 1270000).table
    assert len(table.rows) == 2 and len(table.columns) == 2
    getattr(table, name)
    value = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" if name == "style_id" else True
    setattr(table, name, value)
    assert getattr(table, name) == value
    assert Presentation.from_bytes(prs.to_bytes()).slides[0].shapes[0].table.cell(0, 0).text_frame is not None


@pytest.mark.parametrize("owner", ["core_properties", "app_properties", "custom_properties"])
def test_fx205_properties_public_base(owner):
    prs, _, _, _ = _fx205_text_owner()
    properties = getattr(prs, owner)
    if owner == "core_properties":
        properties.author = "F205 author"
        assert properties.author == "F205 author"
    elif owner == "app_properties":
        properties.application = "F205 app"
        assert properties.application == "F205 app"
    else:
        properties.set("F205", "value")
        assert properties.get("F205").value == "value"
    assert prs.to_bytes()


def test_fx205_sections_public_base():
    prs, _, _, _ = _fx205_text_owner()
    assert prs.sections == ()
    prs.set_sections([("F205 section", [0])])
    assert prs.sections[0].name == "F205 section"


def test_fx205_media_public_base(tmp_path):
    prs, slide, _, _ = _fx205_text_owner()
    import wave
    path = tmp_path / "sound.wav"
    with wave.open(str(path), "wb") as stream:
        stream.setnchannels(1)
        stream.setsampwidth(2)
        stream.setframerate(8000)
        stream.writeframes(b"\x00\x00" * 80)
    movie = slide.shapes.add_movie(path, 0, 0, 1270000, 1270000, mime_type="audio/wav")
    assert slide.media
    assert slide.extract_media(movie.shape_id) == path.read_bytes()
    slide.remove_media(movie.shape_id)
    assert slide.media == ()


@pytest.mark.parametrize("format", ["jpeg", "tiff", "pdfa", "handout", "odp"])
def test_fx205_exports_public_base(format):
    prs, _, _, _ = _fx205_text_owner()
    if format in ("jpeg", "tiff"):
        output = prs.render_slides(format=format, dpi=24.0)
        if format == "jpeg":
            assert len(output) == 1 and output[0].startswith(b"\xff\xd8")
        else:
            assert output.startswith((b"II*\x00", b"MM\x00*"))
    elif format == "pdfa":
        assert prs.to_pdfa(profile="pdfa-2b").startswith(b"%PDF")
    elif format == "handout":
        assert prs.to_handout_pdf(slides_per_page=2).startswith(b"%PDF")
    else:
        output, diagnostics = prs.to_odp()
        assert output.startswith(b"PK") and isinstance(diagnostics, list)


def test_fx205_existing_cell_and_notes_capabilities_are_real():
    from rpptx import Presentation
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(1, 1, 0, 0, 1270000, 1270000).table
    table.cell(0, 0).text_frame.paragraphs[0].add_run("F205 existing cell")
    assert table.cell(0, 0).text == "F205 existing cell"
    assert prs.to_notes_pdf().startswith(b"%PDF")


def test_fx205_fresh_all_layouts_match_native_widescreen_rectangles():
    # Native baseline75261bde, actual PowerPoint16.113.3, retained projection.
    import xml.etree.ElementTree as ET
    from rpptx import Presentation
    prs = Presentation()
    assert prs.slide_width == 12192000 and prs.slide_height == 6858000
    expected = {'Blank': [('dt', '10', None, None), ('ftr', '11', None, None), ('sldNum', '12', None, None)],
 'Comparison': [('title',
                 '0',
                 {'x': '839788', 'y': '365125'},
                 {'cx': '10515600', 'cy': '1325563'}),
                ('body',
                 '1',
                 {'x': '839788', 'y': '1681163'},
                 {'cx': '5157787', 'cy': '823912'}),
                ('obj',
                 '2',
                 {'x': '839788', 'y': '2505075'},
                 {'cx': '5157787', 'cy': '3684588'}),
                ('body',
                 '3',
                 {'x': '6172200', 'y': '1681163'},
                 {'cx': '5183188', 'cy': '823912'}),
                ('obj',
                 '4',
                 {'x': '6172200', 'y': '2505075'},
                 {'cx': '5183188', 'cy': '3684588'}),
                ('dt', '10', None, None),
                ('ftr', '11', None, None),
                ('sldNum', '12', None, None)],
 'Content with Caption': [('title',
                           '0',
                           {'x': '839788', 'y': '457200'},
                           {'cx': '3932237', 'cy': '1600200'}),
                          ('obj',
                           '1',
                           {'x': '5183188', 'y': '987425'},
                           {'cx': '6172200', 'cy': '4873625'}),
                          ('body',
                           '2',
                           {'x': '839788', 'y': '2057400'},
                           {'cx': '3932237', 'cy': '3811588'}),
                          ('dt', '10', None, None),
                          ('ftr', '11', None, None),
                          ('sldNum', '12', None, None)],
 'Picture with Caption': [('title',
                           '0',
                           {'x': '839788', 'y': '457200'},
                           {'cx': '3932237', 'cy': '1600200'}),
                          ('pic',
                           '1',
                           {'x': '5183188', 'y': '987425'},
                           {'cx': '6172200', 'cy': '4873625'}),
                          ('body',
                           '2',
                           {'x': '839788', 'y': '2057400'},
                           {'cx': '3932237', 'cy': '3811588'}),
                          ('dt', '10', None, None),
                          ('ftr', '11', None, None),
                          ('sldNum', '12', None, None)],
 'Section Header': [('title',
                     '0',
                     {'x': '831850', 'y': '1709738'},
                     {'cx': '10515600', 'cy': '2852737'}),
                    ('body',
                     '1',
                     {'x': '831850', 'y': '4589463'},
                     {'cx': '10515600', 'cy': '1500187'}),
                    ('dt', '10', None, None),
                    ('ftr', '11', None, None),
                    ('sldNum', '12', None, None)],
 'Title Only': [('title', '0', None, None),
                ('dt', '10', None, None),
                ('ftr', '11', None, None),
                ('sldNum', '12', None, None)],
 'Title Slide': [('ctrTitle',
                  '0',
                  {'x': '1524000', 'y': '1122363'},
                  {'cx': '9144000', 'cy': '2387600'}),
                 ('subTitle',
                  '1',
                  {'x': '1524000', 'y': '3602038'},
                  {'cx': '9144000', 'cy': '1655762'}),
                 ('dt', '10', None, None),
                 ('ftr', '11', None, None),
                 ('sldNum', '12', None, None)],
 'Title and Content': [('title', '0', None, None),
                       ('obj', '1', None, None),
                       ('dt', '10', None, None),
                       ('ftr', '11', None, None),
                       ('sldNum', '12', None, None)],
 'Title and Vertical Text': [('title', '0', None, None),
                             ('body', '1', None, None),
                             ('dt', '10', None, None),
                             ('ftr', '11', None, None),
                             ('sldNum', '12', None, None)],
 'Two Content': [('title', '0', None, None),
                 ('obj',
                  '1',
                  {'x': '838200', 'y': '1825625'},
                  {'cx': '5181600', 'cy': '4351338'}),
                 ('obj',
                  '2',
                  {'x': '6172200', 'y': '1825625'},
                  {'cx': '5181600', 'cy': '4351338'}),
                 ('dt', '10', None, None),
                 ('ftr', '11', None, None),
                 ('sldNum', '12', None, None)],
 'Vertical Title and Text': [('title',
                              '0',
                              {'x': '8724900', 'y': '365125'},
                              {'cx': '2628900', 'cy': '5811838'}),
                             ('body',
                              '1',
                              {'x': '838200', 'y': '365125'},
                              {'cx': '7734300', 'cy': '5811838'}),
                             ('dt', '10', None, None),
                             ('ftr', '11', None, None),
                             ('sldNum', '12', None, None)],
 '__master__': [('title',
                 '0',
                 {'x': '838200', 'y': '365125'},
                 {'cx': '10515600', 'cy': '1325563'}),
                ('body',
                 '1',
                 {'x': '838200', 'y': '1825625'},
                 {'cx': '10515600', 'cy': '4351338'}),
                ('dt',
                 '2',
                 {'x': '838200', 'y': '6356350'},
                 {'cx': '2743200', 'cy': '365125'}),
                ('ftr',
                 '3',
                 {'x': '4038600', 'y': '6356350'},
                 {'cx': '4114800', 'cy': '365125'}),
                ('sldNum',
                 '4',
                 {'x': '8610600', 'y': '6356350'},
                 {'cx': '2743200', 'cy': '365125'})]}
    p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    actual = {}
    with zipfile.ZipFile(io.BytesIO(prs.to_bytes())) as package:
        for part in package.namelist():
            if not (part.startswith("ppt/slideLayouts/slideLayout") or part == "ppt/slideMasters/slideMaster1.xml") or not part.endswith(".xml"):
                continue
            root = ET.fromstring(package.read(part))
            common = root.find(p + "cSld")
            key = common.get("name") if "slideLayouts/" in part else "__master__"
            owners = []
            for shape in common.find(p + "spTree"):
                placeholder = shape.find("./" + p + "nvSpPr/" + p + "nvPr/" + p + "ph")
                if placeholder is None:
                    continue
                transform = shape.find("./" + p + "spPr/" + a + "xfrm")
                offset = transform.find(a + "off") if transform is not None else None
                extent = transform.find(a + "ext") if transform is not None else None
                owners.append((placeholder.get("type", "obj"), placeholder.get("idx", "0"), dict(offset.attrib) if offset is not None else None, dict(extent.attrib) if extent is not None else None))
            actual[key] = owners
    assert len(actual) == 12
    assert actual == expected


@pytest.mark.parametrize("lexical,expected", [("1pt",12700), ("1in",914400), ("1cm",360000), ("1mm",36000), ("1pc",152400), ("1pi",152400), ("-1pt",-12700)])
def test_fx205_imported_universal_spacing_remains_present(lexical, expected):
    from rpptx import Presentation
    import re
    from rpptx.util import Pt
    prs, _, _, authored_run = _fx205_text_owner()
    authored_run.font.size = Pt(12)
    parts = _package_parts(prs.to_bytes())
    source = parts["ppt/slides/slide1.xml"].decode()
    source, count = re.subn(r"<a:rPr([^>]*)/>", lambda match: '<a:rPr' + match.group(1) + ' spc="' + lexical + '"/>', source, count=1)
    assert count == 1
    parts["ppt/slides/slide1.xml"] = source.encode()
    imported = Presentation.from_bytes(_package_bytes(parts))
    run = imported.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    before = imported.to_bytes()
    assert int(run.font.spacing) == expected
    assert imported.to_bytes() == before
    run.font.spacing = 127
    assert int(run.font.spacing) == 127
    reopened = Presentation.from_bytes(imported.to_bytes())
    assert int(reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.spacing) == 127


# F-X209 controls start on the claimed installed capability surface.
@pytest.mark.parametrize("operation", [
    "master-logo", "layout-logo", "layout-hide", "master-background",
    "master-title-style", "layout-rename", "layout-duplicate", "layout-remove",
    "layout-master", "layout-usage", "layout-follow-background", "slide-hide",
    "theme-color", "theme-font", "apply-theme", "import-master",
])
def test_fx209_part_owned_authoring_capabilities_persist(operation):
    from rpptx import Presentation, Inches
    from rpptx.dml.color import RGBColor
    prs = Presentation()
    layout = prs.slide_layouts[0]
    master = prs.slide_master
    if operation == "master-logo":
        master.shapes.add_picture(_tiny_png(), Inches(0.1), Inches(0.1))
        assert len(master.shapes) > 0
    elif operation == "layout-logo":
        layout.shapes.add_picture(_tiny_png(), Inches(0.1), Inches(0.1))
        assert len(layout.shapes) > 0
    elif operation == "layout-hide":
        layout.show_master_shapes = False
        assert layout.show_master_shapes is False
    elif operation == "master-background":
        master.background.fill.solid()
        master.background.fill.fore_color.rgb = RGBColor(17, 34, 51)
        assert str(master.background.fill.fore_color.rgb) == "112233"
    elif operation == "master-title-style":
        master.text_styles.title.font.name = "Carlito"
        assert master.text_styles.title.font.name == "Carlito"
    elif operation == "layout-rename":
        layout.name = "Owned layout"
        assert layout.name == "Owned layout"
    elif operation == "layout-duplicate":
        count = len(prs.slide_layouts)
        copied = prs.slide_layouts.duplicate(layout)
        assert len(prs.slide_layouts) == count + 1
        assert copied.name == prs.slide_layouts[0].name
    elif operation == "layout-remove":
        count = len(prs.slide_layouts)
        prs.slide_layouts.remove(layout)
        assert len(prs.slide_layouts) == count - 1
    elif operation == "layout-master":
        assert layout.slide_master == master
    elif operation == "layout-usage":
        prs.slides.add_slide(layout)
        assert len(prs.slide_layouts[0].used_by_slides) == 1
    elif operation == "layout-follow-background":
        layout.follow_master_background = False
        assert layout.follow_master_background is False
    elif operation == "slide-hide":
        slide = prs.slides.add_slide(layout)
        slide.show_master_shapes = False
        assert slide.show_master_shapes is False
    elif operation == "theme-color":
        master.theme.colors["accent1"] = RGBColor(17, 34, 51)
        assert str(prs.slide_master.theme.colors["accent1"]) == "112233"
    elif operation == "theme-font":
        master.theme.fonts.major.latin = "Carlito"
        assert prs.slide_master.theme.fonts.major.latin == "Carlito"
    else:
        source = Presentation()
        source_bytes = source.to_bytes()
        before = len(prs.slide_masters)
        prs.apply_theme(source_bytes, import_masters=operation == "import-master")
        assert len(prs.slide_masters) == before + int(operation == "import-master")
    reopened = Presentation.from_bytes(prs.to_bytes())
    assert len(reopened.slide_masters) == len(prs.slide_masters)
    if operation == "layout-rename":
        assert reopened.slide_layouts[0].name == "Owned layout"
    if operation == "theme-color":
        assert str(reopened.slide_master.theme.colors["accent1"]) == "112233"
    if operation == "theme-font":
        assert reopened.slide_master.theme.fonts.major.latin == "Carlito"


def test_fx209_layout_lifecycle_keeps_surviving_handles_and_refuses_used_owner():
    from rpptx import Presentation, RpptxError, StaleElementError
    prs = Presentation()
    layouts = prs.slide_layouts
    source = layouts[0]
    survivor = layouts[1]
    master = source.slide_master
    source_name, survivor_name = source.name, survivor.name
    copied = layouts.duplicate(source)
    assert source.name == source_name and survivor.name == survivor_name
    assert source.slide_master == master and copied.slide_master == master
    assert len(prs.slide_layouts) == 12
    prs.slide_layouts.remove(source)
    with pytest.raises(StaleElementError, match="SlideLayoutCollection.remove"):
        _ = source.name
    assert survivor.name == survivor_name and prs.slide_layouts[0] == survivor
    assert copied.slide_master == master
    slide = prs.slides.add_slide(copied)
    before = prs.to_bytes()
    with pytest.raises(RpptxError, match="used|inbound"):
        prs.slide_layouts.remove(copied)
    assert prs.to_bytes() == before
    assert slide.slide_layout == copied and copied.name == source_name
    reopened = Presentation.from_bytes(before)
    assert len(reopened.slide_layouts) == 11
    assert reopened.slides[0].slide_layout.name == source_name


@pytest.mark.parametrize("script", ["logos", "deck-theme", "brand-potx"])
def test_fx209_original_three_branding_scripts_render_and_preserve_content(script, tmp_path):
    from rpptx import Presentation, Pt
    from rpptx.dml.color import RGBColor
    prs = Presentation()
    if script == "logos":
        master = prs.slide_master
        master.shapes.add_picture(_tiny_png(0x20, 0x80, 0xE0), Pt(20), Pt(20), Pt(40), Pt(40))
        prs.slide_layouts[0].show_master_shapes = False
        prs.slide_layouts[1].shapes.add_picture(_tiny_png(0xE0, 0x20, 0x40), Pt(100), Pt(20), Pt(40), Pt(40))
        for index in [0, 6, 1]:
            prs.slides.add_slide(prs.slide_layouts[index])
        from PIL import Image
        import io
        pages = [Image.open(io.BytesIO(blob)).convert("RGB") for blob in prs.render_slides(dpi=72)]
        assert len(pages) == 3
        assert pages[0].getpixel((40, 40)) != (0x20, 0x80, 0xE0)
        assert pages[1].getpixel((40, 40)) == (0x20, 0x80, 0xE0)
        assert pages[2].getpixel((40, 40)) == (0x20, 0x80, 0xE0)
        assert pages[2].getpixel((120, 40)) == (0xE0, 0x20, 0x40)
    else:
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Retained branded heading"
        if script == "deck-theme":
            prs.slide_master.theme.colors["accent1"] = RGBColor(17, 34, 51)
            prs.slide_master.theme.fonts.major.latin = "Carlito"
        else:
            brand = Presentation()
            brand.slide_master.theme.colors["accent1"] = RGBColor(17, 34, 51)
            brand.slide_master.theme.fonts.major.latin = "Carlito"
            path = tmp_path / "brand.potx"
            brand.save(path)
            prs.apply_theme(path)
        assert prs.slides[0].shapes.title.text == "Retained branded heading"
        assert str(prs.slide_master.theme.colors["accent1"]) == "112233"
        assert prs.slide_master.theme.fonts.major.latin == "Carlito"
        assert len(prs.render_slides(dpi=72)) == 1
    pdf = prs.to_pdf()
    assert pdf.startswith(b"%PDF-") and len(pdf) > 1000
    path = tmp_path / (script + ".pptx")
    prs.save(path)
    reopened = Presentation(path)
    assert len(reopened.slides) == len(prs.slides)
    if script != "logos":
        assert reopened.slides[0].shapes.title.text == "Retained branded heading"
        assert str(reopened.slide_master.theme.colors["accent1"]) == "112233"
        assert reopened.slide_master.theme.fonts.major.latin == "Carlito"


def test_fx209_master_style_fonts_bullets_levels_and_live_proxies_persist():
    from rpptx import Presentation, Pt
    from rpptx.dml.color import RGBColor
    prs = Presentation()
    master = prs.slide_master
    styles = master.text_styles
    title = styles.title[0]
    font = title.font
    font.name = "Carlito"
    font.size = Pt(40)
    font.color.rgb = RGBColor(17, 34, 51)
    assert font.name == "Carlito" and font.size == Pt(40)
    assert str(font.color.rgb) == "112233"
    for level in range(1, 6):
        paragraph = styles.body.level(level)
        paragraph.font.name = "Carlito"
        paragraph.font.size = Pt(28 - level)
        paragraph.bullet = "•"
        paragraph.bullet_color = RGBColor(51, 34, 17)
        paragraph.left_indent = Pt(level * 12)
        paragraph.first_line_indent = Pt(-6)
    styles.other.font.name = "Carlito"
    saved = prs.to_bytes()
    font.name = "Carlito"
    assert prs.to_bytes() == saved
    for bad in [0, 10]:
        with pytest.raises(ValueError, match="level"):
            styles.body.level(bad)
        assert prs.to_bytes() == saved
    reopened = Presentation.from_bytes(saved)
    assert reopened.slide_master.text_styles.title.font.name == "Carlito"
    assert reopened.slide_master.text_styles.title.font.size == Pt(40)
    for level in range(1, 6):
        paragraph = reopened.slide_master.text_styles.body.level(level)
        assert paragraph.bullet == "•"
        assert str(paragraph.bullet_color) == "332211"
        assert paragraph.font.size == Pt(28 - level)
        assert paragraph.left_indent == Pt(level * 12) and paragraph.first_line_indent == Pt(-6)


def test_fx209_master_layout_placeholders_share_physical_owner_and_persist():
    from rpptx import Presentation
    prs = Presentation()
    owners = [prs.slide_master, prs.slide_layouts[0]]
    for index, owner in enumerate(owners):
        shape = next(shape for shape in owner.shapes if shape.is_placeholder)
        idx = shape.placeholder_format.idx
        placeholder = owner.placeholders[idx]
        assert placeholder.shape_id == shape.shape_id
        placeholder.text = f"Physical placeholder {index}"
        assert next(shape for shape in owner.shapes if shape.shape_id == placeholder.shape_id).text == f"Physical placeholder {index}"
        with pytest.raises(IndexError):
            _ = owner.placeholders[999999]
    saved = prs.to_bytes()
    reopened = Presentation.from_bytes(saved)
    for index, owner in enumerate([reopened.slide_master, reopened.slide_layouts[0]]):
        shape = next(shape for shape in owner.shapes if shape.is_placeholder)
        assert owner.placeholders[shape.placeholder_format.idx].text == f"Physical placeholder {index}"
    assert reopened.to_bytes() == saved


def test_fx204_batches_keep_counts_notes_rollback_and_noop_handles():
    import rpptx
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 3000000, 800000)
    shape.text = "A A"
    prs.slides[0].notes_text = "A"
    held = prs.slides[0].shapes[-1].text_frame.paragraphs[0].runs[0]
    before = prs.to_bytes()
    assert prs.replace_all([]) == ()
    assert prs.replace_all([("missing", "x", 0)]) == (0,)
    assert prs.replace_all([("A", "A", 3)]) == (3,)
    assert held.text == "A A" and prs.to_bytes() == before
    with pytest.raises(rpptx.ReplacementCountError) as raised:
        prs.replace_all([("A", "B", 3), ("B", "C", 4)])
    assert raised.value.index == 1 and raised.value.placeholder == "B"
    assert raised.value.expected == 4 and raised.value.found == 3
    assert prs.to_bytes() == before and held.text == "A A"
    with pytest.raises(rpptx.RpptxError):
        prs.replace_all([("A", "B"), ("", "C")])
    assert prs.to_bytes() == before and held.text == "A A"
    assert prs.replace_all([("A", "B", 3), ("B", "C", 3)]) == (3, 3)
    reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
    assert reopened.slides[0].shapes[-1].text == "C C"
    assert reopened.slides[0].notes_text == "C"
    with pytest.raises(rpptx.StaleElementError):
        _ = held.text


def test_fx204_add_at_preserves_physical_survivors_and_refusals():
    import rpptx
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 3000000, 800000).text = "survivor"
    held = prs.slides[0].shapes[-1]
    before = prs.to_bytes()
    with pytest.raises(rpptx.RpptxError):
        prs.slides.add_slide_at(prs.slide_layouts[6], 2)
    assert prs.to_bytes() == before and held.text == "survivor"
    other = rpptx.Presentation()
    with pytest.raises(ValueError):
        prs.slides.add_slide_at(other.slide_layouts[6], 0)
    assert prs.to_bytes() == before
    new_slide = prs.slides.add_slide_at(prs.slide_layouts[6], 0)
    assert len(prs.slides) == 2 and len(new_slide.shapes) == 0
    assert held.text == "survivor" and prs.slides[1].shapes[-1].text == "survivor"
    new_slide.shapes.add_textbox(0, 0, 3000000, 800000).text = "new"
    reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
    assert reopened.slides[0].shapes[-1].text == "new"
    assert reopened.slides[1].shapes[-1].text == "survivor"


def test_fx204_fit_snapshots_and_notes_are_readonly_and_exact():
    import rpptx
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 3000000, 800000).text = "short"
    before = prs.to_bytes()
    report = prs.fit_report()
    assert isinstance(report, rpptx.FitReport)
    assert isinstance(report.frames, tuple) and report.frames
    assert isinstance(report.diagnostics, tuple)
    frame = report.frames[0]
    assert isinstance(frame, rpptx.FitFrameReport)
    assert frame.update.owner_part and isinstance(frame.update.child_path, list)
    assert not frame.overflow
    assert frame.quantized_trial is None or isinstance(frame.quantized_trial, rpptx.FitTrial)
    with pytest.raises(AttributeError):
        frame.overflow = True
    assert prs.to_bytes() == before
    for text in ["first\r\nnext\n\n", "first\nnext", "first\rnext", "  π 日本語  ", ""]:
        prs.slides[0].notes_text = text
        reopened = rpptx.Presentation.from_bytes(prs.to_bytes())
        assert reopened.slides[0].notes_text == text


@pytest.mark.parametrize("kind", ["shape", "table", "cell", "body", "paragraph", "run", "slide", "layout", "master"])
def test_fx211b_element_nine_physical_owners_are_checked_atomic_and_reopened(kind):
    from rpptx import Element, Presentation, RpptxError
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 1000, 1000)
    shape.text = "original"
    table = slide.shapes.add_table(1, 1, 0, 0, 1000, 1000).table
    cell = table.cell(0, 0)
    cell.text = "cell original"
    owner = {"shape": shape, "table": table, "cell": cell,
             "body": shape.text_frame, "paragraph": shape.text_frame.paragraphs[0],
             "run": shape.text_frame.paragraphs[0].runs[0], "slide": slide,
             "layout": prs.slide_layouts[6], "master": prs.slide_masters[0]}[kind]
    element = owner.element
    assert isinstance(element, Element)
    xml = element.xml
    assert isinstance(xml, str)
    before = prs.to_bytes()
    element.replace_xml(xml)
    assert prs.to_bytes() == before
    assert element.xml == xml
    for bad in ("<foreign/>", xml + "<extra/>", "<!DOCTYPE p [<!ENTITY x 'bad'>]>" + xml):
        with pytest.raises(Exception):
            element.replace_xml(bad)
        assert prs.to_bytes() == before
        assert element.xml == xml
    if kind == "body":
        held_run = shape.text_frame.paragraphs[0].runs[0]
        held_cell_run = cell.text_frame.paragraphs[0].runs[0]
        wrong_owner = '<q:txBody xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:p/></q:txBody>'
        for selected in (element, cell.text_frame.element):
            with pytest.raises(RpptxError, match="expected one complete qualified owner root"):
                selected.replace_xml(wrong_owner)
            assert prs.to_bytes() == before
            assert element.xml == xml
            assert held_run.text == "original"
            assert held_cell_run.text == "cell original"
    if kind in ("layout", "master"):
        def renamed(match):
            opening = match.group(0)
            if 'name="' in opening:
                return re.sub(r'name="[^"]*"', 'name="xml edited"', opening, count=1)
            return opening[:-1] + ' name="xml edited">'
        changed = re.sub(r"<p:cSld\b[^>]*>", renamed, xml, count=1)
    elif kind == "table":
        changed = xml.replace("cell original", "cell edited")
    elif kind == "cell":
        changed = xml.replace("cell original", "cell edited")
    else:
        changed = xml.replace("original", "edited")
    assert changed != xml
    element.replace_xml(changed)
    current = element.xml
    assert "edited" in current
    assert "original" not in current or kind == "slide"  # unrelated cell remains
    reopened = Presentation.from_bytes(prs.to_bytes())
    reopened = Presentation.from_bytes(reopened.to_bytes())
    if kind == "layout":
        assert reopened.slide_layouts[6].name == "xml edited"
    elif kind == "master":
        assert 'name="xml edited"' in reopened.slide_masters[0].element.xml
    elif kind in ("table", "cell"):
        assert reopened.slides[0].shapes[1].table.cell(0, 0).text == "cell edited"
    else:
        assert reopened.slides[0].shapes[0].text == "edited"


def test_fx211b_element_precisely_retires_replaced_children_and_preserves_unrelated_owners():
    from rpptx import Presentation, StaleElementError
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 1000, 1000)
    shape.text = "first"
    other = slide.shapes.add_textbox(0, 0, 1000, 1000)
    other.text = "unrelated"
    frame = shape.text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    font = run.font
    other_run = other.text_frame.paragraphs[0].runs[0]
    selected = paragraph.element
    before = prs.to_bytes()
    selected.replace_xml(selected.xml)
    assert prs.to_bytes() == before and font.bold is None
    selected.replace_xml(selected.xml.replace("first", "second"))
    assert paragraph.text == frame.text == shape.text == "second"
    assert "second" in selected.xml
    assert other_run.text == "unrelated"
    for read in (lambda: run.text, lambda: font.bold):
        with pytest.raises(StaleElementError, match="Element.replace_xml"):
            read()
    new_run = paragraph.runs[0]
    new_font = new_run.font
    new_run.element.replace_xml(new_run.element.xml.replace("second", "third"))
    assert new_run.text == "third"
    with pytest.raises(StaleElementError, match="Element.replace_xml"):
        _ = new_font.bold
    assert paragraph.text == "third"


def test_fx211b_shape_xml_retires_same_value_properties_and_body_remove_recreate():
    from rpptx import Presentation, StaleElementError
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 1000, 1000)
    shape.text = "kept"
    fill = shape.fill
    line = shape.line
    body = shape.text_frame
    old_paragraph = body.paragraphs[0]
    original = shape.element.xml
    changed = original.replace('name="TextBox 1"', 'name="renamed"')
    if changed == original:
        changed = re.sub(r'(name=")[^"]*', r'\1renamed', original, count=1)
    shape.element.replace_xml(changed)
    assert shape.text == "kept"
    for read in (lambda: fill.type, lambda: line.width, lambda: body.text, lambda: old_paragraph.text):
        with pytest.raises(StaleElementError, match="Element.replace_xml"):
            read()
    current_body = shape.text_frame
    source = shape.element.xml
    removed, count = re.subn(r"<p:txBody\b[^>]*>.*?</p:txBody>", "", source, count=1, flags=re.S)
    assert count == 1
    shape.element.replace_xml(removed)
    assert not shape.has_text_frame
    shape.element.replace_xml(source)
    assert shape.text == "kept"
    with pytest.raises(StaleElementError, match="Element.replace_xml"):
        _ = current_body.text
    assert shape.text_frame.text == "kept"


def test_fx211b_table_xml_preserves_table_but_retires_cells_and_whole_shape_retires_table():
    from rpptx import Presentation, StaleElementError
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_table(1, 1, 0, 0, 1000, 1000)
    table = shape.table
    cell = table.cell(0, 0)
    cell.text = "before"
    frame = cell.text_frame
    table.element.replace_xml(table.element.xml.replace("before", "after"))
    assert table.cell(0, 0).text == "after"
    for read in (lambda: cell.text, lambda: frame.text):
        with pytest.raises(StaleElementError, match="Element.replace_xml"):
            read()
    xml = shape.element.xml
    changed = re.sub(r'(name=")[^"]*', r'\1renamed table', xml, count=1)
    shape.element.replace_xml(changed)
    with pytest.raises(StaleElementError, match="Element.replace_xml"):
        _ = table.cell(0, 0)
    assert shape.table.cell(0, 0).text == "after"


def test_fx211b_notes_edit_keeps_unrelated_physical_text_owners():
    from rpptx import Presentation
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 1000, 1000)
    shape.text = "outside notes"
    body = shape.text_frame
    paragraph = body.paragraphs[0]
    run = paragraph.runs[0]
    font = run.font
    element = paragraph.element
    slide.notes_text = "new notes"
    assert slide.notes_text == "new notes"
    assert paragraph.text == run.text == body.text == shape.text == "outside notes"
    assert font.bold is None
    assert "outside notes" in element.xml


@pytest.mark.parametrize("kind", ["shape", "table", "cell", "body", "paragraph", "run", "slide", "layout", "master"])
def test_fx211b_element_qualified_alias_and_default_roots_preserve_actual_source(kind):
    from rpptx import Presentation
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 1000, 1000)
    shape.text = "alias text"
    table = slide.shapes.add_table(1, 1, 0, 0, 1000, 1000).table
    cell = table.cell(0, 0)
    cell.text = "cell text"
    owner = {"shape": shape, "table": table, "cell": cell,
             "body": shape.text_frame, "paragraph": shape.text_frame.paragraphs[0],
             "run": shape.text_frame.paragraphs[0].runs[0], "slide": slide,
             "layout": prs.slide_layouts[6], "master": prs.slide_masters[0]}[kind]
    element = owner.element
    fully_modeled = kind in ("body", "paragraph", "run", "cell")
    xml = element.xml
    opaque = '<corpus:keep xmlns:corpus="urn:fx211b:opaque" value="corpus:QName">&#65;</corpus:keep>'
    if kind in ("shape", "slide"):
        opaque = f'<p:extLst><p:ext uri="urn:fx211b:opaque">{opaque}</p:ext></p:extLst>'
    if not fully_modeled:
        head, closing = xml.rsplit("</", 1)
        element.replace_xml(head + opaque + "</" + closing)
        xml = element.xml
        assert opaque in xml
    root = re.search(r"<([ap]):(\w+)\b", xml)
    assert root
    prefix, local = root.groups()
    uri = "http://schemas.openxmlformats.org/" + ("presentationml/2006/main" if prefix == "p" else "drawingml/2006/main")
    # Change only the modeled root name. Its children retain their actual
    # existing qualified namespace bindings and opaque lexical payloads.
    alias = xml.replace(f"<{prefix}:{local}", f'<q:{local} xmlns:q="{uri}"', 1)
    alias = alias.replace(f"</{prefix}:{local}>", f"</q:{local}>")
    default = xml.replace(f"<{prefix}:{local}", f'<{local} xmlns="{uri}"', 1)
    default = default.replace(f"</{prefix}:{local}>", f"</{local}>")
    before = prs.to_bytes()
    for incoming in (alias, default):
        element.replace_xml(incoming)
        if fully_modeled:
            assert prs.to_bytes() == before
            assert element.xml == xml
        else:
            # Unknown vocabulary conservatively declines equivalence. A
            # qualified edit still preserves the actual source payloads.
            assert "alias text" in prs.slides[0].shapes[0].text
            assert prs.slides[0].shapes[1].table.cell(0, 0).text == "cell text"
            assert opaque in element.xml
            saved_parts = _package_parts(prs.to_bytes())
            assert any(opaque.encode() in data for name, data in saved_parts.items() if name.endswith(".xml"))
            reopened = Presentation.from_bytes(prs.to_bytes())
            assert reopened.slides[0].shapes[0].text == "alias text"
            assert reopened.slides[0].shapes[1].table.cell(0, 0).text == "cell text"
    before = prs.to_bytes()
    current = element.xml
    element.replace_xml(current)
    assert prs.to_bytes() == before
    wrong = alias.replace(f'xmlns:q="{uri}"', 'xmlns:q="urn:foreign"', 1)
    with pytest.raises(Exception):
        element.replace_xml(wrong)
    assert prs.to_bytes() == before and element.xml == current


def test_fx211b_footer_replacement_retires_actual_changed_nodes_only():
    from rpptx import Presentation, StaleElementError
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    other = slide.shapes.add_textbox(0, 0, 1000, 1000)
    other.text = "unrelated footer owner"
    prs.set_header_footer(slide_number=False, footer="footer first", hide_on_title=False)
    actual = [shape for shape in slide.shapes if shape.has_text_frame and shape.text == "footer first"]
    assert len(actual) == 1
    old_body = actual[0].text_frame
    old_paragraph = old_body.paragraphs[0]
    old_runs = list(old_paragraph.runs)
    other_run = other.text_frame.paragraphs[0].runs[0]
    prs.set_header_footer(slide_number=False, footer="footer second", hide_on_title=False)
    assert slide.header_footer.footer == "footer second"
    assert other_run.text == "unrelated footer owner"
    for read in [lambda: old_paragraph.text] + [lambda run=run: run.text for run in old_runs]:
        with pytest.raises(StaleElementError, match="Presentation.set_header_footer"):
            read()


def test_fx211b_held_layout_element_tracks_actual_part_after_earlier_removal():
    from rpptx import Presentation
    prs = Presentation()
    held = prs.slide_layouts[6]
    slide = prs.slides.add_slide(held)
    element = held.element
    xml = element.xml
    name = held.name
    prs.slide_layouts.remove(prs.slide_layouts[0])
    assert held.name == name
    assert element.xml == xml == held.element.xml
    element.replace_xml(xml)
    assert slide.slide_layout.name == name


def test_fx211b_element_shape_matrix_preserves_five_kinds_and_retires_group_children():
    from rpptx import Presentation, StaleElementError
    from rpptx.enum.shapes import MSO_CONNECTOR
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    text = slide.shapes.add_textbox(0, 0, 1000, 1000)
    text.text = "unrelated survivor"
    picture = slide.shapes.add_picture(io.BytesIO(_tiny_png()), 0, 0, 1000, 1000)
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 1000, 1000)
    table = slide.shapes.add_table(1, 1, 0, 0, 1000, 1000)
    group = slide.shapes.add_group_shape()
    child = group.shapes.add_textbox(0, 0, 1000, 1000)
    child.text = "same group child"
    for owner, root in ((text, "sp"), (picture, "pic"), (connector, "cxnSp"), (table, "graphicFrame"), (group, "grpSp")):
        element = owner.element
        xml = element.xml
        assert re.search(rf"<p:{root}\b", xml)
        before = prs.to_bytes()
        element.replace_xml(xml)
        assert prs.to_bytes() == before
        wrong = re.sub(rf"<p:{root}\b", "<p:wrong", xml, count=1).replace(f"</p:{root}>", "</p:wrong>")
        with pytest.raises(Exception):
            element.replace_xml(wrong)
        assert prs.to_bytes() == before
        changed = re.sub(r'(name=")[^"]*', r'\1xml ' + root, xml, count=1)
        assert changed != xml
        element.replace_xml(changed)
        assert owner.name == "xml " + root
        assert re.search(rf"<p:{root}\b", element.xml)
    with pytest.raises(StaleElementError, match="Element.replace_xml"):
        _ = child.text
    assert group.shapes[0].text == "same group child"
    reopened = Presentation.from_bytes(prs.to_bytes())
    assert [shape.name for shape in reopened.slides[0].shapes] == ["xml sp", "xml pic", "xml cxnSp", "xml graphicFrame", "xml grpSp"]


def test_fx211b_graphic_frame_payload_transition_is_named_and_atomic():
    from rpptx import Presentation
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_table(1, 1, 0, 0, 1000, 1000)
    table = shape.table
    cell = table.cell(0, 0)
    cell.text = "retained"
    xml = shape.element.xml
    changed, count = re.subn(r"<a:graphicData\b[^>]*>.*?</a:graphicData>", '<a:graphicData uri="urn:fx211b:opaque"><payload xmlns="urn:fx211b:opaque"/></a:graphicData>', xml, count=1, flags=re.S)
    assert count == 1
    before = prs.to_bytes()
    with pytest.raises(NotImplementedError, match="table payload presence"):
        shape.element.replace_xml(changed)
    assert prs.to_bytes() == before and cell.text == table.cell(0, 0).text == "retained"


def test_fx211b_virtual_cell_and_inherited_style_elements_refuse_without_materializing():
    from rpptx import Presentation
    prs, _ = _f_x194_table()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(prs.to_bytes())) as source, zipfile.ZipFile(out, "w") as target:
        for member in source.infolist():
            data = source.read(member.filename)
            if member.filename == "ppt/slides/slide1.xml":
                data, count = re.subn(rb"<a:txBody\b[^>]*>.*?</a:txBody>", b"", data, count=1, flags=re.S)
                assert count == 1
            target.writestr(member, data)
    prs = Presentation.from_bytes(out.getvalue())
    cell = prs.slides[0].shapes[0].table.cell(0, 0)
    frame = cell.text_frame
    style = prs.slide_masters[0].text_styles.title.level(1)
    prior_name = style.font.name
    before = prs.to_bytes()
    with pytest.raises(NotImplementedError, match="physical text body"):
        _ = frame.element
    with pytest.raises(NotImplementedError, match="physical XML owner"):
        _ = style.element
    assert prs.to_bytes() == before
    with pytest.raises(ValueError, match="shape has no text frame"):
        _ = frame.text
    assert style.font.name == prior_name


def test_fx220_ppt_python_font_graph_snapshots_inputs_and_removal(tmp_path):
    from pathlib import Path
    from rpptx import Presentation, EmbeddedFont, FontEmbeddingLicense

    source = Path(__file__).resolve().parents[2] / "oxml-layout/fonts/Selawik-Regular.ttf"
    body = source.read_bytes()
    license = FontEmbeddingLicense(authorized=True, identity="SIL Open Font License 1.1")
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 1000000, 1000000).text = "font control"
    held = deck.slides[0]
    for kind, value in zip(("regular", "bold", "italic", "bold_italic"), (body, source, io.BytesIO(body), EmbeddedFont(body, "bold_italic", license=license))):
        deck.embed_font("FX220 Python PPT", value, kind, license=license)
    assert held.shapes[0].text == "font control"
    fonts = deck.embedded_fonts
    assert isinstance(fonts, tuple) and len(fonts) == 1 and fonts[0].typeface == "FX220 Python PPT"
    assert tuple(face.kind for face in fonts[0].faces) == ("regular", "bold", "italic", "bold_italic")
    for face in fonts[0].faces:
        assert face.format == "eot" and face.data == body and face.stored_len > len(body)
        assert face.part_name.startswith("/ppt/fonts/") and face.relationship_id
        with pytest.raises(AttributeError):
            face.format = "unknown"
    saved = deck.to_bytes()
    package = zipfile.ZipFile(io.BytesIO(saved))
    assert len([name for name in package.namelist() if name.startswith("ppt/fonts/")]) == 4
    reopened = Presentation.from_bytes(saved)
    assert tuple(face.data for face in reopened.embedded_fonts[0].faces) == (body,) * 4
    deck.embed_font("FX220 Python PPT", body, license=license)
    assert deck.to_bytes() == saved and held.shapes[0].text == "font control"
    assert deck.remove_embedded_font("FX220 Python PPT", "italic").data == body
    assert deck.remove_embedded_font("FX220 Python PPT", "italic") is None
    assert len(deck.remove_font("FX220 Python PPT").faces) == 3
    assert deck.remove_font("FX220 Python PPT") is None
    output = tmp_path / "removed.pptx"
    deck.save(output)
    assert Presentation(output).embedded_fonts == ()


def test_fx220_ppt_python_license_flags_and_refusal_are_atomic():
    from pathlib import Path
    from rpptx import Presentation, FontEmbeddingLicense

    body = (Path(__file__).resolve().parents[2] / "oxml-layout/fonts/Selawik-Regular.ttf").read_bytes()
    deck = Presentation()
    deck.slides.add_slide(deck.slide_layouts[6])
    # The bundled PowerPoint template explicitly enables subset saving.
    for name, initial in (("embed_true_type_fonts", None), ("save_subset_fonts", True)):
        assert getattr(deck, name) is initial
        for value in (False, True, None):
            setattr(deck, name, value)
            assert getattr(deck, name) is value
    baseline = deck.to_bytes()
    held = deck.slides[0]
    good = FontEmbeddingLicense(authorized=True, identity="OFL 1.1")
    for value, kind, license in ((body, "regular", FontEmbeddingLicense(authorized=False, identity="OFL")), (body, "regular", FontEmbeddingLicense(authorized=True, identity=" ")), (b"not a font", "regular", good), (body, "unknown", good), (io.StringIO("not bytes"), "regular", good)):
        with pytest.raises(Exception):
            deck.embed_font("FX220 refusal", value, kind, license=license)
        assert deck.to_bytes() == baseline and held.shapes is not None
    class OversizedReader:
        def read(self, size):
            assert 0 < size <= 8192
            return b"x" * (64 * 1024 * 1024 + 1)
    with pytest.raises(ValueError):
        deck.embed_font("FX220 refusal", OversizedReader(), license=good)
    assert deck.to_bytes() == baseline and held.shapes is not None


@pytest.mark.parametrize("mode", ["short_chunks", "later_type", "later_error", "cumulative_bound"])
def test_fx220_rpptx_binary_readers_reach_real_eof_before_publication(mode):
    from pathlib import Path
    import rpptx

    body = (Path(__file__).resolve().parents[2] / "oxml-layout/fonts/Selawik-Regular.ttf").read_bytes()
    value = rpptx.Presentation()
    value.slides.add_slide(value.slide_layouts[6]).shapes.add_textbox(0, 0, 1000000, 1000000).text = "binary owner"
    held = value.slides[0].shapes[0]
    baseline = value.to_bytes()
    license = rpptx.FontEmbeddingLicense(authorized=True, identity="OFL 1.1")
    class Reader:
        def __init__(self):
            self.calls = 0
            self.offset = 0
        def read(self, size):
            self.calls += 1
            if mode == "short_chunks":
                part = body[self.offset:self.offset + min(size, 8192)]
                self.offset += len(part)
                return part
            if self.calls == 1:
                return body
            if mode == "later_type":
                return "invalid later text"
            if mode == "later_error":
                raise OSError("later font read failed")
            return b"x" * (1024 * 1024)
    reader = Reader()
    if mode == "short_chunks":
        value.embed_font("FX220 binary stream", reader, license=license)
        assert reader.calls > 2 and reader.offset == len(body)
        assert value.embedded_fonts[0].faces[0].data == body
        authored_reader = Reader()
        authored = rpptx.EmbeddedFont(authored_reader, license=license)
        assert authored.data == body and authored_reader.calls > 2
        value.embed_font("FX220 authored stream", authored, license=license)
        assert value.embedded_fonts[1].faces[0].data == body
    else:
        if mode == "later_error":
            expected, match = OSError, "later font read failed"
        elif mode == "later_type":
            expected, match = TypeError, "reader must return bytes"
        else:
            expected, match = ValueError, "64 MiB"
        with pytest.raises(expected, match=match):
            value.embed_font("FX220 binary stream", reader, license=license)
        assert reader.calls > 1
        assert value.to_bytes() == baseline and held.text == "binary owner"

        with pytest.raises(expected, match=match):
            rpptx.EmbeddedFont(Reader(), license=license)
        assert value.to_bytes() == baseline and held.text == "binary owner"

@pytest.mark.parametrize("method", ["pdf", "png", "all_slides", "notes_pdf", "notes_png", "handout_pdf", "handout_png", "layout"])
def test_fx220_ppt_embedded_faces_and_caller_override_reach_exports(method):
    from pathlib import Path
    from rpptx import Presentation, FontEmbeddingLicense
    import xml.etree.ElementTree as ET

    directory = Path(__file__).resolve().parents[2] / "oxml-layout/fonts"
    kinds = ("regular", "bold", "italic", "bold_italic")
    styles = ("Regular", "Bold", "Italic", "BoldItalic")
    embedded = [(directory / f"Caladea-{style}.ttf").read_bytes() for style in styles]
    caller = [(directory / f"LiberationMono-{style}.ttf").read_bytes() for style in styles]
    name = "FX220 deterministic export face"
    deck = _fx195_presentation(name)
    for index in range(4):
        if index == 0:
            run = deck.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        else:
            run = deck.slides[0].shapes[0].text_frame.paragraphs[0].add_run()
        run.text = "iiiiWWWW route "
        run.font.name = name
        run.font.bold = bool(index & 1)
        run.font.italic = bool(index & 2)
    # Notes are independently source-authored with four actual styled runs.
    parts = _package_parts(deck.to_bytes())
    notes_part = next(part for part in parts if re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", part))
    namespace = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    tree = ET.fromstring(parts[notes_part])
    paragraph = next(parent for parent in tree.iter(namespace + "p") if any("FX195 NOTES" in (child.text or "") for child in parent.iter(namespace + "t")))
    for child in list(paragraph):
        if child.tag == namespace + "r": paragraph.remove(child)
    for index in range(4):
        run = ET.SubElement(paragraph, namespace + "r")
        properties = ET.SubElement(run, namespace + "rPr", {"sz": "2000", "b": str(index & 1), "i": str((index >> 1) & 1)})
        ET.SubElement(properties, namespace + "latin", {"typeface": name})
        ET.SubElement(run, namespace + "t").text = "iiiiWWWW notes route "
    parts[notes_part] = ET.tostring(tree, encoding="utf-8", xml_declaration=True)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as output:
        for part, data in parts.items(): output.writestr(part, data)
    deck = Presentation.from_bytes(buffer.getvalue())
    license = FontEmbeddingLicense(authorized=True, identity="Apache License 2.0")
    for kind, body in zip(kinds, embedded): deck.embed_font(name, body, kind, license=license)
    source = deck.to_bytes()
    held = deck.slides[0]
    disabled = Presentation.from_bytes(source)
    disabled.embed_true_type_fonts = False
    missing = Presentation.from_bytes(source)
    missing.remove_font(name)
    reference = Presentation.from_bytes(missing.to_bytes())

    def invoke(value, fonts):
        if method == "pdf": return value.to_pdf(fonts=fonts)
        if method == "png": return value.render_slide_to_png(0, dpi=72.0, fonts=fonts)
        if method == "all_slides": return value.render_all_slides(dpi=72.0, fonts=fonts)
        if method == "notes_pdf": return value.to_notes_pdf(fonts=fonts)
        if method == "notes_png": return value.render_all_notes(dpi=72.0, fonts=fonts)
        if method == "handout_pdf": return value.to_handout_pdf(slides_per_page=2, fonts=fonts)
        if method == "handout_png": return value.render_all_handouts(slides_per_page=2, dpi=72.0, fonts=fonts)
        return tuple(tuple((line.bounds.width, line.bounds.height) for line in frame.lines) for frame in value.text_layout(fonts=fonts))

    positive = invoke(deck, [])
    off = invoke(disabled, [])
    absent = invoke(missing, [])
    override = invoke(deck, list(zip([name] * 4, caller)))
    expected = invoke(reference, list(zip([name] * 4, embedded)))
    expected_override = invoke(reference, list(zip([name] * 4, caller)))
    assert positive != off and positive != override
    assert off == absent
    if method.endswith("pdf") or method == "pdf":
        assert positive.startswith(b"%PDF") and b"Caladea" in positive
        assert b"LiberationMono" in override and b"Caladea" not in override
        assert b"Caladea" not in off
    else:
        assert positive == expected and override == expected_override
    assert deck.to_bytes() == source and held.shapes[0].text == "iiiiWWWW route " * 4


@pytest.mark.parametrize("kind", [51, 65])
def test_fx215_ppt_public_create_live_replace_and_atomic_noop(kind):
    from rpptx import Presentation, CategoryChartData, Inches, GraphicFrame
    data = CategoryChartData()
    data.categories = ["East", "West"]
    data.add_series("Sales", [11.0, 23.0])
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_chart(kind, Inches(1), Inches(1), Inches(4), Inches(2), data)
    assert isinstance(frame, GraphicFrame) and frame.has_chart
    chart = frame.chart
    assert chart.chart_type == kind and chart.categories == ("East", "West")
    assert chart.series[0].name == "Sales" and chart.series[0].values == [11.0, 23.0]
    before = prs.to_bytes()
    chart.replace_data(data)
    assert prs.to_bytes() == before and frame.has_chart
    bad = CategoryChartData(); bad.categories = ["East"]; bad.add_series("Bad", [float("inf")])
    with pytest.raises(Exception): chart.replace_data(bad)
    assert prs.to_bytes() == before and chart.categories == ("East", "West")
    replacement = CategoryChartData(); replacement.categories = ["North", "South", "Central"]
    replacement.add_series("Updated", [31.0, 47.0, 59.0])
    chart.replace_data(replacement)
    assert chart.categories == ("North", "South", "Central")
    assert chart.series[0].values == [31.0, 47.0, 59.0]
    output = prs.to_bytes()
    with zipfile.ZipFile(io.BytesIO(output)) as archive:
        name = next(name for name in archive.namelist() if name.startswith("ppt/charts/chart") and name.endswith(".xml"))
        assert b"Updated" in archive.read(name) and b"59" in archive.read(name)


def test_fx215_shared_xy_refuses_heterogeneous_coordinates_before_publish():
    from rpptx import Presentation, XyChartData, Inches
    data = XyChartData(); a = data.add_series("A"); b = data.add_series("B")
    for x, y in [(1.0, 3.0), (2.0, 5.0)]: a.add_data_point(x, y); b.add_data_point(x, y + 1)
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_chart(74, Inches(1), Inches(1), Inches(4), Inches(2), data)
    assert frame.chart.categories == ("1", "2") or frame.chart.categories == ("1.0", "2.0")
    before = prs.to_bytes()
    b.add_data_point(3.0, 8.0)
    with pytest.raises(ValueError, match="same x values"):
        slide.shapes.add_chart(74, Inches(1), Inches(1), Inches(4), Inches(2), data)
    assert prs.to_bytes() == before


def test_fx215_ppt_live_chart_follows_grouped_owner_and_refuses_retired_owner():
    from rpptx import Presentation, CategoryChartData, Inches, StaleElementError
    data = CategoryChartData(); data.categories = ["A"]; data.add_series("Series", [7.0])
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_chart(51, Inches(1), Inches(1), Inches(4), Inches(2), data)
    chart = frame.chart
    group = slide.shapes.group([frame])
    assert chart.categories == ("A",) and group.shapes[0].chart.categories == ("A",)
    replacement = CategoryChartData(); replacement.categories = ["Updated"]; replacement.add_series("New", [19.0])
    chart.replace_data(replacement)
    assert chart.categories == ("Updated",) and group.shapes[0].chart.series[0].values == [19.0]
    slide.shapes.remove(group)
    before = prs.to_bytes()
    with pytest.raises(StaleElementError): chart.replace_data(data)
    assert prs.to_bytes() == before


@pytest.mark.parametrize("replacement", ["properties", "graphic_data"])
def test_fx215_chart_xml_retires_old_view_and_data_edits_preserve_owned_graph(replacement):
    import posixpath
    import xml.etree.ElementTree as ET
    from rpptx import CategoryChartData, Inches, Presentation, StaleElementError

    def data(category, value):
        result = CategoryChartData()
        result.categories = [category]
        result.add_series("Sales", [value])
        return result

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    first = slide.shapes.add_chart(51, Inches(1), Inches(1), Inches(3), Inches(2), data("First", 11.0))
    other = slide.shapes.add_chart(51, Inches(5), Inches(1), Inches(3), Inches(2), data("Other", 23.0))
    held = first.chart
    other_chart = other.chart
    first_id = first.shape_id
    other_xml = other.element.xml
    c = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    rel = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    other_id = ET.fromstring(other_xml).find(".//" + c + "chart").get(r + "id")

    def other_graph():
        with zipfile.ZipFile(io.BytesIO(prs.to_bytes())) as package:
            slide_rels = ET.fromstring(package.read("ppt/slides/_rels/slide1.xml.rels"))
            target = next(node.get("Target") for node in slide_rels.findall(rel) if node.get("Id") == other_id)
            chart_part = posixpath.normpath(posixpath.join("ppt/slides", target))
            chart_rels = posixpath.join(posixpath.dirname(chart_part), "_rels", posixpath.basename(chart_part) + ".rels")
            relationships = ET.fromstring(package.read(chart_rels))
            workbook = next(node.get("Target") for node in relationships.findall(rel) if node.get("Type").endswith("/package"))
            workbook_part = posixpath.normpath(posixpath.join(posixpath.dirname(chart_part), workbook))
            return tuple(package.read(part) for part in (chart_part, chart_rels, workbook_part))

    original_other_graph = other_graph()
    xml = first.element.xml
    expected = "First"
    if replacement == "properties":
        xml = re.sub(r'(name=")[^"]*', r'\1checked chart', xml, count=1)
    else:
        old_id = ET.fromstring(xml).find(".//" + c + "chart").get(r + "id")
        assert old_id != other_id and f'r:id="{old_id}"' in xml
        xml = xml.replace(f'r:id="{old_id}"', f'r:id="{other_id}"', 1)
        expected = "Other"
    first.element.replace_xml(xml)
    assert first.has_chart and first.shape_id == first_id
    before_refusal = prs.to_bytes()
    with pytest.raises(StaleElementError, match="Element.replace_xml"):
        _ = held.categories
    with pytest.raises(StaleElementError, match="Element.replace_xml"):
        held.replace_data(data("Refused", 99.0))
    assert prs.to_bytes() == before_refusal

    current = first.chart
    assert current.categories == (expected,)
    current.replace_data(data("Updated", 47.0))
    assert current.categories == ("Updated",) and current.series[0].values == [47.0]
    assert first.has_chart and first.shape_id == first_id
    assert other_chart.categories == ("Other",) and other_chart.series[0].values == [23.0]
    assert other.element.xml == other_xml and other_graph() == original_other_graph
    saved = prs.to_bytes()
    current.replace_data(data("Updated", 47.0))
    assert prs.to_bytes() == saved and current.categories == ("Updated",)
    reopened = Presentation.from_bytes(saved)
    assert reopened.slides[0].shapes[0].chart.categories == ("Updated",)
    assert reopened.slides[0].shapes[1].chart.categories == ("Other",)


@pytest.mark.parametrize("format", ["html", "markdown"])
def test_fx218_detached_semantic_export_and_atomic_publication(format, tmp_path):
    import rpptx
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    held = slide.shapes.add_textbox(0, 0, 3000000, 800000)
    held.text = "Visible & literal <text>"
    slide.notes_text = "Private speaker note"
    png = _tiny_png()
    slide.shapes.add_picture(io.BytesIO(png), 0, 0)
    slide.shapes.add_picture(io.BytesIO(png), 10000, 0)
    before = prs.to_bytes()
    to_text = getattr(prs, "to_" + format)
    export = getattr(prs, "export_" + format)
    save = getattr(prs, "save_" + format)
    assert isinstance(to_text(), str)
    result = export(include_notes=True, inline_images=False)
    assert isinstance(result, rpptx.PresentationTextExport)
    expected_text = "Visible &amp; literal &lt;text&gt;" if format == "html" else r"Visible \& literal \<text\>"
    assert expected_text in result.text.replace("&#32;", " ")
    assert "Private speaker note" in result.text.replace("&#32;", " ")
    assert isinstance(result.assets, tuple) and len(result.assets) == 1
    name, body = result.assets[0]
    assert isinstance(name, str) and body == png and isinstance(body, bytes)
    assert result.text.count(name) == 2
    assert isinstance(result.diagnostics, tuple)
    with pytest.raises(AttributeError):
        result.text = "changed"
    with pytest.raises(ValueError, match="retain image assets"):
        to_text(inline_images=False)
    output = tmp_path / ("result.html" if format == "html" else "result.md")
    for bad in [1, "yes", object()]:
        with pytest.raises(TypeError):
            save(output, include_notes=bad)
        assert not output.exists()
    asset = tmp_path / name
    asset.write_bytes(b"keep existing asset")
    with pytest.raises(Exception):
        save(output)
    assert asset.read_bytes() == b"keep existing asset"
    assert not output.exists()
    asset.unlink()
    assert isinstance(save(output, include_notes=True), tuple)
    assert asset.read_bytes() == png
    assert output.read_text() == result.text
    published = output.read_bytes()
    with pytest.raises(Exception):
        save(output)
    assert output.read_bytes() == published
    assert prs.to_bytes() == before
    assert held.text == "Visible & literal <text>"
    assert rpptx.Presentation.from_bytes(before).slides[0].notes_text == "Private speaker note"


@pytest.mark.parametrize("method", ["add", "add_at", "single", "diagnostics", "batch"])
def test_fx216_foreign_layout_precedes_stale_collection_without_publication(method):
    import rpptx

    destination = rpptx.Presentation()
    local_layout = destination.slide_layouts[6]
    held = destination.slides
    original = held.add_slide(local_layout)
    source = rpptx.Presentation()
    foreign_layout = source.slide_layouts[6]
    source_slide = source.slides.add_slide(foreign_layout)
    original_xml = original.element.xml
    before = destination.to_bytes(), source.to_bytes()

    def operation(layout):
        if method == "add":
            return held.add_slide(layout)
        if method == "add_at":
            return held.add_slide_at(layout, 0)
        if method == "single":
            return held.import_slide(source_slide, layout=layout)
        if method == "diagnostics":
            return held.import_slide_with_diagnostics(source_slide, layout=layout)
        return held.import_slides(source, [0], layout=layout)

    message = (
        "layout belongs to a different presentation"
        if method in ("add", "add_at")
        else "layout is not a layout of this presentation"
    )
    with pytest.raises(ValueError, match=message):
        operation(foreign_layout)
    _assert_exact_stale(rpptx, lambda: operation(local_layout), "slide collection", 0, 1,
                        "Re-fetch it with prs.slides.", "SlideCollection.add_slide")
    assert (destination.to_bytes(), source.to_bytes()) == before
    assert len(destination.slides) == len(source.slides) == 1
    assert original.element.xml == original_xml and len(original.shapes) == 0
    assert local_layout.name == destination.slide_layouts[6].name
    assert foreign_layout.name == source.slide_layouts[6].name


def test_fx216_batch_returns_live_ordered_owners_and_empty_selection_is_noop():
    import rpptx

    source = rpptx.Presentation()
    source_layout = source.slide_layouts[6]
    for text in ("first", "second", "third"):
        slide = source.slides.add_slide(source_layout)
        slide.shapes.add_textbox(0, 0, 914400, 914400).text = text
    source_before = source.to_bytes()
    destination = rpptx.Presentation()
    layout = destination.slide_layouts[6]
    existing = destination.slides.add_slide(layout)
    existing.shapes.add_textbox(0, 0, 914400, 914400).text = "existing"
    imported, diagnostics = destination.slides.import_slides(source, [2, 0], layout=layout, index=0)
    assert diagnostics == ()
    assert isinstance(imported, tuple) and len(imported) == 2
    assert [slide.shapes[0].text for slide in imported] == ["third", "first"]
    assert [slide.shapes[0].text for slide in destination.slides] == ["third", "first", "existing"]
    assert existing.shapes[0].text == "existing"
    held = destination.slides
    before = destination.to_bytes()
    assert held.import_slides(source, [], layout=layout) == ((), ())
    assert len(held) == 3 and destination.to_bytes() == before
    with pytest.raises(rpptx.RpptxError, match="duplicate selected source slide"):
        held.import_slides(source, [0, 0], layout=layout)
    with pytest.raises(IndexError):
        held.import_slides(source, [0, 99], layout=layout)
    assert len(held) == 3 and destination.to_bytes() == before
    next_slide, next_diagnostics = held.import_slide_with_diagnostics(source.slides[1], layout=layout, index=0)
    assert next_diagnostics == () and next_slide.shapes[0].text == "second"
    assert [slide.shapes[0].text for slide in imported] == ["third", "first"]
    assert existing.shapes[0].text == "existing"
    assert source.to_bytes() == source_before
    reopened = rpptx.Presentation.from_bytes(destination.to_bytes())
    assert [slide.shapes[0].text for slide in reopened.slides] == ["second", "third", "first", "existing"]


def test_fx211c_paragraph_text_preserves_its_owner_and_retires_replaced_runs():
    from rpptx import Presentation, StaleElementError
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 1270000, 1270000)
    paragraph = shape.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "old run"
    paragraph = shape.text_frame.paragraphs[0]
    old_run = paragraph.runs[0]
    paragraph.text = "replacement"
    paragraph.level = 1
    paragraph.font.bold = True
    assert paragraph.text == "replacement"
    before = prs.to_bytes()
    with pytest.raises(StaleElementError, match="Paragraph.text"):
        old_run.text = "unsafe"
    assert prs.to_bytes() == before
    reopened = Presentation.from_bytes(before)
    actual = reopened.slides[0].shapes[0].text_frame.paragraphs[0]
    assert actual.text == "replacement" and actual.level == 1 and actual.font.bold is True


def test_fx211c_row_cells_are_checked_live_grid_views():
    from rpptx import Presentation, StaleElementError
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(2, 2, 0, 0, 1270000, 1270000).table
    row = table.rows[0]
    cells = row.cells
    assert isinstance(cells, tuple) and len(cells) == 2
    cells[0].text = "first"
    cells[1].text = "second"
    assert tuple(cell.text for cell in row.cells) == ("first", "second")
    reopened = Presentation.from_bytes(prs.to_bytes())
    assert tuple(cell.text for cell in reopened.slides[0].shapes[0].table.rows[0].cells) == ("first", "second")
    table.rows.add_row()
    before = prs.to_bytes()
    with pytest.raises(StaleElementError):
        _ = row.cells
    with pytest.raises(StaleElementError):
        cells[0].text = "unsafe"
    assert prs.to_bytes() == before


def test_fx211c_theme_name_is_live_and_preserves_source_parts():
    from rpptx import Presentation
    prs = Presentation()
    master = prs.slide_masters[0]
    theme = master.theme
    prs.slides.add_slide(prs.slide_layouts[6])
    before = zipfile.ZipFile(io.BytesIO(prs.to_bytes()))
    before_parts = {name: before.read(name) for name in before.namelist()}
    theme.name = 'Held & "quoted" theme'
    assert theme.name == master.theme.name == 'Held & "quoted" theme'
    saved = prs.to_bytes()
    after = zipfile.ZipFile(io.BytesIO(saved))
    changed = {name for name in before_parts if before_parts[name] != after.read(name)}
    assert len(changed) == 1 and next(iter(changed)).startswith("ppt/theme/")
    stable = prs.to_bytes()
    theme.name = theme.name
    assert prs.to_bytes() == stable
    assert Presentation.from_bytes(saved).slide_masters[0].theme.name == theme.name


def test_fx211c_alt_text_and_highlight_alias_reach_saved_native_properties():
    from rpptx import Presentation
    import xml.etree.ElementTree as ET
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 1270000, 1270000)
    shape.alt_text = 'Description & "quotes"'
    assert shape.alt_text == 'Description & "quotes"'
    shape.alt_title = "Title"
    shape.decorative = True
    assert shape.alt_title == "Title" and shape.decorative is True
    run = shape.text_frame.paragraphs[0].add_run()
    run.text = "highlight"
    run.font.highlight_color = "FFFF00"
    assert str(run.font.highlight_color.rgb) == "FFFF00"
    saved = prs.to_bytes()
    root = ET.fromstring(zipfile.ZipFile(io.BytesIO(saved)).read("ppt/slides/slide1.xml"))
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    owners = [node for node in root.findall("./p:cSld/p:spTree/p:sp/p:nvSpPr/p:cNvPr", ns) if node.get("id") == str(shape.shape_id)]
    assert len(owners) == 1
    nonvisual = owners[0]
    assert nonvisual.get("descr") == shape.alt_text
    assert nonvisual.get("title") == "Title"
    assert root.find("./p:cSld/p:spTree/p:nvGrpSpPr/p:cNvPr", ns).get("descr") is None
    assert root.find(".//a:rPr/a:highlight/a:srgbClr", ns).get("val") == "FFFF00"
    reopened = Presentation.from_bytes(saved)
    assert reopened.slides[0].shapes[0].alt_text == shape.alt_text
    assert reopened.slides[0].shapes[0].alt_title == "Title"
    assert reopened.slides[0].shapes[0].decorative is True
    before = prs.to_bytes()
    with pytest.raises((ValueError, TypeError)):
        run.font.highlight_color = "invalid"
    assert prs.to_bytes() == before
    run.font.highlight_color = None
    assert run.font.highlight_color.rgb is None


def test_fx211c_arrangement_aliases_and_foreign_refusals_are_atomic():
    from rpptx import Presentation
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shapes = slide.shapes
    a = shapes.add_textbox(0, 0, 127000, 127000)
    b = shapes.add_textbox(254000, 254000, 127000, 127000)
    c = shapes.add_textbox(508000, 508000, 127000, 127000)
    shapes.align([a, b], "left", relative_to="selection")
    assert a.left == b.left == 0
    shapes.distribute([a, b, c], "horizontal", relative_to="selection")
    assert b.left == 254000
    before = prs.to_bytes()
    with pytest.raises(ValueError, match="disagree"):
        shapes.align([a, b], "left", "slide", relative_to="selection")
    with pytest.raises(ValueError, match="disagree"):
        shapes.distribute([a, b, c], "horizontal", "slide", relative_to="selection")
    assert prs.to_bytes() == before
    foreign = Presentation()
    other = foreign.slides.add_slide(foreign.slide_layouts[6])
    other_before = foreign.to_bytes()
    with pytest.raises(ValueError, match="owner|presentation"):
        other.shapes.group([a])
    assert foreign.to_bytes() == other_before and prs.to_bytes() == before


def test_fx211c_absent_core_author_retains_none_and_roundtrips():
    from rpptx import Presentation
    prs = Presentation()
    properties = prs.core_properties
    assert properties.author is None
    before = prs.to_bytes()
    properties.author = None
    assert prs.to_bytes() == before
    properties.author = "Author"
    assert properties.author == "Author"
    assert Presentation.from_bytes(prs.to_bytes()).core_properties.author == "Author"
    properties.author = ""
    assert properties.author is None
    assert Presentation.from_bytes(prs.to_bytes()).core_properties.author is None


def test_fx211c_autofit_approximation_facts_are_owned_and_deduplicated():
    from rpptx import Presentation
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    selected_ids = []
    for index in range(2):
        shape = slide.shapes.add_textbox(0, index * 914400, 4000000, 914400)
        shape.text = "Actual contributing text"
        paragraph = shape.text_frame.paragraphs[0]
        paragraph.runs[0].font.name = "Aptos"
        shape.text_frame.auto_size = 2
        selected_ids.append(shape.shape_id)
    updated = prs.refresh_autofit()
    facts = [value for value in updated.diagnostics if "Font 'Aptos' is approximate" in value]
    assert len(facts) == 1 and "Carlito" in facts[0]
    selected = [update for update in updated.updates if update.owner_part == "/ppt/slides/slide1.xml" and update.shape_id in selected_ids]
    assert len(selected) == 2
    assert all(any("Font 'Aptos' is approximate" in value for value in update.diagnostics) for update in selected)
    again = prs.refresh_autofit()
    assert again.diagnostics == updated.diagnostics
    assert not again.changed
    report = prs.fit_report()
    assert sum("Font 'Aptos' is approximate" in value for value in report.diagnostics) == 1



def test_fx211c_bullet_char_alias_qualifies_numbering_four_families_and_render(tmp_path):
    from pathlib import Path
    import xml.etree.ElementTree as ET
    from rpptx import Presentation
    from rpptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shapes = slide.shapes
    numbered = shapes.add_textbox(0, 0, Inches(3), Inches(1)).text_frame.paragraphs[0]
    numbered.text = "Numbered"
    numbered.auto_number = "arabicPeriod"
    numbered.auto_number_start = 3
    numbered.bullet_color = "112233"
    numbered.bullet_size = 1.25
    numbered.bullet_font = "Arial"
    assert numbered.bullet_char is True
    assert numbered.auto_number == "arabicPeriod" and numbered.auto_number_start == 3
    assert str(numbered.bullet_color) == "112233"
    assert numbered.bullet_size == 1.25 and numbered.bullet_font == "Arial"

    pairs = [("Symbol", "\uf0b7"), ("Wingdings", "\uf0a7"), ("Arial", "\uf0a7"), ("Wingdings", "\uf0ff")]
    held = []
    for index, (family, character) in enumerate(pairs):
        paragraph = shapes.add_textbox(0, Inches((index + 1) * .45), Inches(3), Inches(.4)).text_frame.paragraphs[0]
        paragraph.text = family + " control"
        paragraph.bullet_font = family
        paragraph.bullet_color = "336699"
        paragraph.bullet_size = 1.1
        paragraph.bullet_char = character
        assert paragraph.bullet_char == paragraph.bullet == character
        assert paragraph.bullet_font == family and str(paragraph.bullet_color) == "336699"
        assert paragraph.bullet_size == 1.1
        held.append(paragraph)
    before = prs.to_bytes()
    for paragraph, (_, character) in zip(held, pairs):
        paragraph.bullet_char = character
    assert prs.to_bytes() == before
    # Bundled caller bytes keep this regression independent of Office installation.
    font = (Path(__file__).resolve().parents[2] / "oxml-layout/fonts/Carlito-Regular.ttf").read_bytes()
    pdf, diagnostics = prs.to_pdf_with_diagnostics(fonts=[("Arial", font), ("Calibri", font)])
    assert pdf.startswith(b"%PDF-") and len(re.findall(rb"/Type\s*/Page\b", pdf)) == 1
    assert isinstance(diagnostics, tuple)
    assert any("Wingdings" in value and ("missing" in value.lower() or "unknown" in value.lower()) for value in diagnostics)
    assert prs.to_bytes() == before
    (tmp_path / "four-family-bullets.pdf").write_bytes(pdf)
    reopened = Presentation.from_bytes(before)
    actual = [shape.text_frame.paragraphs[0] for shape in reopened.slides[0].shapes]
    assert actual[0].auto_number == "arabicPeriod" and actual[0].auto_number_start == 3
    assert str(actual[0].bullet_color) == "112233" and actual[0].bullet_size == 1.25 and actual[0].bullet_font == "Arial"
    assert [(paragraph.bullet_font, paragraph.bullet_char) for paragraph in actual[1:]] == pairs
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    tree = ET.fromstring(zipfile.ZipFile(io.BytesIO(before)).read("ppt/slides/slide1.xml"))
    properties = tree.findall(".//a:pPr", ns)
    assert properties[0].find("a:buAutoNum", ns).get("startAt") == "3"
    assert [(node.find("a:buFont", ns).get("typeface"), node.find("a:buChar", ns).get("char")) for node in properties[1:]] == pairs
    assert numbered.text == "Numbered" and all(paragraph.text.endswith(" control") for paragraph in held)


@pytest.mark.parametrize("invalid", ["", True, 17])
def test_fx211c_bullet_char_invalid_alias_is_atomic_with_live_owners(invalid):
    prs, _, paragraph, run = _fx205_text_owner()
    paragraph.bullet_char = "•"
    paragraph.bullet_font = "Arial"
    before = prs.to_bytes()
    with pytest.raises((TypeError, ValueError)):
        paragraph.bullet_char = invalid
    assert prs.to_bytes() == before
    assert paragraph.bullet_char == paragraph.bullet == "•"
    assert run.text == "F205 parity" and paragraph.bullet_font == "Arial"
    paragraph.bullet_char = False
    assert paragraph.bullet is False and paragraph.bullet_char is False
    paragraph.bullet_char = None
    assert paragraph.bullet is None and paragraph.bullet_char is None
    before = prs.to_bytes()
    paragraph.bullet_char = None
    assert prs.to_bytes() == before and run.text == "F205 parity"


@pytest.mark.parametrize("method,count,reference,explanation", [
    ("align", 1, "selection", "selection reference requires at least 2, got 1"),
    ("distribute", 1, "selection", "selection reference requires at least 3, got 1"),
    ("distribute", 2, "selection", "selection reference requires at least 3, got 2"),
    ("align", 0, "slide", "no shapes were given"),
    ("distribute", 0, "slide", "no shapes were given"),
])
def test_fx211c_arrangement_minimum_refusal_keeps_bytes_and_held_owners(method, count, reference, explanation):
    from rpptx import Presentation, RpptxError
    prs = Presentation()
    shapes = prs.slides.add_slide(prs.slide_layouts[6]).shapes
    selected = [shapes.add_textbox(index * 254000, 0, 127000, 127000) for index in range(3)]
    for index, shape in enumerate(selected):
        shape.text = str(index)
    before = prs.to_bytes()
    direction = "left" if method == "align" else "horizontal"
    with pytest.raises(RpptxError) as error:
        getattr(shapes, method)(selected[:count], direction, relative_to=reference)
    assert explanation in str(error.value)
    if count:
        assert "too few selected shapes" in str(error.value)
    assert prs.to_bytes() == before
    assert [shape.text for shape in selected] == ["0", "1", "2"]
    # Each slide-reference operation still accepts its distinct one-shape minimum.
    getattr(shapes, method)(selected[:1], direction, relative_to="slide")
    assert selected[0].text == "0"
    shapes.distribute(selected, "horizontal", relative_to="selection")
    reopened = Presentation.from_bytes(prs.to_bytes())
    assert [shape.left for shape in reopened.slides[0].shapes] == [shape.left for shape in selected]


@pytest.mark.parametrize("first_axis", ["rows", "columns"])
def test_fx211c_table_collections_keep_owner_authority_and_retire_indexed_grid_views(first_axis, tmp_path):
    import xml.etree.ElementTree as ET
    import zipfile

    from rpptx import Presentation, StaleElementError

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table_shape = slide.shapes.add_table(2, 2, 0, 0, 4000000, 2000000)
    table = table_shape.table
    table.cell(0, 0).text = "retained first cell"
    other = slide.shapes.add_textbox(0, 2500000, 1000000, 500000)
    other.text = "unrelated sibling"
    other_xml = other.element.xml
    rows, columns = table.rows, table.columns
    foreign = Presentation()
    foreign_slide = foreign.slides.add_slide(foreign.slide_layouts[6])
    foreign_table = foreign_slide.shapes.add_table(2, 2, 0, 0, 1000, 1000).table
    cell, row, column = table.cell(0, 0), rows[0], columns[0]
    stable = prs.to_bytes()
    for collection, foreign_member in ((rows, foreign_table.rows[0]), (columns, foreign_table.columns[0])):
        with pytest.raises(ValueError, match="not in this collection"):
            collection.remove(foreign_member)
        assert prs.to_bytes() == stable
    for collection, operation in ((rows, "add_row"), (columns, "add_column")):
        with pytest.raises(IndexError):
            getattr(collection, operation)(99)
        assert prs.to_bytes() == stable
    row.height = row.height
    column.width = column.width
    assert prs.to_bytes() == stable
    assert cell.text == "retained first cell"
    assert (len(rows), len(columns)) == (2, 2)

    second_axis = "columns" if first_axis == "rows" else "rows"
    for axis in (first_axis, second_axis):
        collection = rows if axis == "rows" else columns
        add = "add_row" if axis == "rows" else "add_column"
        label = "RowCollection" if axis == "rows" else "ColumnCollection"
        for operation in (add, "remove"):
            old_row, old_column, old_cell = rows[0], columns[0], table.cell(0, 0)
            if operation == add:
                added = getattr(collection, operation)()
                assert (added.height if axis == "rows" else added.width) == (
                    rows[-1].height if axis == "rows" else columns[-1].width
                )
            else:
                collection.remove(collection[-1])
            count = 3 if operation == add else 2
            assert (len(rows), len(columns)) == ((count, 2) if axis == "rows" else (2, count))
            assert table_shape.has_table and table.cell(0, 0).text == "retained first cell"
            assert other.text == "unrelated sibling" and other.element.xml == other_xml
            stable = prs.to_bytes()
            for view, attribute, value in ((old_row, "height", 1), (old_column, "width", 1), (old_cell, "text", "unsafe")):
                with pytest.raises(StaleElementError) as error:
                    setattr(view, attribute, value)
                assert f"First invalidated by {label}.{operation} " in str(error.value)
                assert prs.to_bytes() == stable
            stage = tmp_path / f"collection-grid-{first_axis}-{axis}-{operation}.pptx"
            prs.save(stage)
            assert prs.to_bytes() == stable
            saved = Presentation(stage)
            saved_table = saved.slides[0].shapes[0].table
            assert (len(saved_table.rows), len(saved_table.columns)) == (len(rows), len(columns))
            assert saved_table.cell(0, 0).text == "retained first cell"
            assert saved.slides[0].shapes[1].text == "unrelated sibling"
            with zipfile.ZipFile(stage) as package:
                source = ET.fromstring(package.read("ppt/slides/slide1.xml"))
            namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
            assert len(source.findall(f".//{{{namespace}}}tblGrid/{{{namespace}}}gridCol")) == len(columns)
            assert len(source.findall(f".//{{{namespace}}}tbl/{{{namespace}}}tr")) == len(rows)

    target = tmp_path / f"collection-ancestors-{first_axis}.pptx"
    prs.save(target)
    assert prs.to_bytes() == stable
    reopened = Presentation(target)
    saved_table = reopened.slides[0].shapes[0].table
    assert (len(saved_table.rows), len(saved_table.columns)) == (2, 2)
    assert saved_table.cell(0, 0).text == "retained first cell"
    assert reopened.slides[0].shapes[1].text == "unrelated sibling"
    slide.shapes.remove(table_shape)
    stable = prs.to_bytes()
    for action in (lambda: len(rows), lambda: len(columns), lambda: table.cell(0, 0), lambda: table_shape.has_table,
                   lambda: rows.add_row(), lambda: columns.add_column()):
        with pytest.raises(StaleElementError) as error:
            action()
        assert "First invalidated by ShapeCollection.remove " in str(error.value)
        assert prs.to_bytes() == stable
    assert other.text == "unrelated sibling" and other.element.xml == other_xml


def test_fx231_installed_presentation_imported_whitespace_roundtrip(tmp_path):
    import os
    import rpptx

    source = os.environ.get("FX231_INSTALLED_INPUT")
    if source is None:
        # The ordinary suite independently constructs a read/save control.
        # The equipped story gate instead supplies actual native HTML output.
        source = tmp_path / "source.pptx"
        control = rpptx.Presentation()
        slide = control.slides.add_slide(control.slide_layouts[6])
        frame = slide.shapes.add_textbox(0, 0, 4_000_000, 2_000_000).text_frame
        paragraph = frame.paragraphs[0]
        paragraph.add_run("a ")
        control.slides[0].shapes[0].text_frame.paragraphs[0].add_run("b ").font.bold = True
        control.slides[0].shapes[0].text_frame.paragraphs[0].add_run("c\u00a0\u00a0d")
        linked = control.slides[0].shapes[0].text_frame.add_paragraph()
        linked.add_run("x ").hyperlink.address = "https://example.test/"
        control.slides[0].shapes[0].text_frame.paragraphs[1].add_run("y")
        control.save(source)
    with open(source, "rb") as stream:
        data = stream.read()
    for deck in (rpptx.Presentation(source), rpptx.Presentation.from_bytes(data)):
        paragraph = deck.slides[0].shapes[0].text_frame.paragraphs[0]
        assert [(run.text, run.font.bold) for run in paragraph.runs] == [
            ("a ", None), ("b ", True), ("c\u00a0\u00a0d", None)
        ]
        assert all(run.hyperlink.address is None for run in paragraph.runs)
        output = tmp_path / "saved.pptx"
        deck.save(output)
        reopened = rpptx.Presentation(output)
        paragraph = reopened.slides[0].shapes[0].text_frame.paragraphs[0]
        assert [(run.text, run.font.bold) for run in paragraph.runs] == [
            ("a ", None), ("b ", True), ("c\u00a0\u00a0d", None)
        ]

    linked = reopened.slides[0].shapes[0].text_frame.paragraphs[1]
    assert [(run.text, run.hyperlink.address) for run in linked.runs] == [("x ", "https://example.test/"), ("y", None)]


def test_fx246_ppt_legend_controls_defaults_hidden_noops_and_atomicity():
    import rpptx
    data = rpptx.CategoryChartData()
    data.categories = ["Q1", "Q2"]
    data.add_series("North", [12, 14])
    data.add_series("South", [9, 11])
    deck = rpptx.Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    frame = slide.shapes.add_chart(51, rpptx.Inches(.5), rpptx.Inches(1), rpptx.Inches(6), rpptx.Inches(2.6), data)
    chart = frame.chart
    assert (chart.has_legend, chart.legend_position, chart.legend_overlay) == (True, "r", False)
    before = deck.to_bytes()
    chart.has_legend = True
    chart.legend_position = "r"
    chart.legend_overlay = False
    assert deck.to_bytes() == before
    for position in ("r", "l", "t", "b", "tr"):
        chart = deck.slides[0].shapes[0].chart
        chart.legend_position = position
        chart = deck.slides[0].shapes[0].chart
        chart.legend_overlay = True
        assert chart.legend_position == position and chart.legend_overlay
    chart = deck.slides[0].shapes[0].chart
    chart.has_legend = False
    chart = deck.slides[0].shapes[0].chart
    before = deck.to_bytes()
    chart.legend_position = "l"
    chart.legend_overlay = True
    assert deck.to_bytes() == before
    assert (chart.has_legend, chart.legend_position, chart.legend_overlay) == (False, "r", False)
    chart.has_legend = True
    reopened = rpptx.Presentation.from_bytes(deck.to_bytes())
    assert reopened.slides[0].shapes[0].chart.has_legend
    chart = deck.slides[0].shapes[0].chart
    before = deck.to_bytes()
    for property, value, error in (("has_legend", 1, TypeError), ("legend_overlay", "yes", TypeError), ("legend_position", "bad", ValueError), ("legend_position", 1, TypeError)):
        with pytest.raises(error):
            setattr(chart, property, value)
        assert deck.to_bytes() == before


def test_fx246_ppt_grouped_legend_owner_and_retirement():
    import rpptx
    data = rpptx.CategoryChartData()
    data.categories = ["A"]
    data.add_series("Series", [7])
    deck = rpptx.Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    frame = slide.shapes.add_chart(51, rpptx.Inches(1), rpptx.Inches(1), rpptx.Inches(4), rpptx.Inches(2), data)
    chart = frame.chart
    assert not chart.has_legend
    group = slide.shapes.group([frame])
    chart.has_legend = True
    chart.legend_position = "b"
    assert group.shapes[0].chart.legend_position == "b"
    slide.shapes.remove(group)
    before = deck.to_bytes()
    for operation in (lambda: chart.has_legend, lambda: setattr(chart, "has_legend", True), lambda: setattr(chart, "legend_position", "l"), lambda: setattr(chart, "legend_overlay", True)):
        with pytest.raises(rpptx.StaleElementError):
            operation()
        assert deck.to_bytes() == before


@pytest.mark.parametrize("trailer", (b"NOT_XML_DOCUMENT_CONTENT", b"<![CDATA[tail]]>", b"&amp;"))
def test_fx246_ppt_malformed_chart_document_is_refused_atomically(trailer):
    import io
    import zipfile
    import rpptx
    data = rpptx.CategoryChartData()
    data.categories = ["A"]
    data.add_series("North", [7])
    deck = rpptx.Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.shapes.add_chart(51, rpptx.Inches(1), rpptx.Inches(1), rpptx.Inches(4), rpptx.Inches(2), data)
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(deck.to_bytes())) as source, zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            value = source.read(info.filename)
            if info.filename.startswith("ppt/charts/chart") and info.filename.endswith(".xml"):
                value += trailer
            target.writestr(info, value)
    deck = rpptx.Presentation.from_bytes(buffer.getvalue())
    chart = deck.slides[0].shapes[0].chart
    before = deck.to_bytes()
    for operation in (lambda: chart.has_legend, lambda: setattr(chart, "has_legend", True)):
        with pytest.raises(Exception):
            operation()
        assert deck.to_bytes() == before


def _fx239_text_owner():
    import rpptx
    from rpptx.util import Inches
    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    shape.text_frame.text = "kept"
    paragraph = shape.text_frame.paragraphs[0]
    return presentation, slide, shape, paragraph, paragraph.runs[0].font


@pytest.mark.parametrize("initial", [False, True])
@pytest.mark.parametrize("name,refused", [
    ("space_before", 6), ("space_before", 126), ("space_before", False),
    ("space_after", 6), ("space_after", 126), ("space_after", False),
    ("spacing", 2), ("spacing", 126), ("spacing", -126),
    ("spacing", True), ("spacing", False),
    ("line_spacing", 0), ("line_spacing", False), ("line_spacing", 0.000001),
])
def test_fx239_quantization_refusals_are_source_and_revision_atomic(initial, name, refused):
    from rpptx.util import Pt
    presentation, slide, shape, paragraph, font = _fx239_text_owner()
    target = font if name == "spacing" else paragraph
    if initial:
        setattr(target, name, Pt(12))
    before_value = getattr(target, name)
    before = presentation.to_bytes()
    with pytest.raises((TypeError, ValueError), match="EMU|Pt|positive|bool"):
        setattr(target, name, refused)
    assert presentation.to_bytes() == before
    assert getattr(target, name) == before_value
    assert shape.text_frame.text == "kept" and slide.notes_text is None
    setattr(target, name, Pt(12))
    assert getattr(target, name) == Pt(12)


def test_fx239_notes_names_distinguish_absent_empty_and_missing_body():
    import rpptx
    presentation, slide, shape, paragraph, font = _fx239_text_owner()
    before = presentation.to_bytes()
    assert not slide.has_notes_slide
    assert not slide.has_notes_slide
    assert presentation.to_bytes() == before
    notes = slide.notes_slide
    assert isinstance(notes, rpptx.NotesSlide) and slide.has_notes_slide
    frame = notes.notes_text_frame
    assert isinstance(frame, rpptx.NotesTextFrame) and frame.text == ""
    frame.text = "speaker\nsecond line"
    assert frame.text == "speaker\nsecond line" and shape.text_frame.text == "kept"
    reopened = rpptx.Presentation.from_bytes(presentation.to_bytes())
    assert reopened.slides[0].notes_slide.notes_text_frame.text == frame.text


def test_fx239_theme_alias_and_color_import_scope_are_explicit():
    import pptx.enum.dml as oracle
    import rpptx.enum.dml as actual
    assert importlib.metadata.version("python-pptx") == "1.0.2"
    assert actual.MSO_THEME_COLOR_INDEX is actual.MSO_THEME_COLOR
    assert {x.name: x.value for x in actual.MSO_COLOR_TYPE} == {
        x.name: x.value for x in oracle.MSO_COLOR_TYPE
    }
    import rpptx
    *_, font = _fx239_text_owner()
    assert font.color.type is None
    font.color.rgb = rpptx.RGBColor(0x12, 0x34, 0x56)
    assert font.color.type is actual.MSO_COLOR_TYPE.RGB


def _fx239_notes_variant(kind):
    import rpptx
    presentation, slide, *_ = _fx239_text_owner()
    slide.notes_text = "original body"
    with zipfile.ZipFile(io.BytesIO(presentation.to_bytes())) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    name = next(name for name in parts if "/notesSlides/notesSlide" in name and name.endswith(".xml"))
    xml = parts[name]
    marker = xml.index(b'type="body"')
    start = xml.rfind(b"<p:sp>", 0, marker)
    end = xml.index(b"</p:sp>", marker) + len(b"</p:sp>")
    assert start >= 0
    body = xml[start:end]
    if kind == "missing":
        replacement = body.replace(b'type="body"', b'type="title"')
    elif kind == "duplicate":
        replacement = body + re.sub(rb'(<p:cNvPr id=")\d+', rb'\g<1>9999', body, count=1)
    elif kind == "group":
        replacement = b'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9999" name="group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + body + b'</p:grpSp>'
    elif kind == "alternate":
        replacement = b'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback>' + body + b'</mc:Fallback></mc:AlternateContent>'
    elif kind == "no-txBody":
        tx_start = body.index(b"<p:txBody")
        tx_end = body.index(b"</p:txBody>") + len(b"</p:txBody>")
        replacement = body[:tx_start] + body[tx_end:]
    else:
        replacement = body
    xml = xml[:start] + replacement + xml[end:]
    opaque = b'<x:opaque xmlns:x="urn:fx239" x:version="1"><!--keep--><x:child/></x:opaque>'
    xml = xml.replace(b"</p:notes>", b'<p:extLst><p:ext uri="fx239">' + opaque + b'</p:ext></p:extLst></p:notes>')
    if kind == "alias":
        xml = re.sub(rb"(<\/?)p:", rb"\g<1>q:", xml).replace(b"xmlns:p=", b"xmlns:q=")
    elif kind == "default":
        xml = re.sub(rb"(<\/?)p:", rb"\g<1>", xml).replace(b"xmlns:p=", b"xmlns=")
    parts[name] = xml
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as package:
        for part_name, raw in parts.items():
            package.writestr(part_name, raw)
    return rpptx.Presentation.from_bytes(stream.getvalue()), name, opaque


@pytest.mark.parametrize("kind", ["duplicate", "group", "alternate"])
def test_fx239_notes_frame_owner_is_exact_or_refused(kind):
    presentation, _, _ = _fx239_notes_variant(kind)
    slide = presentation.slides[0]
    before = presentation.to_bytes()
    assert slide.has_notes_slide
    assert "original body" in slide.notes_text
    with pytest.raises(Exception, match="one direct Body.*Slide.notes_text"):
        slide.notes_slide.notes_text_frame
    assert presentation.to_bytes() == before


@pytest.mark.parametrize("kind", ["missing", "no-txBody", "alias", "default"])
def test_fx239_notes_text_preserves_owned_and_opaque_source(kind):
    presentation, name, opaque = _fx239_notes_variant(kind)
    slide = presentation.slides[0]
    before = presentation.to_bytes()
    notes = slide.notes_slide
    frame = notes.notes_text_frame
    assert presentation.to_bytes() == before
    if kind == "missing":
        assert frame is None and slide.has_notes_slide
        return
    assert frame.text == ("" if kind == "no-txBody" else "original body")
    frame.text = "new body\nsecond"
    saved = presentation.to_bytes()
    with zipfile.ZipFile(io.BytesIO(saved)) as package:
        assert opaque in package.read(name)
    assert presentation.slides[0].notes_slide.notes_text_frame.text == "new body\nsecond"
    frame.text = frame.text
    assert presentation.to_bytes() == saved


def test_fx239_notes_wrappers_follow_physical_slide_owners():
    import rpptx
    presentation, slide, shape, paragraph, font = _fx239_text_owner()
    notes = slide.notes_slide
    frame = notes.notes_text_frame
    frame.text = "target"
    other = presentation.slides.add_slide(presentation.slide_layouts[6])
    other.notes_text = "other"
    presentation.slides.move(0, 1)
    assert frame.text == "target"
    frame.text = "moved"
    assert presentation.slides[1].notes_text == "moved"
    assert paragraph.text == "kept" and shape.text_frame.text == "kept"
    presentation.slides.remove(slide)
    with pytest.raises(rpptx.StaleElementError):
        frame.text = "must not retarget"
    with pytest.raises(rpptx.StaleElementError):
        notes.notes_text_frame
    assert presentation.slides[0].notes_text == "other"


def test_fx239_valid_unit_controls_retain_pinned_structure():
    from rpptx.util import Length
    presentation, slide, shape, paragraph, font = _fx239_text_owner()
    paragraph.space_before = 0
    paragraph.space_after = 127
    font.spacing = -127
    paragraph.line_spacing = Length(127)
    with zipfile.ZipFile(io.BytesIO(presentation.to_bytes())) as package:
        from xml.etree import ElementTree as ET
        root = ET.fromstring(package.read("ppt/slides/slide1.xml"))
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    assert next(root.iter(a + "spcBef"))[0].get("val") == "0"
    assert next(root.iter(a + "spcAft"))[0].get("val") == "1"
    assert next(root.iter(a + "lnSpc"))[0].get("val") == "1"
    assert next(root.iter(a + "rPr")).get("spc") == "-1"
    for value in [Length(0), Length(126)]:
        before = presentation.to_bytes()
        with pytest.raises(ValueError, match="Pt|positive"):
            paragraph.line_spacing = value
        assert presentation.to_bytes() == before
    paragraph.line_spacing = 0.000005
    assert paragraph.line_spacing == 0.00001
    for name in ["space_before", "space_after"]:
        setattr(paragraph, name, None)
        assert getattr(paragraph, name) is None
    font.spacing = 0
    assert font.spacing == 0
    font.spacing = None
    assert font.spacing is None


def test_fx239_notes_names_match_pinned_python_pptx(tmp_path):
    import pptx
    import rpptx
    from xml.etree import ElementTree as ET
    assert importlib.metadata.version("python-pptx") == "1.0.2"
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    projections = []
    for index, module in enumerate([rpptx, pptx]):
        presentation = module.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        assert not slide.has_notes_slide
        frame = slide.notes_slide.notes_text_frame
        assert frame is not None and frame.text == "" and slide.has_notes_slide
        frame.text = "one\ntwo"
        path = tmp_path / f"notes-{index}.pptx"
        presentation.save(path)
        with zipfile.ZipFile(path) as package:
            notes_name = next(n for n in package.namelist() if "/notesSlides/notesSlide" in n and n.endswith(".xml"))
            root = ET.fromstring(package.read(notes_name))
            bodies = [shape for shape in root.iter(p + "sp") if any(ph.get("type") == "body" for ph in shape.iter(p + "ph"))]
            assert len(bodies) == 1
            text = ["".join(t.text or "" for t in paragraph.iter(a + "t")) for paragraph in bodies[0].iter(a + "p")]
            rel_name = notes_name.rsplit("/", 1)[0] + "/_rels/" + notes_name.rsplit("/", 1)[1] + ".rels"
            relationships = ET.fromstring(package.read(rel_name))
            projections.append((text, sorted(rel.get("Type").rsplit("/", 1)[1] for rel in relationships)))
    assert projections[0] == projections[1]


def test_fx239_valid_spacing_matches_pinned_python_pptx(tmp_path):
    import pptx
    import rpptx
    from xml.etree import ElementTree as ET
    projections = []
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for index, module in enumerate([rpptx, pptx]):
        presentation = module.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        from_module = __import__(module.__name__ + ".util", fromlist=["Pt", "Inches"])
        paragraph = slide.shapes.add_textbox(*(from_module.Inches(x) for x in [1, 1, 4, 1])).text_frame.paragraphs[0]
        paragraph.text = "spacing"
        paragraph.space_before = from_module.Pt(6)
        paragraph.space_after = from_module.Pt(12)
        paragraph.line_spacing = 1.5
        path = tmp_path / f"spacing-{index}.pptx"
        presentation.save(path)
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read("ppt/slides/slide1.xml"))
        projections.append(tuple((next(root.iter(a + name))[0].tag, next(root.iter(a + name))[0].get("val")) for name in ["spcBef", "spcAft", "lnSpc"]))
    assert projections[0] == projections[1]


def _fx240_material():
    # Authentic existing native test identity, not a production or trusted key.
    import base64
    from pathlib import Path
    source = (Path(__file__).resolve().parents[2] / "oxml-opc" / "src" / "signature.rs").read_text(encoding="utf-8")
    def constant(name):
        return base64.b64decode(re.search(r'const ' + name + r': &str = "([^"]+)";', source)[1], validate=True)
    return constant("TEST_PRIVATE_KEY_PKCS8_BASE64"), constant("TEST_CERTIFICATE_DER_BASE64")


def _fx240_zip_edit(source, part=None, old=None, new=None, extra=None):
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source)) as archive, zipfile.ZipFile(output, "w") as result:
        for info in archive.infolist():
            data = archive.read(info.filename)
            if info.filename == part:
                assert old in data
                data = data.replace(old, new, 1)
            result.writestr(info, data)
        if extra:
            result.writestr("word/uncovered.xml" if "rpptx" == "rdocx" else "ppt/uncovered.xml", b"<uncovered/>")
    return output.getvalue()


def test_fx240_sign_verify_owned_reports_pem_and_multisignatures():
    import base64
    import rpptx as product
    owner = product.Presentation()
    assert owner.verify_signatures() == ()
    opaque = b'<unmodelled xmlns="urn:caller-owned"><payload a="keep">opaque</payload></unmodelled>'
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(owner.to_bytes())) as source, zipfile.ZipFile(output, "w") as result:
        entries = [(item.filename, source.read(item.filename)) for item in source.infolist()]
        for name, data in entries:
            result.writestr(name, data)
        result.writestr("custom/opaque.xml", opaque)
    owner = product.Presentation.from_bytes(output.getvalue())
    key, cert = _fx240_material()
    def pem(label, value):
        encoded = base64.b64encode(value)
        return b" \r\n-----BEGIN " + label + b"-----\r\n" + b"\r\n".join(encoded[i:i+64] for i in range(0, len(encoded), 64)) + b"\r\n-----END " + label + b"-----\t\r\n"
    report = owner.sign(pem(b"PRIVATE KEY", key), pem(b"CERTIFICATE", cert))
    assert type(report) is dict and report["cryptographically_valid"] and report["coverage_complete"]
    assert set(report) == {"signature_part", "signer", "cryptographically_valid", "coverage_complete", "covered_parts", "covered_relationships", "issues"}
    assert set(report["signer"]) == {"subject", "issuer", "serial_number"}
    assert "F-172 Test Signer" in report["signer"]["subject"]
    assert report["issues"] == [] and report["covered_parts"] and report["covered_relationships"]
    assert set(report["covered_relationships"][0]) == {"source_part", "relationship_id"}
    report["covered_parts"].clear()
    report["signer"]["subject"] = "mutated"
    assert owner.verify_signatures()[0]["covered_parts"]
    assert owner.verify_signatures()[0]["signer"]["subject"] != "mutated"
    before = owner.to_bytes()
    with zipfile.ZipFile(io.BytesIO(before)) as preserved:
        assert preserved.read("custom/opaque.xml") == opaque
    with pytest.raises(product.PackageError):
        owner.sign(key, cert)  # The native provider refuses re-signing a signed graph.
    assert owner.to_bytes() == before
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(before)) as archive, zipfile.ZipFile(output, "w") as result:
        signature_bytes = archive.read("_xmlsignatures/sig1.xml")
        for info in archive.infolist():
            result.writestr(info, archive.read(info.filename))
        result.writestr("_xmlsignatures/sig2.xml", signature_bytes)
    reports = product.Presentation.from_bytes(output.getvalue()).verify_signatures()
    assert type(reports) is tuple and len(reports) == 2
    assert reports[1]["signer"] is None and not reports[1]["cryptographically_valid"]
    assert reports[1]["issues"] == [{"kind": "orphan_signature_part", "details": {"part_name": "/_xmlsignatures/sig2.xml"}}]
    assert not any(hasattr(product, name) for name in ("SignatureReport", "_SignatureReport", "_SignatureIssue"))


def test_fx240_current_tamper_partial_unsigned_and_malformed():
    import rpptx as product
    owner = product.Presentation()
    slide = owner.slides.add_slide(owner.slide_layouts[6]); slide.shapes.add_textbox(0, 0, 1000000, 1000000).text = "signed text"
    key, cert = _fx240_material()
    owner.sign(key, cert)
    original = owner.to_bytes()
    owner.slides[0].shapes[0].text = "changed text"
    reports = owner.verify_signatures()
    assert len(reports) == 1 and not reports[0]["cryptographically_valid"]
    issues = reports[0]["issues"]
    changed = [issue for issue in issues if issue["kind"] == "changed_reference"]
    assert any(issue["details"]["uri"].startswith("/ppt/slides/slide1.xml?") for issue in changed)
    assert all(set(issue["details"]) == {"uri", "expected_digest", "actual_digest"} for issue in changed)
    partial = product.Presentation.from_bytes(_fx240_zip_edit(original, extra=True))
    report = partial.verify_signatures()[0]
    assert report["cryptographically_valid"] and not report["coverage_complete"]
    assert {"kind": "uncovered_part", "details": {"part_name": "/ppt/uncovered.xml"}} in report["issues"]
    malformed = _fx240_zip_edit(original, "_xmlsignatures/sig1.xml", b"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256", b"urn:unsupported-signature")
    with pytest.raises(product.PackageError):
        product.Presentation.from_bytes(malformed).verify_signatures()
    assert product.Presentation().verify_signatures() == ()


def test_fx240_refusal_preserves_current_bytes_and_held_authority():
    import rpptx as product
    owner = product.Presentation()
    slide = owner.slides.add_slide(owner.slide_layouts[6]); slide.shapes.add_textbox(0, 0, 1000000, 1000000).text = "signed text"
    slide = owner.slides[0]; shape = slide.shapes[0]
    frame = shape.text_frame; paragraph = frame.paragraphs[0]; run = paragraph.runs[0]
    before = owner.to_bytes()
    key, cert = _fx240_material()
    for bad in (b"", b"x" * (1024 * 1024 + 1), b"-----BEGIN RSA PRIVATE KEY-----Zg==-----END RSA PRIVATE KEY-----", key + b" "):
        with pytest.raises(product.PackageError):
            owner.sign(bad, cert)
        assert owner.to_bytes() == before and owner.verify_signatures() == ()
        assert shape.text == frame.text == paragraph.text == run.text == "signed text" and len(slide.shapes) == 1
    for value in ("not bytes", 123, None):
        with pytest.raises(TypeError):
            owner.sign(value, cert)
    report = owner.sign(key, cert)
    assert report["cryptographically_valid"]
    assert shape.text == frame.text == paragraph.text == run.text == "signed text" and len(slide.shapes) == 1
    saved = owner.to_bytes()
    assert owner.verify_signatures()[0]["cryptographically_valid"]
    assert owner.to_bytes() == saved
    assert shape.text == frame.text == paragraph.text == run.text == "signed text" and len(slide.shapes) == 1


def test_fx238_slide_xml_hack_refuses_with_remove_pointer():
    import rpptx
    presentation = rpptx.Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    before = presentation.to_bytes()
    with pytest.raises(NotImplementedError, match="slides.remove"):
        presentation.slides._sldIdLst
    assert presentation.to_bytes() == before


@pytest.mark.parametrize("authored,expected,old,new", [
    (51, 58, b'<c:barDir val="col"/>', b'<c:barDir val="bar"/>'),
    (51, 52, b'<c:grouping val="clustered"/>', b'<c:grouping val="stacked"/>'),
    (65, 4, b'</c:ser>', b'<c:marker><c:symbol val="none"/></c:marker></c:ser>'),
    (5, 69, b'</c:ser>', b'<c:explosion val="0"/></c:ser>'),
])
def test_fx253_imported_chart_types_are_exact_after_reopen(authored, expected, old, new):
    from rpptx import Presentation, CategoryChartData, Inches, StaleElementError
    data = CategoryChartData()
    data.categories = ["A", "B"]
    data.add_series("Series", [3.0, 7.0])
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    frame = slide.shapes.add_chart(authored, Inches(1), Inches(1), Inches(3), Inches(2), data)
    parts = _package_parts(deck.to_bytes())
    part = next(name for name in parts if re.fullmatch(r"ppt/charts/chart[0-9]+.xml", name))
    assert old in parts[part]
    parts[part] = parts[part].replace(old, new)
    if expected == 58:
        parts[part] = parts[part].replace(b'val="clustered"', b'val="stacked"')
    source = _package_bytes(parts)
    imported = Presentation.from_bytes(source)
    chart = imported.slides[0].shapes[0].chart
    before = imported.to_bytes()
    assert chart.chart_type == expected
    assert chart.categories == ("A", "B") and chart.series[0].values == [3.0, 7.0]
    assert chart.chart_type == expected and imported.to_bytes() == before
    assert _package_parts(imported.to_bytes())[part] == parts[part]
    reopened = Presentation.from_bytes(imported.to_bytes())
    assert reopened.slides[0].shapes[0].chart.chart_type == expected
    grouped = imported.slides[0].shapes.group([imported.slides[0].shapes[0]])
    assert chart.chart_type == expected and grouped.shapes[0].chart.chart_type == expected
    imported.slides[0].shapes.remove(grouped)
    with pytest.raises(StaleElementError):
        _ = chart.chart_type


def test_fx253_unknown_chart_type_is_none_without_requiring_cached_data_or_workbook():
    from rpptx import Presentation, CategoryChartData, Inches
    data = CategoryChartData()
    data.categories = ["A"]
    data.add_series("Series", [3.0])
    deck = Presentation()
    deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_chart(51, Inches(1), Inches(1), Inches(3), Inches(2), data)
    parts = _package_parts(deck.to_bytes())
    part = next(name for name in parts if re.fullmatch(r"ppt/charts/chart[0-9]+.xml", name))
    shell = b'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>BODY</c:plotArea></c:chart></c:chartSpace>'
    for plot, expected in [(b'<c:bubbleChart/>', None), (b'<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/></c:barChart>', 58)]:
        parts[part] = shell.replace(b'BODY', plot)
        imported = Presentation.from_bytes(_package_bytes(parts))
        before = imported.to_bytes()
        assert imported.slides[0].shapes[0].chart.chart_type == expected
        assert imported.to_bytes() == before


@pytest.mark.parametrize("plot,scalar,valid", [("scatterChart", "scatterStyle", "smoothMarker"), ("radarChart", "radarStyle", "filled")])
def test_fx253_invalid_selected_styles_raise_while_valid_unsupported_styles_are_none(plot, scalar, valid):
    from rpptx import Presentation, CategoryChartData, Inches, RpptxError
    data = CategoryChartData()
    data.categories = ["A"]
    data.add_series("Series", [3.0])
    deck = Presentation()
    deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_chart(51, Inches(1), Inches(1), Inches(3), Inches(2), data)
    parts = _package_parts(deck.to_bytes())
    part = next(name for name in parts if re.fullmatch(r"ppt/charts/chart[0-9]+.xml", name))
    for value in [valid, "garbage"]:
        parts[part] = f'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:{plot}><c:{scalar} val="{value}"/></c:{plot}></c:plotArea></c:chart></c:chartSpace>'.encode()
        imported = Presentation.from_bytes(_package_bytes(parts))
        before = imported.to_bytes()
        if value == valid:
            assert imported.slides[0].shapes[0].chart.chart_type is None
        else:
            with pytest.raises(RpptxError, match="invalid.*style"):
                _ = imported.slides[0].shapes[0].chart.chart_type
        assert imported.to_bytes() == before


def test_fx258_slide_id_uses_registered_source_identity():
    import rpptx
    from xml.etree import ElementTree as ET
    presentation, *_ = _fx239_text_owner()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    parts = _package_parts(presentation.to_bytes())
    root = ET.fromstring(parts["ppt/presentation.xml"])
    ids = root.findall("{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst/{http://schemas.openxmlformats.org/presentationml/2006/main}sldId")
    for node, value in zip(ids, (700, 999)):
        node.set("id", str(value))
    parts["ppt/presentation.xml"] = ET.tostring(root)
    presentation = rpptx.Presentation.from_bytes(_package_bytes(parts))
    first, second = presentation.slides
    before = presentation.to_bytes()
    assert (first.slide_id, second.slide_id) == (700, 999)
    assert presentation.to_bytes() == before
    presentation.slides.move(0, 1)
    assert (first.slide_id, second.slide_id) == (700, 999)
    assert [slide.slide_id for slide in presentation.slides] == [999, 700]
    reopened = rpptx.Presentation.from_bytes(presentation.to_bytes())
    assert [slide.slide_id for slide in reopened.slides] == [999, 700]
    presentation.slides.remove(first)
    with pytest.raises(rpptx.StaleElementError, match="stale"):
        _ = first.slide_id
    assert second.slide_id == 999


def test_fx258_powerpoint_color_getters_match_pinned_enums():
    import rpptx
    *_, font = _fx239_text_owner()
    assert font.color.type is None
    with pytest.raises(AttributeError, match="theme_color"):
        _ = font.color.theme_color
    font.color.rgb = "123456"
    assert font.color.type is MSO_COLOR_TYPE.RGB
    assert font.color.theme_color is MSO_THEME_COLOR.NOT_THEME_COLOR
    font.color.theme_color = "accent2"
    assert font.color.type is MSO_COLOR_TYPE.SCHEME
    assert font.color.theme_color is MSO_THEME_COLOR.ACCENT_2
    font.color.theme_color = None
    assert font.color.type is None


@pytest.mark.parametrize("token", ["phClr", "unknown-slot"])
def test_fx258_unknown_theme_projection_preserves_source(token):
    import rpptx
    presentation, _, _, _, font = _fx239_text_owner()
    font.color.theme_color = "accent1"
    parts = _package_parts(presentation.to_bytes())
    parts["ppt/slides/slide1.xml"] = parts["ppt/slides/slide1.xml"].replace(b'val="accent1"', ('val="' + token + '"').encode())
    presentation = rpptx.Presentation.from_bytes(_package_bytes(parts))
    shape = presentation.slides[0].shapes[0]
    font = shape.text_frame.paragraphs[0].runs[0].font
    before = presentation.to_bytes()
    assert font.color.type is MSO_COLOR_TYPE.SCHEME
    with pytest.raises(NotImplementedError, match="element.xml"):
        _ = font.color.theme_color
    assert presentation.to_bytes() == before and token in shape.element.xml
    font.color.theme_color = "accent2"
    assert font.color.theme_color is MSO_THEME_COLOR.ACCENT_2


@pytest.mark.parametrize("surface", _F_X197_PPT_COLOUR_SURFACES)
def test_fx258_color_projection_preserves_all_existing_owners(surface):
    import rpptx
    presentation = _textbox_presentation(rpptx)
    setter, color = _f_x197_ppt_colour_surface(presentation, surface)
    setter("123456")
    before = presentation.to_bytes()
    assert color.type is MSO_COLOR_TYPE.RGB
    assert color.theme_color is MSO_THEME_COLOR.NOT_THEME_COLOR
    assert presentation.to_bytes() == before
    color.theme_color = "accent1"
    color.brightness = .25
    before = presentation.to_bytes()
    assert color.type is MSO_COLOR_TYPE.SCHEME
    assert color.theme_color is MSO_THEME_COLOR.ACCENT_1
    assert color.brightness == .25
    assert presentation.to_bytes() == before
    reopened = rpptx.Presentation.from_bytes(before)
    assert _package_parts(reopened.to_bytes()) == _package_parts(before)


@pytest.mark.parametrize("choice", ["gradient", "pattern"])
def test_fx258_gradient_and_pattern_color_projection(choice):
    presentation, _, shape, *_ = _fx239_text_owner()
    fill = shape.fill
    if choice == "gradient":
        fill.gradient()
        colors = [stop.color for stop in fill.gradient_stops]
    else:
        fill.patterned()
        colors = [fill.fore_color, fill.back_color]
    for color in colors:
        color.rgb = "123456"
        assert color.type is MSO_COLOR_TYPE.RGB
        assert color.theme_color is MSO_THEME_COLOR.NOT_THEME_COLOR
        color.theme_color = "accent2"
        before = presentation.to_bytes()
        assert color.theme_color is MSO_THEME_COLOR.ACCENT_2
        assert presentation.to_bytes() == before


def test_fx258_language_id_uses_finite_pinned_mapping():
    import pptx.enum.lang as oracle
    from rpptx.enum.lang import MSO_LANGUAGE_ID
    assert importlib.metadata.version("python-pptx") == "1.0.2"
    assert {name: int(member) for name, member in MSO_LANGUAGE_ID.__members__.items()} == {name: int(member) for name, member in oracle.MSO_LANGUAGE_ID.__members__.items()}
    assert len(MSO_LANGUAGE_ID.__members__) == 216
    presentation, _, shape, _, font = _fx239_text_owner()
    before = presentation.to_bytes()
    assert font.language_id is MSO_LANGUAGE_ID.NONE and font.language is None
    assert presentation.to_bytes() == before
    for name, expected in oracle.MSO_LANGUAGE_ID.__members__.items():
        if name == "MIXED":
            continue
        font.language_id = MSO_LANGUAGE_ID[name]
        if name == "NONE":
            assert font.language is None and font.language_id is MSO_LANGUAGE_ID.NONE
            continue
        assert font.language == expected.xml_value
        assert int(font.language_id) == int(oracle.MSO_LANGUAGE_ID.from_xml(expected.xml_value))
        assert shape.text_frame.text == "kept"
    font.language_id = MSO_LANGUAGE_ID.NONE
    assert font.language is None
    font.language_id = MSO_LANGUAGE_ID.FRENCH
    font.language_id = None
    assert font.language is None


@pytest.mark.parametrize("value,error", [(True, TypeError), ("en-US", TypeError), (1.5, TypeError), (object(), TypeError), (-2, ValueError), (999999, ValueError), (2**80, ValueError)])
def test_fx258_language_id_refuses_without_publishing(value, error):
    from rpptx.enum.lang import MSO_LANGUAGE_ID
    presentation, slide, shape, _, font = _fx239_text_owner()
    font.language = "fr-FR"
    before = presentation.to_bytes()
    with pytest.raises(error, match="Font.language_id"):
        font.language_id = value
    assert presentation.to_bytes() == before
    assert font.language == "fr-FR" and font.language_id is MSO_LANGUAGE_ID.FRENCH
    assert shape.text_frame.text == "kept" and slide.slide_id >= 256
    font.language_id = MSO_LANGUAGE_ID.ENGLISH_US
    assert font.language == "en-US"


def test_fx258_unknown_language_id_keeps_existing_string_route():
    from rpptx.enum.lang import MSO_LANGUAGE_ID
    presentation, _, shape, _, font = _fx239_text_owner()
    font.language = "zz-Unknown"
    before = presentation.to_bytes()
    with pytest.raises(NotImplementedError, match="Font.language"):
        _ = font.language_id
    assert font.language == "zz-Unknown" and presentation.to_bytes() == before
    assert shape.text_frame.text == "kept"
    font.language_id = MSO_LANGUAGE_ID.FRENCH
    assert font.language == "fr-FR"


@pytest.mark.parametrize("name", ["_element", "_txBody"])
def test_fx258_private_powerpoint_names_point_to_checked_element(name):
    import rpptx
    presentation, slide, shape, *_ = _fx239_text_owner()
    frame = shape.text_frame
    target = shape if name == "_element" else frame
    before = presentation.to_bytes()
    with pytest.raises(NotImplementedError, match="element.*xml.*replace_xml"):
        getattr(target, name)
    assert presentation.to_bytes() == before and frame.text == "kept"
    element = target.element
    element.replace_xml(element.xml.replace("kept", "edited"))
    assert presentation.slides[0].shapes[0].text_frame.text == "edited"
    presentation.slides.remove(presentation.slides[0])
    with pytest.raises(rpptx.StaleElementError, match="stale"):
        getattr(target, name)


@pytest.mark.parametrize("tag", ["en-US", "fr-FR"])
def test_fx258_language_saved_packages_agree_with_pinned_python_pptx(tag):
    import pptx
    import rpptx
    from pptx.enum.lang import MSO_LANGUAGE_ID as OracleLanguage
    from rpptx.enum.lang import MSO_LANGUAGE_ID
    assert importlib.metadata.version("python-pptx") == "1.0.2"
    presentation, _, _, _, font = _fx239_text_owner()
    expected = OracleLanguage.from_xml(tag)
    font.language_id = MSO_LANGUAGE_ID(int(expected))
    saved = presentation.to_bytes()
    oracle = pptx.Presentation(io.BytesIO(saved))
    oracle_font = oracle.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
    assert oracle_font.language_id is expected
    oracle_font.language_id = OracleLanguage.FRENCH if tag == "en-US" else OracleLanguage.ENGLISH_US
    data = io.BytesIO(); oracle.save(data)
    reopened = rpptx.Presentation.from_bytes(data.getvalue())
    actual = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font
    assert int(actual.language_id) == int(oracle_font.language_id)
    assert actual.language == oracle_font.language_id.xml_value


@pytest.mark.parametrize("owner", ["master-style", "master-shape", "layout-shape", "cell-default"])
def test_fx258_language_and_color_keep_master_layout_and_cell_owners(owner):
    import rpptx
    from rpptx.enum.lang import MSO_LANGUAGE_ID
    presentation, slide, *_ = _fx239_text_owner()
    if owner == "master-style":
        font = presentation.slide_master.text_styles.body.level(1).font
    else:
        if owner == "master-shape":
            shapes = presentation.slide_master.shapes
        elif owner == "layout-shape":
            shapes = presentation.slide_layouts[6].shapes
        else:
            shapes = slide.shapes
        if owner == "cell-default":
            frame = shapes.add_table(1, 1, 0, 0, 100, 100).table.cell(0, 0).text_frame
        else:
            frame = shapes.add_textbox(0, 0, 100, 100).text_frame
        font = frame.paragraphs[0].font
    font.language_id = MSO_LANGUAGE_ID.FRENCH
    font.color.theme_color = "accent2"
    before = presentation.to_bytes()
    assert font.language_id is MSO_LANGUAGE_ID.FRENCH
    assert font.language == "fr-FR"
    assert font.color.type is MSO_COLOR_TYPE.SCHEME
    assert font.color.theme_color is MSO_THEME_COLOR.ACCENT_2
    assert presentation.to_bytes() == before
    reopened = rpptx.Presentation.from_bytes(before)
    assert _package_parts(reopened.to_bytes()) == _package_parts(before)


def test_fx258_removed_text_body_keeps_causal_owner_error():
    import rpptx
    presentation, _, shape, *_ = _fx239_text_owner()
    frame = shape.text_frame
    element = shape.element
    xml, count = re.subn(r"<p:txBody\b[^>]*>.*?</p:txBody>", "", element.xml, flags=re.S)
    assert count == 1
    element.replace_xml(xml)
    before = presentation.to_bytes()
    with pytest.raises(rpptx.StaleElementError, match="stale"):
        _ = frame._txBody
    assert presentation.to_bytes() == before


def test_fx258_color_saved_packages_agree_with_pinned_python_pptx():
    import pptx
    import rpptx
    from pptx.enum.dml import MSO_THEME_COLOR as OracleTheme
    assert importlib.metadata.version("python-pptx") == "1.0.2"
    presentation, _, _, _, font = _fx239_text_owner()
    for code in range(1, 17):
        font.color.theme_color = MSO_THEME_COLOR(code)
        font.color.brightness = .25
        oracle = pptx.Presentation(io.BytesIO(presentation.to_bytes()))
        color = oracle.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color
        assert int(color.type) == int(font.color.type)
        assert int(color.theme_color) == code
        assert color.brightness == font.color.brightness
        color.theme_color = OracleTheme(code)
        data = io.BytesIO(); oracle.save(data)
        reopened = rpptx.Presentation.from_bytes(data.getvalue())
        actual = reopened.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color
        assert actual.theme_color is MSO_THEME_COLOR(code)
        assert actual.type is MSO_COLOR_TYPE.SCHEME
        assert actual.brightness == .25
    font.color.rgb = "123456"
    oracle = pptx.Presentation(io.BytesIO(presentation.to_bytes()))
    color = oracle.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color
    assert int(color.type) == int(font.color.type)
    assert int(color.theme_color) == int(font.color.theme_color) == 0

# F-X254 local regressions. Complete installed/oracle suites belong to sprint final.
@pytest.mark.parametrize("suffix", [".pdf", ".PDF"])
def test_fx254_python_save_refusal_is_atomic(tmp_path, suffix):
    from rpptx import Presentation

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    before = prs.to_bytes()
    output = tmp_path / ("wrong" + suffix)
    with pytest.raises(ValueError, match="to_pdf"):
        prs.save(output)
    assert not output.exists()
    assert prs.to_bytes() == before


def _slide_xml(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read("ppt/slides/slide1.xml").decode()


def _rewrite_placeholder_gap(xml):
    shapes = list(re.finditer(r"<p:sp\b.*?</p:sp>", xml, re.DOTALL))
    placeholders = [shape for shape in shapes if "<p:ph" in shape.group(0)]
    if len(placeholders) < 2:
        return xml
    first, second = placeholders[0], placeholders[1]
    first_xml = first.group(0).replace('idx="1"', 'idx="2"', 1)
    second_xml = second.group(0).replace('idx="1"', 'idx="2"', 1)
    return (
        xml[: first.start()]
        + second_xml
        + xml[first.end() : second.start()]
        + first_xml
        + xml[second.end() :]
    )


def _part_from_relationship(owner_part, target):
    parts = owner_part.rsplit("/", 1)[0].split("/")
    for piece in target.split("/"):
        if piece == "..":
            parts.pop()
        elif piece not in ("", "."):
            parts.append(piece)
    return "/".join(parts)


def _relationship_target(rels_xml, type_suffix):
    for relationship in re.finditer(r"<Relationship\b([^>]*)/?>", rels_xml):
        attributes = dict(re.findall(r'([A-Za-z]+)="([^"]*)"', relationship.group(1)))
        if attributes.get("Type", "").endswith(type_suffix):
            return attributes["Target"]
    raise AssertionError(f"missing relationship {type_suffix}")


def test_issue392_placeholder_iteration_crosses_index_gaps(tmp_path):
    import rpptx

    presentation = rpptx.Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[0])
    source = tmp_path / "gaps-source.pptx"
    presentation.save(source)
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    slide_rels = parts["ppt/slides/_rels/slide1.xml.rels"].decode()
    layout_part = _part_from_relationship(
        "ppt/slides/slide1.xml",
        _relationship_target(slide_rels, "/slideLayout"),
    )
    layout_rels = parts[layout_part.rsplit("/", 1)[0] + "/_rels/" + layout_part.rsplit("/", 1)[1] + ".rels"].decode()
    master_part = _part_from_relationship(
        layout_part,
        _relationship_target(layout_rels, "/slideMaster"),
    )
    for part_name in ("ppt/slides/slide1.xml", layout_part, master_part):
        parts[part_name] = _rewrite_placeholder_gap(parts[part_name].decode()).encode()
    gapped = tmp_path / "gaps.pptx"
    with zipfile.ZipFile(gapped, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)

    opened = rpptx.Presentation(gapped)
    slide = opened.slides[0]

    def tree_indexes(owner):
        return [shape.placeholder_format.idx for shape in owner.shapes if shape.is_placeholder]

    def iterated_indexes(owner):
        return [shape.placeholder_format.idx for shape in owner.placeholders]

    slide_tree = tree_indexes(slide)
    slide_order = iterated_indexes(slide)
    assert 0 in slide_tree and 2 in slide_tree and 1 not in slide_tree
    assert slide_order == sorted(slide_tree)
    assert slide_order != slide_tree
    assert len(slide.placeholders) == len(slide_tree)
    assert slide.placeholders[0].placeholder_format.idx == 0
    assert slide.placeholders[2].placeholder_format.idx == 2
    with pytest.raises(IndexError):
        _ = slide.placeholders[1]

    layout = slide.slide_layout
    master = layout.slide_master
    for owner in (layout, master):
        owner_tree = tree_indexes(owner)
        assert iterated_indexes(owner) == owner_tree
        assert owner_tree != sorted(owner_tree)
        assert len(owner.placeholders) == len(owner_tree)
        present = set(owner_tree)
        assert 0 in present
        assert owner.placeholders[0].placeholder_format.idx == 0
        for index in present:
            assert owner.placeholders[index].placeholder_format.idx == index
        if 1 not in present:
            with pytest.raises(IndexError):
                _ = owner.placeholders[1]
    notes = slide.notes_slide
    assert not hasattr(notes, "placeholders")


def test_issue392_percentage_rounding_has_explicit_boundaries():
    import rpptx
    from rpptx.dml.color import RGBColor

    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    expected_alpha = {
        0.8: "20000",
        0.998085: "192",
        0.994255: "574",
        0.57: "43000",
    }
    for value in expected_alpha:
        shape = slide.shapes.add_shape(1, 0, 0, 100, 100)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(1, 2, 3)
        shape.fill.transparency = value
    for value in (0.000015, 0.000025):
        shape = slide.shapes.add_shape(1, 0, 0, 100, 100)
        shape.fill.gradient()
        shape.fill.gradient_stops[0].position = value
    box = slide.shapes.add_textbox(0, 0, 1000000, 500000)
    box.text_frame.text = "tie"
    paragraph = box.text_frame.paragraphs[0]
    paragraph.bullet_size = 0.250015
    second = paragraph.add_run(" next")
    second_paragraph = box.text_frame.add_paragraph()
    second_paragraph.bullet_size = 0.250025
    paragraph.runs[0].font.baseline = -0.000015
    second.font.baseline = -0.000025
    xml = _slide_xml(presentation.to_bytes())
    for alpha in expected_alpha.values():
        assert f'<a:alpha val="{alpha}"/>' in xml
    assert '<a:alpha val="575"/>' not in xml
    assert '<a:alpha val="19999"/>' not in xml
    assert xml.count('pos="2"') == 2
    assert 'pos="3"' not in xml
    assert xml.count('<a:buSzPct val="25002"/>') == 2
    assert '<a:buSzPct val="25003"/>' not in xml
    assert xml.count('baseline="-2"') == 2
    assert 'baseline="-3"' not in xml

    before = presentation.to_bytes()
    with pytest.raises(TypeError, match="rejects bool"):
        slide.shapes[0].fill.transparency = True
    with pytest.raises(TypeError, match="rejects bool"):
        slide.shapes[4].fill.gradient_stops[0].position = False
    with pytest.raises(TypeError, match="rejects bool"):
        paragraph.bullet_size = True
    with pytest.raises(TypeError, match="rejects bool"):
        paragraph.runs[0].font.baseline = False
    with pytest.raises(ValueError):
        paragraph.runs[0].font.baseline = float("nan")
    assert presentation.to_bytes() == before


def test_issue392_up_arrow_has_reproducible_geometry(tmp_path):
    import rpptx
    from rpptx.enum.shapes import MSO_SHAPE

    assert (int(MSO_SHAPE.UP_ARROW), MSO_SHAPE.UP_ARROW.xml_value) == (35, "upArrow")
    presentation = rpptx.Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = presentation.slides[0].shapes.add_shape(MSO_SHAPE.UP_ARROW, 0, 0, 100, 100)
    assert shape.auto_shape_type is MSO_SHAPE.UP_ARROW
    assert list(shape.adjustments) == [0.5, 0.5]
    output = tmp_path / "up-arrow.pptx"
    presentation.save(output)
    reopened = rpptx.Presentation(output)
    imported = reopened.slides[0].shapes[0]
    assert imported.auto_shape_type is MSO_SHAPE.UP_ARROW
    assert list(imported.adjustments) == [0.5, 0.5]
    named = reopened.slides[0].shapes.add_shape("upArrow", 0, 0, 10, 10)
    assert named.auto_shape_type is MSO_SHAPE.UP_ARROW
    assert 'prst="upArrow"' in _slide_xml(reopened.to_bytes())


def test_issue392_text_frame_clear_preserves_approved_properties():
    import rpptx

    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    box = slide.shapes.add_textbox(0, 0, 1000000, 400000)
    paragraph = box.text_frame.paragraphs[0]
    paragraph.alignment = rpptx.PP_ALIGN.CENTER
    paragraph.add_run("one")
    held = paragraph.add_run("two")
    other = slide.shapes.add_textbox(0, 500000, 1000000, 400000)
    other.text = "stay"
    other_run = other.text_frame.paragraphs[0].runs[0]
    table = slide.shapes.add_table(1, 1, 0, 800000, 1000000, 400000)
    cell = table.table.cell(0, 0).text_frame
    cell.paragraphs[0].alignment = rpptx.PP_ALIGN.RIGHT
    cell.paragraphs[0].add_run("cell-one")
    cell_run = cell.paragraphs[0].add_run("cell-two")

    box.text_frame.clear()
    cell.clear()
    assert box.text_frame.paragraphs[0].alignment is rpptx.PP_ALIGN.CENTER
    assert cell.paragraphs[0].alignment is rpptx.PP_ALIGN.RIGHT
    assert (len(box.text_frame.paragraphs), len(box.text_frame.paragraphs[0].runs)) == (1, 0)
    assert (len(cell.paragraphs), len(cell.paragraphs[0].runs)) == (1, 0)
    assert box.text_frame.text == ""
    assert cell.text == ""
    with pytest.raises(rpptx.StaleElementError, match="TextFrame.clear"):
        _ = held.text
    with pytest.raises(rpptx.StaleElementError, match="TextFrame.clear"):
        _ = cell_run.text
    assert other_run.text == "stay"
    xml = _slide_xml(presentation.to_bytes())
    assert 'algn="ctr"' in xml
    assert 'algn="r"' in xml


def test_issue392_empty_text_then_append_has_one_new_run():
    import rpptx

    presentation = rpptx.Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    frame = presentation.slides[0].shapes.add_textbox(0, 0, 1000000, 400000).text_frame
    frame.text = "hello"
    frame.text = ""
    assert len(frame.paragraphs[0].runs) == 0
    added = frame.paragraphs[0].add_run("next")
    assert [run.text for run in frame.paragraphs[0].runs] == ["next"]
    assert added.text == "next"
    xml = _slide_xml(presentation.to_bytes())
    assert "<a:t>next</a:t>" in xml
    assert "<a:t/>" not in xml
    assert "<a:t></a:t>" not in xml


def test_issue392_cell_run_hyperlinks_use_the_actual_slide_scope(tmp_path):
    import rpptx

    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    table = slide.shapes.add_table(1, 1, 0, 0, 1000000, 400000)
    cell_run = table.table.cell(0, 0).text_frame.paragraphs[0].add_run("cell")
    box = slide.shapes.add_textbox(0, 500000, 1000000, 400000)
    box_run = box.text_frame.paragraphs[0].add_run("box")
    before = presentation.to_bytes()
    with pytest.raises(rpptx.RpptxError, match="control characters"):
        cell_run.hyperlink.address = "https://example.com/\nnext"
    with pytest.raises(ValueError, match="cell hyperlink assignment is not supported"):
        cell_run.hyperlink.target_slide = slide
    assert presentation.to_bytes() == before
    cell_run.hyperlink.address = "https://example.com/cell"
    cell_run.hyperlink.address = "https://example.com/changed"
    box_run.hyperlink.address = "https://example.com/box"
    assert cell_run.hyperlink.address == "https://example.com/changed"
    assert box_run.hyperlink.address == "https://example.com/box"
    assert _slide_hyperlink_targets(presentation.to_bytes()) == [
        "https://example.com/box",
        "https://example.com/changed",
    ]
    cell_run.hyperlink.address = None
    assert cell_run.hyperlink.address is None
    assert _slide_hyperlink_targets(presentation.to_bytes()) == ["https://example.com/box"]
    cell_run.hyperlink.address = "https://example.com/kept"
    output = tmp_path / "cell-link.pptx"
    presentation.save(output)
    reopened = rpptx.Presentation(output)
    reopened_run = reopened.slides[0].shapes[0].table.cell(0, 0).text_frame.paragraphs[0].runs[0]
    assert reopened_run.hyperlink.address == "https://example.com/kept"


def test_issue392_bullet_size_distinguishes_points_from_percent():
    import rpptx
    from rpptx.util import Pt

    presentation = rpptx.Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    frame = presentation.slides[0].shapes.add_textbox(0, 0, 1000000, 400000).text_frame
    points = frame.paragraphs[0]
    percent = frame.add_paragraph()
    cleared = frame.add_paragraph()
    refused = frame.add_paragraph()
    points.bullet_size = Pt(17.85)
    percent.bullet_size = 0.29
    refused.bullet_size = 1.0
    before = presentation.to_bytes()
    with pytest.raises(ValueError, match="0.25 to 4.0"):
        refused.bullet_size = 0.1
    with pytest.raises(TypeError, match="rejects bool"):
        refused.bullet_size = False
    assert presentation.to_bytes() == before
    cleared.bullet_size = 0.5
    cleared.bullet_size = None
    assert isinstance(points.bullet_size, rpptx.Length)
    assert int(points.bullet_size) == 1785 * 127
    assert percent.bullet_size == 0.29
    assert cleared.bullet_size is None
    xml = _slide_xml(presentation.to_bytes())
    assert '<a:buSzPts val="1785"/>' in xml
    assert '<a:buSzPct val="29000"/>' in xml
    assert xml.count("<a:buSz") == 3


def test_issue392_style_removal_is_kind_checked():
    import rpptx
    from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE

    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, 100, 100)
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 10, 10)
    plain = slide.shapes.add_textbox(0, 0, 100, 100)
    picture = slide.shapes.add_picture(io.BytesIO(_tiny_png()), 0, 0)
    group = slide.shapes.add_group_shape()
    table = slide.shapes.add_table(1, 1, 0, 0, 100, 100)
    assert _slide_xml(presentation.to_bytes()).count("<p:style") == 2
    assert shape.remove_style() is None
    assert connector.remove_style() is None
    assert _slide_xml(presentation.to_bytes()).count("<p:style") == 0
    unchanged = presentation.to_bytes()
    assert plain.remove_style() is None
    assert shape.remove_style() is None
    assert connector.remove_style() is None
    assert presentation.to_bytes() == unchanged
    for unsupported in (picture, group, table):
        before = presentation.to_bytes()
        with pytest.raises(rpptx.RpptxError, match="remove_style is not supported"):
            unsupported.remove_style()
        assert presentation.to_bytes() == before


@pytest.mark.parametrize("axis", ["slide_width", "slide_height"])
@pytest.mark.parametrize("value", [914399, 51206401])
def test_fx254_python_dimension_refusal_is_atomic(axis, value):
    from rpptx import Presentation

    prs = Presentation()
    before = prs.to_bytes()
    with pytest.raises(ValueError, match="1.*56"):
        setattr(prs, axis, value)
    assert prs.to_bytes() == before


@pytest.mark.parametrize("axis", ["slide_width", "slide_height"])
@pytest.mark.parametrize("value", [914400, 51206400])
def test_fx254_python_dimension_endpoints_remain_valid(axis, value):
    from rpptx import Presentation

    prs = Presentation()
    setattr(prs, axis, value)
    assert getattr(prs, axis) == value


@pytest.mark.parametrize("axis", ["flip_h", "flip_v"])
@pytest.mark.parametrize("kind", ["table", "chart"])
def test_fx254_frame_flip_axes_and_payloads(axis, kind):
    from rpptx import CategoryChartData, Presentation, XL_CHART_TYPE
    from rpptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if kind == "table":
        shape = slide.shapes.add_table(1, 1, 0, 0, Inches(2), Inches(1))
    else:
        data = CategoryChartData()
        data.categories = ["A", "B"]
        data.add_series("Values", [1, 2])
        shape = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 0, Inches(2), Inches(1), data
        )
    before = prs.to_bytes()
    setattr(shape, axis, False)
    assert prs.to_bytes() == before
    with pytest.raises(ValueError, match="flip"):
        setattr(shape, axis, True)
    assert prs.to_bytes() == before
    assert getattr(shape, axis) is False


def test_fx254_style_availability_refuses_unknown_atomically():
    from rpptx import Presentation, RpptxError
    from rpptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_table(1, 1, 0, 0, Inches(2), Inches(1))
    before = prs.to_bytes()
    with pytest.raises((ValueError, RpptxError), match="style"):
        shape.table.style_id = "{00000000-0000-0000-0000-000000000000}"
    assert prs.to_bytes() == before
    assert len(shape.table.rows) == 1

@pytest.mark.parametrize("case", ["shape", "geometry", "paragraph", "slide"])
def test_fx254_four_xml_frontiers_refuse_before_publish(case):
    from rpptx import MSO_SHAPE, Presentation, RpptxError
    from rpptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(2), Inches(1))
    shape.text = "source"
    owner = shape.text_frame.element if case == "paragraph" else slide.element if case == "slide" else shape.element
    old, new = {
        "shape": ("</p:sp>", "<a:p/></p:sp>"),
        "geometry": ("<a:avLst/>", "<a:avLst/><a:bogus/>"),
        "paragraph": ("<a:r>", "<a:bogus/><a:r>"),
        "slide": ("<p:clrMapOvr>", "<p:bogus/><p:clrMapOvr>"),
    }[case]
    xml = owner.xml
    assert old in xml
    before = prs.to_bytes()
    with pytest.raises((ValueError, RpptxError)):
        owner.replace_xml(xml.replace(old, new, 1))
    assert prs.to_bytes() == before
    assert shape.text == "source"
    assert owner.xml == xml


@pytest.mark.parametrize("case", ["shape", "geometry", "paragraph", "slide"])
def test_fx254_import_validation_diagnoses_four_frontiers(case):
    from rpptx import MSO_SHAPE, Presentation
    from rpptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(2), Inches(1))
    shape.text = "source"
    old, new = {
        "shape": (b"</p:sp>", b"<a:p/></p:sp>"),
        "geometry": (b"<a:avLst/>", b"<a:avLst/><a:bogus/>"),
        "paragraph": (b"<a:r>", b"<a:bogus/><a:r>"),
        "slide": (b"<p:clrMapOvr>", b"<p:bogus/><p:clrMapOvr>"),
    }[case]
    target = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(prs.to_bytes())) as source, zipfile.ZipFile(target, "w") as dest:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "ppt/slides/slide1.xml":
                assert old in data
                data = data.replace(old, new, 1)
            dest.writestr(info, data)
    imported = Presentation.from_bytes(target.getvalue())
    before = imported.to_bytes()
    issues = [issue for issue in imported.validate() if issue.kind == "invalid_xml_child"]
    assert issues and all("/ppt/slides/slide1.xml" in issue.message for issue in issues)
    assert imported.to_bytes() == before
    imported.slides[0].shapes.add_textbox(0, 0, 1000000, 500000).text = "unrelated"
    assert new in _package_parts(imported.to_bytes())["ppt/slides/slide1.xml"]
    assert any(issue.kind == "invalid_xml_child" for issue in imported.validate())
    owner = imported.slides[0].shapes[0].element
    existing = owner.xml
    current = imported.to_bytes()
    with pytest.raises((ValueError, __import__("rpptx").RpptxError)):
        owner.replace_xml(existing.replace("</p:sp>", "<a:bogus/></p:sp>", 1))
    assert imported.to_bytes() == current
    assert owner.xml == existing


def test_fx254_imported_size_noop_and_other_axis_correction():
    import rpptx

    parts = _package_parts(rpptx.Presentation().to_bytes())
    xml = parts["ppt/presentation.xml"]
    xml, count = re.subn(rb'(<p:sldSz\b[^>]*\bcx=")[0-9]+', rb'\g<1>60000000', xml, count=1)
    assert count == 1
    parts["ppt/presentation.xml"] = xml
    prs = rpptx.Presentation.from_bytes(_package_bytes(parts))
    before = prs.to_bytes()
    prs.slide_width = 60000000
    assert prs.to_bytes() == before
    prs.slide_height = 914400
    assert prs.slide_width == 60000000 and prs.slide_height == 914400


@pytest.mark.parametrize("style", ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}", "{D7AC3CCA-C797-4891-BE02-D94E43425B78}"])
def test_fx254_builtin_styles_and_clear(style):
    import rpptx

    prs = rpptx.Presentation()
    shape = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_table(1, 1, 0, 0, 1000000, 500000)
    shape.table.style_id = style.lower()
    assert shape.table.style_id == style.lower()
    before = prs.to_bytes()
    shape.table.style_id = style.lower()
    assert prs.to_bytes() == before
    shape.table.style_id = None
    assert shape.table.style_id is None


@pytest.mark.parametrize("mode", ["valid", "duplicate", "foreign", "wrong-content-type"])
def test_fx254_custom_style_uses_actual_unique_deck_definition(mode):
    import xml.etree.ElementTree as ET
    import rpptx

    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    ct = "http://schemas.openxmlformats.org/package/2006/content-types"
    kind = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
    style = "{12345678-1234-1234-1234-123456789ABC}"
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_table(1, 1, 0, 0, 1000000, 500000)
    parts = _package_parts(prs.to_bytes())
    relationships = ET.fromstring(parts["ppt/_rels/presentation.xml.rels"])
    for row in list(relationships):
        if row.get("Type") == kind:
            relationships.remove(row)
    ET.SubElement(relationships, "{" + rel + "}Relationship", {"Id": "rIdFx254Style", "Type": kind, "Target": "tableStyles.xml"})
    parts["ppt/_rels/presentation.xml.rels"] = ET.tostring(relationships)
    types = ET.fromstring(parts["[Content_Types].xml"])
    for row in list(types):
        if row.get("PartName") == "/ppt/tableStyles.xml":
            types.remove(row)
    ET.SubElement(types, "{" + ct + "}Override", {"PartName": "/ppt/tableStyles.xml", "ContentType": "application/xml" if mode == "wrong-content-type" else "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"})
    parts["[Content_Types].xml"] = ET.tostring(types)
    row = f'<a:tblStyle styleId="{style}" styleName="Custom"><a:wholeTbl/></a:tblStyle>'
    if mode == "duplicate":
        row += row.replace(style, style.lower())
    if mode == "foreign":
        row = row.replace("<a:tblStyle ", '<x:tblStyle xmlns:x="urn:foreign" ').replace("</a:tblStyle>", "</x:tblStyle>")
    parts["ppt/tableStyles.xml"] = f'<a:tblStyleLst xmlns:a="{a}" def="{style}">{row}</a:tblStyleLst>'.encode()
    prs = rpptx.Presentation.from_bytes(_package_bytes(parts))
    table = prs.slides[0].shapes[0].table
    before = prs.to_bytes()
    if mode == "valid":
        table.style_id = style.lower()
        assert table.style_id == style.lower()
        assert rpptx.Presentation.from_bytes(prs.to_bytes()).slides[0].shapes[0].table.style_id == style.lower()
    else:
        with pytest.raises((ValueError, rpptx.RpptxError), match="style"):
            table.style_id = style
        assert prs.to_bytes() == before


@pytest.mark.parametrize("payload", ['<z:opaque xmlns:z="urn:foreign"/>', '<z:avLst xmlns:z="urn:foreign"/>'])
def test_fx254_raw_foreign_geometry_cannot_disappear_in_projection(payload):
    import rpptx

    prs = rpptx.Presentation()
    shape = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_shape(rpptx.MSO_SHAPE.RECTANGLE, 0, 0, 1000000, 500000)
    owner = shape.element
    before = prs.to_bytes()
    xml = owner.xml
    assert "<a:avLst/>" in xml
    with pytest.raises((ValueError, rpptx.RpptxError)):
        owner.replace_xml(xml.replace("<a:avLst/>", payload, 1))
    assert prs.to_bytes() == before


@pytest.mark.parametrize("scope", ["root", "group"])
def test_fx254_valid_ignored_extension_uses_physical_ancestor(scope):
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shapes = slide.shapes if scope == "root" else slide.shapes.add_group_shape().shapes
    shapes.add_shape(rpptx.MSO_SHAPE.RECTANGLE, 0, 0, 1000000, 500000).text = "before"
    parts = _package_parts(prs.to_bytes())
    source = parts["ppt/slides/slide1.xml"]
    insertion = b'<p:sld ' if scope == "root" else b'<p:grpSp>'
    replacement = b'<p:sld xmlns:z="urn:fx254" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="z" ' if scope == "root" else b'<p:grpSp xmlns:z="urn:fx254" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="z">'
    assert insertion in source
    parts["ppt/slides/slide1.xml"] = source.replace(insertion, replacement, 1)
    prs = rpptx.Presentation.from_bytes(_package_bytes(parts))
    shape = prs.slides[0].shapes[0] if scope == "root" else prs.slides[0].shapes[0].shapes[0]
    xml = shape.element.xml
    shape.element.replace_xml(xml.replace("</p:sp>", '<z:opaque xmlns:z="urn:fx254"/></p:sp>', 1))
    assert b'<z:opaque xmlns:z="urn:fx254"/>' in _package_parts(prs.to_bytes())["ppt/slides/slide1.xml"]
    shape = prs.slides[0].shapes[0] if scope == "root" else prs.slides[0].shapes[0].shapes[0]
    shape.text = "after"
    assert b'<z:opaque xmlns:z="urn:fx254"/>' in _package_parts(prs.to_bytes())["ppt/slides/slide1.xml"]


@pytest.mark.parametrize("case", ["blank", "additive", "rebound", "understood", "wrong-attribute", "once", "literal"])
def test_fx254_mc_authority_is_uri_based_and_qualified(case):
    import rpptx

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_shape(rpptx.MSO_SHAPE.RECTANGLE, 0, 0, 1000000, 500000)
    parts = _package_parts(prs.to_bytes())
    uri = "urn:fx&#50;54" if case == "once" else "urn:fx&amp;#50;54" if case == "literal" else "urn:fx254"
    grant = 'Ignorable="z"' if case == "wrong-attribute" else 'mc:Ignorable="a"' if case == "understood" else 'mc:Ignorable="z"'
    authority = f'<p:sld xmlns:z="{uri}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" {grant} '
    parts["ppt/slides/slide1.xml"] = parts["ppt/slides/slide1.xml"].replace(b"<p:sld ", authority.encode(), 1)
    prs = rpptx.Presentation.from_bytes(_package_bytes(parts))
    element = prs.slides[0].shapes[0].element
    xml = element.xml
    child = '<z:opaque xmlns:z="urn:other"/>' if case == "rebound" else '<a:bogus/>' if case == "understood" else '<z:opaque xmlns:z="urn:fx254"/>'
    changed = xml.replace("</p:sp>", child + "</p:sp>", 1)
    if case == "blank":
        changed = changed.replace("<p:sp ", '<p:sp mc:Ignorable="" ', 1)
    elif case == "additive":
        changed = changed.replace("<p:sp ", '<p:sp xmlns:q="urn:additional" mc:Ignorable="q" ', 1)
        changed = changed.replace("</p:sp>", '<q:opaque/></p:sp>', 1)
    before = prs.to_bytes()
    if case in ("blank", "additive", "once"):
        element.replace_xml(changed)
        saved = prs.to_bytes()
        assert child.encode() in _package_parts(saved)["ppt/slides/slide1.xml"]
        assert rpptx.Presentation.from_bytes(saved).to_bytes() == saved
    else:
        with pytest.raises((ValueError, rpptx.RpptxError)):
            element.replace_xml(changed)
        assert prs.to_bytes() == before
        assert element.xml == xml


@pytest.mark.parametrize("change", ["fresh-id", "swap-ids", "group-id", "same-owner-text"])
def test_fx254_whole_part_exception_keeps_physical_shape_identity(change):
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shapes = slide.shapes.add_group_shape().shapes if change == "group-id" else slide.shapes
    first = shapes.add_shape(rpptx.MSO_SHAPE.RECTANGLE, 0, 0, 1000000, 500000)
    first.text = "before"
    second = shapes.add_shape(rpptx.MSO_SHAPE.RECTANGLE, 0, 0, 1000000, 500000)
    first_id, second_id = first.shape_id, second.shape_id
    group_id = slide.shapes[0].shape_id
    parts = _package_parts(prs.to_bytes())
    parts["ppt/slides/slide1.xml"] = parts["ppt/slides/slide1.xml"].replace(b"</p:sp>", b'<z:opaque xmlns:z="urn:old-exception"/></p:sp>', 1)
    prs = rpptx.Presentation.from_bytes(_package_bytes(parts))
    slide = prs.slides[0]
    shapes = slide.shapes[0].shapes if change == "group-id" else slide.shapes
    held = shapes[0]
    element = slide.element
    xml = element.xml
    if change == "same-owner-text":
        changed = xml.replace("before", "after", 1).replace("<p:cNvPr ", '<p:cNvPr xmlns:q="urn:unrelated" ', 1)
    elif change == "group-id":
        changed = xml.replace(f'id="{group_id}"', 'id="99"', 1)
    elif change == "fresh-id":
        changed = xml.replace(f'id="{first_id}"', 'id="99"', 1)
    else:
        changed = xml.replace(f'id="{first_id}"', 'id="99"', 1).replace(f'id="{second_id}"', f'id="{first_id}"', 1).replace('id="99"', f'id="{second_id}"', 1)
    assert changed != xml
    before = prs.to_bytes()
    if change == "same-owner-text":
        element.replace_xml(changed)
        assert prs.slides[0].shapes[0].text == "after"
        assert b'<z:opaque xmlns:z="urn:old-exception"/>' in _package_parts(prs.to_bytes())["ppt/slides/slide1.xml"]
    else:
        with pytest.raises((ValueError, rpptx.RpptxError)):
            element.replace_xml(changed)
        assert prs.to_bytes() == before
        assert element.xml == xml
        assert held.shape_id == first_id and held.text == "before"


def test_issue388_ppt_append_owner_identity():
    """Expected on old source: the paragraph handle and the earlier run become
    stale after the first add_run. The contract is that both stay usable
    across a second add_run and add_field, and the new run addresses itself.
    """
    import rpptx

    presentation = _textbox_presentation(rpptx)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    earlier = paragraph.runs[0]
    assert earlier.text == "hello"
    first = paragraph.add_run(" one")
    assert earlier.text == "hello"
    assert first.text == " one"
    assert paragraph.text == "hello one"
    paragraph.add_field("slidenum", "2")
    second = paragraph.add_run(" two")
    assert earlier.text == "hello"
    assert first.text == " one"
    assert second.text == " two"
    assert paragraph.text == "hello one2 two"
    assert paragraph.runs[0].text == "hello"
    assert paragraph.runs[1].text == " one"
    assert paragraph.runs[2].text == " two"


def test_issue388_ppt_append_collection_contract():
    """Expected on old source: add_paragraph retires the earlier run and the
    already materialised paragraph. The contract keeps those owners live, leaves
    the materialised collection at its earlier length, and shows the new
    paragraph only to a later iteration.
    """
    import rpptx

    presentation = _textbox_presentation(rpptx)
    frame = presentation.slides[0].shapes[0].text_frame
    earlier_run = frame.paragraphs[0].runs[0]
    materialised = list(frame.paragraphs)
    assert len(materialised) == 1
    assert materialised[0].text == "hello"
    frame.add_paragraph()
    assert earlier_run.text == "hello"
    assert materialised[0].text == "hello"
    assert len(materialised) == 1
    again = list(frame.paragraphs)
    assert len(again) == 2
    assert again[0].text == "hello"
    assert again[1].text == ""
    assert earlier_run.text == "hello"


def test_issue389_local_edits_do_not_scale_with_unrelated_package():
    """The five PowerPoint operations edit one text box on a one-slide deck
    and on a deck of 40 slides with 40 text boxes. The edited shape changes.
    A shape on the last slide does not. The visit bound is a library test.
    The release ratio is final-sprint evidence.
    """
    import rpptx
    from rpptx.enum.dml import MSO_FILL_TYPE

    for slides, boxes in ((1, 0), (40, 40)):
        presentation = rpptx.Presentation()
        for _ in range(slides):
            presentation.slides.add_slide(presentation.slide_layouts[6])
        for slide in presentation.slides:
            for _ in range(boxes):
                slide.shapes.add_textbox(0, 0, 100, 100)
        shape = presentation.slides[0].shapes.add_textbox(0, 0, 100, 100)
        shape.alt_text = "alt"
        assert shape.alt_text == "alt"
        shape.left = 200_000
        assert int(shape.left) == 200_000
        shape.fill.solid()
        assert shape.fill.type == MSO_FILL_TYPE.SOLID
        xml = shape.element.xml
        assert "TextBox " in xml
        shape.element.replace_xml(xml.replace("TextBox ", "TextBax ", 1))
        assert "TextBax " in shape.element.xml
        if slides > 1:
            other = presentation.slides[-1].shapes[0]
            assert other.alt_text in (None, "")
            assert "TextBox " in other.element.xml
            assert int(other.left) == 0


def test_issue403_shape_edits_do_the_same_inventory_work_at_n_and_4n_slides():
    """Debug builds count the owner inventory work of each edit: the owner
    addresses walked, the survival proofs matched and the owners indexed.
    An edit on the first slide does the same work whatever the number of later
    slides, which only keep their inventory parts. Release builds do not count.
    """
    import rpptx
    from rpptx import _rpptx
    from rpptx.enum.shapes import MSO_SHAPE

    inventory_work = getattr(_rpptx, "_inventory_work", None)
    if inventory_work is None:
        pytest.skip("release build: the inventory work counter is debug-only")

    def work(slides):
        presentation = rpptx.Presentation()
        for _ in range(slides):
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            for _ in range(6):
                slide.shapes.add_textbox(0, 0, 100, 100).text_frame.text = "before"
            slide.shapes.add_table(3, 3, 0, 0, 900, 900)
        slide = presentation.slides[0]
        shapes = slide.shapes
        table = shapes[6].table
        counts = {}

        def count(name, edit):
            inventory_work()
            edit()
            counts[name] = inventory_work()

        def replace_shape_text():
            element = shapes[2].element
            element.replace_xml(element.xml.replace(">before<", ">after<"))

        count("ShapeCollection.add_shape", lambda: shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, 100, 100))
        count("Paragraph.add_run", lambda: shapes[0].text_frame.paragraphs[0].add_run("run"))
        count("TextFrame.add_paragraph", lambda: shapes[0].text_frame.add_paragraph())
        count("Shape.text", lambda: setattr(shapes[1], "text", "after"))
        count("TextFrame.text", lambda: setattr(shapes[3].text_frame, "text", "after"))
        count("Cell.text", lambda: setattr(table.cell(1, 1), "text", "cell"))
        count("RowCollection.add_row", lambda: table.rows.add_row())
        count("Element.replace_xml (shape text)", replace_shape_text)
        count("Slide.try_replace_text", lambda: slide.try_replace_text("before", "later"))
        count("SlideCollection.add_slide", lambda: presentation.slides.add_slide(presentation.slide_layouts[6]))
        return counts

    small, large = work(3), work(12)
    assert all(small.values()), small
    assert large == small


def test_issue403_package_edits_do_the_same_commit_work_at_n_and_4n_slides():
    """Debug builds count the work of staging and reopening a package: the
    parts parsed, the slide and notes parts serialized, the media hashed and
    the packages written as a ZIP, and the owner inventory work. A
    package-bound edit on the first slide of a deck read from bytes does the
    same work whatever the number of later slides, and writes no ZIP.
    Release builds do not count.
    """
    import io

    import rpptx
    from rpptx import _rpptx

    commit_work = getattr(_rpptx, "_commit_work", None)
    if commit_work is None:
        pytest.skip("release build: the commit work counter is debug-only")
    inventory_work = _rpptx._inventory_work

    def work(slides):
        built = rpptx.Presentation()
        for index in range(slides):
            slide = built.slides.add_slide(built.slide_layouts[6])
            for _ in range(4):
                slide.shapes.add_textbox(0, 0, 100, 100).text_frame.text = "before"
            slide.shapes.add_table(2, 2, 0, 0, 900, 900)
            slide.shapes.add_picture(io.BytesIO(_tiny_png(index, 0, 0)), 0, 0, 100, 100)
            slide.notes_slide.notes_text_frame.text = "notes"
        presentation = rpptx.Presentation.from_bytes(built.to_bytes())
        slide = presentation.slides[0]
        shapes = slide.shapes
        counts = {}

        def count(name, edit, inventory=True):
            commit_work()
            inventory_work()
            edit()
            counts[name] = (*commit_work(), inventory_work() if inventory else None)

        count("CoreProperties.title", lambda: setattr(presentation.core_properties, "title", "edited"))
        count("SlideTransition.type", lambda: setattr(slide.transition, "type", "fade"))
        count("SlideTransition.duration", lambda: setattr(slide.transition, "duration", 1.5))
        count("HeaderFooter.footer", lambda: setattr(slide.header_footer, "footer", "footer"))
        count("HeaderFooter.slide_number", lambda: setattr(slide.header_footer, "slide_number", True))
        count("NotesTextFrame.text", lambda: setattr(slide.notes_slide.notes_text_frame, "text", "edited"))
        count("FillFormat.solid (table cell)", lambda: shapes[4].table.cell(0, 0).fill.solid())
        count("Shape.replace_image", lambda: shapes[5].replace_image(io.BytesIO(_tiny_png(1, 2, 3))))
        count("Shape.flip_h", lambda: setattr(shapes[0], "flip_h", True))
        count("ShapeCollection.remove", lambda: shapes.remove(shapes[1]))
        # Removing a slide moves the owners of every later slide to a new
        # ordinal, so only its commit work is compared.
        count(
            "SlideCollection.remove",
            lambda: presentation.slides.remove(presentation.slides[1]),
            inventory=False,
        )
        return counts

    small, large = work(3), work(12)
    assert all(zip_writes == 0 for *_, zip_writes, _ in small.values()), small
    assert large == small


def _issue403_deck_with_other_slide_shapes(extra_shapes):
    """Three slides with a text box each, the third rewritten by `extra_shapes`."""
    import rpptx

    built = rpptx.Presentation()
    for _ in range(3):
        slide = built.slides.add_slide(built.slide_layouts[6])
        for _ in range(2):
            slide.shapes.add_textbox(0, 0, 100, 100).text_frame.text = "t"
    source = zipfile.ZipFile(io.BytesIO(built.to_bytes()))
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "ppt/slides/slide3.xml":
                data = extra_shapes(data.decode()).encode()
            target.writestr(info, data)
    return output.getvalue()


def _issue403_repeated_shape_ids(xml):
    ids = re.findall(r'<p:cNvPr id="(\d+)"', xml)
    return re.sub(
        r'(<p:cNvPr id=")(\d+)(")',
        lambda match: match.group(1) + (match.group(2) if match.group(2) == ids[0] else ids[1]) + match.group(3),
        xml,
    )


def _issue403_alternate_content(inner_ids):
    def edit(xml):
        shapes = []
        for inner in inner_ids:
            shape = (
                f'<p:sp><p:nvSpPr><p:cNvPr id="{inner}" name="AC {inner}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                "<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>ac</a:t></a:r></a:p></p:txBody></p:sp>"
            )
            shapes.append(
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                f'<mc:Choice Requires="p14">{shape}</mc:Choice><mc:Fallback>{shape}</mc:Fallback>'
                "</mc:AlternateContent>"
            )
        xml = xml.replace("</p:spTree>", "".join(shapes) + "</p:spTree>", 1)
        return xml.replace(
            "<p:sld ", '<p:sld xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" ', 1
        )

    return edit


@pytest.mark.parametrize(
    "other_slide",
    ["alternate content, distinct ids", "alternate content, one id", "repeated shape id"],
)
def test_issue403_footer_and_transition_edits_ignore_another_slide(other_slide):
    """Transition and header and footer edits of slide 1 match the owners
    of the slides they change only. Two `mc:AlternateContent` shapes or a
    repeated shape ID on slide 3 used to refuse both families as an ambiguous
    footer owner identity. Header and footer edits add placeholder shapes,
    and every edit that adds or removes shapes, such as add_shape, still
    refuses a deck where another slide repeats a shape ID."""
    import rpptx
    from rpptx.enum.shapes import MSO_SHAPE

    extra = {
        "alternate content, distinct ids": _issue403_alternate_content((50, 51)),
        "alternate content, one id": _issue403_alternate_content((50, 50)),
        "repeated shape id": _issue403_repeated_shape_ids,
    }[other_slide]
    data = _issue403_deck_with_other_slide_shapes(extra)
    edits = {
        "SlideTransition.type": lambda slide: setattr(slide.transition, "type", "fade"),
        "SlideTransition.advance_after": lambda slide: setattr(slide.transition, "advance_after", 2.0),
        "HeaderFooter.footer": lambda slide: setattr(slide.header_footer, "footer", "F"),
        "HeaderFooter.slide_number": lambda slide: setattr(slide.header_footer, "slide_number", True),
        "ShapeCollection.add_shape": lambda slide: slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, 10, 10),
    }
    for name, edit in edits.items():
        presentation = rpptx.Presentation.from_bytes(data)
        kept = presentation.slides[2].shapes[0]
        structural = name.startswith(("HeaderFooter", "ShapeCollection"))
        if other_slide == "repeated shape id" and structural:
            with pytest.raises(ValueError, match="structural owner proof has an ambiguous shape id"):
                edit(presentation.slides[0])
            continue
        edit(presentation.slides[0])
        assert kept.text_frame.text == "t", name
        assert len(rpptx.Presentation.from_bytes(presentation.to_bytes()).slides) == 3


@pytest.mark.parametrize("kind", ["shape", "table", "cell", "body", "paragraph", "run", "slide", "layout", "master"])
def test_issue389_element_publication_and_refusal_for_every_kind(kind):
    import rpptx

    original = rpptx.Presentation()
    first = original.slides.add_slide(original.slide_layouts[6])
    for _ in range(2):
        first.shapes.add_textbox(0, 0, 1000, 1000)
    target_slide = original.slides.add_slide(original.slide_layouts[6])
    target_slide.shapes.add_textbox(0, 0, 1000, 1000).text = "target"
    target_slide.shapes.add_table(1, 1, 0, 0, 1000, 1000).table.cell(0, 0).text = "payload"
    unique = original.to_bytes()
    damaged = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(unique)) as source, zipfile.ZipFile(damaged, "w") as out:
        for member in source.infolist():
            data = source.read(member.filename)
            if member.filename == "ppt/slides/slide1.xml":
                ids = re.findall(rb'<p:cNvPr\b[^>]*\bid="(\d+)"', data)
                assert len(ids) == 3
                data, count = re.subn(rb'(<p:cNvPr\b[^>]*\bid=")' + ids[-1] + rb'(")',
                                      lambda match: match[1] + ids[-2] + match[2], data, count=1)
                assert count == 1
            out.writestr(member, data)

    def select(deck):
        slide = deck.slides[1]
        shape = slide.shapes[0]
        table = slide.shapes[1].table
        return {
            "shape": lambda: shape.element,
            "table": lambda: table.element,
            "cell": lambda: table.cell(0, 0).element,
            "body": lambda: shape.text_frame.element,
            "paragraph": lambda: shape.text_frame.paragraphs[0].element,
            "run": lambda: shape.text_frame.paragraphs[0].runs[0].element,
            "slide": lambda: slide.element,
            "layout": lambda: deck.slide_layouts[6].element,
            "master": lambda: deck.slide_masters[0].element,
        }[kind]()

    def changed(xml):
        if kind in ("shape", "body"):
            closing = "</p:txBody>" if kind == "shape" else "</a:txBody>"
            # The physical shape body uses p:txBody on both owning surfaces.
            if closing not in xml:
                closing = "</p:txBody>"
            assert closing in xml
            return xml.replace(closing, "<a:p><a:r><a:t>new</a:t></a:r></a:p>" + closing, 1)
        if kind == "paragraph":
            assert "</a:p>" in xml
            return xml.replace("</a:p>", "<a:r><a:t>new</a:t></a:r></a:p>", 1)
        if kind in ("table", "cell"):
            # Grow the physical owner inventory. A same-topology text edit is
            # legal even when an unrelated imported slide has duplicate IDs.
            assert "</a:txBody>" in xml
            return xml.replace("</a:txBody>",
                               "<a:p><a:r><a:t>new</a:t></a:r></a:p></a:txBody>", 1)
        if kind == "run":
            assert "target" in xml
            return xml.replace("target", "edited", 1)
        match = re.search(r"<p:cSld\b[^>]*>", xml)
        assert match
        tag = match[0]
        if re.search(r'\sname="[^"]*"', tag):
            tag = re.sub(r'\sname="[^"]*"', ' name="edited"', tag, count=1)
        else:
            tag = tag.replace("<p:cSld", '<p:cSld name="edited"', 1)
        return xml[:match.start()] + tag + xml[match.end():]

    if kind in ("table", "cell"):
        stable = rpptx.Presentation.from_bytes(damaged.getvalue())
        stable_element = select(stable)
        stable_element.replace_xml(stable_element.xml.replace("payload", "edited", 1))
        reopened = rpptx.Presentation.from_bytes(stable.to_bytes())
        assert reopened.slides[1].shapes[1].table.cell(0, 0).text == "edited"

    broken = rpptx.Presentation.from_bytes(damaged.getvalue())
    element = select(broken)
    before_xml = element.xml
    before = broken.to_bytes()
    held_run = broken.slides[1].shapes[0].text_frame.paragraphs[0].runs[0]
    if kind in ("shape", "table", "cell", "body", "paragraph"):
        with pytest.raises(ValueError, match="structural owner proof has an ambiguous shape id"):
            element.replace_xml(changed(before_xml))
        assert broken.to_bytes() == before
        assert element.xml == before_xml
        assert held_run.text == "target"
        assert broken.slides[1].shapes[0].text_frame.paragraphs[0].runs[0].text == "target"
        assert broken.slides[1].shapes[1].table.cell(0, 0).text == "payload"
    else:
        # These edits preserve the relevant physical topology. Existing policy
        # permits them despite unrelated imported duplicate IDs.
        element.replace_xml(changed(before_xml))
        assert element.xml != before_xml
        reopened = rpptx.Presentation.from_bytes(broken.to_bytes())
        assert select(reopened).xml == element.xml

    control = rpptx.Presentation.from_bytes(unique)
    control_element = select(control)
    control_before = control_element.xml
    control_element.replace_xml(changed(control_before))
    assert control_element.xml != control_before
    reopened = rpptx.Presentation.from_bytes(control.to_bytes())
    assert select(reopened).xml == control_element.xml


def test_issue389_empty_body_removal_has_no_ghost_owner_or_resurrection():
    import rpptx

    deck = rpptx.Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 1000, 1000)
    source_bytes = deck.to_bytes()
    result = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source_bytes)) as source, zipfile.ZipFile(result, "w") as out:
        for member in source.infolist():
            data = source.read(member.filename)
            if member.filename == "ppt/slides/slide1.xml":
                data, count = re.subn(rb'<p:txBody\b[^>]*>.*?</p:txBody>',
                                      b'<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody>', data, count=1, flags=re.S)
                assert count == 1
            out.writestr(member, data)
    deck = rpptx.Presentation.from_bytes(result.getvalue())
    shape = deck.slides[0].shapes[0]
    held_body = shape.text_frame
    assert len(held_body.paragraphs) == 0
    xml = shape.element.xml
    without_body, count = re.subn(r'<p:txBody\b[^>]*>.*?</p:txBody>', '', xml, count=1, flags=re.S)
    assert count == 1
    shape.element.replace_xml(without_body)
    with pytest.raises(rpptx.StaleElementError):
        _ = held_body.text
    # A stale inventory body must not poison a later full publication.
    deck.slides.add_slide(deck.slide_layouts[6])
    fresh = deck.slides[0].shapes[0]
    xml = fresh.element.xml
    xml = xml.replace('</p:sp>', '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>', 1)
    fresh.element.replace_xml(xml)
    assert len(fresh.text_frame.paragraphs) == 1
    with pytest.raises(rpptx.StaleElementError):
        _ = held_body.text
