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
    with pytest.raises(rpptx.StaleElementError, match=r"revision 2.*revision 3"):
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


def _assert_exact_stale(rpptx, operation, kind, captured, current, recovery):
    expected = (
        f"{kind} handle was created at document revision {captured}, but the "
        f"document is now at revision {current} (a structural change "
        f"invalidated it). {recovery}"
    )
    with pytest.raises(rpptx.StaleElementError) as raised:
        operation()
    assert str(raised.value) == expected


def test_structural_append_invalidates_every_preexisting_path_handle():
    import rpptx

    prs = rpptx.Presentation()
    layouts = prs.slide_layouts
    layout = layouts[6]
    slides = prs.slides
    slide = slides.add_slide(layout)

    _assert_exact_stale(
        rpptx,
        lambda: len(slides),
        "slide collection",
        0,
        1,
        "Re-fetch it with prs.slides.",
    )
    _assert_exact_stale(
        rpptx,
        lambda: len(layouts),
        "slide layout collection",
        0,
        1,
        "Re-fetch it with prs.slide_layouts.",
    )
    _assert_exact_stale(
        rpptx,
        lambda: layout.name,
        "slide layout",
        0,
        1,
        "Re-fetch it with prs.slide_layouts[6].",
    )
    shapes = slide.shapes
    shapes.add_textbox(0, 0, 100, 100)

    _assert_exact_stale(
        rpptx,
        lambda: slide.shapes,
        "slide",
        1,
        2,
        "Re-fetch it with prs.slides[0].",
    )
    _assert_exact_stale(
        rpptx,
        lambda: len(shapes),
        "shape collection",
        1,
        2,
        "Re-fetch it with prs.slides[0].shapes.",
    )


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
    prs.slides.add_slide(prs.slide_layouts[6])

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
        _assert_exact_stale(rpptx, operation, kind, 4, 5, recovery)


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
        lambda: shape.text,
        lambda: frame.text,
        lambda: paragraph.text,
        lambda: run.text,
    ):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
    assert prs.slides[0].shapes[-1].text == "after"

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_textbox(0, 0, 100, 100).text_frame
    frame.text = "before"
    frame = prs.slides[0].shapes[-1].text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    frame.text = "after"
    for operation in (lambda: frame.text, lambda: paragraph.text, lambda: run.text):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
    assert prs.slides[0].shapes[-1].text_frame.text == "after"

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    frame = slide.shapes.add_textbox(0, 0, 100, 100).text_frame
    frame.text = "before"
    frame = prs.slides[0].shapes[-1].text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.runs[0]
    paragraph.text = "after"
    for operation in (lambda: paragraph.text, lambda: run.text):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
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
    for operation in (
        lambda: shape.left,
        lambda: len(slide.shapes),
        lambda: second.text,
        lambda: font.bold,
    ):
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
    prs.slides.add_slide(prs.slide_layouts[6])

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
        _assert_exact_stale(rpptx, operation, kind, 0, 1, recovery)


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
    authored = (
        _author_documented_decks(
            rpptx.Presentation,
            RpptxInches,
            RpptxPt,
            RpptxShape,
            tmp_path / "rpptx-authored",
            image_path,
        ),
        _author_documented_decks(
            OraclePresentation,
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
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}

    content_types = parts["[Content_Types].xml"].decode()
    parts["[Content_Types].xml"] = content_types.replace(
        "</Types>",
        '<Override PartName="/ppt/notesSlides/notesSlide1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
        "</Types>",
    ).encode()
    slide_rels = parts["ppt/slides/_rels/slide1.xml.rels"].decode()
    parts["ppt/slides/_rels/slide1.xml.rels"] = slide_rels.replace(
        "</Relationships>",
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" '
        'Target="../notesSlides/notesSlide1.xml"/>'
        "</Relationships>",
    ).encode()
    parts["ppt/notesSlides/notesSlide1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/>"
        "<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"
        "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Notes Placeholder\"/>"
        "<p:cNvSpPr/><p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr>"
        "</p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>"
        '<a:r><a:rPr b="1"><a:extLst>'
        '<a:ext uri="{6D487B31-4C56-4F45-AED3-4C741AC43E77}">'
        '<x:payload xmlns:x="urn:rdocx:test"/></a:ext></a:extLst></a:rPr>'
        f"<a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/>"
        "</p:clrMapOvr></p:notes>"
    ).encode()
    parts["ppt/notesSlides/_rels/notesSlide1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" '
        'Target="../notesMasters/notesMaster1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        'Target="../slides/slide1.xml"/></Relationships>'
    ).encode()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


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
    with pytest.raises(rpptx.StaleElementError):
        _ = held.notes_text

    presentation.save(output)
    assert rpptx.Presentation(output).slides[0].notes_text == "Final speaker note"
    with zipfile.ZipFile(output) as archive:
        notes_xml = archive.read("ppt/notesSlides/notesSlide1.xml").decode()
    assert '<p:ph type="body" idx="1"' in notes_xml
    assert 'b="1"' in notes_xml
    assert "x:payload" in notes_xml

    without_notes = rpptx.Presentation()
    held = without_notes.slides.add_slide(without_notes.slide_layouts[6])
    assert held.notes_text is None
    held.notes_text = "Created speaker note"
    assert without_notes.slides[0].notes_text == "Created speaker note"
    with pytest.raises(rpptx.StaleElementError):
        _ = held.notes_text


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
    assert run.font.color == "123456"
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
    return slide_xml[slide_xml.index("<p:txBody>") : slide_xml.index("</p:txBody>")]


def _replace_in_slide(source, target, old, new):
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    slide_xml = parts["ppt/slides/slide1.xml"].decode()
    assert old in slide_xml
    parts["ppt/slides/slide1.xml"] = slide_xml.replace(old, new, 1).encode()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _autofit_textbox(rpptx, lines, *, height=None, size=18):
    """A python-pptx-style 3 x 1.5 inch text box of Arial lines."""
    presentation = rpptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_textbox(
        rpptx.Inches(1), rpptx.Inches(1), rpptx.Inches(3), height or rpptx.Inches(1.5)
    )
    presentation.slides[0].shapes[0].text_frame.word_wrap = True
    _write_arial_lines(rpptx, presentation, lines, size)
    return presentation, presentation.slides[0].shapes[0]


def _write_arial_lines(rpptx, presentation, lines, size=18):
    """Replaces the first shape's text with numbered Arial lines."""
    presentation.slides[0].shapes[0].text_frame.text = "\n".join(
        f"Item {index}" for index in range(lines)
    )
    frame = presentation.slides[0].shapes[0].text_frame
    for paragraph in frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = rpptx.Pt(size)
    return frame


def test_setting_text_to_fit_shape_stores_the_shrink_powerpoint_chose(tmp_path):
    # Six 18 point Arial lines in a 1.5 inch box: PowerPoint for Mac stores
    # fontScale 92.5 % with 20 % less line spacing (ground truth, #310).
    import rpptx
    from rpptx import MSO_AUTO_SIZE

    presentation, shape = _autofit_textbox(rpptx, 6)
    frame = shape.text_frame
    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    output = tmp_path / "autofit.pptx"
    presentation.save(output)

    assert '<a:normAutofit fontScale="92500" lnSpcReduction="20000"/>' in _text_body_xml(output)
    assert frame.auto_size is MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    (layout,) = presentation.text_layout()
    assert (layout.font_scale, layout.overflow) == (0.925, False)
    assert {line.font_size for line in layout.lines} == {17.0}

    # Later edits need an explicit refresh, as PowerPoint refits on edit.
    frame = _write_arial_lines(rpptx, presentation, 12)
    assert '<a:normAutofit fontScale="92500" lnSpcReduction="20000"/>' in _text_body_xml(output)
    result = frame.refresh_autofit()
    assert isinstance(result, rpptx.AutofitResult)
    assert (result.autofit, result.fits, result.height, result.font_size) == (
        "normal",
        True,
        None,
        None,
    )
    assert result.font_scale < 0.925 and result.line_spacing_reduction == 0.2
    assert ("Arial", "Liberation Sans") in result.font_substitutions
    assert presentation.refresh_autofit() == (result,)
    assert not presentation.text_layout()[0].overflow

    frame.auto_size = MSO_AUTO_SIZE.NONE
    assert frame.refresh_autofit() is None
    assert presentation.refresh_autofit() == ()


def test_setting_shape_to_fit_text_resizes_the_shape_keeping_its_anchored_edge():
    import rpptx
    from rpptx import MSO_ANCHOR, MSO_AUTO_SIZE

    # The python-pptx order: autofit first, text afterwards, then save.
    presentation, shape = _autofit_textbox(rpptx, 0)
    bottom = shape.top + shape.height
    shape.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    shape.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    assert shape.height == rpptx.Inches(1.5)
    _write_arial_lines(rpptx, presentation, 2)
    presentation.to_bytes()

    # PowerPoint for Mac sizes two 18 point lines to 646331 EMU.
    shape = presentation.slides[0].shapes[0]
    assert shape.height == 646_331
    assert shape.top + shape.height == bottom
    assert shape.left == rpptx.Inches(1)
    (result,) = presentation.refresh_autofit()
    assert (result.autofit, result.height, result.width) == (
        "shape",
        rpptx.Length(646_331),
        None,
    )


def test_marked_frames_refresh_on_save_and_none_unmarks_them(tmp_path):
    import rpptx
    from rpptx import MSO_AUTO_SIZE

    presentation, shape = _autofit_textbox(rpptx, 0)
    shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    _write_arial_lines(rpptx, presentation, 6)
    output = tmp_path / "marked.pptx"
    presentation.save(output)
    assert '<a:normAutofit fontScale="92500" lnSpcReduction="20000"/>' in _text_body_xml(output)

    # Text edited after a save is refitted by the next one.
    _write_arial_lines(rpptx, presentation, 8)
    presentation.save(output)
    assert '<a:normAutofit fontScale="70000" lnSpcReduction="20000"/>' in _text_body_xml(output)

    # Back to NONE, the frame is no longer refitted.
    frame = presentation.slides[0].shapes[0].text_frame
    frame.auto_size = MSO_AUTO_SIZE.NONE
    presentation.save(output)
    assert "<a:noAutofit/>" in _text_body_xml(output)

    # A marked shape that is removed is forgotten.
    presentation.slides[0].shapes[0].text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    shapes = presentation.slides[0].shapes
    shapes.remove(shapes[0])
    assert presentation.text_layout() == ()
    presentation.save(output)


def test_fit_text_writes_the_largest_fitting_size_on_every_run(tmp_path):
    import pathlib

    import rpptx
    from rpptx import MSO_AUTO_SIZE

    presentation, shape = _autofit_textbox(rpptx, 6)
    frame = shape.text_frame
    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    result = frame.fit_text(font_family="Calibri", max_size=40, bold=True)

    # Six Calibri lines fit the 100.8 point text area at 14 points.
    assert result.font_size == rpptx.Pt(14)
    assert result.font_substitutions == (("Calibri", "Carlito"),)
    assert frame.auto_size is MSO_AUTO_SIZE.NONE
    assert frame.word_wrap is True
    for paragraph in frame.paragraphs:
        font = paragraph.runs[0].font
        assert (font.name, font.size, font.bold, font.italic) == (
            "Calibri",
            rpptx.Pt(14),
            True,
            False,
        )
    output = tmp_path / "fit.pptx"
    presentation.save(output)
    body = _text_body_xml(output)
    assert "<a:noAutofit/>" in body
    assert body.count('sz="1400"') == 12

    # The python-pptx defaults keep each run's typeface, cap at 18 points,
    # and measure with a caller font file when one is given.
    font_file = (
        pathlib.Path(__file__).resolve().parents[2]
        / "oxml-layout"
        / "fonts"
        / "LiberationSans-Regular.ttf"
    )
    result = frame.fit_text(font_file=font_file)
    assert result.font_size == rpptx.Pt(14)
    assert frame.paragraphs[0].runs[0].font.name == "Calibri"
    assert frame.paragraphs[0].runs[0].font.bold is False

    # python-pptx takes int(max_size).
    assert frame.fit_text(max_size=18.9).font_size == rpptx.Pt(14)
    with pytest.raises(ValueError, match="max_size"):
        frame.fit_text(max_size=0)
    with pytest.raises(ValueError, match="font_file"):
        frame.fit_text(font_file=tmp_path / "missing.ttf")
    tiny, tiny_shape = _autofit_textbox(rpptx, 40, height=rpptx.Pt(2))
    with pytest.raises(rpptx.RpptxError, match="even at 1 point"):
        tiny_shape.text_frame.fit_text()
    empty, empty_shape = _autofit_textbox(rpptx, 0)
    assert empty_shape.text_frame.fit_text() is None


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

    assert _text_body_xml(output).startswith(
        '<p:txBody><a:bodyPr lIns="228600" tIns="38100" rIns="0" bIns="50800" '
        'anchor="b" wrap="none"><a:spAutoFit/></a:bodyPr>'
    )
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
    assert _text_body_xml(output).startswith("<p:txBody><a:bodyPr/>")


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
    assert [getattr(font, name) for name in names] == [None] * len(names)

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
    assert [getattr(font, name) for name in names] == [
        True,
        True,
        True,
        True,
        "Georgia",
        "123456",
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
    assert alpha.color == "FF0000"
    assert theme.color is None
    alpha.color = "00ff00"
    theme.color = (1, 2, 3)
    assert (alpha.color, theme.color) == ("00FF00", "010203")
    assert paragraph.font.color is None
    paragraph.font.color = rpptx.RGBColor.from_string("ABCDEF")
    assert paragraph.font.color == "ABCDEF"
    presentation.save(output)

    body = _text_body_xml(output)
    assert '<a:srgbClr val="00FF00"><a:alpha val="50000"/></a:srgbClr>' in body
    assert '<a:solidFill><a:srgbClr val="010203"/></a:solidFill>' in body
    assert "schemeClr" not in body
    assert '<a:defRPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:defRPr>' in body

    alpha.color = None
    presentation.save(output)
    assert "00FF00" not in _text_body_xml(output)


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


def test_add_run_returns_a_live_run_and_stales_the_paragraph_handle_once():
    import rpptx

    presentation = _textbox_presentation(rpptx)
    paragraph = presentation.slides[0].shapes[0].text_frame.paragraphs[0]
    run = paragraph.add_run(" world")
    assert run.text == " world"
    run.font.bold = True
    assert run.font.bold is True
    _assert_stale_after_exactly_one_bump(rpptx, lambda: paragraph.text)

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
    assert (font.italic, font.underline, font.name, font.color, font.size) == (
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
    with pytest.raises(rpptx.StaleElementError, match=r"\.runs\[0\]\.hyperlink\."):
        _ = held.address
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
    with pytest.raises(rpptx.StaleElementError, match=r"\.click_action"):
        _ = held.target_slide
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
    with pytest.raises(rpptx.RpptxError):
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
    second.follow_master_background = False
    assert second.background.fill.type == MSO_FILL_TYPE.BACKGROUND
    output = tmp_path / "slides-out.pptx"
    prs.save(output)

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output)
    assert oracle.slides[0]._element.get("show") == "0"
    assert oracle.slides[0].follow_master_background is False
    assert oracle.slides[0].background.fill.fore_color.rgb == pptx.dml.color.RGBColor(0x12, 0x34, 0x56)
    assert oracle.slides[1].follow_master_background is False

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
    _assert_stale_after_exactly_one_bump(rpptx, lambda: slide.slide_layout)
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
    # Pins a known gap: the ECMA preset table has no upArrow, python-pptx has UP_ARROW.
    with pytest.raises(ValueError, match="MSO_SHAPE has no XML mapping for 'upArrow'"):
        MSO_SHAPE.from_xml("upArrow")

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
    with pytest.raises(ValueError, match="unsupported MSO_SHAPE value"):
        rounded.auto_shape_type = 35
    with pytest.raises(rpptx.RpptxError, match="unknown DrawingML preset geometry: upArrow"):
        rounded.auto_shape_type = "upArrow"

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
    with pytest.raises(ValueError, match="assigned value must be type RGBColor"):
        fill.fore_color.rgb = (1, 2, 3)

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
    with pytest.raises(rpptx.StaleElementError):
        _ = held.type
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
    with pytest.raises(rpptx.StaleElementError, match=r"table\.cell\(2, 0\)\.fill"):
        _ = held_fill.type
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
    with pytest.raises(rpptx.StaleElementError, match=r"table\.cell\(0, 0\)\.border_right\."):
        _ = held_border.width
    with pytest.raises(rpptx.StaleElementError, match=r"table\.rows\[2\]\."):
        _ = held.height
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
    for operation in (lambda: held_cell.text, lambda: len(rows)):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
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
    prs.slides.add_slide(prs.slide_layouts[6])
    with pytest.raises(rpptx.StaleElementError, match=r"table\.rows\[0\]\."):
        prs.slides[0].shapes[1].table.rows.remove(held_row)
    prs.slides.remove(prs.slides[1])
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
    with pytest.raises(rpptx.StaleElementError, match=r"shapes\[1\]\.line\.tail_end"):
        _ = held.type

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
    with pytest.raises(ValueError, match="assigned value must be type RGBColor"):
        shapes[0].shadow.color.rgb = (1, 2, 3)
    held = shapes[0].shadow
    held_color = held.color
    prs.slides.add_slide(prs.slide_layouts[6])
    with pytest.raises(rpptx.StaleElementError):
        _ = held.visible
    with pytest.raises(rpptx.StaleElementError):
        _ = held_color.rgb
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


def test_shapes_move_changes_the_z_order_and_stales_handles_once(tmp_path):
    import rpptx

    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    for name in ("back", "middle", "front"):
        prs.slides[0].shapes.add_shape("rect", 0, 0, 100, 100).name = name
    shapes = prs.slides[0].shapes
    held = shapes[0]
    shapes.move(0, -1)
    assert [shape.name for shape in prs.slides[0].shapes] == ["middle", "front", "back"]
    for operation in (lambda: held.name, lambda: len(shapes)):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
    with pytest.raises(rpptx.StaleElementError, match=r"prs\.slides\[0\]\.shapes\[0\]\."):
        _ = held.name
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
    for operation in (lambda: group.name, lambda: len(shapes)):
        _assert_stale_after_exactly_one_bump(rpptx, operation)
    with pytest.raises(rpptx.StaleElementError, match=r"prs\.slides\[0\]\.shapes\[0\]\.shapes\."):
        shapes.add_textbox(0, 0, 1, 1)
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
    with pytest.raises(ValueError, match="mc:AlternateContent fallback cannot take new shapes"):
        fallback_group.shapes.add_textbox(0, 0, 10, 10)


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
    with pytest.raises(rpptx.StaleElementError):
        _ = held.shapes
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
    with pytest.raises(ValueError, match="shape is not in this collection"):
        prs.slides[1].shapes.remove(shapes[0])
    with pytest.raises(rpptx.StaleElementError):
        shapes.remove(picture)
    shapes.remove(shapes[1])
    with pytest.raises(rpptx.StaleElementError):
        _ = group.name
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
    _assert_stale_after_exactly_one_bump(rpptx, lambda: held.text)
    _assert_stale_after_exactly_one_bump(rpptx, lambda: source.notes_text)
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
    _assert_stale_after_exactly_one_bump(rpptx, lambda: held.text)
    shapes = prs.slides[0].shapes
    assert (shapes[0].text, shapes[1].table.cell(0, 0).text) == ("Ada and Ada", "Ada")
    assert prs.slides[0].notes_text == "Notes for Ada"
    held = prs.slides[0].shapes[0]
    assert prs.try_replace_text("Ada", "Grace") == 4
    _assert_stale_after_exactly_one_bump(rpptx, lambda: held.text)
    assert prs.slides[0].shapes[0].text == "Grace and Grace"


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
    _assert_stale_after_exactly_one_bump(rpptx, lambda: slide.comments)
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
    _assert_stale_after_exactly_one_bump(rpptx, lambda: held.shapes)
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
    _assert_stale_after_exactly_one_bump(rpptx, lambda: frame.text)
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
    with pytest.raises(rpptx.RpptxError, match="has a chart, which is not carried"):
        prs.slides.import_slide(source.slides[1])
    assert prs.to_bytes() == before
    with pytest.raises(ValueError, match="layout is not a layout of this presentation"):
        prs.slides.import_slide(source.slides[0], layout=source.slide_layouts[1])
    with pytest.raises(IndexError):
        prs.slides.import_slide(source.slides[0], index=2)

    held = prs.slides[0]
    imported = prs.slides.import_slide(source.slides[0], index=0)
    with pytest.raises(rpptx.StaleElementError):
        _ = held.shapes
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
    assert len(MSO_SHAPE) == 181
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
    assert set(oracle_members) - {member.name for member in MSO_SHAPE} == {"UP_ARROW"}
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
    gradient = (
        b'<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst>'
        b'<a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>'
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
