import importlib.metadata
import io
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
    # The width the renderer assumed for the unsized deck stays in effect.
    assert (prs.slide_width, prs.slide_height) == (9_144_000, rpptx.Inches(5))


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
    assert slide_xml.count("<a:ln") == 1

    pptx = pytest.importorskip("pptx", reason="python-pptx is the differential oracle")
    oracle = pptx.Presentation(output).slides[0].shapes
    assert oracle[0].fill.type == pptx.enum.dml.MSO_FILL.SOLID
    assert oracle[0].fill.fore_color.rgb == pptx.dml.color.RGBColor(0xFF, 0x00, 0x00)
    assert oracle[0].line.color.rgb == pptx.dml.color.RGBColor(0x00, 0x80, 0x00)
    assert oracle[0].line.width == rpptx.Pt(2)
    assert oracle[1].fill.type == pptx.enum.dml.MSO_FILL.BACKGROUND


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


def test_line_width_none_removes_the_width_attribute_like_python_pptx():
    import rpptx

    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape("rect", 0, 0, 100, 100)
    shape.line.width = rpptx.Pt(2)
    shape.line.width = None
    start_tag = prs.slides[0].shapes[0].xml.split(b"<a:ln", 1)[1].split(b">", 1)[0]
    assert b" w=" not in start_tag
    assert prs.slides[0].shapes[0].line.width == 0


def test_a_group_fill_reads_as_group_and_a_new_fill_replaces_it(tmp_path):
    import rpptx
    from rpptx.enum.dml import MSO_FILL_TYPE

    def build(deck):
        pptx = pytest.importorskip("pptx")
        from lxml import etree
        from pptx.oxml.ns import qn

        shape = deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_shape(
            pptx.enum.shapes.MSO_SHAPE.RECTANGLE, 0, 0, 10, 10
        )
        etree.SubElement(shape._element.spPr, qn("a:grpFill"))

    prs = rpptx.Presentation(_python_pptx_deck(tmp_path / "group-fill.pptx", build))
    fill = prs.slides[0].shapes[0].fill
    assert fill.type == MSO_FILL_TYPE.GROUP
    fill.solid()
    xml = prs.slides[0].shapes[0].xml
    assert b"grpFill" not in xml and xml.count(b"<a:solidFill") == 1
    assert fill.type == MSO_FILL_TYPE.SOLID


def test_assigning_a_line_colour_makes_a_patterned_line_solid(tmp_path):
    import rpptx
    from rpptx.dml.color import RGBColor
    from rpptx.enum.dml import MSO_FILL_TYPE

    def build(deck):
        pptx = pytest.importorskip("pptx")
        shape = deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_shape(
            pptx.enum.shapes.MSO_SHAPE.RECTANGLE, 0, 0, 10, 10
        )
        shape.line.fill.patterned()
        shape.line.fill.pattern = pptx.enum.dml.MSO_PATTERN.PERCENT_5
        shape.line.fill.fore_color.rgb = pptx.dml.color.RGBColor(0, 0, 0xFF)

    prs = rpptx.Presentation(
        _python_pptx_deck(tmp_path / "patterned-line.pptx", build)
    )
    line = prs.slides[0].shapes[0].line
    assert line.fill.type == MSO_FILL_TYPE.PATTERNED
    line.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    assert line.fill.type == MSO_FILL_TYPE.SOLID
    assert str(line.color.rgb) == "FF0000"
