import concurrent.futures
import io
import os
import re
import shutil
import statistics
import struct
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path

import pytest


PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
POPLER_VERSION = "26.01.0"
POPLER_TOOLS = ("pdfinfo", "pdftotext")


def _nontrivial_document(seed):
    from rdocx import Document

    document = Document()
    sentence = (
        f"Independent render {seed} exercises shaping, line breaking, and pagination. "
        "The quick brown fox jumps over the lazy dog while numerals 0123456789 "
        "keep every paragraph substantial enough for the timing gate."
    )
    for index in range(72):
        document.add_paragraph(f"{index + 1}. {sentence} {sentence}")
    return document


def _available_cpu_count():
    if hasattr(os, "sched_getaffinity"):
        return len(os.sched_getaffinity(0))
    return os.cpu_count()


def _render_serial(documents):
    return [document.to_pdf() for document in documents]


def _render_parallel(documents):
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        return list(executor.map(lambda document: document.to_pdf(), documents))


def _timed(render, documents):
    started = time.perf_counter()
    outputs = render(documents)
    elapsed = time.perf_counter() - started
    return elapsed, outputs


def _assert_poppler_version(tool, version_output):
    expected = f"{tool} version {POPLER_VERSION}"
    reported = [
        line.strip()
        for line in version_output.splitlines()
        if line.strip().startswith(f"{tool} version ")
    ]
    assert reported == [expected], f"F-133 requires {expected}, got {reported!r}"


def _poppler_tools():
    resolved = {}
    for tool in POPLER_TOOLS:
        path = shutil.which(tool)
        assert path is not None, f"F-133 requires {tool} from Poppler {POPLER_VERSION}"
        version = subprocess.run(
            [path, "-v"],
            check=True,
            capture_output=True,
            text=True,
        )
        version_output = f"{version.stdout}\n{version.stderr}"
        _assert_poppler_version(tool, version_output)
        resolved[tool] = path
    return resolved


def _pdf_semantics(pdf, poppler):
    with tempfile.TemporaryDirectory() as directory:
        pdf_path = Path(directory) / "render.pdf"
        text_path = Path(directory) / "render.txt"
        pdf_path.write_bytes(pdf)
        page_count = subprocess.run(
            [poppler["pdfinfo"], str(pdf_path)],
            check=True,
            capture_output=True,
            text=True,
        ).stdout
        page_count = next(
            int(line.split(":", 1)[1])
            for line in page_count.splitlines()
            if line.startswith("Pages:")
        )
        subprocess.run(
            [poppler["pdftotext"], str(pdf_path), str(text_path)],
            check=True,
            capture_output=True,
        )
        return page_count, text_path.read_text()


def _assert_releases_gil(operation):
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
        result = operation()
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
    return result


def _font_with_zero_units_per_em():
    font_path = (
        Path(__file__).resolve().parents[2]
        / "oxml-layout"
        / "fonts"
        / "Carlito-Regular.ttf"
    )
    font = bytearray(font_path.read_bytes())
    font[:] = font.replace(b"Carlito", b"FaultyX")
    font[:] = font.replace("Carlito".encode("utf-16-be"), "FaultyX".encode("utf-16-be"))
    table_count = struct.unpack_from(">H", font, 4)[0]
    for table_index in range(table_count):
        record = 12 + table_index * 16
        if font[record : record + 4] == b"head":
            head_offset = struct.unpack_from(">I", font, record + 8)[0]
            struct.pack_into(">H", font, head_offset + 18, 0)
            return bytes(font)
    raise AssertionError("Carlito test font has no head table")


def _document_with_invalid_embedded_font():
    from rdocx import Document

    source = Document()
    run = source.add_paragraph("").add_run("layout must reject this font")
    run.font.name = "FaultyX"
    source_archive = io.BytesIO(source.to_bytes())
    result = io.BytesIO()
    with zipfile.ZipFile(source_archive) as source_zip:
        with zipfile.ZipFile(result, "w") as result_zip:
            for info in source_zip.infolist():
                if info.filename == "word/fontTable.xml":
                    continue
                data = source_zip.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    data = data.replace(
                        b'  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>\n',
                        b"",
                    )
                    data = data.replace(
                        b"</Types>",
                        (
                            b'<Default Extension="ttf" '
                            b'ContentType="application/x-font-ttf"/></Types>'
                        ),
                    )
                elif info.filename == "word/_rels/document.xml.rels":
                    data = data.replace(
                        b'  <Relationship Id="rdocxFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>\n',
                        b"",
                    )
                result_zip.writestr(info, data)
            result_zip.writestr("word/fonts/FaultyX.ttf", _font_with_zero_units_per_em())
    return Document.from_bytes(result.getvalue())


def test_to_pdf_returns_pdf_bytes():
    document = _nontrivial_document(0)

    pdf = document.to_pdf()
    reopened = type(document).from_bytes(document.to_bytes())

    assert isinstance(pdf, bytes)
    assert pdf.startswith(b"%PDF-")
    assert reopened.paragraphs[0].text.startswith("1. Independent render 0")


def test_render_methods_return_png_bytes_and_page_lists():
    document = _nontrivial_document(1)

    first_page = document.render_page_to_png(0, dpi=72.0)
    pages = document.render_all_pages(dpi=72.0)

    assert isinstance(first_page, bytes)
    assert first_page.startswith(PNG_SIGNATURE)
    assert document.render_page_to_png(len(pages), dpi=72.0) is None
    assert isinstance(pages, list)
    assert len(pages) > 1
    assert all(isinstance(page, bytes) and page.startswith(PNG_SIGNATURE) for page in pages)


def test_render_pages_accepts_keyword_options_and_zero_based_pages():
    from rdocx import LayoutError

    document = _nontrivial_document(2)

    png_pages = document.render_pages(
        format="png", transparent=True, dpi=72.0, pages=[0, 1]
    )
    jpeg_pages = document.render_pages(format="jpeg", quality=80, dpi=72.0, pages=[1])
    tiff = document.render_pages(format="tiff", dpi=72.0, pages=[0, 1])

    assert isinstance(png_pages, list)
    assert [page[:8] for page in png_pages] == [PNG_SIGNATURE, PNG_SIGNATURE]
    assert isinstance(jpeg_pages, list)
    assert len(jpeg_pages) == 1
    assert jpeg_pages[0].startswith(b"\xff\xd8")
    assert isinstance(tiff, bytes)
    assert tiff.startswith((b"II*\x00", b"MM\x00*"))
    with pytest.raises(LayoutError):
        document.render_pages(format="jpeg", quality=0, pages=[0])
    with pytest.raises(TypeError):
        document.render_pages("jpeg")


def _caller_font_document():
    from rdocx import Document

    document = Document()
    run = document.add_paragraph("Monospaced caller face").runs[0]
    run.font.name = "Rdocx Test Face"
    return document


def _pdf_base_fonts(pdf):
    return set(re.findall(rb"/BaseFont\s*/([^\s/>]+)", pdf))


def test_to_pdf_takes_caller_fonts_as_bytes_or_from_a_directory(tmp_path):
    mono = Path(__file__).resolve().parents[2] / "oxml-layout" / "fonts"
    mono = mono / "LiberationMono-Regular.ttf"
    document = _caller_font_document()
    assert b"LiberationMono" not in _pdf_base_fonts(document.to_pdf())

    pdf = document.to_pdf(fonts=[("Rdocx Test Face", mono.read_bytes())])
    assert b"LiberationMono" in _pdf_base_fonts(pdf)

    # A directory font takes its family name from the file name.
    shutil.copy(mono, tmp_path / "Rdocx Test Face.ttf")
    assert b"LiberationMono" in _pdf_base_fonts(document.to_pdf(font_dir=tmp_path))
    assert b"LiberationMono" in _pdf_base_fonts(document.to_pdf(font_dir=str(tmp_path)))

    # An empty caller directory falls back to the normal font sources.
    empty = tmp_path / "empty"
    empty.mkdir()
    assert document.to_pdf(font_dir=empty) == document.to_pdf()
    with pytest.raises(FileNotFoundError, match="font directory .*missing"):
        document.to_pdf(font_dir=tmp_path / "missing")
    with pytest.raises(NotADirectoryError, match="is not a directory"):
        document.to_pdf(font_dir=tmp_path / "Rdocx Test Face.ttf")
    with pytest.raises(TypeError):
        document.to_pdf(fonts=[("Rdocx Test Face", "not bytes")])


def test_render_page_to_svg_returns_the_page_and_its_diagnostics():
    from rdocx import SvgDiagnostic, SvgRenderResult

    document = _caller_font_document()

    page = document.render_page_to_svg(0)
    assert isinstance(page, SvgRenderResult)
    assert page.svg.startswith("<svg ")
    assert "Monospaced" in page.svg
    assert all(isinstance(diagnostic, SvgDiagnostic) for diagnostic in page.diagnostics)
    assert document.render_page_to_svg(1) is None
    snapshot = SvgRenderResult(
        svg="<svg/>", diagnostics=[SvgDiagnostic(path="pages[0]", message="omitted")]
    )
    assert snapshot.diagnostics == (SvgDiagnostic(path="pages[0]", message="omitted"),)


def test_to_pdfa_deterministic_writes_the_requested_profile():
    document = _caller_font_document()

    part_two = document.to_pdfa_deterministic()
    assert part_two.startswith(b"%PDF-")
    assert b"<pdfaid:part>2</pdfaid:part>" in part_two
    assert document.to_pdfa_deterministic("pdfa-2b") == part_two
    assert b"<pdfaid:part>3</pdfaid:part>" in document.to_pdfa_deterministic("pdfa-3b")
    with pytest.raises(ValueError, match="profile must be pdfa-2b or pdfa-3b"):
        document.to_pdfa_deterministic("pdfa-1b")


def test_render_errors_reacquire_and_map_cleanly():
    from rdocx import LayoutError, RdocxError

    document = _document_with_invalid_embedded_font()

    # The preserved orphan font is not registered in a Word font table.
    # Feed the same malformed sfnt through the checked caller-font entrypoint.
    before = document.to_bytes()
    with pytest.raises(LayoutError) as raised:
        document.to_pdf(fonts=(("FaultyX", _font_with_zero_units_per_em()),))
    assert document.to_bytes() == before
    assert isinstance(raised.value, RdocxError)
    assert "font parsing error" in str(raised.value)


def test_poppler_pdf_oracle_is_available_at_reviewed_version():
    assert set(_poppler_tools()) == set(POPLER_TOOLS)


def test_poppler_version_pin_rejects_unreviewed_suffix():
    for tool in POPLER_TOOLS:
        reported = f"{tool} version {POPLER_VERSION}-unreviewed"
        with pytest.raises(AssertionError) as raised:
            _assert_poppler_version(tool, reported)
        assert reported in str(raised.value)


def test_to_bytes_releases_gil_for_python_worker():
    document = _nontrivial_document(10)

    def serialize_repeatedly():
        package = b""
        for _ in range(16):
            package = document.to_bytes()
        return package

    package = _assert_releases_gil(serialize_repeatedly)

    assert package.startswith(b"PK")


def test_compare_releases_gil_for_python_worker():
    original = _nontrivial_document(20)
    edited = _nontrivial_document(21)

    diagnostics = _assert_releases_gil(
        lambda: original.compare(
            edited, author="Ada", timestamp="2026-09-14T09:00:00Z"
        )
    )

    assert isinstance(diagnostics, tuple)


def test_revision_resolution_releases_gil_for_python_worker():
    original = _nontrivial_document(30)
    edited = _nontrivial_document(31)
    original.compare(edited, author="Ada", timestamp="2026-09-15T10:30:00Z")
    assert original.revisions

    count = _assert_releases_gil(original.accept_all)

    assert count > 0
    assert original.revisions == ()


def test_layout_releases_gil_for_python_worker():
    document = _nontrivial_document(22)

    fragments = _assert_releases_gil(document.layout)

    assert fragments


def test_rebuild_toc_releases_gil_for_python_worker():
    document = _nontrivial_document(23)

    report = _assert_releases_gil(document.rebuild_toc)

    assert report.entry_count == 0


def test_render_page_to_png_releases_gil_for_python_worker():
    document = _nontrivial_document(11)

    page = _assert_releases_gil(lambda: document.render_page_to_png(0, dpi=72.0))

    assert page.startswith(PNG_SIGNATURE)


def test_font_svg_and_pdfa_renders_release_gil_for_python_worker(tmp_path):
    document = _nontrivial_document(13)
    fonts = Path(__file__).resolve().parents[2] / "oxml-layout" / "fonts"
    shutil.copy(fonts / "Carlito-Regular.ttf", tmp_path)

    pdf = _assert_releases_gil(lambda: document.to_pdf(font_dir=tmp_path))
    page = _assert_releases_gil(lambda: document.render_page_to_svg(0))
    archival = _assert_releases_gil(document.to_pdfa_deterministic)

    assert pdf.startswith(b"%PDF-")
    assert page is not None and page.svg.startswith("<svg ")
    assert archival.startswith(b"%PDF-")


def test_render_all_pages_releases_gil_for_python_worker():
    document = _nontrivial_document(12)

    pages = _assert_releases_gil(lambda: document.render_all_pages(dpi=72.0))

    assert len(pages) > 1
    assert all(page.startswith(PNG_SIGNATURE) for page in pages)


def test_four_concurrent_to_pdf_calls_are_faster_than_serial():
    available_cpus = _available_cpu_count()
    assert available_cpus is not None and available_cpus >= 2, (
        "F-133 concurrency gate requires a supported multi-core test environment"
    )

    _nontrivial_document(-1).to_pdf()
    poppler = _poppler_tools()

    serial_samples = []
    parallel_samples = []
    for trial in range(3):
        seeds = [trial * 4 + i for i in range(4)]
        serial_documents = [_nontrivial_document(seed) for seed in seeds]
        parallel_documents = [_nontrivial_document(seed) for seed in seeds]
        if trial % 2 == 0:
            serial_elapsed, serial_outputs = _timed(_render_serial, serial_documents)
            parallel_elapsed, parallel_outputs = _timed(
                _render_parallel, parallel_documents
            )
        else:
            parallel_elapsed, parallel_outputs = _timed(
                _render_parallel, parallel_documents
            )
            serial_elapsed, serial_outputs = _timed(_render_serial, serial_documents)

        assert all(
            output.rstrip().endswith(b"%%EOF")
            for output in serial_outputs + parallel_outputs
        )
        assert [_pdf_semantics(output, poppler) for output in parallel_outputs] == [
            _pdf_semantics(output, poppler) for output in serial_outputs
        ]
        serial_samples.append(serial_elapsed)
        parallel_samples.append(parallel_elapsed)

    serial_median = statistics.median(serial_samples)
    parallel_median = statistics.median(parallel_samples)
    assert parallel_median < serial_median, (
        f"parallel median {parallel_median:.3f}s was not lower than "
        f"serial median {serial_median:.3f}s"
    )


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

@pytest.mark.parametrize('method', ['to_pdf', 'render_page_to_png', 'render_page_to_svg', 'render_all_pages', 'render_pages', 'layout', 'layout_page'])
def test_fx195_installed_word_caller_selection_and_ordered_inputs(method, tmp_path):
    from rdocx import Document
    family, first, second = _fx195_font_inputs()
    document = Document()
    document.add_paragraph('iiii ' * 960)
    document.paragraphs[0].runs[0].font.name = family
    document.paragraphs[0].runs[0].font.size = 36 * 12700
    before = document.to_bytes()
    first_args, second_args = ({'fonts': [(family, first)]}, {'fonts': [(family, second)]})
    first_count = len(document.render_all_pages(dpi=72, **first_args))
    second_count = len(document.render_all_pages(dpi=72, **second_args))
    assert first_count != second_count, 'independent caller faces must discriminate actual pagination'
    page = min(first_count, second_count)

    def invoke(kwargs):
        if method == 'to_pdf':
            return document.to_pdf(**kwargs)
        if method == 'render_page_to_png':
            return document.render_page_to_png(0, dpi=72, **kwargs)
        if method == 'render_page_to_svg':
            return document.render_page_to_svg(0, **kwargs).svg
        if method == 'render_all_pages':
            return document.render_all_pages(dpi=72, **kwargs)
        if method == 'render_pages':
            return document.render_pages(dpi=72, pages=[0], **kwargs)
        if method == 'layout':
            return document.layout(**kwargs)
        return document.layout_page(page, **kwargs)
    reference, alternative = (invoke(first_args), invoke(second_args))
    assert reference != alternative, f'{method} must use the selected physical caller face'
    if method == 'to_pdf':
        assert any((b'FX195CX' in name for name in _pdf_base_fonts(reference)))
    if method == 'render_page_to_svg':
        import base64
        embedded = re.findall("data:font/[^;]+;base64,([^')]+)", reference)
        assert first in [base64.b64decode(value) for value in embedded]
    first_dir, second_dir, empty = (tmp_path / 'first', tmp_path / 'second', tmp_path / 'empty')
    for directory in (first_dir, second_dir, empty):
        directory.mkdir()
    (first_dir / (family + '.ttf')).write_bytes(first)
    (second_dir / (family + '.ttf')).write_bytes(second)
    assert invoke({'font_dir': first_dir}) == reference
    assert invoke({'font_dirs': [second_dir, first_dir]}) == reference
    assert invoke({'font_dirs': [first_dir, second_dir]}) == alternative
    assert invoke({'font_dirs': [second_dir], **first_args}) == reference
    assert invoke(first_args) == reference, 'same-family request must not retain the previous face'
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
        from rdocx import LayoutError
        with pytest.raises((PermissionError, LayoutError), match=re.escape(str(denied))) as failure:
            invoke({'font_dirs': [denied]})
        assert 'permission denied' in str(failure.value).lower()
    finally:
        denied.chmod(448)
    assert document.to_bytes() == before

def test_fx195_word_selected_pages_use_the_configured_bundle(tmp_path):
    from rdocx import Document
    family, first, second = _fx195_font_inputs()
    document = Document()
    document.add_paragraph('iiii ' * 960)
    document.paragraphs[0].runs[0].font.name = family
    document.paragraphs[0].runs[0].font.size = 36 * 12700
    before = document.to_bytes()
    configured = {'fonts': [(family, first)]}
    all_pages = document.render_all_pages(dpi=72, **configured)
    assert len(all_pages) > 1
    assert document.render_pages(dpi=72, **configured) == all_pages
    assert document.render_pages(dpi=72, pages=[1, 0], **configured) == [all_pages[1], all_pages[0]]
    assert len(document.render_all_pages(dpi=72, fonts=[(family, second)])) != len(all_pages)
    with pytest.raises(Exception):
        document.render_pages(dpi=72, pages=[], **configured)
    with pytest.raises(Exception):
        document.render_pages(dpi=72, pages=[len(all_pages)], **configured)
    from rdocx import LayoutError
    with pytest.raises(LayoutError, match='page index 1 was selected more than once'):
        document.render_pages(dpi=72, pages=[1, 0, 1], **configured)
    assert document.to_bytes() == before
    tiff = document.render_pages(dpi=72, format='tiff', pages=[1, 0], **configured)
    from PIL import Image
    image = Image.open(io.BytesIO(tiff))
    assert image.n_frames == 2
    for frame_index, page_index in enumerate((1, 0)):
        image.seek(frame_index)
        expected = Image.open(io.BytesIO(all_pages[page_index]))
        assert image.convert("RGBA").tobytes() == expected.convert("RGBA").tobytes()
    png = document.render_pages(dpi=72, transparent=True, pages=[0], **configured)
    assert Image.open(io.BytesIO(png[0])).mode == 'RGBA'
    jpeg = document.render_pages(dpi=72, format='jpeg', quality=80, pages=[0], **configured)
    assert Image.open(io.BytesIO(jpeg[0])).format == 'JPEG'
    with pytest.raises(Exception):
        document.render_pages(format='jpeg', quality=0, **configured)
    assert document.to_bytes() == before

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

def test_fx195_word_caller_png_matches_independent_font_ink():
    from rdocx import Document
    family, first, second = _fx195_font_inputs()
    paintings = []
    for source in (first, second):
        document = Document()
        document.add_paragraph('iiii')
        document.paragraphs[0].runs[0].font.name = family
        document.paragraphs[0].runs[0].font.size = 36 * 12700
        before = document.to_bytes()
        png = document.render_page_to_png(0, dpi=72, fonts=[(family, source)])
        expected = _fx195_four_i_ink_width(source, 36)
        assert abs(_fx195_painted_width(png) - expected) <= 2, 'painted width must match independent selected-font hmtx/glyf metrics'
        assert document.to_bytes() == before
        paintings.append(png)
    assert paintings[0] != paintings[1]

def test_fx195_word_collection_selects_actual_second_face(tmp_path):
    from rdocx import Document
    family, first, selected = _fx195_font_inputs()
    ordinary = (Path(__file__).resolve().parents[2] / 'oxml-layout' / 'fonts' / 'Carlito-Regular.ttf').read_bytes()
    collection = _fx195_collection(ordinary, selected)
    directory = tmp_path / 'collection'
    directory.mkdir()
    (directory / 'filename-alias.ttc').write_bytes(collection)
    document = Document()
    document.add_paragraph('iiii')
    document.paragraphs[0].runs[0].font.name = family
    document.paragraphs[0].runs[0].font.size = 36 * 12700
    before = document.to_bytes()
    direct = document.render_page_to_png(0, dpi=72, fonts=[(family, selected)])
    actual = document.render_page_to_png(0, dpi=72, font_dirs=[directory])
    assert actual == direct
    assert actual != document.render_page_to_png(0, dpi=72, fonts=[(family, first)])
    assert abs(_fx195_painted_width(actual) - _fx195_four_i_ink_width(selected, 36)) <= 2
    assert document.to_bytes() == before

def test_fx195_word_caller_precedes_real_embedded_font():
    import xml.etree.ElementTree as ET
    from rdocx import Document
    family, caller, embedded = _fx195_font_inputs()
    source = Document()
    source.add_paragraph('iiii')
    source.paragraphs[0].runs[0].font.name = family
    source.paragraphs[0].runs[0].font.size = 36 * 12700
    key = '{00112233-4455-6677-8899-AABBCCDDEEFF}'
    mask = bytes.fromhex(key.replace('{', '').replace('}', '').replace('-', ''))[::-1]
    obfuscated = bytearray(embedded)
    for index in range(32):
        obfuscated[index] ^= mask[index % 16]
    with zipfile.ZipFile(io.BytesIO(source.to_bytes())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    relns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    relations = ET.fromstring(parts['word/_rels/document.xml.rels'])
    font_table_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable'
    for relationship in list(relations):
        if relationship.attrib.get('Type') == font_table_type:
            relations.remove(relationship)
    ET.SubElement(relations, '{' + relns + '}Relationship', {'Id': 'fx195FontTable', 'Type': font_table_type, 'Target': 'fontTable.xml'})
    parts['word/_rels/document.xml.rels'] = ET.tostring(relations, encoding='utf-8', xml_declaration=True)
    parts['word/fontTable.xml'] = f'<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="{family}"><w:embedRegular r:id="fx195Embedded" w:fontKey="{key}"/></w:font></w:fonts>'.encode()
    parts['word/_rels/fontTable.xml.rels'] = f'<Relationships xmlns="{relns}"><Relationship Id="fx195Embedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/fx195.odttf"/></Relationships>'.encode()
    parts['word/fonts/fx195.odttf'] = bytes(obfuscated)
    ctns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    types = ET.fromstring(parts['[Content_Types].xml'])
    for name, content_type in (('/word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'), ('/word/fonts/fx195.odttf', 'application/vnd.openxmlformats-officedocument.obfuscatedFont')):
        for old in list(types):
            if old.attrib.get('PartName') == name:
                types.remove(old)
        ET.SubElement(types, '{' + ctns + '}Override', {'PartName': name, 'ContentType': content_type})
    parts['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8', xml_declaration=True)
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    document = Document.from_bytes(output.getvalue())
    before = document.to_bytes()
    embedded_png = source.render_page_to_png(0, dpi=72, fonts=[(family, embedded)])
    caller_png = source.render_page_to_png(0, dpi=72, fonts=[(family, caller)])
    assert embedded_png != caller_png
    assert document.render_page_to_png(0, dpi=72) == embedded_png
    assert document.render_page_to_png(0, dpi=72, fonts=[(family, caller)]) == caller_png
    assert document.to_pdf(fonts=[(family, caller)]) == source.to_pdf(fonts=[(family, caller)])
    assert document.to_bytes() == before


def test_fx195_word_empty_fonts_preserve_bundled_measurement_fallback(tmp_path):
    from rdocx import Document

    family, first, second = _fx195_font_inputs()
    document = Document()
    document.add_paragraph("iiii " * 960)
    document.paragraphs[0].runs[0].font.name = family
    document.paragraphs[0].runs[0].font.size = 36 * 12700
    before = document.to_bytes()
    ordinary = document.layout()
    ordinary_count = max(fragment.physical_page for fragment in ordinary)
    configured = [(face, document.layout(fonts=[(family, face)]))
                  for face in (first, second)]
    counts = [max(fragment.physical_page for fragment in layout)
              for _, layout in configured]
    assert counts[0] != counts[1], "distinct bundled faces must discriminate physical pagination"
    caller_face, normal = next((face, layout) for face, layout in configured
                              if max(fragment.physical_page for fragment in layout) != ordinary_count)
    normal_count = max(fragment.physical_page for fragment in normal)
    boundary = min(ordinary_count, normal_count)
    assert (document.layout_page(boundary) is None) == (boundary >= ordinary_count)
    assert (document.layout_page(boundary, fonts=[(family, caller_face)]) is None) == (boundary >= normal_count)
    ordinary_page = document.layout_page(0)
    # Page-zero metadata is font-independent. Actual font-sensitive page availability
    # is tested in the entrypoint matrix above; this control checks empty-input equality.
    empty = tmp_path / "empty"
    empty.mkdir()
    for inputs in ({"fonts": []}, {"font_dirs": []}, {"font_dir": empty}, {"font_dirs": [empty]}):
        assert document.layout(**inputs) == ordinary
        assert document.layout_page(0, **inputs) == ordinary_page
        assert document.layout_page(boundary, **inputs) == document.layout_page(boundary)
    # Empty explicit bytes cannot bypass validation of a provided invalid path.
    missing = tmp_path / "missing"
    with pytest.raises(FileNotFoundError, match=re.escape(str(missing))):
        document.layout(fonts=[], font_dirs=[missing])
    assert document.to_bytes() == before




def _fx195_variable_svg_owners(svg, weight, variable):
    import base64
    import hashlib
    import json
    import os
    import struct
    from pathlib import Path
    import xml.etree.ElementTree as ET

    authority = json.loads((Path(os.environ['F_X195_VARIABLE_ORACLES']) / 'receipt.json').read_text())
    row = next(row for row in authority['controls'] if row['script'] == 'Thai' and row['weight'] == weight)
    expected_sha = row['source_sha256'] if variable else row['output_sha256']
    tree = ET.fromstring(svg)
    ns = '{http://www.w3.org/2000/svg}'
    definitions = {}
    for style in tree.iter(ns + 'style'):
        for family, encoded in re.findall(r"@font-face\{font-family:'([^']+)';src:url\('data:[^;]+;base64,([^']+)'\)", style.text or ''):
            assert family not in definitions, 'one physical definition per emitted family'
            definitions[family] = base64.b64decode(encoded, validate=True)
    owners = list(tree.iter(ns + 'text'))
    assert len(owners) == 5, 'the authenticated Thai fixture has five positioned text owners'
    assert [''.join(node.itertext()) for node in owners] == ['ภาษา', 'ไทยยิน', 'ดี', 'ต้อน', 'รับ']
    assert ''.join(''.join(node.itertext()) for node in owners) == row['text']
    families = {node.get('font-family') for node in owners}
    assert len(families) == 1 and None not in families
    family = next(iter(families))
    assert family in definitions
    font = definitions[family]
    assert hashlib.sha256(font).hexdigest() == expected_sha, 'every positioned owner must select authenticated physical source'
    assert len(font) >= 12 and font[:4] in (b'\x00\x01\x00\x00', b'OTTO')
    count = struct.unpack_from('>H', font, 4)[0]
    assert 12 + count * 16 <= len(font)
    tables = {}
    for offset in range(12, 12 + count * 16, 16):
        tag = font[offset:offset + 4]
        assert tag not in tables
        position, length = struct.unpack_from('>II', font, offset + 8)
        assert position <= len(font) and length <= len(font) - position
        tables[tag] = (position, length)
    defaults = {}
    if variable:
        position, length = tables[b'fvar']
        assert length >= 16 and struct.unpack_from('>I', font, position)[0] == 0x00010000
        axis_offset, _, axis_count, axis_size = struct.unpack_from('>HHHH', font, position + 4)
        assert axis_offset >= 16 and axis_size >= 20
        assert axis_offset + axis_count * axis_size <= length
        for index in range(axis_count):
            axis = position + axis_offset + index * axis_size
            tag = font[axis:axis + 4].decode('ascii')
            assert tag not in defaults
            lower, default, upper = struct.unpack_from('>iii', font, axis + 4)
            assert lower <= default <= upper
            defaults[tag] = default / 65536
        assert defaults == {'wght': 400.0, 'wdth': 100.0}, 'implicit selection admitted only for authenticated fixture defaults'
    else:
        assert b'fvar' not in tables, 'independent reference is the selected static instance'
    evidence = []
    for node in owners:
        assert node.get('font-family') == family
        assert 'stroke' not in node.attrib and 'transform' not in node.attrib and 'filter' not in node.attrib
        assert node.get('font-style') == 'normal'
        assert node.get('font-weight') == ('bold' if weight == 700 else 'normal')
        assert node.get('textLength') is None, 'these positioned Thai owners require no textLength oracle correction'
        xs = [float(value) for value in node.get('x', '').split()]
        assert len(xs) == len(''.join(node.itertext())) and xs
        assert float(node.get('y')) > 0
        style = node.get('style')
        selectors = {}
        if variable:
            if style is None:
                assert weight == 400 and node.get('font-weight') == 'normal'
            else:
                assert style.startswith('font-synthesis:none;font-variation-settings:')
                payload = style.removeprefix('font-synthesis:none;font-variation-settings:')
                for declaration in payload.split(','):
                    match = re.fullmatch(r"'([A-Za-z0-9]{4})' ([-+0-9.]+)", declaration.strip())
                    assert match is not None, 'no unqualified CSS or synthetic operation'
                    tag, value = match.groups()
                    assert tag not in selectors
                    selectors[tag] = float(value)
                assert selectors and set(selectors) <= {'wght', 'wdth'}
            effective = defaults | selectors
            assert effective == {'wght': float(weight), 'wdth': 100.0}
            if weight == 700:
                assert selectors.get('wght') == 700, 'nondefault selection must be explicit on every owner'
        else:
            assert style is None
        evidence.append({'text': ''.join(node.itertext()), 'family': family, 'physicalFontSHA': expected_sha,
                         'explicitCoordinates': selectors, 'effectiveCoordinates': defaults | selectors if variable else {},
                         'implicitDefault': variable and style is None})
    return tree, owners, family, evidence


def _fx195_variable_svg_geometry(svg, weight, variable):
    import xml.etree.ElementTree as ET
    tree, owners, _, _ = _fx195_variable_svg_owners(svg, weight, variable)
    for node in owners:
        node.attrib.pop('style', None)
    # Authenticate physical source and every owner's effective coordinates first.
    # Remove only font-definition/selection representation differences. Preserve
    # all five owners, positions, glyph grouping, text, order and other attributes.
    for parent in tree.iter():
        for child in list(parent):
            if child.tag == '{http://www.w3.org/2000/svg}style':
                assert child.text and '@font-face{' in child.text
                parent.remove(child)
    # C14N rewrites prefixes locally, independent of ElementTree's global registry.
    # Keep expanded names, every remaining attribute, text/tails and child order.
    return ET.canonicalize(ET.tostring(tree, encoding='unicode'), rewrite_prefixes=True)


def _fx195_variable_path_references(expected_owners,row,dimensions,directory):
    # Paint at the already compared static-owner positions. This does not
    # promote those positions into a new independent absolute-placement oracle.
    import hashlib, json
    from pathlib import Path
    import xml.etree.ElementTree as ET
    source=Path(row['output']).read_bytes()
    assert hashlib.sha256(source).hexdigest()==row['output_sha256']
    assert ''.join(''.join(node.itertext()) for node in expected_owners)==row['text']
    glyphs=row['shape']
    # This exact retained Thai control has one emitted glyph per scalar. Refuse
    # other cardinalities rather than invent a scalar-to-shaped-glyph mapping.
    assert len(glyphs)==len(row['text'])
    xs=[];ys=[]
    for node in expected_owners:
        text=''.join(node.itertext());positions=[float(x) for x in node.get('x','').split()]
        assert len(positions)==len(text) and float(node.get('font-size'))==24.
        xs.extend(positions);ys.extend([float(node.get('y'))]*len(text))
    assert len(xs)==len(glyphs) and len(set(ys))==1
    factor=24/row['units_per_em'];cursor=0.;origin=xs[0]
    # Static searchable SVG x contains the scalar pen before GPOS. Qualify
    # that representation separately, then apply independent HB offsets once.
    clusters=[glyph['cl'] for glyph in glyphs]
    assert clusters==sorted(clusters) and clusters[0]==0
    assert all(0<=cluster<len(row['text']) for cluster in clusters)
    assert len(expected_owners)==5
    geometry=[];position_proof=[]
    for index,glyph in enumerate(glyphs):
        pen=origin+cursor*factor
        assert abs(xs[index]-pen)<=.000001, 'static scalar pens must match retained independent HB advances before GPOS'
        x=pen+glyph['dx']*factor;y=ys[index]-glyph['dy']*factor
        record=_fx195_source_glyph(source,'',glyph_id=glyph['g'])
        assert record['upem']==row['units_per_em']
        geometry.append((record,x,y))
        position_proof.append({'glyph':glyph['g'],'cluster':glyph['cl'],'scalarIndex':index,'scalar':row['text'][index],'pen':[pen,ys[index]],'glyphOrigin':[x,y],'dx':glyph['dx'],'dy':glyph['dy'],'advance':glyph['ax']})
        cursor+=glyph['ax']
    assert cursor==row['total_x_advance']
    references={};proofs={}
    for label,cubic in [('Quadratic',False),('Cubic',True)]:
        ns='http://www.w3.org/2000/svg';ET.register_namespace('',ns)
        tree=ET.Element('{'+ns+'}svg',{'width':str(int(dimensions[0])),'height':str(int(dimensions[1])),'viewBox':f'0 0 {int(dimensions[0])} {int(dimensions[1])}'})
        # One compound path retains every contour and fits the admitted painter
        # path bound without changing the independently authored glyph topology.
        commands=[]
        for record,x,y in geometry:
            commands.extend(_fx195_source_path_commands(contour,lambda p:(x+p[0]*factor,y-p[1]*factor),cubic) for contour in record['contours'])
        ET.SubElement(tree,'{'+ns+'}path',{'d':' '.join(commands),'fill':'#000000','fill-rule':'nonzero'})
        raw=ET.tostring(tree,encoding='utf-8',xml_declaration=True)
        assert not any(token in raw for token in [b'<text',b'<image',b'<use',b'<defs',b'font',b'href',b'url(',b'<filter'])
        image,q,result=_fx195_render_source_paths(raw)
        path=directory/('variable-independent-'+label.lower()+'.png');image.save(path)
        references['expected'+label]=str(path);references['expected'+label+'SHA']=hashlib.sha256(path.read_bytes()).hexdigest()
        proofs[label]={'svgSHA':hashlib.sha256(raw).hexdigest(),'sourceSHA':row['output_sha256'],'glyphs':[{k:v for k,v in record.items() if k!='contours'} for record,_,_ in geometry],'rendererSHA':q['renderer_binary_sha256'],'actualExit':result.returncode,'positionAuthority':'static scalar pens checked before GPOS, independent retained HB14.4 offsets applied once','positions':position_proof,'completeSourceText':row['text'],'ownerTexts':[''.join(node.itertext()) for node in expected_owners],'curve':'cubic' if cubic else 'quadratic'}
    proof=directory/'variable-independent-paths.json';proof.write_text(json.dumps(proofs,indent=2)+'\n')
    references.update(pathProof=str(proof),pathProofSHA=hashlib.sha256(proof.read_bytes()).hexdigest())
    return references


def _fx195_save_variable_svg_artifact(actual, expected, weight, directory):
    import hashlib
    import json
    import os
    from pathlib import Path
    authority = json.loads((Path(os.environ['F_X195_VARIABLE_ORACLES']) / 'receipt.json').read_text())
    row = next(row for row in authority['controls'] if row['script'] == 'Thai' and row['weight'] == weight)
    root, owners, family, evidence = _fx195_variable_svg_owners(actual, weight, True)
    _, expected_owners, _, expected_evidence = _fx195_variable_svg_owners(expected, weight, False)
    assert len(owners) == len(expected_owners) == 5
    assert [item['text'] for item in evidence] == [item['text'] for item in expected_evidence]
    # Validate the last owner too. No global CSS substring or first-owner proof
    # can admit the wrong physical source or effective coordinates elsewhere.
    import xml.etree.ElementTree as ET
    for defect in ('physical-family', 'coordinates', 'missing-owner'):
        wrong = ET.fromstring(actual)
        wrong_owners = list(wrong.iter('{http://www.w3.org/2000/svg}text'))
        last = wrong_owners[-1]
        if defect == 'physical-family':
            last.set('font-family', 'unadmitted-physical-face')
        elif defect == 'coordinates':
            last.set('style', "font-synthesis:none;font-variation-settings:'wght' " + str(700 if weight == 400 else 400))
        else:
            next(parent for parent in wrong.iter() if last in list(parent)).remove(last)
        with pytest.raises(AssertionError):
            _fx195_variable_svg_owners(ET.tostring(wrong, encoding='unicode'), weight, True)
    if weight == 700:
        wrong = ET.fromstring(actual)
        last = list(wrong.iter('{http://www.w3.org/2000/svg}text'))[-1]
        last.attrib.pop('style')
        with pytest.raises(AssertionError):
            _fx195_variable_svg_owners(ET.tostring(wrong, encoding='unicode'), weight, True)
    a, e = directory / 'variable-installed.svg', directory / 'variable-independent-static.svg'
    a.write_text(actual)
    e.write_text(expected)
    dimensions = [float(root.get(key).removesuffix('pt')) for key in ('width', 'height')]
    assert dimensions == [round(value) for value in dimensions], '72dpi companion requires integer page dimensions'
    artifact = {'case': 'variable-' + str(weight), 'svg': str(a), 'svgSHA': hashlib.sha256(a.read_bytes()).hexdigest(),
                'expectedSVG': str(e), 'expectedSVGSHA': hashlib.sha256(e.read_bytes()).hexdigest(),
                'fontFamily': family, 'fontSHA': row['source_sha256'], 'fontIndex': 0,
                'oracleFont': row['output'], 'oracleFontSHA': row['output_sha256'],
                'width': int(dimensions[0]), 'height': int(dimensions[1]), 'threshold': .99,
                'selector': "'wght' " + str(weight), 'owners': evidence,
                'authenticatedFvarDefaults': {'wght': 400.0, 'wdth': 100.0}}
    artifact.update(_fx195_variable_path_references(expected_owners,row,dimensions,directory))
    # Preserve emitted SVG unchanged. Serial native stage checks every owner,
    # then admits only an independently selected physical static font resolver.
    (directory / 'variable-installed-svg-artifact.json').write_text(json.dumps(artifact, indent=2) + '\n')
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


@pytest.mark.parametrize('weight', [400, 700])
@pytest.mark.parametrize('method', ['to_pdf', 'render_page_to_png', 'render_page_to_svg', 'render_all_pages', 'render_pages', 'layout', 'layout_page'])
def test_fx195_installed_word_variable_instances_match_independent_static(method, weight, tmp_path):
    from rdocx import Document
    oracles = _fx195_style_oracles()
    variable, static, text = oracles[weight]
    family = 'FX195CallerThai'
    document = Document()
    document.add_paragraph(text)
    font = document.paragraphs[0].runs[0].font
    font.name = family
    font.size = 24 * 12700
    font.bold = weight == 700
    before = document.to_bytes()

    def invoke(data):
        kwargs = {'fonts': [(family, data)]}
        if method == 'to_pdf':
            return document.to_pdf(**kwargs)
        if method == 'render_page_to_png':
            return document.render_page_to_png(0, dpi=72, **kwargs)
        if method == 'render_page_to_svg':
            return document.render_page_to_svg(0, **kwargs).svg
        if method == 'render_all_pages':
            return document.render_all_pages(dpi=72, **kwargs)
        if method == 'render_pages':
            return document.render_pages(dpi=72, pages=[0], **kwargs)
        if method == 'layout_page':
            return document.layout_page(0, **kwargs)
        return document.layout(**kwargs)

    expected_value = invoke(static)
    actual_value = invoke(variable)
    if method == 'render_page_to_svg':
        expected = _fx195_variable_svg_geometry(expected_value,weight,False)
        actual = _fx195_variable_svg_geometry(actual_value,weight,True)
    else:
        expected = _fx195_style_observation(expected_value, method, tmp_path, 'static')
        actual = _fx195_style_observation(actual_value, method, tmp_path, 'variable')
    if method == 'render_page_to_svg':
        _fx195_save_variable_svg_artifact(actual_value, expected_value, weight, tmp_path)
    assert actual == expected, 'selected variable coordinates must match independently instantiated static geometry'
    repeated = invoke(variable)
    repeated = _fx195_variable_svg_geometry(repeated,weight,True) if method == 'render_page_to_svg' else _fx195_style_observation(repeated,method,tmp_path,'repeat')
    assert repeated == actual
    # A genuinely different weight control must paint differently even when
    # the selected installed entrypoint itself returns only measurement.
    requested = document.render_page_to_png(0, dpi=72, fonts=[(family, variable)])
    font.bold = weight != 700
    opposite = document.render_page_to_png(0, dpi=72, fonts=[(family, variable)])
    assert requested != opposite
    font.bold = weight == 700
    assert document.to_bytes() == before



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

def _fx195_source_glyph(source,char,*,glyph_id=None):
 import struct,hashlib
 def u16(at):return struct.unpack_from('>H',source,at)[0]
 tables={source[o:o+4]:struct.unpack_from('>II',source,o+8) for o in range(12,12+16*u16(4),16)}
 assert source[:4]==b'\x00\x01\x00\x00'
 assert all(o+n<=len(source) for o,n in tables.values())
 g=glyph_id
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



def _fx195_render_source_paths(raw):
    # Independent source authoring callers only, no product text or contour API.
    import hashlib, json, os, subprocess, tempfile
    from pathlib import Path
    from PIL import Image
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
    return image,q,result


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
    advance = record['advance'] * factor
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
    image,q,result = _fx195_render_source_paths(raw)
    qualification_path = Path(os.environ['F_X195_PATH_REFERENCE_QUALIFICATION'])
    record.update({'source_sha256':hashlib.sha256(source).hexdigest(),'svg_sha256':hashlib.sha256(raw).hexdigest(),
                   'origin':list(origin),'scale':font_scale,'roi':list(roi),'negative':negative,'cubic':cubic,
                   'qualification':str(qualification_path),'qualificationSHA':os.environ['F_X195_PATH_REFERENCE_QUALIFICATION_SHA'],
                   'rendererSHA':q['renderer_binary_sha256'],'stdout':result.stdout.decode(),'actualExit':result.returncode})
    return image,roi,len(case['text'])*advance,record


def _fx195_style_fixture_owners(before, case, page_size, notes=False):
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
        return [{'owner':'body','origin':(int(margin.get(w+'left'))/20,int(margin.get(w+'top'))/20+48*ascent/upem),'scale':1.}]
    p='{http://schemas.openxmlformats.org/presentationml/2006/main}';a='{http://schemas.openxmlformats.org/drawingml/2006/main}'
    slide_names=[name for name in parts if re.fullmatch(r'ppt/slides/slide\d+\.xml',name)];assert len(slide_names)==1
    slide=ET.fromstring(parts[slide_names[0]]);shapes=[shape for shape in slide.iter(p+'sp') if ''.join(n.text or '' for n in shape.iter(a+'t'))==case['text']];assert len(shapes)==1
    presentation=ET.fromstring(parts['ppt/presentation.xml']);size=presentation.find(p+'sldSz');assert size is not None
    slide_size=tuple(float(size.get(key))/12700 for key in ['cx','cy'])
    asc,desc=struct.unpack_from('>hh',source,tables[b'hhea'][0]+4);height=57.6;descent=-desc*48/upem
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
@pytest.mark.parametrize('method', ['to_pdf','render_page_to_png','render_page_to_svg','render_all_pages','render_pages','layout','layout_page'])
def test_fx195_installed_word_styles_match_independent_paint(method, style_case, tmp_path):
    from rdocx import Document
    case = _fx195_style_case(style_case)
    document = Document()
    document.add_paragraph(case['text'])
    font = document.paragraphs[0].runs[0].font
    font.name = case['family']
    font.size = 48 * 12700
    font.bold, font.italic = case['bold'],case['italic']
    before = document.to_bytes()
    kwargs = {'fonts':[(case['family'],data) for data in case['files']]}
    def invoke():
        if method == 'render_page_to_png': return document.render_page_to_png(0,dpi=72,**kwargs)
        if method == 'render_page_to_svg': return document.render_page_to_svg(0,**kwargs).svg
        if method == 'render_all_pages': return document.render_all_pages(dpi=72,**kwargs)
        if method == 'render_pages': return document.render_pages(pages=[0],dpi=72,**kwargs)
        if method == 'layout_page': return document.layout_page(0,**kwargs)
        return getattr(document,method)(**kwargs)
    _fx195_assert_installed_style(case,method,invoke,lambda:document.to_pdf(**kwargs),lambda:document.render_page_to_png(0,dpi=72,**kwargs),before,document.to_bytes,tmp_path)


def test_fx195_svg_geometry_normalizer_keeps_semantics_without_global_prefix_state(monkeypatch):
    import xml.etree.ElementTree as ET
    # This tests only normalization. Physical-font/coordinate authentication remains
    # in every real installed caller, before this tree reaches the normalizer.
    monkeypatch.setattr(sys.modules[__name__], '_fx195_variable_svg_owners',
                        lambda svg, weight, variable: (ET.fromstring(svg), [], None, None))
    default = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 30"><g data-owner="first"><text x="2 4" y="7" xml:space="preserve"> A </text> tail </g><g data-owner="second"/></svg>'
    prefixed = '<s:svg xmlns:s="http://www.w3.org/2000/svg" viewBox="0 0 20 30"><s:g data-owner="first"><s:text x="2 4" y="7" xml:space="preserve"> A </s:text> tail </s:g><s:g data-owner="second"/></s:svg>'
    def normalized(svg):
        return _fx195_variable_svg_geometry(svg, 400, False)
    assert normalized(default) == normalized(prefixed)
    # Every semantic change must survive normalization and fail equality.
    changes = [default.replace('2 4', '2 5'), default.replace('y="7"', 'y="8"'),
               default.replace('0 0 20 30', '0 0 21 30'), default.replace('> A <', '> B <'),
               default.replace(' tail ', ' different tail '), default.replace('preserve', 'default'),
               default.replace('data-owner="first"', 'data-owner="changed"'),
               default.replace('http://www.w3.org/2000/svg', 'urn:foreign')]
    tree = ET.fromstring(default)
    tree[:] = list(reversed(list(tree)))
    changes.append(ET.tostring(tree, encoding='unicode'))
    for changed in changes:
        assert normalized(changed) != normalized(default)


def test_fx210_word_companions_preserve_payload_and_owned_messages():
    from rdocx import Document

    document = Document()
    document.add_paragraph("").add_run("Actual contributing text").font.name = "Aptos"
    source = (Path(__file__).resolve().parents[2] / "oxml-layout/fonts/Carlito-Regular.ttf").read_bytes()
    kwargs = {"fonts": [("Aptos", source)]}
    pdf, facts = document.to_pdf_with_diagnostics(**kwargs)
    assert pdf == document.to_pdf(**kwargs)
    assert isinstance(facts, tuple) and all(isinstance(value, str) for value in facts)
    assert sum("Font 'Aptos' is approximate" in value for value in facts) == 1
    png, png_facts = document.render_page_to_png_with_diagnostics(0, 72.0, **kwargs)
    assert png == document.render_page_to_png(0, 72.0, **kwargs)
    assert png_facts == facts
    pages, page_facts = document.render_pages_with_diagnostics(dpi=72.0, **kwargs)
    assert pages == document.render_pages(dpi=72.0, **kwargs)
    assert page_facts == facts
    from rdocx import LayoutError

    before_empty = document.to_bytes()
    with pytest.raises(LayoutError, match="^raster export error: at least one raster page is required$"):
        document.render_pages_with_diagnostics(pages=[], dpi=72.0, **kwargs)
    assert document.to_bytes() == before_empty
    tiff, _ = document.render_pages_with_diagnostics(format="tiff", dpi=72.0, **kwargs)
    assert isinstance(tiff, bytes)



@pytest.mark.parametrize("owner", ["symbol", "numbering"])
def test_fx211c_word_unknown_symbol_source_and_same_pass_diagnostics(owner, tmp_path):
    import xml.etree.ElementTree as ET
    from rdocx import Document, ListLevel
    document = Document()
    document.add_paragraph("Visible symbol owner")
    if owner == "numbering":
        definition = document.add_numbering_definition([ListLevel(format="bullet", text="\uf0ff")])
        instance = document.add_numbering_instance(definition)
        document.paragraphs[0].numbering = (instance, 0)
    source = zipfile.ZipFile(io.BytesIO(document.to_bytes()))
    parts = {name: source.read(name) for name in source.namelist()}
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    if owner == "symbol":
        tree = ET.fromstring(parts["word/document.xml"])
        run = next(tree.iter(ns + "r"))
        ET.SubElement(run, ns + "sym", {ns + "font": "Wingdings", ns + "char": "F0FF"})
        parts["word/document.xml"] = ET.tostring(tree, encoding="utf-8")
    else:
        tree = ET.fromstring(parts["word/numbering.xml"])
        level = next(node for node in tree.iter(ns + "lvl") if node.find(ns + "lvlText").get(ns + "val") == "\uf0ff")
        properties = level.find(ns + "rPr")
        if properties is None:
            properties = ET.SubElement(level, ns + "rPr")
        fonts = properties.find(ns + "rFonts")
        if fonts is None:
            fonts = ET.SubElement(properties, ns + "rFonts")
        fonts.set(ns + "ascii", "Wingdings")
        fonts.set(ns + "hAnsi", "Wingdings")
        parts["word/numbering.xml"] = ET.tostring(tree, encoding="utf-8")
    data = io.BytesIO()
    with zipfile.ZipFile(data, "w") as package:
        for name, body in parts.items():
            package.writestr(name, body)
    document = Document.from_bytes(data.getvalue())
    before = document.to_bytes()
    font = (Path(__file__).resolve().parents[2] / "oxml-layout/fonts/Carlito-Regular.ttf").read_bytes()
    kwargs = {"fonts": [("Wingdings", font), ("Calibri", font)]}
    pdf, facts = document.to_pdf_with_diagnostics(**kwargs)
    (tmp_path / (owner + "-facts.txt")).write_text(repr(facts))
    assert pdf.startswith(b"%PDF-") and document.to_bytes() == before
    expected = "Unknown Wingdings symbol U+F0FF: original code retained"
    assert facts.count(expected) == 1, facts
    assert document.to_pdf_with_diagnostics(**kwargs)[1] == facts
    reopened = Document.from_bytes(before)
    saved = zipfile.ZipFile(io.BytesIO(reopened.to_bytes()))
    if owner == "symbol":
        symbol = ET.fromstring(saved.read("word/document.xml")).find(".//" + ns + "sym")
        assert symbol.get(ns + "font") == "Wingdings" and symbol.get(ns + "char") == "F0FF"
    else:
        levels = ET.fromstring(saved.read("word/numbering.xml")).iter(ns + "lvl")
        assert any(level.find(ns + "lvlText").get(ns + "val") == "\uf0ff" for level in levels)
    assert reopened.to_pdf_with_diagnostics(**kwargs)[1] == facts


# CI preparation runs explicitly before pytest. Both binding families consume
# these independent controls. Importing this module never prepares or compiles.
def _fx195_prepare_inputs(directory, family=None, harfbuzz_archive=None, native_tools_only=False):
    import hashlib
    import json
    import os
    from pathlib import Path
    import shutil
    import stat
    import subprocess
    import tarfile
    import tomllib
    import urllib.request
    import xml.etree.ElementTree as ET

    root = Path(__file__).resolve().parents[3]
    directory = Path(directory).resolve()
    directory.mkdir(parents=True, exist_ok=False)
    digest = lambda p: hashlib.sha256(Path(p).read_bytes()).hexdigest()
    commands = []

    def run(argv, timeout=120):
        index = len(commands)
        result = subprocess.run([str(a) for a in argv], cwd=root,
                                capture_output=True, timeout=timeout, check=False)
        stdout, stderr = directory / f'command-{index}.stdout', directory / f'command-{index}.stderr'
        stdout.write_bytes(result.stdout)
        stderr.write_bytes(result.stderr)
        commands.append({'argv': [str(a) for a in argv], 'exit': result.returncode,
                         'stdout': str(stdout), 'stdoutSHA256': digest(stdout),
                         'stderr': str(stderr), 'stderrSHA256': digest(stderr)})
        (directory / 'commands.json').write_text(json.dumps(commands, indent=2) + '\n')
        if result.returncode:
            raise RuntimeError(f'font preparation failed: {argv}: {result.stderr.decode(errors="replace")}')
        return result.stdout

    def source_guard():
        files = subprocess.check_output(['git', 'ls-files', '-z'], cwd=root).decode().split('\0')
        return {p: {'sha256': digest(root / p), 'mode': stat.S_IMODE((root / p).stat().st_mode)}
                for p in files if p}

    before = source_guard()
    archive_sha = '2357ed966c6ced7bfa720b0640c0231065af01158fbea215093ffa15aed44371'
    archive_url = 'https://github.com/harfbuzz/harfbuzz/releases/download/14.4.0/harfbuzz-14.4.0.tar.xz'
    archive = directory / 'harfbuzz-14.4.0.tar.xz'
    if harfbuzz_archive is not None:
        shutil.copyfile(harfbuzz_archive, archive)
    else:
        with urllib.request.urlopen(archive_url, timeout=120) as response, archive.open('wb') as out:
            size = 0
            while chunk := response.read(1024 * 1024):
                size += len(chunk)
                if size > 32 * 1024 * 1024:
                    raise RuntimeError('HarfBuzz source archive exceeds its bound')
                out.write(chunk)
    assert digest(archive) == archive_sha
    source = directory / 'harfbuzz-source'
    source.mkdir()
    total, count = 0, 0
    with tarfile.open(archive, 'r|xz') as tar:
        for member in tar:
            count += 1
            total += member.size
            if count > 4096 or total > 128 * 1024 * 1024 or member.size < 0:
                raise RuntimeError('HarfBuzz source inventory exceeds its bounds')
            target = source / member.name
            target.resolve().relative_to(source)
            if member.issym() and member.name == 'harfbuzz-14.4.0/CLAUDE.md' and member.linkname == 'AGENTS.md':
                continue  # Reviewed documentation link, not a build input.
            if member.isdir():
                target.mkdir(parents=True, exist_ok=True)
            elif member.isfile():
                target.parent.mkdir(parents=True, exist_ok=True)
                with tar.extractfile(member) as data, target.open('wb') as out:
                    shutil.copyfileobj(data, out)
                target.chmod(stat.S_IMODE(member.mode))
            else:
                raise RuntimeError('unexpected HarfBuzz source entry')
    build = directory / 'harfbuzz-build'
    run(['meson', 'setup', build, source / 'harfbuzz-14.4.0', '--buildtype=release',
         '--default-library=static', '-Dtests=disabled', '-Ddocs=disabled',
         '-Dintrospection=disabled', '-Dicu=disabled', '-Dgraphite=disabled',
         '-Dcairo=enabled', '-Dfreetype=enabled', '-Dglib=enabled',
         '-Dgobject=disabled', '-Dcoretext=disabled', '-Dutilities=enabled'], 300)
    run(['meson', 'compile', '-C', build, '-j', '4', 'hb-shape', 'hb-subset', 'hb-view'], 1200)
    tools = {name: build / 'util' / name for name in ('hb-shape', 'hb-subset', 'hb-view')}
    identities = {}
    for name, tool in tools.items():
        identities[name] = run([tool, '--version']).decode().splitlines()[0]
        assert identities[name] == name + ' (HarfBuzz) 14.4.0'
    tool_guards = {str(p): digest(p) for p in tools.values()}

    if native_tools_only:
        assert source_guard() == before
        assert all(digest(p) == sha for p, sha in tool_guards.items())
        receipt = {'status': 'actual-independent-native-tools-prepared',
                   'sourceGuardsBeforeAfter': True, 'sourceInputs': before,
                   'commands': commands, 'harfbuzzSourceURL': archive_url,
                   'harfbuzzSourceSHA256': archive_sha, 'harfbuzzTools': tool_guards,
                   'toolVersions': identities}
        (directory / 'receipt.json').write_text(json.dumps(receipt, indent=2) + '\n')
        if os.environ.get('GITHUB_PATH'):
            with Path(os.environ['GITHUB_PATH']).open('a') as out:
                out.write(str(build / 'util') + '\n')
        print(json.dumps(receipt, indent=2))
        return receipt

    fonts = root / 'crates/oxml-layout/fonts'
    pinned_fonts = {
        'Carlito-Regular.ttf': 'b4ff23ba370cc95a3c349336b73f9c28514a1371210f89832efc85c4b1ea7131',
        'Carlito-Italic.ttf': '718a0663864d37a4868220a19b9668a5fe10a46197f6df367b4c2c30c04c026c',
        'NotoSansThai.ttf': '5a1c559bb539583c8a1fd99d1c5b9491e5e14478c9cd2bd0970d5c3096cc9ef8',
        'LICENSE-Carlito': 'bbcf8ce9c8acc91355ab5d263ec9e37d498de2222336e7d42e589f791afe56c4',
        'LICENSE-Noto': 'df5143cdf3380169f2d03bf6d2cd243e85621fd097e608c3607cf7f9c9884ae6',
    }
    assert all(digest(fonts / name) == sha for name, sha in pinned_fonts.items())
    guards = directory / 'style-font-guards.json'
    guards.write_text(json.dumps(pinned_fonts, indent=2) + '\n')

    def tables(data):
        return {data[o:o + 4]: data[off:off + length]
                for o in range(12, 12 + 16 * struct.unpack_from('>H', data, 4)[0], 16)
                for off, length in [struct.unpack_from('>II', data, o + 8)]}

    def default_static(data):
        # Keep the original default outline and metrics, remove only its axes.
        selected = {k: v for k, v in tables(data).items() if k not in (b'fvar', b'gvar')}
        selected[b'head'] = selected[b'head'][:8] + b'\0' * 4 + selected[b'head'][12:]
        count = len(selected)
        power = 1 << (count.bit_length() - 1)
        result = bytearray(struct.pack('>IHHHH', 65536, count, 16 * power, power.bit_length() - 1, 16 * (count - power)))
        result.extend(b'\0' * (16 * count))
        head = None
        for index, (tag, value) in enumerate(sorted(selected.items())):
            offset = len(result)
            struct.pack_into('>4sIII', result, 12 + index * 16, tag, _fx195_checksum(value), offset, len(value))
            result.extend(value + b'\0' * (-len(value) % 4))
            if tag == b'head':
                head = offset
        struct.pack_into('>I', result, head + 8, (0xB1B0AFBA - _fx195_checksum(result)) & 0xFFFFFFFF)
        assert _fx195_checksum(result) == 0xB1B0AFBA
        return bytes(result)

    def glyph_one_points(data):
        ts = tables(data)
        long_loca = struct.unpack_from('>h', ts[b'head'], 50)[0]
        offset = struct.unpack_from('>I' if long_loca else '>H', ts[b'loca'], 4 if long_loca else 2)[0]
        if not long_loca:
            offset *= 2
        raw = ts[b'glyf'][offset:]
        assert struct.unpack_from('>h', raw)[0] == 1
        count = struct.unpack_from('>H', raw, 10)[0] + 1
        at = 14 + struct.unpack_from('>H', raw, 12)[0]
        flags = []
        while len(flags) < count:
            flag = raw[at]
            at += 1
            repetitions = raw[at] if flag & 8 else 0
            at += bool(flag & 8)
            flags.extend([flag] * (repetitions + 1))
        coordinates = []
        for short, same in [(2, 16), (4, 32)]:
            values, current = [], 0
            for flag in flags:
                if flag & short:
                    delta = raw[at] if flag & same else -raw[at]
                    at += 1
                elif flag & same:
                    delta = 0
                else:
                    delta = struct.unpack_from('>h', raw, at)[0]
                    at += 2
                current += delta
                values.append(current)
            coordinates.append(values)
        return [list(p) for p in zip(*coordinates)]

    axis = directory / 'axis-oracles'
    axis.mkdir()
    fixture_source = (root / 'crates/oxml-pdf/src/font.rs').read_text()
    controls = []
    original_axis = None
    for name in ('ital', 'slnt', 'none'):
        if name != 'none':
            section = fixture_source.split(f'fn fx195_{name}_axis_causal_selected_source_and_embedded_outline()', 1)[1]
            arrays = re.findall(r'&\[([\d,\s]+)\]', section)[:2]
            assert len(arrays) == 2
            original, independent_static = [bytes(int(n) for n in re.findall(r'\d+', a)) for a in arrays]
            if name == 'ital':
                original_axis = original
            assert glyph_one_points(independent_static) == [[50, 0], [450, 0], [590, 700], [190, 700]]
        else:
            original = default_static(original_axis)
        input_path = axis / (name + '.ttf')
        input_path.write_bytes(original)
        coords = [f'--variations={name}={1 if name == "ital" else -12}'] if name != 'none' else []
        shaped = run([tools['hb-shape'], input_path, 'A', '--font-size=upem', '--no-glyph-names', '--output-format=json', *coords])
        assert json.loads(shaped) == [{'g': 1, 'cl': 0, 'dx': 0, 'dy': 0, 'ax': 600, 'ay': 0}]
        row = {'axis': name, 'input_sha256': digest(input_path), 'exit_code': 0, 'stdout': shaped.decode()}
        if name != 'none':
            static = axis / (name + '-static.ttf')
            run([tools['hb-subset'], input_path, '--unicodes=*', '--name-IDs=*', '--output-file=' + str(static), *coords])
            points = glyph_one_points(static.read_bytes())
            assert points == [[50, 0], [450, 0], [590, 700], [190, 700]]
            assert b'fvar' not in tables(static.read_bytes())
            row['static'] = {'sha256': digest(static), 'glyph1_points': points, 'fvar_absent': True}
        controls.append(row)
    (axis / 'receipt.json').write_text(json.dumps({'controls': controls, 'tools': identities}, indent=2) + '\n')

    def svg_facts(path):
        node = ET.fromstring(path.read_bytes())
        definitions = {e.get('id'): e for e in node.iter() if e.get('id')}
        href = '{http://www.w3.org/1999/xlink}href'
        token = re.compile(r'[A-DF-Za-df-z]|[-+]?(?:\d*\.?\d+)(?:[eE][-+]?\d+)?')
        uses = []
        for use in node.iter():
            if use.tag.endswith('}use'):
                paths = [token.findall(e.get('d', '')) for e in definitions[use.get(href)[1:]].iter() if e.tag.endswith('}path')]
                uses.append(({k: v for k, v in use.attrib.items() if k != href}, paths))
        return node.attrib, uses

    variable = directory / 'variable-oracles'
    variable.mkdir()
    rows = []
    for weight in (400, 700):
        source_font = fonts / 'NotoSansThai.ttf'
        static = variable / f'Thai-{weight}-static.ttf'
        coords = f'--variations=wght={weight},wdth=100'
        run([tools['hb-subset'], source_font, '--unicodes=*', '--name-IDs=*', coords, '--output-file=' + str(static)])
        shapes, paintings = [], []
        for label, font in [('source', source_font), ('static', static)]:
            variations = [coords] if label == 'source' else []
            shape = json.loads(run([tools['hb-shape'], font, 'ภาษาไทยยินดีต้อนรับ', '--font-size=upem', '--no-glyph-names', '--output-format=json', *variations]))
            shapes.append(shape)
            svg = variable / f'Thai-{weight}-{label}.svg'
            run([tools['hb-view'], font, 'ภาษาไทยยินดีต้อนรับ', '--font-size=1000', '--margin=0', '--output-format=svg', '--output-file=' + str(svg), *variations])
            paintings.append(svg_facts(svg))
        positions = lambda shape: [{k: v for k, v in glyph.items() if k != 'g'} for glyph in shape]
        assert positions(shapes[0]) == positions(shapes[1])
        left, right = paintings
        assert left[0] == right[0] and len(left[1]) == len(right[1]) > 0
        delta, coordinates = 0., 0
        for (lp, lpaths), (rp, rpaths) in zip(left[1], right[1]):
            assert lp == rp and len(lpaths) == len(rpaths)
            for lpath, rpath in zip(lpaths, rpaths):
                assert len(lpath) == len(rpath)
                for l, r in zip(lpath, rpath):
                    try:
                        lf, rf = float(l), float(r)
                    except ValueError:
                        assert l == r
                    else:
                        delta = max(delta, abs(lf - rf))
                        coordinates += 1
        assert coordinates > 0 and delta <= .008
        rows.append({'script': 'Thai', 'weight': weight, 'units_per_em': 1000,
                     'text': 'ภาษาไทยยินดีต้อนรับ', 'source': str(source_font), 'output': str(static),
                     'source_sha256': digest(source_font), 'output_sha256': digest(static),
                     'shape': shapes[1], 'source_shape': shapes[0],
                     'total_x_advance': sum(glyph['ax'] for glyph in shapes[1]),
                     'source_instance_positioning_matches_static': True,
                     'source_static_svg_path_topology_matches': True,
                     'source_static_svg_path_max_coordinate_delta_font_units': delta,
                     'source_static_svg_coordinate_pairs_compared': coordinates})
    (variable / 'receipt.json').write_text(json.dumps({'tools': identities, 'controls': rows, 'sourceArchiveSHA256': archive_sha, 'toolPins': tool_guards}, indent=2) + '\n')

    # Cargo supplies the real CLI path and the actual independent renderer ABI.
    cargo = json.loads(run(['cargo', 'metadata', '--locked', '--format-version=1', '--no-deps']))
    assert any(p['name'] == family + '-cli' for p in cargo['packages'])
    cli_output = run(['cargo', 'build', '--locked', '-p', family + '-cli', '--message-format=json-render-diagnostics'], 1200)
    cli_records = [json.loads(line) for line in cli_output.splitlines() if line.startswith(b'{')]
    cli = next(r['executable'] for r in cli_records if r.get('reason') == 'compiler-artifact' and r.get('executable') and r['target']['name'] == family)
    output = run(['cargo', 'test', '--locked', '-p', 'rdocx', '--lib', '--no-run', '--message-format=json-render-diagnostics'], 1200)
    artifacts = [json.loads(line) for line in output.splitlines() if line.startswith(b'{')]
    versions = {'resvg': '0.48.1', 'usvg': '0.48.1', 'tiny-skia': '0.12.0'}
    lock = tomllib.loads((root / 'Cargo.lock').read_text())
    selected = {}
    for name, version in versions.items():
        assert any(p['name'] == name and p['version'] == version for p in lock['package'])
        records = [r for r in artifacts if r.get('reason') == 'compiler-artifact' and r['package_id'].endswith('#' + name + '@' + version)]
        paths = {p for r in records for p in r['filenames'] if p.endswith('.rlib')}
        assert len(paths) == 1, (name, paths)
        selected[name] = Path(paths.pop())
    # Never link against an open shared cache or qualify a PATH shim as rustc.
    rustc = Path(run(['rustup', 'which', 'rustc']).decode().strip()).resolve()
    compiler_version = run([rustc, '-vV']).decode()
    assert compiler_version.startswith('rustc 1.97.1 ')
    sysroot = Path(run([rustc, '--print', 'sysroot']).decode().strip()).resolve()
    # This superset includes compiler libraries, target stdlib and proc-macro ABI.
    compiler_inputs = [rustc, *sorted(p for p in (sysroot / 'lib').rglob('*') if p.is_file())]
    assert len(compiler_inputs) > 1
    dependency_paths = sorted({Path(p).resolve() for record in artifacts
                               if record.get('reason') == 'compiler-artifact'
                               for p in record['filenames']
                               if p.endswith(('.rlib', '.dylib', '.so'))})
    assert set(selected.values()).issubset(set(dependency_paths))
    dependency_directory = directory / 'reference-dependencies'
    dependency_directory.mkdir()
    assert len({p.name for p in dependency_paths}) == len(dependency_paths)
    for path in dependency_paths:
        (dependency_directory / path.name).symlink_to(path)
    reference = directory / 'original-path-reference.rs'
    reference.write_text(_FX195_REFERENCE_RENDERER_SOURCE)
    binary = directory / 'original-path-reference'
    input_guards = {str(p): digest(p) for p in [reference, root / 'Cargo.lock',
                                               *compiler_inputs, *dependency_paths]}
    inventory = {str(p): str(p.resolve()) for p in dependency_directory.iterdir()}
    run([rustc, '--sysroot', sysroot, '--edition=2024', reference,
         '--extern', 'resvg=' + str(dependency_directory / selected['resvg'].name),
         '-L', 'dependency=' + str(dependency_directory), '-o', binary], 300)
    assert all(digest(p) == sha for p, sha in input_guards.items())
    assert {str(p): str(p.resolve()) for p in dependency_directory.iterdir()} == inventory
    assert sorted(str(p) for p in (sysroot / 'lib').rglob('*') if p.is_file()) == sorted(str(p) for p in compiler_inputs[1:])
    assert source_guard() == before
    compile_receipt = directory / 'reference-compile-receipt.json'
    compile_receipt.write_text(json.dumps({'actualExit': 0, 'guardsBefore': True, 'guardsAfter': True,
                                         'binarySHA': digest(binary), 'inputs': input_guards,
                                         'actualCompiler': str(rustc), 'actualCompilerSHA256': digest(rustc),
                                         'sysroot': str(sysroot), 'isolatedDependencyInventory': inventory,
                                         'cargoCompilerArtifacts': [record for record in artifacts if record.get('reason') == 'compiler-artifact'],
                                         'rustcVerboseVersion': compiler_version,
                                         'command': commands[-1]}, indent=2) + '\n')
    qualification = directory / 'path-reference-qualification.json'
    qualification.write_text(json.dumps({'actual_compiler_exit': 0, 'actual_transitive_abi_qualified': True,
                                        'renderer_source_sha256': digest(reference), 'rustc_verbose_version': compiler_version,
                                        'actual_transitive_versions': versions, 'actualCompileReceipt': str(compile_receipt),
                                        'actualCompileReceiptSHA': digest(compile_receipt), 'renderer_binary': str(binary),
                                        'renderer_binary_sha256': digest(binary), 'qualified_resvg_rlib': str(selected['resvg']),
                                        'qualified_resvg_rlib_sha256': digest(selected['resvg'])}, indent=2) + '\n')
    assert all(digest(p) == sha for p, sha in tool_guards.items())
    environment = {'F_X195_AXIS_ORACLES': str(axis), 'F_X195_VARIABLE_ORACLES': str(variable),
                   'F_X195_STYLE_FONT_GUARDS': str(guards), 'F_X195_PATH_REFERENCE_QUALIFICATION': str(qualification),
                   'F_X195_PATH_REFERENCE_QUALIFICATION_SHA': digest(qualification),
                   family.upper() + '_CLI': cli}
    (directory / 'environment.json').write_text(json.dumps(environment, indent=2) + '\n')
    if os.environ.get('GITHUB_ENV'):
        with Path(os.environ['GITHUB_ENV']).open('a') as out:
            for key, value in environment.items():
                out.write(key + '=' + value + '\n')
    (directory / 'receipt.json').write_text(json.dumps({'status': 'actual-independent-inputs-prepared',
        'sourceGuardsBeforeAfter': True, 'sourceInputs': before, 'commands': commands,
        'harfbuzzSourceURL': archive_url, 'harfbuzzSourceSHA256': archive_sha,
        'harfbuzzTools': tool_guards, 'environment': environment}, indent=2) + '\n')
    print(json.dumps(environment, indent=2))
    return environment


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Prepare independent font controls before binding pytest.')
    parser.add_argument('--prepare-fx195-inputs', required=True)
    parser.add_argument('--family', choices=('rdocx', 'rpptx'))
    parser.add_argument('--native-tools-only', action='store_true')
    parser.add_argument('--harfbuzz-archive')
    args = parser.parse_args()
    if args.native_tools_only and args.family is not None:
        parser.error('--native-tools-only cannot be combined with --family')
    if not args.native_tools_only and args.family is None:
        parser.error('--family is required for binding input preparation')
    _fx195_prepare_inputs(args.prepare_fx195_inputs, args.family, args.harfbuzz_archive,
                          args.native_tools_only)
