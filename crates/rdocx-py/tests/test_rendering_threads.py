import concurrent.futures
import io
import os
import shutil
import statistics
import subprocess
import sys
import tempfile
import threading
import time
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


def test_render_errors_reacquire_and_map_cleanly():
    from rdocx import LayoutError, RdocxError

    document = _nontrivial_document(3)

    with pytest.raises(LayoutError) as raised:
        document.render_pages(format="jpeg", quality=0, pages=[0])
    assert isinstance(raised.value, RdocxError)
    assert "JPEG quality 0 is outside 1 through 100" in str(raised.value)


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

    package = _assert_releases_gil(document.to_bytes)

    assert package.startswith(b"PK")


def test_render_page_to_png_releases_gil_for_python_worker():
    document = _nontrivial_document(11)

    page = _assert_releases_gil(lambda: document.render_page_to_png(0, dpi=72.0))

    assert page.startswith(PNG_SIGNATURE)


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
