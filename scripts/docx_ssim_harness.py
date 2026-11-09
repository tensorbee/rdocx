#!/usr/bin/env python3
"""Compare deterministic WordprocessingML renders with a pinned oracle."""

from __future__ import annotations

import argparse
from concurrent.futures import ProcessPoolExecutor
from contextlib import contextmanager, nullcontext
import ctypes
try:
    import fcntl
except ImportError:  # pragma: no cover - the harness targets POSIX CI and macOS Writer
    fcntl = None
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory, gettempdir
import time
import unittest
from unittest.mock import patch
import zipfile

from fetch_docx_corpus import EXPECTED_COUNT, load_manifest, verify_directory
from golden_png_harness import decode_png
from pptx_ssim_harness import composite_luminance, structural_similarity


REPO_ROOT = Path(__file__).resolve().parent.parent
CORPUS_MANIFEST = REPO_ROOT / "scripts" / "docx-corpus-manifest.tsv"
DEFAULT_CORPUS = REPO_ROOT / "corpus" / "docx"
SOFFICE = "soffice"
PDFTOPPM = "pdftoppm"
SOFFICE_VERSION = "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb"
PDFTOPPM_VERSION = "pdftoppm version 26.01.0"
SOFFICE_PDF_FILTER = "pdf:writer_pdf_Export"
REVISION_VIEW = "accepted"
DPI = 150
SSIM_TARGET = 0.95
COVERAGE_TARGET = 0.80
RESULT_HEADER = (
    "document\tcategory\tpage\trust_width\trust_height\toracle_width"
    "\toracle_height\tnormalization\tssim\trust_png\toracle_png"
)
EVIDENCE_ENV = "RDOCX_DOCX_GATE_EVIDENCE"
ACCEPTOR_BUILD_COMMAND = (
    "cargo",
    "build",
    "--locked",
    "--offline",
    "-p",
    "rdocx",
    "--no-default-features",
    "--message-format=json-render-diagnostics",
)
ACCEPTOR_SOURCE = r'''use rdocx::Document;
use std::env;
use std::fs;
use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args_os().skip(1);
    let input = PathBuf::from(arguments.next().ok_or("missing input DOCX")?);
    let output = PathBuf::from(arguments.next().ok_or("missing output DOCX")?);
    if arguments.next().is_some() {
        return Err("unexpected argument".into());
    }
    let mut document = Document::open(&input)?;
    if document.revisions().is_empty() {
        fs::copy(input, output)?;
        return Ok(());
    }
    document.accept_all()?;
    document.save(&output)?;
    let reopened = Document::open(output)?;
    if !reopened.revisions().is_empty() {
        return Err("accepted DOCX retains modeled revisions".into());
    }
    Ok(())
}
'''

MULTI_SCRIPT_FIXTURES = (
    (
        "arabic.docx",
        "arabic",
        "ar-SA",
        "Noto Sans Arabic",
        "العربية مرحبا بالعالم",
    ),
    (
        "devanagari.docx",
        "devanagari",
        "hi-IN",
        "Noto Sans Devanagari",
        "देवनागरी नमस्ते दुनिया",
    ),
    (
        "thai.docx",
        "thai",
        "th-TH",
        "Noto Sans Thai",
        "ภาษาไทยยินดีต้อนรับ",
    ),
    (
        "simplified-chinese.docx",
        "simplified-chinese",
        "zh-CN",
        "Noto Sans SC",
        "〈中〉、你好世界",
    ),
)
MULTI_SCRIPT_FONT_FILES = (
    "NotoSansArabic.ttf",
    "NotoSansDevanagari.ttf",
    "NotoSansThai.ttf",
)
ORACLE_FONT_DIRECTORY = REPO_ROOT / "scripts" / "oracle-fonts"
ORACLE_CJK_FONT = "NotoSansSC-FX058-oracle-thin.ttf"
ORACLE_CJK_SOURCE_SHA256 = (
    "b06144fa7b2d5212fe21344261449c9350f603e3e2ae625e76306022d024fbe5"
)
ORACLE_CJK_OUTPUT_SHA256 = (
    "390ba9f55d4dd69915736d2b225d602b40012cd2c50db4c1e6d2bbdfd61e63a6"
)
ORACLE_CJK_GENERATOR = "hb-subset (HarfBuzz) 13.2.1"
ORACLE_CJK_GENERATION_COMMAND = (
    "hb-subset crates/oxml-layout/fonts/NotoSansSC-FX058-subset.ttf "
    "--output-file=scripts/oracle-fonts/NotoSansSC-FX058-oracle-thin.ttf "
    "--unicodes='*' --variations='wght=100' --name-IDs='*' --name-languages='*'"
)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def assert_oracle_font_inventory() -> None:
    product_fonts = REPO_ROOT / "crates" / "oxml-layout" / "fonts"
    source = product_fonts / "NotoSansSC-FX058-subset.ttf"
    output = ORACLE_FONT_DIRECTORY / ORACLE_CJK_FONT
    expected = {
        "LICENSE-Noto",
        "NotoSansSC-FX058-oracle-thin.ttf",
        "PROVENANCE.md",
    }
    actual = {path.name for path in ORACLE_FONT_DIRECTORY.iterdir() if path.is_file()}
    if actual != expected:
        raise ValueError(
            f"oracle font inventory expected {sorted(expected)!r}, got {sorted(actual)!r}"
        )
    if sha256(source) != ORACLE_CJK_SOURCE_SHA256:
        raise ValueError("oracle CJK source font SHA-256 does not match provenance")
    if sha256(output) != ORACLE_CJK_OUTPUT_SHA256:
        raise ValueError("oracle CJK output font SHA-256 does not match provenance")
    if (ORACLE_FONT_DIRECTORY / "LICENSE-Noto").read_bytes() != (
        product_fonts / "LICENSE-Noto"
    ).read_bytes():
        raise ValueError("oracle CJK licence differs from the approved Noto licence")
    provenance = (ORACLE_FONT_DIRECTORY / "PROVENANCE.md").read_text(encoding="utf-8")
    for required in (
        ORACLE_CJK_SOURCE_SHA256,
        ORACLE_CJK_OUTPUT_SHA256,
        ORACLE_CJK_GENERATOR,
        ORACLE_CJK_GENERATION_COMMAND,
    ):
        if required not in provenance:
            raise ValueError(f"oracle CJK provenance is missing {required!r}")


def xml_text(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def write_deterministic_zip_member(
    archive: zipfile.ZipFile, name: str, data: str
) -> None:
    member = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
    member.compress_type = zipfile.ZIP_DEFLATED
    member.external_attr = 0o100644 << 16
    archive.writestr(member, data.encode("utf-8"))


def build_multi_script_fixtures(output: Path) -> list[dict[str, object]]:
    output.mkdir(parents=True, exist_ok=True)
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'''
    relationships = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    fixtures = []
    for filename, script, language, family, text in MULTI_SCRIPT_FIXTURES:
        language_attributes = f'w:val="{language}"'
        if script == "arabic":
            language_attributes += f' w:bidi="{language}"'
        if script == "simplified-chinese":
            language_attributes += f' w:eastAsia="{language}"'
        paragraphs = "".join(
            f'''<w:p><w:pPr><w:spacing w:after="0" w:line="480" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}" w:eastAsia="{family}" w:cs="{family}"/><w:kern w:val="1"/><w:sz w:val="48"/><w:szCs w:val="48"/><w:lang {language_attributes}/></w:rPr><w:t>{xml_text(text)}</w:t></w:r></w:p>'''
            for _ in range(6)
        )
        document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'''
        path = output / filename
        with zipfile.ZipFile(path, "w") as archive:
            write_deterministic_zip_member(archive, "[Content_Types].xml", content_types)
            write_deterministic_zip_member(archive, "_rels/.rels", relationships)
            write_deterministic_zip_member(archive, "word/document.xml", document_xml)
        fixtures.append(
            {
                "document": filename,
                "category": script,
                "language": language,
                "font": family,
                "text": text,
                "path": path,
            }
        )
    rtl_filename = "bidirectional.docx"
    rtl_language = "ar-SA"
    rtl_family = "Noto Sans Arabic"
    rtl_text = "العربية"
    rtl_paragraphs = "".join(
        f'''<w:p><w:pPr><w:bidi/><w:spacing w:after="0" w:line="480" w:lineRule="exact"/><w:ind w:start="360" w:end="720"/><w:jc w:val="start"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="{rtl_family}" w:hAnsi="{rtl_family}" w:eastAsia="{rtl_family}" w:cs="{rtl_family}"/><w:rtl/><w:kern w:val="1"/><w:sz w:val="48"/><w:szCs w:val="48"/><w:lang w:val="en-US" w:bidi="{rtl_language}"/></w:rPr><w:t xml:space="preserve">العربية </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="{rtl_family}" w:hAnsi="{rtl_family}" w:eastAsia="{rtl_family}" w:cs="{rtl_family}"/><w:rtl w:val="0"/><w:kern w:val="1"/><w:sz w:val="48"/><w:szCs w:val="48"/><w:lang w:val="en-US" w:bidi="{rtl_language}"/></w:rPr><w:t>ABC 123</w:t></w:r></w:p>'''
        for _ in range(6)
    )
    rtl_document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{rtl_paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'''
    rtl_path = output / rtl_filename
    with zipfile.ZipFile(rtl_path, "w") as archive:
        write_deterministic_zip_member(archive, "[Content_Types].xml", content_types)
        write_deterministic_zip_member(archive, "_rels/.rels", relationships)
        write_deterministic_zip_member(archive, "word/document.xml", rtl_document_xml)
    fixtures.append(
        {
            "document": rtl_filename,
            "category": "bidirectional",
            "language": rtl_language,
            "font": rtl_family,
            "text": rtl_text,
            "path": rtl_path,
        }
    )
    return fixtures


def install_oracle_fonts(profile: Path) -> None:
    assert_oracle_font_inventory()
    fonts = profile / "user" / "fonts"
    fonts.mkdir(parents=True, exist_ok=True)
    source = REPO_ROOT / "crates" / "oxml-layout" / "fonts"
    for filename in MULTI_SCRIPT_FONT_FILES:
        shutil.copyfile(source / filename, fonts / filename)
    shutil.copyfile(ORACLE_FONT_DIRECTORY / ORACLE_CJK_FONT, fonts / ORACLE_CJK_FONT)


ORACLE_FONT_LOCK_PATH = Path(gettempdir()) / "rdocx-docx-ssim-coretext-fonts.lock"
ORACLE_FONT_LOCK_TIMEOUT_SECONDS = 240.0
ORACLE_FONT_LOCK_POLL_SECONDS = 0.05


@contextmanager
def oracle_font_process_lock(
    path: Path = ORACLE_FONT_LOCK_PATH,
    timeout_seconds: float = ORACLE_FONT_LOCK_TIMEOUT_SECONDS,
):
    """Serialize process-scoped CoreText registration across harness runs."""
    if fcntl is None:
        raise OSError("oracle font registration requires POSIX advisory locking")
    descriptor = os.open(path, os.O_CREAT | os.O_RDWR, 0o600)
    deadline = time.monotonic() + timeout_seconds
    acquired = False
    try:
        while not acquired:
            try:
                fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
                acquired = True
            except BlockingIOError:
                if time.monotonic() >= deadline:
                    raise TimeoutError(
                        f"timed out waiting for oracle font registration lock {path}"
                    )
                time.sleep(ORACLE_FONT_LOCK_POLL_SECONDS)
        yield
    finally:
        try:
            if acquired:
                fcntl.flock(descriptor, fcntl.LOCK_UN)
        finally:
            os.close(descriptor)


@contextmanager
def oracle_font_registration(profile: Path):
    """Expose the approved fonts to the isolated Writer process."""
    install_oracle_fonts(profile)
    if sys.platform != "darwin":
        yield
        return
    with oracle_font_process_lock():
        with coretext_oracle_font_registration():
            yield


@contextmanager
def coretext_oracle_font_registration():
    """Register approved fonts for exactly one serialized macOS session."""

    core_foundation = ctypes.CDLL(
        "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation"
    )
    core_text = ctypes.CDLL("/System/Library/Frameworks/CoreText.framework/CoreText")
    core_foundation.CFURLCreateFromFileSystemRepresentation.argtypes = (
        ctypes.c_void_p,
        ctypes.c_char_p,
        ctypes.c_long,
        ctypes.c_bool,
    )
    core_foundation.CFURLCreateFromFileSystemRepresentation.restype = ctypes.c_void_p
    core_foundation.CFRelease.argtypes = (ctypes.c_void_p,)
    core_text.CTFontManagerRegisterFontsForURL.argtypes = (
        ctypes.c_void_p,
        ctypes.c_uint32,
        ctypes.POINTER(ctypes.c_void_p),
    )
    core_text.CTFontManagerRegisterFontsForURL.restype = ctypes.c_bool
    core_text.CTFontManagerUnregisterFontsForURL.argtypes = (
        ctypes.c_void_p,
        ctypes.c_uint32,
        ctypes.POINTER(ctypes.c_void_p),
    )
    core_text.CTFontManagerUnregisterFontsForURL.restype = ctypes.c_bool

    font_source = REPO_ROOT / "crates" / "oxml-layout" / "fonts"
    font_paths = [font_source / filename for filename in MULTI_SCRIPT_FONT_FILES]
    font_paths.append(ORACLE_FONT_DIRECTORY / ORACLE_CJK_FONT)
    registered = []
    session_scope = 3
    try:
        for path in font_paths:
            filename = path.name
            encoded = os.fsencode(path)
            url = core_foundation.CFURLCreateFromFileSystemRepresentation(
                None, encoded, len(encoded), False
            )
            if not url:
                raise ValueError(f"CoreText could not create a URL for {filename}")
            error = ctypes.c_void_p()
            if not core_text.CTFontManagerRegisterFontsForURL(
                url, session_scope, ctypes.byref(error)
            ):
                core_foundation.CFRelease(url)
                if error:
                    core_foundation.CFRelease(error)
                raise ValueError(f"CoreText could not register {filename}")
            if error:
                core_foundation.CFRelease(error)
            registered.append((filename, url))
        yield
    finally:
        for filename, url in reversed(registered):
            error = ctypes.c_void_p()
            if not core_text.CTFontManagerUnregisterFontsForURL(
                url, session_scope, ctypes.byref(error)
            ):
                if error:
                    core_foundation.CFRelease(error)
                core_foundation.CFRelease(url)
                raise ValueError(f"CoreText could not unregister {filename}")
            if error:
                core_foundation.CFRelease(error)
            core_foundation.CFRelease(url)


def tool_version(command: str, arguments: list[str]) -> str:
    completed = subprocess.run(
        [command, *arguments],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
    )
    lines = (completed.stdout + completed.stderr).splitlines()
    if not lines:
        raise ValueError(f"{command} did not report a version")
    return lines[0].strip()


def assert_tool_versions() -> tuple[str, str]:
    soffice = tool_version(SOFFICE, ["--version"])
    pdftoppm = tool_version(PDFTOPPM, ["-v"])
    if soffice != SOFFICE_VERSION:
        raise ValueError(
            f"LibreOffice version expected {SOFFICE_VERSION!r}, got {soffice!r}"
        )
    if pdftoppm != PDFTOPPM_VERSION:
        raise ValueError(
            f"pdftoppm version expected {PDFTOPPM_VERSION!r}, got {pdftoppm!r}"
        )
    return soffice, pdftoppm


def numbered_pages(directory: Path, prefix: str) -> list[Path]:
    pattern = re.compile(rf"^{re.escape(prefix)}([0-9]+)\.png$")
    numbered = []
    for candidate in directory.glob("*.png"):
        match = pattern.fullmatch(candidate.name)
        if match is None:
            raise ValueError(f"unexpected page output {candidate}")
        number = int(match.group(1))
        if number < 1:
            raise ValueError(f"invalid page number in {candidate}")
        numbered.append((number, candidate))
    numbered.sort()
    if not numbered:
        raise ValueError(f"renderer emitted no pages in {directory}")
    actual = [number for number, _ in numbered]
    expected = list(range(1, len(numbered) + 1))
    if actual != expected:
        raise ValueError(f"page sequence is {actual}, expected {expected}")
    return [path for _, path in numbered]


def run_rust_renderer(document: Path, output: Path) -> list[Path]:
    output.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            "cargo",
            "run",
            "--quiet",
            "--locked",
            "-p",
            "rdocx-cli",
            "--",
            "render",
            str(document),
            "--output-dir",
            str(output),
            "--dpi",
            str(DPI),
            "--format",
            "png",
        ],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
        timeout=300,
    )
    return numbered_pages(output, f"{document.stem}_page")


def build_acceptor(helper: Path) -> Path:
    source = helper / "src" / "main.rs"
    source.parent.mkdir(parents=True, exist_ok=True)
    source.write_text(ACCEPTOR_SOURCE, encoding="utf-8")
    built = subprocess.run(
        ACCEPTOR_BUILD_COMMAND,
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
        timeout=300,
    )
    artifacts = []
    for line in built.stdout.splitlines():
        message = json.loads(line)
        target = message.get("target", {})
        if (
            message.get("reason") == "compiler-artifact"
            and target.get("name") == "rdocx"
            and "lib" in target.get("kind", [])
        ):
            artifacts.extend(
                Path(filename)
                for filename in message.get("filenames", [])
                if filename.endswith(".rlib")
            )
    if len(artifacts) != 1:
        raise ValueError(f"rdocx build reported {len(artifacts)} library artifacts")
    dependency_directory = artifacts[0].parent
    if dependency_directory.name != "deps":
        dependency_directory /= "deps"
    binary = helper / "rdocx-acceptor"
    subprocess.run(
        [
            "rustc",
            "--edition=2024",
            str(source),
            "--extern",
            f"rdocx={artifacts[0]}",
            "-L",
            f"dependency={dependency_directory}",
            "-o",
            str(binary),
        ],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
        timeout=300,
    )
    if not binary.is_file():
        raise ValueError(f"rustc did not create revision acceptor {binary}")
    return binary


def prepare_accepted_document(document: Path, output: Path, helper: Path) -> Path:
    binary = helper / "rdocx-acceptor"
    if not binary.is_file():
        binary = build_acceptor(helper)
    output.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [str(binary), str(document), str(output)],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
        timeout=300,
    )
    if not output.is_file():
        raise ValueError(f"rdocx did not create accepted document {output}")
    return output


def render_oracle_document(
    document: Path, pdf_dir: Path, png_dir: Path, profile: Path, helper: Path
) -> list[Path]:
    for directory in (pdf_dir, png_dir, profile):
        directory.mkdir(parents=True, exist_ok=True)
    accepted_document = prepare_accepted_document(
        document, pdf_dir.parent / "accepted" / document.name, helper
    )
    with oracle_font_registration(profile):
        subprocess.run(
            [
                SOFFICE,
                "--headless",
                "--norestore",
                "--nodefault",
                "--nolockcheck",
                f"-env:UserInstallation={profile.as_uri()}",
                "--convert-to",
                SOFFICE_PDF_FILTER,
                "--outdir",
                str(pdf_dir),
                str(accepted_document),
            ],
            cwd=REPO_ROOT,
            check=True,
            capture_output=True,
            text=True,
            timeout=180,
        )
    pdf = pdf_dir / f"{accepted_document.stem}.pdf"
    if not pdf.is_file():
        raise ValueError(f"LibreOffice did not create {pdf}")
    prefix = png_dir / "oracle-page"
    subprocess.run(
        [PDFTOPPM, "-r", str(DPI), "-png", str(pdf), str(prefix)],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
        timeout=180,
    )
    return numbered_pages(png_dir, "oracle-page-")


def union_page_inputs(
    document: str, rust_pages: list[Path], oracle_pages: list[Path]
) -> list[tuple[str, int, Path | None, Path | None]]:
    union_count = max(len(rust_pages), len(oracle_pages))
    if union_count == 0:
        raise ValueError(f"{document} rendered no pages")
    return [
        (
            document,
            page + 1,
            rust_pages[page] if page < len(rust_pages) else None,
            oracle_pages[page] if page < len(oracle_pages) else None,
        )
        for page in range(union_count)
    ]


def white_image(width: int, height: int) -> tuple[int, int, bytes]:
    return width, height, bytes((255, 255, 255, 255)) * (width * height)


def composite_on_white_canvas(
    image: tuple[int, int, bytes], target_width: int, target_height: int
) -> tuple[int, int, bytes]:
    width, height, rgba = image
    if width > target_width or height > target_height:
        raise ValueError(
            f"image {width}x{height} exceeds canvas {target_width}x{target_height}"
        )
    if len(rgba) != width * height * 4:
        raise ValueError("decoded image has an invalid pixel buffer")
    if (width, height) == (target_width, target_height):
        return image
    canvas = bytearray(bytes((255, 255, 255, 255)) * (target_width * target_height))
    row_bytes = width * 4
    target_row_bytes = target_width * 4
    for row in range(height):
        source_start = row * row_bytes
        target_start = row * target_row_bytes
        canvas[target_start : target_start + row_bytes] = rgba[
            source_start : source_start + row_bytes
        ]
    return target_width, target_height, bytes(canvas)


def normalize_page_pair(
    rust_image: tuple[int, int, bytes] | None,
    oracle_image: tuple[int, int, bytes] | None,
) -> tuple[
    tuple[int, int, bytes],
    tuple[int, int, bytes],
    tuple[int, int, int, int, str],
]:
    if rust_image is None and oracle_image is None:
        raise ValueError("a union page has no Rust or oracle image")
    rust_width, rust_height = rust_image[:2] if rust_image is not None else (0, 0)
    oracle_width, oracle_height = (
        oracle_image[:2] if oracle_image is not None else (0, 0)
    )
    target_width = max(rust_width, oracle_width)
    target_height = max(rust_height, oracle_height)
    if rust_image is None:
        action = "blank-rust"
        normalized_rust = white_image(target_width, target_height)
    else:
        normalized_rust = composite_on_white_canvas(
            rust_image, target_width, target_height
        )
        action = "none"
    if oracle_image is None:
        action = "blank-oracle"
        normalized_oracle = white_image(target_width, target_height)
    else:
        normalized_oracle = composite_on_white_canvas(
            oracle_image, target_width, target_height
        )
    if (
        rust_image is not None
        and oracle_image is not None
        and (rust_width, rust_height) != (oracle_width, oracle_height)
    ):
        action = "shared-white-canvas"
    return (
        normalized_rust,
        normalized_oracle,
        (rust_width, rust_height, oracle_width, oracle_height, action),
    )


def score_union_page(
    page: tuple[str, int, Path | None, Path | None]
) -> tuple[float, int, int, int, int, str]:
    _, _, rust_png, oracle_png = page
    rust_image = decode_png(rust_png) if rust_png is not None else None
    oracle_image = decode_png(oracle_png) if oracle_png is not None else None
    normalized_rust, normalized_oracle, dimensions = normalize_page_pair(
        rust_image, oracle_image
    )
    return structural_similarity(normalized_rust, normalized_oracle), *dimensions


def meets_coverage(scores: list[float]) -> bool:
    if not scores:
        return False
    passing = sum(score >= SSIM_TARGET for score in scores)
    return passing / len(scores) >= COVERAGE_TARGET


def multi_script_gate_met(scores: list[float]) -> bool:
    return meets_coverage(scores)


def evidence_payload(
    documents: int,
    scores: list[float],
    results: Path,
    per_document: list[dict[str, object]],
) -> dict[str, object]:
    if documents != EXPECTED_COUNT:
        raise ValueError(f"rendered {documents} documents, expected {EXPECTED_COUNT}")
    if not scores:
        raise ValueError("the corpus rendered no pages")
    ordered = sorted(scores)
    passing = sum(score >= SSIM_TARGET for score in scores)
    middle = len(ordered) // 2
    median = (
        ordered[middle]
        if len(ordered) % 2
        else (ordered[middle - 1] + ordered[middle]) / 2
    )
    return {
        "documents": documents,
        "pages": len(scores),
        "rust_pages": sum(int(item["rust_pages"]) for item in per_document),
        "oracle_pages": sum(int(item["oracle_pages"]) for item in per_document),
        "dimension_mismatches": sum(
            int(item["dimension_mismatches"]) for item in per_document
        ),
        "rust_only_pages": sum(
            int(item["rust_only_pages"]) for item in per_document
        ),
        "oracle_only_pages": sum(
            int(item["oracle_only_pages"]) for item in per_document
        ),
        "per_document": per_document,
        "minimum": ordered[0],
        "median": median,
        "maximum": ordered[-1],
        "passing": passing,
        "coverage": passing / len(scores),
        "target": SSIM_TARGET,
        "coverage_target": COVERAGE_TARGET,
        "trend_target_met": meets_coverage(scores),
        "dpi": DPI,
        "revision_view": REVISION_VIEW,
        "results": str(results),
    }


def assert_expected_artifacts(paths: list[Path]) -> None:
    for path in paths:
        if not path.is_file() or path.stat().st_size == 0:
            raise ValueError(f"missing expected artifact {path}")


def assert_rtl_oracle_evidence(evidence: Path) -> None:
    payload = json.loads(evidence.read_text(encoding="utf-8"))
    if not payload["multi_script"]["gate_met"]:
        raise ValueError("multi-script oracle gate did not pass")
    results = Path(payload["results"]).read_text(encoding="utf-8").splitlines()
    header = results[0].split("\t")
    rows = [dict(zip(header, row.split("\t"))) for row in results[1:]]
    bidi_scores = [
        float(row["ssim"])
        for row in rows
        if row["category"] == "bidirectional"
    ]
    if not bidi_scores:
        raise ValueError("bidirectional oracle evidence is missing")
    if any(score < SSIM_TARGET for score in bidi_scores):
        raise ValueError("bidirectional oracle page missed the 0.95 SSIM threshold")


def run_gate(corpus: Path, output: Path) -> tuple[dict[str, object], Path]:
    entries = load_manifest(CORPUS_MANIFEST)
    if len(entries) != EXPECTED_COUNT:
        raise ValueError(f"corpus has {len(entries)} documents, expected {EXPECTED_COUNT}")
    verify_directory(corpus, entries)
    soffice_version, pdftoppm_version = assert_tool_versions()
    all_pages = []
    categories = {}
    page_counts = {}
    for relative_path, category, _, _, _, _, _ in entries:
        document = corpus / relative_path
        document_root = output / "documents" / relative_path
        rust_pages = run_rust_renderer(document, document_root / "rust")
        oracle_pages = render_oracle_document(
            document,
            document_root / "oracle-pdf",
            document_root / "oracle-png",
            document_root / "libreoffice-profile",
            output / "rdocx-acceptor",
        )
        all_pages.extend(union_page_inputs(relative_path, rust_pages, oracle_pages))
        categories[relative_path] = category
        page_counts[relative_path] = (len(rust_pages), len(oracle_pages))
    corpus_page_count = len(all_pages)
    fixtures = build_multi_script_fixtures(output / "multi-script-corpus")
    for fixture in fixtures:
        relative_path = str(fixture["document"])
        document = Path(fixture["path"])
        document_root = output / "multi-script-documents" / relative_path
        rust_pages = run_rust_renderer(document, document_root / "rust")
        oracle_pages = render_oracle_document(
            document,
            document_root / "oracle-pdf",
            document_root / "oracle-png",
            document_root / "libreoffice-profile",
            output / "rdocx-acceptor",
        )
        all_pages.extend(union_page_inputs(relative_path, rust_pages, oracle_pages))
        categories[relative_path] = str(fixture["category"])
        page_counts[relative_path] = (len(rust_pages), len(oracle_pages))
    workers = min(8, os.cpu_count() or 1)
    with ProcessPoolExecutor(max_workers=workers) as executor:
        scored_pages = list(executor.map(score_union_page, all_pages))
    scores = [score for score, *_ in scored_pages]
    corpus_scores = scores[:corpus_page_count]
    multi_script_scores = scores[corpus_page_count:]
    results = output / "ssim-results.tsv"
    rows = [RESULT_HEADER]
    rows.extend(
        f"{document}\t{categories[document]}\t{page}\t{rust_width}\t"
        f"{rust_height}\t{oracle_width}\t{oracle_height}\t{normalization}\t"
        f"{score:.9f}\t{rust_png or ''}\t{oracle_png or ''}"
        for (
            (document, page, rust_png, oracle_png),
            (score, rust_width, rust_height, oracle_width, oracle_height, normalization),
        ) in zip(all_pages, scored_pages)
    )
    results.write_text("\n".join(rows) + "\n", encoding="utf-8")
    per_document = []
    for relative_path, category, _, _, _, _, _ in entries:
        rust_count, oracle_count = page_counts[relative_path]
        document_scores = [
            scored
            for page, scored in zip(all_pages, scored_pages)
            if page[0] == relative_path
        ]
        per_document.append(
            {
                "document": relative_path,
                "category": category,
                "rust_pages": rust_count,
                "oracle_pages": oracle_count,
                "union_pages": max(rust_count, oracle_count),
                "dimension_mismatches": sum(
                    scored[-1] == "shared-white-canvas" for scored in document_scores
                ),
                "rust_only_pages": sum(
                    scored[-1] == "blank-oracle" for scored in document_scores
                ),
                "oracle_only_pages": sum(
                    scored[-1] == "blank-rust" for scored in document_scores
                ),
            }
        )
    multi_script_per_document = []
    for fixture in fixtures:
        relative_path = str(fixture["document"])
        rust_count, oracle_count = page_counts[relative_path]
        document_scores = [
            scored
            for page, scored in zip(all_pages, scored_pages)
            if page[0] == relative_path
        ]
        multi_script_per_document.append(
            {
                "document": relative_path,
                "category": fixture["category"],
                "language": fixture["language"],
                "font": fixture["font"],
                "rust_pages": rust_count,
                "oracle_pages": oracle_count,
                "union_pages": max(rust_count, oracle_count),
                "dimension_mismatches": sum(
                    scored[-1] == "shared-white-canvas" for scored in document_scores
                ),
                "rust_only_pages": sum(
                    scored[-1] == "blank-oracle" for scored in document_scores
                ),
                "oracle_only_pages": sum(
                    scored[-1] == "blank-rust" for scored in document_scores
                ),
            }
        )
    payload = evidence_payload(len(entries), corpus_scores, results, per_document)
    multi_script_passing = sum(
        score >= SSIM_TARGET for score in multi_script_scores
    )
    payload["multi_script"] = {
        "fixtures": len(fixtures),
        "pages": len(multi_script_scores),
        "passing": multi_script_passing,
        "coverage": multi_script_passing / len(multi_script_scores),
        "target": SSIM_TARGET,
        "coverage_target": COVERAGE_TARGET,
        "gate_met": multi_script_gate_met(multi_script_scores),
        "minimum": min(multi_script_scores),
        "maximum": max(multi_script_scores),
        "per_document": multi_script_per_document,
    }
    payload["libreoffice"] = soffice_version
    payload["pdftoppm"] = pdftoppm_version
    evidence = output / "gate-evidence.json"
    evidence.write_text(
        json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    assert_expected_artifacts([results, evidence])
    if not payload["multi_script"]["gate_met"]:
        raise ValueError(
            "multi-script corpus missed the hard 0.95 SSIM on 80 percent of pages gate"
        )
    return payload, evidence


def run_suite(test_names: list[str] | None = None) -> bool:
    loader = unittest.defaultTestLoader
    suite = (
        loader.loadTestsFromNames(test_names, module=sys.modules[__name__])
        if test_names
        else loader.loadTestsFromTestCase(DocxSsimHarnessTests)
    )
    return unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful()


class DocxSsimHarnessTests(unittest.TestCase):
    def test_oracle_font_inventory_is_exact_and_reproducible(self) -> None:
        assert_oracle_font_inventory()

    @unittest.skipUnless(os.name == "posix", "requires POSIX advisory locking")
    def test_oracle_font_lock_serializes_processes_and_releases_after_error(self) -> None:
        child = """
import sys
from pathlib import Path
sys.path.insert(0, sys.argv[1])
from docx_ssim_harness import oracle_font_process_lock
try:
    with oracle_font_process_lock(Path(sys.argv[2]), float(sys.argv[3])):
        pass
except TimeoutError:
    raise SystemExit(3)
"""
        with TemporaryDirectory() as temporary:
            lock_path = Path(temporary) / "coretext.lock"

            def child_attempt(timeout: float) -> subprocess.CompletedProcess[str]:
                return subprocess.run(
                    [
                        sys.executable,
                        "-c",
                        child,
                        str(REPO_ROOT / "scripts"),
                        str(lock_path),
                        str(timeout),
                    ],
                    cwd=REPO_ROOT,
                    capture_output=True,
                    text=True,
                    check=False,
                )

            with self.assertRaisesRegex(RuntimeError, "registration failed"):
                with oracle_font_process_lock(lock_path, 1.0):
                    blocked = child_attempt(0.1)
                    self.assertEqual(
                        blocked.returncode,
                        3,
                        blocked.stdout + blocked.stderr,
                    )
                    raise RuntimeError("registration failed")

            acquired = child_attempt(1.0)
            self.assertEqual(acquired.returncode, 0, acquired.stdout + acquired.stderr)

    def test_identical_images_have_ssim_one(self) -> None:
        image = (2, 1, bytes((10, 20, 30, 255, 40, 50, 60, 128)))
        self.assertEqual(structural_similarity(image, image), 1.0)

    def test_alpha_is_composited_over_white(self) -> None:
        self.assertEqual(
            composite_luminance(bytes((0, 0, 0, 0, 0, 0, 0, 255))),
            [255, 0],
        )

    def test_a_one_pixel_layout_shift_moves_word_ssim(self) -> None:
        pixels = [255] * 25
        for index in (6, 7, 11, 12, 13, 17):
            pixels[index] = 0
        shifted = [255] * 25
        for row in range(5):
            shifted[row * 5 + 1 : row * 5 + 5] = pixels[row * 5 : row * 5 + 4]
        first = (5, 5, bytes(value for pixel in pixels for value in (pixel,) * 3 + (255,)))
        second = (
            5,
            5,
            bytes(value for pixel in shifted for value in (pixel,) * 3 + (255,)),
        )
        score = structural_similarity(first, second)
        self.assertLess(score, 1.0)
        self.assertTrue(math.isclose(score, 0.343724364, abs_tol=0.000000001))

    def test_oracle_export_uses_an_rdocx_accepted_copy(self) -> None:
        with TemporaryDirectory() as temporary, patch(
            f"{__name__}.prepare_accepted_document",
            return_value=Path(temporary) / "accepted.docx",
        ) as prepare, patch(f"{__name__}.subprocess.run") as run, patch(
            f"{__name__}.numbered_pages", return_value=[Path("oracle-page-1.png")]
        ), patch(
            f"{__name__}.oracle_font_registration", return_value=nullcontext()
        ):
            root = Path(temporary)
            (root / "oracle-pdf" / "accepted.pdf").parent.mkdir(parents=True)
            (root / "oracle-pdf" / "accepted.pdf").touch()
            pages = render_oracle_document(
                root / "original.docx",
                root / "oracle-pdf",
                root / "oracle-png",
                root / "profile",
                root / "helper",
            )
        prepare.assert_called_once()
        self.assertEqual(run.call_args_list[0].args[0][-1], str(root / "accepted.docx"))
        self.assertEqual(pages, [Path("oracle-page-1.png")])

    def test_acceptor_resolves_and_verifies_every_modeled_revision(self) -> None:
        self.assertIn("document.revisions().is_empty()", ACCEPTOR_SOURCE)
        self.assertIn("fs::copy(input, output)?;", ACCEPTOR_SOURCE)
        self.assertIn("document.accept_all()?;", ACCEPTOR_SOURCE)
        self.assertIn("!reopened.revisions().is_empty()", ACCEPTOR_SOURCE)
        self.assertIn("--locked", ACCEPTOR_BUILD_COMMAND)
        self.assertIn("--offline", ACCEPTOR_BUILD_COMMAND)

    def test_raw_ssim_requires_equal_dimensions(self) -> None:
        with self.assertRaisesRegex(ValueError, "image dimensions differ"):
            structural_similarity((1, 1, bytes(4)), (2, 1, bytes(8)))

    def test_union_coverage_keeps_every_page_index(self) -> None:
        pages = union_page_inputs(
            "sample.docx",
            [Path("rust-1.png"), Path("rust-2.png")],
            [Path("oracle-1.png"), Path("oracle-2.png"), Path("oracle-3.png")],
        )
        self.assertEqual(len(pages), 3)
        self.assertEqual(pages[-1], ("sample.docx", 3, None, Path("oracle-3.png")))

    def test_dimension_mismatch_uses_a_shared_white_canvas(self) -> None:
        rust = (1, 2, bytes((0, 0, 0, 255)) * 2)
        oracle = (2, 1, bytes((0, 0, 0, 255)) * 2)
        normalized_rust, normalized_oracle, dimensions = normalize_page_pair(
            rust, oracle
        )
        self.assertEqual(dimensions, (1, 2, 2, 1, "shared-white-canvas"))
        self.assertEqual(normalized_rust[:2], (2, 2))
        self.assertEqual(normalized_oracle[:2], (2, 2))
        self.assertEqual(normalized_rust[2][4:8], bytes((255, 255, 255, 255)))
        self.assertEqual(normalized_oracle[2][8:], bytes((255, 255, 255, 255)) * 2)

    def test_unmatched_page_scores_against_a_blank_white_counterpart(self) -> None:
        rust = (1, 1, bytes((0, 0, 0, 255)))
        normalized_rust, normalized_oracle, dimensions = normalize_page_pair(
            rust, None
        )
        self.assertEqual(dimensions, (1, 1, 0, 0, "blank-oracle"))
        self.assertEqual(normalized_oracle, white_image(1, 1))
        self.assertLess(structural_similarity(normalized_rust, normalized_oracle), 1.0)

    def test_page_sequence_must_be_contiguous(self) -> None:
        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            (directory / "page-1.png").touch()
            (directory / "page-3.png").touch()
            with self.assertRaisesRegex(ValueError, "page sequence"):
                numbered_pages(directory, "page-")

    def test_zero_padded_oracle_page_numbers_are_accepted(self) -> None:
        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            first = directory / "page-01.png"
            second = directory / "page-02.png"
            first.touch()
            second.touch()
            self.assertEqual(numbered_pages(directory, "page-"), [first, second])

    def test_trend_target_classifies_eighty_percent(self) -> None:
        self.assertTrue(meets_coverage([0.95] * 4 + [0.94]))
        self.assertFalse(meets_coverage([0.95] * 3 + [0.94] * 2))

    def test_missed_trend_is_advisory(self) -> None:
        per_document = [
            {
                "rust_pages": 1,
                "oracle_pages": 1,
                "dimension_mismatches": 0,
                "rust_only_pages": 0,
                "oracle_only_pages": 0,
            }
        ]
        payload = evidence_payload(
            EXPECTED_COUNT, [0.94], Path("scores.tsv"), per_document
        )
        self.assertFalse(payload["trend_target_met"])
        self.assertEqual(payload["documents"], EXPECTED_COUNT)

    def test_multi_script_corpus_pages_meet_the_reviewed_oracle_contract(self) -> None:
        from copy import deepcopy
        import xml.etree.ElementTree as ET

        word = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        expected = (
            (
                "arabic.docx", "arabic", "ar-SA", "Noto Sans Arabic",
                "العربية مرحبا بالعالم",
            ),
            (
                "devanagari.docx", "devanagari", "hi-IN", "Noto Sans Devanagari",
                "देवनागरी नमस्ते दुनिया",
            ),
            ("thai.docx", "thai", "th-TH", "Noto Sans Thai", "ภาษาไทยยินดีต้อนรับ"),
            (
                "simplified-chinese.docx", "simplified-chinese", "zh-CN", "Noto Sans SC",
                "〈中〉、你好世界",
            ),
            ("bidirectional.docx", "bidirectional", "ar-SA", "Noto Sans Arabic", "العربية"),
        )

        def assert_fixture(tree, spec):
            _, category, language, family, text = spec
            bidi = category == "bidirectional"
            self.assertEqual(tree.tag, word + "document")
            self.assertEqual([child.tag for child in tree], [word + "body"])
            body = tree.find(word + "body")
            self.assertEqual(
                [child.tag for child in body], [word + "p"] * 6 + [word + "sectPr"]
            )
            section = body.find(word + "sectPr")
            self.assertEqual(
                [child.tag for child in section], [word + "pgSz", word + "pgMar"]
            )
            self.assertEqual(section[0].attrib, {word + "w": "12240", word + "h": "15840"})
            self.assertEqual(section[1].attrib, {
                word + key: value for key, value in (
                    ("top", "1440"), ("right", "1440"), ("bottom", "1440"), ("left", "1440"),
                    ("header", "720"), ("footer", "720"), ("gutter", "0"),
                )
            })
            for paragraph in body.findall(word + "p"):
                self.assertEqual(
                    [child.tag for child in paragraph],
                    [word + "pPr"] + [word + "r"] * (2 if bidi else 1),
                )
                properties = paragraph[0]
                expected_properties = (
                    ["bidi", "spacing", "ind", "jc"] if bidi else ["spacing"]
                )
                self.assertEqual(
                    [child.tag for child in properties],
                    [word + tag for tag in expected_properties],
                )
                self.assertEqual(properties.find(word + "spacing").attrib, {
                    word + "after": "0", word + "line": "480", word + "lineRule": "exact",
                })
                if bidi:
                    self.assertEqual(properties.find(word + "bidi").attrib, {})
                    self.assertEqual(
                        properties.find(word + "ind").attrib,
                        {word + "start": "360", word + "end": "720"},
                    )
                    self.assertEqual(
                        properties.find(word + "jc").attrib, {word + "val": "start"}
                    )
                for index, run in enumerate(paragraph.findall(word + "r")):
                    self.assertEqual([child.tag for child in run], [word + "rPr", word + "t"])
                    rpr = run[0]
                    kern = rpr.findall(word + "kern")
                    self.assertEqual(
                        len(kern), 1, "each qualified run requires exactly one kern1"
                    )
                    self.assertEqual(kern[0].attrib, {word + "val": "1"})
                    expected_rpr = (
                        ["rFonts"] + (["rtl"] if bidi else [])
                        + ["kern", "sz", "szCs", "lang"]
                    )
                    self.assertEqual(
                        [child.tag for child in rpr], [word + tag for tag in expected_rpr]
                    )
                    self.assertEqual(rpr.find(word + "rFonts").attrib, {
                        word + slot: family for slot in ("ascii", "hAnsi", "eastAsia", "cs")
                    })
                    for size in ("sz", "szCs"):
                        self.assertEqual(rpr.find(word + size).attrib, {word + "val": "48"})
                    languages = {word + "val": "en-US" if bidi else language}
                    if bidi or category == "arabic":
                        languages[word + "bidi"] = language
                    if category == "simplified-chinese":
                        languages[word + "eastAsia"] = language
                    self.assertEqual(rpr.find(word + "lang").attrib, languages)
                    if bidi:
                        self.assertEqual(
                            rpr.find(word + "rtl").attrib,
                            {} if index == 0 else {word + "val": "0"},
                        )
                    self.assertEqual(
                        run[1].text,
                        ("العربية " if index == 0 else "ABC 123") if bidi else text,
                    )
                    self.assertEqual(
                        run[1].attrib,
                        {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}
                        if bidi and index == 0 else {},
                    )

        with TemporaryDirectory() as temporary:
            fixtures = build_multi_script_fixtures(Path(temporary))
            self.assertEqual(
                [fixture["category"] for fixture in fixtures],
                [spec[1] for spec in expected],
            )
            self.assertEqual(
                [fixture["document"] for fixture in fixtures],
                [spec[0] for spec in expected],
            )
            total_runs = 0
            for fixture, spec in zip(fixtures, expected):
                with self.subTest(fixture=spec[0]):
                    self.assertEqual(
                        tuple(fixture[key] for key in (
                            "document", "category", "language", "font", "text",
                        )),
                        spec,
                    )
                    path = Path(fixture["path"])
                    self.assertTrue(path.is_file())
                    self.assertEqual(path.name, spec[0])
                    with zipfile.ZipFile(path) as archive:
                        self.assertEqual(
                            archive.namelist(),
                            ["[Content_Types].xml", "_rels/.rels", "word/document.xml"],
                        )
                        tree = ET.fromstring(archive.read("word/document.xml"))
                    total_runs += len(tree.findall(".//" + word + "r"))
                    assert_fixture(tree, spec)
                    # These reject an oracle-fixture policy, not valid Word
                    # authoring/imported zero or above-size kerning properties.
                    for mutation in ("absent", "zero", "above-size", "duplicate", "unqualified"):
                        with self.subTest(fixture=spec[0], mutation=mutation):
                            changed = deepcopy(tree)
                            rpr = changed.findall(".//" + word + "rPr")[-1]
                            kern = rpr.find(word + "kern")
                            if mutation == "absent":
                                rpr.remove(kern)
                            elif mutation in ("zero", "above-size"):
                                kern.set(word + "val", "0" if mutation == "zero" else "49")
                            elif mutation == "duplicate":
                                rpr.append(deepcopy(kern))
                            else:
                                kern.tag = "kern"
                            with self.assertRaises(AssertionError):
                                assert_fixture(changed, spec)
                    assert_fixture(tree, spec)
            self.assertEqual(total_runs, 36)
        self.assertTrue(multi_script_gate_met([0.95] * 4 + [0.94]))
        self.assertFalse(multi_script_gate_met([0.95] * 3 + [0.94] * 2))

    def test_rtl_corpus_document_contains_directional_fixture_xml(self) -> None:
        with TemporaryDirectory() as temporary:
            fixtures = build_multi_script_fixtures(Path(temporary))
            rtl = fixtures[-1]
            self.assertEqual(rtl["category"], "bidirectional")
            with zipfile.ZipFile(Path(rtl["path"])) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("<w:bidi/>", document_xml)
            self.assertIn("<w:rtl/>", document_xml)
            self.assertIn('<w:rtl w:val="0"/>', document_xml)
            self.assertIn('<w:jc w:val="start"/>', document_xml)
            self.assertIn('<w:ind w:start="360" w:end="720"/>', document_xml)
            self.assertIn("ABC 123", document_xml)

    def test_rtl_corpus_document_matches_the_reviewed_oracle(self) -> None:
        evidence = os.environ.get(EVIDENCE_ENV)
        if evidence is None:
            self.skipTest("requires corpus gate evidence")
        assert_rtl_oracle_evidence(Path(evidence))

    def test_rtl_oracle_gate_rejects_below_threshold_render_evidence(self) -> None:
        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            results = directory / "ssim-results.tsv"
            results.write_text(
                RESULT_HEADER
                + "\n"
                + "bidirectional.docx\tbidirectional\t1\t10\t10\t10\t10\tnone"
                + "\t0.949999999\trust.png\toracle.png\n",
                encoding="utf-8",
            )
            evidence = directory / "gate-evidence.json"
            evidence.write_text(
                json.dumps(
                    {
                        "multi_script": {"gate_met": True},
                        "results": str(results),
                    }
                ),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(ValueError, "missed the 0.95 SSIM"):
                assert_rtl_oracle_evidence(evidence)

    def test_missing_expected_artifact_is_a_hard_failure(self) -> None:
        with TemporaryDirectory() as temporary:
            missing = Path(temporary) / "missing.tsv"
            with self.assertRaisesRegex(ValueError, "missing expected artifact"):
                assert_expected_artifacts([missing])

    def test_oracle_environment_is_exactly_pinned(self) -> None:
        self.assertEqual(
            SOFFICE_VERSION,
            "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb",
        )
        self.assertEqual(PDFTOPPM_VERSION, "pdftoppm version 26.01.0")
        self.assertEqual(SOFFICE_PDF_FILTER, "pdf:writer_pdf_Export")
        self.assertEqual(REVISION_VIEW, "accepted")
        self.assertEqual(DPI, 150)

    def test_tool_version_drift_is_a_hard_failure(self) -> None:
        with patch(
            f"{__name__}.tool_version",
            side_effect=("LibreOffice 0.0.0", PDFTOPPM_VERSION),
        ), self.assertRaisesRegex(ValueError, "LibreOffice version expected"):
            assert_tool_versions()

    def test_full_corpus_evidence_is_complete(self) -> None:
        evidence = os.environ.get(EVIDENCE_ENV)
        if evidence is None:
            self.skipTest("requires corpus gate evidence")
        payload = json.loads(Path(evidence).read_text(encoding="utf-8"))
        self.assertEqual(payload["documents"], EXPECTED_COUNT)
        self.assertGreater(payload["pages"], 0)
        self.assertEqual(
            payload["pages"],
            sum(item["union_pages"] for item in payload["per_document"]),
        )
        self.assertEqual(
            payload["rust_pages"],
            sum(item["rust_pages"] for item in payload["per_document"]),
        )
        self.assertEqual(
            payload["oracle_pages"],
            sum(item["oracle_pages"] for item in payload["per_document"]),
        )
        self.assertEqual(payload["dpi"], DPI)
        self.assertEqual(payload["revision_view"], REVISION_VIEW)
        self.assertEqual(payload["libreoffice"], SOFFICE_VERSION)
        self.assertEqual(payload["pdftoppm"], PDFTOPPM_VERSION)
        self.assertEqual(payload["multi_script"]["fixtures"], 5)
        self.assertGreater(payload["multi_script"]["pages"], 0)
        self.assertTrue(payload["multi_script"]["gate_met"])
        self.assertGreaterEqual(payload["multi_script"]["coverage"], COVERAGE_TARGET)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--self-test", action="store_true")
    mode.add_argument("--check", action="store_true")
    parser.add_argument(
        "--corpus-dir",
        type=Path,
        default=Path(os.environ.get("RDOCX_DOCX_CORPUS_DIR", DEFAULT_CORPUS)),
    )
    parser.add_argument("--output-dir", type=Path)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.self_test:
        return 0 if run_suite() else 1
    if not run_suite():
        return 1
    temporary = None
    try:
        if args.output_dir is None:
            temporary = TemporaryDirectory(prefix="rdocx-ssim-")
            output = Path(temporary.name)
        else:
            output = args.output_dir.resolve()
            if output.exists() and any(output.iterdir()):
                raise ValueError(f"output directory is not empty: {output}")
            output.mkdir(parents=True, exist_ok=True)
        payload, evidence = run_gate(args.corpus_dir.resolve(), output)
        os.environ[EVIDENCE_ENV] = str(evidence)
        if not run_suite(
            [
                "DocxSsimHarnessTests.test_full_corpus_evidence_is_complete",
                "DocxSsimHarnessTests.test_rtl_corpus_document_matches_the_reviewed_oracle",
            ]
        ):
            return 1
        print(
            "docx_ssim_harness: SSIM trend "
            f"{payload['passing']}/{payload['pages']} pages at SSIM >= {SSIM_TARGET:.2f} "
            f"({payload['coverage']:.3%}), min {payload['minimum']:.6f}, "
            f"median {payload['median']:.6f}, max {payload['maximum']:.6f}, "
            f"target met {str(payload['trend_target_met']).lower()}"
        )
        multi_script = payload["multi_script"]
        print(
            "docx_ssim_harness: multi-script hard gate "
            f"{multi_script['passing']}/{multi_script['pages']} pages at SSIM >= "
            f"{SSIM_TARGET:.2f} ({multi_script['coverage']:.3%}), target met "
            f"{str(multi_script['gate_met']).lower()}"
        )
        print(f"docx_ssim_harness: results {payload['results']}")
        return 0
    except (OSError, ValueError, subprocess.CalledProcessError, subprocess.TimeoutExpired) as error:
        print(f"docx_ssim_harness: {error}", file=sys.stderr)
        return 2
    finally:
        if temporary is not None:
            temporary.cleanup()


if __name__ == "__main__":
    raise SystemExit(main())
