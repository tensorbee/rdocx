#!/usr/bin/env python3
"""Validate every workspace README and compile its Rust examples."""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass
import datetime
import json
from pathlib import Path
import re
import subprocess
import sys
import tarfile
import tomllib
from urllib.parse import unquote, urlsplit
from urllib.request import Request, urlopen


REPO_ROOT = Path(__file__).resolve().parent.parent
RUST_FENCE = re.compile(r"^```rust(?P<attributes>[^\n]*)$", re.MULTILINE)
EXAMPLE_FENCE = re.compile(
    r"^```(?:rust[^\n]*|toml|python|javascript|sh|text)$",
    re.MULTILINE,
)
WORKSPACE_PACKAGE_COUNT = 27
PUBLISHABLE_PACKAGE_COUNT = 22
LOCAL_PATCHES = (
    ("oxml-core", "crates/oxml-core"),
    ("oxml-drawing", "crates/oxml-drawing"),
    ("oxml-layout", "crates/oxml-layout"),
    ("oxml-media", "crates/oxml-media"),
    ("oxml-opc", "crates/oxml-opc"),
    ("oxml-pdf", "crates/oxml-pdf"),
    ("oxml-sml", "crates/oxml-sml"),
    ("oxml-cli-support", "crates/oxml-cli-support"),
    ("oxml-chart", "crates/oxml-chart"),
    ("rdocx", "crates/rdocx"),
    ("rdocx-cli", "crates/rdocx-cli"),
    ("rdocx-html", "crates/rdocx-html"),
    ("rdocx-layout", "crates/rdocx-layout"),
    ("rdocx-opc", "crates/rdocx-opc"),
    ("rdocx-oxml", "crates/rdocx-oxml"),
    ("rdocx-pdf", "crates/rdocx-pdf"),
    ("rpptx", "crates/rpptx"),
    ("rpptx-cli", "crates/rpptx-cli"),
    ("rpptx-chart", "crates/rpptx-chart"),
    ("rpptx-layout", "crates/rpptx-layout"),
    ("rpptx-oxml", "crates/rpptx-oxml"),
    ("rpptx-render", "crates/rpptx-render"),
)


@dataclass(frozen=True)
class ReadmeCase:
    package: str
    crate_name: str
    readme: Path
    expected_rust_fences: int
    companion_crates: tuple[tuple[str, str], ...] = ()


README_CASES = (
    ReadmeCase("rdocx", "rdocx", REPO_ROOT / "README.md", 3),
    ReadmeCase(
        "rdocx-opc",
        "rdocx_opc",
        REPO_ROOT / "crates/rdocx-opc/README.md",
        1,
    ),
    ReadmeCase(
        "rdocx-oxml",
        "rdocx_oxml",
        REPO_ROOT / "crates/rdocx-oxml/README.md",
        1,
    ),
    ReadmeCase(
        "rdocx-layout",
        "rdocx_layout",
        REPO_ROOT / "crates/rdocx-layout/README.md",
        1,
    ),
    ReadmeCase(
        "rdocx-html",
        "rdocx_html",
        REPO_ROOT / "crates/rdocx-html/README.md",
        1,
    ),
    ReadmeCase(
        "rdocx-pdf",
        "rdocx_pdf",
        REPO_ROOT / "crates/rdocx-pdf/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-cli-support",
        "oxml_cli_support",
        REPO_ROOT / "crates/oxml-cli-support/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-core", "oxml_core", REPO_ROOT / "crates/oxml-core/README.md", 1
    ),
    ReadmeCase(
        "oxml-drawing",
        "oxml_drawing",
        REPO_ROOT / "crates/oxml-drawing/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-layout",
        "oxml_layout",
        REPO_ROOT / "crates/oxml-layout/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-media", "oxml_media", REPO_ROOT / "crates/oxml-media/README.md", 1
    ),
    ReadmeCase(
        "oxml-opc", "oxml_opc", REPO_ROOT / "crates/oxml-opc/README.md", 1
    ),
    ReadmeCase(
        "oxml-pdf",
        "oxml_pdf",
        REPO_ROOT / "crates/oxml-pdf/README.md",
        1,
        (("oxml-layout", "oxml_layout"),),
    ),
    ReadmeCase(
        "oxml-py-support",
        "oxml_py_support",
        REPO_ROOT / "crates/oxml-py-support/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-sml",
        "oxml_sml",
        REPO_ROOT / "crates/oxml-sml/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-chart",
        "oxml_chart",
        REPO_ROOT / "crates/oxml-chart/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx",
        "rpptx",
        REPO_ROOT / "crates/rpptx/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx-chart",
        "rpptx_chart",
        REPO_ROOT / "crates/rpptx-chart/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx-layout",
        "rpptx_layout",
        REPO_ROOT / "crates/rpptx-layout/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx-oxml",
        "rpptx_oxml",
        REPO_ROOT / "crates/rpptx-oxml/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx-render",
        "rpptx_render",
        REPO_ROOT / "crates/rpptx-render/README.md",
        1,
    ),
)

README_REQUIRED_TEXT = {
    REPO_ROOT / "README.md": (
        'rdocx = "0.14.0"',
        'rdocx = { version = "0.14.0", default-features = false }',
        "rdocx convert report.docx --to pdf -o report.pdf",
        "rdocx convert report.docx --to html -o report.html",
        "rdocx convert report.docx --to md -o report.md",
        'rdocx replace report.docx --placeholder "Draft" --value "Final" -o final.docx',
    ),
    REPO_ROOT / "crates/rdocx-cli/README.md": (
        "cargo install rdocx-cli --version '^0.14.0'",
        "rdocx convert report.docx --to pdf -o report.pdf",
    ),
    REPO_ROOT / "crates/rdocx-html/README.md": ('rdocx-html = "0.14.0"',),
    REPO_ROOT / "crates/rdocx-layout/README.md": ('rdocx-layout = "0.14.0"',),
    REPO_ROOT / "crates/rdocx-opc/README.md": (
        'rdocx-opc = "0.14.0"',
        "use rdocx_opc::OpcPackage;",
    ),
    REPO_ROOT / "crates/rdocx-oxml/README.md": ('rdocx-oxml = "0.14.0"',),
    REPO_ROOT / "crates/rdocx-pdf/README.md": (
        'rdocx-pdf = "0.14.0"',
        "use rdocx_pdf::render_to_pdf;",
    ),
    REPO_ROOT / "crates/oxml-cli-support/README.md": (
        'oxml_cli_support::parse_range("2,4-6")?',
    ),
    REPO_ROOT / "crates/oxml-core/README.md": ("Length::inches(8.5)",),
    REPO_ROOT / "crates/oxml-drawing/README.md": ("Fill::from_xml",),
    REPO_ROOT / "crates/oxml-layout/README.md": ('Color::from_hex("3366CC")',),
    REPO_ROOT / "crates/oxml-media/README.md": ("resolve(b\"\\x89PNG",),
    REPO_ROOT / "crates/oxml-opc/README.md": ("ContentTypes::from_xml",),
    REPO_ROOT / "crates/oxml-pdf/README.md": ("render_to_pdf(&layout)",),
    REPO_ROOT / "crates/oxml-py-support/README.md": (
        "emu_from_inches(8.5)",
    ),
    REPO_ROOT / "crates/rdocx-py/README.md": (
        "doc.add_paragraph(\"Hello from Python\")",
        'doc.save("hello.docx")',
        "runs the Rust `rdocx` document engine locally",
        "presents a typed\nPython API",
        "Review workflows cover comments, tracked revisions, comparison",
        "Complete package saves preserve safe producer XML",
    ),
    REPO_ROOT / "crates/rdocx-wasm/README.md": (
        'from "@tensorbee/rdocx-wasm"',
        "doc.toDocxBytes()",
        "wasm-pack build --target bundler --scope tensorbee crates/rdocx-wasm --out-name rdocx_wasm",
    ),
    REPO_ROOT / "crates/rpptx/README.md": (
        "use rpptx::Presentation;",
        "Presentation::new()?",
    ),
    REPO_ROOT / "crates/oxml-chart/README.md": ("AxisId::new(10_000_001)?",),
    REPO_ROOT / "crates/rpptx-chart/README.md": ("AxisId::new(10_000_001)?",),
    REPO_ROOT / "crates/rpptx-cli/README.md": (
        "cargo install rpptx-cli --version '^0.12.1'",
        "rpptx convert deck.pptx --to pdf -o deck.pdf",
    ),
    REPO_ROOT / "crates/rpptx-py/README.md": (
        'Presentation("deck.pptx")',
        "len(presentation.slides)",
        "presentation.to_pdf()",
        "presentation.render_slide_to_png(0)",
        "Read speaker-note text and inspect or mutate modern comment threads.",
        "runs the Rust `rpptx` presentation engine locally",
        "exposes a typed\nPython API",
        "renders slides and speaker notes without a remote service",
        "preserves safe package content",
    ),
    REPO_ROOT / "crates/rpptx-layout/README.md": ("ScopedMediaIds::default()",),
    REPO_ROOT / "crates/rpptx-oxml/README.md": (
        "CT_Presentation::from_xml",
        "part-level parsing and",
    ),
    REPO_ROOT / "crates/rpptx-render/README.md": (
        "RelScopes::default()",
        "does not write PDF or raster files",
    ),
    REPO_ROOT / "crates/rpptx-wasm/README.md": (
        'from "@tensorbee/rpptx-wasm"',
        "deck.toBytes()",
        "wasm-pack build --target bundler --scope tensorbee crates/rpptx-wasm --out-name rpptx_wasm",
    ),
}

ROOT_WORKFLOW_CLAIMS = {
    "DOCX": "Create, open, edit, validate, and save complete packages, with encryption and signing through opt-in features",
    "Rich authoring": "Complete paragraph and run properties, multilingual and vertical typography, conditional and floating tables, section page semantics, settings, styles, numbering, fields, forms, equations, drawings, comments, and metadata",
    "Preservation": "Retain unknown safe producer XML byte for byte when it is not modelled",
    "Native layout": "Resolve Word flow content into positioned pages with bundled, system, embedded, or caller-provided fonts",
    "Fixed output": "PDF, PDF/A, PNG, JPEG, TIFF, and SVG",
    "Flow output": "HTML, HTML fragments, Markdown, MHTML, RTF, ODT, and EPUB",
    "Automation": "Rust facade, CLI, Python binding, and locally built browser binding",
}
CAPABILITY_CLASSIFICATIONS = {
    "complete",
    "partial",
    "unsupported",
    "preserve-only",
    "permanent-non-goal",
}
COMPARISON_ROWS = {
    "rdocx": (
        "Open, create, edit, and save",
        "Unknown safe producer XML is retained byte for byte when it is not modelled",
        "Yes, Word flow layout with deterministic bundled fonts",
        "PDF and page images",
        "HTML and Markdown export",
        "`rdocx-cli`",
        "`rdocx-py`, with a narrower facade",
        "Workspace `rdocx-wasm` facade, deliberately unpublished",
    ),
    "python-docx": (
        "Create, open, change, and save",
        "Existing content that its API cannot manipulate is left alone on load and save. No byte-exact guarantee is stated",
        "ND",
        "ND",
        "ND",
        "ND",
        "Primary API",
        "ND",
    ),
    "docx-rs": (
        "Create and parse into the model used by its writer. Editing a parsed document is not separately documented",
        "ND. The project states that its OOXML support is not exhaustive",
        "ND",
        "ND",
        "ND",
        "ND",
        "ND",
        "Document generation and DOCX-to-JSON parsing through WebAssembly",
    ),
    "docx4j": (
        "Open, create, edit, and save",
        "ND for unknown XML or byte-exact round trips",
        "No project-owned page engine is documented. PDF paths use XSL-FO with Apache FOP, Microsoft Word through documents4j, or Microsoft Graph",
        "PDF through the documented conversion paths. Raster output is ND",
        "HTML export and first-party Markdown import and export",
        "ND",
        "ND",
        "ND",
    ),
    "Aspose.Words": (
        "Create, load, modify, and save DOCX",
        "Feature-level preservation is documented during conversion. No byte-exact unknown-XML contract is stated",
        "Yes, its own page-layout engine",
        "PDF plus PNG, JPEG, BMP, and TIFF",
        "HTML and Markdown",
        "ND",
        "Official Python via .NET API",
        "ND",
    ),
}
COMPARISON_EVIDENCE = (
    "https://github.com/tensorbee/rdocx",
    "https://python-docx.readthedocs.io/en/stable/user/documents.html",
    "https://github.com/bokuweb/docx-rs",
    "https://github.com/plutext/docx4j",
    "https://www.docx4java.org/blog/2020/09/office-pptxxlsxdocx-to-pdf-to-in-docx4j-8-2-3/",
    "https://github.com/plutext/docx4j/tree/VERSION_17_1_1/docx4j-markdown",
    "https://docs.aspose.com/words/python-net/product-overview/",
    "https://docs.aspose.com/words/python-net/converting-to-fixed-page-format/",
    "https://docs.aspose.com/words/python-net/supported-document-formats/",
    "https://docs.aspose.com/words/python-net/supported-features-on-document-load/",
)
COMPARISON_EVIDENCE_USAGE = COMPARISON_EVIDENCE + (
    "https://python-docx.readthedocs.io/en/stable/user/documents.html",
    "https://github.com/bokuweb/docx-rs",
)
ROOT_UNIQUENESS_CLAIM = (
    "Among these reviewed projects, rdocx alone documents the complete combination\n"
    "of a native Rust API, project-owned Word layout, fixed and flow outputs, a CLI,\n"
    "Python, and a browser surface. That statement is bounded to the official\n"
    "evidence set and review date."
)
MeasurementRow = tuple[str, str, str, str, str, str, str, str, str]
MEASUREMENT_COLUMNS = (
    "Measurement",
    "Value",
    "Version",
    "Platform",
    "Build mode",
    "Input",
    "Command",
    "Statistic",
    "Measured on",
)
MEASUREMENT_DATE = "2026-09-19"
MEASUREMENT_PLATFORM = "macOS 26.6.2, Apple M5 Max, arm64"
ARCHIVE_COMPRESSION_TOLERANCE_BYTES = 64
ARCHIVE_MEASUREMENTS = {
    "oxml-chart": (102_042, 659_367, 6),
    "oxml-cli-support": (6_718, 21_586, 6),
    "oxml-core": (20_677, 100_124, 15),
    "oxml-drawing": (159_703, 1_121_577, 24),
    "oxml-layout": (4_623_324, 9_227_483, 51),
    "oxml-media": (12_252, 50_992, 6),
    "oxml-opc": (93_152, 360_300, 12),
    "oxml-pdf": (66_015, 304_432, 14),
    "oxml-sml": (12_511, 49_803, 6),
    "rdocx": (1_085_772, 6_463_361, 36),
    "rdocx-cli": (33_805, 145_256, 8),
    "rdocx-html": (15_486, 63_894, 11),
    "rdocx-layout": (251_567, 1_368_637, 15),
    "rdocx-opc": (3_655, 9_668, 6),
    "rdocx-oxml": (367_792, 2_380_607, 32),
    "rdocx-pdf": (8_111, 26_758, 6),
    "rpptx": (389_176, 2_023_170, 16),
    "rpptx-chart": (6_648, 21_136, 6),
    "rpptx-cli": (27_236, 108_831, 8),
    "rpptx-layout": (79_109, 458_112, 11),
    "rpptx-oxml": (152_243, 1_038_449, 20),
    "rpptx-render": (57_790, 320_838, 8),
}
PACKAGE_VERSIONS = {
    **{name: "0.12.1" for name, _ in LOCAL_PATCHES if not name.startswith("rdocx")},
    **{name: "0.14.0" for name, _ in LOCAL_PATCHES if name.startswith("rdocx")},
}
PERFORMANCE_OBSERVATIONS = {
    "layout-throughput": "31,019.1 pages/s",
    "layout-peak": "29.03 MiB",
    "pdf-throughput": "60,058.0 pages/s",
    "pdf-peak": "1.73 MiB",
}
PERFORMANCE_COMMAND = (
    "`cargo test -p rdocx --test regression_test --release "
    "a_thousand_page_document_paginates_and_renders_within_the_declared_limits "
    "-- --ignored --exact --nocapture --test-threads=1`"
)


def archive_row(package: str) -> MeasurementRow:
    compressed, members, count = ARCHIVE_MEASUREMENTS[package]
    return (
        f"Crates.io archive: {package}",
        f"{compressed:,} compressed bytes, {members:,} member bytes, {count} members",
        PACKAGE_VERSIONS[package],
        MEASUREMENT_PLATFORM,
        "`cargo package --locked --no-verify`",
        f"Tracked `{package}` package inventory",
        "`python3 scripts/readme_doctests.py --record-measurements`",
        "gzip archive bytes, tar member bytes, tar member count",
        MEASUREMENT_DATE,
    )


MEASUREMENT_ROWS: dict[str, MeasurementRow] = {
    **{f"archive:{name}": archive_row(name) for name in ARCHIVE_MEASUREMENTS},
    "layout-throughput": (
        "Large-document layout throughput",
        f"minimum 250 pages/s, observed {PERFORMANCE_OBSERVATIONS['layout-throughput']}",
        "rdocx 0.14.0",
        MEASUREMENT_PLATFORM,
        "release, one test thread",
        "1,000 one-page paragraphs with deterministic fonts",
        PERFORMANCE_COMMAND,
        "pages per wall-clock second",
        MEASUREMENT_DATE,
    ),
    "layout-peak": (
        "Large-document layout peak allocation",
        f"maximum 64 MiB, observed {PERFORMANCE_OBSERVATIONS['layout-peak']}",
        "rdocx 0.14.0",
        MEASUREMENT_PLATFORM,
        "release, one test thread",
        "1,000 one-page paragraphs with deterministic fonts",
        PERFORMANCE_COMMAND,
        "peak live allocation",
        MEASUREMENT_DATE,
    ),
    "pdf-throughput": (
        "Large-document PDF throughput",
        f"minimum 1,000 pages/s, observed {PERFORMANCE_OBSERVATIONS['pdf-throughput']}",
        "rdocx 0.14.0",
        MEASUREMENT_PLATFORM,
        "release, one test thread",
        "1,000 deterministic layout pages",
        PERFORMANCE_COMMAND,
        "pages per wall-clock second",
        MEASUREMENT_DATE,
    ),
    "pdf-peak": (
        "Large-document PDF peak allocation",
        f"maximum 16 MiB, observed {PERFORMANCE_OBSERVATIONS['pdf-peak']}",
        "rdocx 0.14.0",
        MEASUREMENT_PLATFORM,
        "release, one test thread",
        "1,000 deterministic layout pages",
        PERFORMANCE_COMMAND,
        "peak live allocation",
        MEASUREMENT_DATE,
    ),
}
SPEED_MEASUREMENTS = (
    "layout-throughput",
    "layout-peak",
    "pdf-throughput",
    "pdf-peak",
)
MEASUREMENT_PAGES: dict[Path, tuple[str, ...]] = {
    **{
        (REPO_ROOT / "README.md" if name == "rdocx" else REPO_ROOT / path / "README.md"):
        (f"archive:{name}",)
        for name, path in LOCAL_PATCHES
    },
}
for measurement_page in (
    REPO_ROOT / "README.md",
    REPO_ROOT / "crates/rdocx-layout/README.md",
    REPO_ROOT / "crates/oxml-pdf/README.md",
    REPO_ROOT / "crates/rdocx-py/README.md",
):
    MEASUREMENT_PAGES[measurement_page] = (
        *MEASUREMENT_PAGES.get(measurement_page, ()),
        *SPEED_MEASUREMENTS,
    )
MEASUREMENT_TIERS = {
    **{f"archive:{name}": "rederived" for name in ARCHIVE_MEASUREMENTS},
    **{measurement_id: "gated" for measurement_id in SPEED_MEASUREMENTS},
}
MEASUREMENT_TIER_NAMES = ("rederived", "gated", "recorded")
UNBOUNDED_MEASUREMENT_CLAIMS = re.compile(
    r"\b(fastest|smallest|lightest|every library|any other library|all other|"
    r"industry-leading|unmatched|best-in-class)\b",
    re.IGNORECASE,
)
RELATIVE_MEASUREMENT_CLAIM = re.compile(
    r"\b(?:faster|smaller) than\b(?P<target>.*?)"
    r"(?:\.(?=\s|$)|[!?](?=\s|$)|\n|$)",
    re.IGNORECASE,
)
NAMED_COMPARISON_SUBJECT = re.compile(
    r"`[^`]+`|\[[^]]+\]\([^)]+\)|\b(?:rdocx|rpptx|python-docx|docx-rs|"
    r"docx4j|Aspose\.Words)\b|\b[vV]?\d+(?:\.\d+)+\b",
    re.IGNORECASE,
)
DEFERRED_MEASUREMENT_MARKERS = (
    "Python wheel and source-distribution sizes",
    "installed Python site-packages footprint",
    "CLI release archive sizes",
    "WASM bundle sizes",
    "Python boundary timing",
)
MARKDOWN_LINK = re.compile(r"(?<!!)\[[^]]+\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")
MARKDOWN_IMAGE = re.compile(r"!\[[^]]*\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")


def markdown_destinations(text: str) -> tuple[str, ...]:
    images = tuple(MARKDOWN_IMAGE.findall(text))
    text_without_images = MARKDOWN_IMAGE.sub("image", text)
    return images + tuple(MARKDOWN_LINK.findall(text_without_images))


def markdown_section(text: str, heading: str) -> str | None:
    marker = f"## {heading}\n"
    if text.count(marker) != 1:
        return None
    section = text.split(marker, 1)[1]
    return section.split("\n## ", 1)[0]


def measurement_table(text: str) -> tuple[MeasurementRow, ...] | None:
    section = markdown_section(text, "Measured footprint and speed")
    if section is None:
        return None
    rows: list[MeasurementRow] = []
    header_seen = False
    for line in section.splitlines():
        if not line.startswith("|"):
            continue
        cells = tuple(cell.strip() for cell in line.strip("|").split("|"))
        if cells == MEASUREMENT_COLUMNS:
            header_seen = True
            continue
        if cells and all(set(cell) <= {"-", ":"} for cell in cells):
            continue
        if len(cells) != len(MEASUREMENT_COLUMNS):
            return ()
        rows.append(cells)  # type: ignore[arg-type]
    return tuple(rows) if header_seen else ()


def validate_unbounded_claims(
    readmes: set[Path], overrides: dict[Path, str] | None = None
) -> bool:
    overrides = {} if overrides is None else overrides
    valid = True
    for readme in sorted(readmes):
        text = overrides.get(readme, readme.read_text(encoding="utf-8"))
        unbounded = UNBOUNDED_MEASUREMENT_CLAIMS.search(text)
        if unbounded is not None:
            print(
                f"README doctest error: {readme} contains unbounded claim "
                f"{unbounded.group(0)!r}",
                file=sys.stderr,
            )
            valid = False
        for comparison in RELATIVE_MEASUREMENT_CLAIM.finditer(text):
            if NAMED_COMPARISON_SUBJECT.search(comparison.group("target")) is None:
                print(
                    f"README doctest error: {readme} contains unbounded claim "
                    f"{comparison.group(0)!r}",
                    file=sys.stderr,
                )
                valid = False
    return valid


def valid_measurement_date(value: str) -> bool:
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value) is None:
        return False
    try:
        datetime.date.fromisoformat(value)
    except ValueError:
        return False
    return True


def performance_thresholds(source: str | None = None) -> dict[str, float]:
    if source is None:
        source = (
            REPO_ROOT / "crates/rdocx/tests/regression_test.rs"
        ).read_text(encoding="utf-8")
    patterns = {
        "layout-throughput": r"MIN_LAYOUT_PAGES_PER_SECOND: f64 = ([\d_]+(?:\.\d+)?)",
        "pdf-throughput": r"MIN_PDF_PAGES_PER_SECOND: f64 = ([\d_]+(?:\.\d+)?)",
        "layout-peak": r"MAX_LAYOUT_PEAK_BYTES: usize = ([\d_]+) \* MIB",
        "pdf-peak": r"MAX_PDF_PEAK_BYTES: usize = ([\d_]+) \* MIB",
    }
    thresholds: dict[str, float] = {}
    for measurement_id, pattern in patterns.items():
        match = re.search(pattern, source)
        if match is None:
            return {}
        thresholds[measurement_id] = float(match.group(1).replace("_", ""))
    return thresholds


def validate_speed_bounds(
    rows: dict[str, MeasurementRow] | None = None,
    thresholds: dict[str, float] | None = None,
) -> bool:
    rows = MEASUREMENT_ROWS if rows is None else rows
    thresholds = performance_thresholds() if thresholds is None else thresholds
    valid = True
    for measurement_id in SPEED_MEASUREMENTS:
        row = rows.get(measurement_id)
        if row is None:
            return False
        if measurement_id.endswith("throughput"):
            match = re.match(r"minimum ([\d,]+(?:\.\d+)?) pages/s, observed ", row[1])
            claim = None if match is None else float(match.group(1).replace(",", ""))
            bounded = claim is not None and claim <= thresholds.get(measurement_id, -1)
        else:
            match = re.match(r"maximum ([\d,]+(?:\.\d+)?) MiB, observed ", row[1])
            claim = None if match is None else float(match.group(1).replace(",", ""))
            bounded = claim is not None and claim >= thresholds.get(measurement_id, float("inf"))
        if not bounded:
            print(
                f"README doctest error: {measurement_id} outruns its code gate",
                file=sys.stderr,
            )
            valid = False
    return valid


def expected_measurement_version(
    measurement_id: str, versions: dict[str, str]
) -> str | None:
    if measurement_id.startswith("archive:"):
        return versions.get(measurement_id.removeprefix("archive:"))
    version = versions.get("rdocx")
    return None if version is None else f"rdocx {version}"


def validate_measurement_evidence(
    overrides: dict[Path, str] | None = None,
    metadata: dict[str, object] | None = None,
) -> bool:
    overrides = {} if overrides is None else overrides
    metadata = cargo_metadata() if metadata is None else metadata
    packages = None if metadata is None else metadata.get("packages")
    if not isinstance(packages, list):
        print(
            "README doctest error: invalid metadata for measurement evidence",
            file=sys.stderr,
        )
        return False
    versions = {
        package.get("name"): package.get("version")
        for package in packages
        if isinstance(package, dict)
        and isinstance(package.get("name"), str)
        and isinstance(package.get("version"), str)
    }
    readmes = {
        readme
        for package in packages
        if isinstance(package, dict)
        for readme in (package_readme(package),)
        if readme is not None
    }
    valid = validate_unbounded_claims(readmes, overrides)
    for readme in sorted(readmes):
        text = overrides.get(readme, readme.read_text(encoding="utf-8"))
        expected_ids = MEASUREMENT_PAGES.get(readme, ())
        table = measurement_table(text)
        if not expected_ids:
            if table is not None:
                print(
                    f"README doctest error: {readme} has an unapproved measurement table",
                    file=sys.stderr,
                )
                valid = False
            continue
        if table is None or len(table) != len(expected_ids):
            print(
                f"README doctest error: {readme} measurement row count differs "
                f"from approved ids {expected_ids!r}",
                file=sys.stderr,
            )
            valid = False
            continue
        expected_rows = tuple(MEASUREMENT_ROWS[row_id] for row_id in expected_ids)
        if table != expected_rows:
            print(
                f"README doctest error: {readme} measurement rows differ from approved evidence",
                file=sys.stderr,
            )
            valid = False
        for measurement_id, row in zip(expected_ids, table):
            if any(not cell for cell in row):
                print(
                    f"README doctest error: {readme} has an incomplete measurement row",
                    file=sys.stderr,
                )
                valid = False
            if not valid_measurement_date(row[8]):
                print(
                    f"README doctest error: {readme} has a non-ISO measurement date",
                    file=sys.stderr,
                )
                valid = False
            expected_version = expected_measurement_version(measurement_id, versions)
            if expected_version is None or row[2] != expected_version:
                print(
                    f"README doctest error: {readme} measurement version is stale",
                    file=sys.stderr,
                )
                valid = False

    thresholds = performance_thresholds()
    if thresholds != {
        "layout-throughput": 250.0,
        "pdf-throughput": 1_000.0,
        "layout-peak": 64.0,
        "pdf-peak": 16.0,
    }:
        print(
            "README doctest error: performance gate constants differ from approved bounds",
            file=sys.stderr,
        )
        valid = False
    if not validate_speed_bounds(thresholds=thresholds):
        valid = False

    if set(MEASUREMENT_TIERS) != set(MEASUREMENT_ROWS):
        print("README doctest error: measurement tiers are incomplete", file=sys.stderr)
        valid = False
    if set(MEASUREMENT_TIERS.values()) - set(MEASUREMENT_TIER_NAMES):
        print("README doctest error: unknown measurement tier", file=sys.stderr)
        valid = False
    if "recorded" in MEASUREMENT_TIERS.values():
        print(
            "README doctest error: deferred recorded measurements must stay absent",
            file=sys.stderr,
        )
        valid = False
    backlog = (REPO_ROOT / "docs/hld/14-development-backlog.md").read_text(
        encoding="utf-8"
    )
    for marker in DEFERRED_MEASUREMENT_MARKERS:
        if marker not in backlog:
            print(
                f"README doctest error: deferred measurement is untracked: {marker}",
                file=sys.stderr,
            )
            valid = False
    forbidden_deferred = re.compile(
        r"\b(wheel|sdist|site-packages|CLI release archive|WASM bundle|"
        r"Python boundary)\b[^\n|]*\b\d+(?:\.\d+)?\s*(?:bytes?|KiB|MiB|ms|s)\b",
        re.IGNORECASE,
    )
    for readme in sorted(readmes):
        text = overrides.get(readme, readme.read_text(encoding="utf-8"))
        if forbidden_deferred.search(text):
            print(
                f"README doctest error: {readme} publishes a deferred measurement",
                file=sys.stderr,
            )
            valid = False
    return valid


def capability_matrix(text: str) -> dict[str, str]:
    matrix: dict[str, str] = {}
    for line in text.splitlines():
        if not line.startswith("| DOCX-"):
            continue
        cells = tuple(cell.strip() for cell in line.strip("|").split("|"))
        if len(cells) != 19 or cells[16] not in CAPABILITY_CLASSIFICATIONS:
            continue
        matrix[cells[0]] = cells[16]
    return matrix


def validate_root_capability_claims(readme: str, matrix_text: str) -> bool:
    section = markdown_section(readme, "Built for complete document workflows")
    matrix = capability_matrix(matrix_text)
    if section is None or not matrix:
        print(
            "README doctest error: missing capability section or approved matrix",
            file=sys.stderr,
        )
        return False
    observed: dict[str, str] = {}
    for line in section.splitlines():
        if not line.startswith("|"):
            continue
        cells = tuple(cell.strip() for cell in line.strip("|").split("|"))
        if len(cells) != 2 or cells[0] in {"Workflow", "---"}:
            continue
        if cells[0] in observed:
            print(f"README doctest error: duplicate workflow {cells[0]}", file=sys.stderr)
            return False
        observed[cells[0]] = cells[1]
    if observed != ROOT_WORKFLOW_CLAIMS:
        print(
            "README doctest error: root workflow claims differ from the approved "
            f"summary, observed={observed!r}",
            file=sys.stderr,
        )
        return False
    required_matrix_state = {
        "DOCX-001": "complete",
        "DOCX-002": "complete",
        "DOCX-080": "preserve-only",
        "DOCX-082": "permanent-non-goal",
        "DOCX-083": "permanent-non-goal",
    }
    if any(matrix.get(key) != value for key, value in required_matrix_state.items()):
        print(
            "README doctest error: approved matrix no longer supports the root boundary",
            file=sys.stderr,
        )
        return False
    return True


def validate_opt_in_security_claims(
    root_readme: str, rpptx_readme: str, metadata: dict[str, object]
) -> bool:
    packages = metadata.get("packages")
    if not isinstance(packages, list):
        print("README doctest error: invalid metadata for opt-in features", file=sys.stderr)
        return False
    package_metadata = {
        package.get("name"): package
        for package in packages
        if isinstance(package, dict)
    }
    expected = {"rdocx": root_readme, "rpptx": rpptx_readme}
    valid = True
    for package_name, readme in expected.items():
        package = package_metadata.get(package_name)
        features = package.get("features") if isinstance(package, dict) else None
        version = package.get("version") if isinstance(package, dict) else None
        defaults = features.get("default") if isinstance(features, dict) else None
        dependency = (
            f'{package_name} = {{ version = "{version}", features = '
            '["agile-encryption", "digital-signatures"] }'
        )
        if (
            not isinstance(version, str)
            or not isinstance(features, dict)
            or "agile-encryption" not in features
            or "digital-signatures" not in features
            or not isinstance(defaults, list)
            or "agile-encryption" in defaults
            or "digital-signatures" in defaults
            or readme.count(dependency) != 1
        ):
            print(
                f"README doctest error: {package_name} security features are not "
                "documented as exact opt-ins",
                file=sys.stderr,
            )
            valid = False
    return valid


def validate_root_narrative(readme: str) -> bool:
    headings = (
        "## Built for complete document workflows",
        "## Examples",
        "## Installation",
        "## Evidence-based alternatives",
        "## Honest boundaries",
    )
    positions = tuple(readme.find(heading) for heading in headings)
    required_intro = (
        "one `Document` for the complete Word workflow",
        "integrated native document stack",
        "without an Office installation or conversion service",
    )
    valid = all(position >= 0 for position in positions)
    valid = valid and positions == tuple(sorted(positions))
    valid = valid and all(item in readme[: positions[0]] for item in required_intro)
    valid = valid and "## Capability status" not in readme
    valid = valid and "## Project status" not in readme
    if not valid:
        print(
            "README doctest error: root README does not lead with the approved product narrative",
            file=sys.stderr,
        )
    return valid


def validate_crate_narratives(overrides: dict[Path, str] | None = None) -> bool:
    overrides = {} if overrides is None else overrides
    metadata = cargo_metadata()
    if metadata is None or not isinstance(metadata.get("packages"), list):
        print("README doctest error: invalid metadata for crate narratives", file=sys.stderr)
        return False
    valid = True
    for package in metadata["packages"]:
        if not isinstance(package, dict) or package.get("name") == "rdocx":
            continue
        readme = package_readme(package)
        if readme is None or not readme.is_file():
            valid = False
            continue
        text = overrides.get(readme, readme.read_text(encoding="utf-8"))
        headings = ("## Capabilities", "## Use it when", "## Relationship", "## Example")
        positions = tuple(text.find(heading) for heading in headings)
        section = markdown_section(text, "Capabilities")
        bullet_count = 0 if section is None else sum(
            line.startswith("- ") for line in section.splitlines()
        )
        if any(position < 0 for position in positions) or positions != tuple(
            sorted(positions)
        ) or bullet_count < 5:
            print(
                f"README doctest error: {readme} lacks the capability-led crate narrative",
                file=sys.stderr,
            )
            valid = False
    forbidden = {
        REPO_ROOT / "crates/rpptx-py/README.md": ("reading, editing, and rendering",),
        REPO_ROOT / "crates/rpptx-render/README.md": ("layout, raster, and PDF output",),
        REPO_ROOT / "crates/rpptx-oxml/README.md": ("package-preserving parse",),
        REPO_ROOT / "crates/rdocx-oxml/README.md": ("## Migrating from 0.4",),
    }
    for readme, claims in forbidden.items():
        text = overrides.get(readme, readme.read_text(encoding="utf-8"))
        if any(claim in text for claim in claims):
            print(
                f"README doctest error: {readme} retains a rejected capability claim",
                file=sys.stderr,
            )
            valid = False
    return valid


def validate_root_versions(readme: str, metadata: dict[str, object]) -> bool:
    packages = metadata.get("packages")
    if not isinstance(packages, list):
        print("README doctest error: invalid Cargo metadata", file=sys.stderr)
        return False
    versions = {
        package.get("name"): package.get("version")
        for package in packages
        if isinstance(package, dict)
    }
    rdocx_version = versions.get("rdocx")
    cli_version = versions.get("rdocx-cli")
    if not isinstance(rdocx_version, str) or not isinstance(cli_version, str):
        print(
            "README doctest error: Cargo metadata lacks the rdocx package family",
            file=sys.stderr,
        )
        return False
    requirements = (
        f'rdocx = "{rdocx_version}"',
        f'rdocx = {{ version = "{rdocx_version}", default-features = false }}',
        f"cargo install rdocx-cli --version '^{cli_version}'",
    )
    valid = True
    for requirement in requirements:
        if readme.count(requirement) != 1:
            print(
                "README doctest error: expected one metadata-derived root "
                f"requirement {requirement!r}",
                file=sys.stderr,
            )
            valid = False
    return valid


def markdown_anchors(text: str) -> set[str]:
    anchors: set[str] = set()
    for line in text.splitlines():
        match = re.match(r"^#{1,6}\s+(.+?)\s*#*$", line)
        if match is None:
            continue
        heading = re.sub(r"[`*_~]", "", match.group(1)).strip().lower()
        heading = re.sub(r"[^\w\- ]", "", heading)
        anchors.add(re.sub(r"[ ]+", "-", heading))
    return anchors


def validate_local_links(readme: str, source: Path | None = None) -> bool:
    source = REPO_ROOT / "README.md" if source is None else source
    valid = True
    for destination in markdown_destinations(readme):
        parsed = urlsplit(destination)
        if parsed.scheme or parsed.netloc:
            continue
        relative_path = unquote(parsed.path)
        target = (
            (source.parent / relative_path).resolve()
            if relative_path
            else source
        )
        try:
            target.relative_to(REPO_ROOT)
        except ValueError:
            print(
                f"README doctest error: local link escapes repository {destination!r}",
                file=sys.stderr,
            )
            valid = False
            continue
        if not target.is_file():
            print(
                f"README doctest error: local link target is missing {destination!r}",
                file=sys.stderr,
            )
            valid = False
            continue
        if parsed.fragment and target.suffix.lower() == ".md":
            anchors = markdown_anchors(target.read_text(encoding="utf-8"))
            if unquote(parsed.fragment).lower() not in anchors:
                print(
                    f"README doctest error: local anchor is missing {destination!r}",
                    file=sys.stderr,
                )
                valid = False
    return valid


def validate_all_local_links(overrides: dict[Path, str] | None = None) -> bool:
    overrides = {} if overrides is None else overrides
    metadata = cargo_metadata()
    if metadata is None or not isinstance(metadata.get("packages"), list):
        print("README doctest error: invalid metadata for README links", file=sys.stderr)
        return False
    readmes = {
        readme
        for package in metadata["packages"]
        if isinstance(package, dict)
        for readme in (package_readme(package),)
        if readme is not None
    }
    return all(
        validate_local_links(
            overrides.get(readme, readme.read_text(encoding="utf-8")), readme
        )
        for readme in sorted(readmes)
    )


def validate_comparison_evidence(readme: str) -> bool:
    section = markdown_section(readme, "Evidence-based alternatives")
    if section is None:
        print("README doctest error: missing alternatives section", file=sys.stderr)
        return False
    volatile = re.compile(
        r"\b(downloads?|stars?|pricing?|prices?|costs?|memory|cold start|"
        r"binary size|install size|faster|fastest|most popular)\b",
        re.IGNORECASE,
    )
    if volatile.search(section):
        print(
            "README doctest error: alternatives contain a volatile claim",
            file=sys.stderr,
        )
        return False
    urls = tuple(
        destination
        for destination in markdown_destinations(section)
        if urlsplit(destination).scheme
    )
    if Counter(urls) != Counter(COMPARISON_EVIDENCE_USAGE):
        print(
            "README doctest error: comparison links differ from approved "
            f"official evidence, observed={urls!r}",
            file=sys.stderr,
        )
        return False
    rows: list[tuple[str, tuple[str, ...]]] = []
    for line in section.splitlines():
        if not line.startswith("|"):
            continue
        cells = tuple(cell.strip() for cell in line.strip("|").split("|"))
        if len(cells) != 9:
            print(
                "README doctest error: comparison table has a malformed row",
                file=sys.stderr,
            )
            return False
        if cells[0] in {"Project", "---"}:
            continue
        product = re.sub(r"^\[([^]]+)\]\([^)]+\)$", r"\1", cells[0])
        rows.append((product, cells[1:]))
    if tuple(product for product, _ in rows) != tuple(COMPARISON_ROWS):
        print(
            "README doctest error: comparison projects differ from the approved set",
            file=sys.stderr,
        )
        return False
    for product, cells in rows:
        if cells != COMPARISON_ROWS[product]:
            print(
                f"README doctest error: unsupported comparison claim for {product}",
                file=sys.stderr,
            )
            return False
    if section.count(ROOT_UNIQUENESS_CLAIM) != 1:
        print(
            "README doctest error: scoped comparison conclusion changed",
            file=sys.stderr,
        )
        return False
    return True


def check_official_links() -> bool:
    valid = True
    for url in sorted(COMPARISON_EVIDENCE):
        request = Request(url, headers={"User-Agent": "rdocx-readme-check/1"})
        try:
            with urlopen(request, timeout=20) as response:
                status = response.status
        except OSError as error:
            print(
                "README doctest error: official evidence did not resolve: "
                f"{url}: {error}",
                file=sys.stderr,
            )
            valid = False
            continue
        if not 200 <= status < 400:
            print(
                f"README doctest error: official evidence returned {status}: {url}",
                file=sys.stderr,
            )
            valid = False
    return valid


def validate_fences(readme: Path, expected: int) -> bool:
    text = readme.read_text(encoding="utf-8")
    attributes = RUST_FENCE.findall(text)
    if len(attributes) != expected or any(
        attribute != ",no_run" for attribute in attributes
    ):
        print(
            f"README doctest error: expected {expected} "
            f"exact rust,no_run fences, found {len(attributes)} with "
            f"attributes {attributes!r}",
            file=sys.stderr,
        )
        return False
    return True


def cargo_metadata() -> dict[str, object] | None:
    result = subprocess.run(
        ["cargo", "metadata", "--no-deps", "--format-version", "1"],
        cwd=REPO_ROOT,
        stdout=subprocess.PIPE,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        return None
    return json.loads(result.stdout)


def package_readme(package: dict[str, object]) -> Path | None:
    value = package.get("readme")
    manifest = package.get("manifest_path")
    if not isinstance(value, str) or not isinstance(manifest, str):
        return None
    return (Path(manifest).parent / value).resolve()


def validate_local_patches(packages: list[object]) -> bool:
    expected: set[tuple[str, str]] = set()
    for package in packages:
        if not isinstance(package, dict) or package.get("publish") == []:
            continue
        name = package.get("name")
        manifest = package.get("manifest_path")
        if not isinstance(name, str) or not isinstance(manifest, str):
            print(
                "README doctest error: invalid publishable package metadata",
                file=sys.stderr,
            )
            return False
        try:
            package_path = Path(manifest).parent.resolve().relative_to(REPO_ROOT)
        except ValueError:
            print(
                f"README doctest error: {name} is outside the repository",
                file=sys.stderr,
            )
            return False
        expected.add((name, package_path.as_posix()))

    actual = set(LOCAL_PATCHES)
    if len(actual) != len(LOCAL_PATCHES) or actual != expected:
        missing = sorted(expected - actual)
        unexpected = sorted(actual - expected)
        print(
            "README doctest error: local patches differ from publishable "
            f"metadata, missing={missing!r}, unexpected={unexpected!r}",
            file=sys.stderr,
        )
        return False
    return True


def build_package_archive(package: dict[str, object]) -> Path | None:
    name = package["name"]
    version = package.get("version")
    if not isinstance(name, str) or not isinstance(version, str):
        print("README doctest error: invalid package identity", file=sys.stderr)
        return None
    command = [
        "cargo",
        "package",
        "--locked",
        "--allow-dirty",
        "--no-verify",
        "-p",
        name,
    ]
    for patch_name, patch_path in LOCAL_PATCHES:
        command.extend(
            ["--config", f'patch.crates-io.{patch_name}.path="{patch_path}"']
        )
    result = subprocess.run(
        command,
        cwd=REPO_ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr, end="")
        return None
    archive = REPO_ROOT / f"target/package/{name}-{version}.crate"
    if not archive.is_file():
        print(
            f"README doctest error: missing generated archive {archive}",
            file=sys.stderr,
        )
        return None
    return archive


def archive_measurement(archive: Path) -> tuple[int, int, int]:
    with tarfile.open(archive, "r:gz") as package_archive:
        members = package_archive.getmembers()
        member_bytes = 0
        for member in members:
            if not member.name.endswith("/.cargo_vcs_info.json"):
                member_bytes += member.size
                continue
            extracted = package_archive.extractfile(member)
            if extracted is None:
                member_bytes += member.size
                continue
            vcs_info = json.loads(extracted.read())
            git = vcs_info.get("git")
            if isinstance(git, dict):
                git.pop("dirty", None)
                git["sha1"] = "0" * 40
            member_bytes += len(json.dumps(vcs_info, indent=2).encode())
    return archive.stat().st_size, member_bytes, len(members)


def recorded_archive_measurement(row: MeasurementRow) -> tuple[int, int, int] | None:
    match = re.fullmatch(
        r"([\d,]+) compressed bytes, ([\d,]+) member bytes, (\d+) members",
        row[1],
    )
    if match is None:
        return None
    return tuple(int(value.replace(",", "")) for value in match.groups())  # type: ignore[return-value]


def pinned_rustc_is_active() -> bool:
    toolchain = tomllib.loads(
        (REPO_ROOT / "rust-toolchain.toml").read_text(encoding="utf-8")
    )["toolchain"]["channel"]
    result = subprocess.run(
        ["rustc", "--version"],
        cwd=REPO_ROOT,
        stdout=subprocess.PIPE,
        text=True,
        check=False,
    )
    return result.returncode == 0 and result.stdout.startswith(f"rustc {toolchain} ")


def validate_archive_measurement(package: str, archive: Path) -> bool:
    observed = archive_measurement(archive)
    recorded = recorded_archive_measurement(MEASUREMENT_ROWS[f"archive:{package}"])
    if recorded is None:
        print(
            f"README doctest error: malformed archive measurement for {package}",
            file=sys.stderr,
        )
        return False
    observed_compressed, observed_members, observed_count = observed
    recorded_compressed, recorded_members, recorded_count = recorded
    valid = (
        observed_members == recorded_members
        and observed_count == recorded_count
        and observed_compressed < 10 * 1024 * 1024
    )
    if pinned_rustc_is_active():
        valid = valid and (
            abs(observed_compressed - recorded_compressed)
            <= ARCHIVE_COMPRESSION_TOLERANCE_BYTES
        )
    else:
        valid = valid and (
            observed_compressed
            <= recorded_compressed + ARCHIVE_COMPRESSION_TOLERANCE_BYTES
        )
    if not valid:
        print(
            f"README doctest error: {package} archive measurement differs, "
            f"recorded={recorded!r}, observed={observed!r}",
            file=sys.stderr,
        )
    return valid


def validate_package_archive(package: dict[str, object], readme: Path) -> bool:
    name = package["name"]
    if not isinstance(name, str):
        print("README doctest error: invalid package name", file=sys.stderr)
        return False
    archive = build_package_archive(package)
    if archive is None:
        return False
    with tarfile.open(archive, "r:gz") as package_archive:
        readmes = [
            member
            for member in package_archive.getmembers()
            if Path(member.name).name == "README.md"
        ]
        if len(readmes) != 1:
            print(
                f"README doctest error: {name} archive has {len(readmes)} "
                "README files",
                file=sys.stderr,
            )
            return False
        packaged = package_archive.extractfile(readmes[0])
        if packaged is None or packaged.read() != readme.read_bytes():
            print(
                f"README doctest error: {name} archive README differs from "
                f"{readme}",
                file=sys.stderr,
            )
            return False
    return validate_archive_measurement(name, archive)


def validate_inventory() -> bool:
    metadata = cargo_metadata()
    if metadata is None:
        return False
    packages = metadata.get("packages")
    if not isinstance(packages, list) or len(packages) != WORKSPACE_PACKAGE_COUNT:
        observed = len(packages) if isinstance(packages, list) else "invalid"
        print(
            f"README doctest error: expected {WORKSPACE_PACKAGE_COUNT} workspace "
            f"packages, found {observed}",
            file=sys.stderr,
        )
        return False

    valid = True
    readme_paths: set[Path] = set()
    publishable: list[tuple[dict[str, object], Path]] = []
    for package in packages:
        if not isinstance(package, dict) or not isinstance(package.get("name"), str):
            print("README doctest error: invalid Cargo package metadata", file=sys.stderr)
            valid = False
            continue
        name = package["name"]
        readme = package_readme(package)
        manifest_value = package.get("manifest_path")
        if not isinstance(manifest_value, str):
            print(
                f"README doctest error: {name} lacks a manifest path",
                file=sys.stderr,
            )
            valid = False
            continue
        manifest = Path(manifest_value)
        manifest_data = tomllib.loads(manifest.read_text(encoding="utf-8"))
        declared_readme = manifest_data.get("package", {}).get("readme")
        if not isinstance(declared_readme, str):
            print(
                f"README doctest error: {manifest} lacks an explicit package.readme",
                file=sys.stderr,
            )
            valid = False
        if readme is None or not readme.is_file():
            print(
                f"README doctest error: {name} does not declare an existing README",
                file=sys.stderr,
            )
            valid = False
            continue
        readme_paths.add(readme)
        text = readme.read_text(encoding="utf-8")
        if not text.startswith(f"# {name}\n"):
            print(
                f"README doctest error: {readme} must start with '# {name}'",
                file=sys.stderr,
            )
            valid = False
        if name != "rdocx":
            for heading in ("## Use it when", "## Relationship", "## Example"):
                if heading not in text:
                    print(
                        f"README doctest error: {readme} lacks {heading!r}",
                        file=sys.stderr,
                    )
                    valid = False
        if not EXAMPLE_FENCE.search(text):
            print(
                f"README doctest error: {readme} lacks a supported example fence",
                file=sys.stderr,
            )
            valid = False
        if package.get("publish") != []:
            publishable.append((package, readme))

    if len(readme_paths) != WORKSPACE_PACKAGE_COUNT:
        print(
            f"README doctest error: expected {WORKSPACE_PACKAGE_COUNT} distinct "
            f"README sources, found {len(readme_paths)}",
            file=sys.stderr,
        )
        valid = False
    if len(publishable) != PUBLISHABLE_PACKAGE_COUNT:
        print(
            f"README doctest error: expected {PUBLISHABLE_PACKAGE_COUNT} "
            f"publishable packages, found {len(publishable)}",
            file=sys.stderr,
        )
        valid = False
    if not validate_local_patches(packages):
        valid = False

    root_readme = (REPO_ROOT / "README.md").read_text(encoding="utf-8")
    capability_source = (
        REPO_ROOT / "docs/hld/02-scope-and-non-goals.md"
    ).read_text(encoding="utf-8")
    if not validate_root_capability_claims(root_readme, capability_source):
        valid = False
    if not validate_root_narrative(root_readme):
        valid = False
    if not validate_crate_narratives():
        valid = False
    if not validate_root_versions(root_readme, metadata):
        valid = False
    rpptx_readme = (REPO_ROOT / "crates/rpptx/README.md").read_text(encoding="utf-8")
    if not validate_opt_in_security_claims(root_readme, rpptx_readme, metadata):
        valid = False
    if not validate_all_local_links():
        valid = False
    if not validate_comparison_evidence(root_readme):
        valid = False
    if not validate_measurement_evidence(metadata=metadata):
        valid = False

    for readme, required_items in README_REQUIRED_TEXT.items():
        text = readme.read_text(encoding="utf-8")
        for required in required_items:
            if required not in text:
                print(
                    f"README doctest error: {readme} does not contain "
                    f"{required!r}",
                    file=sys.stderr,
                )
                valid = False
    if valid:
        for package, readme in sorted(
            publishable, key=lambda item: str(item[0]["name"])
        ):
            if not validate_package_archive(package, readme):
                valid = False
    if valid:
        print(
            f"readme_doctests: {WORKSPACE_PACKAGE_COUNT} distinct workspace "
            f"READMEs and {PUBLISHABLE_PACKAGE_COUNT} publishable package "
            "inventories validated"
        )
    return valid


def markdown_measurement_row(row: MeasurementRow) -> str:
    return "| " + " | ".join(row) + " |"


def record_measurements() -> bool:
    metadata = cargo_metadata()
    packages = None if metadata is None else metadata.get("packages")
    if not isinstance(packages, list):
        print("README doctest error: invalid metadata for recording", file=sys.stderr)
        return False
    thresholds = performance_thresholds()
    if not validate_speed_bounds(thresholds=thresholds):
        return False
    versions = {
        package.get("name"): package.get("version")
        for package in packages
        if isinstance(package, dict)
    }
    rows: list[MeasurementRow] = []
    for package in sorted(
        (
            package
            for package in packages
            if isinstance(package, dict) and package.get("publish") != []
        ),
        key=lambda package: str(package.get("name")),
    ):
        name = package.get("name")
        if not isinstance(name, str):
            return False
        archive = build_package_archive(package)
        if archive is None:
            return False
        compressed, members, count = archive_measurement(archive)
        row = list(archive_row(name))
        row[1] = (
            f"{compressed:,} compressed bytes, {members:,} member bytes, "
            f"{count} members"
        )
        version = versions.get(name)
        if not isinstance(version, str):
            return False
        row[2] = version
        rows.append(tuple(row))  # type: ignore[arg-type]
    rows.extend(MEASUREMENT_ROWS[row_id] for row_id in SPEED_MEASUREMENTS)
    print("| " + " | ".join(MEASUREMENT_COLUMNS) + " |")
    print("|" + "|".join("---" for _ in MEASUREMENT_COLUMNS) + "|")
    for row in rows:
        print(markdown_measurement_row(row))
    return True


def build_rlibs(package: str, crate_names: tuple[str, ...]) -> dict[str, Path] | None:
    command = [
        "cargo",
        "build",
        "--locked",
        "-p",
        package,
        "--message-format=json-render-diagnostics",
    ]
    result = subprocess.run(
        command,
        cwd=REPO_ROOT,
        stdout=subprocess.PIPE,
        text=True,
        check=False,
    )
    artifacts: dict[str, set[Path]] = {name: set() for name in crate_names}
    for line in result.stdout.splitlines():
        try:
            message = json.loads(line)
        except json.JSONDecodeError:
            continue
        if message.get("reason") == "compiler-message":
            rendered = message.get("message", {}).get("rendered")
            if rendered:
                print(rendered, file=sys.stderr, end="")
        if message.get("reason") != "compiler-artifact":
            continue
        target = message.get("target", {})
        crate_name = target.get("name")
        if crate_name not in artifacts or "lib" not in target.get("crate_types", []):
            continue
        artifacts[crate_name].update(
            Path(filename).resolve()
            for filename in message.get("filenames", [])
            if filename.endswith(".rlib")
        )
    if result.returncode != 0:
        return None
    for crate_name, crate_artifacts in artifacts.items():
        if len(crate_artifacts) != 1:
            print(
                f"README doctest error: expected one {crate_name} rlib from "
                f"the {package} build, found "
                f"{sorted(str(path) for path in crate_artifacts)!r}",
                file=sys.stderr,
            )
            return None
    return {name: paths.pop() for name, paths in artifacts.items()}


def compile_readme(case: ReadmeCase) -> bool:
    if not validate_fences(case.readme, case.expected_rust_fences):
        return False

    crate_names = (case.crate_name,) + tuple(
        crate_name for _, crate_name in case.companion_crates
    )
    rlibs = build_rlibs(case.package, crate_names)
    if rlibs is None:
        return False
    resolved_rlibs = list(rlibs.items())
    rlib = rlibs[case.crate_name]
    dependency_dir = rlib.parent / "deps"
    if not dependency_dir.is_dir():
        dependency_dir = rlib.parent
    command = [
        "rustdoc",
        "--test",
        str(case.readme),
        "--crate-name",
        f"{case.crate_name}_readme",
        "--edition=2024",
        "-Dwarnings",
        "-L",
        f"dependency={dependency_dir}",
    ]
    for crate_name, crate_rlib in resolved_rlibs:
        command.extend(("--extern", f"{crate_name}={crate_rlib}"))
    result = subprocess.run(
        command,
        cwd=REPO_ROOT,
        check=False,
    )
    if result.returncode == 0:
        print(
            f"readme_doctests: {case.expected_rust_fences} Rust examples "
            f"compiled from {case.readme}"
        )
    return result.returncode == 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("readme", nargs="?", type=Path)
    parser.add_argument("--check-official-links", action="store_true")
    parser.add_argument("--record-measurements", action="store_true")
    args = parser.parse_args()
    if args.record_measurements:
        return 0 if record_measurements() else 1
    if args.check_official_links:
        return 0 if check_official_links() else 1
    if args.readme is not None:
        case = ReadmeCase("rdocx", "rdocx", args.readme.resolve(), 3)
        return 0 if compile_readme(case) else 1

    if not validate_inventory():
        return 1
    return 0 if all(compile_readme(case) for case in README_CASES) else 1


if __name__ == "__main__":
    raise SystemExit(main())
