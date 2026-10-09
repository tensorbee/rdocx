import importlib.metadata
import hashlib
import os
import re
import struct
import subprocess
import urllib.request
import zipfile
import zlib
from pathlib import Path

import pytest


ORACLE_DISTRIBUTION = "python-docx"
ORACLE_VERSION = "1.2.0"
TAGGED_SOURCE_ROOT = (
    "https://raw.githubusercontent.com/python-openxml/python-docx/"
    "v1.2.0/docs/user"
)
TAGGED_SOURCE_PAGES = {
    "documents": f"{TAGGED_SOURCE_ROOT}/documents.rst",
    "quickstart": f"{TAGGED_SOURCE_ROOT}/quickstart.rst",
    "text": f"{TAGGED_SOURCE_ROOT}/text.rst",
}
HELD_ROW_SOURCE_BODY = (
    "row = table.rows[1]\n"
    "row.cells[0].text = 'Foo bar to you.'\n"
    "row.cells[1].text = 'And a hearty foo bar to you too sir!'\n"
)
HELD_ROW_RDOCX_BODY = (
    "row = table.rows[1]\n"
    "row.cells[0].text = 'Foo bar to you.'\n"
    "row = document.tables[0].rows[1]\n"
    "row.cells[1].text = 'And a hearty foo bar to you too sir!'\n"
)

EXPECTED_MANIFEST_IDS = frozenset(
    {
        "documents.opening-a-document",
        "documents.really-opening-a-document",
        "quickstart.opening-a-document",
        "quickstart.adding-a-paragraph",
        "quickstart.adding-a-table",
        "quickstart.table-cell-text",
        "quickstart.held-row-two-cell-assignment",
        "quickstart.table-iteration",
        "quickstart.table-style",
        "quickstart.adding-runs",
        "text.horizontal-alignment",
        "text.indentation",
        "text.line-spacing",
        "text.paragraph-spacing",
        "text.pagination-properties",
        "text.font-name-and-size",
        "text.font-tristate-and-underline",
    }
)

DOCUMENTED_S33_EXAMPLES = (
    {
        "id": "documents.opening-a-document",
        "page": TAGGED_SOURCE_PAGES["documents"],
        "heading": "Opening a document",
        "body": "from docx import Document\n\ndocument = Document()\ndocument.save('test.docx')\n",
        "transformation": "namespace-only",
        "setup": "empty",
        "observation": "files",
        "expected": ("test.docx",),
    },
    {
        "id": "documents.really-opening-a-document",
        "page": TAGGED_SOURCE_PAGES["documents"],
        "heading": "REALLY opening a document",
        "body": "document = Document('existing-document-file.docx')\ndocument.save('new-file-name.docx')\n",
        "transformation": "namespace-only",
        "setup": "existing-document",
        "observation": "paragraph-texts",
        "expected": ("existing",),
    },
    {
        "id": "quickstart.opening-a-document",
        "page": TAGGED_SOURCE_PAGES["quickstart"],
        "heading": "Opening a document",
        "body": "from docx import Document\n\ndocument = Document()\n",
        "transformation": "namespace-only",
        "setup": "empty",
        "observation": "paragraph-texts",
        "expected": (),
    },
    {
        "id": "quickstart.adding-a-paragraph",
        "page": TAGGED_SOURCE_PAGES["quickstart"],
        "heading": "Adding a paragraph",
        "body": "paragraph = document.add_paragraph('Lorem ipsum dolor sit amet.')\n",
        "transformation": "namespace-only",
        "setup": "document",
        "observation": "paragraph-texts",
        "expected": ("Lorem ipsum dolor sit amet.",),
    },
    {
        "id": "quickstart.adding-a-table",
        "page": TAGGED_SOURCE_PAGES["quickstart"],
        "heading": "Adding a table",
        "body": "table = document.add_table(rows=2, cols=2)\n",
        "transformation": "namespace-only",
        "setup": "document",
        "observation": "table-shape",
        "expected": ((2, 2),),
    },
    {
        "id": "quickstart.table-cell-text",
        "page": TAGGED_SOURCE_PAGES["quickstart"],
        "heading": "Adding a table",
        "body": "cell = table.cell(0, 1)\ncell.text = 'parrot, possibly dead'\n",
        "transformation": "namespace-only",
        "setup": "table",
        "observation": "table-cells",
        "expected": (("", "parrot, possibly dead"), ("", "")),
    },
    {
        "id": "quickstart.held-row-two-cell-assignment",
        "page": TAGGED_SOURCE_PAGES["quickstart"],
        "heading": "Adding a table",
        "body": (
            "row = table.rows[1]\n"
            "row.cells[0].text = 'Foo bar to you.'\n"
            "row.cells[1].text = 'And a hearty foo bar to you too sir!'\n"
        ),
        "transformation": "namespace-and-held-row-refetch",
        "setup": "table",
        "observation": "table-cells",
        "expected": (
            ("", ""),
            ("Foo bar to you.", "And a hearty foo bar to you too sir!"),
        ),
    },
    {
        "id": "quickstart.table-iteration",
        "page": TAGGED_SOURCE_PAGES["quickstart"],
        "heading": "Adding a table",
        "body": "for row in table.rows:\n    for cell in row.cells:\n        print(cell.text)\n",
        "transformation": "namespace-only",
        "setup": "populated-table",
        "observation": "table-cells",
        "expected": (("alpha", "beta"), ("gamma", "delta")),
    },
    {
        "id": "quickstart.table-style",
        "page": TAGGED_SOURCE_PAGES["quickstart"],
        "heading": "Adding a table",
        "body": "table.style = 'LightShading-Accent1'\n",
        "transformation": "namespace-only",
        "setup": "table",
        "observation": "table-style",
        "expected": "LightShading-Accent1",
    },
    {
        "id": "quickstart.adding-runs",
        "page": TAGGED_SOURCE_PAGES["quickstart"],
        "heading": "Applying bold and italic",
        "body": "paragraph = document.add_paragraph('Lorem ipsum ')\nparagraph.add_run('dolor sit amet.')\n",
        "transformation": "namespace-only",
        "setup": "document",
        "observation": "runs",
        "expected": (("Lorem ipsum ", "dolor sit amet."),),
    },
    {
        "id": "text.horizontal-alignment",
        "page": TAGGED_SOURCE_PAGES["text"],
        "heading": "Horizontal alignment (justification)",
        "body": "from docx.enum.text import WD_ALIGN_PARAGRAPH\nparagraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER\n",
        "transformation": "namespace-only",
        "setup": "paragraph-format",
        "observation": "alignment",
        "expected": 1,
    },
    {
        "id": "text.indentation",
        "page": TAGGED_SOURCE_PAGES["text"],
        "heading": "Indentation",
        "body": "from docx.shared import Inches\nfrom docx.shared import Pt\nparagraph_format.left_indent = Inches(0.5)\nparagraph_format.right_indent = Pt(24)\nparagraph_format.first_line_indent = Inches(-0.25)\n",
        "transformation": "namespace-only",
        "setup": "paragraph-format",
        "observation": "indentation",
        "expected": (457200, 304800, -228600),
    },
    {
        "id": "text.line-spacing",
        "page": TAGGED_SOURCE_PAGES["text"],
        "heading": "Line spacing",
        "body": "paragraph_format.line_spacing = Pt(18)\nparagraph_format.line_spacing = 1.75\n",
        "transformation": "namespace-only",
        "setup": "paragraph-format",
        "observation": "line-spacing",
        "expected": 1.75,
    },
    {
        "id": "text.paragraph-spacing",
        "page": TAGGED_SOURCE_PAGES["text"],
        "heading": "Paragraph spacing",
        "body": "paragraph_format.space_before = Pt(18)\nparagraph_format.space_after = Pt(12)\n",
        "transformation": "namespace-only",
        "setup": "paragraph-format",
        "observation": "spacing",
        "expected": (228600, 152400),
    },
    {
        "id": "text.pagination-properties",
        "page": TAGGED_SOURCE_PAGES["text"],
        "heading": "Pagination properties",
        "body": "paragraph_format.keep_with_next = True\nparagraph_format.page_break_before = False\n",
        "transformation": "namespace-only",
        "setup": "paragraph-format",
        "observation": "pagination",
        "expected": (True, False),
    },
    {
        "id": "text.font-name-and-size",
        "page": TAGGED_SOURCE_PAGES["text"],
        "heading": "Apply character formatting",
        "body": "from docx.shared import Pt\nfont.name = 'Calibri'\nfont.size = Pt(12)\n",
        "transformation": "namespace-only",
        "setup": "font",
        "observation": "font-name-and-size",
        "expected": ("Calibri", 152400),
    },
    {
        "id": "text.font-tristate-and-underline",
        "page": TAGGED_SOURCE_PAGES["text"],
        "heading": "Apply character formatting",
        "body": "font.italic = True\nfont.italic = False\nfont.italic = None\nfont.underline = True\nfont.underline = WD_UNDERLINE.DOT_DASH\n",
        "transformation": "namespace-only",
        "setup": "font",
        "observation": "font-tristate-and-underline",
        "expected": (None, 9),
    },
)


def _assert_oracle_version():
    assert importlib.metadata.version(ORACLE_DISTRIBUTION) == ORACLE_VERSION


def _namespace_only(body, namespace):
    if namespace == "docx":
        return body
    assert namespace == "rdocx"
    transformed = body.replace("from docx ", "from rdocx ")
    transformed = transformed.replace("from docx.", "from rdocx.")
    assert transformed.replace("from rdocx ", "from docx ").replace(
        "from rdocx.", "from docx."
    ) == body
    return transformed


def _example_body(example, namespace):
    transformed = _namespace_only(example["body"], namespace)
    if namespace == "docx" or example["transformation"] == "namespace-only":
        return transformed
    assert example["id"] == "quickstart.held-row-two-cell-assignment"
    assert example["transformation"] == "namespace-and-held-row-refetch"
    assert transformed == HELD_ROW_SOURCE_BODY
    return HELD_ROW_RDOCX_BODY


def _fresh_cell(document, row, column):
    return document.tables[0].rows[row].cells[column]


def _example_context(setup, root):
    import rdocx

    namespace = {
        "Document": rdocx.Document,
        "Pt": rdocx.Pt,
        "WD_UNDERLINE": rdocx.WD_UNDERLINE,
    }
    if setup == "empty":
        return namespace

    document = rdocx.Document()
    namespace["document"] = document
    if setup == "existing-document":
        source = rdocx.Document()
        source.add_paragraph("existing")
        source.save(root / "existing-document-file.docx")
        return namespace
    if setup == "document":
        return namespace

    if setup in ("table", "populated-table"):
        table = document.add_table(rows=2, cols=2)
        namespace["table"] = table
    if setup == "table":
        return namespace
    if setup == "populated-table":
        for row, values in enumerate((("alpha", "beta"), ("gamma", "delta"))):
            for column, value in enumerate(values):
                _fresh_cell(document, row, column).text = value
        namespace["table"] = document.tables[0]
        return namespace

    paragraph = document.add_paragraph("")
    if setup == "paragraph-format":
        namespace["paragraph_format"] = paragraph.paragraph_format
        return namespace
    if setup == "font":
        run = paragraph.add_run("")
        namespace["font"] = run.font
        return namespace
    raise AssertionError(f"unknown example setup: {setup}")


def _table_cells(document):
    return tuple(
        tuple(cell.text for cell in row.cells) for row in document.tables[0].rows
    )


def _example_observation(example, namespace, root):
    observation = example["observation"]
    if observation == "files":
        return tuple(sorted(path.name for path in root.glob("*.docx")))
    document = namespace["document"]
    if observation == "paragraph-texts":
        return tuple(paragraph.text for paragraph in document.paragraphs)
    if observation == "table-shape":
        return tuple(
            (len(table.rows), len(table.rows[0].cells)) for table in document.tables
        )
    if observation == "table-cells":
        return _table_cells(document)
    if observation == "table-style":
        return document.tables[0].style
    if observation == "runs":
        return tuple(
            tuple(run.text for run in paragraph.runs)
            for paragraph in document.paragraphs
        )
    paragraph = document.paragraphs[-1]
    paragraph_format = paragraph.paragraph_format
    if observation == "alignment":
        value = paragraph_format.alignment
        return int(value) if value is not None else None
    if observation == "indentation":
        return (
            int(paragraph_format.left_indent),
            int(paragraph_format.right_indent),
            int(paragraph_format.first_line_indent),
        )
    if observation == "line-spacing":
        return paragraph_format.line_spacing
    if observation == "spacing":
        return (int(paragraph_format.space_before), int(paragraph_format.space_after))
    if observation == "pagination":
        return (
            paragraph_format.keep_with_next,
            paragraph_format.page_break_before,
        )
    font = paragraph.runs[-1].font
    if observation == "font-name-and-size":
        return (font.name, int(font.size))
    if observation == "font-tristate-and-underline":
        underline = font.underline
        return (font.italic, int(underline) if underline is not None else None)
    raise AssertionError(f"unknown example observation: {observation}")


def _optional_int(value):
    return int(value) if value is not None else None


def _optional_enum(value):
    if value is None or isinstance(value, bool):
        return value
    return int(value)


def _line_spacing_record(value):
    if value is None:
        return None
    if isinstance(value, float):
        return ("relative", value)
    return ("length", int(value))


def _paragraph_record(paragraph, oracle):
    paragraph_format = paragraph.paragraph_format
    runs = []
    for run in paragraph.runs:
        font = run.font
        color = font.color.rgb
        runs.append(
            (
                run.text,
                font.name,
                _optional_int(font.size),
                tuple(color) if color is not None else None,
                font.bold,
                font.italic,
                _optional_enum(font.underline),
                font.strike,
            )
        )
    return (
        paragraph.text,
        _optional_enum(paragraph_format.alignment),
        _optional_int(paragraph_format.space_before),
        _optional_int(paragraph_format.space_after),
        _optional_int(paragraph_format.left_indent),
        _optional_int(paragraph_format.right_indent),
        _optional_int(paragraph_format.first_line_indent),
        _line_spacing_record(paragraph_format.line_spacing),
        paragraph_format.keep_with_next,
        paragraph_format.keep_together,
        paragraph_format.page_break_before,
        paragraph_format.widow_control,
        tuple(runs),
    )


def _table_style(table, oracle):
    if table.style is None:
        return None
    style = table.style.style_id if oracle else table.style
    return None if oracle and style == "TableNormal" else style


def _document_record(document, oracle):
    paragraphs = tuple(
        _paragraph_record(paragraph, oracle) for paragraph in document.paragraphs
    )
    tables = []
    for table in document.tables:
        rows = []
        for row in table.rows:
            rows.append(
                tuple(
                    (
                        cell.text,
                        _optional_int(cell.width),
                        _optional_enum(cell.vertical_alignment),
                        tuple(
                            _paragraph_record(paragraph, oracle)
                            for paragraph in cell.paragraphs
                        ),
                    )
                    for cell in row.cells
                )
            )
        tables.append(
            (
                _table_style(table, oracle),
                _optional_enum(table.alignment),
                tuple(rows),
            )
        )
    return paragraphs, tuple(tables)


def _author_parity_document(
    Document,
    Inches,
    Pt,
    RGBColor,
    WD_ALIGN_PARAGRAPH,
    WD_UNDERLINE,
    WD_TABLE_ALIGNMENT,
    WD_CELL_VERTICAL_ALIGNMENT,
    set_color,
    source_path,
    path,
):
    document = Document(source_path)
    paragraph = document.add_paragraph("Alpha ")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(3)
    paragraph_format.space_after = Pt(6)
    paragraph_format.left_indent = Inches(0.5)
    paragraph_format.right_indent = Inches(0.25)
    paragraph_format.first_line_indent = Inches(-0.25)
    paragraph_format.line_spacing = Pt(18)
    paragraph_format.keep_with_next = True
    paragraph_format.keep_together = False
    paragraph_format.page_break_before = False
    paragraph_format.widow_control = True
    run = paragraph.add_run("beta")
    font = run.font
    font.name = "Aptos"
    font.size = Pt(12)
    set_color(font, RGBColor(0x12, 0x34, 0x56))
    font.bold = True
    font.italic = False
    font.underline = WD_UNDERLINE.DOT_DASH
    font.strike = True

    second = document.add_paragraph("Second")
    second.paragraph_format.line_spacing = 1.75
    table = document.add_table(rows=2, cols=2)
    table.style = "LightShading-Accent1"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row, values in enumerate((("one", "two"), ("three", "four"))):
        for column, value in enumerate(values):
            cell = document.tables[0].rows[row].cells[column]
            cell.text = value
            cell = document.tables[0].rows[row].cells[column]
            cell.width = Inches(2)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    cell = document.tables[0].rows[0].cells[0]
    nested = cell.add_paragraph("nested")
    nested.paragraph_format.space_after = Pt(6)
    document.save(path)


def test_documented_s33_examples_run_with_declared_transformations(
    tmp_path, monkeypatch
):
    _assert_oracle_version()
    assert TAGGED_SOURCE_ROOT == (
        "https://raw.githubusercontent.com/python-openxml/python-docx/"
        "v1.2.0/docs/user"
    )
    assert {example["id"] for example in DOCUMENTED_S33_EXAMPLES} == (
        EXPECTED_MANIFEST_IDS
    )
    assert len(DOCUMENTED_S33_EXAMPLES) == len(EXPECTED_MANIFEST_IDS)
    for example in DOCUMENTED_S33_EXAMPLES:
        section = example["id"].split(".", 1)[0]
        assert example["page"] == TAGGED_SOURCE_PAGES[section]
        expected_transformation = (
            "namespace-and-held-row-refetch"
            if example["id"] == "quickstart.held-row-two-cell-assignment"
            else "namespace-only"
        )
        assert example["transformation"] == expected_transformation

    held_row = next(
        example
        for example in DOCUMENTED_S33_EXAMPLES
        if example["id"] == "quickstart.held-row-two-cell-assignment"
    )
    assert held_row["body"] == HELD_ROW_SOURCE_BODY
    assert _example_body(held_row, "docx") == HELD_ROW_SOURCE_BODY
    assert _example_body(held_row, "rdocx") == HELD_ROW_RDOCX_BODY
    line_spacing = next(
        example
        for example in DOCUMENTED_S33_EXAMPLES
        if example["id"] == "text.line-spacing"
    )
    assert line_spacing["body"] == (
        "paragraph_format.line_spacing = Pt(18)\n"
        "paragraph_format.line_spacing = 1.75\n"
    )

    for example in DOCUMENTED_S33_EXAMPLES:
        root = tmp_path / example["id"]
        root.mkdir()
        monkeypatch.chdir(root)
        namespace = _example_context(example["setup"], root)
        exec(_example_body(example, "rdocx"), namespace)
        assert _example_observation(example, namespace, root) == example["expected"]


def test_rdocx_and_python_docx_round_trip_the_same_normalized_content(tmp_path):
    _assert_oracle_version()

    import rdocx
    from docx import Document as OracleDocument
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT as OracleCellAlignment
    from docx.enum.table import WD_TABLE_ALIGNMENT as OracleTableAlignment
    from docx.enum.text import WD_ALIGN_PARAGRAPH as OracleParagraphAlignment
    from docx.enum.text import WD_UNDERLINE as OracleUnderline
    from docx.shared import Inches as OracleInches
    from docx.shared import Pt as OraclePt
    from docx.shared import RGBColor as OracleRGBColor

    assert int(rdocx.Inches(1)) == int(OracleInches(1)) == 914400
    assert int(rdocx.Pt(12)) == int(OraclePt(12)) == 152400
    assert int(rdocx.WD_ALIGN_PARAGRAPH.CENTER) == int(
        OracleParagraphAlignment.CENTER
    )
    assert int(rdocx.WD_UNDERLINE.DOT_DASH) == int(OracleUnderline.DOT_DASH)
    assert int(rdocx.WD_TABLE_ALIGNMENT.CENTER) == int(OracleTableAlignment.CENTER)
    assert int(rdocx.WD_CELL_VERTICAL_ALIGNMENT.BOTTOM) == int(
        OracleCellAlignment.BOTTOM
    )

    rdocx_path = tmp_path / "rdocx-authored.docx"
    oracle_path = tmp_path / "python-docx-authored.docx"
    source_path = tmp_path / "styles-source.docx"
    OracleDocument().save(source_path)
    _author_parity_document(
        rdocx.Document,
        rdocx.Inches,
        rdocx.Pt,
        rdocx.RGBColor,
        rdocx.WD_ALIGN_PARAGRAPH,
        rdocx.WD_UNDERLINE,
        rdocx.WD_TABLE_ALIGNMENT,
        rdocx.WD_CELL_VERTICAL_ALIGNMENT,
        lambda font, color: setattr(font, "color", color),
        source_path,
        rdocx_path,
    )
    _author_parity_document(
        OracleDocument,
        OracleInches,
        OraclePt,
        OracleRGBColor,
        OracleParagraphAlignment,
        OracleUnderline,
        OracleTableAlignment,
        OracleCellAlignment,
        lambda font, color: setattr(font.color, "rgb", color),
        source_path,
        oracle_path,
    )

    records = []
    for path in (rdocx_path, oracle_path):
        rdocx_record = _document_record(rdocx.Document(path), oracle=False)
        oracle_record = _document_record(OracleDocument(path), oracle=True)
        assert rdocx_record == oracle_record
        assert tuple(
            paragraph[7] for paragraph in rdocx_record[0]
        ) == (("length", 228600), ("relative", 1.75))
        assert rdocx_record[1][0][0] == "LightShading-Accent1"
        records.append(rdocx_record)
    assert records[0] == records[1]

# Issues 157, 159 and 160: the attached producer and identity matrices are
# built in source so every row exercises the same document operations.
_MATRIX_TIMESTAMP = "2026-09-27T12:00:00Z"
_MATRIX_W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
_MATRIX_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_MATRIX_IDENTITY_ROWS = (
    (None, None),
    ("paragraphs", "w:rsidR"),
    ("paragraphs", "w:rsidRDefault"),
    ("paragraphs", "w:rsidP"),
    ("paragraphs", "w14:paraId"),
    ("paragraphs", "w14:textId"),
    ("runs", "w:rsidR"),
    ("runs", "w:rsidRPr"),
    ("runs", "w:rsidDel"),
    ("field runs", "w:rsidR"),
    ("field runs", "w:rsidRPr"),
    ("footer field runs", "w:rsidR"),
    ("footer field runs", "w:rsidRPr"),
    ("content control", "w:id"),
    ("content control", "w:tag"),
    ("table rows", "w:rsidR"),
    ("table rows", "w:rsidTr"),
    ("table rows", "w14:paraId"),
)
_MATRIX_PRODUCER_ROWS = (
    "control: none",
    "xml:space=preserve on every w:t",
    'w:val="0" toggles on every run',
    "empty <w:pPr/> on plain paragraphs",
    'pageBreakBefore/keepNext w:val="0"',
    'w:orient="portrait" on pgSz',
    "default namespace on the root",
    "packed TOC field run",
    "packed footer fields, no cached result",
    "inline content control on first runs",
    "empty comments part",
)


def _matrix_field(paragraph, instruction, cached=None):
    from docx.oxml.ns import qn

    result = []
    for kind, text in (("begin", None), ("instr", instruction), ("separate", None)):
        run = paragraph.makeelement(qn("w:r"), {})
        if kind == "instr":
            item = run.makeelement(qn("w:instrText"), {})
            item.text = text
            item.set(_MATRIX_XML_SPACE, "preserve")
        else:
            item = run.makeelement(qn("w:fldChar"), {})
            item.set(qn("w:fldCharType"), kind)
        run.append(item)
        paragraph.append(run)
        result.append(run)
    if cached is not None:
        run = paragraph.makeelement(qn("w:r"), {})
        text = run.makeelement(qn("w:t"), {})
        text.text = cached
        run.append(text)
        paragraph.append(run)
        run = paragraph.makeelement(qn("w:r"), {})
        end = run.makeelement(qn("w:fldChar"), {})
        end.set(qn("w:fldCharType"), "end")
        run.append(end)
        paragraph.append(run)
        result.append(run)
    return result


def _matrix_fixture(path, where=None, attr=None, word="alpha"):
    from docx import Document as OracleDocument
    from docx.oxml.ns import qn

    document = OracleDocument()
    body = document.element.body
    field_runs = []
    entries = ["Chapter 1", "Section 1.1", "Chapter 2", "Section 2.1", "Chapter 3", "Section 3.1"]
    for index, title in enumerate(entries):
        paragraph = document.add_paragraph()
        if index == 0:
            field_runs.extend(_matrix_field(paragraph._p, ' TOC \\o "1-3" \\h \\z \\u '))
        paragraph.add_run(title + "\t1")
        if index == len(entries) - 1:
            run = paragraph._p.makeelement(qn("w:r"), {})
            end = run.makeelement(qn("w:fldChar"), {})
            end.set(qn("w:fldCharType"), "end")
            run.append(end)
            paragraph._p.append(run)
            field_runs.append(run)
    for index in range(1, 4):
        document.add_heading(f"Chapter {index}", level=1)
        document.add_paragraph(
            f"Body text of chapter {index}, lorem {word if index == 2 else 'ipsum'} dolor."
        )
        document.add_heading(f"Section {index}.1", level=2)
        document.add_paragraph(f"Body text of section {index}.1.")
    table = document.add_table(rows=2, cols=2)
    for cell in table._cells:
        cell.text = "cell"
    control = body.makeelement(qn("w:sdt"), {})
    properties = control.makeelement(qn("w:sdtPr"), {})
    tag = properties.makeelement(qn("w:tag"), {})
    tag.set(qn("w:val"), "block")
    properties.append(tag)
    control.append(properties)
    content = control.makeelement(qn("w:sdtContent"), {})
    for value in ("Inside the content control, one.", "Inside the content control, two."):
        paragraph = document.add_paragraph(value)._p
        body.remove(paragraph)
        content.append(paragraph)
    control.append(content)
    body.find(qn("w:tbl")).addnext(control)
    footer = document.sections[0].footer.paragraphs[0]._p
    run = footer.makeelement(qn("w:r"), {})
    text = run.makeelement(qn("w:t"), {})
    text.text = "Page "
    text.set(_MATRIX_XML_SPACE, "preserve")
    run.append(text)
    footer.append(run)
    footer_runs = _matrix_field(footer, " PAGE ", "1") + _matrix_field(footer, " NUMPAGES ", "1")
    targets = {
        "paragraphs": list(body.iter(qn("w:p"))),
        "runs": list(body.iter(qn("w:r"))),
        "field runs": field_runs,
        "footer field runs": footer_runs,
        "content control": [properties],
        "table rows": list(body.iter(qn("w:tr"))),
    }
    if where:
        assert targets[where], where
        for element in targets[where]:
            if attr == "w:tag":
                element.find(qn("w:tag")).set(qn("w:val"), "block-2")
            elif attr == "w:id":
                identity = element.makeelement(qn("w:id"), {})
                identity.set(qn("w:val"), "-2000000001")
                element.insert(0, identity)
            elif attr.startswith("w14:"):
                element.set(_MATRIX_W14 + attr[4:], "1A2B3C4D")
            else:
                element.set(qn(attr), "00A1B2C3")
    document.save(path)
    return path


def _matrix_part(path, part):
    with zipfile.ZipFile(path) as package:
        return package.read(part)


def _matrix_attribute_count(path, attr):
    data = _matrix_part(path, "word/document.xml") + _matrix_part(path, "word/footer1.xml")
    if attr == "w:tag":
        return data.count(b"<w:tag ")
    if attr == "w:id":
        return data.count(b"<w:id ")
    return data.count(b" " + attr.encode() + b"=")


def _matrix_comparison_count(original, edited):
    import rdocx

    compared = rdocx.Document(original)
    compared.compare(rdocx.Document(edited), "R", _MATRIX_TIMESTAMP)
    return len(compared.revisions)


@pytest.mark.parametrize("where,attr", _MATRIX_IDENTITY_ROWS)
def test_issue_159_identity_matrix_across_operations(tmp_path, where, attr):
    _assert_oracle_version()
    import rdocx

    plain = _matrix_fixture(tmp_path / "plain.docx")
    source = _matrix_fixture(tmp_path / "source.docx", where, attr)
    edited = _matrix_fixture(tmp_path / "edited.docx", where, attr, word="ALPHA")
    assert len(_MATRIX_IDENTITY_ROWS) == 18
    assert len(_MATRIX_IDENTITY_ROWS) * 7 == 126
    if attr is not None:
        assert _matrix_attribute_count(source, attr) > 0
    if attr == "w:tag":
        assert b'block-2' in _matrix_part(source, "word/document.xml")
    document = rdocx.Document(source)
    assert document.try_replace_text("Body text of section 3.1.", "Body text of section three.") == 1
    saved = tmp_path / "saved.docx"
    document.save(saved)
    if attr is not None:
        assert _matrix_attribute_count(source, attr) == _matrix_attribute_count(saved, attr)
    else:
        assert _matrix_part(source, "word/document.xml") != _matrix_part(saved, "word/document.xml")

    document = rdocx.Document(source)
    assert document.try_replace_text("lorem", "LOREM") == 3
    replaced = tmp_path / "replaced.docx"
    document.save(replaced)
    assert b"LOREM" in _matrix_part(replaced, "word/document.xml")
    assert rdocx.Document(source).rebuild_toc().entry_count == 6
    fields = rdocx.Document(source)
    fields.update_page_fields()
    refreshed = tmp_path / "refreshed.docx"
    fields.save(refreshed)
    assert len(re.findall(rb'fldCharType="separate"/>(?:</w:r><w:r>)?<w:t>([^<]*)</w:t>', _matrix_part(refreshed, "word/footer1.xml"))) == 2
    assert rdocx.Document(source).to_pdf().startswith(b"%PDF")
    assert _matrix_comparison_count(plain, source) == 0
    assert _matrix_comparison_count(source, edited) == 2


def _matrix_rewrite(path, trait):
    empty_comments = trait == "empty comments part"
    with zipfile.ZipFile(path) as package:
        items = [(item, package.read(item.filename)) for item in package.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for item, data in items:
            if item.filename == "word/document.xml":
                xml = data.decode()
                if trait == "xml:space=preserve on every w:t":
                    xml = xml.replace("<w:t>", '<w:t xml:space="preserve">')
                elif trait == 'w:val="0" toggles on every run':
                    xml = xml.replace("<w:r>", '<w:r><w:rPr><w:b w:val="0"/><w:i w:val="0"/></w:rPr>')
                elif trait == "empty <w:pPr/> on plain paragraphs":
                    xml = xml.replace("<w:p><w:r>", "<w:p><w:pPr/><w:r>")
                elif trait == 'pageBreakBefore/keepNext w:val="0"':
                    xml = xml.replace("<w:p><w:r>", '<w:p><w:pPr><w:keepNext w:val="0"/><w:pageBreakBefore w:val="0"/></w:pPr><w:r>')
                elif trait == 'w:orient="portrait" on pgSz':
                    xml = xml.replace("<w:pgSz ", '<w:pgSz w:orient="portrait" ')
                elif trait == "default namespace on the root":
                    xml = xml.replace("<w:document ", '<w:document xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks" ', 1)
                elif trait == "packed TOC field run":
                    xml = re.sub(
                        r'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r>(<w:instrText[^>]*>[^<]*</w:instrText>)</w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>',
                        r'<w:r><w:fldChar w:fldCharType="begin"/>\1<w:fldChar w:fldCharType="separate"/></w:r>',
                        xml,
                    )
                elif trait == "inline content control on first runs":
                    xml = re.sub(
                        r'(<w:p>(?:<w:pPr>.*?</w:pPr>)?)(<w:r>.*?</w:r>)',
                        r'\1<w:sdt><w:sdtPr><w:tag w:val="goog_rdk_0"/></w:sdtPr><w:sdtContent>\2</w:sdtContent></w:sdt>',
                        xml,
                    )
                data = xml.encode()
            elif item.filename == "word/footer1.xml" and trait == "packed footer fields, no cached result":
                xml = data.decode()
                xml = re.sub(
                    r'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r>(<w:instrText[^>]*>[^<]*</w:instrText>)</w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>',
                    r'<w:r><w:fldChar w:fldCharType="begin"/>\1<w:fldChar w:fldCharType="separate"/></w:r>',
                    xml,
                )
                xml = xml.replace(
                    '<w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
                    '<w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/></w:r>',
                )
                data = xml.encode()
            if empty_comments and item.filename == "[Content_Types].xml":
                data = data.replace(b"</Types>", b'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')
            if empty_comments and item.filename == "word/_rels/document.xml.rels":
                data = data.replace(b"</Relationships>", b'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')
            package.writestr(item, data)
        if empty_comments:
            package.writestr(
                "word/comments.xml",
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>',
            )
    _matrix_assert_trait(path, trait)
    return path


def _matrix_assert_trait(path, trait):
    empty_comments = trait == "empty comments part"
    document = _matrix_part(path, "word/document.xml")
    footer = _matrix_part(path, "word/footer1.xml")
    signatures = {
        "xml:space=preserve on every w:t": (document, b'xml:space="preserve">Body text'),
        'w:val="0" toggles on every run': (document, b'<w:b w:val="0"/>'),
        "empty <w:pPr/> on plain paragraphs": (document, b'<w:p><w:pPr/><w:r>'),
        'pageBreakBefore/keepNext w:val="0"': (document, b'<w:pageBreakBefore w:val="0"/>'),
        'w:orient="portrait" on pgSz': (document, b'<w:pgSz w:orient="portrait"'),
        "default namespace on the root": (document, b'xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks"'),
        "packed TOC field run": (document, b'fldCharType="begin"/><w:instrText'),
        "packed footer fields, no cached result": (footer, b'fldCharType="begin"/><w:instrText'),
        "inline content control on first runs": (document, b'goog_rdk_0'),
        "empty comments part": (_matrix_part(path, "word/comments.xml") if empty_comments else b"", b'<w:comments'),
    }
    if trait in signatures:
        payload, marker = signatures[trait]
        assert marker in payload, trait


def _matrix_png():
    def chunk(name, data):
        return struct.pack(">I", len(data)) + name + data + struct.pack(">I", zlib.crc32(name + data))

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00"))
        + chunk(b"IEND", b"")
    )


@pytest.mark.parametrize("trait", _MATRIX_PRODUCER_ROWS)
def test_issue_160_producer_matrix_across_operations_and_picture(tmp_path, trait):
    _assert_oracle_version()
    import rdocx

    assert len(_MATRIX_PRODUCER_ROWS) == 11
    assert len(_MATRIX_PRODUCER_ROWS) * 8 == 88

    source = _matrix_rewrite(_matrix_fixture(tmp_path / "source.docx"), trait)
    edited = _matrix_rewrite(_matrix_fixture(tmp_path / "edited.docx", word="ALPHA"), trait)
    document = rdocx.Document(source)
    noop = tmp_path / "noop.docx"
    document.save(noop)
    assert _matrix_part(source, "word/document.xml") == _matrix_part(noop, "word/document.xml")

    document = rdocx.Document(source)
    assert document.try_replace_text("alpha", "ALPHA") == 1
    replaced = tmp_path / "replaced.docx"
    document.save(replaced)
    assert rdocx.Document(source).rebuild_toc().entry_count == 6
    fields = rdocx.Document(source)
    fields.update_page_fields()
    refreshed = tmp_path / "refreshed.docx"
    fields.save(refreshed)
    footer = _matrix_part(refreshed, "word/footer1.xml").decode()
    assert _matrix_part(refreshed, "word/footer1.xml") == _matrix_part(source, "word/footer1.xml")
    expected_caches = 0 if trait == "packed footer fields, no cached result" else 2
    assert len(re.findall(r'fldCharType="separate"/>(?:</w:r><w:r>)?<w:t>([^<]*)</w:t>', footer)) == expected_caches
    assert rdocx.Document(source).to_pdf().startswith(b"%PDF")
    assert isinstance(_matrix_comparison_count(source, refreshed), int)
    assert _matrix_comparison_count(source, edited) == 2
    assert _matrix_comparison_count(source, replaced) == 2

    pictured = rdocx.Document(source)
    pictured.add_picture(_matrix_png(), "matrix.png", width=914400, height=914400)
    pictured_path = tmp_path / "pictured.docx"
    pictured.save(pictured_path)
    reopened = rdocx.Document(pictured_path)
    assert "word/media/image1.png" in zipfile.ZipFile(pictured_path).namelist()
    _matrix_assert_trait(pictured_path, trait)
    assert reopened.try_replace_text("alpha", "ALPHA") == 1
    assert reopened.to_pdf().startswith(b"%PDF")


_ISSUE158_REPORT_SHA256 = "d05f9c753c00eb804c6e345126ef7a1f7a4fc635d2c9b653b829922030cd875e"
_ISSUE158_REPORT_URL = "https://github.com/user-attachments/files/32701104/fixture-report.docx"


def _issue158_report(tmp_path):
    source = os.environ.get("RDOCX_ISSUE158_REPORT")
    if source:
        data = Path(source).read_bytes()
    else:
        with urllib.request.urlopen(_ISSUE158_REPORT_URL, timeout=30) as response:
            data = response.read()
    assert hashlib.sha256(data).hexdigest() == _ISSUE158_REPORT_SHA256
    path = tmp_path / "fixture-report.docx"
    path.write_bytes(data)
    return path


def test_issue_158_word_fixture_acceptance(tmp_path):
    _assert_oracle_version()
    import rdocx
    from docx import Document as OracleDocument

    assert len(_MATRIX_IDENTITY_ROWS) == 18
    assert len(_MATRIX_PRODUCER_ROWS) == 11
    assert (len(_MATRIX_IDENTITY_ROWS) * 7, len(_MATRIX_PRODUCER_ROWS) * 8) == (126, 88)
    report = _issue158_report(tmp_path)
    oracle = OracleDocument(report)
    assert len(oracle.paragraphs) == 90
    assert [(len(table.rows), len(table.columns)) for table in oracle.tables] == [
        (57, 6), (8, 5), (6, 2)
    ]
    document = rdocx.Document(report)
    assert [paragraph.text for paragraph in document.paragraphs[:2]] == [
        "Riverton Footbridge",
        "Principal inspection and condition survey, 2025",
    ]
    assert [len(table.rows) for table in document.tables] == [57, 8, 6]


def test_issue_158_complete_word_workflow(tmp_path):
    _assert_oracle_version()
    import rdocx

    source = _issue158_report(tmp_path)
    assert rdocx.Document(source).rebuild_toc().entry_count == 21
    document = rdocx.Document(source)
    word = "described"
    assert sum(paragraph.text.count(word) for paragraph in document.paragraphs) == 1
    probe = rdocx.Document.from_bytes(document.to_bytes())
    assert probe.try_replace_text(word, word.upper()) == 1
    assert document.try_replace_text(word, word.upper()) == 1

    xml = _matrix_part(source, "word/document.xml").decode()
    controls = re.findall(
        r'<w:sdt>(?:(?!</w:sdtPr>).)*?<w:tag w:val="([^"]+)"/>.*?</w:sdtPr><w:sdtContent>(.*?)</w:sdtContent>',
        xml,
    )
    assert [tag for tag, _ in controls] == ["goog_rdk_0", "goog_rdk_1"]
    for _, inner in controls:
        text = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", inner))
        probe = rdocx.Document.from_bytes(document.to_bytes())
        assert probe.try_replace_text(text.split(".")[0], "X") == 1

    index = next(i for i, paragraph in enumerate(document.paragraphs) if word.upper() in paragraph.text)
    body_index = document.find_content_index(document.paragraphs[index])
    document.clone_content(document.paragraphs[index], body_index + 1)
    for i, run in enumerate(document.paragraphs[index + 1].runs):
        run.text = "Inserted paragraph." if i == 0 else ""
    assert document.paragraphs[index + 1].text == "Inserted paragraph."
    table = document.tables[0]
    row_count = len(table.rows)
    table.clone_row(row_count - 1)
    for column in range(len(document.tables[0].rows[row_count].cells)):
        document.tables[0].rows[row_count].cells[column].text = "cloned"
    document.tables[0].remove_row(row_count)
    assert len(document.tables[0].rows) == row_count == 57

    first_table = document.find_content_index(document.tables[0])
    def direct_body_index(paragraph):
        try:
            return document.find_content_index(paragraph)
        except ValueError:
            return -1

    target_index = next(
        index for index, paragraph in enumerate(document.paragraphs)
        if len(paragraph.runs) == 1 and len(paragraph.text) > 60
        and direct_body_index(paragraph) > first_table
    )
    target = document.paragraphs[target_index]
    text = target.text
    body_index = document.find_content_index(target)
    document.split_run(body_index, 0, 30)
    document.split_run(body_index, 0, 10)
    assert document.paragraphs[target_index].runs[1].text == text[10:30]
    bounds = rdocx.RunRange(
        start=rdocx.RunPosition(body_index=body_index, run_index=1),
        end=rdocx.RunPosition(body_index=body_index, run_index=2),
    )
    comment_id = document.add_comment(bounds, author="Reviewer", text="Comment on a piece of text.", initials="R")
    document.reply_to(comment_id, author="Reviewer", text="Reply.")
    document.resolve_comment(comment_id)
    assert document.paragraphs[target_index].text == text

    drawing = next(
        item.xml.decode() if isinstance(item.xml, bytes) else item.xml
        for item in document.story_items
        if item.xml and "r:embed" in (item.xml.decode() if isinstance(item.xml, bytes) else item.xml)
    )
    image_id = re.search(r'r:embed="([^"]+)"', drawing).group(1)
    document.replace_image(image_id, _matrix_png())
    assert document.image_data(image_id) == _matrix_png()
    assert document.update_layout_backed_fields().updated_count == 0
    assert document.rebuild_toc().entry_count == 21
    edited = tmp_path / "edited.docx"
    document.save(edited)
    assert zipfile.ZipFile(edited).namelist()[0] == "[Content_Types].xml"

    redline = tmp_path / "redline.docx"
    cli = os.environ.get("RDOCX_CLI")
    command = [cli] if cli else [
        "cargo", "run", "--quiet", "--locked", "-p", "rdocx-cli", "--"
    ]
    result = subprocess.run(
        command + ["compare", str(source), str(edited), "--author", "Reviewer",
                   "--timestamp", _MATRIX_TIMESTAMP, "-o", str(redline)],
        cwd=Path(__file__).resolve().parents[3], capture_output=True, text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert len(rdocx.Document(redline).revisions) > 0
    assert rdocx.Document(edited).to_pdf().startswith(b"%PDF")


def test_issue_161_rebuilt_toc_compares_and_resolves_both_sides(tmp_path):
    _assert_oracle_version()
    import rdocx
    from docx import Document as OracleDocument
    from docx.oxml import OxmlElement

    source = _matrix_fixture(tmp_path / "source.docx")
    oracle = OracleDocument(source)
    body = oracle.element.body
    control = OxmlElement("w:sdt")
    control.append(OxmlElement("w:sdtPr"))
    content = OxmlElement("w:sdtContent")
    for paragraph in list(body)[:6]:
        body.remove(paragraph)
        content.append(paragraph)
    control.append(content)
    body.insert(0, control)
    oracle.save(source)
    edited = rdocx.Document(source)
    assert edited.rebuild_toc().entry_count == 6
    edited_path = tmp_path / "edited.docx"
    edited.save(edited_path)
    redline = rdocx.Document(source)
    assert redline.compare(
        rdocx.Document(edited_path), "Ada", _MATRIX_TIMESTAMP, granularity="word"
    ) == ()
    accepted = rdocx.Document.from_bytes(redline.to_bytes())
    accepted.accept_all()
    assert accepted.compare(rdocx.Document(edited_path), "Ada", _MATRIX_TIMESTAMP) == ()
    rejected = rdocx.Document.from_bytes(redline.to_bytes())
    rejected.reject_all()
    assert rejected.compare(rdocx.Document(source), "Ada", _MATRIX_TIMESTAMP) == ()


def test_issue_161_toc_entry_hyperlink_transition_tracks_boundary(tmp_path):
    _assert_oracle_version()
    import rdocx
    from docx import Document as OracleDocument
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    original = OracleDocument()
    original.add_paragraph("Chapter 1")
    original.add_paragraph("Following body paragraph")
    original_path = tmp_path / "toc_entry_original.docx"
    original.save(original_path)
    edited = OracleDocument(original_path)
    paragraph = edited.paragraphs[0]._p
    for child in list(paragraph):
        paragraph.remove(child)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), "_Toc1")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "Chapter 1"
    run.append(text)
    hyperlink.append(run)
    paragraph.append(hyperlink)
    edited_path = tmp_path / "toc_entry_edited.docx"
    edited.save(edited_path)

    redline = rdocx.Document(original_path)
    assert redline.compare(rdocx.Document(edited_path), "Ada", _MATRIX_TIMESTAMP) == ()
    accepted = rdocx.Document.from_bytes(redline.to_bytes())
    accepted.accept_all()
    assert accepted.compare(rdocx.Document(edited_path), "Ada", _MATRIX_TIMESTAMP) == ()
    rejected = rdocx.Document.from_bytes(redline.to_bytes())
    rejected.reject_all()
    assert rejected.compare(rdocx.Document(original_path), "Ada", _MATRIX_TIMESTAMP) == ()
