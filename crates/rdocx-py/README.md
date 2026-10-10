# rdocx-py

Native rdocx for Python

`rdocx` gives Python applications a native DOCX workflow for creating,
opening, editing, comparing, laying out, and rendering Word documents. It
works directly with OOXML packages and does not require Microsoft Word,
LibreOffice, a conversion service, or a Java or .NET runtime.

The package runs the Rust `rdocx` document engine locally, presents a typed
Python API, and needs no remote service for editing, layout, or rendering.
Review workflows cover comments, tracked revisions, comparison, and table of
contents rebuilding. Complete package saves preserve safe producer XML and
parts that the focused Python surface does not model.

## Installation

Install the current release from PyPI:

```sh
python -m pip install rdocx
```

`rdocx` requires CPython 3.9 or newer. Its platform wheels use the Python 3.9
stable ABI, so one wheel supports multiple compatible Python versions. Source
installation requires a Rust toolchain and maturin.

## Quick start

```python
from rdocx import Document

doc = Document()
doc.add_paragraph("Hello from Python")
doc.save("hello.docx")

pdf = doc.to_pdf()
with open("report.pdf", "wb") as output:
    output.write(pdf)
```

## Capabilities

- File and byte-based DOCX input and output.
- Paragraphs, runs, fonts, tables, rows, cells, sections, and styles.
- Rich per-section headers and footers, related stories, and hyperlinks
  resolved, retargeted, or removed in any story.
- Paragraph text replacement that keeps paragraph formatting, comments, and
  bookmarks.
- Counted literal replacement with `Paragraph.replace_text`, `Cell.replace_text`
  and `Document.replace_text_at`, scoped to the selected owner and preserving
  run formatting. `expect` checks the local count before publishing any edit.
- Paragraph style assignment by style ID or name, checked against the styles
  the document defines.
- New documents with Word's usual styles, such as `Heading 2`, `Title`,
  `List Paragraph`, `Caption`, and `Table Grid`.
- Style creation with a font, spacing, and indentation through
  `Document.add_style`, checked updates through `Document.set_style`, plus
  style removal and default selection.
- Numbering definitions and instances built from `ListLevel` values, and
  paragraph styles linked to a numbering level.
- Core document properties such as title, author, and revision, read and
  written through `Document.core_properties` under python-docx's names.
- Picture replacement and resizing by image relationship.
- Complete paragraph and run formatting, multilingual and vertical typography,
  conditional and floating tables, section semantics, settings, fields, forms,
  equations, drawings, comments, and metadata.
- Tracked comparison, main-body comment threads, bookmarks, revision
  resolution, and TOC insertion and rebuilding.
- Deterministic layout fragments, page geometry, and PDF, PDF/A, SVG, PNG,
  JPEG, and TIFF output through the native document engine, with caller fonts
  or a font directory for PDF. PDF and raster methods accept keyword-only
  `revision_view="tracked"` to show tracked changes.
- The python-docx calls a first draft reaches for: `add_heading`,
  `add_page_break`, `add_paragraph(text, style=...)` with python-docx's list
  styles such as `'List Bullet'` added on demand, `add_table(..., style=...)`,
  `insert_paragraph_before`, `add_section`, `run.bold`, `run.italic`,
  `run.underline`, `font.highlight_color`, `styles['Normal']`, and
  `Document(stream)` and `save(stream)`.
- Raw XML for a missing feature: `xml` and `replace_xml` on paragraphs, runs,
  tables and cells, and `section_xml` and `replace_section_xml` for section
  properties. A replacement that holds another element, would lose an element,
  misplaces an element Word checks, or references an unknown relationship id
  or one of the wrong type raises `ValueError`.
- Python collections with negative indexes, slices, and iteration. Paragraph,
  run and table handles survive appends and in-place edits, and a handle a
  structural change moved raises a stale-handle error naming that change.
- An `AttributeError` that names the rdocx call when a python-docx name is
  spelled differently, and a `ValueError` instead of a silently wrong file for
  a bare int length that rounds to zero, an empty font name, a table without
  rows or columns, or a `.pdf` save path.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Large-document layout throughput | minimum 250 pages/s, observed 31,019.1 pages/s | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 one-page paragraphs with deterministic fonts | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | pages per wall-clock second | 2026-09-19 |
| Large-document layout peak allocation | maximum 64 MiB, observed 29.03 MiB | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 one-page paragraphs with deterministic fonts | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | peak live allocation | 2026-09-19 |
| Large-document PDF throughput | minimum 1,000 pages/s, observed 60,058.0 pages/s | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 deterministic layout pages | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | pages per wall-clock second | 2026-09-19 |
| Large-document PDF peak allocation | maximum 16 MiB, observed 1.73 MiB | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 deterministic layout pages | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | peak live allocation | 2026-09-19 |

## Use it when

Use `rdocx` when a Python service, desktop application, or automation task
needs to work with Word documents locally. It is suited to template updates,
document assembly, review workflows, comparison, and deterministic output.

## Relationship

The Python API is intentionally focused. It is source compatible with the
documented `rdocx` surface, but it is not a drop-in replacement for every
python-docx method or private lxml object.

## Example

Open and update an existing document:

```python
from rdocx import Document

document = Document("template.docx")
document.add_paragraph("Approved")
document.save("approved.docx")
```

## Package XML escape hatch

For an unmodelled package part, edit the DOCX ZIP with lxml, then reopen it
with `rdocx`. Keep the original ZIP members and update package relationships
and content types if an edit adds or removes a part. For example, to change
an existing application property:

```python
from io import BytesIO
from zipfile import ZipFile
from lxml import etree
from rdocx import Document

source = Document("report.docx").to_bytes()
output = BytesIO()
with ZipFile(BytesIO(source)) as original, ZipFile(output, "w") as edited:
    for member in original.infolist():
        data = original.read(member.filename)
        if member.filename == "docProps/app.xml":
            root = etree.fromstring(data)
            ns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            root.find(f"{{{ns}}}Application").text = "Report service"
            data = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
        edited.writestr(member, data)
document = Document.from_bytes(output.getvalue())
document.save("report-edited.docx")
```

The typed API handles style formatting, comments, fields, relationships and
other modeled operations. `set_style` keeps properties not supplied in the
call and cannot clear an existing theme font or theme colour.

## Type checking

The distribution includes hand-written extension stubs and a `py.typed`
marker. Editors and type checkers can inspect concrete document, paragraph,
run, table, section, story, comment, comparison, and layout types.

```sh
python -m pip install mypy
python -m mypy your_application.py
```

The release gate validates the installed package with strict mypy and
`stubtest` in addition to its runtime suite.

`Document.set_header(text)` and `Document.set_footer(text)` replace a complete
story through one checked transaction. Removing a complete owned comment thread
cleans its definitions and companion metadata. Partial ranges or malformed
ownership raise `RdocxError` and preserve the document and live handles.
Successful replacement invalidates earlier handles once.

## Project links

- [Source repository](https://github.com/tensorbee/rdocx)
- [Issue tracker](https://github.com/tensorbee/rdocx/issues)
- [Changelog](https://github.com/tensorbee/rdocx/blob/main/CHANGELOG.md)
- [Binding specification](https://github.com/tensorbee/rdocx/blob/main/docs/hld/10-bindings-spec.md)

## License

Licensed under either the
[MIT License](https://github.com/tensorbee/rdocx/blob/main/LICENSE) or the
[Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0), at your
option.
