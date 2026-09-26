# rdocx

[![CI](https://github.com/tensorbee/rdocx/actions/workflows/ci.yml/badge.svg)](https://github.com/tensorbee/rdocx/actions/workflows/ci.yml)
[![crates.io](https://img.shields.io/crates/v/rdocx.svg)](https://crates.io/crates/rdocx)
[![docs.rs](https://docs.rs/rdocx/badge.svg)](https://docs.rs/rdocx)
[![License: MIT/Apache-2.0](https://img.shields.io/crates/l/rdocx.svg)](https://github.com/tensorbee/rdocx/blob/main/LICENSE)
[![MSRV: 1.93](https://img.shields.io/badge/MSRV-1.93-blue.svg)](https://blog.rust-lang.org/2026/01/09/Rust-1.93.0.html)

rdocx gives Rust applications one `Document` for the complete Word workflow.
Create a business document, open and edit a customer template, retain producer
content you do not model, lay out the result, and deliver DOCX, PDF, page
images, HTML, or Markdown without starting Microsoft Word or LibreOffice.

This is an integrated native document stack, not only a DOCX writer. The same
engine powers Rust, the command line, Python, and a locally built browser
package. Project-owned layout and bundled fonts make deterministic output
possible without an Office installation or conversion service.

## Built for complete document workflows

| Workflow | Implemented result |
|---|---|
| DOCX | Create, open, edit, validate, and save complete packages, with encryption and signing through opt-in features |
| Rich authoring | Complete paragraph and run properties, multilingual and vertical typography, conditional and floating tables, section page semantics, settings, styles, numbering, fields, forms, equations, drawings, comments, and metadata |
| Preservation | Retain unknown safe producer XML byte for byte when it is not modelled |
| Native layout | Resolve Word flow content into positioned pages with bundled, system, embedded, or caller-provided fonts |
| Fixed output | PDF, PDF/A, PNG, JPEG, TIFF, and SVG |
| Flow output | HTML, HTML fragments, Markdown, MHTML, RTF, ODT, and EPUB |
| Automation | Rust facade, CLI, Python binding, and locally built browser binding |

Most applications need only `rdocx`. Specialist crates expose the OPC,
WordprocessingML, layout, HTML, and fixed-output layers for applications that
already own one of those boundaries.

## Measured footprint and speed

Archive size is regenerated from tracked package contents. The performance
rows are the enforced release-mode bounds plus one dated observation.

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rdocx | 1,084,462 compressed bytes, 6,457,069 member bytes, 36 members | 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `rdocx` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-09-19 |
| Large-document layout throughput | minimum 250 pages/s, observed 31,019.1 pages/s | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 one-page paragraphs with deterministic fonts | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | pages per wall-clock second | 2026-09-19 |
| Large-document layout peak allocation | maximum 64 MiB, observed 29.03 MiB | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 one-page paragraphs with deterministic fonts | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | peak live allocation | 2026-09-19 |
| Large-document PDF throughput | minimum 1,000 pages/s, observed 60,058.0 pages/s | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 deterministic layout pages | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | pages per wall-clock second | 2026-09-19 |
| Large-document PDF peak allocation | maximum 16 MiB, observed 1.73 MiB | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 deterministic layout pages | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | peak live allocation | 2026-09-19 |

## Examples

### Create a document

```rust,no_run
use rdocx::{Document, Length};

let mut document = Document::new();
document.add_paragraph("Quarterly report");

let mut summary = document.add_paragraph("");
summary.add_run("Status: ").bold(true);
summary.add_run("approved");

let mut table = document.add_table(1, 2);
assert!(table.set_column_width(0, Length::inches(2.0)));
assert!(table.set_column_width(1, Length::inches(4.0)));

document.save("report.docx")?;
# Ok::<(), rdocx::Error>(())
```

### Read and update a document

```rust,no_run
use rdocx::Document;
use std::collections::HashMap;

let mut document = Document::open("template.docx")?;
for paragraph in document.paragraphs() {
    println!("{}", paragraph.text());
}

let mut replacements = HashMap::new();
replacements.insert("{{status}}", "Approved");
document.replace_all(&replacements);
document.save("approved.docx")?;
# Ok::<(), rdocx::Error>(())
```

### Render and export

```rust,no_run
use rdocx::Document;

let document = Document::open("report.docx")?;
document.save_pdf("report.pdf")?;
let first_page_png = document.render_page_to_png_deterministic(0, 150.0)?;
let html = document.to_html();
let markdown = document.to_markdown();

assert!(first_page_png.as_ref().is_some_and(|png| !png.is_empty()));
assert!(!html.is_empty());
assert!(!markdown.is_empty());
# Ok::<(), rdocx::Error>(())
```

## Installation

Most Rust applications need only the facade:

```toml
[dependencies]
rdocx = "0.14.0"
```

Bundled metric-compatible fonts are always available through deterministic
rendering. The default feature also discovers system fonts. Disable default
features when an application must use only the bundled set:

```toml
[dependencies]
rdocx = { version = "0.14.0", default-features = false }
```

Native encryption and signing APIs are opt-in:

```toml
[dependencies]
rdocx = { version = "0.14.0", features = ["agile-encryption", "digital-signatures"] }
```

The workspace requires Rust 1.93 or newer and uses edition 2024.

## CLI

Install the CLI version from the same stable family:

```sh
cargo install rdocx-cli --version '^0.14.0'
```

Common commands:

```sh
rdocx inspect report.docx
rdocx text report.docx
rdocx convert report.docx --to pdf -o report.pdf
rdocx convert report.docx --to html -o report.html
rdocx convert report.docx --to md -o report.md
rdocx replace report.docx --placeholder "Draft" --value "Final" -o final.docx
rdocx diff before.docx after.docx
```

## Surfaces

| Surface | Boundary |
|---|---|
| [rdocx](https://docs.rs/rdocx) | Native Rust facade for complete DOCX packages, authoring, conversion, and rendering |
| [rdocx-oxml](https://docs.rs/rdocx-oxml) | Typed WordprocessingML for callers that already own the lower-level model |
| [rdocx-layout](https://docs.rs/rdocx-layout) | Word flow layout for callers that already own layout input |
| [rdocx-html](https://docs.rs/rdocx-html) | HTML and Markdown export from parsed Word content |
| [rdocx-cli](https://github.com/tensorbee/rdocx/blob/main/crates/rdocx-cli/README.md) | Shell automation for inspection, conversion, validation, rendering, replacement, and diffing |
| [rdocx-py](https://github.com/tensorbee/rdocx/blob/main/crates/rdocx-py/README.md) | Python editing, DOCX save, PDF output, and page images |
| [rdocx-wasm](https://github.com/tensorbee/rdocx/blob/main/crates/rdocx-wasm/README.md) | Browser and JavaScript round trips with PDF, HTML, and Markdown export |

The [binding specification](https://github.com/tensorbee/rdocx/blob/main/docs/hld/10-bindings-spec.md#native-word-facade-stability)
defines where Python, WebAssembly, and CLI intentionally expose less than
native Rust.

## Evidence-based alternatives

Comparison checked 2026-09-09. `ND` means the capability was not documented in
the official sources linked below. It does not mean the capability is
impossible. Native layout means a page-layout engine implemented by the
library, rather than conversion through Microsoft Office, a hosted service, or
another renderer.

| Project | DOCX open, create, edit | Preservation | Native layout and render | PDF and raster | HTML and Markdown | CLI | Python | Browser and WASM |
|---|---|---|---|---|---|---|---|---|
| [rdocx](https://github.com/tensorbee/rdocx) | Open, create, edit, and save | Unknown safe producer XML is retained byte for byte when it is not modelled | Yes, Word flow layout with deterministic bundled fonts | PDF and page images | HTML and Markdown export | `rdocx-cli` | `rdocx-py`, with a narrower facade | Workspace `rdocx-wasm` facade, deliberately unpublished |
| [python-docx](https://python-docx.readthedocs.io/en/stable/user/documents.html) | Create, open, change, and save | Existing content that its API cannot manipulate is left alone on load and save. No byte-exact guarantee is stated | ND | ND | ND | ND | Primary API | ND |
| [docx-rs](https://github.com/bokuweb/docx-rs) | Create and parse into the model used by its writer. Editing a parsed document is not separately documented | ND. The project states that its OOXML support is not exhaustive | ND | ND | ND | ND | ND | Document generation and DOCX-to-JSON parsing through WebAssembly |
| [docx4j](https://github.com/plutext/docx4j) | Open, create, edit, and save | ND for unknown XML or byte-exact round trips | No project-owned page engine is documented. PDF paths use XSL-FO with Apache FOP, Microsoft Word through documents4j, or Microsoft Graph | PDF through the documented conversion paths. Raster output is ND | HTML export and first-party Markdown import and export | ND | ND | ND |
| [Aspose.Words](https://docs.aspose.com/words/python-net/product-overview/) | Create, load, modify, and save DOCX | Feature-level preservation is documented during conversion. No byte-exact unknown-XML contract is stated | Yes, its own page-layout engine | PDF plus PNG, JPEG, BMP, and TIFF | HTML and Markdown | ND | Official Python via .NET API | ND |

Among these reviewed projects, rdocx alone documents the complete combination
of a native Rust API, project-owned Word layout, fixed and flow outputs, a CLI,
Python, and a browser surface. That statement is bounded to the official
evidence set and review date. See the official [python-docx document guide](https://python-docx.readthedocs.io/en/stable/user/documents.html),
[docx-rs project](https://github.com/bokuweb/docx-rs), [docx4j PDF paths](https://www.docx4java.org/blog/2020/09/office-pptxxlsxdocx-to-pdf-to-in-docx4j-8-2-3/),
[docx4j Markdown module](https://github.com/plutext/docx4j/tree/VERSION_17_1_1/docx4j-markdown),
[Aspose fixed-page documentation](https://docs.aspose.com/words/python-net/converting-to-fixed-page-format/),
[Aspose format table](https://docs.aspose.com/words/python-net/supported-document-formats/),
and [Aspose load behavior](https://docs.aspose.com/words/python-net/supported-features-on-document-load/).

## Honest boundaries

rdocx preserves safe unmodeled XML, but it does not execute VBA, ActiveX, OLE,
add-ins, or embedded applications. Binary DOC and Word 2003 XML are permanent
non-goals. Exact authoring, reading, rendering, and preservation coverage lives
in the [modern DOCX capability matrix](https://github.com/tensorbee/rdocx/blob/main/docs/hld/02-scope-and-non-goals.md#modern-docx-capability-matrix).

## License

Licensed under either of:

- MIT license ([LICENSE](https://github.com/tensorbee/rdocx/blob/main/LICENSE) or <https://opensource.org/licenses/MIT>)
- Apache License, Version 2.0 (<https://www.apache.org/licenses/LICENSE-2.0>)

at your option.
