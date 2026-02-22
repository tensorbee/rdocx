# rdocx

[![CI](https://github.com/tensorbee/rdocx/actions/workflows/ci.yml/badge.svg)](https://github.com/tensorbee/rdocx/actions/workflows/ci.yml)
[![crates.io](https://img.shields.io/crates/v/rdocx.svg)](https://crates.io/crates/rdocx)
[![docs.rs](https://docs.rs/rdocx/badge.svg)](https://docs.rs/rdocx)
[![License: MIT/Apache-2.0](https://img.shields.io/crates/l/rdocx.svg)](LICENSE)
[![MSRV: 1.93](https://img.shields.io/badge/MSRV-1.93-blue.svg)](https://blog.rust-lang.org/2026/01/09/Rust-1.93.0.html)

A pure Rust DOCX library — create, read, and modify Word documents programmatically. Additionally, render pixel-identical PDFs and export to HTML and Markdown, all from the same document object. No LibreOffice, no unoconv, no C dependencies.

## Why rdocx?

Most DOCX solutions in the ecosystem shell out to LibreOffice or wrap C/C++ libraries. rdocx is written entirely in Rust, so it compiles to a single binary with zero runtime dependencies. It works everywhere Rust does — including WASM.

The core focus is **DOCX**: a high-level, python-docx-inspired API for building and editing Word documents with paragraphs, tables, images, headers/footers, styles, and lists. On top of that, rdocx includes a built-in layout engine that paginates your document and can render it to **PDF** (with font subsetting, bookmarks, and selectable text) or export to **HTML** and **Markdown** — so you get faithful output in every format without leaving Rust.

## DOCX Features

- **Read & write** DOCX files with a high-level API
- **Tables** with merged cells, borders, shading, and content-based column sizing
- **Images** — inline and anchored, with header/footer background images
- **Headers & footers** with first-page support and per-section overrides
- **Styles** — paragraph and character styles, theme color resolution
- **Lists** with automatic numbering ID management
- **Template engine** with placeholder replacement (plain text and regex)
- **TOC generation** with internal hyperlinks and dot-leader tabs
- **Document merging** with style deduplication and numbering remapping

## Output Formats

- **PDF** — built-in layout engine with text shaping (rustybuzz), Unicode line breaking, multi-section pagination, font subsetting, ToUnicode CMap, bookmarks, and images
- **HTML** — semantic mapping from OOXML with CSS styling and base64-embedded images
- **Markdown** — GFM-compatible output with pipe tables and formatting
- **PNG** — page-to-image rendering via tiny-skia rasterizer

## Extras

- **WASM support** via standalone `rdocx-wasm` crate
- **CLI tool** (`rdocx-cli`) — inspect, convert, diff, replace, validate, render

## Installation

```toml
[dependencies]
rdocx = "0.1"
```

To include bundled metric-compatible fonts (Carlito, Caladea, Liberation family):

```toml
[dependencies]
rdocx-layout = { version = "0.1", features = ["bundled-fonts"] }
```

## Quick Start

### Create a document

```rust
use rdocx::{Document, Length};

let mut doc = Document::new();

doc.add_paragraph("Hello, World!");

let mut para = doc.add_paragraph("");
para.add_run("Bold text").bold(true);
para.add_run(" and ");
para.add_run("italic text").italic(true);

doc.add_table(3, 4);

doc.save("output.docx").unwrap();
```

### Read a document

```rust
use rdocx::Document;

let doc = Document::open("report.docx").unwrap();

for para in doc.paragraphs() {
    println!("{}", para.text());
}

for table in doc.tables() {
    for row in table.rows() {
        for cell in row.cells() {
            print!("{}\t", cell.text());
        }
        println!();
    }
}
```

### Convert to PDF

```rust
use rdocx::Document;

let doc = Document::open("report.docx").unwrap();
doc.save_pdf("report.pdf").unwrap();

// Or get bytes directly
let pdf_bytes = doc.to_pdf().unwrap();
```

### Convert to HTML / Markdown

```rust
use rdocx::Document;

let doc = Document::open("report.docx").unwrap();

let html = doc.to_html();
let markdown = doc.to_markdown();
```

### Template replacement

```rust
use rdocx::Document;
use std::collections::HashMap;

let mut doc = Document::open("template.docx").unwrap();

let mut replacements = HashMap::new();
replacements.insert("{{name}}", "Jane Doe");
replacements.insert("{{date}}", "2025-01-15");
doc.replace_all(&replacements);

doc.save("filled.docx").unwrap();
```

### Merge documents

```rust
use rdocx::{Document, SectionBreak};

let mut doc = Document::open("part1.docx").unwrap();
let part2 = Document::open("part2.docx").unwrap();

doc.append_with_break(&part2, SectionBreak::NextPage);
doc.save("combined.docx").unwrap();
```

## CLI

Install the CLI:

```sh
cargo install rdocx-cli
```

```sh
# Inspect document structure
rdocx inspect report.docx

# Extract plain text
rdocx text report.docx

# Convert to PDF
rdocx convert report.docx -o report.pdf

# Convert to HTML or Markdown
rdocx convert report.docx -o report.html
rdocx convert report.docx -o report.md

# Find and replace text
rdocx replace report.docx --find "Draft" --replace "Final" -o final.docx

# Diff two documents
rdocx diff v1.docx v2.docx
```

## How rdocx Compares

### vs. Python Libraries

| | rdocx | python-docx | docx2pdf | pypandoc |
|---|---|---|---|---|
| Create DOCX | Yes | Yes | -- | -- |
| Read DOCX | Yes | Yes | -- | -- |
| DOCX to PDF | Yes (built-in) | No | Via MS Word | Via Pandoc + LaTeX |
| DOCX to HTML | Yes (built-in) | No | No | Yes (lossy) |
| DOCX to Markdown | Yes (built-in) | No | No | Yes (lossy) |
| Layout engine | Yes | None | Delegates to Word | Delegates to LaTeX |
| External runtime | **None** | None (but no PDF) | **MS Word required** | **Pandoc + LaTeX** |
| Install size | **4 MB binary** | ~5 MB | ~31 KB + Word | 300-650 MB |
| Runs in Docker / CI | Yes | Yes (no PDF) | No | Yes (huge image) |
| WASM / browser | Yes | No | No | No |

**python-docx** is the most popular DOCX library in any language (~14M PyPI downloads/month), but it has **zero conversion capabilities** — no PDF, no HTML, no Markdown. Users who need PDF must bolt on a separate tool like LibreOffice (~500 MB) or a commercial API. rdocx gives you the same read/write API *plus* built-in conversion in a single 4 MB binary.

### vs. Java Libraries

| | rdocx | Apache POI | docx4j | Aspose.Words |
|---|---|---|---|---|
| Create DOCX | Yes | Yes | Yes | Yes |
| Read DOCX | Yes | Yes | Yes | Yes |
| PDF (built-in) | Yes | No | Via FOP (limited) | Yes (high fidelity) |
| HTML (built-in) | Yes | No | Yes | Yes |
| License | MIT / Apache-2.0 | Apache-2.0 | Apache-2.0 | **$1,199+** |
| Total dependency size | **4 MB** | 18-28 MB + JRE | 50-80 MB + JRE | 14 MB + JRE |
| Typical memory (moderate doc) | **10-50 MB** | 256 MB - 1 GB | 256 MB - 2 GB | 50-300 MB |
| Cold start | **< 10 ms** | 2-5 sec | 2-5 sec | 2-5 sec |
| Runtime required | None | JVM (~200 MB) | JVM (~200 MB) | JVM (~200 MB) |

Java solutions carry the JVM's baseline overhead: 50-100 MB of RAM before a single document is loaded, and 2-5 second cold starts from class loading. Apache POI has **no built-in PDF** at all. docx4j's FOP pipeline is acknowledged by its own maintainer as limited in fidelity. Aspose has excellent PDF output but costs $1,199+ per developer. rdocx delivers comparable capabilities as a zero-dependency native binary.

### vs. Other Rust Crates

| | rdocx | docx-rs | docx-rust | ooxmlsdk |
|---|---|---|---|---|
| Create DOCX | Yes | Yes | Yes | Low-level |
| Read DOCX | Yes | Yes | Yes | Low-level |
| Round-trip preservation | Yes | Limited | Limited | N/A |
| Tables, images, headers | Yes | Yes | Basic | Raw XML |
| PDF conversion | **Yes** | No | No | No |
| HTML / Markdown export | **Yes** | No | No | No |
| Layout engine | **Yes** | No | No | No |
| Page-to-image rendering | **Yes** | No | No | No |
| Template engine | **Yes** | No | No | No |
| Document merging | **Yes** | No | No | No |
| Regex find/replace | **Yes** | No | No | No |
| CLI tool | **Yes** | No | No | No |
| WASM | Yes | Yes | No | No |

**docx-rs** (1M+ downloads, 500+ stars) is the most popular Rust DOCX crate, but it is a read/write library only — no conversion, no layout engine, no PDF. The same is true for every other Rust DOCX crate. rdocx is the only Rust crate that combines DOCX read/write with a built-in layout engine and multi-format output (PDF, HTML, Markdown, PNG).

### Resource Footprint

| Metric | rdocx (native) | Python + LibreOffice | Java (POI + FOP) |
|---|---|---|---|
| Binary / install size | **4 MB** | ~500 MB | ~250 MB (JARs + JRE) |
| Memory (moderate document) | **10-50 MB** | ~200-500 MB | ~300 MB - 1.5 GB |
| Cold start | **< 10 ms** | ~2-4 sec (LibreOffice) | ~2-5 sec (JVM) |
| Serverless / Lambda friendly | Yes | Difficult | Difficult |
| Docker image overhead | **~10 MB** (musl static) | ~500 MB+ | ~250 MB+ |
| WASM compatible | Yes | No | No |

## Crate Architecture

| Crate | Purpose |
|---|---|
| `rdocx` | High-level Document API |
| `rdocx-opc` | OPC/ZIP package I/O |
| `rdocx-oxml` | OOXML types (CT_Document, CT_PPr, CT_RPr, CT_Tbl, ...) |
| `rdocx-layout` | Layout engine (text shaping, line breaking, pagination) |
| `rdocx-pdf` | PDF rendering with font subsetting |
| `rdocx-html` | HTML and Markdown conversion |
| `rdocx-cli` | CLI binary |
| `rdocx-wasm` | WASM bindings (standalone, excluded from workspace) |

## Minimum Supported Rust Version

1.93 (edition 2024)

## License

Licensed under either of

- MIT license ([LICENSE](LICENSE) or http://opensource.org/licenses/MIT)
- Apache License, Version 2.0 (http://www.apache.org/licenses/LICENSE-2.0)

at your option.
