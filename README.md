# rdocx

Pure Rust library for reading, writing, and converting DOCX documents. No LibreOffice, no unoconv, no C dependencies — just add the crate and go.

## Why rdocx?

Most DOCX solutions in the ecosystem shell out to LibreOffice or wrap C/C++ libraries. rdocx is written entirely in Rust, so it compiles to a single binary with zero runtime dependencies. It works everywhere Rust does — including WASM.

## Features

- **Read & write** DOCX files with a high-level, python-docx-inspired API
- **PDF conversion** with font subsetting, ToUnicode CMap, bookmarks, and images
- **HTML & Markdown** export with semantic mapping and CSS styling
- **Layout engine** with text shaping (rustybuzz), Unicode line breaking, and multi-section pagination
- **Tables** with merged cells, borders, shading, and content-based column sizing
- **Images** — inline and anchored, with header/footer background images
- **Headers & footers** with first-page support and per-section overrides
- **Lists** with automatic numbering ID management
- **Template engine** with placeholder replacement (plain text and regex)
- **TOC generation** with internal hyperlinks and dot-leader tabs
- **Document merging** with style deduplication and numbering remapping
- **Page-to-image rendering** via tiny-skia rasterizer
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

1.85 (edition 2024)

## License

Licensed under either of

- MIT license ([LICENSE](LICENSE) or http://opensource.org/licenses/MIT)
- Apache License, Version 2.0 (http://www.apache.org/licenses/LICENSE-2.0)

at your option.
