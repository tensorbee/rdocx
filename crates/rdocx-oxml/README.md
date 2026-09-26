# rdocx-oxml

`rdocx-oxml` provides typed WordprocessingML parts for tools that need
schema-level control. It accepts namespace aliases, serializes children in
schema order, and retains unmodelled XML alongside typed edits.

## Capabilities

- Document bodies, paragraphs, runs, tables, styles, and numbering.
- Headers, footers, notes, comments, revisions, and content controls.
- Transitional OfficeMath parsing and serialization.
- Prefix-tolerant input, schema-ordered output, and raw preservation state.
- Complete paragraph and run properties, multilingual and vertical text,
  conditional and floating tables, section semantics, settings, and web settings.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rdocx-oxml | 367,296 compressed bytes, 2,378,294 member bytes, 32 members | 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `rdocx-oxml` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-09-19 |

## Use it when

Use this crate when a tool already owns the XML model. Most applications should
use the package-preserving [`rdocx`](https://docs.rs/rdocx) facade instead.

## Relationship

`rdocx` owns complete DOCX packages and uses these types for WordprocessingML
parts. The shared `oxml-*` crates remain format-neutral.

## Example

```rust,no_run
use rdocx_oxml::{BodyContent, CT_Document, CT_P};

let mut document = CT_Document::new();
let mut paragraph = CT_P::new();
paragraph.add_run("Low-level WordprocessingML");
document.body.content.push(BodyContent::Paragraph(paragraph));

assert_eq!(document.body.paragraphs().count(), 1);
```

```toml
[dependencies]
rdocx-oxml = "0.14.0"
```
