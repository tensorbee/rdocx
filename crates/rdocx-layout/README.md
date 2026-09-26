# rdocx-layout

`rdocx-layout` turns semantic Word content into positioned pages. It resolves
styles and numbering, shapes text, lays out tables and notes, paginates
sections, and retains source provenance.

## Capabilities

- Word style, numbering, field, and paragraph resolution.
- Line breaking, tables, footnotes, endnotes, sections, and pagination.
- Deterministic bundled fonts plus system, embedded, and caller fonts.
- Page-reference lookup, Word source provenance, and top-level body extents.
- Shared `oxml-layout` output for downstream backends.
- Section-specific page geometry, columns, headers, footers, notes, floating
  tables, vertical text, and conditional table formatting retain Word meaning.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rdocx-layout | 256,485 compressed bytes, 1,388,431 member bytes, 15 members | 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `rdocx-layout` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-09-19 |
| Large-document layout throughput | minimum 250 pages/s, observed 31,019.1 pages/s | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 one-page paragraphs with deterministic fonts | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | pages per wall-clock second | 2026-09-19 |
| Large-document layout peak allocation | maximum 64 MiB, observed 29.03 MiB | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 one-page paragraphs with deterministic fonts | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | peak live allocation | 2026-09-19 |
| Large-document PDF throughput | minimum 1,000 pages/s, observed 60,058.0 pages/s | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 deterministic layout pages | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | pages per wall-clock second | 2026-09-19 |
| Large-document PDF peak allocation | maximum 16 MiB, observed 1.73 MiB | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 deterministic layout pages | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | peak live allocation | 2026-09-19 |

## Use it when

Use `layout_document_deterministic` for reproducible output with bundled fonts.
Applications that start from a DOCX file should normally call the rendering
methods on [`rdocx::Document`](https://docs.rs/rdocx) instead.

The provenance variants return `WordLayoutResult`. Its
`body_layout_fragments` accessor reports the point-space block extent on every
occupied page for one zero-based direct body item. The result keeps an empty
fragment slice for preserved body content that does not enter layout.

## Relationship

This crate converts Word-specific semantic input into the shared positioned
model from `oxml-layout`. PDF and raster backends consume that model. It does
not emit PDF or pixels itself.

## Example

```rust,no_run
use rdocx_layout::{LayoutInput, Result, layout_document_deterministic};

fn page_count(input: &LayoutInput) -> Result<usize> {
    Ok(layout_document_deterministic(input)?.pages.len())
}
```

```toml
[dependencies]
rdocx-layout = "0.14.0"
```
