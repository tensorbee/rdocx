# oxml-pdf

Render format-neutral `LayoutResult` pages to PDF, archival PDF/A, or raster
images.

## Capabilities

- PDF output from positioned page elements.
- PDF/A-2b and PDF/A-3b preflight and rendering.
- Font embedding and subsetting, metadata, outlines, links, and structure.
- Selected-page PNG and JPEG plus multi-page TIFF output.
- Tagged reading order, PDF/A output intents, gradients, clipping, opacity,
  transformed groups, and shared raster composition use the same page model.

## Measured footprint and speed

Archive size is regenerated from tracked package contents. The performance
rows are the enforced release-mode bounds plus one dated observation.

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: oxml-pdf | 76,052 compressed bytes, 347,842 member bytes, 14 members | 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `oxml-pdf` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-09 |
| Large-document layout throughput | minimum 250 pages/s, observed 31,019.1 pages/s | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 one-page paragraphs with deterministic fonts | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | pages per wall-clock second | 2026-09-19 |
| Large-document layout peak allocation | maximum 64 MiB, observed 29.03 MiB | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 one-page paragraphs with deterministic fonts | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | peak live allocation | 2026-09-19 |
| Large-document PDF throughput | minimum 1,000 pages/s, observed 60,058.0 pages/s | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 deterministic layout pages | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | pages per wall-clock second | 2026-09-19 |
| Large-document PDF peak allocation | maximum 16 MiB, observed 1.73 MiB | rdocx 0.14.0 | macOS 26.6.2, Apple M5 Max, arm64 | release, one test thread | 1,000 deterministic layout pages | `cargo test -p rdocx --test regression_test --release a_thousand_page_document_paginates_and_renders_within_the_declared_limits -- --ignored --exact --nocapture --test-threads=1` | peak live allocation | 2026-09-19 |

## Use it when

Use this crate when a custom OOXML frontend already produces shared layout frames. Use `rdocx::Document::to_pdf` or `rpptx::Presentation::to_pdf_deterministic` for normal document conversion.

## Relationship

This is the shared successor to the deprecated `rdocx-pdf` shim. It is an
output backend and does not open OOXML packages or lay out DOCX and PPTX
content. Raster page indices are zero-based.

## Example

```rust,no_run
use oxml_layout::LayoutResult;
use oxml_pdf::render_to_pdf;

let layout = LayoutResult::new(Vec::new(), Vec::new(), None, Vec::new());
let pdf = render_to_pdf(&layout);
assert!(pdf.starts_with(b"%PDF-"));
```

Add `oxml-pdf = "0.14.0"` and `oxml-layout = "0.14.0"` to your dependencies. See the [renderer API](https://docs.rs/oxml-pdf) for the accepted layout model.
