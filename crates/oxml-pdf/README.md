# oxml-pdf

Format-neutral PDF generation from the positioned output model in `oxml-layout`.

## Use it when

Use this crate when a custom OOXML frontend already produces shared layout frames. Use `rdocx::Document::to_pdf` or `rpptx::Presentation::to_pdf_deterministic` for normal document conversion.

## Relationship

This is the shared successor to the deprecated `rdocx-pdf` shim.

## Example

```rust,no_run
use oxml_layout::LayoutResult;
use oxml_pdf::render_to_pdf;

let layout = LayoutResult::new(Vec::new(), Vec::new(), None, Vec::new());
let pdf = render_to_pdf(&layout);
assert!(pdf.starts_with(b"%PDF-"));
```

Add `oxml-pdf = "0.11.0"` and `oxml-layout = "0.11.0"` to your dependencies. See the [renderer API](https://docs.rs/oxml-pdf) for the accepted layout model.
