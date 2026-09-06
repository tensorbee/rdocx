# rdocx-layout

`rdocx-layout` converts an assembled Word `LayoutInput` into positioned page
frames. It owns DOCX style resolution, line breaking, table layout, and
pagination.

## Use it when

Use `layout_document_deterministic` for reproducible output with bundled fonts.
Applications that start from a DOCX file should normally call the rendering
methods on [`rdocx::Document`](https://docs.rs/rdocx) instead.

## Relationship

This crate converts Word-specific semantic input into the shared positioned
model from `oxml-layout`. PDF and raster backends consume that model.

## Example

```rust,no_run
use rdocx_layout::{LayoutInput, Result, layout_document_deterministic};

fn page_count(input: &LayoutInput) -> Result<usize> {
    Ok(layout_document_deterministic(input)?.pages.len())
}
```

```toml
[dependencies]
rdocx-layout = "0.13.1"
```
