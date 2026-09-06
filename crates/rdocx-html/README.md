# rdocx-html

`rdocx-html` converts parsed Word content to a complete HTML document, an HTML
fragment, or Markdown. It works from semantic WordprocessingML and does not
run the page layout engine.

## Use it when

Use this crate when an application already owns semantic Word content. Use the
high-level [`rdocx`](https://docs.rs/rdocx) facade when starting from a DOCX
file.

## Relationship

`rdocx` prepares the `HtmlInput` consumed here. This conversion path is
independent of pagination and PDF rendering.

## Example

```rust,no_run
use rdocx_html::{HtmlInput, HtmlOptions, to_html_document, to_markdown};

fn export(input: &HtmlInput) -> (String, String) {
    (
        to_html_document(input, &HtmlOptions::default()),
        to_markdown(input),
    )
}
```

```toml
[dependencies]
rdocx-html = "0.13.0"
```
