# rpptx-oxml

Typed PresentationML object model with package-preserving parse and serialization support.

## Use it when

Use this crate for schema-level slide, shape, text, relationship, and presentation-part work. Use `rpptx` for a stable high-level facade.

## Relationship

`rpptx` owns complete package behavior, while layout and rendering crates consume this typed model.

## Example

```rust,no_run
use rpptx_oxml::presentation::CT_Presentation;

let presentation = CT_Presentation::from_xml(br#"
  <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:notesSz cx="6858000" cy="9144000"/>
  </p:presentation>
"#)?;
assert!(presentation.slide_ids.is_empty());
# Ok::<(), Box<dyn std::error::Error>>(())
```

Add `rpptx-oxml = "0.11.0"` to your dependencies. The [API documentation](https://docs.rs/rpptx-oxml) lists the modeled PresentationML types.
