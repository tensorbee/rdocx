# oxml-core

Format-neutral OOXML units, XML helpers, and core document-property types.

## Use it when

Use this crate when implementing OOXML infrastructure shared by WordprocessingML and PresentationML. End-user document applications should prefer `rdocx` or `rpptx`.

## Relationship

Higher-level `oxml-*`, `rdocx-*`, and `rpptx-*` crates build on these primitives without introducing a format-specific dependency here.

## Example

```rust,no_run
use oxml_core::Length;

let page_width = Length::inches(8.5);
assert_eq!(page_width.to_twips(), 12_240);
```

Add `oxml-core = "0.11.0"` to your dependencies. See the [API documentation](https://docs.rs/oxml-core) for XML and property types.
