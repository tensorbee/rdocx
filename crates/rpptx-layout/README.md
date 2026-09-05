# rpptx-layout

Presentation layout and placeholder inheritance resolution for PowerPoint-compatible slides.

## Use it when

Use this crate when resolving slide, layout, master, theme, and placeholder state into a renderable slide model. Use `rpptx` for complete deck operations.

## Relationship

It sits between `rpptx-oxml` and `rpptx-render` and does not own package I/O.

## Example

```rust,no_run
use rpptx_layout::{FlattenedSource, ScopedMediaIds};

let media = ScopedMediaIds::default();
assert_eq!(media.get(FlattenedSource::Slide, "rId1"), None);
```

Add `rpptx-layout = "0.10.0"` to your dependencies. See the [resolver API](https://docs.rs/rpptx-layout) for the resolved-slide contract.
