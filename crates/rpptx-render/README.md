# rpptx-render

Rendering bridge from resolved presentation models to shared layout, raster, and PDF output.

## Use it when

Use this crate when integrating the presentation renderer below the `rpptx` facade. Applications normally call deterministic render methods on `rpptx::Presentation`.

## Relationship

It consumes `rpptx-layout` output and uses `oxml-layout` and `oxml-pdf` backends.

## Example

```rust,no_run
use rpptx_render::{RelScope, RelScopes};

let relationships = RelScopes::default();
let missing = relationships.get(RelScope::Slide, "rId1");
assert!(missing.is_err());
```

Add `rpptx-render = "0.10.0"` to your dependencies. See the [rendering API](https://docs.rs/rpptx-render) for input and output types.
