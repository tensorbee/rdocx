# rpptx-chart

`rpptx-chart` is a deprecated compatibility shim for the shared
[`oxml-chart`](https://docs.rs/oxml-chart) model and renderer. Existing imports
continue to work because every public item is an exact re-export.

## Use it when

Use this crate only while migrating an existing dependency. New code should
use `oxml-chart` directly.

## Relationship

The shim preserves the former PowerPoint-family package name without owning
ChartML parsing, serialization, validation, or rendering.

## Example

```rust,no_run
use rpptx_chart::AxisId;

let category_axis = AxisId::new(10_000_001)?;
let value_axis = AxisId::new(10_000_002)?;
assert_ne!(category_axis, value_axis);
# Ok::<(), rpptx_chart::ChartError>(())
```

```toml
[dependencies]
rpptx-chart = "0.10.0"
```

For new code, replace both the dependency and the import with `oxml-chart` and
`oxml_chart`.
