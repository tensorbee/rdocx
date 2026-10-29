# rpptx-chart

`rpptx-chart` keeps existing presentation chart imports compiling while callers
migrate to [`oxml-chart`](https://docs.rs/oxml-chart). Every public item is an
exact re-export of the shared ChartML model and renderer.

## Capabilities

- Existing `rpptx_chart` imports continue to compile.
- Exact `oxml-chart` models, validators, and geometry.
- No duplicate chart implementation or presentation package policy.
- A direct dependency and import migration path.
- The shared surface includes typed series, axes, labels, legends, validation,
  workbook bindings, geometry, and deterministic page-model lowering.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rpptx-chart | 6,841 compressed bytes, 21,809 member bytes, 6 members | 0.14.5 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo publish --workspace --dry-run --allow-dirty --locked` | Tracked `rpptx-chart` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-29 |

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
rpptx-chart = "0.14.5"
```

For new code, replace both the dependency and the import with `oxml-chart` and
`oxml_chart`.
