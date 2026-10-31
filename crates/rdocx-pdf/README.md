# rdocx-pdf

`rdocx-pdf` preserves the former Word-family renderer import path while
applications migrate to [`oxml-pdf`](https://docs.rs/oxml-pdf). Its public API
is an exact re-export of the shared fixed-output backend.

## Capabilities

- Existing `rdocx_pdf` imports continue to compile.
- Exact `oxml-pdf` functions and types.
- No duplicate renderer implementation or document parsing.
- A direct migration path to `oxml-pdf` or `rdocx::Document`.
- The re-export retains PDF, PDF/A, raster, metadata, link, bookmark, and
  tagged-structure output options from the shared backend.

## Measured footprint and speed

The archive row is a historical 0.16.5 observation. Current 0.17.0 archive measurements remain pending.

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rdocx-pdf | 8,128 compressed bytes, 26,782 member bytes, 6 members | 0.16.5 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo publish --workspace --dry-run --allow-dirty --locked` | Tracked `rdocx-pdf` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-29 |

## Use it when

Use this crate only while migrating an existing dependency. New code should use
`oxml-pdf`, or call PDF and PNG rendering directly on
[`rdocx::Document`](https://docs.rs/rdocx).

## Relationship

The shim forwards the shared renderer API without owning document layout or
package behavior.

## Example

```rust,no_run
use rdocx_pdf::render_to_pdf;

let renderer = render_to_pdf;
let _ = renderer;
```

```toml
[dependencies]
rdocx-pdf = "0.17.0"
```

For new code, replace both the dependency and the import with `oxml-pdf` and
`oxml_pdf`.
