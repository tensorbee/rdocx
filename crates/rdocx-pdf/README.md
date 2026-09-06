# rdocx-pdf

`rdocx-pdf` is a deprecated compatibility shim for the shared
[`oxml-pdf`](https://docs.rs/oxml-pdf) renderer. Existing imports continue to
work because the public functions are exact re-exports.

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
rdocx-pdf = "0.13.1"
```

For new code, replace both the dependency and the import with `oxml-pdf` and
`oxml_pdf`.
