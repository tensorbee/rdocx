# rdocx-opc

`rdocx-opc` is a deprecated compatibility shim for the shared
[`oxml-opc`](https://docs.rs/oxml-opc) package layer. Existing code can keep
its old imports while migrating.

## Use it when

Use this crate only while migrating an existing `rdocx-opc` dependency. New
code should depend on `oxml-opc` directly.

## Relationship

The retained types are exact re-exports. Word-specific package construction
belongs in the high-level [`rdocx`](https://docs.rs/rdocx) facade.

## Example

```rust,no_run
use rdocx_opc::OpcPackage;

let package = OpcPackage::new();
assert!(package.parts.is_empty());
```

```toml
[dependencies]
rdocx-opc = "0.13.0"
```

For new code, replace both the dependency and the import with `oxml-opc` and
`oxml_opc`.
