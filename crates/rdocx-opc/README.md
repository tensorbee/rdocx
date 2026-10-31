# rdocx-opc

`rdocx-opc` keeps existing Word package code compiling while it moves to
[`oxml-opc`](https://docs.rs/oxml-opc). Every public item is an exact re-export,
so migration changes the dependency and import path without changing package
behavior.

## Capabilities

- Exact `oxml-opc` re-exports under the former crate name.
- Existing `OpcPackage` imports continue to compile.
- No added Word-specific package policy or forwarding layer.
- A direct migration path for applications moving to the shared crate.
- The re-export includes relationships, content types, package properties,
  ZIP limits, signatures, encryption, and safe part-name handling.

## Measured footprint and speed

This exact published migration package supplies the archive row.

The archive row uses `cargo package --locked --allow-dirty --no-verify` with all 22 local package patches. Compilation and publication verification are separate gates.

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rdocx-opc | 3,953 compressed bytes, 10,694 member bytes, 6 members | 0.17.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --allow-dirty --no-verify` | Tracked `rdocx-opc` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-31 |

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
rdocx-opc = "0.17.0"
```

For new code, replace both the dependency and the import with `oxml-opc` and
`oxml_opc`.
