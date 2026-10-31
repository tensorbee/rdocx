# oxml-core

Share exact OOXML units, XML helpers, and package-independent document property
models.

## Capabilities

- Lengths in twips, EMUs, half-points, centipoints, and physical units.
- Angles and thousandths-of-a-percent values used by OOXML schemas.
- Core, application, and custom property parsing and writing.
- XML 1.0 lexical validation, namespace handling, decoding, and raw capture.
- Unknown property values and subtrees retain their producer bytes while
  modeled values use schema-aware units and expanded names.

## Measured footprint and speed

The archive row is regenerated from the package that carries this README.

The archive row uses `cargo package --locked --allow-dirty --no-verify` with all 22 local package patches. Compilation and publication verification are separate gates.

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: oxml-core | 21,854 compressed bytes, 104,587 member bytes, 15 members | 0.15.0 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --allow-dirty --no-verify` | Tracked `oxml-core` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-31 |

## Use it when

Use this crate when implementing OOXML infrastructure shared by WordprocessingML and PresentationML. End-user document applications should prefer `rdocx` or `rpptx`.

## Relationship

Higher-level `oxml-*`, `rdocx-*`, and `rpptx-*` crates build on these primitives
without introducing a format-specific dependency here. These are focused OOXML
building blocks, not a general DOM or schema implementation. Unit conversions
follow the repository's exact truncation rules.

## Example

```rust,no_run
use oxml_core::Length;

let page_width = Length::inches(8.5);
assert_eq!(page_width.to_twips(), 12_240);
```

Add `oxml-core = "0.15.0"` to your dependencies. See the [API documentation](https://docs.rs/oxml-core) for XML and property types.
