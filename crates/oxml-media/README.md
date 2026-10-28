# oxml-media

Identify OOXML media safely, derive intrinsic image sizing, and allocate
collision-free package part names without dependencies.

## Capabilities

- Magic-byte and filename identification for common image formats.
- Native sizing for PNG, JPEG, GIF, BMP, and WebP.
- Parameter-free MIME token validation and sequential media naming.
- Limited MP3, RIFF WAVE, and ISO base media signature checks.
- Truncation-safe probes, explicit DPI fallback, and collision-free numbering
  keep media decisions deterministic before a package is mutated.

## Measured footprint and speed

The archive row is regenerated from the package that carries this README.

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: oxml-media | 12,250 compressed bytes, 51,016 member bytes, 6 members | 0.14.4 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo publish --workspace --dry-run --allow-dirty --locked` | Tracked `oxml-media` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-28 |

## Use it when

Use this crate when an OOXML writer must identify or size image bytes, validate
a package content type, or check MP3, RIFF WAVE, and ISO base media signatures
without decoding the complete payload.

## Relationship

DOCX and PPTX package facades use these helpers before adding media parts and
relationships. This crate identifies containers and metadata. It does not
decode media or identify codecs.

## Example

```rust,no_run
use oxml_media::{ImageFormat, resolve};

let format = resolve(b"\x89PNG\r\n\x1a\n", "image.bin");
assert_eq!(format, ImageFormat::Png);
```

Add `oxml-media = "0.14.4"` to your dependencies. See the [API documentation](https://docs.rs/oxml-media) for supported formats and sizing functions.
