# rpptx

`rpptx` gives Rust applications one `Presentation` for complete
PowerPoint-compatible package workflows. Open or create decks, edit content,
retain unmodelled package data, validate invariants, and produce deterministic
presentation, notes, handout, PDF, and animation outputs.

## Capabilities

- Open, create, validate, and save PPTX or PPSX packages.
- Encrypt and sign packages through the opt-in `agile-encryption` and
  `digital-signatures` features.
- Add, remove, move, duplicate, and transfer slides.
- Author and edit text, pictures, shapes, groups, tables, charts, comments,
  SmartArt text, and media.
- Produce PDF, PDF/A, resolved slide frames, notes, handouts, and animations.
- Import HTML, ODP, and PDF through explicit facade APIs.
- Resolve master, layout, placeholder, theme, chart, SmartArt, media, notes,
  comments, and timing state while preserving unmodelled package content.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rpptx | 390,209 compressed bytes, 2,030,685 member bytes, 16 members | 0.12.1 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `rpptx` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-09-19 |

## Use it when

Use this crate for complete PPTX applications. Choose the lower-level `rpptx-oxml` only for schema-level PresentationML work.

## Relationship

The facade owns package preservation and delegates part modeling, inheritance
resolution, charts, and layout lowering to the specialist `rpptx-*` and
`oxml-*` crates.

## Example

```rust,no_run
use rpptx::Presentation;

let deck = Presentation::new()?;
let bytes = deck.to_bytes()?;
# Ok::<(), Box<dyn std::error::Error>>(())
```

```toml
[dependencies]
rpptx = "0.12.1"
```

Enable native encryption and signing explicitly:

```toml
[dependencies]
rpptx = { version = "0.12.1", features = ["agile-encryption", "digital-signatures"] }
```
