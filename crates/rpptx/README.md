# rpptx

`rpptx` gives Rust applications one `Presentation` for complete
PowerPoint-compatible package workflows. Open or create decks, edit content,
retain unmodelled package data, validate invariants, and produce deterministic
presentation, notes, handout, PDF, and animation outputs.

## Capabilities

- Open, create, validate, and save PPTX, PPTM, POTX, POTM, PPSX, and PPSM
  packages. The save path extension selects the declared package class, so a
  template saved as `.pptx` becomes a presentation.
- Encrypt and sign packages through the opt-in `agile-encryption` and
  `digital-signatures` features.
- Add, remove, move, duplicate, and transfer slides.
- Author and edit text, pictures, shapes, groups, tables, charts, comments,
  SmartArt text, and media.
- Populate groups, nested to any depth, with text boxes, preset shapes,
  connectors, groups, tables, and pictures.
- Insert and remove table rows and columns, extending or shrinking merged
  cells and growing or shrinking the frame by the row or column size.
- Produce PDF, PDF/A, resolved slide frames, notes, handouts, and animations.
- Import HTML, ODP, and PDF through explicit facade APIs.
- Resolve master, layout, placeholder, theme, chart, SmartArt, media, notes,
  comments, and timing state while preserving unmodelled package content.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rpptx | 750,591 compressed bytes, 4,003,521 member bytes, 16 members | 0.14.3 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo publish --workspace --dry-run --allow-dirty --locked` | Tracked `rpptx` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-10-23 |

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
rpptx = "0.14.3"
```

Enable native encryption and signing explicitly:

```toml
[dependencies]
rpptx = { version = "0.14.3", features = ["agile-encryption", "digital-signatures"] }
```

## Licensed font embedding

```rust,no_run
use rpptx::{Presentation, EmbeddedFont, EmbeddedFontKind, FontEmbeddingLicense};

let mut deck = Presentation::new()?;
let bytes = std::fs::read("MyFace-Regular.ttf")?;
deck.embed_font("My Face", EmbeddedFont::new(
    EmbeddedFontKind::Regular, bytes,
    FontEmbeddingLicense::new(true, "SIL Open Font License 1.1"),
))?;
let faces = deck.embedded_fonts()?;
assert_eq!(faces[0].typeface, "My Face");
deck.save("embedded.pptx")?;
# Ok::<(), Box<dyn std::error::Error>>(())
```

Confirm the caller font's embedding license before authorizing it. Native
embedding writes checked uncompressed EOT and stages the face relationship
with the presentation model and flags. Imported compressed/unknown data remains
preserved with absent decoded bytes. Caller render faces retain precedence.
