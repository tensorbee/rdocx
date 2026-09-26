# rpptx-oxml

`rpptx-oxml` provides typed PresentationML parts for schema-level slide work.
It parses and serializes presentations, slides, layouts, masters, shape trees,
notes, comments, diagrams, and timing data.

## Capabilities

- Typed presentation, slide, layout, master, and notes parts.
- Shapes, groups, pictures, connectors, and graphic frames.
- Comments, replies, placeholders, SmartArt, and animation timing structures.
- Part-level round trips with schema-aware ordering and identifier helpers.
- Typed text, tables, charts, media relationships, notes masters, handout
  masters, custom shows, transitions, extensions, and raw preservation state.

## Measured footprint and speed

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: rpptx-oxml | 153,106 compressed bytes, 1,042,305 member bytes, 20 members | 0.12.1 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `rpptx-oxml` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-09-19 |

## Use it when

Use this crate for schema-level slide, shape, text, relationship, and
presentation-part work. Use the incubating `rpptx` facade for complete deck
operations.

## Relationship

`rpptx` owns complete package behavior, while layout and rendering crates
consume this typed model. `rpptx-oxml` promises part-level parsing and
serialization, not whole-package preservation by itself.

## Example

```rust,no_run
use rpptx_oxml::presentation::CT_Presentation;

let presentation = CT_Presentation::from_xml(br#"
  <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:notesSz cx="6858000" cy="9144000"/>
  </p:presentation>
"#)?;
assert!(presentation.slide_ids.is_empty());
# Ok::<(), Box<dyn std::error::Error>>(())
```

Add `rpptx-oxml = "0.12.1"` to your dependencies. The [API documentation](https://docs.rs/rpptx-oxml) lists the modeled PresentationML types.
