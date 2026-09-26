# oxml-drawing

Parse, edit, resolve, and serialize reusable DrawingML content without
depending on a document format.

## Capabilities

- Typed colors with theme and color-map resolution.
- Solid, gradient, pattern, image, and no-fill models.
- Preset and custom geometry, lines, transforms, text bodies, and tables.
- Schema-ordered writing with ordered preservation of unmodelled children.
- Shared chart, table, color, theme, and geometry models serve both document
  families without importing either facade.

## Measured footprint and speed

The archive row is regenerated from the package that carries this README.

| Measurement | Value | Version | Platform | Build mode | Input | Command | Statistic | Measured on |
|---|---|---|---|---|---|---|---|---|
| Crates.io archive: oxml-drawing | 159,717 compressed bytes, 1,121,654 member bytes, 24 members | 0.12.1 | macOS 26.6.2, Apple M5 Max, arm64 | `cargo package --locked --no-verify` | Tracked `oxml-drawing` package inventory | `python3 scripts/readme_doctests.py --record-measurements` | gzip archive bytes, tar member bytes, tar member count | 2026-09-19 |

## Use it when

Use this crate when reading or writing DrawingML shared by DOCX and PPTX packages. Use `rpptx` or `rdocx` for complete documents.

## Relationship

It consumes format-neutral OOXML primitives and supplies drawing models to
presentation and rendering crates. Format-specific anchors, wrappers,
packaging, and rendering belong elsewhere. Effects and DrawingML elements that
are not typed remain preserved rather than being interpreted.

## Example

```rust,no_run
use oxml_drawing::fill::Fill;

let fill = Fill::from_xml(
    br#"<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
)?;
assert!(matches!(fill, Fill::NoFill(_)));
# Ok::<(), oxml_drawing::fill::FillError>(())
```

Add `oxml-drawing = "0.12.1"` to your dependencies. Browse the [typed DrawingML API](https://docs.rs/oxml-drawing) before constructing schema-level values directly.
