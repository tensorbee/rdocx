# 03, Architecture

## Three families, one workspace

```
crates/
  # format-neutral infrastructure
  oxml-core          units, xml helpers, entity decoding, raw-XML capture,
                     core / app / custom properties
  oxml-opc           ZIP and OPC package, relationships, content types
  oxml-media         image sniffing, dimensions and DPI, MIME, media naming
  oxml-drawing       DrawingML: colour, transforms, geometry, fills, lines,
                     effects, theme, text body
  oxml-layout        output types, font manager, bundled fonts, line breaking
  oxml-pdf           PDF writer and tiny-skia rasteriser
  oxml-sml           minimal SpreadsheetML writer, chart workbooks only
  oxml-cli-support   range parsing, JSON envelope, output-path defaulting
  oxml-py-support    content paths, stale-element errors, the Length pyclass

  # WordprocessingML
  rdocx-opc          deprecated shim over oxml-opc
  rdocx-oxml         WordprocessingML types, re-exports oxml-core
  rdocx-layout       flow engine, paginator, blocks, tables, style resolver
  rdocx-pdf          deprecated shim over oxml-pdf
  rdocx-html         HTML and Markdown emitter
  rdocx              the python-docx-shaped facade
  rdocx-cli  rdocx-wasm  rdocx-py

  # PresentationML
  rpptx-oxml         PresentationML types
  rpptx-layout       inheritance resolver and flattener
  rpptx-render       resolved slides to page frames
  rpptx-chart        ChartML model and renderer
  rpptx              the python-pptx-shaped facade, plus assets/default.pptx
  rpptx-cli  rpptx-wasm  rpptx-py
```

## The dependency rule

The graph is acyclic and layered. **Nothing in `oxml-*` may depend on
`rdocx-*` or `rpptx-*`,** with exactly one documented exception below.

```
oxml-core ──┬─→ oxml-drawing ──→ rpptx-oxml ──→ rpptx-layout ──→ rpptx-render
            │         │                                              │
            │         └────────────────→ rdocx-oxml ──→ rdocx-layout │
            ├─→ oxml-opc                                    │        │
            ├─→ oxml-media                                  ↓        ↓
            └─→ oxml-layout ──→ oxml-pdf ←──────────── rdocx-pdf   rpptx
                                                            ↓        ↓
                                                          rdocx   rpptx-cli
```

**The one exception.** `rdocx_oxml::theme::Theme` becomes a thin adapter over
`oxml_drawing::CT_OfficeStyleSheet` (`impl From<&CT_OfficeStyleSheet> for
Theme`), so that `rdocx-layout`'s existing `LayoutInput.theme` field does not
churn. The edge runs `oxml-drawing → rdocx-oxml`, never the reverse.

## Why these seams

**`oxml-opc` does not depend on `oxml-core`.** It has its own small local-name
handling. Staying independent means it is publishable first and consumable
alone, which matters for `rdocx-wasm`, which wants only `OpcPackage`.

**`oxml-media` has no dependencies at all.** It is pure byte sniffing, so it is
a leaf that anything can take cheaply.

**`oxml-layout` is where the format boundary genuinely falls.** Its
`output.rs` is already 100 percent docx-free: page frames, positioned elements,
glyph runs, colours and fonts. Its sibling `input.rs` is 100 percent
docx-specific and stays in `rdocx-layout`. That seam is the reason the PDF
backend transfers for free.

**`oxml-pdf` consumes only `LayoutResult`.** It depends on `oxml-layout` and
nothing else in the workspace. A slide is a page with a fixed size, so the same
crate serves both formats without knowing either exists.

**`rpptx-layout` is separate from `rpptx-render`.** The inheritance resolver
produces a `ResolvedSlide` in which every theme reference, colour transform and
inherited property is already collapsed to a concrete value. The renderer
consumes that and nothing else. Freezing this contract is what lets the resolver
and the renderer be built and tested independently.

## What stays put

`rdocx-oxml` remains a real crate holding roughly 8,700 lines of
WordprocessingML: text, properties, tables, styles, numbering, borders, headers
and footers, footnotes, placeholder replacement, and `drawing.rs`. The
`wp:` inline and anchor code in the latter is Word-only and has no pptx value,
so it is not migrated.

`rdocx-layout` keeps the flow model: the engine, the paginator, blocks, tables
and the style resolver. Slides do not paginate, so none of it transfers.

## Versioning

`oxml-*` and `rdocx-*` share `version.workspace = true` and move together.

`rpptx-*` crates **opt out** with an explicit `version = "0.1.0"`, and carry
their own `keywords` and `categories`, because the workspace values say
`["docx", "word"]` which would be wrong on a presentation crate. This means
rpptx can churn through breaking 0.1.x releases without dragging rdocx's version
number along, and rdocx can ship patches without implying rpptx stability.

They fold into the lockstep train once rpptx stabilises.

## Crate-level conventions

- **quick-xml pull parsing only.** No serde, no derive, no macros, no codegen.
  Every element's parser and serialiser is hand-written. This is a deliberate
  existing choice and the new crates follow it.
- **Spec names.** Types are `CT_*` and `ST_*` after the schema, under a
  crate-level `#![allow(non_camel_case_types)]`.
- **Root parts** get `from_xml(&[u8]) -> Result<Self>` and
  `to_xml(&self) -> Result<Vec<u8>>`. **Nested elements** get
  `from_xml(reader: &mut Reader<&[u8]>)` and
  `to_xml<W: Write>(&self, writer: &mut Writer<W>)`.
- **Prefix-tolerant on read, fixed prefix on write.** `matches_local_name`
  strips any prefix and compares the local part.
- **Unmodelled subtrees are preserved verbatim** via `capture_element` into
  `raw_xml` fields. This matters far more for PresentationML than for
  WordprocessingML, and it is the scope control for an otherwise unbounded
  format: parse only what you render, preserve the rest.
- **`thiserror`, no `anyhow`.** One error enum per crate plus a `Result` alias.
- Edition 2024, MSRV 1.93.

## Facade conventions

Both facades use the same borrow-handle idiom rdocx already has: a mutable
`Foo<'a>` wrapping `&'a mut CT_Foo` and a read-only `FooRef<'a>`, with
consuming builders for formatting so calls chain, `&mut self` methods for adding
content that return a nested handle, and index-based `Option`-returning
accessors that never panic.

Every consuming formatting builder on `Paragraph`, `Run`, `Table`, `Row`, and
`Cell` has a non-consuming `set_*` twin because a `mut self -> Self` builder
cannot back a Python property setter. The 61 consuming builders delegate to
their setter twins, so Rust callers retain chaining while borrowed handles and
Python properties use in-place mutation.
