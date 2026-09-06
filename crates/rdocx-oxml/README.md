# rdocx-oxml

`rdocx-oxml` provides typed WordprocessingML elements with prefix-tolerant
parsing, schema-ordered serialization, and verbatim preservation of unmodelled
XML.

## Use it when

Use this crate when a tool already owns the XML model. Most applications should
use the package-preserving [`rdocx`](https://docs.rs/rdocx) facade instead.

## Relationship

`rdocx` owns complete DOCX packages and uses these types for WordprocessingML
parts. The shared `oxml-*` crates remain format-neutral.

## Example

```rust,no_run
use rdocx_oxml::{BodyContent, CT_Document, CT_P};

let mut document = CT_Document::new();
let mut paragraph = CT_P::new();
paragraph.add_run("Low-level WordprocessingML");
document.body.content.push(BodyContent::Paragraph(paragraph));

assert_eq!(document.body.paragraphs().count(), 1);
```

```toml
[dependencies]
rdocx-oxml = "0.13.1"
```

## Migrating from 0.4

Version 0.5 adds explicit preservation state to the public numbering structs.
Code that constructs `CT_Lvl`, `CT_AbstractNum`, `CT_Num`, or `CT_Numbering`
with a full struct literal must supply the new raw XML fields. Prefer
`CT_Lvl::new`, `CT_AbstractNum::new`, and `CT_Numbering::new`. When a literal is
required, `CT_Lvl` adds empty `extra_xml` and `extra_attributes` vectors plus
`ppr_raw: None` and `rpr_raw: None`. `CT_AbstractNum` and `CT_Num` add empty
`extra_xml` and `extra_attributes` vectors. `CT_Numbering` adds empty
`root_attributes` and `extra_xml` vectors. These fields retain producer XML
during typed numbering edits and are intentionally part of the 0.5 model.

`CT_TabStop` full literals also add `source_occurrence: None`. Parsed numbering
tabs use this field to keep producer XML attached to the same occurrence after
typed edits, insertions, or removals. `CT_TabStop::new` initializes it for a new
tab. Semantic equality ignores the field and continues to compare alignment,
position, and leader values. `CT_Tabs::from_xml_with_prefixes` accepts the
in-scope WordprocessingML prefixes when callers parse aliased tab collections
directly. Numbering, style, and document-family readers route complete
paragraph property subtrees through one internal namespace projection rather
than exposing a partially contextual public `CT_PPr` parser.
