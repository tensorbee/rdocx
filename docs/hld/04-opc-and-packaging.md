# 04, OPC and packaging

Owner: `oxml-opc`, with media naming in `oxml-media`.

## The package

```rust
pub struct OpcPackage {
    pub content_types: ContentTypes,
    pub package_rels:  Relationships,
    pub part_rels:     HashMap<String, Relationships>,  // key: "/ppt/slides/slide1.xml"
    pub parts:         HashMap<String, Vec<u8>>,        // same key shape
}
```

Fully in memory. Every part is decompressed at open. Part names are normalised
to a leading slash. This design is already format-neutral and is carried over
essentially unchanged from `rdocx-opc`.

**Saves are deterministic.** Both `part_rels` and `parts` are emitted in sorted
key order, so writing the same package twice produces byte-identical output.
That property is load-bearing for the round-trip corpus and must not regress.

## Generalising the constructors

The only docx-specific code in the existing crate is two constructors. They are
replaced by:

```rust
impl OpcPackage {
    /// Empty package: minimal content types, no parts, no relationships.
    pub fn new() -> Self;

    /// Package whose officeDocument relationship points at `part_name`.
    /// Package-relative, no leading slash: "word/document.xml", "ppt/presentation.xml".
    pub fn with_main_part(part_name: &str, content_type: &str) -> Self;
}

impl ContentTypes {
    /// Only the two universal defaults, "rels" and "xml".
    pub fn minimal() -> Self;
}
```

The docx presets become a short private helper in `crates/rdocx/src/document.rs`,
and the pptx presets one in `crates/rpptx/src/package.rs`.

Rejected alternatives, recorded so they are not revisited: a `PackageKind` enum
forces the leaf crate to carry every format's content-type table and grows a
variant per format. Feature-gated `new_docx` / `new_pptx` helpers make a leaf
crate feature-conditional for two functions' worth of string constants, and
features are additive so a workspace containing both compiles both anyway.

## What transfers unmodified

**`main_document_part()`** keys off the `officeDocument` relationship type,
which PowerPoint uses for `/ppt/presentation.xml`. It reads a `.pptx` today.

**`resolve_rel_target(source_part, target)`** joins a relative target against
the source part's directory and collapses `.` and `..`. It already handles
`../slideLayouts/slideLayout1.xml` from `/ppt/slides/slide1.xml` correctly, and
it clamps a traversal that escapes the root rather than allowing zip-slip.

**`rels_path_to_part_name`** and its inverse are generic path algebra.

## Relationship types

`rel_types` stays one flat module, grouped by comment. The existing thirteen
constants are kept. Added:

```rust
// Package-level. Note the package namespace, not officeDocument.
CORE_PROPERTIES, THUMBNAIL

// Shared officeDocument
EXTENDED_PROPERTIES   // docProps/app.xml
CUSTOM_PROPERTIES     // docProps/custom.xml

// PresentationML
SLIDE, SLIDE_LAYOUT, SLIDE_MASTER, NOTES_SLIDE, NOTES_MASTER,
PRES_PROPS, VIEW_PROPS, TABLE_STYLES, HANDOUT_MASTER
```

A `content_types` constants module is added alongside, so neither format crate
hand-types the long MIME strings.

rdocx resolves core properties through the package-level `CORE_PROPERTIES`
relationship and retains its normalized target. Metadata is written back to
that part with its content-type override. A document that creates metadata
without an existing relationship uses `/docProps/core.xml` and adds the missing
package relationship.

## Part naming

**Numeric suffixes are allocated after the greatest positive parsed suffix,
never as `count + 1`.** Deleting slide 2 and then adding a slide must not
collide with slide 3. The rdocx image counter parses the consecutive decimal
digits after `/word/media/image`, including `usize::MAX`, and ignores missing,
signed, zero, nonnumeric and unrelated suffixes. Ordinary packages allocate
`1 + max(existing suffix)`. At the finite boundary, checked increment wraps
from `usize::MAX` to 1 and skips every occupied parsed suffix until a free
positive number is found. Allocation never creates `image0` or overwrites an
existing numbered image part.

Canonical part layouts:

```
docx                      pptx
/word/document.xml        /ppt/presentation.xml
/word/styles.xml          /ppt/slides/slideN.xml
/word/media/imageN.ext    /ppt/slideLayouts/slideLayoutN.xml
                          /ppt/slideMasters/slideMasterN.xml
                          /ppt/notesSlides/notesSlideN.xml
                          /ppt/theme/themeN.xml
                          /ppt/media/imageN.ext
                          /ppt/charts/chartN.xml
                          /ppt/embeddings/WorkbookN.xlsx
```

## Media

`oxml-media` owns everything about an image byte string.

```rust
pub enum ImageFormat { Png, Jpeg, Gif, Bmp, Tiff, Webp, Svg, Emf, Wmf }

impl ImageFormat {
    pub fn sniff(data: &[u8]) -> Option<Self>;
    pub fn from_extension(ext: &str) -> Option<Self>;
    pub fn extension(self) -> &'static str;
    pub fn content_type(self) -> &'static str;
}

/// Sniff first, fall back to the extension, default to PNG.
pub fn resolve(data: &[u8], filename: &str) -> ImageFormat;

pub struct ImageInfo {
    pub format: ImageFormat,
    pub width_px: u32, pub height_px: u32,
    pub dpi_x: Option<f64>, pub dpi_y: Option<f64>,  // None means the file declares none
    pub bit_depth: u8, pub channels: u8, pub has_alpha: bool,
}
pub fn probe(data: &[u8]) -> Option<ImageInfo>;

impl ImageInfo {
    pub fn native_size(&self, default_dpi: f64) -> (Length, Length);
}

pub struct MediaNamer { /* dir, stem, next */ }
impl MediaNamer {
    pub fn scan<'a>(dir: &str, stem: &str, existing: impl Iterator<Item = &'a str>) -> Self;
    pub fn next_part_name(&mut self, ext: &str) -> String;
}
```

**Sniffing beats the extension.** Today the extension is trusted, so a `.png`
that is really a JPEG gets the wrong content type. This is the bug class the
crate exists to eliminate.

**`native_size` takes the DPI rather than baking one in**, because the right
default differs by consumer: python-docx assumes 72 when a file declares none,
Word assumes 96. rdocx passes 72 so the Python bindings match python-docx, and
that constant is documented rather than buried.

Header parsing is lifted from the PDF crate, where `jpeg_dimensions` and the
PNG IHDR reader are currently private. The JPEG walk classifies SOI, TEM and
RST0 through RST7 as standalone markers with no length field, and EOI terminates
the codestream. Length-bearing segments are validated for a present length of
at least two bytes and bounds within the input before the walk indexes or
advances. Fill bytes and truncated input return safely. Preserve these
invariants when the reader moves to `oxml-media`.

## Package integrity

Both facades expose a `validate()` that is cheap, non-panicking, and run
automatically under `debug_assertions` before `save`. It checks dangling
relationship ids, missing content-type overrides, relationship targets that
resolve to no part, and orphan media. `rpptx` adds its own presentation-specific
checks, listed in `06-presentationml-model.md`.
