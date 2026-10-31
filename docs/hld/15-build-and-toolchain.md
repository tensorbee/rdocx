# 15, Build and toolchain

## Toolchain pinning

The repository pins its development toolchain in `rust-toolchain.toml`:

```toml
# rust-toolchain.toml
[toolchain]
channel = "1.97.1"
components = ["rustfmt", "clippy"]
targets = ["wasm32-unknown-unknown"]
```

MSRV stays declared separately as `rust-version = "1.93"` in
`[workspace.package]`, and the MSRV CI job continues to pin it explicitly. The
two answer different questions: the toolchain file says what this repository is
developed with, `rust-version` says what it is guaranteed to compile with.

Public setter inlining is selective. The native Word Run bold, italic and size
setters and their optional-value implementations carry inline annotations for
separate downstream consumers. Qualification uses an external crate with
opt-level3, LTO disabled and one codegen unit, with assembly inspection and
paired archived executables. Paragraph-mark setters already inline under that
consumer profile. There is no blanket annotation or end-to-end speed guarantee.

## Deterministic rendering

Normal `oxml-layout` construction may load system fonts. System fonts differ
between machines and between a developer's machine and a CI runner.

**This makes any recorded render baseline unreproducible**, which would poison
both the hash harness and the SSIM gate: a baseline recorded locally would
mismatch CI in a way indistinguishable from a real regression.

`oxml-layout` provides an explicit deterministic mode:

```rust
impl FontManager {
    /// Bundled fonts only. System font loading is bypassed entirely.
    /// Rendering is then reproducible across machines and platforms.
    pub fn new_deterministic() -> Result<Self>;
}
```

The constructor starts with a fresh font database and loads only the checked-in
bundled font bytes. It never calls system-font discovery or clones the normal
process snapshot. Document-embedded fonts remain explicit layout inputs, so
they are deterministic too. Caller-font construction starts from an empty
database and likewise cannot observe bundled or system fonts.

Word font authoring takes bytes directly from the caller and stores the
obfuscated bytes as document package parts. The font-table relationship and
OOXML font key are the only source for deterministic deobfuscation. The font
record name becomes the layout alias for the loaded face before bundled
fallback. No authoring or render path discovers a host font for embedding.

`Engine::new_deterministic()` and `layout_document_deterministic()` carry that
database through layout. The public facade exposes
`Document::render_page_to_png_deterministic()`,
`Document::to_pdf_deterministic()`, and deterministic SVG page export. All
three reuse the separate cached bundled-font-only layout. The PDF facade passes
that layout directly to
`oxml_pdf::render_to_pdf_checked`, which gives the golden-PNG gate
deterministic PDF input and refuses failed selected-font preparation before
output. The legacy vector operation retains its signature and returns no
partial bytes on preparation failure. Existing constructors and rendering
methods still load system fonts for library users.

The hash harness, golden-PNG gate and SSIM harness use the deterministic path.
The normal rendering API still discovers system fonts, but now captures the
bundled plus system face table once per process. Installing, removing, or
replacing system fonts requires a process restart. File-backed font bytes use a
separate bounded process cache keyed by canonical file identity, so faces at
different TTC indices share one byte buffer. Both process caches are compiled
only with `system-fonts`, and poisoned file-cache locks recover by rebuilding
the requested entry.

Font sources enabled by Cargo dependency feature unification, including
`fontdb/memmap` shared files, use `Database::with_face_data` to retain owned
font bytes and the selected face index. This also handles file sources when
host discovery is disabled. Binary and ordinary system-file caches retain
their existing paths. Unified source features do not enable host discovery.

The Word SSIM harness reaches that deterministic path through the production
`rdocx-cli render` command. Oracle-only normalization happens after both raster
trees exist. It cannot alter layout, pagination, font selection, or renderer
dimensions.

The M23 authoring gate is owned by
`scripts/docx_authoring_conformance.py`. CI invokes its public mode before the
workspace test suite. That mode builds an offline temporary consumer with only
the public `rdocx` dependency and uses the same deterministic 150 DPI render
path twice. The public Cargo child keeps its300second timeout, original argv,
working directory, environment and target fallback. A timeout or nonzero exit
retains the call-site phase, exact command and typed stdout/stderr in the deciding
failure before temporary source cleanup. Bytes use reconstructable escapes,
including non-UTF8 bytes. Empty bytes, empty text and absent streams remain
distinct. Only allowlisted inherited target, toolchain and dev-profile settings
are reported, with unspecified values explicit. The diagnostic identifies the
Python host platform without inferring an unobserved compiler version.

The generated public client emits stderr checkpoints at main entry and before
each of its two deterministic renders. Missing entry output establishes only
that entry was not observed. It does not prove compilation or a Cargo lock wait.
Checkpoint output narrows observed progress without diagnosing its wait cause.
Synthetic harness checks qualify diagnostic retention only. Actual source-bound
MSRV public conformance and clean final CI remain required to qualify the child
execution. Private diagnostics and artifact non-disclosure remain unchanged.

Its local private mode uses exact LibreOffice 26.2.5.2 build
`cd7284b4cbbfeb507e630c1aac019f4157393acb` and pdftoppm 26.01.0 identities.
Private source packages, digests, thresholds, and rendered evidence remain in
the ignored local corpus directory and are never fetched by CI.

The pinned Writer process receives one oracle-only static Thin instance for
the Noto Sans SC fixture. `hb-subset (HarfBuzz) 13.2.1` instantiates it at
`wght=100` from the exact checked-in product font with source SHA-256
`b06144fa7b2d5212fe21344261449c9350f603e3e2ae625e76306022d024fbe5`.
The output SHA-256 is
`390ba9f55d4dd69915736d2b225d602b40012cd2c50db4c1e6d2bbdfd61e63a6`.
The binary, authentic OFL licence, exact command, and provenance record live in
`scripts/oracle-fonts/`. Harness inventory assertions bind all three files and
both hashes. Product code continues to consume the published variable font,
and hosted verification does not need the generation tool.

Tagged Word PDFs use the same deterministic layout and writer path. Structure
node references, page-local MCIDs, parent-tree keys, conditional PDF/UA
metadata, and XMP bytes are derived only from ordered layout input. A
`.notdef` glyph suppresses the PDF/UA claim. These outputs add no clock, random
identifier, host lookup, or system-font dependency. Untagged Presentation
layouts keep the existing writer path.

PDF/A-2b and PDF/A-3b use the same deterministic layout boundary. The writer
adds no clock, random source, network lookup, or host colour service. Its fixed
XMP and file identifiers depend only on the selected profile and ordered
layout metadata. Ordinary PDF output remains on the existing byte-compatible
path.

Reusable managers bound shaping to 2,048 entries and 16 MiB, file bytes to 256
entries and 128 MiB, and coverage, resolution, and paragraph traces by explicit
entry ceilings. The reusable Word engine bounds paragraph state at 4,096
entries and 50 MiB, table state at 32 entries and 2 MiB, header and footer state
at 64 entries and 4 MiB, and aligned restart page and checkpoint state at 1,024
entries. Restart candidates have no independent byte partition. Checked
admission charges current and pending block caches plus the complete candidate
against the 5,216-entry and 64 MiB aggregate limits, and arithmetic overflow
fails closed. Both pending and published queues use retained-capacity
accounting. The thousand-page incremental gate exercises the deterministic bundled-fallback
facade and requires a middle edit to paginate at most two pages while matching
a fresh layout exactly. The related-story gate source-builds 700 paragraphs
with a footnote, endnote, default header, and page-number footer. An unchanged
context retains all but at most two page frames and equals a fresh deterministic
layout. Changed stories and note-reference sequences fall back fully, and
continued notes publish only clean boundaries. These caches add no feature flag
or dependency and remain available in the default-off graphs without enabling
host discovery.

The caller-font relayout gate passes five generated fonts totalling 22 MiB and
40 aliases through the deterministic bundled-fallback engine. Structural test
accounting requires zero repeated retained-context font bytes on an unchanged
warm call while exact font-manager invalidation and checked transfer remain
byte-exact. The unchanged 49-entry hash harness, both WASM target checks, the
no-default-features suite, and the existing release-mode 1,000-page performance
test cover this private comparison split. No binding metadata, feature, public
type, dependency, or package inventory changes.

`rpptx-render::layout_presentation_deterministic` applies the same rule to a
whole presentation. It shares page lowering with
`layout_presentation`, bypasses system-font discovery, and adds only explicit
font files from `RenderInput`. The `rpptx` corpus example is an unpublished
development target and does not change any crate publication setting.

Native notes and handout export is available only with the existing `render`
feature and adds no dependency, feature flag, module, binding package, or
published crate. Its public pre-1.0 `rpptx` API requires rustdoc with warnings
denied, both WASM checks, dependency-direction validation, the patched
workspace publish dry run, archive-size enforcement, and separate release
review before publication.

The `oxml-layout` `--no-default-features` path disables host system font
discovery while retaining bundled fonts for deterministic construction. This is
also the same font-isolation path the WASM build needs.

`oxml-layout` pins `hypher` exactly at 0.1.7 with default features disabled and
only `alloc`, `english`, `french`, `german`, and `spanish` enabled. Its Liang
pattern data is embedded in the published crate graph, adds no runtime file or
network lookup, and is covered by the repository licence policy. Package dry
runs and the archive-size gate cover the added pattern payload. The dependency
lives only in `oxml-layout`, so it does not add a format-family edge or feature
flag.

Builds otherwise stay native. There is no development container. The Linux-only
work, manylinux wheels, `wasm32` checks and the LibreOffice render oracle, runs
on CI runners as it always would have.

Pandoc 3.10 is test infrastructure for the MathML and LaTeX conversion gate.
The Linux x86-64 release archive is pinned to SHA-256
`e0f8af62d0f267d22baa5bcefe6d5dda3a097ccc60de794b759fe03159923244`.
The bounded installer rejects an unexpected platform, archive layout, member,
size, digest, or executable identity, then exposes the verified binary only to
CI. Its 160 MiB extracted-size ceiling admits the authenticated 162,406,703-byte
payload. The installer skips without materializing only the archive's exact
`pandoc-lua -> pandoc` and `pandoc-server -> pandoc` symlinks and rejects every
other unsupported member type. Pandoc and texmath are absent from every
published crate and production dependency graph.

The exact Issue365 conversion fixture has a separate live Pandoc 3.11 gate,
`fx234_issue365_matches_actual_pinned_pandoc_3_11`. Its macOS executable identity
is `pandoc 3.11`, with SHA-256
`23908d6bc347072befa2bfe312f2c1c44a7c9c2e9e60bb5d41069245454b3236`.
The gate runs `--from=latex --to=html --math-method=mathml` and
`--from=html --to=latex`, comparing complete normalized OfficeMath trees in
both directions. This fixture does not replace the checksum-pinned Linux 3.10
installer or its historical evidence.

Native equation qualification separately pins Microsoft Word16.113.2
build16.113.26092012 and Poppler26.09.0. The existing OfficeMath manifest binds
actual application identity, tool bytes, source recipes, native PDF and normal
save/reopen evidence. Authentic Word-bundled Cambria and Cambria Math remain
external caller-font inputs through the existing deterministic render API.
They are neither installed nor redistributed, and bundled Caladea defaults
remain unchanged. Historical Word16.104 build16.104.25121423 and Poppler26.01.0
controls retain their identities and results. Production math reads the existing
resolved ttf-parser0.25.1 dependency directly from rdocx-layout, with no new
resolved package, public font API or external executable. Both WASM graphs,
denied-warning rustdoc, verified local-patched workspace publication dry runs,
actual affected archive inventories and the10MiB ceiling cover this direct edge.

ODP conversion adds no external production tool and reuses the workspace
`zip`, `quick-xml`, and `oxml-media` dependencies. LibreOffice 26.2.5.2 remains
an explicit ignored acceptance gate that runs with an isolated user profile.
Publish dry runs and the 10 MiB archive ceiling cover the additive native API
and private converter module.

Presentation HTML import adds no production executable. The private `rpptx`
module reuses workspace `scraper` 0.27 behind `default-template`. Google Chrome
152.0.7977.65 is an ignored differential oracle only. The native archive,
WASM, rustdoc, dependency policy, and 10 MiB package ceiling remain gates.

Word MHTML conversion adds no dependency, module, feature, executable, or
asset. It reuses the private `rdocx` HTML owner, existing `scraper` and `base64`
dependencies, and the shared atomic path writer. Microsoft Word 16.104 build
16.104.25121423 is an ignored differential oracle only. The exact native
oracle, both WASM checks, rustdoc, the patched workspace publish dry run, the
10 MiB archive ceiling, and the unchanged 49-entry hash harness are gates.

PDF presentation import adds no production executable. The private `rpptx`
module uses `lopdf` 0.44.0 with default features disabled behind the existing
`render` feature. Poppler 26.01.0 is an ignored differential oracle only. The
native archive, both WASM graphs, rustdoc, dependency policy, patched publish
dry run, and 10 MiB package ceiling remain gates.

Modern presentation package classes add no dependency, feature, module, or
asset. The additive pre-1.0 `rpptx` enum and methods plus the new `oxml-opc`
content-type constants require rustdoc, README inventory, patched publish
dry-run, archive-size, and release review before publication.

Modern Word package classes add one approved private `rdocx` module and four
`oxml-opc` content-type constants, with no dependency, feature, executable, or
asset. The additive pre-1.0 native enum and methods require both WASM target
checks, rustdoc with warnings denied, dependency policy, patched workspace
publish dry runs, the 10 MiB archive ceiling, and the unchanged 49-entry hash
harness. Microsoft Word 16.104 build 16.104.25121423 is an ignored no-repair
oracle only.

Independent installed style paint requires the full104 Word and PowerPoint
method matrix and the separately executed eight-owner SVG artifact gate.
PNG and SVG use independently decoded original/static physical contours
painted by qualified resvg0.48.1, usvg0.48.1 and tiny-skia0.12.0 controls.
Quadratic and equivalent cubic representations each meet the unchanged0.99
threshold with source-derived placement, scale and ROI. Wrong-style, origin
and geometry controls remain below0.99. A shared raster backend comparison
cannot expose common backend defects.

PDF uses original/static font-text paint under pinned Poppler26.01.0. Variable
PDF subset glyphs reachable through actual font resources and text operators
match source-derived first-use glyph mappings, contours, bounds and advances
exactly. Original-font per-owner ROIs meet0.99, wrong-weight and wrong-origin
negatives remain below0.99, and drawing commands, repeats and pixels outside
those ROIs remain exact. Equivalent quadratic point encodings need not be
bit-identical under the pinned PDF rasterizer. This qualification permits no
original-font substitution, fitted tolerance or product canonicalization.
Non-PDF exact default, instance, axes, owner and geometry controls remain.

The SVG artifact gate is an explicitly selected ignored library test whose
installed wheel and authenticated artifacts are external prerequisites. It
requires one actual passed test, zero ignored outcomes and all eight owners
against both quadratic and cubic controls. Ordinary library ignores earn no
acceptance credit. Python acceptance never launches Cargo.

Variable-font style controls pin `hb-subset`14.4.0 and authenticate licensed
source bytes, tool identity and independent static400/700 outputs. Actual
advances, outlines and painted geometry are distinct checks. Synthetic controls
use independently specified stroke and shear geometry, without substituting
producer metadata for painting. The existing Thin CJK13.2.1 generation contract
remains separate. Original static no-input deterministic outputs stay exact,
and actual intentional variable or synthesis deltas require independent
attribution before the exclusive hash baseline moves. Font licensing and the
packaged font inventory remain unchanged.

The SVG raster oracle independently shapes each selected static script and
weight with pinned HarfBuzz14.4.0. Only a unique textLength with
spacingAndGlyphs whose target equals that independent natural advance within
half a six-decimal serialization unit qualifies for identity-only removal
from the oracle copy. The raw product SVG and raw raster comparison remain
retained. Nonidentity, unknown, nonfinite and nonpositive adjustments refuse
normalization. This isolates the pinned usvg cluster-width adjustment
limitation without changing product geometry, the0.99 paint threshold or
weight400-versus700 discrimination. It establishes no general browser result.

## Feature flags

| Crate | Feature | Default | Notes |
|---|---|---|---|
| `oxml-opc` | `agile-encryption` | off | Adds native read and fixed-profile write support for password-protected OOXML packages |
| `rdocx` | `agile-encryption` | off | Forwards encrypted native document opens and saves to `oxml-opc` |
| `rpptx` | `agile-encryption` | off | Forwards encrypted native presentation opens and saves to `oxml-opc` |
| `oxml-opc` | `digital-signatures` | off | Adds OPC signature creation, verification, and coverage reports |
| `rdocx` | `digital-signatures` | off | Forwards native signature creation and verification to `oxml-opc` |
| `rpptx` | `digital-signatures` | off | Forwards native signature creation and current-state verification to `oxml-opc` |
| `oxml-cli-support` | `digital-signatures` | off | Converts native reports for the two explicit CLI consumers |
| `oxml-layout` | `system-fonts` | on | Off for wasm, where `fontconfig` will not build |
| `rdocx-layout` | `system-fonts` | on | Forwards host discovery to `oxml-layout` |
| `rdocx` | `system-fonts` | on | Forwards through the complete native layout graph |
| `rpptx-render` | `system-fonts` | on | Preserves host discovery for normal presentation rendering |
| `rpptx` | `system-fonts` | on | Preserves native presentation font resolution |
| `rpptx` | `default-template` | on | The bundled `default.pptx` and private bounded HTML importer |
| `rpptx` | `render` | on | Pulls in `rpptx-render`, `oxml-pdf`, and the private bounded PDF importer |
| `rpptx-wasm` | `render` | off | Adds `toPdf` through the deterministic facade renderer |
| `rdocx-py`, `rpptx-py` | `extension-module` | off | Must stay off for `cargo test` |

The workspace dependency entries for `oxml-layout`, `oxml-pdf`,
`rdocx-layout`, and `rdocx` are default-off so a member can select the exact
graph. Direct native `rdocx` and `rdocx-layout` builds retain default-on system
fonts. The CLI and Python binding opt in explicitly, while `rdocx-wasm` does
not. Its generated `toPdf` method calls the normal `Document::to_pdf` facade,
which therefore uses document-embedded and bundled fonts without host font
discovery in that graph. Native `rpptx`, `rpptx-render`, and the presentation
Python binding retain system fonts through the same explicit forwarding
pattern. Bundled font bytes remain available in both modes. The `rpptx`
security features are native-only named-consumer edges to `oxml-opc`. The
Python and CLI manifests explicitly select digital-signatures through their
facades. Both CLIs also select the default-off shared report feature through
their dependency declarations, so no-default-features builds retain signing.
The unpublished Python support crate has a mandatory signature-enabled OPC
edge for its two concrete binding consumers. Isolated CLI-support default and
no-default graphs exclude OPC and crypto. Native defaults and WASM do not
select signing. Python, WASM and CLI manifests do not select agile-encryption.

Raster encoding dependencies are intentionally narrow. `jpeg-encoder` and
`tiff` are direct `oxml-pdf` dependencies with default features disabled,
because the format-neutral backend is their named consumer. The selected
versions pass the repository license and advisory policy, compile for the WASM
target graph, and do not add a format-family dependency edge.

The finite OfficeMath conversion controls keep source-only recipes and their
actual native producer receipts outside published archives. Their current
Word16.114 font qualification retains physical collection face indices and
Cambria/Cambria Math source hashes, independent of historical Word16.113.2 and
Word16.104 producer pins. Caller fonts are read through the existing APIs,
never installed, copied into archives or redistributed. Fresh installed Python
wheels, scoped native rustdoc, WASM and verified package dry runs qualify the
unchanged public conversion signatures and their bounded semantic expansion.


Python form and CLI values changes require matched ordinary release wheel and
CLI observations before and after implementation. Preserve a permanent copy of
the prefeature CLI before rebuilding its target, and retain source/profile/hash
receipts for both artifacts. Wheel and binary size deltas are observations,
not new publication ceilings. Fresh installations use --no-compile so wheel
members remain byte-identical. Authentication binds the current source and
isolated output directory explicitly, preserving the prefeature artifacts.

## Packaging
The native bibliography module is ordinary `rdocx` source and belongs in its
published source archive. Bound Word inputs, UI receipts, native output graphs
and other external oracle artifacts stay in ignored storage, outside crate
archives. Bibliography introduces no collation dependency or data bundle.
Future dependency additions belong to the separately approved F-X192 contract.
The public layout font resolver retains its existing concrete algorithm and
requires already-cascaded properties, with both layout and bibliography
consumers covered by their scoped checks.

Current source archives must pass verified package dry runs, archived source
and test membership checks, README footprint checks and the crates.io10MiB
limit. The shared Rust diff retains MSRV, WASM, hash and existing Python
runtime/typing compatibility gates. These checks do not claim dedicated
bibliography Python or WASM APIs. Feature preparation is not publication.
Release family verification and the separate approval at the reviewed SHA
remain required through the release workflow.


`oxml-layout` packages its source and bundled font assets through an explicit
manifest inventory:

```toml
[package]
include = [
    "src/**/*",
    "fonts/*.ttf",
    "fonts/LICENSE-*",
    "fonts/NOTICE-*",
    "fonts/SUBSET-*.md",
]
```

The dedicated package CI job compares `cargo package -p oxml-layout --list`
against all 27 TTFs, the four family licence files, the Caladea and Noto
notices. The manifest also includes the Simplified Chinese, Hebrew, Korean and
Japanese subset records. The job then runs verified packaging without
`--no-verify` and rejects a missing archive or one larger than the crates.io
10 MiB limit. `oxml-layout` is a published 0.1.2 package, while the release
workflow remains the authority for every later publication.

The external PowerPoint and Word corpora remain outside every published crate
under the ignored `corpus/` directory. Their tracked manifests pin immutable
source URLs and SHA-256 values. The Word manifest also pins one of each required
category plus the reviewed `Apache-2.0` or `MIT` licence identity and immutable
licence URL.

Oracle-only fonts under `scripts/oracle-fonts/` are test infrastructure. They
are absent from every crate include list and generated archive. The static CJK
fixture therefore changes neither the 24-font `oxml-layout` package inventory
nor its archive size contract.

Caller-authorized embedded fonts are runtime document parts, not workspace
assets. Their exact license identity travels in the document font table and
does not alter any crate include list, bundled-font license file, or published
archive inventory. Package dry-runs therefore retain the existing 10 MiB
archive ceiling.

A separate crate-local packaging rule applies to
`crates/rpptx/assets/default.pptx`. **An asset must live under its own crate's
directory**: a workspace-root `assets/` compiles locally but is not collected
into the published tarball.

Public API changes in published crates run the verified workspace packaging
dry run against local patches for every internal crate. Every generated
archive must remain below the crates.io 10 MiB limit. Default-off features are
present in package metadata even when archive verification builds the default
graph.

The format-neutral semantic types are part of the published `oxml-layout`
source inventory. The tagged writer implementation, including
`src/structure.rs`, is part of the published `oxml-pdf` source inventory. No
runtime oracle, installer, validation profile, or generated PDF is packaged in
either crate.

`oxml-pdf` also packages `assets/sRGB2014.icc` and
`assets/LICENSE-sRGB2014`. The 3,024-byte profile comes from the International
Color Consortium registry and has SHA-256
`384b832de3412066743b52a75ee906b6fb9fb8d9e09e936fc2c43223815c6e0a`.
The adjacent legal file records the ICC distribution terms, source, retrieval
date, byte size, and digest. The verified package inventory must contain both
files and the shared raster encoder source graph while every generated archive
stays below the crates.io 10 MiB ceiling.

The publishable `rpptx-cli` binary contains ten commands. Its `thumbnail`
command uses the deterministic presentation renderer, and its `outline` and
`comment` commands depend only on the facade. The package dry run and
archive-size gate therefore cover the complete command surface without adding
runtime assets to the CLI crate.

Rust tags build one selected CLI family across six native targets. GNU Linux
x86-64 and arm64, macOS Intel and arm64, and Windows x86-64 use the CLI default
system font feature. Linux x86-64 musl builds with that feature disabled so the
binary is static and uses the bundled font inventory. Tar archives contain an
executable named `rdocx` or `rpptx`, while the Windows ZIP contains the matching
`.exe`. Every archive also contains the selected CLI crate README as
`README.md` and the workspace `LICENSE`. The aggregate job checks that exact
inventory and the executable mode, compares both prose files with the reviewed
sources after CRLF-to-LF normalization, and writes `SHA256SUMS` over all six
archives.

Every bundled font family has its licence under the crate's `fonts/` directory.
Caladea ships with the full Apache License 2.0 text in `LICENSE-Caladea` and its
copyright, trademark and designer attribution in `NOTICE-Caladea`. The
`bundled_fonts.rs` module documentation identifies Caladea as Apache-2.0 and
the Carlito and Liberation families as SIL Open Font License 1.1.

The bundled inventory includes original Gelasio upright and italic variable
faces and Selawik Regular and Bold. The9 approved added assets total444812
bytes, including licenses, notices and the3968-byte Noto Sans Symbols2 subset.
Gelasio retains upstream commit9228e69b160d79f33950e026293f6e13ba9780d0.
Selawik retains official release1.01 at commit
b65ff9ce7c4673e75ec5c622924a6b6f61f2f876. Original family names, axes and
license texts remain unchanged. Gelasio and Selawik use their own
`LICENSE-*` and `NOTICE-*` files. The symbols subset uses `LICENSE-Noto` and
its exact pinned recipe and attribution in `NOTICE-Noto`. Package membership
includes these assets without replacing existing unrelated font bytes.

Licensed Georgia, Segoe UI and original symbol-font references remain local
oracle inputs read in place, never redistributed. Independent metric controls
pin physical font hashes, selected face/axis, HarfBuzz14.4.0, scale, language,
direction and kern mode. Isolated advances, shaped strings, generic extents,
Word line selection and wrapping are separate comparisons. Real ligature,
kerning and vertical-metric differences remain explicit and asserted.
A finite measured maximum is not a universal tolerance or a paint guarantee.
No genuine Segoe UI Bold reference is asserted where only Regular was measured.
Deterministic baselines require reviewed actual integrated output deltas.

Caller-supplied licensed directories remain the optional-font route. There is
no downloaded default font pack, automatic network loader or new pack crate.

## Publishing

The dependency order grows to roughly twenty crates:

```
oxml-core -> oxml-opc -> oxml-media -> oxml-drawing -> oxml-layout -> oxml-pdf
  -> oxml-sml -> oxml-cli-support -> oxml-chart
  -> rdocx-oxml -> rdocx-layout -> rdocx-html -> rdocx -> rdocx-cli
  -> rpptx-oxml -> rpptx-chart -> rpptx-layout -> rpptx-render -> rpptx -> rpptx-cli
```

The fifteen crates.io candidates in this graph are owned by `mantissaman`:
`oxml-core`, `oxml-opc`, `oxml-media`,
`oxml-drawing`, `oxml-layout`, `oxml-pdf`, `oxml-sml`, `oxml-cli-support`,
`oxml-chart`, `rpptx-oxml`, `rpptx-layout`, `rpptx-render`, `rpptx-chart`,
`rpptx`, and `rpptx-cli`. All 15 implemented packages use the reviewed release
path described below. The earlier 12-package family is published at 0.1.2.
`oxml-cli-support` and `rpptx-cli` are publishable but remain unpublished at
that version. The original 14-package family is published at the immutable
0.1.3 and 0.2.0 boundaries.

`oxml-py-support`, `rpptx-py`, and `rpptx-wasm` are not reserved on crates.io.
The binding crates are not published there. `rpptx-wasm` is an implemented
workspace crate with no crates.io publication path. Its reviewed npm surface is
the local `@tensorbee/rpptx-wasm` bundler tarball. Registry publication remains
unconfigured and unauthorized.

The exact incubating crates.io allowlist now contains 15 implemented shared
and PowerPoint packages. They are
`oxml-core`, `oxml-opc`, `oxml-media`, `oxml-layout`, `oxml-drawing`,
`oxml-pdf`, `oxml-sml`, `oxml-cli-support`, `oxml-chart`, `rpptx-oxml`,
`rpptx-chart`, `rpptx-layout`, `rpptx-render`, `rpptx`, and `rpptx-cli`. All15
were published in S73 at0.12.1 from immutable annotated `rpptx-v0.12.1` tag at
reviewed SHA `58ca5a279277f7cd8de0b8f250fb4650de14371b`. The earlier 0.11.0, 0.10.0,
0.9.0, 0.8.0, 0.7.0, 0.6.0, 0.5.0,
and 0.4.0 registry releases remain available, and no existing version or tag
was moved. Manifest eligibility and allowlist membership do not authorize a
later publication without a separately approved `/release` invocation at the
exact reviewed SHA. The unpublished `rpptx-wasm` preparation member remains at
0.12.1 without gaining another publication path. The failed immutable
`rpptx-v0.12.0` tag
published no registry packages and created no GitHub release.

The complete stable 0.11.1 family is published against the shared 0.8.0 family
from the immutable annotated `v0.11.1` tag at reviewed SHA
`5a850ce9ae6c31f8365594ed2970193266f8b2a6`. It recovers the immutable partial
v0.11.0 attempt, whose tag targets reviewed SHA
`25350d000ed7ed96bf4f6e371f01f8fbc8e2cec4`. That attempt published only
`rdocx-opc` and `rdocx-oxml` before package verification exposed the missing
shared `TextSegment.direction` registry contract. It created no GitHub release
and posted no contribution notifications. All seven 0.11.1 packages and six
leave-open notifications are verified. After separate approval, exactly
`rdocx-opc@0.11.0` and `rdocx-oxml@0.11.0` are yanked. Complete coherent stable
releases remain live and unyanked. The tag is never moved or deleted, no
v0.11.0 GitHub release exists, and no other external state changes. Historical S73
stable source was published as the exact seven-package0.14.0 family from
immutable annotated `v0.14.0` tag at reviewed SHA
`58ca5a279277f7cd8de0b8f250fb4650de14371b`. All seven registry entries and
their sole owner are verified, and the release depends on shared 0.12.1.
The immutable v0.13.0 tag at reviewed SHA
`05332b17f481741e7d5ab4e39699c6d1536475af` published five stable packages,
then stopped before `rdocx`, `rdocx-cli`, and the GitHub release because
registry `oxml-opc@0.10.0` lacks the F-238 Word main content-type constants.

`wheels.yml` accepts new stable `v*` and incubating `rpptx-v*` tags. Before
publication it reproduces the hash harness and validates the tag-driven
selected family: every selected Rust crate and binding version, each workspace
pin when present, and the Python project version. Historical fixed-version
regressions continue to protect the earlier immutable release records in the
local policy suite. The tag preflight is version-agnostic so a reviewed patch
or minor release does not inherit an S73 version literal.

**The same regressions run in the canonical local gate.** `/verify` step 6 runs
`python3 -m unittest scripts.test_sprint_workflow`, the module holding both
preflights and the pinned-toolchain assertions, so a version carrier that moves
without its assertion moving with it fails before the sprint closes rather than
on the tag. Without that step the preflights run for the first time at
publication, which is what S42 demonstrated when F-X022 passed the entire local
gate and still left the incubating preflight and the `ci.yml` WASM literal
asserting the previous version.

The pull-request CI workflow runs the same complete module in its dedicated
`release-regressions` job. The job has no condition or failure-tolerant path,
and checkout plus a locked cargo-release 1.1.3 installation precede the exact
whole-module command. This keeps both release family preflights and future
release-contract regressions in the ordinary CI gate with the external command
their stable-family checks require.

Every stable or incubating tag also requires one reviewed `CHANGELOG.md`
section whose second-level heading is the exact tag. The required ordered
subsections are Highlights, Added, Fixed, Compatibility, and Contributors.
The deterministic workflow CLI checks that contract and renders only the
reviewed body. Both modes are read-only. `wheels.yml` runs the check before
either crates.io allowlist, stores one render in runner-temporary storage, and
byte-compares a fresh render with that artifact immediately before passing it
to `gh release create --notes-file`. Generated GitHub notes are not a release
source.

The workflow then runs
`cargo publish --workspace --dry-run` with an exact local source patch for each
member of the 22-package publishable union. Cargo rewrites packaged path
dependencies to the registry, so the patches keep verification on the reviewed
workspace graph before those versions exist there. They do not enter generated
archives and the dry run uploads nothing. The stable path then publishes only
the seven released rdocx packages in dependency order. The incubating path
publishes only the 15 candidates above in dependency order. Every real command
keeps archive verification enabled. Registry waits separate dependency layers,
and authentication, network, compilation and duplicate-version failures fail
the job.

A single `wheels.yml` tag run builds the selected Rust CLI, Python wheels and
source distribution. The six CLI and seven Python build jobs attest their
produced files using GitHub build provenance with job-scoped OIDC and
attestation permissions. The aggregate job validates exactly thirteen
selected-family payloads, verifies each downloaded subject with `gh attestation
verify FILE -R tensorbee/rdocx` against the workflow and tag source SHA, and
writes one `SHA256SUMS` over all thirteen files. Both registry publication
jobs depend on that gate. Only the crates.io job receives
`CARGO_REGISTRY_TOKEN`. The PyPI job stays directly in `wheels.yml`, under
`pypi` environment OIDC, to preserve its trusted publisher filename.
The GitHub release job waits for both registry jobs and publishes the thirteen
payloads, checksum file and reviewed notes as one release. A partial registry
publication retains the immutable tag and leaves the GitHub release absent.
All third-party actions are pinned to reviewed full commit SHAs. The CLI
manifest cargo-binstall URLs point to these same family tags and archives.

The generated archives remain subject to the crates.io 10 MiB ceiling.
`oxml-layout` contains all 27 bundled fonts and their required legal files, and
`rpptx` contains `assets/default.pptx`. No binding or WASM package is in either
crates.io allowlist.

The deterministic inventory includes official Noto Sans Arabic, Devanagari,
and Thai fonts plus four FontTools-reproducible subsets, Noto Sans Simplified
Chinese, Hebrew, Korean and Japanese, each carrying only its approved fixture
repertoire. A subset covers the fixtures the rendering gates draw and nothing
more, so bundling Korean and Japanese costs tens of kilobytes instead of the
several megabytes a full family weighs. The two upstream sources measured
9.6 MB and 10.4 MB when they were retrieved. The source and output hashes,
subset command, licence, and notice ship with `oxml-layout`. Package checks
verify that inventory and keep the archive below the same 10 MiB ceiling.

Two new release tag namespaces:

| Tag | Workflow | Publishes |
|---|---|---|
| `v*` | `wheels.yml` | stable crates.io family, `rdocx` CLI and PyPI distribution |
| `rpptx-v*` | `wheels.yml` | incubating crates.io family, `rpptx` CLI and PyPI distribution |

The historical `py-rdocx-v*` and `py-rpptx-v*` tags remain immutable release
records, but do not trigger new publication. Manual `wheels.yml` dispatch
builds both Python distributions and uploads twelve `cp39-abi3` wheels and two
source distributions without registry or GitHub release authority. A new
family tag builds only the selected Python project cells. The selected
project's six wheels and one source distribution are validated before PyPI
publishes them through OIDC. There is no long-lived PyPI secret.

## Release process

The unsafe `scripts/release.sh` is deleted. Version changes are targeted F-ID
edits to `[workspace.package]`, the internal pins in
`[workspace.dependencies]`, and `Cargo.lock`. They are reviewed before a tag is
possible and never rewrite README prose by pattern.

`cargo-release` preparation is configured in Cargo metadata. The ten packages
that inherit `[workspace.package].version`, including the unpublished
`rdocx-wasm`, `rdocx-py`, and `oxml-py-support` packages, use cargo-release's
effective `workspace` shared-version group and the `v{{version}}` tag template.
That shared-version group, the `rdocx` Python project, and rdocx WASM
contract literals are prepared at0.17.0, together with the exact seven-package
stable family and its `v0.17.0` release notes. The prepared stable packages
require shared OOXML0.15.0. The previous complete unified stable release is
immutable `v0.16.4` at24af95d3d94c25e95f61a75e7e21f18fa95b6363. The historical
S73 family retains its immutable annotated `v0.14.0` tag at reviewed SHA
`58ca5a279277f7cd8de0b8f250fb4650de14371b`. Its published archives
require shared 0.12.1. The unpublished 0.13.2 crates.io train is superseded
rather than backfilled. The Python, binding, and
WASM carriers remain unpublished on crates.io. The source version move gives no
Rust publication authority.
The immutable v0.11.0 attempt published only `rdocx-opc` and `rdocx-oxml`
before package verification failed against the published shared 0.7.0 API.
The remaining five packages and GitHub release were not published at that
version. Shared 0.8.0 and stable 0.11.1 form the published recovery sequence.
After separate immediate approval, the post-recovery cleanup yanked exactly
`rdocx-opc@0.11.0` and `rdocx-oxml@0.11.0`. Complete coherent stable releases
remain live and unyanked. The complete S73 stable family is 0.14.0. Earlier
immutable registry releases, including the complete 0.12.0 family, remain
available. No binding, WASM, Python, npm, or
incubating package gained publication authority from the stable release.
The 15 publishable `oxml-*` and `rpptx*` package manifests and the unpublished
`rpptx-py` binding crate use explicit version0.15.0, the named `incubating`
group, and the `rpptx-v{{version}}` template. The unpublished `rpptx-wasm`
crate remains at 0.12.1 with its separate npm boundary. The crates.io
allowlist remains exactly15 packages. The previous complete unified family is
immutable `rpptx-v0.14.4` at24af95d3d94c25e95f61a75e7e21f18fa95b6363. The
historical S73 family retains the immutable `rpptx-v0.12.1` release at reviewed SHA
`58ca5a279277f7cd8de0b8f250fb4650de14371b`, and earlier registry releases
remain available. The stable 0.10.1 historical graph proof authenticates the
original registry archive and unchanged published manifests and lock, then
resolves with `--locked`. Its normal dependencies remain registry-sourced,
including `oxml-layout@0.6.0`, rather than the current workspace family. Fresh
unlocked resolution is unavailable because the admissible zip8 versions are
yanked. This historical replay does not change current supply-chain policy. Shared 0.11.0 contains the additive `oxml-opc` constants required by
the stable recovery.
Workspace settings consolidate the preparation commit, upgrade internal
dependency requirements, and retain archive verification. Publishing, tag
creation, and pushing are disabled, and no README replacement is configured.
Preparation changes release carriers, assertions, and selected-family notes
without changing runtime behavior.
Current local verification checks all22 locally patched published archives,
normalized inventory and exact source payloads below10MiB. Fresh local wheels
and source distributions for both families run clean Python3.9/3.12 runtime and
Python3.12 typing checks. Both WASM bindings and current npm carrier contracts
retain their existing separate authority. Final integrated full verification,
sprint review and SHA-bound hosted build-only rehearsal remain mandatory before
close. The shared0.15.0 family publishes before Word0.17.0 where the verified
internal pins require it, each with its own approval at reviewed main.
External release actions remain owned by `/release`.

Strict SemVer runs in three fixed native jobs on Linux x86_64, Mac arm64 and
Windows x86_64 with Rust1.97.1 and official archive/executable pins. Existing
CI signing uses only fixed completed producer artifacts, binding target, job
ID and native wrapper digest in the signed subject name. Certificate source,
workflow, immutable repository IDs and run attempt are independently checked.
OIDC and attestation write permissions belong only to the signing job. The
existing owner-only corpus dispatch remains separate and nonqualifying.

Fresh owned SemVer checkouts enforce `core.autocrlf=false` through checkout-step
Git environment configuration. Measured bytes must equal canonical Git blobs.
Local differing bytes are refused without normalization or global Git changes.
Generated current and baseline projects, locks, library JSON, effective features
and original Cargo cwd are retained separately. Canonical Rustup proxies are
qualified by their actual Cargo-home paths, current and explicit pinned-toolchain
selection and observed bare-tool agreement with the concrete native sysroot.
Proxy byte identity is not assumed, and arbitrary PATH tools remain refused.
Compiler witnesses retain their actual check lock, then acquire all locked target
dependencies through bounded successful fetch before unchanged offline metadata.
The lock, configuration, command and raw fetch observations remain bound through
portable evidence consumption. Optional reuse acquisition failure is retained
and reacquires fresh rows, while mandatory deciding collectors still refuse.
 A failed tool or compiler row
retains available diagnostic bytes and cannot be silently omitted.

Deciding CI, build-only wheels, full verification and release callers acquire
complete paginated API catalogs into the canonical current sprint state outside
supplied evidence. They preserve existing runstate and bind exact source,
policy, artifacts and cutoff. Frozen validation uses only local pinned gh,
unchanged bundles and embedded approved roots, with no hidden network fetch.
The running gate requires completed native/sign jobs without requiring its own
future completion. External consumers require actual successful final-source
CI and reject newer relevant failures. SemVer proof artifacts stay outside all
product distribution directories and checksum globs. The twenty-six build-only
payloads and four skipped publication jobs keep their existing contract.


`/release {vX.Y.Z | rpptx-vX.Y.Z}` is the only command allowed to create or
push a new family release tag or start crates.io or PyPI publication. It
selects one Rust allowlist, one CLI and one Python distribution at a matching
version. The tagged run must contain six CLI archives, six `cp39-abi3` wheels,
one source distribution, one complete checksum file and verifiable build
provenance for every payload. The selected PyPI trusted publisher remains
`wheels.yml` in the `pypi` environment.

`/release-notes TAG` is the deliberate preparation ceremony for the two
new unified namespaces. It derives human-written highlights, additions, fixes,
compatibility guidance, and contributor credit from reviewed repository
evidence, then updates the exact changelog section for review with the code.
`/release` renders and inspects that section at the reviewed SHA before its
separate final approval. After publication it requires the GitHub release body
to equal the same rendered bytes.

The ceremony also builds one selected-family inventory of every included
GitHub issue and pull request. The reviewed notes link each record directly and
credit its authenticated external reporter or contributor with the specific
outcome that landed, including hardened equivalents of unmerged reference
implementations. Before approval, `/release` reports the complete inventory and
the planned record-specific comments. After successful registry publication
and exact release-body verification, it posts those comments with the final tag
and GitHub release link, then retains their URLs in the release evidence. A
missing link, credit, inventory entry, or notification blocks completion of the
release F-ID.

Release preparation F-IDs complete on the sprint branch with local artifact,
workflow and version evidence. `/run-sprint` pushes that reviewed sprint, and
`/close-sprint` merges and pushes it to `main`. Publication does not occur
inside `/run-sprint`. `/release` requires clean `main` at the exact closed
sprint merge SHA, a clean committed sprint review, a tree comparison showing
only the close-sprint tracker summary changed from the reviewed sprint source,
and a fresh read-only `/verify --full` at the main SHA. It refuses a material
tree difference rather than writing a new review record on `main`. A separate
explicit approval is required immediately before each family tag is created.
The tag points to the reviewed `main` merge SHA. `/close-sprint` alone owns
main merges and sprint tags.

The manual `wheels.yml` build-only preflight runs at a reviewed canonical
sprint SHA before sprint close. When actual hosted acceptance needs remote
source, the source availability checkpoint follows passing local checks and
independent review. Preparation remains approved and in-progress until the
named SHA-bound external gates pass. Downloaded artifacts and metadata are
validated, installed and tested in Python3.9 and3.12. Final integrated full
verification and sprint review remain due after preparation.
Bare native and musllinux wheel runners execute one exact positive inventory
of86 installed priority and compatibility cases, partitioned into42 Word and44
PowerPoint nodes. Both paths pass the family list as structured pytest
arguments. Missing or renamed nodes fail collection, and policy checks reject
missing, additional or duplicate inventory entries. Metadata, native imports
and fresh Alpine installation remain mandatory. These hosts do not prepare
the external font, Poppler, LibreOffice or CLI acceptance environment.
The equipped Python CI retains both full suites, including all104 font cases
and the separate eight-owner native SVG gate. The complete Issue158 Word and
Issue217 PowerPoint CLI chains remain in that equipped environment. The
Windows `rpptx` CLI runs on an eight MiB thread stack so complete deck
validation does not overflow the default main-thread stack. The `rdocx` CLI
uses the same Windows stack for complete document comparison. Geometry
comparisons use explicit equal-length checks that work on Python3.9. Strict
typing and stub checks run separately under Python3.12. The same
version carriers and source are then checked
at the main release boundary. Earlier published `py-*` tags and PyPI versions
remain historical records. New releases use the two unified family tags.

When a later sprint wave depends on an integrated and reviewed F-ID that is not
completed, `/run-sprint` uses a resumable dependency-prefix checkpoint before
that consumer. Passing scoped evidence and focused integration reconciliation
complete the prefix and its delivery records before dependent work begins.
The checkpoint does not run full workspace verification or sprint review.
A release preparation dependency follows that ordinary completion route and
does not publish during the sprint. After `/close-sprint` merges the reviewed
sprint source, `/release` checks the exact main SHA, runs full verification
and requests separate approval before each family tag. Final integrated
verification and review remain bound to their actual HEAD, and earlier
checkpoint evidence cannot qualify a changed final source.

`sprint_workflow.py init --resume` treats `CURRENT_SPRINT.md` as the canonical
source for each F-ID's title and size. It refreshes those fields and adds newly
listed F-IDs while preserving the existing phase, feature state, owner, wave,
worker handoff, review, and verification records.

A new family release tag starts `wheels.yml`. Before either registry job, the
asset job proves that the tag points to the current `main` commit and its
closed sprint tag. It checks the exact selected publishable workspace crates,
their lockfile versions and their unpacked package inventory after the full
workspace publication dry run. Cargo removes dry-run `.crate` files on a fresh
target directory, so the selected inventory checks each package directory and
its manifest. The job also checks all thirteen assets and their
producer attestations. The Rust job reproduces the deterministic hash baseline
before crates.io publication. The Python job publishes through the existing
PyPI trusted publisher. A successful release requires every selected registry
entry and owner to report the requested version, plus one matching GitHub
release at the reviewed `main` SHA. `rdocx-wasm` inherits the stable workspace version
but stays `publish = false` because its distribution path is npm.

The current Python release boundary is reviewed SHA
`58ca5a279277f7cd8de0b8f250fb4650de14371b`. Immutable tags
`py-rdocx-v0.14.0` and `py-rpptx-v0.12.1` select only `rdocx 0.14.0` and
`rpptx 0.12.1` respectively. Each live PyPI version contains six `cp39-abi3`
wheels and one source distribution. Tag workflows `35314444102` and
`35317913616` passed exact publication-set validation and trusted publishing,
and each matching GitHub release uses the byte-identical reviewed changelog
body. Both projects are owned by `mantissaman`. Earlier immutable Python
releases remain available.

Both Python `pyproject.toml` versions and their Rust binding crate versions are
exact release carriers. The Python preflight checks all four values instead of
rewriting them during the release ceremony.

Closed sprint release guards accept one annotated integer or fractional sprint
tag, such as `s91` or `s91.1`, at the exact current `origin/main` SHA. The shape
is `^s[0-9]+(\.[0-9]+)?$`. Exactly-one, annotated-object and main-SHA checks
remain mandatory before publication archive validation.

## CI job matrix

Listed in `12-testing-strategy.md`. The matrix carries these
repository-specific gates:

**Path-filtered expensive jobs with one aggregate gate.** A `changes` job uses
`dorny/paths-filter` v4.0.3 at immutable reviewed commit
`ceb8a2b8f2d89434be7ff52d3de7ec3738c5cc9d`. Its inline filters cover the full
inputs of `test`, `msrv`, `wasm`, `python-bindings`,
`presentation-fidelity`, `word-fidelity`, `hash-harness`, `supply-chain`, and
`prose`. Every filter also covers the CI workflow itself. The detector alone receives
`pull-requests: read` in addition to repository content read.

The `ci-gate` job always runs and depends on change detection plus every
filtered job. It fails unless selected jobs succeeded and unselected jobs were
skipped. This includes failure, cancellation, unexpected-skip, and
change-detector failure paths. A documentation-only change runs prose while
the filtered product jobs skip, yet the same stable aggregate gate reports.
The scheduled path skips change detection and requires supply-chain success.
Unfiltered jobs retain their ordinary triggers.

Repository ruleset `21823007` is active for repository `tensorbee/rdocx`,
targets branches, and includes `~DEFAULT_BRANCH`. Its only rule requires exact
status context `CI gate` with
`strict_required_status_checks_policy=false` and
`do_not_enforce_on_create=false`. Its sole bypass actor is `RepositoryRole`
ID 5, role `admin`, in `always` mode. Effective rules on `main` expose only
this required status check. This narrow bypass preserves the direct reviewed
sprint-close push without exempting ordinary pull requests.

The hosted protection proof uses reviewed and verified S58 SHA
`31c51f04f1a9e7c6a198ef16eebba0d782a5827a`. Docs-only PR 59 passed run
`33275852961` and required CI gate job `99162339881` while all eight
path-filtered product jobs skipped. Selected-failure PR 60 failed run
`33276064981` and CI gate job `99162924862`, and GitHub reported its merge state
as blocked while the administrator viewer retained bypass authority. Both
disposable pull requests are closed and unmerged. Their verified remote head
refs were deleted and are absent. Their disposable worktrees and local branches
were removed cleanly.

**A dedicated `oxml-layout` package job.** It checks the exact 32-font and ten
licence-and-notice-file inventory, builds and verifies the generated archive,
and enforces the crates.io 10 MiB limit. Package membership retains all
existing unrelated font and legal bytes.

**`--exclude rdocx-py --exclude rpptx-py` on every `--all-features` job.**
`pyo3/extension-module` tells the linker the Python symbols come from the host
interpreter, which is false for a test binary. On Linux this is an
unresolved-symbol link failure that is easy to misdiagnose as something else.

**A dedicated no-default-features job.** It runs `cargo test -p oxml-layout
--no-default-features`, which exercises the font-isolation path used by WASM.

**The golden-PNG gate in the test job.** After the full workspace suite, the
same Ubuntu 24.04 job runs `python3 scripts/golden_png_harness.py --check` with
the Poppler 26.01.0 installation already on `PATH`. The step is unconditional
and propagates a decoded-pixel mismatch as a CI failure. The external raster
manifest is distinct from the product hash manifest. An approved bundled-font
behavior change must propagate its measured affected samples to both under
the exclusive baseline lease, with attributed review. Unrelated samples,
dimensions, DPI and exact comparison remain unchanged.

**The Pandoc texmath gate in the test job.** The Ubuntu 24.04 runner installs
the exact SHA-256-pinned Pandoc 3.10 archive, verifies the executable identity,
and runs the ignored `rdocx` structural differential by its exact unit-test
path before the workspace suite. The installer caps extraction at 160 MiB and
skips only the exact unneeded `pandoc-lua` and `pandoc-server` aliases to the
verified `pandoc` executable. The installer and gate have mutation coverage in
the release-regression suite.

**The large-document gate in the test job.** The Ubuntu 24.04 runner executes
the exact locked `rdocx` regression binary in release mode, selects only the
ignored 1,000-page performance test, and uses one test thread. The test measures
deterministic layout and direct PDF rendering separately, including their
generation-isolated peak heap allocations. Workflow mutation coverage rejects
removing or weakening any part of this invocation and rejects swallowed
failures.

The same constants bound the four dated README speed rows. The root,
`rdocx-layout`, `oxml-pdf`, and `rdocx-py` pages repeat those rows exactly.
The validator reads the four regression constants and rejects a published
throughput guarantee above its floor or an allocation guarantee below its
ceiling.

**Workspace package READMEs in the docs job.** Every one of the 27 workspace
packages explicitly declares one distinct README. The root file is the
high-level `rdocx` guide. It has exactly three compiling Rust examples and a
seven-row implemented-workflow summary backed by the modern DOCX matrix. It
does not lead with capability classifications, the backlog, or sprint state.
Its dated alternatives table is bounded to exact reviewed official sources and
claims for rdocx, python-docx, docx-rs, docx4j, and Aspose.Words. The other 26
packages use focused crate-local files. Each leads with an outcome and at least
three implemented capabilities before direct-use guidance, neighbouring
package boundaries, publication status, and an example suited to the actual
consumer surface. The three deprecated shims direct new consumers to
`oxml-opc`, `oxml-pdf`, and `oxml-chart`.

Measurement evidence stays outside the alternatives table. Every measurement
row supplies an exact value, version, platform, build mode, input, command,
statistic, and date. The root and 22 publishable crate pages record their
re-derived `.crate` archive footprint. Speed rows appear only on the four pages
whose consumers can reproduce the large-document gate. Python wheel and source
distribution sizes, installed Python footprint, CLI release archive sizes,
WASM bundle sizes, and Python boundary timings stay absent until their pinned
release or human workflow records them.

After the workspace documentation build, `scripts/readme_doctests.py` checks
the exact 27-package inventory, validates Rust, shell, Python, and JavaScript
snippets, derives root stable-family versions from Cargo metadata, checks local
paths and anchors from every README location, and rejects root narrative,
workflow-claim, section-order, crate-boundary, or comparison-evidence drift.
It requires scoped WASM build commands to match the documented package imports
and binds advertised encryption and signing to named default-off Cargo
features. Mutation coverage removes required narrative pieces, adds comparison
rows or duplicate evidence, broadens the dated uniqueness conclusion, and
breaks crate-local links. Its focused network mode resolves the approved official comparison
sources during review, not in default CI. It compiles 23 Rust examples across
the 21 Rust-library READMEs. It discovers each primary and companion rlib from one Cargo build
graph and passes
them to rustdoc with the repository edition, dependency search path, matching
external crate bindings, and warnings denied.
The same runner is part of canonical non-fast verification. It creates each of
the 22 publishable archives, requires exactly one packaged README, and
byte-compares it with the declared source. It re-derives compressed size, tar
member bytes, and tar member count. Cargo's generated
`.cargo_vcs_info.json` is normalized to a fixed-length clean revision for the
exact member-byte total. The compressed archive may differ from the recorded
observation by at most 64 bytes under the pinned toolchain, or grow by at most
that allowance elsewhere, because the generated commit hash and dirty marker
are not package payload. The `--record-measurements` mode prints the exact
Markdown rows for review. The default validator binds those rows to Cargo
metadata, the pinned toolchain, the 10 MiB ceiling, their approved pages, and
the performance constants. It also rejects incomplete provenance, deferred
placeholder numbers, and unbounded superlatives across the 27-page family.
Version, tag, publication, and release-family metadata remain unchanged.

**A WASM target and Node job.** It installs the `wasm32-unknown-unknown` target,
uses exact Node 24.11.1 and wasm-pack 0.15.0, and checks both facade-backed WASM
crates with the locked workspace graph. It then runs both packages' inline Node
regressions. It installs the official Binaryen version 125 Linux archive after
checking exact SHA-256
`7c3bc16599c8274a04d34a504fe4be2047884f900e0e2da2f6fb9cd667183be4`,
places its `wasm-opt` on `PATH`, and verifies the exact official identity
`wasm-opt version 125 (version_125)`.

Both WASM manifests use release optimization arguments `-Oz`,
`--enable-bulk-memory`, and `--enable-nontrapping-float-to-int`. The job builds
the exact scoped release bundler packages with locked dependencies, packs each
one locally, installs it into a separate fresh consumer with an isolated cache
and scripts disabled, and checks exact identity, WASM, JavaScript glue, public
TypeScript declarations, and imports. The package gate grants no registry
authentication, token, OIDC, publication, release, or tag authority. The
document suite also requires generated `toPdf` to return a complete PDF with an
embedded bundled Carlito font. Checkout, setup-node, the Rust toolchain, and the
Rust cache use reviewed full commit SHAs. The presentation render-profile and
optimized-size gates remain local.

**A dedicated Python artifact workflow.** Its product matrix is the Cartesian
product of `rdocx` and `rpptx` with manylinux_2_28 x86_64 and aarch64,
musllinux_1_2 x86_64, macOS x86_64 and arm64, and Windows x86_64. A second
two-package job builds the source distributions. Native cells install wheels
into fresh environments and run the exact positive family priority and
compatibility inventory, followed by exact mypy and stubtest gates. Each
musllinux cell performs a fresh Alpine install and runs the same family
inventory. Full binding suites and external-tool acceptance run in equipped
CI, retaining their mandatory font, native and rendering controls.

**A pull-request Python bindings job.** It runs on macOS 26 with one matrix row
for each Python package. Every row creates an isolated Python 3.12.9 environment,
installs exact maturin, pytest, and package-oracle versions, builds the current
extension with `maturin develop --locked`, and runs the package's complete
pytest directory. Job-local dev optimization is2, with debug assertions and
overflow checks explicitly enabled. It retains the debug profile and all
existing assertions. No measured speedup follows merely from that setting.
Before pytest, the existing Word rendering test carrier
prepares independent inputs for either family without importing either product.
Pillow12.3.0 and Meson1.9.1 are explicit test and build tools. A SHA-pinned
HarfBuzz14.4.0 source build shapes and subsets the licensed bundled Thai400/700
controls and the original synthetic ital/slnt SFNT literals. Original and static
glyph positions, SVG topology and every coordinate retain the0.008-font-unit
bound. Bundled static-family hashes and licenses are authenticated separately.
Actual Cargo JSON artifacts identify the current family CLI and the existing
rdocx development dependency used to compile the standalone path-only reference
renderer before pytest. Compiler, artifact, source and transitive-version receipts
bind that executable. Tests never compile a missing reference or use product
output as its expected paint. All104 font cases and their0.99 comparisons remain
mandatory, including the separate serial native-eight SVG acceptance.
The Poppler installation supplies reviewed26.01.0 tools and a retained source
archive, actual compiler/command and output-hash receipt. Font comparisons accept
the exact historical binary or this authenticated source build, always using
the same tool for product PDFs and independent references. Version output alone
does not establish that provenance. Build and test failures propagate
directly, with no failure-tolerant condition, inherited pytest override, or
fallback. The top-level pull-request trigger is operative and the job condition
uses only the change detector's `python_bindings` output. Workflow root
permissions are repository content read, and only the detector adds pull-request
metadata read. No job grants an OIDC token. Checkout v6.0.2, setup-python
v6.2.0, rust-cache v2.9.1, and the selected stable rust-toolchain revision use
reviewed full commit SHAs and exact input maps.

**An admitted owner-only compatibility job.** Manual dispatch selects a reviewed
source SHA and one task-owned ephemeral macOS ARM64 runner. Both actor fields
must name the owner. Existing public jobs and their gates remain mandatory.
The private job has one literal shell step and no action, service, container,
publication or credential API. Its pre-Worker admission layer checks the exact
server identity and script before the official Worker can prepare any action.
A bounded credential-redacted proposal and exact same-frame owner admission
precede launch. This is a narrow trusted-job admission boundary, not a host
filesystem sandbox. Runner registration waits for independent review and
owned negative controls. The runner is removed after its one job.

The existing sprint workflow script authenticates the frozen100 executable
closure, a separately admitted two-input native revision, clean Git identity,
the unchanged compiled provider graph and both installed distributions. A
finite process group runs each unchanged specimen once. The job retains all
saved files locally and waits for independent current-attempt qualification.
Only bounded counts and source/run identities enter its summary. No Office
font, media, saved document or captured specimen output is uploaded. A pending
native fact cannot produce a final correct-file metric.

**One checksum-pinned Poppler installer.**
`scripts/install_pinned_poppler.py` downloads the official 26.01.0 source
archive, verifies SHA-256
`1cb944a4b88847f5fb6551683bc799db59f04990f5d8be07aba2acbf38601089`,
and builds only `pdftoppm`, `pdfinfo`, and `pdftotext` in an isolated directory.
It caps the download at 8 MiB and streams at most 2,048 safe archive members
with at most 64 MiB of expanded content. A populated prefix fails closed, so a
successful invocation cannot substitute unrelated binaries that print the
right version. All three finished tools must report exact 26.01.0 identities.
Test, MSRV, both Python binding rows, Presentation fidelity, and Word fidelity
use this single unconditional step before any oracle-dependent command.
Package managers may install build prerequisites but do not install Poppler
itself.

**Pinned corpus-test runtime.** Test and MSRV install uv 0.10.2 with official
`astral-sh/setup-uv` commit
`20cfd1bf945f4377ade1205e4dbc17946fc9a30d`. Each job disables the action cache,
uses only its runner-temporary uv cache, fetches and verifies both pinned
corpora, and runs the broad workspace suite with `RUST_MIN_STACK=8388608`. The
pin makes the python-pptx oracle executable available on a clean Ubuntu host.
Both native jobs prepare HarfBuzz14.4.0 before their font consumers through
`--native-tools-only` in the existing Word rendering test carrier. The exact
source archive, bounded extraction, actual three-tool build, binary identities
and tracked source guards use the same producer as full binding preparation.
The early return requires no installed binding, font fixtures or current Rust
compiler. Separate native environments pin pytest9.1.1 and Meson1.9.1.
The authenticated utility directory is exported through `GITHUB_PATH`.
Full binding preparation and its104 and native-eight controls remain mandatory.

Test and MSRV use job-local development optimization1 with debug assertions
and overflow checks explicitly enabled. Workspace profiles remain unchanged,
and the Python binding rows retain their separate development optimization2.
The complete native suites and oracle controls still execute. Profile settings
alone establish no speedup. Qualification records actual current-source job
and step durations with runner and cache limitations.

The Test and Presentation fidelity deterministic animation commands also use
`RUST_MIN_STACK=8388608`. This is a native test-thread budget and does not alter
product runtime behavior or the separate WASM runtime.

The same two clean Ubuntu 24.04 jobs plus the Presentation and Word fidelity
jobs install
LibreOffice 26.2.5.2 from the official Linux x86-64 Debian archive before the
oracle-dependent command. Retrieval uses the exact old-release route at
`https://downloadarchive.documentfoundation.org/libreoffice/old/26.2.5.2/`,
with the full four-component release in the archive filename. The macOS arm64
Python job uses the corresponding old-release DMG, bounds its download at
320 MiB before readonly mounting, and checks SHA-256
`c99fb4fe574437fc4cb820a4ca15271bca325920861f7139858b36d7f9df78ad`.
Both platforms require the same exact executable build identity. Archive
authentication does not substitute for supported-target installation or actual
source-qualified oracle jobs. The archive SHA-256 is
`2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e`.
The installer streams under fixed download, member, and expanded-byte bounds,
rejects unsafe entries and any populated `/opt/libreoffice26.2` prefix, then
requires exact identity
`LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb`.
It installs the explicit Ubuntu NSS, NSPR, D-Bus, Cairo, GLib, X11, CUPS,
font, and Kerberos runtime-library set needed by that official build. This
makes the unconditional `oxml-chart` viewer tests and the `rdocx` ODT
structural differential and both fidelity gates self-contained. The ODT gate
uses an isolated LibreOffice profile
and rejects any runtime identity other than the exact pinned build.

**A prose and generated-skill job.** It runs `scripts/prose_check.py` and
`scripts/sync_agent_skills.py --check` as separate steps, so voice-rule or
adapter drift fails before integration.

**A dedicated release regression job.** It runs
`python3 -m unittest scripts.test_sprint_workflow` after checkout, without a job
condition, successful fallback or `continue-on-error`. The complete module is
the pull-request gate for release-family version carriers and their workflow
contracts.

The pinned PresentationML corpus fetcher retains all50manifest paths, HTTPS
URLs and exact digests for Test, MSRV and Presentation fidelity. A missing or
stale subject gets at most3attempts after classified transient GET/open/read
failures, with1and2second retry delays. HTTP429/500/502/503/504, temporary
connection or DNS failures and raised incomplete reads qualify. Certificate,
generic SSL, permanent HTTP/DNS, string-valued reasons and unknown errors are
terminal. Exception-valued URLError reasons are unwrapped with a cycle guard.
The60second timeout applies to individual operations, not a total subject deadline.
Local file open/write/close, digest, replace and cleanup failures are terminal,
even when they resemble a transport timeout. Partial bytes are discarded and
failed responses close before retry. Silent EOF with a wrong digest is terminal.
Exclusive owned temporary files protect colliding foreign downloads and existing
destinations. Atomic publication requires complete checksum-matching bytes, and
final whole-directory verification still rejects missing or unmanifested files.
No endpoint fallback, source replacement, broad cleanup or fidelity waiver exists.

**A dedicated Presentation fidelity job** runs on Ubuntu 24.04, primes locked
Cargo dependencies, fetches the pinned 50-deck corpus, installs LibreOffice and
Poppler through the shared checksum-pinned installers, and runs
`scripts/pptx_ssim_harness.py --check`. Its existing corpus command builds and
runs `render_deck` in the release profile with deterministic bundled fonts.
The presentation path filter covers the actual renderer and example closure:
`oxml-chart`, `oxml-core`, `oxml-drawing`, `oxml-layout`, `oxml-media`,
`oxml-opc`, `oxml-pdf`, `oxml-sml`, `rpptx`, `rpptx-layout`, `rpptx-oxml`
and `rpptx-render`, plus root build settings, corpus inputs, pinned installers
and its imported golden PNG helper. Dependency-derived regression coverage
rejects missing routes and unrelated crate selection. Same-source profile
qualification requires exact PNG and render-manifest equality. The harness
rejects any LibreOffice
version other than
26.2.5.2 build `cd7284b4cbbfeb507e630c1aac019f4157393acb` and any pdftoppm
version other than 26.01.0 before rendering begins. This turns a package-manager
upgrade into an explicit pin review rather than an unexplained score delta. The
job records the 0.95 SSIM on 80 percent trend reference. It fails on incomplete
corpus coverage, renderer or oracle failure, missing output, dimension mismatch,
or a dropped bounded shape, but not solely on a missed SSIM trend. The job
retains the gate JSON, render manifest, and per-slide score TSV as its evidence
artifact.

**A dedicated Word fidelity job** runs on Ubuntu 24.04 with exact Rust 1.97.1,
LibreOffice Writer 26.2.5.2, and Poppler 26.01.0. Immediately after the pinned
Rust cache, one `cargo fetch --locked` step materializes the reviewed Cargo
graph before corpus and harness work. It fetches and verifies the five-document
corpus, runs `scripts/docx_ssim_harness.py --check`, and retains the required
JSON and TSV evidence. The harness rejects corpus or tool drift, renderer or
oracle failure, zero output, and a missing or empty evidence artifact. It
records page-count and dimension differences through union-index SSIM scoring
and does not fail solely because the five-document 0.95 on 80 percent trend is
missed. The five source-built script pages separately require raw SSIM of at
least0.95on at least80percent of pages, with every bidirectional page reaching
0.95in its separate evidence check. All36runs explicitly request qualified
kern1at24pt, the shared enabled-kerning intersection of Word and pinned Writer.
Text, four font slots, language, nominal and complex size, direction, paragraph
geometry and deterministic ZIP ownership stay fixed. This oracle domain does
not qualify absent, zero, above-size or inherited Word kerning, whose native K1
controls remain independently deciding. The final source-qualified hosted gate
must execute all five fixtures and the complete existing harness. Diagnostic
fixture scores do not substitute for that run. The aggregate CI gate requires
success when the path filter selects the job.
The harness prepares each Writer input through a locked, offline helper that
calls the existing `rdocx::Document::accept_all` API. It reopens the resulting
untracked copy and rejects remaining modeled revisions before the isolated
Writer profile exports it. This helper shares the workspace Cargo lock and
target directory and adds no published dependency or production entry point.
Workflow regressions keep the fetch singular and immediately after the cache,
reject weakened or misplaced commands, and keep all dependency network access
ahead of the offline helper.

Independent Word legal-paint geometry evidence uses actual Microsoft Word
16.113.2 build 16.113.26092012 exports of frozen static colour and decoration
inputs, with retained input and PDF digests and native export receipts. Pinned
Poppler 26.01.0 supplies raster verification. This component-level authority
is separate from the existing historical native corpus. Product baseline
rendering keeps deterministic bundled fonts. Static native paint inputs do
not qualify tracked change-bar selection, and native versus product font and
page differences prevent a full-page similarity claim. Archive observations
for affected published crates are re-derived from current packaged sources
under the existing exact member inventory, 64 byte compression allowance and
10 MiB ceiling.

## Dependency policy

`deny.toml` keeps the unused ZIP codec bans (`zstd`, `bzip2`, `lzma-rust2`, and
`ppmd-rust`) across the workspace, since `zip` is in every graph through
`oxml-opc`. AES, CBC, CFB, HMAC, SHA, and operating-system randomness
dependencies have the named
default-off `oxml-opc/agile-encryption` consumer and do not enter ordinary,
Python, WASM, or CLI graphs.
Ring, SHA-256, and X.509 parsing have the named default-off
`oxml-opc/digital-signatures` consumer. Ordinary Python wheels and both CLIs
explicitly select that consumer. Native defaults, WASM and isolated base
CLI-support graphs exclude the edge. That feature also uses `base64`, whose ordinary runtime
consumer and binding-graph consequences are described below. The creator uses
operating-system randomness for strict RSA-SHA256 signing and requires a
matching embedded public key. The verifier authenticates with that embedded
key. Certificate-chain trust requires caller policy and no ambient trust-store
dependency is part of the workspace graph.

PDF/A conformance adds no crate dependency. The sRGB2014 bytes are a
crate-local compile-time asset, and veraPDF remains external test
infrastructure rather than production logic.

The single advisory exception, an unmaintained transitive font dependency,
carries its exit route in a comment. Keep that discipline: an ignore without a
stated exit route is a permanent ignore.

The private native Word SVG renderer uses the already pinned workspace
`ttf-parser` edge for selected TTC face extraction and instance-aware ink
bounds. Both installed Python bindings use the existing `oxml-layout` edge for
strict caller font loading. Word selected-page raster rendering uses the
existing `oxml-pdf` edge to paint its one configured layout. Only existing
package dependency-list edges change. Package identities, versions and
checksums remain unchanged, with no public reexport or host-directory WASM API.
These source changes retain Rust1.93, no-default layout and both WASM graphs,
including the Presentation render profile, as completion gates. Actual changed
crate archives require source/member correspondence, the unchanged compression
tolerance and10MiB ceiling, and verified patched publication dry runs.

New dependencies are added only with a named consumer. The workspace already
minimises feature sets deliberately, and the comments in the root manifest
explaining why `zip` and `fontdb` are trimmed should be preserved rather than
regenerated.

`oxml-layout` is the direct consumer of exact `hypher` 0.1.7 for conditional
hyphens, `icu_segmenter` 2.3 with compiled data for complex-script boundaries,
and `unicode-bidi` 0.3.18 for paragraph levels and line-local visual order.
These edges remain format-neutral and compile in the no-default-features and
wasm32 graphs.

The private native Word SVG renderer is the direct runtime consumer of
`base64`, which embeds exact layout font bytes and page image bytes into
self-contained data URLs. No `oxml-*`, Presentation, Python, WASM, or CLI crate
adds a direct edge. The Python, WASM, and CLI graphs inherit `base64`
transitively through their ordinary `rdocx` dependency. Exact resvg 0.48.1 is
an `rdocx` development dependency for the 150 dpi SSIM oracle only. It receives
explicit layout fonts, never system fonts, and does not enter the runtime graph
or generated `rdocx` archive.

`pulldown-cmark` is pinned to `=0.13.5` in the existing `rdocx` facade.
Default features are disabled and only the HTML renderer feature is enabled.
Runtime parser options enable tables and strikethrough. Its optional pinned
`pulldown-cmark-escape` renderer edge stays inside that facade graph. Python and
CLI use the same native importer, with no external Markdown executable.

`scraper` 0.27 has two direct named consumers, the private inbound HTML
importers inside `rdocx` and `rpptx`. The `rdocx` edge disables default
features and enables `errors`. The optional `rpptx` edge is activated only by
`default-template`. HTML5 parser repair diagnostics remain available without a
new public feature. No `oxml-*`, `rdocx-html`, Python, WASM, or CLI crate
declares the dependency directly. Both facade graphs pass Rust 1.93, the wasm32
checks, dependency policy, patched package dry runs, and the 10 MiB archive
ceiling.

`lopdf` 0.44.0 has one direct named consumer, the private inbound PDF importer
inside `rpptx`. The optional edge disables default features and is activated
only by `render`. It owns strict bounded PDF object, page, resource, stream,
font-encoding, image, annotation, and content decoding. It does not render and
does not enter an `oxml-*`, Python, WASM, or CLI crate directly.

The workspace `zip` 9.0.3 dependency has direct named consumers in the
private ODT reader and writer inside `rdocx`. It retains the workspace's
disabled default features and enables only `deflate-flate2-zlib-rs`. Deflate64
and clock-backed timestamp support remain disabled. ODT adds no separate archive provider or `oxml-*` dependency edge.
The reader validates bounded archive
metadata before parsing. The writer fixes entry order, compression,
permissions, the 1980-01-01 timestamp, and other metadata while bounding XML,
media, entries, and total retained output. Neither direction reuses
`OpcPackage` for a non-OPC format.

The private EPUB writer inside `rdocx` is another named consumer of that same
workspace `zip` dependency. It adds no runtime package. EPUBCheck 5.3.0 remains
external development and CI validation infrastructure. The reviewed W3C
release ZIP SHA-256 is
`6c07e68584b2e2ce2f89fe06e1246dfead3eb36b46b340e7d93524f29dcff6c5`,
and the extracted validator JAR SHA-256 is
`f7f96617c929371821609b88c8484d6dc9f24fe916499863c46094c5fb778a65`.
The tracked test job downloads that exact release, verifies both identities,
exports the verified JAR path, and runs the source-built EPUB oracle as a
required CI step before the full workspace suite. The validator is not included
in any crate archive.

## Bundled PowerPoint template authority

The crate-local default template uses the authenticated Office16:9 master and
all eleven layout rectangles. Fresh core metadata has absent creator,
lastModifiedBy and description, with fixed created/modified timestamps
2000-01-01T00:00:00Z. Imported and user-assigned timestamps remain explicit
source values. Constructors and ordinary save add no ambient wall-clock data.
Template geometry and metadata are separately labelled behaviour changes with
measured hash attribution before any exclusive baseline update. The asset
remains packaged under rpptx/assets and every archive measurement uses actual
current source members.

## Embedded font inputs and consumer recipes

Embedding reads caller-owned licensed font bytes, with no network discovery,
automatic host font lookup or installed-font prerequisite. Single-face sfnt
authoring and bounded EOT/ODTTF decoding reuse existing dependencies and
existing native providers. Python and CLI inputs are bounded at64MiB before
copying. Word checked listing also bounds aggregate projected copies, while
PowerPoint repeated references share one decoded physical body.

Consumer recipes use explicit license authorization and exact identity, and
exercise source-built current APIs and actual CLI help. Deterministic rendering
retains caller-first embedded-second selection and entrypoint-specific fallback.
Notes and unit-consumer qualification pins python-docx1.2.0 and
python-pptx1.0.2 in fresh isolated installed environments. The exact five-input
script and122-file PR333 source retain authenticated archive, body, check,
mode and membership identities. Raw runner output and independent requirement
classification remain separate. Native read-only notes projection changes
require actual affected archive verification and measured inventories before
feature completion, in addition to genuine wheel and consumer checks.
Fresh paired wheel qualification authenticates every archive member, RECORD
digest/length, local Python/stub source and native binary before isolated install.
The first-draft name corpus is an independent immutable122-file input at
PR333 commit193819ca0a1dcf3415017c0be1186fe8df6bcb86. Its76/44 script membership,
both original runners, archive and per-file SHA256/Git blob identities accompany
raw reports and separate requirement classification. Pinned python-docx1.2.0
and python-pptx1.0.2 run against authenticated installed wheels in isolated
Python environments. Package/source, stub, native-library and RECORD identities
are checked without source-path imports. Raw runner failures remain recorded,
including the historical passing-merge divergence marker.
Affected crate inventories and measured archive tuples come from actual verified
packages, with the existing10MiB ceiling. Separate downstream skill guidance
requires its own reviewed patch and submission boundary. Local README changes
do not imply that an external skill repository was updated.

Native font-use fixtures derive fresh names from licensed existing fonts while
pinning all unchanged outline, metric, style and license tables. Their inputs
are not installed on the host. Independent font parsing, bounded stored-part
inspection and pinned PDF extraction/raster tools qualify actual glyph use
without changing the product or the managed Python environment. Native Word
16.113.2 build16.113.26092012 and PowerPoint16.113.4 build16.113.26100421
qualifications retain input, output, app executable and tool identities.

Office PDF comparisons are application evidence, not deterministic product
baselines. Native font-unavailable warnings, font-part removal and synthetic
styles remain recorded. Own-renderer baselines continue to use deterministic
font mode and true face selection. External downstream guidance remains a
reviewed proposal until separately submitted, and integrated sprint gates
remain independent of isolated feature qualifications.


Chart legend native controls pin Microsoft Word16.113.2
build16.113.26092012 and PowerPoint16.113.4 build16.113.26100421, with
source/output hashes and application export receipts. PowerPoint uses local
PDF export with Best for printing. Poppler26.01.0 supplies150DPI raster and
text extraction. Separate font-matched legend controls use the identical
native bundled Aptos font bytes, SHA256
`95980114fcfd42f2f9c446dae429b70582bf2f03097d68433ea9e7d85a49da0b`,
through the existing deterministic caller-font route. No font installation is
required. Original font mismatches and empty-legend native overlay remain
recorded. Component paint qualification does not establish whole-chart or
outer-frame equivalence, and deterministic shared baselines remain unchanged.

## ZIP supply-chain selection

The workspace requires zip9.0.3, with default features disabled and only deflate-flate2-zlib-rs enabled. Its MSRV1.88 is below workspace MSRV1.93, whose actual compile check remains mandatory. The three direct consumers are oxml-opc, rdocx and rpptx. The upstream major changes the native exposed error identity, decoded filename behavior and ordinary raw-container metadata. No new codec, alias or policy ignore is added. Current cargo-deny yanked rejection remains enabled.

Matched ordinary wheel and CLI size deltas include this approved dependency migration and signing feature wiring. Original ZIP8 artifacts and receipts remain separately labelled predecessors. Genuine current archives require fresh measured catalog/README values and verified compilation, exact normalized members/count,64-byte compressed allowance and the existing10MiB ceiling.

ZIP9 also selects its upstream macOS/BSD-only nix0.31.3 dependency with fs and
disabled defaults, plus cfg_aliases0.2.2 for that dependency build. They do not
add a ZIP codec and remain subject to actual target, license and MSRV checks.
