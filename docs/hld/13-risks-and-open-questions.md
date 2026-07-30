# 13, Risks and open questions

## Open questions, to settle before the milestone that needs them

### Q1, preset shape definition provenance (blocks M10)

The roughly 187 preset geometries are defined in ECMA-376 as parameterised
paths. Three sources exist and only one is usable:

- **LibreOffice's table is MPL-2.0**, which is file-level copyleft and
  incompatible with this repository's MIT OR Apache-2.0. **Rejected.**
- The ECMA-376 accompanying files. Licensing needs confirming.
- Deriving the tables from the specification text, which enumerates every
  preset's guides. Slower, but unencumbered.

**Settle this before writing the generator, not after.** The guide evaluator is
needed for `a:custGeom` regardless, so that work is never wasted whichever way
this lands.

### Q2, PyPI name availability

The future crates.io names in the publishing graph are controlled by
`mantissaman` through 0.0.0 placeholders. PyPI has not been checked. If either
`rdocx` or `rpptx` is taken there, maturin's `module-name` allows shipping a
distribution such as `rdocx-python` while keeping `import rdocx`. Claim both
PyPI names as soon as the decision to ship wheels is confirmed.

### Q3, the deck corpus

Fifty real decks are needed, spanning producers. Sourcing them is not
technically hard but they must be redistributable, or the corpus must be
fetched rather than committed. Decide before M8, since the corpus is that
milestone's entry gate.

## Risks, ranked

### R1, silent output drift during the extraction

**The top risk.** M4 and M5 change unit conversion, text-shaping input types and
content-type derivation. All three alter output without failing to compile, and
the existing 320 tests cannot see it.

*Mitigation*: the hash harness, built first and gating every PR. Every
intentional change lands as its own labelled commit with a reviewed delta. Do
not fold behaviour changes into moves.

### R2, the PDF coordinate-system flip

Replacing the per-element Y flip with one global CTM touches every element type
in a shipped, working renderer.

*Mitigation*: its own reviewable commit, landed **before any pptx code exists**,
gated on golden-PNG diffs comparing pixels rather than PDF bytes.

### R3, `Group`-blind collection passes

Three passes in the PDF backend iterate elements flat. Missing one produces PDFs
with absent fonts, images or link annotations, **and only for pptx content**, so
rdocx's suite never catches it.

*Mitigation*: the `walk()` helper is mandatory, all three passes are rewritten
on it, and each gets an explicit test.

### R4, inheritance correctness

The layout-to-master match is by type, not `idx`. The text chain is seven levels
deep and level-indexed. The colour map is per-master and a dark master inverts
it. Wrong here means subtly wrong fonts, positions and colours on real
templates, which users notice and cannot describe.

*Mitigation*: M9 is a standalone milestone with visual differential tests and a
table of sampled theme-colour resolutions asserted to exact RGB.

### R5, schema child ordering

Diffuse, because it touches every writer, and violations are silent until
PowerPoint refuses the file.

*Mitigation*: `OrderedRawChildren`, plus the corpus-wide "opens without repair"
gate at M8 and M11.

### R6, raw-XML preservation against relationship remapping

The preservation strategy that makes lossless round-tripping possible is exactly
what makes deep copy dangerous, because `r:id` attributes hide inside preserved
blobs.

*Mitigation*: `rewrite_rel_ids`, and `add_slide` synthesising rather than
deep-copying so the common path never needs it.

### R7, scope

Charts plus full parity in one release is **17 to 18 months solo** with nothing
shipping until the end. Both halves of that were chosen deliberately. The
duration was not: the plan was estimated at phase granularity and said nine to
twelve months, and the story-level sizing in `14-development-backlog.md`
disagrees. The story-level number is the trustworthy one.

*Mitigations, neither requiring rework*: a second developer takes it to roughly
9 to 11 months, since M7 and M8 parallelise once M6 lands. Or cut a
read-plus-render release at the end of M10, about 12 months solo, which is the
point where the library becomes useful. Charts are self-contained and nothing
depends on them, so M12 can move independently at any time.

### R8, `oxml-layout` packaging

The bundled fonts are 6.8 MB outside `src/`, published today only because
`cargo publish --no-verify` skips the build-from-archive check, and there is no
`include` or `exclude`.

*Mitigation*: an explicit `include`, drop `--no-verify`, and assert the `.crate`
size against the 10 MiB crates.io limit in CI. Roughly 3.5 to 4 MB compressed is
expected, but verify rather than assume.

### R9, index-path aliasing in the Python bindings

An index path addresses a position, not an object, so a handle held across a
structural mutation would silently read the wrong element.

*Mitigation*: the revision counter making it a loud `StaleElementError`, lazy
collections so the idiomatic loop never holds a stale handle, and stable ids in
v0.2.

### R10, the `rpptx-oxml` scope surface

PresentationML is enormous and there is no natural stopping point.

*Mitigation*: the raw-XML passthrough discipline. **If you do not render it, do
not model it.** Stated as a rule in `06-presentationml-model.md` and enforced in
review.

## Known defects being carried

Found during the audit that produced this plan, and each has a story in M1.

### The `rdocx-wasm` save path

`to_docx_bytes` rebuilds a minimal package, silently discarding every part
except `document.xml` and `styles.xml`. Detailed in `10-bindings-spec.md`.

## Assumptions that would invalidate the plan if wrong

- **That a slide is a page.** The entire rendering reuse argument rests on it.
  Verified: `output.rs` is docx-free and `rdocx-pdf` depends on nothing else.
- **That `OpcPackage` reads a `.pptx` unmodified.** Verified through
  `main_document_part()` keying off `officeDocument`. Worth an actual test in M2
  rather than continued confidence.
- **That preset geometry is a data problem once the evaluator exists.** True
  only if Q1 resolves favourably. If it does not, M10 grows by roughly a week
  for a hand-derived table.
