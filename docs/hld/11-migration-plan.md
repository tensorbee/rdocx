# 11, Migration plan

How the `oxml-*` crates are extracted without breaking a shipped library.

Covers milestones M1 through M6. Every step is an in-place `git mv`, so history
is preserved for free, and `cargo test --workspace` is green at every step.

## The safety net comes first

The migration touches **unit conversion** and **text-shaping input types**. Both
change output silently rather than failing to compile. The 64 in-memory
round-trip tests prove structure survives, not that bytes are identical, so they
cannot catch this class of defect.

**M1 therefore builds an output-stability harness before anything moves.**
`crates/rdocx/examples/generate_all_samples.rs` already exercises nearly every
API. For each sample, record a digest of the flushed `document.xml`,
`styles.xml` and `numbering.xml`, plus the page-one PNG at 150 dpi. Re-run after
every step, and treat any delta as a defect until it is explained.

**Deterministic font mode is a prerequisite for that harness**, not an
optimisation. `crates/rdocx-layout/src/font.rs:93` calls `load_system_fonts()`,
and system fonts differ by platform, so a digest recorded on one machine would
not match one recorded on another. The harness and the SSIM gate render from
bundled fonts only, with system loading bypassed. This is also the first thing
to exercise the `--no-default-features` path.

## The facade trick

`matches_local_name` has **323 call sites across 13 files** in `rdocx-oxml`.
Migrating them individually would be a large, risky, reviewer-hostile diff.

Instead, `rdocx-oxml` becomes a facade over `oxml-core`, and **not one call site
changes**:

```rust
// crates/rdocx-oxml/src/lib.rs
pub use oxml_core::{core_properties, raw_xml, units};
pub use oxml_core::error::{OxmlError, Result};
pub(crate) use oxml_core::xml_text;

// crates/rdocx-oxml/src/namespace.rs keeps W_NS and W_PREFIX, and adds:
pub use oxml_core::xml::{matches_local_name, R_NS, MC_NS};
```

The acceptance check is mechanical: `git diff --stat` shows only `lib.rs`,
`namespace.rs` and `Cargo.toml` modified, plus five deletions. The same pattern
moves `Length` with zero churn.

This is what makes the bulk of the extraction low-risk, and it is worth stating
plainly: most of this migration is a re-export block.

## Order of operations

| Step | Crate | Note |
|---|---|---|
| 1 | `oxml-core` | Move five files plus `length.rs`. Add the new unit types, app and custom properties. Make `xml_text` public |
| 2 | `rdocx-oxml` facade | The re-export block above |
| 3 | `Length` | Delete `crates/rdocx/src/length.rs`, re-export from `oxml-core` |
| 4 | `oxml-opc` | Move verbatim, generalise the constructors, add the pptx relationship types |
| 5 | `rdocx-opc` shim | `pub use oxml_opc::*` with a deprecation note. Consumers still compile untouched |
| 6 | Consumers | Flip imports to `oxml_opc` directly. `rdocx::Error::Opc` changes its inner type |
| 7 | `oxml-media` | All new. Then adopt it in rdocx and delete the old helpers |
| 8 | `oxml-layout` | Move the output types, font manager and bundled fonts |
| 9 | **`line.rs` decoupling** | Its own PR, its own review |
| 10 | `PositionedElement` extension | `Transform`, `Path`, `Paint`, `Group`, `walk`, `#[non_exhaustive]` |
| 11 | `oxml-pdf` | Rename, rewrite the three collection passes on `walk`, add the new arms |
| 12 | `rdocx-pdf` shim | `pub use oxml_pdf::*` |

Each edge points only at an already-migrated crate, and each step is
independently revertable.

## The one piece of real API design

**`crates/rdocx-layout/src/line.rs` is the only file in the extraction that
cannot move verbatim.** It imports `CT_TabStop`, `ST_Jc`, `ST_TabJc`,
`ST_Underline` and `Twips` from `rdocx-oxml`.

| Today | In `oxml-layout` |
|---|---|
| `CT_TabStop` | `TabStop { pos_pt: f64, align: TabAlign, leader: Option<TabLeader> }` |
| `ST_Jc` | `Align { Start, Center, End, Justify, Distribute }` |
| `ST_TabJc` | `TabAlign { Left, Center, Right, Decimal, Bar }` |
| `ST_Underline` | `Underline { Single, Double, Thick, Dotted, Dash, Wave, ... }` |
| `line_spacing: Option<Twips>` plus `line_rule: Option<String>` | `LineSpacing { Single, Multiple(f64), Exact(f64), AtLeast(f64) }` |

Tab positions become points rather than twips, because the layout engine already
works in points everywhere else. Replacing the stringly-typed `line_rule` with a
proper enum is a strict improvement.

A roughly 40-line `LineBreakParams::from_docx` in a new
`crates/rdocx-layout/src/convert.rs` keeps the docx side intact. Budget 150 to
250 changed lines across `engine.rs`, `paginator.rs`, `block.rs` and `table.rs`,
plus rewriting `line.rs`'s 11 tests. **Gate hard on the hash harness.**

## Preserve behaviour, do not improve it

Three things are deliberately left wrong during the move, because correcting
them mid-extraction would produce hash deltas indistinguishable from migration
bugs:

- **Unit truncation.** Float constructors truncate toward zero with `as i64` or
  `as i32`. Positive and negative fractional tests pin every `Length`, `Twips`
  and `Emu` constructor. A rounding change shifts every twip, which shifts
  layout, which moves the regression tests' output.
- **`apply_tint_shade`.** Keep Word's 0-255 convention and its naive sRGB
  interpolation, byte for byte. `oxml-drawing` adds spec-correct functions
  alongside under different names.
- **Everything else that looks improvable.** File it as a story, do not fold it
  into a move.

The exception is behaviour that is a **defect**, which is fixed in M1 as its own
commit with a reviewed hash delta: the image counter, the JPEG marker walk, and
core-property resolution.

One intentional delta is expected in M3: content types become sniffed from magic
bytes, so a mislabelled `.png` that is really a JPEG now gets `image/jpeg`. The
harness will flag it. Label the commit accordingly.

## What happens to the published crates

All seven are published at 0.2.0. Downloads are roughly 4,000 on each sub-crate,
almost entirely transitive, and **59** on `rdocx-cli`, which is the honest
human-install signal. This is the cheapest moment this rename will ever be.

| Crate | Fate |
|---|---|
| `rdocx-opc` | 0.3.0 deprecation shim, then stop publishing. The 0.3.x stays on crates.io forever |
| `rdocx-pdf` | Same, over `oxml-pdf` |
| `rdocx-oxml` | **Stays a real crate permanently.** It keeps ~8,700 lines of WordprocessingML |
| `rdocx-layout` | Stays. Keeps the flow model |
| `rdocx`, `rdocx-cli`, `rdocx-html` | Names unaffected |

**Do not yank anything.** Yanking is for broken or insecure releases. It breaks
fresh resolution for existing users and does not remove the crate.

Set each deprecated crate's `description` to "deprecated: moved to `oxml-opc`".
That string is what appears on crates.io search results and docs.rs, and it is
the only whole-crate deprecation signal Cargo surfaces.

A shim is cheap insurance specifically for `rdocx-oxml`, because rdocx's public
API currently **leaks** its types (`CT_PPr`, `CT_SectPr`, `VMerge`, `Twips`)
without re-exporting them, so a downstream user may depend on it directly.

## Repository and link impact

The repository keeps the name `tensorbee/rdocx`, so **no existing link is
affected at all**. crates.io indexes by crate name, docs.rs builds from the
uploaded tarball, and no redirect is involved.

rdocx goes to **0.3.0**. It is a breaking release regardless: `Error::Opc` and
`Error::Layout` change their inner types, `line.rs` is a public module whose
types change, and `PositionedElement` becomes `#[non_exhaustive]`.

## Release tooling

The unsafe `scripts/release.sh` is gone. Version changes are prepared as
reviewable F-ID commits with targeted manifest and lockfile edits. `/release`
then tags the exact fully verified commit after a separate final approval. It
owns the `v*` release namespace, while `/close-sprint` owns `sNN` tags and
`/spec-bump` owns local `spec-v*` tags.

`publish.yml` uses `cargo publish --workspace`, available at the pinned
toolchain. It performs archive verification and propagates authentication,
network, compilation and duplicate-version failures without relabelling them.
