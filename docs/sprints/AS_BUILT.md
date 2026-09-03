# As Built

Append-only completion log. One entry per F-ID, written by `/complete-feature`
at the moment of completion, describing what was actually built rather than what
was planned.

Entries are never edited after the fact. When a later story changes something
recorded here, the later story gets its own entry. The design intent lives in
`docs/hld/`, the plan lives in `.claude/plans/F-XXX-design.md`, and this file is
the record of what happened.

Newest entries at the bottom.

## Entry template

```markdown
### F-XXX, Short title

**Sprint.** SNN
**Completed.** YYYY-MM-DD
**Size.** S | M | L, estimated N days, actual N days

**What was built.** One paragraph. What exists now that did not before, in terms
a reader who has not seen the diff can follow.

**Non-obvious choices.** Anything a future reader would otherwise have to
reverse-engineer from the code, and the reason for it. Rejected alternatives
belong here, not in a comment.

**Deviations from the design plan.** What changed between
`.claude/plans/F-XXX-design.md` and the implementation, and why. "None" is a
valid and common answer.

**Spec sections touched.** The `docs/hld/` sections this story implements or
contradicts. If it contradicts one, say which and confirm the spec was updated.

**Tests.** The test gate from `docs/hld/14-development-backlog.md`, plus any
others added. Name them.

**Hash harness.** Unchanged, or the expected delta and its justification.
Mandatory for every story in M1 through M6.

**Notes for future sessions.** Anything that will not be obvious in three
months. Traps found, assumptions made, follow-up worth filing.
```

## Entries

### F-001, Deterministic font mode

**Sprint.** S01
**Completed.** 2026-07-29
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `FontManager`, the layout engine, and `Document` now expose
an explicit rendering path that loads checked-in bundled fonts and never
discovers host fonts. The existing `bundled-fonts` feature is default-on for
the current `rdocx` consumer.

**Non-obvious choices.** Determinism is explicit rather than ambient. Normal
library rendering keeps its system-font path, while deterministic rendering
returns a clear error when bundled fonts are disabled.

**Deviations from the design plan.** The plan was revised to correct the
manifest's missing default declaration after implementation discovery showed
that the code and HLD already described bundled fonts as default-on. Microscope
pass 1 also strengthened the golden gate to inspect the actual resolved font
bytes rather than compare two calls under one environment.

**Spec sections touched.** `docs/hld/15-build-and-toolchain.md`, "Deterministic
rendering" and "Feature flags".

**Tests.** `deterministic_font_manager_uses_only_bundled_fonts`,
`deterministic_font_manager_requires_bundled_fonts`, and
`deterministic_render_is_independent_of_system_fonts`.

**Hash harness.** Unchanged. F-003 recorded the first baseline after this path
was integrated.

**Notes for future sessions.** The end-to-end test verifies every font buffer
used by the inspected layout belongs to the checked-in bundled set.

### F-002, rust-toolchain.toml

**Sprint.** S01
**Completed.** 2026-07-29
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The repository now selects Rust 1.97.1 with `rustfmt`,
`clippy`, and `wasm32-unknown-unknown` through `rust-toolchain.toml`.

**Non-obvious choices.** The workspace and CI MSRV declarations remain 1.93.
The development toolchain and the compatibility floor answer different
questions.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/15-build-and-toolchain.md`, "Toolchain
pinning".

**Tests.** `rustup show active-toolchain`, installed component and target
inspection, and confirmation of every 1.93 MSRV declaration.

**Hash harness.** Unchanged.

**Notes for future sessions.** Rustup may synchronize channel metadata before
reporting the repository override, even when the toolchain is installed.

### F-003, Output-stability hash harness

**Sprint.** S01
**Completed.** 2026-07-29
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `scripts/hash_harness.py` regenerates seven samples and
compares SHA-256 values for three OOXML parts and deterministic page-one PNG
output at 150 dpi. Check mode is read-only, while update mode requires a
non-empty reason.

**Non-obvious choices.** Missing optional parts are recorded as JSON `null`
rather than omitted. The baseline has 28 sorted entries, so additions,
removals, and byte changes are reported separately.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, "The hash
harness".

**Tests.** Python comparison and reason-refusal unit tests,
`python3 scripts/hash_harness.py --check`, and a temporary writer whitespace
injection that left the structural round-trip test green while changing all
seven `document.xml` digests.

**Hash harness.** Expected initial delta. Added 28 entries with reason
`F-003 initial deterministic baseline`. Manifest SHA-256 is
`9a3c64d61df793b9d8f7203df9cb966fb67201518b4f7fc0f2e68d276aaaca8f`.

**Notes for future sessions.** `invoice` has no `word/numbering.xml`, which is
the single explicit null entry in the initial baseline.

### F-004, Caladea licence and the false OFL claim

**Sprint.** S01
**Completed.** 2026-07-29
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The `rdocx-layout` package now carries the Apache-2.0
licence and Caladea notice beside the four TTFs. Bundled-font documentation
names the correct licence per family, and a test enforces licence coverage.

**Non-obvious choices.** Attribution files live under the crate's `fonts/`
directory so they are included in the published archive with the assets they
cover.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/13-risks-and-open-questions.md`, "Known
defects being carried", and `docs/hld/15-build-and-toolchain.md`, "Packaging".

**Tests.** `every_bundled_font_family_has_a_licence_file`, the full
`rdocx-layout` suite, upstream TTF provenance checks, and the package file list.

**Hash harness.** Unchanged.

**Notes for future sessions.** The checked-in Caladea files match the
`crosextrafonts-20130214` source archive, and the notice fields match embedded
TTF metadata.

### F-005, Fix the image counter

**Sprint.** S01
**Completed.** 2026-07-29
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Imported media names now seed allocation from the greatest
positive numeric suffix rather than the part count. Allocation avoids existing
suffixes and remains collision-free at the finite `usize` boundary.

**Non-obvious choices.** When the greatest suffix cannot be incremented, the
counter wraps to one and skips occupied suffixes until it finds a free name.
Ordinary packages still allocate exactly maximum plus one.

**Deviations from the design plan.** Microscope passes 1 and 2 exposed overflow
and overwrite cases at `usize::MAX`. The plan was clarified to add checked
wrapping and occupied-suffix skipping without adding a media-namer abstraction.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`, "Part naming",
and `docs/hld/13-risks-and-open-questions.md`, "Known defects being carried".

**Tests.** `next_image_name_uses_the_highest_existing_index_not_the_part_count`,
`malformed_media_names_do_not_change_the_highest_image_index`,
`occupied_max_image_suffix_wraps_to_a_free_low_number`, and
`max_minus_one_allocates_max_then_wraps_safely`.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Suffix parsing is extension-independent and
reads only consecutive ASCII digits immediately after `image`.

### F-006, Fix the JPEG standalone-marker walk

**Sprint.** S01
**Completed.** 2026-07-29
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The JPEG dimension walk now handles SOI, TEM, and restart
markers without reading nonexistent lengths, validates every length-bearing
segment, tolerates marker fill bytes, and terminates at EOI.

**Non-obvious choices.** The parser remains a small header walk because PDF
output passes JPEG bytes through unchanged and needs only dimensions.

**Deviations from the design plan.** Microscope pass 1 found that EOI was being
skipped like a restart marker. Pass 2 verified immediate termination and the
new trailing-data regression.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`, "Media", and
`docs/hld/13-risks-and-open-questions.md`, "Known defects being carried".

**Tests.** `jpeg_restart_marker_before_sof_preserves_dimensions`,
`every_truncated_jpeg_header_returns_without_panicking`, and
`jpeg_bytes_after_eoi_cannot_supply_dimensions`.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** SOF still has to appear before EOI. Trailing
bytes after a completed JPEG cannot supply dimensions.

### F-007, Resolve core properties through the relationship

**Sprint.** S02
**Completed.** 2026-07-30
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Document metadata now resolves through the package-level
core-properties relationship, preserves a custom part target across load and
save, and creates the conventional target only when the relationship is
missing. `rdocx-opc` exposes the standard relationship type publicly.

**Non-obvious choices.** The facade retains a private copy of the stable
relationship URI so the `rdocx 0.3.0` package can still verify against the
published `rdocx-opc 0.3.0` dependency before both move to 0.4.1.

**Deviations from the design plan.** The full packaging gate exposed the
published-dependency compatibility issue after workspace tests passed. An
independent microscope pass approved the private URI because both integration
gates cross-check it against the public constant.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`, "Relationship
types" and "Part naming".

**Tests.** `core_properties_at_relationship_target_round_trip_in_place`,
`metadata_round_trip`, focused rdocx and OPC suites, and the clean package
dry-run.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** A non-standard target is authoritative. Saving
must not create an orphaned `/docProps/core.xml` part.

### F-008, Non-consuming setter twins

**Sprint.** S02
**Completed.** 2026-07-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** All 61 consuming builders across `Paragraph`, `Run`,
`Table`, `Row`, and `Cell` now delegate to non-consuming `set_*` twins.

**Non-obvious choices.** Action builders receive literal `set_*` names as the
story required. Existing builder names and chaining behavior remain unchanged.

**Deviations from the design plan.** The backlog's paragraph-level bold gate
was corrected to obtain a `Run`, where bold formatting belongs. Integration
with F-007 required retaining two independent additions to the shared test
file, followed by a clean microscope pass.

**Spec sections touched.** `docs/hld/03-architecture.md`, "Facade conventions",
`docs/hld/10-bindings-spec.md`, "Two supporting decisions", and
`docs/hld/14-development-backlog.md`, "F-008, Non-consuming setter twins (M)".

**Tests.** `non_consuming_setters_mutate_borrowed_wrappers` and
`non_consuming_setters_match_consuming_builders`, plus all 68 integrated rdocx
integration tests.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep mutation bodies in the in-place setters so
builder and binding behavior remain single-sourced.

### F-009, Cache the layout result

**Sprint.** S02
**Completed.** 2026-07-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Document` caches normal and deterministic layout results
in separate thread-safe slots, exposes cloned page layout access, and clears
both caches across direct mutations and mutable-accessor paths.

**Non-obvious choices.** `Mutex<Option<Arc<LayoutResult>>>` preserves the
`Document: Send + Sync` binding contract. Caller-supplied font layouts remain
uncached because their inputs are not part of a stable document cache key.

**Deviations from the design plan.** None. The approved plan had already
replaced the backlog's thread-local `RefCell<Option<Rc<_>>>` proposal.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`, "Performance",
`docs/hld/10-bindings-spec.md`, "Two supporting decisions",
`docs/hld/13-risks-and-open-questions.md`, "Known defects being carried", and
`docs/hld/14-development-backlog.md`, "F-009, Cache the layout result (M)".

**Tests.** `rendering_all_pages_performs_one_layout`,
`document_mutation_invalidates_cached_layout`,
`mutable_accessor_invalidates_cached_layout`,
`font_modes_use_isolated_layout_caches`, and
`document_remains_send_and_sync`.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Any new mutable accessor must invalidate both
layout modes before returning the borrow.

### F-010, Reserve crate names

**Sprint.** S02
**Completed.** 2026-07-30
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Fourteen approved `oxml-*` and `rpptx*` names were
published as dependency-free `0.0.0` placeholders and verified as owned by
`mantissaman`.

**Non-obvious choices.** Python and wasm binding names were excluded because
their documented distribution channels are PyPI and npm. Publications ran
sequentially through crates.io's rolling new-crate rate limit.

**Deviations from the design plan.** None. The registry required repeated
cooldown windows, and the workflow stopped after every HTTP 429 before
resuming at the exact rejected name.

**Spec sections touched.** `docs/hld/13-risks-and-open-questions.md`,
"Q2, PyPI name availability", and `docs/hld/15-build-and-toolchain.md`,
"Publishing".

**Tests.** Exact `cargo info <name>@0.0.0` and owner checks for all fourteen
names, package inspection, publish dry-runs, and archive-size checks.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** The placeholders reserve names only. They expose
no implementation API and do not change any existing `rdocx 0.3.0` crate.

### F-011, Pin unit truncation behaviour

**Sprint.** S02
**Completed.** 2026-07-30
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Fractional positive and negative tests now pin truncation
toward zero for every float constructor on `Length`, `Twips`, and `Emu`.

**Non-obvious choices.** The vectors cross the half-unit boundary so temporary
rounding mutations fail while the existing production casts remain unchanged.

**Deviations from the design plan.** Microscope pass 1 corrected one invalid
HLD heading citation. Pass 2 found no defects or smells.

**Spec sections touched.** `docs/hld/11-migration-plan.md`, "Preserve
behaviour, do not improve it", and `docs/hld/12-testing-strategy.md`, "New
tests the extracted crates need".

**Tests.** `length_float_constructors_truncate_toward_zero`,
`twips_float_constructors_truncate_toward_zero`, and
`emu_float_constructors_truncate_toward_zero`, including temporary rounding
mutations that made every gate fail.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** A change from casts to rounding is a behavior
change even when whole-unit conversion tests continue to pass.

### F-012, Tag v0.4.1

**Sprint.** S02
**Completed.** 2026-07-30
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The workspace was published as seven lockstep rdocx crates
at 0.4.1 from the reviewed S02 SHA. A dedicated `/release` command now owns
`v*` tags and publication, while the tag workflow verifies the deterministic
hash baseline and publishes only the approved rdocx allowlist.

**Non-obvious choices.** The published `v0.4.0` mainline was merged into S02
before release, preserving its contract changes and retargeting the planned
0.3.1 release to 0.4.1. The fourteen `oxml-*` and `rpptx*` placeholders remain
at 0.0.0 until PowerPoint development is complete.

**Deviations from the design plan.** The original plan targeted 0.3.1 before
the separate 0.4.0 release appeared. The reconciled plan and release evidence
target 0.4.1. The publication workflow retains deliberate registry-index waits
because real publication is explicitly allowlisted instead of workspace-wide.

**Spec sections touched.** `docs/hld/11-migration-plan.md`, release boundary,
`docs/hld/13-risks-and-open-questions.md`, release risks,
`docs/hld/14-development-backlog.md`, M1 gate and F-012, and
`docs/hld/15-build-and-toolchain.md`, "Publishing" and "Release process".

**Tests.** `/verify --full` passed at
`6e02a4b6417c9bb0c245237bdf8168dd06310c39`. The package dry-run produced
exactly seven archives below 10 MiB, including all 20 TTFs and required licence
files in `rdocx-layout`. GitHub Actions run 30522998328 passed, every exact
`cargo info <crate>@0.4.1` lookup succeeded, all owners were `mantissaman`, and
the GitHub release tag peeled to the reviewed SHA.

**Hash harness.** Unchanged. All 28 entries matched locally and on the Linux
publication runner.

**Notes for future sessions.** The release workflow must remain restricted to
the seven rdocx crates until PowerPoint development is complete. After S02 is
merged, forward-merge `main` into `feature/release-0.5.0` before that release
branch continues.

### F-013, Create oxml-core

**Sprint.** S03
**Completed.** 2026-07-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A new private `oxml-core` workspace crate now owns staged
copies of the format-neutral error, units, raw XML, XML text, core properties,
and `Length` implementations. It also provides shared namespace-aware XML
helpers and public XML text handling with focused event coverage.

**Non-obvious choices.** The crate remains at 0.0.0 with `publish = false`.
The existing Word implementations stay in place until F-015 and F-016 can
switch the facades without putting an unpublished dependency into a published
rdocx package.

**Deviations from the design plan.** None. The approved plan already specified
the staged copy and delayed facade switch.

**Spec sections touched.** `docs/hld/03-architecture.md`, "Three families, one
workspace", `docs/hld/11-migration-plan.md`, "The facade trick" and "Order of
operations", and `docs/hld/15-build-and-toolchain.md`, "Publishing".

**Tests.** The moved unit, raw XML, core properties, and XML text tests,
`xml_text_handles_cdata_mixed_nested_and_general_refs`, workspace compilation,
package verification, and the dependency-tree direction check.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Do not publish this crate or connect published
rdocx packages to it until the PowerPoint development publication boundary is
explicitly lifted.

### F-014, New unit types

**Sprint.** S03
**Completed.** 2026-07-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml_core::units` now exposes `Centipoints`, `Angle`, and
`Percent1000` with the exact OOXML storage scales. `Length::mm` adds direct
millimetre construction through 36,000 EMUs per millimetre.

**Non-obvious choices.** Float conversions retain Rust cast semantics and
truncate positive and negative fractional values toward zero. No generic unit
abstraction or `Length::to_mm` accessor was added.

**Deviations from the design plan.** None.

**Spec sections touched.** None. The existing glossary and DrawingML model
already specify the implemented types and scales.

**Tests.** `centipoints_round_trip_points`, `angle_round_trip_degrees`,
`percent1000_round_trip_percent`,
`new_unit_float_constructors_truncate_toward_zero`, and
`length_millimetres_round_trip`.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve the pinned truncation rule when these
types gain format consumers.

### F-017, App and custom properties

**Sprint.** S03
**Completed.** 2026-07-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-core` now provides a shared application-properties
union for Word and PowerPoint plus typed custom properties for text, signed
integers, floating point values, Booleans, file times, and empty values.
Parsers preserve child order and retain unsupported property subtrees as raw
XML.

**Non-obvious choices.** Parsed application children replay their encountered
order, while constructed values use canonical schema order. Unsupported custom
value types remain raw XML instead of being coerced to strings.

**Deviations from the design plan.** Microscope passes 1 and 2 found malformed
root acceptance and inconsistent self-closing typed values. Both were fixed
with regression tests before the clean pass 3.

**Spec sections touched.** None. The existing scope, architecture, and testing
documents already describe the shared model.

**Tests.** `word_app_properties_round_trip_without_presentation_fields`,
`powerpoint_app_properties_round_trip_without_word_fields`,
`unknown_app_property_subtree_is_preserved_verbatim`,
`custom_property_value_types_round_trip`,
`unknown_custom_property_value_is_preserved_verbatim`, and malformed-root and
self-closing-value regressions.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep unknown XML preservation and schema child
order intact when adding further extended-property variants.

### F-018, Create oxml-opc

**Sprint.** S04
**Completed.** 2026-07-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A new private `oxml-opc` workspace crate now owns a staged
copy of the format-neutral OPC package, content-type, relationship, and error
implementation. Generic `OpcPackage::new`, `OpcPackage::with_main_part`, and
`ContentTypes::minimal` replace the DOCX-specific public constructors in the
new crate.

**Non-obvious choices.** The crate remains at 0.0.0 with `publish = false` and
depends only on `quick-xml`, `thiserror`, and `zip`. The existing `rdocx-opc`
implementation and every released rdocx consumer remain untouched until the
real shared crates have an approved publication path.

**Deviations from the design plan.** None.

**Spec sections touched.** None. The existing architecture, OPC, migration,
testing, and build specifications already describe the staged extraction.

**Tests.** Eleven moved OPC tests rebuilt around private DOCX helpers,
`minimal_content_types_contain_only_universal_defaults`,
`with_main_part_resolves_and_round_trips`, independent dependency inspection,
local package verification, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep `rdocx-opc` and released consumers on the
published implementation until the deferred F-022 cutover is packageable.

### F-019, PresentationML relationship and content types

**Sprint.** S04
**Completed.** 2026-07-30
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `oxml-opc` now exposes package, shared-property, and
PresentationML relationship constants plus universal, shared-property, and
PresentationML content-type constants. The generic minimal constructor reuses
the universal values.

**Non-obvious choices.** Word-specific MIME constants remain outside this
story. The table-driven gate lists every public value and asserts uniqueness,
the correct relationship namespace, the `application` MIME top-level type,
and the absence of whitespace or extra slashes.

**Deviations from the design plan.** Microscope pass 1 found that the initial
MIME-shape assertion accepted arbitrary top-level types and extra slashes. The
gate was tightened before clean pass 2.

**Spec sections touched.** None. The existing OPC and PresentationML
specifications already enumerate the required constants.

**Tests.** `relationship_and_content_type_constants_are_unique_and_well_formed`,
all-target compilation, a 12,155-byte integrated local package, dependency
inspection, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Add future package constants to the exhaustive
table so a new value cannot bypass namespace and MIME-shape classification.

### F-020, oxml-opc reads a pptx

**Sprint.** S04
**Completed.** 2026-07-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A code-built PowerPoint package fixture writes and reopens
a presentation, slide, and slide-layout graph entirely in memory. It proves
main-part discovery, relationship round-tripping, normalized package keys, and
parent-directory layout resolution.

**Non-obvious choices.** The fixture lives in the existing `package.rs` test
module. It adds no binary fixture, integration-test target, dependency,
production API, or production-code change.

**Deviations from the design plan.** Sprint review pass 1 required a second
fixture built directly as a valid PresentationML ZIP, independent of
`OpcPackage::write_to`, so the M2 real-package gate does not rely on a
self-round-trip.

**Spec sections touched.** None. The existing OPC and testing specifications
already require this exact package graph.

**Tests.** `pptx_package_resolves_main_slide_and_layout_parts`,
`presentation_layout_target_resolves_one_directory_up`,
`independently_built_pptx_opens_and_resolves_relationships`, all integrated
`oxml-opc` tests, and the integrated full gate. The original two named tests
were observed failing for their intended reasons before completion.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** OPC navigation is now proven against both Word
and PowerPoint package shapes before any PresentationML parser is introduced.
The close workflow preserves the S32.2 target for publication-bound cutovers
instead of forcing them through every intervening sprint.

### F-021, Zip-slip hardening tests

**Sprint.** S04
**Completed.** 2026-07-30
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The `oxml-opc` reader now normalizes ZIP entry names before
classifying package metadata and parts. Root-escaping traversal clamps to the
package root, and absolute entries become canonical leading-slash part names.

**Non-obvious choices.** Normalization uses OPC package-path algebra rather
than host filesystem canonicalization, and no archive entry is extracted to
the filesystem. The released `rdocx-opc` reader remains unchanged.

**Deviations from the design plan.** None.

**Spec sections touched.** None. The existing OPC and testing specifications
already require canonical package names and both hostile cases.

**Tests.** `zip_entry_that_escapes_root_is_clamped_to_root`,
`absolute_zip_entry_is_normalized_to_package_root`, the opaque package
round-trip, deterministic save, all integrated OPC parser tests, and the full
gate. Both hostile cases were observed failing before the normalization fix.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve package-path normalization before raw
metadata classification when the reader gains further ZIP validation.

### F-023, oxml-media format sniffing

**Sprint.** S05
**Completed.** 2026-07-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The new dependency-free `oxml-media` crate classifies PNG,
JPEG, GIF, BMP, TIFF, WebP, SVG, EMF, and WMF from bytes. It exposes canonical
extension and content-type mappings plus sniff-first resolution with extension
fallback and a PNG default.

**Non-obvious choices.** The staged crate remains at version 0.0.0 with
publication disabled. Released rdocx consumers still use their existing image
paths until the deferred cutover after PowerPoint development.

**Deviations from the design plan.** Microscope review tightened SVG prolog
handling and standard WMF recognition before the final clean pass.

**Spec sections touched.** None. The existing media and migration contracts
already specify the staged crate and sniffing precedence.

**Tests.** Magic-byte coverage for every format, canonical mappings,
misleading-extension precedence, unknown-image fallback, every-prefix safety,
the dependency and package riders, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep format detection byte-led when the released
consumers move to `oxml-media`.

### F-024, Image probing and DPI

**Sprint.** S05
**Completed.** 2026-07-30
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `oxml-media` now probes pixel dimensions, optional DPI,
bit depth, channel count, and alpha metadata from bounded PNG, JPEG, GIF, BMP,
and WebP headers.

**Non-obvious choices.** All fixtures are constructed in code. Unsupported,
inconsistent, and truncated headers return `None`, and every indexed read is
bounds checked without adding a decoder dependency.

**Deviations from the design plan.** Microscope review found and drove fixes
for BMP mask placement and alpha classification plus WebP frame, profile,
canvas, RIFF-padding, and parity validation before clean pass 4.

**Spec sections touched.** None. The implementation follows the existing
binary-header parsing contract.

**Tests.** PNG `pHYs`, JFIF density units, EXIF before progressive SOF, GIF,
BMP, all WebP layouts, every-prefix truncation, 22 integrated media tests, and
the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve bounded parsing and the truncation
loops whenever another header layout is accepted.

### F-025, MediaNamer

**Sprint.** S05
**Completed.** 2026-07-30
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `MediaNamer` scans exact directory and stem matches for
positive numeric suffixes, allocates after the maximum, and wraps safely after
`usize::MAX` without emitting zero or reusing an occupied name.

**Non-obvious choices.** The caller's extension is preserved verbatim. The
allocator retains occupied suffixes so integer wrap can search safely from one.

**Deviations from the design plan.** Microscope pass 1 found that root package
parts were not recognized. The fix normalizes the empty root split back to `/`
and adds root, trailing-slash, and non-PNG regression coverage.

**Spec sections touched.** None. The existing media contract already requires
maximum-suffix allocation.

**Tests.** All four F-005 sentence-named regressions, root and directory
normalization, caller extension handling, 22 integrated media tests, and the
integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** The released rdocx allocator remains active
until the deferred consumer cutover.

### F-026, native_size with explicit DPI

**Sprint.** S05
**Completed.** 2026-07-30
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `ImageInfo::native_size` returns dependency-free
`NativeSize` values with explicit width and height EMU fields. Each axis uses
finite positive declared DPI when available and otherwise uses the caller's
finite positive default.

**Non-obvious choices.** Pixel conversion uses 914400 EMU per inch and
truncates toward zero. Invalid effective DPI or dimensions outside the `i64`
range return `None`.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, and
`docs/hld/14-development-backlog.md` now state the dependency-free return type.

**Tests.** Declared-DPI precedence, independent per-axis fallback, fractional
truncation, invalid input, dependency and package riders, 22 integrated media
tests, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** The later consumer cutover must supply its own
default DPI and convert no units outside this API.

### F-029, Create oxml-layout

**Sprint.** S06
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The unpublished `oxml-layout` crate now stages
format-neutral layout output, font management, bundled deterministic fonts,
and layout errors without changing a released consumer.

**Non-obvious choices.** Bundled fonts are always available for deterministic
rendering. The default `system-fonts` feature controls only host font discovery,
and the no-default path keeps the same bundled archive.

**Deviations from the design plan.** Sprint review pass 1 found that two HLD
passages still described an older bundled-fonts-off no-default path. The
implementation followed the intended boundary, and the stale current-intent
wording was corrected during sprint remediation.

**Spec sections touched.** `docs/hld/12-testing-strategy.md` and
`docs/hld/15-build-and-toolchain.md` now state that the no-default path disables
system discovery while retaining bundled deterministic fonts.

**Tests.** Default and no-default font-manager tests, empty-font error handling,
bundled-font licence coverage, dependency isolation, archive contents and size,
released-crate isolation, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep deterministic fonts inside `oxml-layout`
when consumers migrate. System discovery remains an optional capability, not a
determinism requirement.

### F-030, Decouple line.rs

**Sprint.** S06
**Completed.** 2026-07-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `oxml-layout` now has a greedy line breaker with owned
layout types for alignment, tabs, leaders, underline, and spacing. Explicit
wrapping control can retain width overflow while forced line, page, and column
breaks continue to split content.

**Non-obvious choices.** Tab positions and exact or minimum spacing are stored
in points. Multiple spacing stores a factor, and the staged boundary contains
neither twips nor stringly line rules.

**Deviations from the design plan.** Microscope pass 1 strengthened the copied
tests to use deterministic bundled fonts and made the leader regression prove
real glyph shaping. Production behavior remained on the approved design.

**Spec sections touched.** None. The existing migration and rendering contracts
already define the owned line-breaking boundary and deferred rdocx converter.

**Tests.** All 11 copied compatibility tests, four owned spacing and wrapping
regressions, deterministic leader shaping, both feature modes, dependency and
package riders, released-line-breaker isolation, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** The later rdocx cutover must perform schema-unit
conversion before constructing these types. Do not move Word-specific enums
back across this boundary.

### F-031, Transform

**Sprint.** S06
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-layout` now exports a concrete six-coefficient affine
transform with rotation about a point, self-first composition, point
application, exact identity checks, and four-corner rectangle bounds.

**Non-obvious choices.** `self.then(next)` applies `self` first and `next`
second, matching the point equations and PDF `cm` concatenation order. Identity
comparison is exact so small intentional transforms are never discarded.

**Deviations from the design plan.** Microscope pass 1 strengthened the
composition gate with fully nonzero matrices and replaced an exact quarter-turn
bounds case with a negative 30-degree rotation. Production algebra was already
correct.

**Spec sections touched.** None. The existing rendering and testing contracts
already specify the matrix representation and composition order.

**Tests.** Identity neutrality, fractional and positive rotation, fully nonzero
hand-computed PDF composition, negative-rotation four-corner bounds, exact
identity, package and dependency riders, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Convert DrawingML rotation units to degrees
before this boundary, and preserve self-first composition when group transforms
begin consuming the type.

### F-032, Path and PathCommand

**Sprint.** S07
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-layout` now exports backend-neutral path commands,
fill rules, conservative bounds, and constructors for rectangles, rounded
rectangles, and four-cubic ellipses.

**Non-obvious choices.** Bounds include cubic control points without solving
curve extrema. Rounded rectangles use one circular radius clamped from zero to
half the shorter side.

**Deviations from the design plan.** None.

**Spec sections touched.** None. The existing rendering contract already
defines the path representation and conservative bounds.

**Tests.** Ellipse bounds within the control hull, cubic control-point bounds,
empty paths, fill-rule independence, closure, and radius clamping.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Backend path conversion must preserve command
order and choose the fill operator from `FillRule`.

### F-033, Paint and Stroke

**Sprint.** S07
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-layout` now exports solid, linear, radial, and tiled
paint plus gradient stops, line caps, line joins, and arbitrary dash arrays for
strokes.

**Non-obvious choices.** Only a one-stop gradient is normalized at
construction, becoming solid paint. Empty and multi-stop gradients remain
unchanged for the later backend normalization stage.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/14-development-backlog.md` now records
F-036 as an explicit dependency because tiled paint stores `MediaId`.

**Tests.** Single-stop linear and radial degradation, multi-stop preservation,
stroke defaults, tile media identity, package, dependency, and feature-mode
checks.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Sorting, clamping, and duplicate-stop precedence
remain backend construction work and must not be moved silently into this
model.

### F-034, Path and Group arms

**Sprint.** S07
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The staged positioned-element model now carries painted
paths and nested groups with transforms, clips, opacity, effects, and children.
Page backgrounds, layout diagnostics, and constructors for the two
non-exhaustive result structs are also available.

**Non-obvious choices.** `Diagnostic` initially carries one message.
`PositionedElement` and `Effect` are non-exhaustive enums, while `PageFrame`
and `LayoutResult` are non-exhaustive structs with neutral constructor defaults.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` and
`docs/hld/14-development-backlog.md` now state the exact non-exhaustive targets,
minimal diagnostic shape, constructor contract, and unpublished staging
boundary.

**Tests.** Path and group payload preservation, transform direction, neutral
page and result defaults, external constructor doctests, and the integrated
full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Group transforms map child-local coordinates
into their parent. Preserve that direction in every recursive backend.

### F-035, The walk helper

**Sprint.** S07
**Completed.** 2026-07-31
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `oxml-layout::walk` visits every non-group element once in
depth-first document order while carrying its accumulated child-to-page
transform.

**Non-obvious choices.** Group containers are not yielded. With self-first
composition, each group transform is composed before the accumulated parent
transform.

**Deviations from the design plan.** None.

**Spec sections touched.** None. The existing recursion-hazard contract already
defines the helper and its consumers.

**Tests.** Three-deep traversal with three ordered leaves, hand-computed nested
transform order, group exclusion, and root identity.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Font, image, and link collection passes must use
this helper when the PDF backend migrates.

### F-036, MediaId

**Sprint.** S07
**Completed.** 2026-07-31
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Staged output and line image values now use a stable
content-addressed `MediaId` derived from raw bytes instead of relationship-local
embed identifiers.

**Non-obvious choices.** The compact handle uses fixed 64-bit FNV-1a and is
documented as a renderer key rather than a collision-free content guarantee.
This story adds no media store.

**Deviations from the design plan.** None.

**Spec sections touched.** None. The existing rendering contract already
defines the content-addressed handle and replacement boundary.

**Tests.** Equal bytes collapse to one set key, different fixtures differ,
staged output images use the handle, and line conversion preserves it.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Relationship resolution must happen before
constructing this renderer key. Part-local relationship names must not cross
the shared layout boundary.

### F-037, Create oxml-pdf

**Sprint.** S08
**Completed.** 2026-07-31
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The workspace now contains an unpublished `oxml-pdf`
backend at version 0.0.0. It consumes `oxml-layout` and `oxml-media`, keeps the
eight moved backend gates, and removes the copied image-header parsers.

**Non-obvious choices.** The staged crate has no dependency on an `rdocx-*` or
`rpptx-*` crate. It remains excluded from publication while the PowerPoint
development line is incomplete.

**Deviations from the design plan.** A normal staged package archive was built,
but extracted verification resolves the unpublished 0.0.0 dependencies from
crates.io placeholder packages. The reviewed package rider therefore used
`cargo package --no-verify`. The 19.6 KiB archive stayed below the 10 MiB gate,
and no package was published.

**Spec sections touched.** `docs/hld/03-architecture.md` and
`docs/hld/08-rendering-spec.md` now record the staged backend boundary and its
dependency direction.

**Tests.** Fifteen staged backend tests, including the eight moved gates,
dependency-tree inspection, archive inspection and size, and the integrated
full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep `oxml-pdf` unpublished and independent of
released format crates until the shared publication and cutover sprint.

### F-038, Golden-PNG harness

**Sprint.** S08
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Document::to_pdf_deterministic` exposes the bundled-font
PDF path, and the golden harness rasterises page one of all seven samples at
150 DPI before comparing decoded RGBA dimensions and SHA-256 digests exactly.

**Non-obvious choices.** The manifest records `pdftoppm version 26.01.0` and
contains digests rather than binary fixtures. Comparison has no tolerance.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/12-testing-strategy.md` and
`docs/hld/15-build-and-toolchain.md` now define the deterministic facade,
rasterizer identity, manifest contents, and exact pixel gate.

**Tests.** Four harness unit tests, seven exact sample comparisons, a synthetic
one-pixel `proposal` failure that names only that sample, facade coverage,
package checks, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Update the golden manifest only for a declared,
reviewed rendering change with a non-empty reason. Continue comparing decoded
pixels exactly.

### F-039, Global CTM flip

**Sprint.** S08
**Completed.** 2026-07-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Both PDF writers now emit one page-level
`q 1 0 0 -1 0 H cm`. Text and images cancel that outer reflection locally,
while lines and rectangles use top-left coordinates and link annotations stay
outside the content transform.

**Non-obvious choices.** The correct image operator is
`[w 0 0 -h x y+h]`. Omitting the added height moves images outside the page
under the outer CTM.

**Deviations from the design plan.** The original operator used `y`, which was
corrected to `y+h` before integration. Poppler 26.01.0 then produced exactly
four one-pixel vertical antialias swaps at x 112, two in `invoice` and two in
`quote`. The user approved that exact delta, the manifest changed once with an
F-039 reason, and all seven buffers now compare exactly.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/13-risks-and-open-questions.md`,
and `docs/hld/14-development-backlog.md` now record the corrected operator,
mirrored released-backend exception, and exact rendering evidence.

**Tests.** Page CTM, top-left geometry, upright text and images, unchanged
annotation coordinates, exact pre-update four-pixel evidence, seven exact
post-update buffers, one-pixel injection rejection, and the integrated full
gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve the same page and image matrices in
both PDF writers until the F-046 cutover removes the released duplicate. Do not
introduce a pixel tolerance.

### F-040, Group rendering

**Sprint.** S09
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The staged PDF writer recursively emits nested groups with
balanced graphics-state saves and restores, local matrices, optional clipping,
shared opacity states, and children in document order.

**Non-obvious choices.** Group clips reuse F-041's private geometry emitter and
group opacity reuses F-044's document-wide alpha registry. Effects remain
staged, and the raster group path remains owned by F-045.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, and `docs/hld/14-development-backlog.md` now
describe recursive group emission and its exact ordering.

**Tests.** `three_deep_groups_balance_graphics_state`, transform ordering,
non-zero and even-odd clipping, shared group opacity, staged effects, the exact
seven-sample golden gate, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Content emission must preserve group boundaries.
Flattening with `walk` is correct for collection passes but would lose clip and
opacity scope here.

### F-041, Path rendering

**Sprint.** S09
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Staged PDF paths now emit move, line, cubic, and close
geometry plus solid fill and stroke state. Paint selection covers `f`, `f*`,
`S`, `B`, and `B*` with balanced graphics state.

**Non-obvious choices.** The private geometry emitter also serves group clips.
Gradient and tile components remain staged, while a supported solid component
still renders when the other paint component is not yet supported.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, and `docs/hld/14-development-backlog.md` now
state the supported geometry, stroke state, and paint operators.

**Tests.** Fill-only, stroke-only, combined, even-odd, command-order, cap, join,
miter, dash, staged-component, exact seven-sample golden, and integrated full
gates.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** F-043 owns gradient resources. Preserve the
single geometry emitter so visible paths and group clips cannot diverge.

### F-042, Rewrite the three collection passes on walk

**Sprint.** S09
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Font usage, image registration, and every link annotation
pass now traverse nested leaves through `walk`. Image resources, annotations,
and recursive emission share depth-first leaf ordinals, and grouped link
rectangles apply the accumulated transform before PDF page conversion.

**Non-obvious choices.** A private per-page emission state carries the resource
maps and leaf ordinal through recursion. It keeps identity aligned without
public model fields, pointer keys, or parallel traversal implementations.

**Deviations from the design plan.** None. Microscope pass 1 strengthened the
grouped-link test to prove the page `/Annots` reference as well as the
transformed annotation dictionary.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/13-risks-and-open-questions.md`,
and `docs/hld/14-development-backlog.md` now close the R3 mitigation with the
implemented traversal and identity contract.

**Tests.** Nested font subsetting, nested XObject registration and use, nested
transformed link annotations, depth-first leaf identity, top-level stability,
the exact seven-sample golden gate, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Any new resource collection pass must use the
same depth-first leaf contract as recursive content emission.

### F-044, ExtGState alpha

**Sprint.** S09
**Completed.** 2026-07-31
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The staged PDF writer allocates one document-wide
`/ExtGState` for each normalized non-opaque alpha value, writes matching `CA`
and `ca`, and exposes only the states each page uses. Text, lines, rectangles,
solid paths, and group opacity share the registry.

**Non-obvious choices.** Keys use the serialized `f32` value, finite inputs are
clamped, negative zero is normalized, and non-finite values remain opaque. A
path with different fill and stroke alpha repeats its geometry so each paint
operation can select the correct state.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/13-risks-and-open-questions.md`,
and `docs/hld/14-development-backlog.md` now describe shared PDF alpha and its
deterministic raster gate.

**Tests.** Equal-state reuse, distinct states, matching `CA` and `ca`, opaque
content, shared path alpha, differing fill and stroke alpha, exact midpoint
raster compositing, the exact seven-sample golden gate, and the integrated full
gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Reuse this private registry for any new PDF
primitive with alpha. Do not allocate graphics states per element.

### F-043, Gradient shading dictionaries

**Sprint.** S10
**Completed.** 2026-07-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The staged PDF writer now emits deterministic type 2
shading patterns, axial and radial shadings, and type 3 stitching functions
over interval type 2 functions. Page-local pattern resources and fill or stroke
operators address each gradient occurrence by stable depth-first identity.

**Non-obvious choices.** Pattern matrices compose the accumulated group
transform with the global page flip so paint and path geometry share top-left
coordinates. Stops are clamped and sorted, the last repeated offset wins, and
stop alpha uses the documented opaque DeviceRGB fallback over white.

**Deviations from the design plan.** Microscope pass 1 corrected the pattern
matrix to include the global page flip. Pass 2 strengthened the page-local
resource test so it fails if pattern names leak between pages.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, and `docs/hld/14-development-backlog.md` now
describe the implemented PDF gradient resource graph, normalization, matrices,
operators, and exact sampled gate.

**Tests.** Linear and radial resource structure, stop normalization, mixed
solid paint, gradient stroke operators, page-local resources, and exact rotated
gradient samples under Poppler 26.01.0. The integrated 57-test `oxml-pdf`
suite, seven-buffer golden gate, and full workspace gate also pass.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep PDF and raster gradient stop normalization
aligned. The integrated gate completed normally after a worker-only duplicate
hash rerun had stalled in the macOS loader without changing files or evidence.

### F-045, Rasteriser groups, paths, gradients, dashes, and background

**Sprint.** S10
**Completed.** 2026-07-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The staged tiny-skia backend now recursively renders group
transforms and intersecting clips, composites group opacity once per subtree,
draws backend-neutral paths with solid or gradient paint, honours line and path
dashes, and paints supported page backgrounds.

**Non-obvious choices.** Scoped opacity uses a scratch pixmap so overlapping
children are not attenuated individually. Non-extended gradient domains receive
an explicit mask, while tile paint and group effects remain outside the S10
contract.

**Deviations from the design plan.** Microscope pass 1 added explicit raster
stop normalization with clamping, stable sorting, and last-repeated-offset
semantics to match F-043. Pass 2 was clean.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, and `docs/hld/14-development-backlog.md` now
describe recursive raster state, paint translation, dashes, backgrounds, and
the deterministic sampled gates.

**Tests.** Twelve raster tests cover the rotated-rectangle and dashed-line
gates, nested transform order, clip intersection, group opacity, fill rules,
linear and radial gradients, gradient domains and normalization, path dashes,
and page backgrounds. The integrated 57-test `oxml-pdf` suite, exact golden
gate, deliberate one-pixel rejection, and full workspace gate also pass.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve recursive draw order and scoped state.
Flattening groups would lose clip and opacity semantics even though `walk` is
correct for collection passes.

### F-052, Create oxml-drawing and namespace constants

**Sprint.** S12
**Completed.** 2026-07-31
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The workspace now contains the unpublished `oxml-drawing`
crate with the DrawingML, relationships, and package namespace constants used
by later model stories.

**Non-obvious choices.** The crate remains at version 0.0.0 with publication
disabled. It started without dependencies so the format-neutral boundary was
established before parsers were added.

**Deviations from the design plan.** None.

**Spec sections touched.** No HLD file changed. The implementation follows the
existing crate boundary in `docs/hld/03-architecture.md` and development
publication policy in `docs/hld/15-build-and-toolchain.md`.

**Tests.** Namespace URI assertions, workspace membership and publication
state checks, package inspection, dependency inspection, and the integrated
full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep every `oxml-*` production edge
format-neutral. The final S12 graph contains no `rdocx-*` or `rpptx-*` edge.

### F-053, OrderedRawChildren

**Sprint.** S12
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `OrderedRawChildren` stores unmodelled XML subtrees at
caller-defined schema boundaries so modelled children can be written in schema
order without moving or dropping unknown siblings.

**Non-obvious choices.** The helper is concrete and schema-boundary based. It
does not own a generic parser policy or hide which child sequence the caller
implements.

**Deviations from the design plan.** None.

**Spec sections touched.** No HLD file changed. The helper implements the
existing child-order and verbatim-preservation contracts in
`docs/hld/05-drawingml-model.md`.

**Tests.** Raw children before, between, and after modelled children, multiple
children at one boundary, byte-for-byte nested subtree preservation, and the
integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Owning parsers still decide their schema
boundaries. Do not append unknown children at the end of a parent.

### F-054, Colour choices

**Sprint.** S12
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `ColorChoice` now models sRGB, scheme, system, and preset
DrawingML colours with validated RGB values, prefix-tolerant parsing,
fixed-prefix writing, and ordered raw-child preservation.

**Non-obvious choices.** System colours preserve `lastClr` as their portable
fallback. Unknown child subtrees remain byte-for-byte raw rather than becoming
partially modelled data.

**Deviations from the design plan.** Microscope pass 1 corrected the shared
`OxmlError` conversion so parser errors retain their source contract. Pass 2
was clean.

**Spec sections touched.** No HLD file changed. The four choices and
preservation behaviour were already specified in
`docs/hld/05-drawingml-model.md`.

**Tests.** All four colour forms parse and round-trip, malformed RGB and system
fallback values fail, unknown nested children retain their exact bytes, and
the integrated full gate passes.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Parse only children the shared renderer
consumes. Preserve every other subtree at its original boundary.

### F-055, The colour transform stack

**Sprint.** S12
**Completed.** 2026-07-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** All 28 DrawingML colour transforms parse, serialize, and
apply in document order. A readable 40-case table records exact RGBA sampled
from PowerPoint 16.104 build 16.104.25121423.

**Non-obvious choices.** The oracle starts from a PowerPoint-authored native
shape shell, validates the transformed deck without repair, and captures each
shape through the native clipboard PNG payload. This replaced the PowerPoint
`save as picture` command, which returned success without creating a file on
the pinned build.
Linear-light, HSL, alpha, and PNG quantization rules are kept explicit.

**Deviations from the design plan.** The approved plan was revised before
completion to record the clipboard transport. Microscope pass 1 found that
explicit empty start and end transform pairs were preserved raw instead of
modelled. The parser now models those pairs while preserving unexpected
nonempty transform content verbatim. Pass 2 was clean.

**Spec sections touched.** `docs/hld/05-drawingml-model.md`, "Colour, the part
everyone gets wrong", now lists all 28 transforms and the exact resolution and
partial-alpha boundary rules.

**Tests.** The 40 exact PowerPoint cases, all 28 XML mappings, transform order,
linear-gamma conversion, partial alpha, explicit empty pairs, unexpected
nested content, raw-child order, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Expected oracle values are evidence, not values
generated from the Rust formulas. Re-run the ignored generator only with the
pinned PowerPoint build and an explicit native shell.

### F-056, Colour map resolution

**Sprint.** S12
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A concrete 12-slot `ColorMap` provides the standard Office
mapping, selective layout or slide overrides, and resolution in map, theme,
then transform order. Direct RGB, system, and preset choices bypass the map.

**Non-obvious choices.** Semantic and theme slots are validated enums. The
resolver takes a concrete lookup slice instead of a trait or generic, and a
missing system name uses its `lastClr` fallback.

**Deviations from the design plan.** Microscope pass 1 strengthened all 12
default mappings, all 11 untouched override slots, exact direct-colour results,
and the system fallback path. Pass 2 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
three-stage resolution contract in `docs/hld/05-drawingml-model.md` and leaves
`p:clrMap` parsing to F-069.

**Tests.** Standard mapping, selective override composition, the exact dark
master inversion gate with transforms, direct-colour bypass and system
fallback, plus the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** PresentationML parsers should construct this
format-neutral value. They must not move `p:` parsing into `oxml-drawing`.

### F-057, a:xfrm

**Sprint.** S13
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-drawing` now models DrawingML transforms with shape
and child coordinate rectangles, rotation, flips, prefix-tolerant parsing,
fixed-prefix writing, ordered raw-child preservation, and finite affine
composition.

**Non-obvious choices.** Child coordinates map into the parent rectangle before
rotation and flips are composed. Zero child extents return a typed error rather
than producing non-finite matrix coefficients.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
transform contract in `docs/hld/05-drawingml-model.md`.

**Tests.** `nested_group_transform_composes_to_the_hand_computed_matrix`,
prefix and schema-order writing, raw-child preservation, zero-extent rejection,
and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep DrawingML coordinate conversion in this
format model. Renderer transforms remain backend-neutral values.

### F-058, Guide evaluator

**Sprint.** S13
**Completed.** 2026-07-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `oxml-drawing` now parses and evaluates all 17 DrawingML
guide formula tokens from an owned seeded environment, applies adjust-value
overrides in declaration order, evaluates local path commands, and lowers
clockwise arcs to finite cubic Beziers in segments no larger than 90 degrees.

**Non-obvious choices.** Office interoperability defines `mod` as a Euclidean
norm and applies `sqrt` to the absolute input. Multi-turn sweeps are valid and
bounded by a segment-count guard rather than rejected at one full turn.

**Deviations from the design plan.** The approved plan was corrected before
implementation to use the standard 17 formula tokens. Microscope pass 1 found
that valid multi-turn arcs were rejected. A 450-degree regression fixed the
defect, and pass 2 was clean.

**Spec sections touched.** `docs/hld/05-drawingml-model.md`, "Geometry", now
records the standard formula set, owned environment, Office deviations, and
arc-lowering semantics.

**Tests.** `hand_written_custom_geometry_guides_produce_expected_path_coordinates`,
all formula operations, Office deviations, invalid math, finite arc endpoints,
multi-turn sweeps, unit-angle conversion, truncation riders, and the integrated
full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep formula operands owned because custom
geometry supplies document data. Reject every non-finite intermediate before a
renderer sees it.

### F-059, a:custGeom

**Sprint.** S13
**Completed.** 2026-07-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-drawing` now parses, writes, and evaluates
`a:custGeom` adjust lists, guide lists, text rectangles, path lists, and path
commands. Reads tolerate arbitrary prefixes, writes use the fixed `a:` prefix,
and unknown subtrees remain byte-for-byte at their schema boundaries.

**Non-obvious choices.** The approved schema-valid custom geometry fixture is
inline because the repository has no fetched deck corpus. The real corpus gate
remains at the M7 boundary.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
geometry and preservation contracts in `docs/hld/05-drawingml-model.md`.

**Tests.** `corpus_custom_geometry_round_trips_and_evaluates_to_a_closed_path`,
prefix and child-order writing, raw-child preservation, malformed-input
handling, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Replace or supplement the inline fixture when
the separately fetched M7 deck corpus is available. Do not commit a binary deck
only to duplicate the same XML boundary test.

### F-060, Fills

**Sprint.** S13
**Completed.** 2026-07-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `oxml-drawing` now models no fill, solid fill, linear and
path gradients, pattern fill, and stretched or tiled picture fill with source
rectangles. Gradient stops retain document order, reads tolerate arbitrary
prefixes, and writers emit fixed-prefix schema order.

**Non-obvious choices.** Picture fills retain relationship identifiers as
owned strings without introducing an OPC dependency. Modelled leaf elements
also retain ordered raw children so nested extensions survive round trips.

**Deviations from the design plan.** Microscope pass 1 found that a pattern
colour wrapper containing only an extension was dropped. Pass 2 found the same
class of loss in modelled leaf elements. Both were fixed with focused
regressions, and pass 3 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
fill module and preservation contracts in `docs/hld/05-drawingml-model.md`.

**Tests.** `every_fill_form_round_trips_and_gradient_stops_keep_document_order`,
prefix and schema-order writing, nested raw preservation, malformed-value
rejection, released Word theme isolation, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Resolve picture relationships and media bytes in
the PresentationML package layer. Do not add an OPC dependency to this crate.

### F-061, Lines

**Sprint.** S14
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-drawing` now models `a:ln` width, fill, preset and
custom dashes, cap and join choices, and head and tail ends. Reads accept any
prefix, writes use fixed `a:` prefixes, and unsupported children retain their
schema boundaries.

**Non-obvious choices.** Every preset dash token maps to a concrete dash array
without a fallback. The wire model retains schema values while the mapping
provides the later renderer boundary.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
line and preservation contracts in `docs/hld/05-drawingml-model.md`.

**Tests.** `every_preset_line_dash_value_maps_to_a_dash_array`, complete line
round trips, schema-order writing, raw-child preservation, malformed-value
handling, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep dash interpretation in the renderer seam.
The XML layer should continue to preserve the declared line vocabulary.

### F-062, Effects

**Sprint.** S14
**Completed.** 2026-08-01
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `oxml-drawing` now models effect lists and outer shadows,
including geometry and colour, while retaining unsupported effects such as
glow as raw XML at their exact positions.

**Non-obvious choices.** The effect list models only the values needed by the
current renderer contract. Unsupported effect subtrees remain authoritative
wire data rather than partial models.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
effect and preservation contracts in `docs/hld/05-drawingml-model.md`.

**Tests.** `a_shape_with_glow_round_trips_with_glow_intact_as_raw_xml`, outer
shadow round trips, schema-order output, malformed-value handling, and the
integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Add another typed effect only when a current
consumer needs it. Raw preservation already protects unsupported effects.

### F-063, Shape properties and style references

**Sprint.** S14
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-drawing` now composes transforms, geometry, fills,
lines, and effects through `a:spPr`, and models line, fill, effect, and font
style references with optional colour choices.

**Non-obvious choices.** Style index zero is valid. Fill indices greater than
1000 select the background-fill list with an offset of 1000, so index 1001
resolves to background fill style 1.

**Deviations from the design plan.** Microscope pass 1 found that zero indices
and colourless style references were rejected. Both schema-valid cases were
added, and pass 2 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
shape composition and style-matrix rules in `docs/hld/05-drawingml-model.md`.

**Tests.** `fill_ref_1001_resolves_to_background_fill_style_1`, all four style
reference forms, zero and colourless references, shape schema order,
malformed-input handling, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** The format-neutral reference model reuses
`ColorChoice`. It does not alter the released Word theme path.

### F-064a, Text body properties and shell

**Sprint.** S14
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-drawing` now owns a typed text-body shell and body
properties for insets, anchoring, wrapping, vertical direction, autofit, and
schema-ordered raw children.

**Non-obvious choices.** The first child kept later list styles and paragraphs
opaque until their dependent stories arrived. This preserved a usable shell
without anticipating their models.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
text-body and autofit contracts in `docs/hld/05-drawingml-model.md`.

**Tests.** `every_body_property_autofit_form_round_trips_in_schema_order`,
prefix handling, raw-child boundaries, malformed attributes, and the
integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Autofit remains a wire-model choice. Layout
policy belongs to the later renderer milestone.

### F-064b, Text paragraphs and runs

**Sprint.** S14
**Completed.** 2026-08-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `oxml-drawing` now models paragraphs, paragraph and run
properties, regular runs, fields, line breaks, text spans, fonts, hyperlinks,
spacing values, and ordered unsupported XML.

**Non-obvious choices.** Significant leading or trailing text is emitted with
`xml:space="preserve"`. Qualified relationship and whitespace attributes are
matched exactly so hostile prefixes cannot masquerade as `r:id` or
`xml:space`.

**Deviations from the design plan.** Microscope pass 1 found that local-name
fallback could interpret hostile qualified attributes as modelled attributes.
Exact matching and regressions fixed the issue, and pass 2 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
text and preservation contracts in `docs/hld/05-drawingml-model.md`.

**Tests.** `leading_and_trailing_text_whitespace_survives_via_xml_space_preserve`,
paragraph content order, property units, hyperlink attributes, hostile
prefixes, malformed input, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Relationship resolution remains outside this
crate. The parser stores relationship identifiers without adding an OPC edge.

### F-064c, Text bullets

**Sprint.** S14
**Completed.** 2026-08-01
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Paragraph properties now carry character, automatic, and
explicit no-bullet choices with optional font, point or percentage size, and
colour components in DrawingML schema order.

**Non-obvious choices.** All 41 automatic-numbering tokens are explicit. The
wire model retains literal bullet characters, while Wingdings conversion stays
with the later renderer work.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
bullet and preservation contracts in `docs/hld/05-drawingml-model.md`.

**Tests.** `every_modelled_bullet_form_round_trips_in_schema_order`, every
numbering token, optional component order, raw preservation, malformed values,
and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep font-symbol conversion out of structural
round trips so original bullet codepoints remain recoverable.

### F-064d, Nine-level list styles

**Sprint.** S14
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `a:lstStyle` now has nine explicit optional paragraph
property slots. The text-body model composes typed body properties, list
styles, and paragraphs into one structural round trip.

**Non-obvious choices.** Fixed slots make invalid list-level numbers
unrepresentable. Unsupported list-style siblings remain raw at their captured
boundaries, and modelled levels write in ascending schema order.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
nine-level text style chain in `docs/hld/07-inheritance-and-resolution.md`.

**Tests.** `schema_valid_text_body_using_all_nine_list_levels_round_trips_structurally`,
ascending level order, raw sibling preservation, invalid levels, and the
integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** The inline schema-valid fixture covers this
sprint. The fetched deck corpus remains required at the M7 boundary.

### F-064, DrawingML text model

**Sprint.** S14
**Completed.** 2026-08-01
**Size.** XL, estimated 0 days after split, actual 1 day

**What was built.** The umbrella closed after F-064a through F-064d delivered
the complete staged DrawingML text hierarchy for body properties, list styles,
paragraphs, runs, fields, breaks, whitespace, and bullets.

**Non-obvious choices.** The XL story carries no duplicate implementation.
Its four child stories own the natural schema boundaries and their individual
evidence, while this entry records the integrated contract.

**Deviations from the design plan.** The approved split used inline
schema-valid XML for S14. The fetched external deck corpus remains the M7
boundary gate, as planned.

**Spec sections touched.** `docs/hld/14-development-backlog.md` already records
the four-child split and the parent closure rule. No further prose change was
required.

**Tests.** The complete nine-level text-body structural round trip,
`leading_and_trailing_text_whitespace_survives_via_xml_space_preserve`, every
child test gate, and the integrated full workspace gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Run the separately fetched PowerPoint deck
corpus before closing M7. Do not publish the PowerPoint development crates
before PowerPoint development is complete.

### F-065, Theme read and write

**Sprint.** S15
**Completed.** 2026-08-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `oxml-drawing` now reads and writes complete DrawingML
themes with twelve colour slots, major and minor font collections, supplemental
script fonts, and typed fill, line, effect, and background-fill style lists.
`office_default()` supplies the standard Aptos-era Office theme.

**Non-obvious choices.** Unsupported attributes and children remain raw at
their schema boundaries. The canonical default is pinned to PowerPoint 16.104,
plist build 16.104.25121423, and AppleScript build 1214. The generated theme
opened in that build without repair.

**Deviations from the design plan.** Microscope pass 1 found that modelled XML
attributes were not decoded before canonical rewriting and that private writer
helpers had unjustified single-use generics. Entity decoding, its regression,
and concrete writer helpers fixed both findings. Pass 2 was clean.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, "The deck
corpus", and `docs/hld/14-development-backlog.md`, the M7 gate and F-065 and
F-067 entries. The external corpus gate now runs at S16 entry after F-067
creates its harness.

**Tests.** `powerpoint_office_theme_round_trips_structurally`,
`office_default_theme_is_accepted_by_powerpoint`, schema-order and fixed-prefix
writing, four format-style lists, raw-child preservation, entity decoding,
malformed input, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** F-067 must execute the carried M7 `a:txBody` and
`a:spPr` corpus gate before M8 model work. Keep development crates unpublished.

### F-066, The rdocx Theme adapter

**Sprint.** S15
**Completed.** 2026-08-01
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `oxml-drawing` now implements
`From<&CT_OfficeStyleSheet>` for the stable `rdocx_oxml::theme::Theme`,
projecting twelve concrete colour slots and the two Latin font families.

**Non-obvious choices.** The dependency runs only from unpublished
`oxml-drawing` to released `rdocx-oxml`. The adapter ignores shared fields the
legacy type cannot represent, prefers a system colour's `lastClr`, and retains
the legacy symbolic fallback when no resolved value exists.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
single dependency exception in `docs/hld/03-architecture.md` and the frozen
Word tint and shade contract in `docs/hld/05-drawingml-model.md`.

**Tests.** `shared_theme_adapter_matches_the_legacy_theme_projection`,
`shared_theme_adapter_does_not_project_unresolved_colour_forms`, the legacy
`tint_shade_modifiers` regression, both dependency trees, the released
`rdocx-oxml` package dry-run, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Do not reverse the dependency or install the
shared parser into the active Word path before the separately reviewed shared
crate publication and cutover work.

### F-067, Create rpptx-oxml and the corpus harness

**Sprint.** S16
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The workspace now contains unpublished `rpptx-oxml` at
version 0.0.0, PresentationML namespace constants, a pinned 50-deck public
corpus manifest and fetcher, opaque OPC round trips, and the carried M7
DrawingML corpus gate.

**Non-obvious choices.** Byte identity means equality of every decompressed
package part after canonical OPC save. ZIP metadata and compression are not
model state. The corpus remains in ignored storage and every fetch is checked
against its pinned SHA-256 value.

**Deviations from the design plan.** The corpus exposed boundary whitespace,
empty hyperlink relationship ids, and an empty custom-geometry path list.
Their canonical parser states and focused regressions were added before the
carried M7 gate passed.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/13-risks-and-open-questions.md` now record the implemented crate,
corpus source, gate, and publication boundary.

**Tests.** `corpus_manifest_is_complete_and_verified`,
`all_corpus_decks_round_trip_opaquely`,
`carried_m7_drawingml_gate_passes_for_the_corpus`, all 50 pinned decks, 6,898
text bodies, 8,643 shape-property elements, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep every PowerPoint development crate at
version 0.0.0 with publication disabled until PowerPoint development is
complete. Do not commit the fetched corpus binaries.

### F-068, presentation.xml

**Sprint.** S16
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rpptx-oxml` now models the presentation root, slide and
notes sizes, ordered slide and master identifiers, and the default text style.
It validates slide-id bounds and uniqueness while preserving unsupported XML
at its schema boundaries.

**Non-obvious choices.** Relationship identifiers remain strings for the OPC
layer to resolve. Reads use namespace URIs and tolerate alternate prefixes,
while writes use fixed PresentationML, DrawingML, and relationship prefixes.
Canonical-prefix collisions are rejected rather than changing the meaning of
preserved raw XML.

**Deviations from the design plan.** Microscope passes found local-name-only
matching, qualified id ambiguity, and nested canonical-prefix rebinding. URI
aware element matching, qualified relationship attributes, collision checks,
and focused regressions fixed them. Pass 3 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
presentation and preservation contracts in
`docs/hld/06-presentationml-model.md`.

**Tests.** `every_corpus_presentation_part_round_trips_structurally`, all 50
presentation roots, slide-id validation, alternate-prefix writing,
zero-slide templates, malformed input, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep relationship target resolution out of the
XML model and continue rejecting namespace rebinding that would corrupt raw
payload semantics.

### F-069, Slide, layout and master parts

**Sprint.** S16
**Completed.** 2026-08-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `rpptx-oxml` now has distinct schema-ordered models for
slides, layouts, masters, common slide data, master text styles, colour maps,
and colour-map overrides. Unsupported timing, transition, extension, and
producer-specific XML remains at ordered raw boundaries.

**Non-obvious choices.** The three roots use concrete types because their
schema sequences differ. Colour maps reuse the DrawingML value model but keep
their PresentationML element ownership. OPC relationship cardinality remains
an integration concern outside the XML structs.

**Deviations from the design plan.** None. Microscope pass 1 was clean. F-070
subsequently replaced the deliberately raw shape-tree boundary with its typed
model.

**Spec sections touched.** No HLD file changed. The implementation follows the
part, colour-map, and preservation contracts in
`docs/hld/06-presentationml-model.md`.

**Tests.** `every_corpus_slide_layout_and_master_round_trips_structurally`,
`corpus_part_relationship_counts_are_valid`, schema-order and colour-map
fixtures, 421 slides, 766 layouts, 76 masters, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep required package relationships in the OPC
layer and preserve unsupported root children in their captured schema slots.

### F-070, The shape tree

**Sprint.** S16
**Completed.** 2026-08-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Common slide data now owns a typed, schema-ordered shape
tree with required non-visual group properties, DrawingML group properties,
and recursive group shapes. All six child variants retain document z-order.

**Non-obvious choices.** Only group shapes recurse. Shapes, pictures, graphic
frames, connectors, and alternate content own their captured XML bytes until
their named later stories model them. Group properties expose the existing
DrawingML transform while preserving unsupported children.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
shape-tree and preservation contracts in
`docs/hld/06-presentationml-model.md`.

**Tests.** `nested_group_shape_tree_round_trips_with_tree_shape_preserved`,
`every_corpus_shape_tree_round_trips_structurally`, all six child variants,
required child order, 1,263 trees, 63 recursive groups, and the integrated full
gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** F-071 through F-074 own the opaque payload
variants. Preserve their current XML until each later story replaces one
boundary with an approved typed model.

### F-071, Placeholders

**Sprint.** S17
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rpptx-oxml` now models placeholder-bearing partial shapes
inside the ordered shape tree. Placeholder keys retain optional indices,
default an absent type to body, and implement index priority plus the title and
body equivalence classes.

**Non-obvious choices.** An optional `u32` preserves the distinction between a
missing index and index zero. Matching compares indices only when both sides
provide one, then falls back to effective placeholder types. Unrelated shape
content remains in ordered raw slots.

**Deviations from the design plan.** Microscope passes found local-name-only
matching, qualified identifier ambiguity, and nested canonical-prefix
rebinding. URI-aware element matching, qualified relationship attributes,
prefix-collision checks, and focused regressions fixed them. Pass 4 was clean.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` now records
the presence-sensitive placeholder key used by the matching contract.

**Tests.** Index and type matching, absent-type defaulting, both equivalence
classes, opaque preservation, nested group shapes, all 50 corpus decks, and
the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep placeholder matching presence-sensitive
and preserve unsupported shape content at its schema boundary.

### F-072, Pictures

**Sprint.** S17
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rpptx-oxml` now models pictures in root and recursive
shape trees, including non-visual properties, optional placeholders, embedded
and linked image relationships, source-rectangle crops, shape properties,
style, and extensions.

**Non-obvious choices.** Existing DrawingML blip-fill and shape-property types
gained concrete root-aware writers so picture-owned element names remain
schema-correct without adding forwarding wrappers. Relationship identifiers
stay strings for the OPC layer to resolve.

**Deviations from the design plan.** None. Microscope pass 2 was clean after
the first pass findings were remediated.

**Spec sections touched.** No HLD file changed. The implementation follows the
picture, prefix, ordering, and preservation contracts in
`docs/hld/05-drawingml-model.md` and `docs/hld/06-presentationml-model.md`.

**Tests.** Cropped picture round-trip, qualified relationships, alternate read
prefixes, fixed write prefixes, required child order, opaque alternate-content
preservation, 240 corpus pictures, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Resolve image targets in the OPC layer and keep
unsupported blip choices verbatim.

### F-073, Graphic frames

**Sprint.** S17
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rpptx-oxml` now models graphic frames and dispatches exact
graphic-data URIs to typed tables or opaque chart, SmartArt, OLE, and unknown
payloads. Root and recursive shape-tree arms use the typed frame model.

**Non-obvious choices.** Only the table branch is parsed because F-074 owns its
model. Every other payload remains opaque until its named story. The existing
DrawingML transform gained a concrete `p:xfrm` writer while retaining its
DrawingML path.

**Deviations from the design plan.** None. Microscope pass 2 was clean after
the first pass findings were remediated.

**Spec sections touched.** No HLD file changed. The implementation follows the
graphic-frame dispatch and preservation contracts in
`docs/hld/06-presentationml-model.md`.

**Tests.** Exact URI dispatch, required child order, fixed prefixes,
root-aware transforms, opaque payload preservation, all 86 corpus frames with
all four required kinds observed, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Add typed payload branches only in their owning
stories and retain unknown graphic data verbatim.

### F-074, DrawingML tables

**Sprint.** S17
**Completed.** 2026-08-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `oxml-drawing` now models table properties, style and
banding flags, grid columns, rows, cells, text bodies, merge origins, spans,
and horizontal and vertical continuations. Unsupported table and cell content
remains at ordered raw boundaries.

**Non-obvious choices.** Ambiguous edits to preserved grid metadata return a
typed error instead of silently attaching metadata to the wrong column. Cell
text reuses the existing concrete text-body model.

**Deviations from the design plan.** Microscope passes found defects in grid
metadata edits and preservation boundaries. The model now rejects ambiguous
mutations and has focused regressions. Pass 5 was clean.

**Spec sections touched.** `docs/hld/05-drawingml-model.md` now records the
table model, merge semantics, and preservation contract.

**Tests.** Merged-cell origins and continuations, banding and style flags,
schema order, alternate prefixes, opaque preservation, all 26 corpus tables
with 724 cells, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep grid metadata aligned with column identity
and reject edits that cannot preserve that identity unambiguously.

### F-075, Connectors

**Sprint.** S18
**Completed.** 2026-08-01
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `rpptx-oxml` now models connectors in root and recursive
shape trees, including optional typed start and end connections, required
non-visual and shape properties, and ordered preservation of unsupported
content.

**Non-obvious choices.** Each endpoint keeps its required unqualified shape id
and connection-site index. Unsupported locks, style, extensions, attributes,
and children remain at their original schema boundaries.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` now records
the typed connector arm and its optional endpoint contract.

**Tests.** Endpoint round-trip, namespace aliases, fixed prefixes, required
order, qualified-attribute rejection, raw preservation, 85 corpus connectors
with 30 starts, 28 ends, and 6 nested connectors, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep routing and unsupported non-visual content
opaque until a named story owns those boundaries.

### F-076, mc:AlternateContent

**Sprint.** S18
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Shape trees now expose ordered typed members from the
immediate `mc:Fallback` branch while retaining the complete alternate-content
subtree as the only serialisation source.

**Non-obvious choices.** Choices are not evaluated, an absent fallback returns
no selection, and an empty fallback remains distinct from an absent fallback.
Namespace URI resolution identifies the MC fallback instead of its prefix.

**Deviations from the design plan.** None. Microscope pass 2 was clean after
the first pass corrected stale HLD wording.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` now records
fallback-only selection, opaque choices, absent-fallback behaviour, and
raw-only serialisation.

**Tests.** Fallback selection, no-fallback and duplicate-fallback cases,
namespace aliases, recursive order, exact raw preservation, all 21 corpus
alternate-content subtrees, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Add choice evaluation only with an explicit
capability policy. Never serialise this model from the selected fallback.

### F-077, Notes slides and notes master

**Sprint.** S18
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rpptx-oxml` now models notes slides and notes masters,
types shape text bodies with the existing DrawingML text model, and extracts
speaker notes from effective body placeholders only.

**Non-obvious choices.** Plain text retains run, field, explicit-break, and
paragraph order. Slide images, numbers, dates, footers, headers, and master
prompt text do not enter speaker-note output.

**Deviations from the design plan.** Design review against the authoritative
schema corrected `p:notesStyle` from required to optional. The model, writer,
HLD, and regression gate use the optional contract. Microscope pass 2 was
clean.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` now records
the notes-root sequences, relationship cardinalities, optional notes style,
and body-placeholder extraction contract.

**Tests.** Text extraction order, placeholder filtering, schema order,
optional notes style, raw preservation, relationship completeness, all 210
corpus notes slides and 24 notes masters with 72 nonempty bodies, and the
integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Treat the typed text body as the single source
of truth and keep notes-master prompt text outside speaker-note extraction.

### F-078, relmap rewrite_rel_ids

**Sprint.** S18
**Completed.** 2026-08-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rpptx-oxml` now rewrites mapped numeric relationship ids
inside preserved XML by replacing only eligible relationship-namespace
attribute value ranges.

**Non-obvious choices.** Namespace URI scope, including aliases and nested
shadowing, decides eligibility. Element spelling, declarations, attribute
order, quote choice, comments, processing instructions, and every untouched
byte remain unchanged.

**Deviations from the design plan.** None. Design review corrected one HLD
section citation before implementation, and microscope pass 1 was clean.

**Spec sections touched.** No HLD file changed. The implementation follows the
relationship-remapping contract in `docs/hld/06-presentationml-model.md`.

**Tests.** Embed, link, and diagram relationship rewriting, aliases and
shadowing, unmapped and nonnumeric values, exact surrounding-byte preservation,
malformed XML, empty-map identity across preserved payloads in all 50 corpus
decks, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Use this byte-splice helper before deep-copying
opaque payloads. Do not reconstruct preserved XML through an event writer.

### F-079, The rpptx read facade

**Sprint.** S19
**Completed.** 2026-08-02
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The new unpublished `rpptx` crate opens presentations from
paths or bytes, resolves ordered slides and notes through OPC relationships,
and exposes safe borrowed slide and recursive shape handles with text, table
text, and speaker-note access. It also saves facade-owned modelled parts through
the deterministic package writer.

**Non-obvious choices.** Immediate shape iteration preserves z-order. Groups
and selected alternate-content fallbacks expose children explicitly. Indexed
access returns `Option`, missing notes remain distinct from empty notes, and
the public facade does not expose schema-layer `CT_*` values.

**Deviations from the design plan.** Microscope pass 1 found package-root
relationship targets with dot segments were resolved incorrectly and found a
forwarding-only helper. Both were corrected, and pass 2 was clean.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` records the
read facade and relationship-resolved ownership model. `docs/hld/12-testing-strategy.md`
records the normalized python-pptx differential gate.

**Tests.** Eight facade integration tests cover ordered slides and notes, total
indexed access, all six shape kinds, recursive text, contextual graph errors,
deterministic reopen, package-root dot segments, and workspace publication
metadata. `dump_deck_matches_python_pptx_1_0_2_for_the_corpus` matched the
pinned oracle across all 50 decks. The integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep `rpptx` at version 0.0.0 with
`publish = false` until PowerPoint development is complete. Preserve the
normalized facade boundary instead of leaking lower-level model types.

### F-080, Modelled round-trip gate

**Sprint.** S19
**Completed.** 2026-08-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The 50-deck gate now parses, serialises, reparses, and
structurally compares all seven modelled PresentationML and theme roots. It
builds an exact expected package, checks canonical bytes for rewritten parts,
preserves original bytes for unmodelled parts, and verifies the facade read
surface after save and reopen.

**Non-obvious choices.** `CT_StyleMatrix.name` is optional because accepted
producer themes omit `a:fmtScheme/@name`. Canonical `a:blip` output declares
the relationship namespace locally whenever it writes `r:embed` or `r:link`,
so independently serialised fills remain namespace-valid.

**Deviations from the design plan.** The approved plan was extended when the
first corpus run exposed an absent format-scheme name and the first native
PowerPoint run exposed an undeclared relationship prefix. Both compatibility
repairs received focused regressions. Microscope pass 3 was clean.

**Spec sections touched.** `docs/hld/05-drawingml-model.md` records optional
format-scheme names and self-contained blip namespaces.
`docs/hld/06-presentationml-model.md` and
`docs/hld/12-testing-strategy.md` record the structural and exact-byte
round-trip boundary.

**Tests.** The required 50-deck corpus passed all 11 `rpptx` tests, including
seven-root structural equality, exact expected packages, facade reopen, and
the pinned python-pptx differential. Both compatibility regressions passed.
A Codex-operated native run opened and closed all 50 generated decks without
repair in PowerPoint 16.104 build 16.104.25121423. The output digest was
`19609644c12923fad63939656fc54681c667efa2e066fbd2a080bb717aa037fc`.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep lexical byte equality for unmodelled parts
and exact expected canonical bytes for rewritten roots. XML well-formedness is
not a substitute for the native PowerPoint acceptance gate.

### F-081, ResolveCtx skeleton and placeholder chain

**Sprint.** S20
**Completed.** 2026-08-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The new unpublished `rpptx-layout` crate owns a concrete
per-slide `ResolveCtx` with theme, colour-map, presentation text defaults,
master, layout, and slide inputs. It resolves an ordinary slide placeholder
through nested groups and selected fallback branches to its layout and master
counterparts.

**Non-obvious choices.** Matching uses the existing index-first and
type-fallback rule. The master hop uses the matched layout placeholder key, and
a missing layout match terminates the chain instead of skipping a level.

**Deviations from the design plan.** Microscope pass 1 found that the chain
method had the wrong visibility. It was corrected to the crate boundary, and
pass 2 was clean.

**Spec sections touched.** No HLD file changed. The existing architecture and
inheritance HLD already assigned this boundary to `rpptx-layout`.

**Tests.** Four focused tests cover the complete two-hop gate, layout-key
master matching, recursive group and fallback lookup, and missing or
non-placeholder shapes. The integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Build each context once per slide and preserve
the exact hierarchy order when adding later resolver consumers.

### F-082, Effective transform and body properties

**Sprint.** S20
**Completed.** 2026-08-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Ordinary `p:spPr` is typed as DrawingML shape properties.
The resolver returns the first slide, layout, or master transform and merges
body properties per field over exact OOXML inset, anchor, wrap, direction, and
autofit defaults.

**Non-obvious choices.** Transform results are owned clones so hierarchy
lifetimes do not leak through the API. Body values overlay master, layout, and
slide in that order, retaining unrelated inherited fields.

**Deviations from the design plan.** A canonical-prefix assertion was updated
when typed `p:spPr` began using the fixed `p:` root. The first corpus run also
exposed excessive debug-stack pressure from storing shape properties by value.
F-084 completed the reviewed boxed-storage remediation while preserving field
access semantics and proving the normal-stack corpus gate.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` records typed
ordinary-shape properties. `docs/hld/07-inheritance-and-resolution.md` records
owned transform precedence and the exact per-field body cascade.

**Tests.** Transform precedence, exact defaults, per-field body overlay,
ordinary-shape schema order and raw preservation, all 68 PresentationML parser
tests, and the normal-stack 50-deck structural corpus gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the exact integer EMU defaults and do not
replace the property-level body cascade with whole-value selection.

### F-083, The seven-step list style merge

**Sprint.** S20
**Completed.** 2026-08-02
**Size.** L, estimated 4 days, actual 1 day

**What was built.** DrawingML list styles now type `a:defPPr`, and
`rpptx-layout` resolves all nine levels through presentation defaults, the
selected master style, master and layout placeholders, shape list style,
paragraph properties, and run properties.

**Non-obvious choices.** The cache stores only the first four sources by
optional placeholder key. Shape formatting is applied after cloning the cached
prefix, so two shapes occupying one placeholder cannot leak direct formatting.
Bullet components and nested character properties merge independently.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/05-drawingml-model.md` records typed
default paragraph properties. `docs/hld/07-inheritance-and-resolution.md`
records level selection, merge granularity, and prefix-only cache semantics.

**Tests.** Eleven resolver tests cover the named seven-source gate, all nine
levels, `defPPr`, property retention, bullet components, atomic fill and font
slots, raw-action exclusion, and cache isolation. DrawingML parser and required
corpus structural tests also passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Apply shape-owned style after cache cloning and
keep paragraph level selection separate from inherited formatting.

### F-084, Format scheme reference resolution

**Sprint.** S20
**Completed.** 2026-08-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Ordinary shapes now type `p:style` in schema order. The
resolver selects one-based fill, line, and effect format entries, applies the
background-fill rule above 1000, substitutes modelled placeholder colours, and
layers explicit shape properties over theme values.

**Non-obvious choices.** Index zero means no referenced entry, while a positive
out-of-range value is an error. Opaque effect DAGs replace referenced effects
atomically and are rejected when they retain unresolved `phClr`. Font
references retain their `major`, `minor`, or `none` collection selector rather
than being treated as numeric indices.

**Deviations from the design plan.** The plan was revised after the integrated
F-082 parser overflowed on a real slide master at normal stack size. Boxing
shape properties and private shape style removed the large transient values
without requiring a stack-size workaround. Microscope pass 1 then found opaque
effect DAGs bypassed unresolved-colour checks. Three regressions fixed that
defect, and microscope pass 2 was clean. Sprint review pass 1 then found that
the raw scanners matched local names without resolving namespaces. Two focused
regressions made foreign producer extensions inert.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` records typed
shape properties and style with bounded storage. `docs/hld/07-inheritance-and-resolution.md`
records numeric format-list rules, font collection selection, placeholder
colour substitution, explicit overlays, and malformed-reference errors.

**Tests.** Thirty-five integrated resolver tests include the named fill gate,
background fills, zero and out-of-range indices, transform order, explicit
overlays, modelled and opaque effects, unresolved placeholder rejection, and
same-named foreign effect extensions. All 68 parser tests, the 40-case exact
colour table, and the normal-stack 50-deck corpus gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Never claim an opaque effect with `phClr` is
concrete. Type it, substitute it, or return a resolver error.

### F-085, Typeface resolution

**Sprint.** S20
**Completed.** 2026-08-02
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `ResolveCtx::resolve_typeface` maps major and minor Latin,
East Asian, and complex-script tokens to concrete theme faces and applies
supplemental per-script overrides.

**Non-obvious choices.** Script is an explicit optional input because the
current run model does not perform script segmentation. The first matching
supplemental entry wins in document order, missing matches fall back to the
token-specific base face, and ordinary or unknown typefaces pass through.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/07-inheritance-and-resolution.md` records
the explicit script input, all six tokens, override order, fallback, and
pass-through behavior.

**Tests.** Six focused tests cover the named minor Latin gate, all major and
minor aliases, supplemental overrides, missing-script fallback, pass-through,
and duplicate-script document order. The integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep Unicode script segmentation in the text
shaping layer and pass the resulting ISO 15924 tag into this resolver.

### F-086, Draw order and the flattener

**Sprint.** S21
**Completed.** 2026-08-02
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `ResolveCtx::flatten` now emits the effective background,
allowed master artwork, allowed layout artwork, and slide shape-tree leaves in
final source order. It walks nested groups and selected fallback content while
retaining each leaf's source. Slide and layout `showMasterSp`, ordinary
placeholder suppression, typed header-footer controls, and occupied latent
placeholders are applied before the renderer boundary.

**Non-obvious choices.** Background selection keeps the producing part and
effective master colour map. The layout visibility flag controls only the
master pass, while the slide flag controls only the layout pass. Ordinary
master and layout placeholders remain templates rather than drawable content.

**Deviations from the design plan.** Microscope pass 1 found that the tests did
not isolate the two visibility controls and did not prove all four background
fallback sources. Both coverage gaps were corrected, and pass 2 was clean.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` records typed
visibility and header-footer inputs. `docs/hld/07-inheritance-and-resolution.md`
records the borrowed flattened view, final draw order, and suppression policy.

**Tests.** Forty-one resolver tests, 68 PresentationML integration tests, all
50 pinned corpus decks, exact-colour checks, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep source order and suppression in the
flattener. Renderers must not reconstruct the slide, layout, and master passes.

### F-087, ResolvedSlide contract

**Sprint.** S21
**Completed.** 2026-08-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rpptx-layout` now exposes an owned `ResolvedSlide`
boundary with point geometry, accumulated group transforms, concrete paint,
lines, shadows, text, tables, unsupported categories, and diagnostics. The
contract contains no PresentationML or DrawingML model types and retains
visible bounds fallbacks when content cannot yet be represented exactly.

**Non-obvious choices.** Custom paths scale in their own declared coordinate
spaces. Unsupported gradient geometry, media relationships, charts, SmartArt,
OLE, and pending presets remain explicit instead of being approximated as
concrete output. Character and automatic-number bullets retain independently
inherited font, colour, size, and choice values.

**Deviations from the design plan.** The plan added a test-only `oxml-opc`
dependency so both named gates could traverse all 50 real decks. Microscope
pass 1 found six contract and corpus defects, and pass 2 found group-composition
and automatic-number bullet defects. All were remediated. Pass 3 was clean.
Sprint review pass 1 then found that the named corpus gate accepted contextual
resolver errors. The strict gate exposed 20 affected slides. Preset black and
white now resolve concretely, and invalid custom geometry retains a diagnosed
bounds fallback. All corpus slides now produce an owned contract.

**Spec sections touched.** `docs/hld/07-inheritance-and-resolution.md` freezes
the owned output and fallback boundary. `docs/hld/08-rendering-spec.md` records
the accumulated group transform supplied to renderers.

**Tests.** Fifty-seven resolver tests, two independent 50-deck gates, exact
colour checks, dependency-direction riders, publication dry-run, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Renderers consume only the owned contract.
Relationship-to-media resolution and pending geometry evaluation must fill the
documented gaps without exposing source-model types.

### F-088, Visual differential tests

**Sprint.** S21
**Completed.** 2026-08-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The existing `rpptx` integration binary now assembles full
resolver inputs through package relationships and emits normalized visual
records with ordered shape kinds, bounds, concrete shape and run paint, text,
unsupported categories, and diagnostics. It compares shared fields with pinned
python-pptx 1.0.2, proves exact cyan inheritance, prompt suppression, draw
order, and master artwork multiplicity, and records the one-time native
PowerPoint acceptance.

**Non-obvious choices.** Python supplies only mutually observable structure.
Rust separately asserts effective latent-placeholder visibility because
python-pptx exposes raw collections. Automated gates skip when the ignored
external corpus is absent and optional, but fail when the corpus is required.
The manual acceptance record remains available without the corpus files.

**Deviations from the design plan.** The first native review exposed inherited
date and slide-number fields that PowerPoint hid and an explicit slide footer
that the resolver dropped. The private flattener was repaired to require an
inherited header-footer container, preserve occupied slide latent content, and
match latent types across level-specific indices. Microscope pass 1 found two
evidence gaps, pass 2 was clean after remediation, and integrator review found
the clean-clone corpus issue. The compact repair received a clean pass 3.

**Spec sections touched.** `docs/hld/07-inheritance-and-resolution.md` records
source-sensitive latent visibility and the executable visual differential.
`docs/hld/12-testing-strategy.md` records the selected decks, pinned oracle,
external-corpus policy, and native acceptance evidence.

**Tests.** Sixteen `rpptx` tests, 57 resolver tests, all 50 pinned decks, the
40-case exact PowerPoint colour table, optional and required corpus modes, and
the integrated full gate passed. Microsoft PowerPoint 16.104 build
16.104.25121423 opened and exported all four selected originals without repair
or clipping. Native master artwork, backgrounds, exact cyan, prompt
suppression, and footer visibility matched the remediated evidence.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep native visibility policy in the flattener,
keep the Python comparison structural, and preserve explicit unsupported
diagnostics until modelled background paint and media resolution land.

### F-089, Resolve the preset geometry licensing question

**Sprint.** S22
**Completed.** 2026-08-02
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The renderer specification and risk record now identify
the official ECMA-376 fifth-edition Part 1 electronic addendum as the permitted
source for preset geometry definitions. They record the inner archive path,
187-definition count, exact SHA-256, Ecma software-policy basis, and retained
BSD three-clause notice requirement.

**Non-obvious choices.** The decision permits only the official Ecma data set.
The MPL-2.0 LibreOffice implementation table remains rejected, and derivation
from specification text remains the fallback if the official file or notice
cannot be reproduced exactly.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` records the chosen
source and generator input. `docs/hld/13-risks-and-open-questions.md` closes the
provenance question with its licensing evidence.

**Tests.** `preset_geometry_provenance_is_recorded`,
`libreoffice_preset_table_remains_rejected`, repository prose checks, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Retain the vendored Ecma notice with the source
XML and generated table. Do not substitute an implementation-owned table.

### F-090, Preset table generator

**Sprint.** S22
**Completed.** 2026-08-02
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `tools/gen-presets/` now vendors the permitted Ecma XML and
licence notice and generates a checked-in Rust lookup table offline. Check mode
verifies the exact source hash, byte-identical regeneration, all 187 direct
definitions, 186 unique preset names, and coverage of every preset used by the
50-deck corpus.

**Non-obvious choices.** The official source repeats `upDownArrow` twice with
byte-identical XML. The generator rejects conflicting duplicate names and
deduplicates only this identical pair, leaving a deterministic 186-key table.

**Deviations from the design plan.** Source inspection corrected the planned
187 unique names to 187 direct definitions and 186 unique names. Microscope
pass 1 also found that the corpus scan accepted foreign-namespace elements and
that duplicate comparison passed through XML reserialization. Both checks now
operate on namespace-qualified input and exact source bytes. Pass 2 was clean.

**Spec sections touched.** None. F-089 had already recorded the source and the
rendering specification already required the offline checked-in mechanism.

**Tests.** `generator_reproduces_checked_in_table`,
`generated_table_covers_every_corpus_preset`,
`source_has_187_direct_definitions`, `generated_lookup_has_known_and_unknown_cases`,
the 50-deck scan covering 2,141 uses and 26 corpus names, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Regenerate only from the pinned XML after its
hash check. A repeated name with different source bytes is a hard failure.

### F-091, Preset evaluation and fallback

**Sprint.** S22
**Completed.** 2026-08-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** DrawingML preset geometry now models the preset name,
adjustment guides, and ordered raw children while preserving unknown XML.
Known presets use the generated definitions and shared custom-geometry guide
engine to produce backend-neutral paths and text rectangles. Unknown presets
retain shape bounds and text and emit a diagnostic naming the preset.

**Non-obvious choices.** Preset definitions are parsed through the existing
custom-geometry path instead of gaining a second evaluator. Shape-level
adjustments override generated defaults, and custom geometry retains schema
choice precedence when both forms are encountered.

**Deviations from the design plan.** Microscope pass 1 found that the corpus
gate could pass without proving evaluation. The strengthened gate exposed the
standard `wd12` and `hd10` fractional guides, which are now seeded by the
existing evaluator. Pass 2 was clean.

**Spec sections touched.** None. The implementation follows the existing
DrawingML parsing, resolver output, and preset fallback contracts.

**Tests.** `preset_geometry_round_trips_with_unknown_children_verbatim`,
`rectangle_preset_evaluates_to_expected_bounds_and_text_rect`,
`preset_adjustments_override_generated_defaults`,
`unknown_preset_keeps_bounds_text_and_diagnostic`, the non-vacuous 50-deck
corpus gate across 921 preset inputs, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep preset and custom geometry on the shared
guide engine. Unknown names must remain visible and diagnosed.

### F-092, rpptx-render skeleton and RenderInput

**Sprint.** S22
**Completed.** 2026-08-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The workspace now contains an unpublished `rpptx-render`
crate. Its `RenderInput` consumes owned `ResolvedSlide` values, content-addressed
media, fonts, and metadata. Upstream `SlideBundle` assembly carries slide,
layout, master, parsed theme, notes, visibility, and three explicitly scoped
relationship maps.

**Non-obvious choices.** Relationship lookup always names slide, layout, or
master scope, so identical relationship IDs cannot alias. Media keys derive
from bytes and deduplicate shared content. Raw PresentationML stays upstream of
the rendering boundary.

**Deviations from the design plan.** The implementation added the direct
inward `oxml-drawing` dependency required by `SlideBundle`'s concrete
`CT_OfficeStyleSheet` field. The dependency-direction rider confirms that no
reverse `oxml-*` to `rpptx-*` edge was introduced. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` now separates raw
assembly through `SlideBundle` from renderer consumption through `RenderInput`.

**Tests.** `same_relationship_id_resolves_independently_in_all_three_scopes`,
`equal_media_bytes_deduplicate_to_one_media_entry`,
`missing_relationship_reports_scope_and_id`,
`render_input_contains_only_resolved_slides`,
`rpptx_render_dependency_direction_is_one_way`, the publication dry-run, and
the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep raw OOXML in `SlideBundle` assembly and
feed renderers only the frozen owned resolver contract. All PowerPoint crates
remain version 0.0.0 with publication disabled until development is complete.

### F-093, Shape geometry, fills and lines

**Sprint.** S23
**Completed.** 2026-08-03
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `rpptx-render` now lowers resolved slides into ordered page
frames and shape geometry into backend-neutral paths. Solid fills, gradients,
outlines, diagnostics, metadata, and page order cross the renderer boundary.
An otherwise unpainted bounds fallback receives a deterministic 1 point black
outline so unsupported geometry remains visible.

**Non-obvious choices.** Shape paths stay in local coordinates beneath one
group transform. This keeps geometry, paint, and later text on one transform
path and avoids rewriting gradient coordinates into page space.

**Deviations from the design plan.** The sampled-pixel gate exposed an existing
`oxml-pdf` raster defect that applied an accumulated group transform twice to
gradient coordinates. The backend now keeps shader coordinates local, with a
focused regression test. Microscope pass 1 was clean.

**Spec sections touched.** None. The implementation follows the existing
page-frame and path-lowering contract.

**Tests.** `solid_gradient_and_outlined_shapes_rasterise_at_sampled_pixels`,
`preset_and_custom_geometry_lower_to_ordered_paths`,
`bounds_fallback_emits_a_visible_black_outline`,
`layout_slide_rejects_an_out_of_range_index`,
`layout_presentation_preserves_page_order_and_diagnostics`,
`translated_group_gradient_uses_local_coordinates_exactly_once`, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep gradient coordinates local to the same
group as their paths. The visible fallback outline is part of the approved
renderer contract.

### F-094, Rotation, flips and groups

**Sprint.** S23
**Completed.** 2026-08-03
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Resolved shapes now compose local rotation, centre-based
horizontal and vertical flips, bounds translation, and the accumulated parent
group transform into one backend-neutral group transform.

**Non-obvious choices.** The exact order is child rotation, flips,
translation, then parent transform. Geometry, gradients, outlines, and later
content remain beneath the same group so every visual component shares the
same placement.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** None. The implementation follows the documented
DrawingML composition and shared group boundary.

**Tests.** `rotated_shape_corners_match_hand_computed_coordinates`,
`horizontal_and_vertical_flips_are_about_the_shape_centre`,
`nested_group_transform_applies_child_before_parent`,
`rotated_gradient_and_outline_share_the_shape_transform`, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Use the shared group transform for every new
shape-local element. Do not pre-transform individual path coordinates.

### F-095, Arrowheads

**Sprint.** S23
**Completed.** 2026-08-03
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The resolved shape contract now carries source-neutral
line-end kind, width, and length values. The renderer derives stable endpoint
tangents and lowers triangle, stealth, diamond, oval, and arrow ends into
closed filled paths using the resolved line paint.

**Non-obvious choices.** Missing dimensions use DrawingML medium defaults.
Small, medium, and large dimensions are 2, 3, and 5 times the stroke width.
Degenerate segments omit their decoration without producing invalid geometry.

**Deviations from the design plan.** Microscope pass 1 found that structural
path checks did not prove rendered output. A deterministic raster assertion
was added, and pass 2 was clean.

**Spec sections touched.** `docs/hld/07-inheritance-and-resolution.md` and
`docs/hld/08-rendering-spec.md` now include the approved neutral line-end
contract and filled-path lowering.

**Tests.** `line_end_resolution_keeps_kind_width_and_length`,
`triangular_tail_end_emits_an_extra_filled_path`,
`head_end_uses_the_reversed_start_tangent`,
`all_supported_line_end_kinds_produce_finite_geometry`,
`zero_length_segment_omits_arrowhead_without_panicking`, deterministic raster
evidence, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Arrowheads remain presentation-side geometry,
not a shared stroke or backend primitive.

### F-096, Pictures with crop and tile

**Sprint.** S23
**Completed.** 2026-08-03
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Slide, layout, and master relationship identifiers now map
to content-addressed media in separate source scopes. Resolved picture content
carries neutral stretch or tile placement, and the renderer lowers crop,
alignment, translation, scale, flip, DPI, and rotation policy into clipped
shared image elements with bounded row-major tiling.

**Non-obvious choices.** Missing tile values normalize to zero translation,
100 percent scale, no flip, top-left alignment, and rotation with the shape.
Declared DPI wins over embedded DPI, which wins over 96 DPI. External linked
media remains unsupported without network access.

**Deviations from the design plan.** Microscope pass 1 found that
`rotateWithShape=false` did not cover every stretch and tile branch. The fix
made coverage, clipping, flip phase, and rotation policy explicit. Pass 2 was
clean.

**Spec sections touched.** `docs/hld/07-inheritance-and-resolution.md` and
`docs/hld/08-rendering-spec.md` now describe source-scoped media and neutral
picture placement.

**Tests.** `cropped_picture_renders_only_its_crop_region`,
`same_relationship_id_resolves_to_distinct_media_in_each_source_scope`,
`picture_model_resolves_to_neutral_stretch_and_tile_placement`,
`crop_lowers_to_clipped_source_image_geometry`,
`tile_picture_repeats_media_in_row_major_order_inside_shape_clip`,
`tile_dpi_prefers_declared_then_embedded_then_96`,
`equal_picture_bytes_reuse_one_media_id_across_elements`,
`missing_external_media_and_empty_crop_are_contextual`, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Resolve relationships to `MediaId` before the
renderer boundary. Keep repeated image lowering bounded even for malformed
placement values.

### F-097, Backgrounds

**Sprint.** S23
**Completed.** 2026-08-03
**Size.** S, estimated 1 day, actual 1 day

**What was built.** PresentationML backgrounds now retain their complete raw
subtree as the sole serialization source while exposing a typed rendering
projection for `p:bgPr` and `p:bgRef`. The resolver applies slide, layout,
master, then theme precedence, resolves style references and `phClr`, and the
renderer assigns concrete paint to the page background before shape content.

**Non-obvious choices.** The typed projection is read-only and never becomes a
second writer. Unsupported paint records a specific diagnostic, while the raw
subtree remains byte-for-byte preserved in its schema position.

**Deviations from the design plan.** Microscope pass 1 strengthened the exact
theme transform-order assertion. A later lint exposed an oversized fill enum
variant, which was boxed without changing the serialized contract. Microscope
pass 3 was clean.

**Spec sections touched.** `docs/hld/06-presentationml-model.md` and
`docs/hld/07-inheritance-and-resolution.md` now describe the preserving
background projection and concrete resolution path.

**Tests.** `background_projection_preserves_the_source_subtree_verbatim`,
`background_precedence_is_slide_layout_master_then_theme`,
`master_gradient_background_renders_when_slide_and_layout_omit_one`,
`background_reference_resolves_phclr_through_the_master_colour_map`,
`background_is_not_duplicated_in_page_elements`, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the captured background subtree as the
only serialization source and the typed form as a rendering projection only.

### F-098a, Text content box

**Sprint.** S24
**Completed.** 2026-08-05
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Shape text now starts from the evaluated preset or custom
geometry text rectangle, falls back to local shape bounds when needed, applies
all four resolved insets, and clamps malformed negative extents to zero.

**Non-obvious choices.** The content box remains in shape-local coordinates so
the existing shape group transform places paths and text through one boundary.
Missing geometry uses a visible bounds fallback instead of dropping text.

**Deviations from the design plan.** Microscope pass 1 found that the fallback
regression omitted the diagnosed bounds-fallback case. The test was expanded,
and pass 2 was clean.

**Spec sections touched.** None. The implementation follows the existing text
rectangle and body-inset contract.

**Tests.** `preset_text_rectangle_minus_unequal_insets_produces_the_computed_content_box`,
`missing_text_rectangle_falls_back_to_local_shape_bounds`,
`insets_larger_than_the_text_rectangle_do_not_create_negative_extents`, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the box local and finite before shaping.
Do not clip overflowing text at the shape boundary.

### F-098b, Paragraph inline resolution

**Sprint.** S24
**Completed.** 2026-08-05
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Resolved paragraphs now become shaped inline segments with
concrete size, fill, style, typeface, fields, and explicit break boundaries.
One font manager and shaping cache are reused across presentation layout.

**Non-obvious choices.** Typeface selection follows the script actually present
in the text, then uses the resolved Latin, East Asian, or complex-script slot.
Visible 18 point black sans-serif defaults cover missing resolved styling.

**Deviations from the design plan.** Microscope pass 1 found that a populated
Latin slot incorrectly overrode script-specific faces. Text-driven slot
selection and regressions fixed it, and pass 2 was clean.

**Spec sections touched.** None. The frozen resolver and shared shaping
boundaries already describe the implemented ownership.

**Tests.** `resolved_runs_emit_glyph_items_with_concrete_style_and_break_boundaries`,
`script_specific_text_selects_its_resolved_concrete_typeface`,
`repeated_resolved_runs_reuse_one_shaped_cache_entry`,
`text_shaping_failures_return_a_render_error_without_panicking`, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Select a concrete typeface before shaping and
keep shaping failures contextual instead of substituting silent empty output.

### F-098c, Line stacking

**Sprint.** S24
**Completed.** 2026-08-05
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Paragraph margins, hanging indents, wrapping, explicit
breaks, point and percentage spacing, horizontal alignment, and stacked
baselines now lower into positioned text and marker items without clipping.

**Non-obvious choices.** Text and markers share one baseline emitter. Percentage
line spacing uses the effective first-run size, while justified final lines
retain their ordinary alignment.

**Deviations from the design plan.** Microscope pass 1 corrected percentage
spacing that used natural metrics. Pass 2 added proof for justified lines and
production draw order. Pass 3 was clean.

**Spec sections touched.** None. The line-breaking and paragraph-order
contracts already cover the implementation.

**Tests.** `paragraphs_stack_wrapped_lines_with_spacing_and_alignment`,
`wrap_none_breaks_only_at_explicit_line_breaks`,
`percentage_line_spacing_uses_effective_first_run_font_size`,
`shape_text_stays_above_the_path_and_overflows_without_a_clip`, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep line breaking and baseline emission shared
between body text and markers so later decorations cannot drift vertically.

### F-098d, Text anchoring

**Sprint.** S24
**Completed.** 2026-08-05
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Complete text blocks now use top, centre, bottom, justified,
or distributed vertical placement inside the content box, then render above the
shape path within the existing group transform.

**Non-obvious choices.** Overflow remains visible. Justified anchoring allocates
spare height between line boxes, while distributed anchoring uses equal line
gaps with half a gap before the first line and after the last.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` now defines the exact
justified and distributed anchoring policy.

**Tests.** `bottom_center_text_in_an_inset_box_lands_at_the_computed_baseline`,
`top_center_and_bottom_anchors_use_zero_half_and_full_spare_height`,
`justified_and_distributed_anchors_allocate_positive_spare_height`,
`overflowing_anchored_text_remains_visible_without_a_clip`, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Measure the whole unanchored block first, then
apply one vertical placement policy without adding a clip.

### F-098, Shape text layout

**Sprint.** S24
**Completed.** 2026-08-05
**Size.** XL, estimated 0 days, actual 1 day

**What was built.** The umbrella closes the integrated content-box, inline
resolution, line-stacking, and anchoring pipeline delivered by F-098a through
F-098d. Shape text now follows the frozen resolved-slide boundary end to end.

**Non-obvious choices.** The umbrella has no independent source diff. Its gate
is the combined child evidence and deterministic bottom-centre baseline.

**Deviations from the design plan.** None. Every child completed its own plan,
review, tests, and delivery record before the parent closed.

**Spec sections touched.** `docs/hld/14-development-backlog.md` defines the four
implemented ownership boundaries and the combined parent gate.

**Tests.** F-098a through F-098d focused gates, the deterministic
`bottom_center_text_in_an_inset_box_lands_at_the_computed_baseline` regression,
and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Extend the existing private text module and
shared layout primitives. Do not introduce a second presentation text model.

### F-099, Bullets

**Sprint.** S24
**Completed.** 2026-08-05
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Character and automatic bullets now render with independent
size, colour, and typeface styling. Automatic counters are scoped per body and
level across eight common formats, and Wingdings F0B7 maps to a visible Unicode
bullet.

**Non-obvious choices.** Scheme changes, character bullets, no-bullet
paragraphs, and shallower levels reset the relevant sequence. Unsupported
schemes retain a visible Arabic-period marker, and the hanging slot stays fixed
even when a marker is wider than it.

**Deviations from the design plan.** Microscope pass 1 found that long markers
expanded the fixed hanging slot. The slot was fixed and an oversized-marker
regression added. Pass 2 was clean.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` now records the eight
formats, counter resets, visible fallback, and Wingdings mapping.

**Tests.** `wingdings_f0b7_bullet_renders_as_a_visible_unicode_glyph`,
`automatic_bullets_increment_and_reset_by_level`,
`eight_common_auto_number_schemes_format_exact_markers`,
`wide_auto_number_marker_keeps_text_on_the_paragraph_margin`, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep numbering state inside one resolved body and
keep marker measurement independent from the fixed text margin.

### F-100, Autofit

**Sprint.** S24
**Completed.** 2026-08-05
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Stored normal-autofit font scale and line-spacing reduction
now apply verbatim. Bare normal autofit selects from a deterministic 2.5 percent
ladder down to 25 percent, while no-autofit and shape-autofit keep visible
overflow behavior.

**Non-obvious choices.** Only extra leading is reduced, never the font metrics.
Each ladder attempt measures inside paragraph margins and reuses shaping work
within the calculation.

**Deviations from the design plan.** None. Microscope pass 1 was clean and added
no remediation.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` now defines the
25 percent floor, line-spacing floor, and visible smallest-candidate overflow.

**Tests.** `stored_font_scale_renders_at_exactly_sixty_two_point_five_percent`,
`stored_line_spacing_reduction_reduces_only_extra_leading`,
`bare_normal_autofit_uses_quantised_two_point_five_percent_steps`,
`bare_normal_autofit_keeps_the_twenty_five_percent_floor_visible`, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the ladder deterministic and avoid clipping
when even its smallest candidate cannot fit.

### F-101, Vertical text

**Sprint.** S24
**Completed.** 2026-08-05
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Vertical and vertical-270 text reuse horizontal layout in a
transposed content box, then apply opposite centre-preserving quarter turns.
East Asian vertical and other unsupported variants remain visible through
documented rotations and stable diagnostics.

**Non-obvious choices.** The resolver owns fallback diagnostics, while the
renderer receives only concrete direction and applies the affine transform to
one grouped text block.

**Deviations from the design plan.** Microscope pass 1 found an exact renderer
mapping coverage gap. The regression was added, and pass 2 was clean.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md` and
`docs/hld/08-rendering-spec.md` now define direction mapping and visible
diagnosed fallbacks.

**Tests.** `vertical_text_uses_a_transposed_box_and_rotated_group`,
`vertical_270_uses_the_opposite_quarter_turn`,
`east_asian_vertical_text_degrades_to_rotated_with_a_diagnostic`,
`other_vertical_variants_remain_visible_with_diagnostics`, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep vertical fallbacks visible and diagnosed.
Do not add a separate vertical shaping pipeline.

### F-102, Table rendering

**Sprint.** S25
**Completed.** 2026-08-08
**Size.** L, estimated 4 days, actual 1 day

**What was built.** DrawingML table styles, cell fills, margins, borders, and
text styles now parse and preserve unsupported XML, resolve through concrete
table-region precedence, and lower to fills, text, and one physical stroke per
border segment. Two-dimensional merges render only from their origin while
using the correct covered-cell outer edges.

**Non-obvious choices.** Table text style enters the character cascade before
explicit paragraph and run formatting. Corner regions require their matching
row and column options, adjacent borders use one deterministic conflict policy,
and unsupported table paint or cell autofit stays visible through stable
diagnostics.

**Deviations from the design plan.** Microscope passes exposed explicit text
overrides applied in the wrong order, unequal right-to-left column sizing,
merged-edge sourcing, inside-border selection, option-independent corners, and
missing unsupported-paint diagnostics. Each was corrected with a distinguishing
regression before the clean third pass.

**Spec sections touched.** `docs/hld/05-drawingml-model.md`,
`docs/hld/06-presentationml-model.md`,
`docs/hld/07-inheritance-and-resolution.md`,
`docs/hld/08-rendering-spec.md`, and `docs/hld/12-testing-strategy.md`.

**Tests.** `table_style_and_cell_properties_preserve_unmodelled_xml_byte_for_byte`,
`table_style_regions_resolve_in_documented_precedence`,
`table_cell_autofit_is_ignored_and_records_a_diagnostic`,
`banded_merged_table_renders_correct_fills_without_duplicated_borders`,
`merged_continuation_cells_do_not_render_fill_border_or_text_twice`,
`table_cell_margins_place_text_in_the_fixed_content_box`, and the integrated
full workspace gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Resolve logical cell ownership separately from
physical edge ownership. A merge origin owns content, but its far border can
come from the last covered cell.

### F-103, Hyperlinks, fields and diagnostics

**Sprint.** S25
**Completed.** 2026-08-08
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Direct external run hyperlinks now resolve in the producing
slide, layout, or master relationship scope and emit transformed URI link
annotations. Typed and effective slide-number fields substitute the one-based
page number before shaping, while broken or unsupported actions retain visible
text and add stable diagnostics.

**Non-obvious choices.** Hyperlink actions remain direct run state rather than
joining the inheritance cascade. The resolver freezes only the resolved URI,
and the existing text-segment emitter supplies annotation bounds after the
normal recursive transform walk.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/07-inheritance-and-resolution.md` and
`docs/hld/08-rendering-spec.md`.

**Tests.** `slide_number_field_renders_current_page_and_hyperlink_emits_annotation`,
`same_relationship_id_resolves_hyperlink_in_its_shape_source_scope`,
`missing_hyperlink_relationship_keeps_text_and_records_diagnostic`,
`untyped_slide_number_placeholder_uses_the_current_page_number`,
`grouped_hyperlink_annotation_keeps_transformed_run_bounds`, and the integrated
full workspace gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep package relationship objects upstream of
the frozen renderer contract. Unsupported actions should lose only the
annotation, never their visible text.

### F-104, SSIM fidelity harness

**Sprint.** S25
**Completed.** 2026-08-08
**Size.** L, estimated 4 days, actual 1 day

**What was built.** A deterministic whole-presentation entry point, a concrete
50-deck renderer, and a version-pinned LibreOffice SSIM harness now produce and
retain per-slide fidelity evidence at 150 dpi. CI enforces complete rendering,
records the SSIM trend, uploads the detailed evidence even on failure, and keeps
native PowerPoint review as the hard manual fidelity gate. Evidence-ranked
renderer corrections raised the integrated trend to 30 of 421 slides at or
above 0.95 with median SSIM 0.622465 and zero dropped bounded shapes.

**Non-obvious choices.** SSIM against LibreOffice is a trend rather than a
conformance threshold. Native PowerPoint 16.104 versus the same LibreOffice
oracle reached zero of 34 representative slides at 0.95, with median
0.650406194, so completeness and native review remain the hard gates. The
normal renderer still discovers system fonts, while the evidence entry point
uses bundled fonts exclusively.

**Deviations from the design plan.** Native calibration superseded the initial
hard interpretation of 0.95 SSIM on 80 percent of slides. Microscope required
durable CI evidence uploads and clearer acceptance sources. Completion also
removed an out-of-scope mirrored Word layout change after it produced an
undeclared hash delta.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/07-inheritance-and-resolution.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `test_all_corpus_slides_render_without_panic_or_dropped_shape`,
`test_corpus_render_fidelity_records_ssim_trend`, the seven local metric and
oracle-contract self-tests, exact tool-version assertions, the accepted M10
native PowerPoint spot-check, and the integrated 50-deck gate passed for all
421 slides.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Integrated evidence is retained at
`/private/tmp/s25-integrated-fidelity-escalated-20260808`. On macOS, sandboxed
headless LibreOffice can abort during AppKit application registration before it
opens a deck. The same exact command and deck succeed outside that boundary.

### F-105, Bundled default.pptx

**Sprint.** S26
**Completed.** 2026-08-08
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The unpublished `rpptx` facade now owns a crate-local
zero-slide PowerPoint template and exposes `Presentation::new()` when the
default-on `default-template` feature is enabled. The 16:9 template contains
one master, eleven layouts, a full theme, notes infrastructure, and table
styles.

**Non-obvious choices.** The template is loaded through the normal parser and
serialized through the deterministic package writer. Shipping a reviewed
binary template avoids thousands of lines of write-only theme and layout
construction code.

**Deviations from the design plan.** None. The native PowerPoint gate and both
feature modes passed as planned.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/06-presentationml-model.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `new_presentation_uses_the_bundled_zero_slide_template`,
`bundled_template_has_the_documented_part_graph`, default and no-default
feature checks, package-list inspection, native PowerPoint no-repair
acceptance, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the binary asset inside the `rpptx` crate
and retain its recorded source and licence evidence when replacing it.

### F-106, ShapeIdAllocator and MediaStore

**Sprint.** S26
**Completed.** 2026-08-08
**Size.** M, estimated 2 days, actual 1 day

**What was built.** PresentationML shape trees now expose typed non-visual ids
and allocate fresh ids across root shapes, nested groups, and selected
alternate-content fallbacks. The facade-owned media store deduplicates equal
bytes by content hash and allocates collision-free package part names.

**Non-obvious choices.** Hash equality is confirmed with the original bytes so
a hash collision cannot alias different media. Parsed non-visual properties
remain backed by their preserved raw XML, and allocation starts at id 2 because
the shape-tree root owns id 1 in generated decks.

**Deviations from the design plan.** Microscope pass 1 found that insertion
order could change which duplicate media part was reused. Sorting existing
part names made reuse deterministic, and pass 2 was clean.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, and
`docs/hld/06-presentationml-model.md`.

**Tests.** `shape_id_allocator_scans_nested_groups_and_alternate_content`,
`shape_id_allocator_starts_at_two_and_skips_sparse_ids`,
`typed_non_visual_ids_preserve_original_shape_xml`,
`equal_media_bytes_inserted_twice_reuse_one_part`,
`media_store_compares_bytes_inside_a_hash_bucket`,
`media_store_allocates_after_the_highest_existing_suffix`, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep `MediaStore` private to the facade and keep
the allocator scan aligned with every recursive shape-tree container.

### F-107, add_slide

**Sprint.** S26
**Completed.** 2026-08-08
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `Presentation::add_slide()` now selects a resolved layout,
synthesizes a minimal slide with its non-latent placeholders, assigns unique
shape and slide ids, creates the relative layout and presentation
relationships, registers the content type, and appends the slide to the owned
read model.

**Non-obvious choices.** Placeholder type and `idx` are copied without cloning
the layout XML, which avoids hidden relationship identifiers. Date, footer,
and slide-number placeholders remain latent, while every synthesized text body
contains its required paragraph.

**Deviations from the design plan.** Microscope passes strengthened sparse
part-name allocation and the test oracle for placeholder inheritance. The
native three-slide deck then opened in PowerPoint 16.104 without repair.

**Spec sections touched.** `docs/hld/01-glossary.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/06-presentationml-model.md`, and
`docs/hld/13-risks-and-open-questions.md`.

**Tests.** `three_added_slides_have_unique_ids_and_reopen`,
`add_slide_allocates_after_the_highest_existing_part_suffix`,
`add_slide_synthesizes_only_non_latent_layout_placeholders`,
`synthesized_slide_uses_schema_order_and_one_relative_layout_relationship`,
`synthesized_text_bodies_always_contain_a_paragraph`,
`add_slide_rejects_an_unknown_layout_index_without_mutation`, native
PowerPoint acceptance, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Continue using the nine-step package mutation
order and allocate from observed maxima so sparse producer packages are never
overwritten.

### F-108, validate()

**Sprint.** S26
**Completed.** 2026-08-08
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Presentation::validate()` now returns all twelve exact
package and PresentationML issue variants in deterministic order. It checks
slide and shape ids, text bodies, placeholders, content types, relationships,
media reachability, custom shows, layouts, and themes. `save()` writes the same
deterministic bytes as `to_bytes()`, and both debug save boundaries assert a
clean validation result before writing.

**Non-obvious choices.** Semantic slide-id and empty-text checks are deferred
from parsers so corrupted decks remain inspectable. Namespace-aware XML scans
are observational, uppercase media extensions use case-insensitive default
lookup, and shape traversal uses an explicit heap stack to keep validation
non-panicking for deeply nested trees.

**Deviations from the design plan.** Microscope pass 1 found a hard-coded root
shape id, skipped XML parts whose relationship collection was entirely absent,
and recursive traversal that could overflow the stack. All three were repaired
with distinguishing regressions, and pass 2 was clean.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** `every_validation_issue_variant_detects_its_corrupted_deck`,
`validate_collects_all_issues_in_deterministic_order`,
`all_pinned_corpus_decks_validate_cleanly`,
`debug_save_boundaries_assert_on_invalid_presentations`,
`save_writes_the_same_bytes_as_to_bytes`,
`validation_xml_scan_is_prefix_tolerant_and_non_mutating`, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep validation observational and preserve
corrupt input long enough to report every issue. Relationship scans must cover
parts even when their entire `.rels` collection is absent.

### F-109, Shape mutation facade

**Sprint.** S27
**Completed.** 2026-08-08
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The presentation facade now exposes borrowed mutable slide
and recursive shape handles. Supported shapes can update position, size,
rotation, name, fill, line, and preset adjustments, and every change survives
save and reload without replacing the owning shape tree.

**Non-obvious choices.** Mutable access follows the typed shape-tree projection
and deliberately leaves selected `AlternateContent` fallback children
read-only. Setters update the narrow typed field while preserving unmodelled
attributes, children, sibling order, and non-visual ids.

**Deviations from the design plan.** Microscope pass 1 found that creating an
absent group transform could place it after preserved group properties. The
repair shifted the raw boundary and strengthened nested-group coverage to prove
two sibling ids and their order remain unchanged. Pass 2 was clean.

**Spec sections touched.** `docs/hld/06-presentationml-model.md`.

**Tests.** `shape_mutation_setters_survive_save_and_reload`,
`shape_mutation_preserves_unmodelled_xml_and_schema_order`,
`shape_mutation_handles_nested_group_children`,
`alternate_content_fallback_is_not_mutable`,
`shape_mutation_indices_and_kinds_are_total`,
`preset_adjustment_setter_inserts_and_replaces_named_values`,
`shape_name_mutation_escapes_xml_and_preserves_children`, and the integrated
full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep behavior-bearing raw-boundary repair beside
the typed mutation that creates an absent schema child. Do not expose mutable
handles into compatibility branches that remain serialized from preserved XML.

### F-110, Shape constructors

**Sprint.** S27
**Completed.** 2026-08-08
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Mutable slides can now append text boxes, preset shapes,
connectors, and empty groups at the top of z-order. Each constructor allocates
a tree-wide unique id, emits a canonical schema-ordered shell, and returns a
borrowed handle to the exact appended shape.

**Non-obvious choices.** The id allocator scans typed members, preserved raw
members, and every markup-compatibility branch using namespace-resolved
`cNvPr` matching. Connector bounds normalize every endpoint direction into
nonnegative extents and flips, and preset names are checked against all 187
pinned ECMA definitions before mutation.

**Deviations from the design plan.** Microscope pass 1 found append ordering
after a trailing extension, opaque-id collisions, unvalidated preset strings,
and incomplete reopen assertions. The repairs added a preservation-aware append
path, a complete raw scan, preset validation, and full geometry assertions.
Pass 2 was clean.

**Spec sections touched.** `docs/hld/06-presentationml-model.md`.

**Tests.** `all_shape_constructors_open_in_powerpoint_without_repair`,
`ordinary_shape_and_textbox_constructors_emit_canonical_shells`,
`connector_constructor_normalizes_every_direction`,
`empty_group_constructor_has_required_children`,
`four_appended_shapes_have_unique_ids_and_reopen`,
`constructor_names_are_deterministic_from_allocated_ids`, pinned PowerPoint
16.104 build 16.104.25121423 acceptance, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Append through `CT_ShapeTree::append_child` so
schema-final preserved content remains final. Keep the allocator scan aligned
with every typed and opaque shape-tree container.

### F-111, add_picture

**Sprint.** S27
**Completed.** 2026-08-08
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Presentation::add_picture` now inserts a picture into a
chosen slide, using exact 72-DPI native dimensions when both axes are omitted
and truncating aspect-ratio inference when one axis is supplied. Equal image
bytes share one media part while each slide retains its own relationship scope.

**Non-obvious choices.** Image bytes determine the stored extension and MIME
type even when the supplied filename is misleading. Package, media, and
relationship changes are staged in clones so every fallible operation finishes
before the live presentation commits an atomic mutation.

**Deviations from the design plan.** None. Microscope pass 1 was clean. The
pinned python-pptx 1.0.2 comparison produced the same picture kind and 12,700 by
12,700 EMU bounds, and PowerPoint 16.104 build 16.104.25121423 opened the deck
without repair.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** `picture_without_explicit_size_uses_native_dimensions`,
`picture_constructor_round_trips_in_schema_order`,
`picture_one_dimension_preserves_aspect_ratio_with_truncation`,
`duplicate_picture_bytes_share_one_media_part_across_slides`,
`picture_sniffs_bytes_when_extension_is_misleading`,
`invalid_picture_input_does_not_mutate_the_presentation`,
`picture_native_size_matches_python_pptx_1_0_2`,
`added_picture_validates_and_opens_without_repair`, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep media deduplication package-wide and image
relationship reuse slide-scoped. Preserve the staged commit boundary whenever
new fallible picture options are added.

### F-112, Text frame mutation

**Sprint.** S27
**Completed.** 2026-08-08
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Ordinary shapes now expose borrowed text-frame, paragraph,
and run handles. Callers can replace or clear text, append paragraphs and runs,
and set typed paragraph, character, Latin font, and bullet properties while
retaining the required nonempty paragraph invariant.

**Non-obvious choices.** Replacing placeholder text preserves placeholder type
and `idx`, body properties, list style, and paragraph-level state. The mutation
surface keeps fields, breaks, raw compatibility content, whitespace intent, and
schema order unless the caller replaces the owning typed value.

**Deviations from the design plan.** Microscope pass 1 found that creating an
absent paragraph property node could leave preserved boundary-0 content before
it. The repair moved that raw boundary behind the new property and added a
complete markup-compatibility run substitution regression. Pass 2 was clean.

**Spec sections touched.** `docs/hld/05-drawingml-model.md` and
`docs/hld/06-presentationml-model.md`.

**Tests.** `setting_text_on_placeholder_round_trips_and_renders`,
`clearing_text_preserves_required_paragraph`,
`paragraph_run_font_and_bullet_properties_round_trip`,
`text_mutation_preserves_placeholder_identity`,
`text_mutation_preserves_unmodelled_xml_and_schema_order`,
`text_mutation_indices_and_shape_kinds_are_total`,
`text_frame_handles_append_paragraphs_and_runs_in_order`, and the integrated
full gate passed with deterministic fonts.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the minimum one-paragraph invariant in the
DrawingML model and keep raw-boundary shifts local to the property insertion
that changes schema order.

### F-113, Table facade

**Sprint.** S28
**Completed.** 2026-08-09
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Mutable slides can now append tables and expose borrowed
table and cell handles. Callers can edit cell text, fill, margins, banding,
column widths, and rectangular merges, then split a merged origin back into
the original grid while retaining valid table structure.

**Non-obvious choices.** Merge moves source-cell content into the origin in
row-major order, matching pinned python-pptx 1.0.2 semantics. Width updates use
checked EMU arithmetic and keep the table grid synchronized with the graphic
frame extent. Unsupported XML remains at its original schema boundary.

**Deviations from the design plan.** Microscope pass 1 found missing required
paragraphs after content migration, missing constructor validation, and a
preservation test that did not prove byte identity. All three were repaired,
and pass 2 was clean.

**Spec sections touched.** `docs/hld/05-drawingml-model.md` and
`docs/hld/06-presentationml-model.md`.

**Tests.** `merge_then_split_restores_the_original_grid`,
`add_table_round_trips_cells_formatting_banding_and_widths`,
`table_mutation_rejects_invalid_ranges_without_partial_changes`,
`table_mutation_preserves_unmodelled_xml_and_schema_order`,
`table_graphic_frame_constructor_writes_the_canonical_shell`, the pinned
python-pptx 1.0.2 differential gate, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep cell-grid mutation transactional. A merged
source must retain a schema-valid text body even after its content moves to the
origin.

### F-114, remove_slide, move_slide, duplicate_slide

**Sprint.** S28
**Completed.** 2026-08-09
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Presentation` can remove, move, and duplicate slides while
keeping the slide-id list, relationships, parts, content types, notes, custom
shows, media, shape ids, and connector endpoints consistent. Duplicated images
resolve through the new slide's own relationship scope.

**Non-obvious choices.** Duplication remaps typed and preserved relationship
ids, creates fresh notes and back relationships, and allocates fresh shape ids
through compatibility content. Removal prunes only candidate media that is no
longer reachable from any package relationship.

**Deviations from the design plan.** Microscope pass 1 found package-root media
reachability, notes normalization, and compatibility shape-id gaps. Pass 2
found a remaining nonnumeric preserved notes reference. All four defects were
repaired, and pass 3 was clean.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md` and
`docs/hld/06-presentationml-model.md`.

**Tests.** `duplicated_slides_images_resolve_to_the_new_slides_own_relationships`,
`remove_slide_removes_its_part_relationship_notes_and_custom_show_entries`,
`move_slide_reorders_the_slide_id_list_without_rewriting_relationships`,
`duplicate_slide_rewrites_typed_and_preserved_relationship_ids_without_other_byte_changes`,
`slide_id_list_raw_children_follow_surviving_ids_after_collection_edits`,
`every_corpus_preserved_payload_is_identity_with_an_empty_map`, and the
integrated full gate passed against all 50 pinned decks.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Treat slide collection edits as package-graph
transactions. Relationship ids are scoped by producer part, and media pruning
must include package-root reachability.

### F-115, Slide and presentation properties

**Sprint.** S28
**Completed.** 2026-08-09
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The facade now reads and writes slide size, hidden state,
direct slide backgrounds, and core properties. It can also save presentation
bytes as a slideshow package while preserving a valid part graph and changing
only the main content type.

**Non-obvious choices.** Hidden state follows inverse `p:sld/@show` semantics.
Direct background edits preserve theme references and raw producer XML. Core
properties are created only when the package owns or lacks the conventional
part, avoiding unrelated-part replacement.

**Deviations from the design plan.** Microscope pass 1 found direct-background
replacement losing preserved XML, unsafe reuse of an unowned conventional core
part, and incomplete value and preservation assertions. All three were
repaired, and pass 2 was clean.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md` and
`docs/hld/06-presentationml-model.md`.

**Tests.** `slide_and_presentation_properties_round_trip`,
`slide_size_mutation_preserves_kind_and_unmodelled_xml`,
`hidden_flag_uses_inverse_show_semantics`,
`background_set_and_clear_preserve_theme_references_and_raw_xml`,
`core_properties_are_loaded_lazily_and_written_with_valid_graph`,
`save_as_show_changes_only_the_main_content_type`, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep slideshow conversion output-only and keep
core-property creation guarded by relationship ownership rather than a
conventional path alone.

### F-116, Cross-viewer acceptance

**Sprint.** S28
**Completed.** 2026-08-09
**Size.** M, estimated 2 days, actual 1 day

**What was built.** One deterministic ten-slide deck exercises every F-107
through F-115 write feature, validates cleanly, reopens as `.pptx` and `.ppsx`,
and binds four-viewer evidence to SHA-256
`d36da6e8849eabd4487d2572baea19c3716ee7d0fe03aaa4714a28ce3c41de4f`.

**Non-obvious choices.** The ignored gate reruns automatable PowerPoint and
LibreOffice checks, then validates all four tracked evidence rows. Keynote is
user-confirmed human-action evidence because its UI import is not reliably
scriptable. Google Slides is identified by acceptance date and Chrome build,
without retaining the private imported-document URL.

**Deviations from the design plan.** Review hardened the evidence schema
against unobserved counts, pending or blank clean records, vacuous pending
coverage, and shared temporary paths. It also aligned the gate and plan with
the supported Keynote human-action path. The test-only one-pixel PNG was
replaced with a valid precomputed fixture after LibreOffice rejected its IDAT
stream. Microscope pass 8 was clean.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`.

**Tests.** `ten_slide_write_api_deck_validates_and_reopens`,
`ten_slide_write_api_deck_saves_as_presentation_and_show`,
`cross_viewer_acceptance_evidence_is_complete_and_bound_to_one_artifact`,
`generated_ten_slide_write_api_deck_opens_clean_in_all_four_viewers`, pinned
PowerPoint 16.104, Keynote 14.4, Google Slides on 2026-08-09 through Chrome
151.0.7922.76, LibreOffice 26.2.5.2 acceptance, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep every viewer row bound to the same frozen
artifact SHA. Never promote a pending row without a positive observed open or
import, exact slide count, clean conversion result, and close or export result.

### F-117, oxml-sml workbook writer

**Sprint.** S29
**Completed.** 2026-08-10
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The workspace now contains an unpublished `oxml-sml`
crate that writes one-sheet `.xlsx` packages with string and numeric columns,
shared strings, number formats, defined ranges, deterministic relationships,
and the complete minimal OPC part graph.

**Non-obvious choices.** The API is column-oriented and validates all lengths,
formula ranges, finite numbers, sheet names, shared-string counts, and XML
string escapes before package construction. Shared-string indexes are stable,
and workbook output is byte-identical for the same input.

**Deviations from the design plan.** Microscope review added complete
SpreadsheetML escaping for attribute normalization and reserved sequences,
row-limit validation, aggregate shared-string overflow protection, and an
executable viewer-artifact binding. The approved crate boundary did not
change.

**Spec sections touched.** `docs/hld/09-charts-spec.md` and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `workbook_package_has_the_minimal_editable_part_graph`,
`formula_ranges_quote_sheet_names_and_track_column_lengths`,
`spreadsheet_strings_escape_xml_and_reserved_sequences_exactly`,
`viewer_gate_candidate_is_bound_to_recorded_sha`, and the integrated full
gate. Excel 16.104 and LibreOffice Calc 26.2.5.2 opened the same artifact
without repair at SHA-256
`8f8d12aa4ebe94f86c8164fd251cdb23845f985090be0fb6c77242aaa0fba329`.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep `oxml-sml` deliberately smaller than a
spreadsheet library. Chart authoring should consume its deterministic workbook
package and formula ranges rather than adding workbook-reading or calculation
features here.

### F-118, ChartML core types

**Sprint.** S29
**Completed.** 2026-08-10
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The workspace now contains an unpublished `rpptx-chart`
crate with typed ChartML space, chart, plot-area, title, legend, flags, shape
properties, and text properties. It reads aliases by namespace URI, writes
fixed `c`, `a`, and `r` prefixes in schema order, and preserves unsupported
ChartML at stable schema boundaries.

**Non-obvious choices.** The core plot shells remain intentionally opaque for
later plot and axis stories. DrawingML text parsing accepts the caller-owned
`c:txPr` root while reusing the existing concrete text-body implementation.
Corpus preservation evidence retains parent path, schema boundary, sibling
order, and exact bytes.

**Deviations from the design plan.** Microscope review hardened schema-slot
stability after public edits, comments and processing instructions, scalar
extension data, nested namespace handling, first-parse preservation evidence,
and trailing root validation. The approved core-only type boundary remained
unchanged.

**Spec sections touched.** `docs/hld/09-charts-spec.md` and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `chart_space_reads_aliases_and_writes_fixed_prefixes_in_schema_order`,
`core_chart_shells_preserve_unmodelled_children_byte_for_byte`,
`malformed_core_chart_values_return_errors_without_panicking`,
`every_corpus_chart_part_round_trips_structurally`, and the integrated full
gate. The required corpus gate verified 26 chart parts across 9 of 50 pinned
decks.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep plot kinds, axes, and series attached to
the stable raw schema seams until their owning F-IDs type them. Namespace
resolution must remain URI-aware even when output prefixes are canonical.

### F-119, Series and data references

**Sprint.** S29
**Completed.** 2026-08-10
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `rpptx-chart` now models series indexes and order, names,
string and numeric references, category and value data, bubble sizes, formulae,
format codes, and literal caches. Constructors derive cache counts and point
indexes from one value vector, while corpus parsing retains valid sparse
producer caches and unsupported series payloads.

**Non-obvious choices.** Series projection resolves mixed and inherited
prefixes by namespace URI. Preserved payloads use stable schema slots so public
edits cannot reorder markers, labels, error bars, shapes, or extensions. Typed
fixed-prefix rewrites reject conflicting local bindings, and cache edits
reconcile preserved point boundaries without dropping schema-final payloads.

**Deviations from the design plan.** The corpus demonstrated that a logical
point count may exceed the number of cached points when all retained indexes
remain in range, so the plan and HLD were aligned with valid sparse caches.
Nine microscope passes hardened namespace propagation, duplicate wrapper
detection, public-edit behavior, cache resizing, and schema-slot stability
before the independent review became clean.

**Spec sections touched.** `docs/hld/09-charts-spec.md`.

**Tests.** `series_formula_and_cache_are_consistent_with_one_source`,
`string_and_numeric_references_write_fixed_prefixes_in_schema_order`,
`malformed_series_and_cache_values_return_errors_without_panicking`,
`series_preserves_unmodelled_children_byte_for_byte`,
`public_series_edits_do_not_duplicate_or_drop_preserved_payloads`,
`every_corpus_series_round_trips_structurally`, and the integrated full gate.
The required corpus gate verified 66 series across all 26 chart parts in the
50 pinned decks.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Authoring code should supply formulae and value
vectors once and let the ChartML writer derive cache metadata. Plot stories
must preserve the stable series schema slots and URI-aware namespace rules.

### F-120, Axes

**Sprint.** S30
**Completed.** 2026-08-10
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `rpptx-chart` now models category, value, date, and series
axes with scaling, gridlines, titles, number formats, ticks, shape and text
properties, positions, and reciprocal cross-axis references. Plot areas expose
validated typed axes while preserving unsupported children and producer
markup in schema order.

**Non-obvious choices.** Axis identifiers accept the producer-compatible range
from signed 32-bit minimum through unsigned 32-bit maximum because the corpus
contains negative PowerPoint identifiers. Parsed root-family provenance blocks
unsafe relabelling when opaque family-specific content exists, while newly
constructed axes remain freely editable.

**Deviations from the design plan.** Five microscope passes hardened parsed
axis relabelling, constructed versus parsed provenance, structural equality,
and lexical identifier preservation. The approved API and HLD impact did not
change.

**Spec sections touched.** `docs/hld/09-charts-spec.md`.

**Tests.** `axis_id_pairs_are_reciprocal`,
`all_axis_forms_write_fixed_prefixes_in_schema_order`,
`malformed_axis_values_return_errors_without_panicking`,
`axes_preserve_unmodelled_children_byte_for_byte`,
`every_corpus_axis_round_trips_structurally`, and the integrated full gate.
The required corpus gate verified 40 axes across 26 chart parts in all 50
pinned decks.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep axes owned by the plot area and plots
limited to identifier references. Preserve unchanged identifier lexemes for
round-trip output, but compare and validate their normalized numeric values.

### F-123, Data labels and number formats

**Sprint.** S30
**Completed.** 2026-08-10
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Series now carry typed collection-level data labels with
number format, position, separator, and visibility flags. The shared
`NumberFormat` value projects `General`, fixed decimal, and percentage forms
deterministically while preserving unsupported valid producer codes for
round-trip output.

**Non-obvious choices.** Cache source formatting and label formatting remain
separate XML states. Individual point labels, leader lines, shape properties,
text properties, and extensions remain ordered raw payloads. Native glyph
placement remains outside this model boundary.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/09-charts-spec.md`.

**Tests.** `data_labels_write_fixed_prefixes_in_schema_order`,
`common_number_formats_project_cached_values_deterministically`,
`malformed_data_labels_and_number_formats_return_errors_without_panicking`,
`data_labels_preserve_point_overrides_and_extensions_byte_for_byte`,
`every_corpus_data_label_collection_round_trips_structurally`,
`percentage_formatted_label_renders_with_correct_text`, and the integrated
full gate. The corpus gate verified 34 label collections and 35 axis number
formats. LibreOffice 26.2.5.2 and Poppler 26.01.0 extracted `25%` from candidate
SHA-256 `4ba02faa8e4cff6cefa7a7dc73fc0eb0c08d62d180f83fa0d3fd56a7e4136242`.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Treat `format_value` as a deliberately small
renderer projection, not an Excel format-language engine. F-126 owns label
placement and should consume this typed state without reparsing ChartML.

### F-121, Bar and line plots

**Sprint.** S30
**Completed.** 2026-08-10
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Plot areas now own typed single-family bar and line plots
with validated properties, series, data labels, and exactly two references to
the plot-area axis set. Unsupported 3-D plots and combination choices remain
opaque and byte-preserved.

**Non-obvious choices.** Mutable repeated series and axis references reconcile
preserved raw boundaries through stable exact and positional matching. Parsed
bar and line families cannot be relabelled while incompatible preserved
payload remains. Viewer acceptance compares decoded pixels with a zero-error
threshold rather than adding native chart geometry.

**Deviations from the design plan.** Eight microscope passes hardened malformed
single-family validation, exact viewer equality, typed versus opaque corpus
counts, repeated-child raw boundaries, mutable identity reconciliation,
family replacement, and the binary PPM parser.

**Spec sections touched.** `docs/hld/09-charts-spec.md`.

**Tests.** `bar_and_line_plots_round_trip_and_render`,
`bar_and_line_plots_write_fixed_prefixes_in_schema_order`,
`malformed_bar_and_line_plots_return_errors_without_panicking`,
`unsupported_and_combo_plots_remain_byte_preserved`,
`public_plot_edits_preserve_axes_and_unselected_payloads`,
`every_corpus_bar_and_line_plot_round_trips_structurally`, and the integrated
full gate. The corpus gate verified 11 typed bar plots, 2 typed line plots, and
one preserved bar-line combination. Pinned original and candidate renders had
normalized RGB mean absolute error `0.00000000` for both families.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** A plot owns axis references, not axis objects.
Keep combination choices opaque until a dedicated story can model and validate
the entire choice without partially rewriting it.

### F-122, Pie, doughnut, area, scatter and radar plots

**Sprint.** S30
**Completed.** 2026-08-10
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The typed plot boundary now covers all seven v1 families.
Pie and doughnut plots are axis-free, area and radar plots use paired axes, and
scatter plots map existing numeric category and value caches to `c:xVal` and
`c:yVal`. Unsupported bubble, stock, surface, `ofPie`, and 3-D choices remain
opaque.

**Non-obvious choices.** Scatter wrapper provenance remains private on the
shared public `Series` value, so standalone and plot-level round trips retain
the correct wrappers. Typed families reject both public and preserved bubble
payload. Family-specific raw boundaries remain stable when optional typed
children are inserted or removed.

**Deviations from the design plan.** Three microscope passes hardened optional
child insertion order, standalone scatter wrapper preservation, bubble-only
payload rejection, the malformed-input matrix, and live raw-boundary updates.
The corpus gap for four families remained as designed and is covered by inline
fixtures plus pinned viewer candidates.

**Spec sections touched.** `docs/hld/09-charts-spec.md`.

**Tests.** `remaining_v1_plots_round_trip_and_render`,
`remaining_plot_families_write_fixed_prefixes_in_schema_order`,
`scatter_series_map_numeric_categories_and_values_to_x_and_y`,
`malformed_remaining_plots_return_errors_without_panicking`,
`unsupported_plot_families_and_children_remain_byte_preserved`,
`every_supported_corpus_plot_round_trips_structurally`, and the integrated full
gate. The corpus supplied one typed pie plot. SHA-bound candidates for pie,
doughnut, area, scatter, and radar exceeded the 1,000 nonblank-pixel threshold
with counts of 309,502, 233,915, 308,569, 9,865, and 7,161.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep one public series type and retain scatter
wrapper provenance privately. F-125 owns native geometry for every plot family
and should consume these typed values rather than duplicating ChartML parsing.

### F-124, add_chart

**Sprint.** S31
**Completed.** 2026-08-10
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `Presentation::add_chart` now validates one `ChartData`
value and atomically writes the typed ChartML part, editable workbook, slide
and chart relationships, content-type overrides, and canonical graphic frame.
All seven supported two-dimensional chart families share this authoring path.

**Non-obvious choices.** Chart and workbook part suffixes advance independently
from their greatest occupied positive suffix. Workbook cells and ChartML caches
are derived from the same source value. Every fallible package and slide change
is staged before the live presentation is updated.

**Deviations from the design plan.** None. Three microscope passes hardened
independent part numbering, nonpositive extent rollback, relationship-id
rollover, exact cache-to-cell mapping, and the native PowerPoint gate.

**Spec sections touched.** `docs/hld/09-charts-spec.md`.

**Tests.** `add_chart_writes_complete_relationship_graph`,
`add_chart_uses_collision_free_part_numbers`,
`add_chart_caches_and_workbook_share_one_source`,
`add_chart_rejects_invalid_data_without_mutation`,
`authored_chart_graphic_frame_round_trips`, and the integrated full gate.
`authored_chart_enters_renderer_deterministically` parses the ChartML produced
by the owning facade and proves finite, nonempty, repeatable paths and labels.
Microsoft PowerPoint 16.104, build 16.104.25121423, opened candidate SHA-256
`e6e9f7eef1c774d0414c5d0c3f1202da1a28635b5d089e15455b7adc3f66cb00`
without repair. Edit Data showed the authored Category, Revenue, and Cost
values for North, South, and West.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep package mutation on the owning presentation
facade. Do not split the workbook, ChartML, relationships, and slide frame
across separate public operations.

### F-125, Chart rendering: geometry

**Sprint.** S31
**Completed.** 2026-08-10
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `rpptx-chart` now lowers all seven supported plot families
to finite backend-neutral paths and markers inside stable chart-local bounds.
The geometry covers clustered and stacked bars, line and area blank policies,
pie and doughnut wedges, matched scatter points, and radar polygons.

**Non-obvious choices.** Sparse caches retain logical indexes instead of being
densified. Gap, Zero, and Span are projected without allocating an arbitrary
declared count. Domain normalization scales before subtraction so opposite
finite extremes cannot create nonfinite geometry.

**Deviations from the design plan.** Three microscope passes hardened sparse
blank handling, aggregate overflow, nonfinite scatter categories, and adjacent
finite-coordinate cases. The approved API and HLD impact did not change.

**Spec sections touched.** `docs/hld/03-architecture.md` and
`docs/hld/09-charts-spec.md`.

**Tests.** `bar_chart_rasterises_at_computed_positions`,
`bar_geometry_handles_direction_grouping_gap_and_overlap`,
`line_scatter_and_radar_emit_paths_and_markers`,
`pie_doughnut_and_area_emit_closed_paths`,
`sparse_cache_indexes_preserve_slots_and_scatter_pairing`,
`finite_extremes_never_produce_nonfinite_geometry`, and the integrated full
gate. Deterministic raster checks passed with the new
`rpptx-chart` to `oxml-layout` dependency pointing inward.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep plot geometry independent of package
resolution and final theme colours. F-127 and F-128 own those boundaries.

### F-126, Chart rendering: axes, gridlines and labels

**Sprint.** S31
**Completed.** 2026-08-10
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `render_chart` now adds nice-number value axes, category
axes, gridlines, tick marks, deterministic glyph labels, point-level data
labels, and legends around the shared plot geometry. Radar charts use radial
spokes, perimeter labels, and concentric value grids.

**Non-obvious choices.** Annotation expansion is bounded at 16,384 logical
slots. Label anchors derive from clipped family geometry and retain direction
for degenerate bars and zero-radius radar points. Parsed point overrides affect
rendering while their original raw `c:dLbl` subtrees remain the sole
serialization source.

**Deviations from the design plan.** Seven microscope passes hardened scale
selection, explicit and reversed bounds, sparse joins, effective number
formats, short geometry, radar annotations, and exact direction and margin
coverage. The approved public surface and HLD impact did not change.

**Spec sections touched.** `docs/hld/09-charts-spec.md`.

**Tests.** `zero_to_one_hundred_axis_uses_expected_ticks`,
`nice_number_ticks_cover_unpinned_extents`,
`axes_gridlines_and_tick_marks_follow_model_state`,
`labels_and_legend_shape_with_deterministic_fonts`,
`inside_and_outside_label_positions_follow_family_geometry`,
`radar_annotations_use_spokes_perimeter_labels_and_radial_gridlines`,
`point_label_overrides_render_without_changing_preserved_xml`,
`labelled_chart_raster_is_deterministic`, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve the shared plot rectangle and z-order
of gridlines, clipped plot, axes, ticks, legend swatches, and text. Keep every
text run on the caller-provided deterministic `FontManager`.

### F-127, Chart colour resolution

**Sprint.** S32
**Completed.** 2026-08-10
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Presentation charts now resolve direct series fill and
line paint before a mapped accent1 through accent6 cycle. The effective theme,
colour map, theme-slot transforms, series transforms, and alpha flow through
every geometry family, marker, data label, and legend swatch.

**Non-obvious choices.** Direct `a:noFill` remains transparent. Filled area and
radar paths compose their established 55 percent policy with resolved alpha.
Only the selected concrete theme slot is required, so unrelated theme entries
cannot block a direct series colour.

**Deviations from the design plan.** None. Three microscope passes hardened
transparent filled plots, unused theme slots, and the order of theme-slot and
series transforms.

**Spec sections touched.** `docs/hld/09-charts-spec.md`.

**Tests.** `unstyled_four_series_use_accent_one_through_four`,
`direct_series_solid_colour_overrides_theme_accent`,
`series_accent_cycle_repeats_after_six`,
`series_colours_honor_colour_map_and_transform_order`,
`unsupported_direct_series_paint_is_contextual`,
`resolved_chart_palette_raster_is_deterministic`,
`authored_chart_enters_renderer_deterministically`, and the integrated full
gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep presentation chart colour resolution on
the exact DrawingML pipeline. The deliberately naive Word tint and shade helper
remains outside this path.

### F-128, Preserved chart fallback

**Sprint.** S32
**Completed.** 2026-08-10
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Presentation package rendering now resolves chart
relationships in slide, layout, and master scope. Supported charts enter the
ordinary renderer as frozen backend-neutral groups. Unsupported charts use a
compatible immediate cached picture or a labelled placeholder with a stable
diagnostic.

**Non-obvious choices.** AlternateContent and ChartML raw bytes remain the only
serialization source. Cached preview admission matches sniffed MIME and backend
capabilities, accepts only 8-bit three-component JPEG, and caps encoded,
decoded, and PNG inflation storage at 16 MiB before decoding. Integration tests
and the corpus driver call the same package-rendering function.

**Deviations from the design plan.** Seven microscope passes hardened missing,
malformed, corrupt, mismatched, sparse, oversized, grayscale, and CMYK cached
previews. They also tightened schema-positioned chart projection, immutable raw
payload access, source-scoped resolution, and production-path test coverage.
The approved routing and HLD scope did not change.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/06-presentationml-model.md`,
`docs/hld/07-inheritance-and-resolution.md`,
`docs/hld/08-rendering-spec.md`, and `docs/hld/09-charts-spec.md`.

**Tests.** `three_dimensional_chart_uses_cached_image_and_diagnostic`,
`authored_chart_relationship_enters_presentation_renderer`,
`same_chart_relationship_id_is_scoped_to_its_source_part`,
`unsupported_chart_without_preview_keeps_labelled_bounds`,
`missing_or_external_chart_relationship_is_contextual`,
`chart_choice_and_picture_fallback_remain_byte_preserved`,
`non_chart_choice_with_descendant_chart_uri_remains_opaque`,
`supported_and_fallback_charts_render_deterministically`, and the integrated
full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Treat cached chart previews as untrusted package
media. Keep admission bounded and aligned with every backend before allowing a
preview to suppress the visible labelled fallback.

### F-047, Packaging include and size gate

**Sprint.** S32.1
**Completed.** 2026-08-11
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `oxml-layout` is now a publication candidate with an
explicit package boundary containing its source, all 20 bundled TTFs, three
family licence files, and the Caladea notice. CI packages the crate, verifies
the archive, compares its exact inventory, and rejects archives above 10 MiB.

**Non-obvious choices.** The gate compares the package list to an exact sorted
inventory instead of checking only globs. This makes a missing legal file and
an accidental extra asset equally visible.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/15-build-and-toolchain.md`, "Packaging"
and "CI job matrix".

**Tests.** `cargo package -p oxml-layout --list`, verified workspace packaging,
the exact inventory assertion for 20 TTFs and four legal files, the 3,596,626
byte archive-size assertion, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep `NOTICE-Caladea` beside the full Apache
licence. The archive ceiling applies to the compressed `.crate`, while the
inventory check proves the required uncompressed assets are present.

### F-048, Automate split-family release preparation

**Sprint.** S32.1
**Completed.** 2026-08-11
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Cargo release metadata now prepares the inherited stable
rdocx train and the explicit incubating shared and PowerPoint train as separate
version groups with `v{{version}}` and `rpptx-v{{version}}` tag templates.
Preparation consolidates each version change while publication, tags, pushes,
and README replacements remain disabled.

**Non-obvious choices.** Stable packages use cargo-release's effective
`workspace` group because they inherit the root version. The 12 incubating
packages use the named `incubating` group because their versions are explicit.

**Deviations from the design plan.** Microscope pass 1 corrected the stable
metadata to cargo-release's effective `workspace` group. The approved family
boundary and external-action restrictions did not change.

**Spec sections touched.** `docs/hld/15-build-and-toolchain.md`, "Release
process".

**Tests.** Cargo-release 1.1.3 configuration assertions, the workflow
regression suite, a disposable stable preparation from 0.4.1 to 0.4.2, a
disposable incubating preparation from 0.0.0 to 0.1.0, manifest and lockfile
diff inspection, `cargo metadata --no-deps`, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** `/release` remains the sole authority for real
release tags and publication. Cargo-release prepares reviewed version commits
only.

### F-049, Extend publish.yml to the extracted workspace

**Sprint.** S32.1
**Completed.** 2026-08-11
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The publication workflow now accepts stable `v*` and
incubating `rpptx-v*` tags, preflights the exact 19-package publishable union,
and routes each namespace to its own explicit dependency-ordered allowlist.
The release command validates and creates only the requested family tag after
the reviewed-SHA and separate final-approval gates.

**Non-obvious choices.** The workspace dry run patches all 19 internal
dependencies to reviewed local sources. Cargo otherwise rewrites packaged path
dependencies to crates.io, where the reserved incubating 0.0.0 packages expose
no API. The patches verify the source graph without entering any archive or
weakening archive verification.

**Deviations from the design plan.** Microscope pass 1 added the missing
incubating `/release` authority and stronger exact workflow mutations. The
integrated packaging gate then disproved Cargo's assumed automatic local
staging, so the plan was corrected to require the exact local patch set and a
third clean microscope pass.

**Spec sections touched.** `docs/hld/11-migration-plan.md`, "Release tooling",
and `docs/hld/15-build-and-toolchain.md`, "Publishing" and "Release process".

**Tests.** Twenty-one workflow regression tests including swapped predicates,
extra and missing packages, a missing local patch, `continue-on-error`, and
successful fallback mutations. The locally patched
`cargo publish --workspace --dry-run` verified all 19 candidates without an
upload, every archive remained below 10 MiB, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep local patches on the dry-run preflight
only. Real dependency-ordered publish commands stay bare and wait for each
registry layer before publishing its consumers.

### F-050, CI matrix additions

**Sprint.** S32.1
**Completed.** 2026-08-11
**Size.** S, estimated 1 day, actual 1 day

**What was built.** CI now runs the `oxml-layout` no-default-features path, the
supported `rdocx-wasm` target check, and separate prose and generated-skill
drift checks. Every workspace all-feature test, lint, docs, and MSRV command
excludes the two Python extension packages.

**Non-obvious choices.** `rpptx-wasm` remains absent until F-138. The PyO3
exclusions are carried on every all-feature job because extension-module test
binaries cannot link against host Python symbols on Linux.

**Deviations from the design plan.** Microscope pass 1 found missing binding
exclusions on the clippy and docs jobs. The remediation made the exclusion
contract uniform across all all-feature jobs.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, "CI matrix", and
`docs/hld/15-build-and-toolchain.md`, "CI job matrix".

**Tests.** Current and MSRV workspace checks, exact CI command inspection,
`cargo test -p oxml-layout --no-default-features`,
`cargo check --target wasm32-unknown-unknown -p rdocx-wasm`, prose and skill
sync checks, and the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Add `rpptx-wasm` to the target job only when
F-138 creates that package. Keep the binding exclusions synchronized across
every new all-feature job.

### F-X005, Tag rpptx-v0.1.2

**Sprint.** S32.2
**Completed.** 2026-08-11
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete 12-crate incubating family is published at
0.1.2 under the `rpptx-v0.1.2` tag. Every selected manifest has non-empty
release metadata, and the publication workflow runs a self-contained metadata
regression before its archive checks and dependency-ordered crates.io uploads.
The matching GitHub release targets the reviewed sprint commit.

**Non-obvious choices.** The immutable `rpptx-v0.1.0` tag remains the partial
publication that contains only `oxml-core` 0.1.0. The immutable
`rpptx-v0.1.1` tag remains the CI-only failed recovery. A new 0.1.2 family was
required because release tags and published registry versions are never moved
or overwritten.

**Deviations from the design plan.** None. The approved 0.1.2 recovery ran as
designed after a fresh full verification, clean sprint review, and separate
final approval.

**Spec sections touched.** `docs/hld/03-architecture.md`, "Version trains",
`docs/hld/14-development-backlog.md`, "F-X005, Tag rpptx-v0.1.2", and
`docs/hld/15-build-and-toolchain.md`, "Packaging" and "Release process".

**Tests.** The targeted 0.1.2 metadata regression, all workflow regressions,
`cargo metadata --no-deps`, the exact patched 19-package publication dry run,
archive size and bundled asset assertions, supply-chain checks, the full
workspace gate, and all 28 output hashes passed. GitHub Actions run 31496676517
published all 12 packages and created the release. Independent `cargo info`
and owner checks confirmed every 0.1.2 registry entry under `mantissaman`, and
the annotated GitHub tag resolved to commit
`27a8bb8aa494759568d40bf66c167c214e759500`.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Released rdocx consumers may now cut over to
the 0.1.2 shared crates without local registry patches. The stable rdocx family
was not published by this tag.

### F-015, rdocx-oxml becomes a facade

**Sprint.** S32.2
**Completed.** 2026-08-11
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `rdocx-oxml` now re-exports the shared `oxml-core` XML,
raw XML, unit, and property implementations while preserving the established
Word-facing module paths. Five duplicate source files were removed without
call-site changes.

**Non-obvious choices.** The namespace facade retains the Word-specific
constants beside shared namespace helpers. This preserves the public surface
without moving Word vocabulary into the format-neutral crate.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/11-migration-plan.md`, "Consumer
cutovers", and `docs/hld/14-development-backlog.md`, "F-015".

**Tests.** Focused `rdocx-oxml` tests, dependency direction and package checks,
the integrated workspace gate, and all 28 output hashes passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Add format-neutral XML primitives to
`oxml-core`. Keep Word namespace vocabulary in the facade.

### F-016, Length re-export

**Sprint.** S32.2
**Completed.** 2026-08-11
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `rdocx::Length` is now the shared `oxml_core::Length` type.
The duplicate Word implementation was deleted, and every existing constructor,
accessor, conversion, and caller continues through the retained public path.

**Non-obvious choices.** The crate re-exports the type directly instead of
wrapping it. This preserves type identity and keeps the conversion behavior in
one implementation.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/11-migration-plan.md`, "Consumer
cutovers".

**Tests.** Shared unit conversion tests, focused rdocx checks, package
verification, the integrated workspace gate, and all 28 output hashes passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Use `oxml_core::Length` internally and retain
`rdocx::Length` as the compatibility import path.

### F-022, rdocx-opc deprecation shim

**Sprint.** S32.2
**Completed.** 2026-08-11
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `rdocx-opc` is now a deprecated exact re-export shim over
`oxml-opc`. The rdocx library, CLI, and WASM consumer use the shared crate
directly, construct Word packages explicitly, and expose the shared OPC error
type through the high-level error surface.

**Non-obvious choices.** Word-specific new-document setup remains in rdocx.
The shared OPC crate owns generic package mechanics and does not gain a reverse
dependency on a document format.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`, "WebAssembly",
`docs/hld/11-migration-plan.md`, "Consumer cutovers", and
`docs/hld/15-build-and-toolchain.md`, "Published crate graph".

**Tests.** Shared error identity, new-document graph, CLI and WASM checks,
package verification, dependency direction, the integrated workspace gate,
and all 28 output hashes passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** New consumers should depend on `oxml-opc`.
Retain `rdocx-opc` only for compatibility through the next stable transition.

### F-027, rdocx adopts oxml-media

**Sprint.** S32.2
**Completed.** 2026-08-11
**Size.** M, estimated 2 days, actual 1 day

**What was built.** rdocx now uses `oxml-media` for byte-first image format
detection, collision-safe media naming, MIME resolution, and downstream media
extraction. Duplicate local media helpers were removed. Mislabelled JPEG bytes
now produce a JPEG part, content type, and relationship target regardless of
the supplied filename extension.

**Non-obvious choices.** When loaded content-type defaults conflict with the
sniffed bytes, rdocx writes a per-part override. This changes only the new media
part and preserves the existing package default for other parts.

**Deviations from the design plan.** Microscope pass 1 found that replacing a
loaded package default could relabel unrelated existing parts. The remediation
used a per-part override and added a loaded-package regression before pass 2
returned clean.

**Spec sections touched.** `docs/hld/03-architecture.md`, "Crate boundaries",
`docs/hld/04-opc-and-packaging.md`, "Media", `docs/hld/11-migration-plan.md`,
"Consumer cutovers", and `docs/hld/14-development-backlog.md`, "F-027".

**Tests.** The exact mislabelled-JPEG package regression, loaded-package
content-type preservation, naming regressions, package verification,
dependency direction, the integrated workspace gate, and all 28 output hashes
passed.

**Hash harness.** Unchanged. All 28 integrated entries match. The intentional
metadata behavior is covered by the focused package regression because the
harness does not inspect that part name or content type.

**Notes for future sessions.** Resolve media metadata from bytes first. Treat
caller extensions as hints and preserve unrelated loaded-package defaults.

### F-028, add_picture_auto

**Sprint.** S32.2
**Completed.** 2026-08-11
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The rdocx write API now provides `add_picture_auto`, which
probes image metadata and inserts the native EMU dimensions using a 72 DPI
caller default. Invalid or unsupported image data returns a typed error before
any document mutation.

**Non-obvious choices.** The method computes dimensions first and delegates a
successful insertion to the existing explicit-size path. This keeps numbering,
relationships, and drawing construction in one implementation.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`, "Native image
sizing", and `docs/hld/14-development-backlog.md`, "F-028".

**Tests.** Exact 72 DPI extent and round-trip checks, declared and fallback DPI
cases, atomic failure coverage, explicit-size regressions, package verification,
the integrated workspace gate, and all 28 output hashes passed.

**Hash harness.** Unchanged. Existing samples continue to use explicit sizes.

**Notes for future sessions.** Keep the convenience API additive and preserve
the probe-before-mutation boundary.

### F-046, rdocx layout and PDF cutover

**Sprint.** S32.2
**Completed.** 2026-08-11
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rdocx-layout` retains the Word flow model while converting
its result into shared `oxml-layout` pages, elements, fonts, media IDs, and
diagnostics. `rdocx-pdf` is now an exact deprecated shim over `oxml-pdf`, and
the high-level render path uses the shared backend. Duplicate neutral layout,
font, media, and PDF backend sources and bundled font assets were removed.

**Non-obvious choices.** A concrete conversion function is the only boundary
between Word flow layout and shared output. It preserves the pre-cutover Word
glyph slices and line-height behavior while using shared output types and
renderers.

**Deviations from the design plan.** Initial integration exposed four Word PNG
deltas. The converter was corrected to preserve the established wrap and line
height semantics, returning all 28 hashes to baseline. Microscope pass 1 then
found an empty-image omission regression, which was restored before pass 2
returned clean. Sprint review pass 4 found that distinct image bytes could
overwrite each other when their compact `MediaId` values collided. Collision
resolution now assigns deterministic alternate IDs, and a forced-collision
regression covers both inline and anchored images. Sprint review pass 5 then
found repeated registry construction for each image occurrence. The final path
builds the relationship and media maps once per layout and reuses them through
paragraphs, tables, headers, footers, footnotes, shapes, and pagination.
Sprint review pass 6 found that the lower-level public entry points could not
share that private result. `MediaRegistry` is now their common public argument,
so direct callers retain both collision-resolved IDs and their image bytes.

**Spec sections touched.** `docs/hld/03-architecture.md`, "Rendering
boundaries", `docs/hld/08-rendering-spec.md`, "Word conversion boundary",
`docs/hld/11-migration-plan.md`, "Consumer cutovers", and
`docs/hld/15-build-and-toolchain.md`, "Packaging".

**Tests.** Exact conversion cases, layout and PDF suites, both no-default
layout paths, WASM, dependency direction, archive inventory, the integrated
workspace gate, and all 28 output hashes passed.

**Hash harness.** Unchanged. All 28 integrated entries match after preserving
the established Word conversion semantics.

**Notes for future sessions.** Keep format-specific flow logic in
`rdocx-layout`. Add neutral output and backend behavior to the shared crates.

### F-051, CHANGELOG and migration notes

**Sprint.** S32.2
**Completed.** 2026-08-11
**Size.** S, estimated 1 day, actual 1 day

**What was built.** A root `CHANGELOG.md` now documents the Unreleased stable
rdocx cutover, the published shared 0.1.2 family, retained facades, deprecated
shims, breaking surfaces, and automatic picture sizing. The README crate table
names the shared replacements and links the migration notes.

**Non-obvious choices.** One migration table covers every moved or deprecated
crate. This keeps version and compatibility guidance in a single durable
artifact.

**Deviations from the design plan.** Sprint review pass 4 found three retained
`rdocx-layout` breaking changes missing from the migration notes. The completed
table now documents the shared `MediaRegistry` argument on lower-level layout
and pagination, `AnchoredContent::Image` media ID, and `ParagraphBlock::jc`
alignment type. Sprint review pass 7 corrected those function references to
their retained `engine`, `table`, and `paginator` module paths.

**Spec sections touched.** None. The documentation reflects the completed HLD
contract without changing system intent.

**Tests.** Exact migration-path and version assertions, rustdoc with warnings
denied, prose checks, the integrated workspace gate, and all 28 output hashes
passed.

**Hash harness.** Unchanged. Documentation does not affect generated output.

**Notes for future sessions.** Keep the stable rdocx train under Unreleased
until its own release workflow runs. Do not imply that `rpptx-v0.1.2` published
the stable family.

### F-129, oxml-py-support

**Sprint.** S33
**Completed.** 2026-08-12
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A new unpublished `oxml-py-support` crate provides ordered
Word `ContentPath` and `PathSeg` values, `RevisionCounter`, a concrete Rust
`StaleElementError`, and canonical positive and negative `Length` conversions
delegated to `oxml-core`.

**Non-obvious choices.** The shared crate owns stale-domain classification but
accepts caller-supplied recovery guidance. Package-specific wording and Python
exception inheritance remain in the consuming binding. Presentation path
variants remain deferred until F-136 has a concrete consumer.

**Deviations from the design plan.** The approved plan was revised to include
`docs/hld/03-architecture.md` after its crate summary assigned the `Length`
pyclass incorrectly. Microscope pass 1 added caller-owned recovery guidance.
Pass 2 added `docs/hld/15-build-and-toolchain.md`, release metadata, and the
updated workspace-version package count.

**Spec sections touched.** `docs/hld/03-architecture.md`, "Three families, one
workspace", `docs/hld/10-bindings-spec.md`, "The chosen design", "The
invalidation problem, handled loudly", and "Python API shape",
`docs/hld/14-development-backlog.md`, "F-129, oxml-py-support" and "F-132,
Python enums, units and exceptions", and `docs/hld/15-build-and-toolchain.md`,
"Release process".

**Tests.** `stale_path_reports_both_revisions`, current-revision acceptance,
revision bumping, ordered Word paths, positive and negative Length truncation,
release-family metadata, focused crate checks, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep this crate format-neutral. Package bindings
own recovery paths and Python exception classes. Add presentation path variants
only with F-136's concrete consumer.

### F-130, rdocx-py core

**Sprint.** S33
**Completed.** 2026-08-12
**Size.** L, estimated 4 days, actual 1 day

**What was built.** A new unpublished mixed `rdocx-py` package exposes
`Document()`, `Document(path)`, lazy paragraph and run collections, path-only
handles, Python indexing, slicing and iteration, structural mutation, byte
round trips, and named stale-handle failures.

**Non-obvious choices.** `PyDocument` owns the Rust document while every child
handle stores only a Python document reference and an F-129 content path.
Immutable run reads use total facade accessors and preserve both layout caches.
Revision counters advance only after successful structural mutations.

**Deviations from the design plan.** The approved plan was revised to include
release-family metadata, its count regression, and
`docs/hld/15-build-and-toolchain.md`. Microscope pass 1 added the documented
optional path constructor and immutable cache-preserving run accessors. F-130
kept only the temporary stale exception bridge required by its gate, and F-132
replaced it with the final hierarchy. The consolidated gate upgraded PyO3 to
the first fixed 0.29.0 release after two RustSec advisories blocked completion.

**Spec sections touched.** `docs/hld/03-architecture.md`, "Facade conventions",
`docs/hld/10-bindings-spec.md`, "The chosen design" and "The invalidation
problem, handled loudly", and `docs/hld/15-build-and-toolchain.md`, "Release
process".

**Tests.** `stale_paragraph_after_structural_removal_raises_named_error`,
`constructor_accepts_an_optional_input_path`,
`immutable_run_accessors_preserve_cached_layout`, lazy collection coverage,
total facade accessors, byte round trips, 31 installed-package tests, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep lazy handles path-only and re-resolve them
on every operation. Read-only binding access must stay on immutable facade
methods so it cannot invalidate layout caches.

### F-131, rdocx-py formatting and tables

**Sprint.** S33
**Completed.** 2026-08-12
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Path-only font and paragraph-format subhandles expose the
bounded S33 formatting inventory with tri-state clearing. Lazy table, row,
cell, and nested paragraph handles expose table and cell formatting through
total public facade accessors.

**Non-obvious choices.** Binding-only underline variants use a bounded
integer-code facade API so the published exhaustive Rust `UnderlineStyle` enum
remains compatible. Signed Python indentation clearing is separate from the
established Rust helper. Cell text replacement invalidates nested handles with
exactly one revision bump.

**Deviations from the design plan.** There was no approved scope or HLD
deviation. Microscope remediation restored exhaustive underline and legacy
indent semantics, added single-bump cell invalidation and full tri-state
clearing coverage, made unrepresentable table justification and automatic font
colour read as `None`, proved rejected underline values do not mutate state,
and supplied complete recovery paths for stale nested paragraphs.

**Spec sections touched.** `docs/hld/03-architecture.md`, "Facade conventions",
`docs/hld/10-bindings-spec.md`, "Python API shape", and
`docs/hld/14-development-backlog.md`, "F-131, rdocx-py formatting and tables".

**Tests.** `unset_run_bold_is_none`, `none_clears_direct_formatting`,
`facade_table_and_tristate_accessors_are_total`,
`established_underline_enum_and_first_line_indent_remain_compatible`,
`cell_text_replacement_invalidates_nested_run_and_font`, table reopen tests,
the installed binding suite, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve public exhaustive enums and use checked
binding-only integer accessors for a wider Python literal set. Structural
replacement must bump the revision exactly once, and nested stale errors must
name the complete recovery path.

### F-132, Python enums, units and exceptions

**Sprint.** S33
**Completed.** 2026-08-12
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Pure-Python immutable `Length` integer subclasses,
`RGBColor`, four bounded `IntEnum` families, compatibility import paths,
top-level exports, and the public `RdocxError` hierarchy now ship in the mixed
package. Rust package, XML, stale, and layout failures map to their exact
concrete Python classes.

**Non-obvious choices.** Pure-Python value types preserve the Python 3.9
limited ABI while native conversion helpers retain truncation toward zero. A
direct `oxml-layout` dependency exists only under dev dependencies so a private
Rust test can construct the concrete layout failure and prove exact mapping.

**Deviations from the design plan.** There was no product or HLD scope
deviation. Microscope remediation added exact `LayoutError` mapping sensitivity
and its inward, test-only `oxml-layout` dependency, then documented the added
dependency-graph rider in the plan.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`, "Python API shape",
and `docs/hld/14-development-backlog.md`, "F-132, Python enums, units and
exceptions".

**Tests.** `alignment_center_and_inches_match_python_contract`,
`length_is_an_int_with_unit_properties`, fractional truncation, exact enum
values and docs, exception hierarchy,
`layout_error_maps_to_the_exact_public_layout_error_class`, installed abi3
package tests, dependency-direction checks, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep Python value and enum types outside the
native ABI boundary. Preserve the direct `oxml-layout` edge as dev-only unless
a separately designed production dependency requires it.

### F-133, rdocx-py rendering with allow_threads

**Sprint.** S33
**Completed.** 2026-08-12
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `Document.to_pdf`, `to_bytes`, `render_page_to_png`, and
`render_all_pages` convert Python arguments before releasing the GIL, perform
only Rust-owned work while detached, then rebuild Python bytes, lists, and
mapped exceptions after reattachment.

**Non-obvious choices.** The concurrency gate uses independent uncached
nontrivial documents and compares equivalent serial and parallel work. It
validates complete PDFs and extracted semantics through exactly pinned Poppler
26.01.0 instead of treating cache timing or byte identity as the oracle.

**Deviations from the design plan.** An approved plan revision added the
pinned Poppler differential-test oracle without changing implementation or HLD
scope. Review remediation added semantic equivalence and completeness checks,
progress-sensitive GIL gates for all four methods, exact version parsing for
both Poppler tools, and rejection of suffixed unreviewed versions.

**Spec sections touched.** None. The implementation fulfills the existing
binding threading contract without changing architectural intent.

**Tests.** `four_concurrent_to_pdf_calls_are_faster_than_serial`,
`poppler_pdf_oracle_is_available_at_reviewed_version`,
`poppler_version_pin_rejects_unreviewed_suffix`, the additional
`releases_gil_for_python_worker` gates, result and error mapping tests, 31
installed binding tests, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Every blocking native Python method needs a
progress-sensitive GIL-release regression. Keep correctness comparisons
outside timing assertions, require complete outputs, and pin external semantic
oracles to an exact reviewed version.

### F-134, Type stubs and py.typed

**Sprint.** S34
**Completed.** 2026-08-13
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Both mixed Python packages now ship hand-written native
extension stubs and zero-byte `py.typed` markers. Strict consumer programs and
live `stubtest` coverage describe lazy handles, collections, units, enums,
optional values, and factory-only construction exactly.

**Non-obvious choices.** Python 3.12 runs exact `mypy==2.3.0` because that mypy
release no longer supports Python 3.9. The generated extensions retain the
cp39-abi3 floor and were installed separately before the typing gates.

**Deviations from the design plan.** Review remediation narrowed rpptx shape
and length types, included every inline-typed module in strict checking, and
made all non-root native handles statically non-constructible.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`, Packaging,
`docs/hld/12-testing-strategy.md`, Python bindings, and
`docs/hld/14-development-backlog.md`, F-134.

**Tests.** Fresh rdocx and rpptx wheels contained the expected stubs and
markers. Strict mypy passed seven rdocx and six rpptx sources. Stubtest passed
six rdocx and five rpptx modules. The integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Change a native signature and its hand-written
stub in the same story, then run both strict installed-wheel checks and live
stubtest.

### F-135, python-docx parity suite

**Sprint.** S34
**Completed.** 2026-08-13
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A tagged `python-docx==1.2.0` parity suite covers seventeen
documented examples within the delivered rdocx surface. Both libraries author
documents, both libraries read both outputs, and normalized paragraphs, runs,
formatting, tables, cells, units, and enums must agree.

**Non-obvious choices.** The held-row Quickstart example performs one declared
public re-fetch after a structural cell write. This preserves strict global
revision invalidation while keeping the remaining sixteen example bodies as
namespace-only substitutions.

**Deviations from the design plan.** Review expanded the manifest from sixteen
to seventeen tagged examples, distinguished relative line spacing from length
spacing, and moved table-style coverage into both saved writer paths.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** The documented-example gate, bidirectional saved round trip, oracle
pin, manifest, line-spacing, and table-style mutation gates passed against a
fresh installed cp39-abi3 wheel. The integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the parity manifest bounded and tied to a
tagged upstream source. Compare public structure, never package bytes or XML.

### F-136, rpptx-py

**Sprint.** S34
**Completed.** 2026-08-13
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The unpublished `rpptx-py` mixed package exposes lazy
presentation, slide, shape, text, and table handles over the Rust facade. It
ships Python-compatible lengths, the bounded shape enum, mirrored exceptions,
and seven Getting Started examples with two-way `python-pptx==1.0.2`
structural comparison.

**Non-obvious choices.** Every handle and collection is path-only and carries
one captured global revision. Successful structural mutation invalidates all
previous views, including the mutating receiver. Recovery messages are derived
from the concrete repeated-shape path.

**Deviations from the design plan.** The documented examples gained the
minimal required public re-fetches after structural writes. Review also fixed
the omitted placeholder-index default, shape value 51 compatibility, complete
writer-drift sensitivity, and nested recovery paths.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Ten installed binding tests, six shared path tests, 103 rpptx tests,
the seven-example bidirectional differential, exhaustive stale-view probes,
WASM isolation, dependency trees, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Do not add owner references or revision
bypasses to preserve source compatibility. Re-fetch explicitly after every
structural write.

### F-137, wheels.yml

**Sprint.** S34
**Completed.** 2026-08-13
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A pinned GitHub Actions workflow builds both distributions
as cp39-abi3 wheels for six approved platform targets, plus one source
distribution per package. It validates compatible wheels in fresh
environments and reserves PyPI trusted publication for successful `py-v*` tag
runs in the `pypi` environment.

**Non-obvious choices.** Only the final publication job receives
`id-token: write`. The reviewed workflow has a raw-byte SHA-256 attestation in
addition to structural semantic tests, so any unreviewed workflow-byte change
fails closed before its release graph can be trusted.

**Deviations from the design plan.** Review hardened the workflow contract
against 155 non-vacuous matrix, execution, permission, action-pin, artifact,
trigger, and publication mutations. Native wheels and source distributions
were built locally, while the first real hosted cross-platform run remains
future GitHub evidence as planned.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Workflow contract and mutation tests, two native wheels, two source
distributions, clean imports, strict typing, stubtest, archive inventory, and
the integrated full gate passed. No tag, dispatch, or publication occurred.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Any intentional workflow edit must update the
semantic contract and its reviewed raw-byte digest in the same reviewed story.

### F-138, PR-time Python job

**Sprint.** S34
**Completed.** 2026-08-13
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Pull requests now run one fail-fast-disabled two-package
matrix that creates a fresh environment, builds each extension with maturin,
installs exact test and oracle dependencies, and runs the complete package
pytest directory.

**Non-obvious choices.** The job uses exact immutable action commits, root
`contents: read` permission, no OIDC authority, and direct failure propagation.
Its Python 3.12.9 runtime supports the exact pytest 9.1.1 pin while exercising
the cp39-abi3 extensions.

**Deviations from the design plan.** Review added structural trigger,
permission, action-input, step-order, and failure-suppression checks. An
inherited prose violation in an F-137 review artifact was fixed separately
before the integrated gate.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, Python bindings,
and `docs/hld/15-build-and-toolchain.md`, CI job matrix.

**Tests.** Twenty-eight workflow regressions, thirty-three installed rdocx
tests, ten installed rpptx tests, a real failing-test propagation mutation,
WASM and dependency isolation, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the pull-request job least privilege and
make any new package test failure propagate without a conditional or fallback.

### F-139, Rewrite rdocx-wasm

**Sprint.** S35
**Completed.** 2026-08-13
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `WasmDocument` now owns one real `rdocx::Document` and
delegates its established JavaScript surface to the facade. DOCX byte round
trips retain the complete package, ordered text comes through the additive
facade getter, and generated Node tests exercise the actual JavaScript byte
boundary. The native Word and presentation paths retain system-font discovery
through explicit feature forwarding, while the WASM graph disables it and
keeps bundled fonts available.

**Non-obvious choices.** Workspace rendering dependencies are defaults-off so
each concrete consumer selects its graph. Native consumers opt into
`system-fonts`, while `rdocx-wasm` does not. The wrapper keeps the existing
JavaScript names, maps concrete facade errors to string-valued `JsValue`s, and
does not maintain a second package model or add byte aliases.

**Deviations from the design plan.** Microscope remediation added exact
generated-JavaScript reflection, restored presentation system-font defaults,
and strengthened the root workspace manifest sensitivity. The approved facade
ownership, public surface, and HLD impact did not change.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/13-risks-and-open-questions.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `document_with_images_headers_and_numbering_round_trips_every_part_intact`,
`document_text_preserves_body_and_table_order`,
`wasm_round_trip_preserves_the_complete_package_in_node`, native and WASM
feature-contract mutations, no-default gates, package and publication checks,
and the integrated full gate at `fecfd0a` passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep the WASM package as a facade consumer.
Feature isolation depends on defaults-off workspace edges and explicit native
opt-ins, while bundled font bytes remain unconditional.

### F-140, wasm CI job

**Sprint.** S35
**Completed.** 2026-08-13
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The pull-request WASM job now target-checks both facade
wrappers with the locked graph and runs both inline Node suites. It installs
exact Node 24.11.1 and wasm-pack 0.15.0, and a structured workflow contract
enforces package coverage, step order, immutable action pins, least privilege,
and ordinary failure propagation.

**Non-obvious choices.** The operative setup-node pin is the reviewed v6.5.0
commit. Both Node suites run as separate commands without conditions,
`continue-on-error`, fallback success, or listing-only substitutions. The
incubating cargo-release preparation group includes unpublished `rpptx-wasm`,
while the crates.io allowlist remains limited to published packages.

**Deviations from the design plan.** Microscope remediation corrected the
setup-node release label in the testing HLD and added independent provenance
sensitivity. The approved two-package workflow and HLD scope did not change.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `test_wasm_pr_job_checks_both_targets_and_runs_node_tests`,
`test_wasm_pr_job_rejects_skipped_or_weakened_gates`, setup-node provenance
and release-family mutations, locked wasm32 checks, both complete Node suites,
a real propagated Node failure, and the integrated full gate at `fecfd0a`
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep action annotations aligned with their
immutable commits. Any new WASM package must receive an executable Node gate,
not only a target check.

### F-141, to_pdf in the browser

**Sprint.** S35
**Completed.** 2026-08-13
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `WasmDocument.toPdf` now returns bytes directly from the
normal `rdocx::Document::to_pdf` facade. A generated-JavaScript Node regression
adds text through the public binding, calls `toPdf` reflectively, and requires
a complete PDF with a Type 0 font, an embedded TrueType stream, and bundled
Carlito.

**Non-obvious choices.** There is no deterministic alias or WASM-only renderer.
The wrapper's defaults-off graph makes the normal facade path browser-safe by
excluding host discovery while retaining unconditional bundled fonts.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `to_pdf_in_node_returns_a_complete_pdf_with_an_embedded_bundled_font`
proved the generated `toPdf` name, `Uint8Array` boundary, `%PDF-` through
`%%EOF`, `/Subtype /Type0`, `/FontFile2`, and Carlito. The two-test Node suite,
feature-isolation contract, font and rendering riders, mutation sensitivity,
and the integrated full gate at `fecfd0a` passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Browser PDF behavior belongs to the normal
facade method under the WASM feature graph. Do not fork layout or PDF assembly
inside the binding.

### F-142, rpptx-wasm

**Sprint.** S35
**Completed.** 2026-08-13
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The workspace now contains an unpublished `rpptx-wasm`
package backed by one real `rpptx::Presentation`. Its bounded default profile
constructs, opens, serializes, counts, and mutates presentations without the
renderer. The opt-in `render` profile adds only `toPdf`. Package-to-render
assembly moved from the corpus example into the owning facade, and the example
now delegates to that single deterministic path.

**Non-obvious choices.** Native `rpptx` defaults retain template, rendering,
and system fonts, while the wrapper selects a bundled-template-only facade and
adds deterministic rendering explicitly. The exact optimized normal-default
artifact is 519,060 decimal gzip bytes, below the 1,000,000-byte gate. The
size check binds reviewed wasm-pack, wasm-opt, and deterministic gzip arguments
to the freshly built artifact.

**Deviations from the design plan.** Microscope remediation bound the size gate
to the current artifact, hardened normal-default and render feature contracts,
and strengthened facade-to-example parity and rendering completeness. The
approved profiles, facade boundary, and HLD impact did not change.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `default_profile_is_under_one_megabyte_and_round_trips_a_deck`,
`wasm_presentation_uses_the_real_facade_in_node`,
`render_profile_returns_a_complete_pdf`, facade-to-example render parity,
default and render dependency graphs, exact 519,060-byte size and mutation
sensitivity, publication riders, and the integrated full gate at `fecfd0a`
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep rendering optional for the wrapper and
package interpretation in the facade. Re-run the exact optimized size gate
whenever the normal-default dependency graph changes.

### F-143, oxml-cli-support

**Sprint.** S36
**Completed.** 2026-08-13
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The publishable, currently unpublished
`oxml-cli-support` crate owns bounded
one-based range parsing, default output-path construction, and the schema-one
JSON envelope shared by both command-line tools. `rdocx-cli` now uses the
shared output-path and JSON helpers without changing its command surface.

**Non-obvious choices.** Range parsing charges requested expansion work before
deduplication and accepts exactly 100,000 requested values. This prevents large
or overlapping ranges from amplifying memory or CPU work while retaining
sorted and deduplicated results.

**Deviations from the design plan.** Review added the explicit materialization
and cumulative-work bounds, the exact accepted boundary, and full compatibility
coverage for every existing rdocx inspect field and default conversion path.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Seven shared helper tests, two rdocx-cli compatibility tests,
oversized and overlapping range mutations, the 21-package publication dry run,
dependency-direction checks, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep format-neutral command plumbing in this
crate and charge range work before materializing user input.

### F-144, rpptx-cli

**Sprint.** S36
**Completed.** 2026-08-13
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The publishable, currently unpublished `rpptx-cli` binary
provides `inspect`,
`text`, `convert`, `diff`, `replace`, `validate`, and `render`. It consumes the
real presentation facade and shared CLI support, preserves package content and
run formatting during replacement, and uses deterministic rendering for PDF
and PNG output.

**Non-obvious choices.** Raster commands reject more than 8,000,000 pixels per
page. PNG conversion preflights every page and then renders, writes, and drops
one encoded page at a time. Text diff retains its established LCS behavior but
rejects matrices above 1,000,000 cells before allocation.

**Deviations from the design plan.** Review added complete core metadata to
plain inspect output, zero-slide PNG failure, bounded DPI and diff resources,
streaming multi-slide PNG output, and verified corpus provision for both clean
CI jobs.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Fourteen command integrations, all 50 pinned corpus decks, the
same-run, cross-run, grouped, and table replacement matrix, resource-boundary
mutations, deterministic rendering checks, workflow regressions, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep OOXML ownership in the facade. New CLI
operations must remain bounded before allocating or collecting output.

### F-145, rpptx-cli thumbnail and outline

**Sprint.** S36
**Completed.** 2026-08-13
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rpptx thumbnail` renders slide one to a deterministic PNG
at exactly 320 pixels wide with proportional height. `rpptx outline` prints
each slide title once and recursively emits textual paragraphs in shape order
with two spaces per paragraph level.

**Non-obvious choices.** Outline title suppression compares the actual
`ShapeRef` node identity through additive `PartialEq` and `Eq` implementations.
This remains total for field-only titles and does not depend on collapsed
placeholder indexes. Paragraph line breaks normalize to printable spaces.

**Deviations from the design plan.** Review exposed unindexed placeholder and
field-only title cases. The user approved the bounded equality trait addition,
and the owning facade HLD was added to the exact work list before completion.

**Spec sections touched.** `docs/hld/06-presentationml-model.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Fourteen CLI integrations cover portrait aspect ratio, output-path
precedence, grouped text, paragraph levels, embedded breaks, unindexed
placeholders, field-only titles, all 50 corpus decks, targeted mutations, and
the integrated full gate.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Compare facade handles by their defined node
identity when exact shape suppression is required. Do not infer identity from
placeholder indexes or text.

### F-146, npm publication

**Sprint.** S36
**Completed.** 2026-08-13
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The CI WASM job can build local bundler packages, run
`npm pack`, and install `@tensorbee/rdocx-wasm` and
`@tensorbee/rpptx-wasm` into separate fresh consumers. The packages contain
their scoped metadata, WebAssembly binary, JavaScript glue, and public type
declarations. No registry publication path or authority was added.

**Non-obvious choices.** Both manifests use wasm-opt 125 with `-Oz`,
`--enable-bulk-memory`, and `--enable-nontrapping-float-to-int`. CI downloads
the reviewed official Binaryen asset and verifies its SHA-256 before use.
Fresh installs disable scripts, audits, and funding calls.

**Deviations from the design plan.** Actual Rust output required the approved
third wasm-opt feature flag. The user also approved installing exact wasm-opt
125 because wasm-pack otherwise falls back to a different bundled version.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Local bundler builds, two scoped tarballs, separate fresh installs,
package inventories, installed imports, both locked WASM checks, both Node
suites, 36 workflow regressions, dependency isolation, and the integrated full
gate passed. No package was published.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Treat real npm publication as a separate
reviewed story with explicit authority. Keep this path local and install-only.

### F-X001, rdocx-cli tests

**Sprint.** S36
**Completed.** 2026-08-13
**Size.** M, estimated 2 days, actual 1 day

**What was built.** One integration binary invokes all seven `rdocx-cli`
commands through `std::process::Command` with isolated in-code fixtures. The
text command now uses facade document-order text, and both render branches use
the bundled-font deterministic facade.

**Non-obvious choices.** The tests bind visible font output byte-for-byte to
the deterministic renderer for both selected-page and all-page paths. Their
temporary workspaces combine process identity with a local counter.

**Deviations from the design plan.** Review exposed interleaved body-order and
system-font rendering defects in the product. The approved plan was revised to
include those bounded command fixes and exactly three HLD files.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** Seven command integrations, misspelled-command and false-validation
mutations, interleaved paragraph and table text, deterministic selected and
all-page rendering mutations, golden PNG checks, and the integrated full gate
passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep one command integration entrypoint and
bind rendering tests to the deterministic facade rather than host coincidence.

### F-X002, README example correctness

**Sprint.** S36
**Completed.** 2026-08-13
**Size.** S, estimated 1 day, actual 1 day

**What was built.** All six root README Rust examples compile as `no_run`
rustdoc tests. The table example uses the real indexed row and cell APIs. One
canonical Python runner obtains the exact locked rdocx rlib through Cargo JSON
and invokes rustdoc with warnings denied.

**Non-obvious choices.** README remains the only snippet source. The CI docs
job and canonical full verification call the same runner, which prevents drift
without duplicating examples into crate documentation.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The runner compiled six examples, a disposable `rows()` and
`cells()` mutation failed with E0599, output scans remained clean, CI and
generated-adapter contracts passed, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Keep README as the snippet source and discover
the actual locked rlib rather than assuming a target filename.

### F-X003, Deduplicate the sample generators

**Sprint.** S36
**Completed.** 2026-08-13
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The obsolete `generate_samples` example was deleted.
`generate_all_samples` is now the single source for every document and output
consumed by the hash and golden-image harnesses.

**Non-obvious choices.** This is a behavior-neutral deletion. The surviving
generator was not rewritten, and no baseline was recorded or moved.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** None.

**Tests.** All 28 deterministic hashes, all seven golden PNG buffers, every
example compile, the canonical seven-sample inventory, a missing-contract
mutation, repository invocation search, and the integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Add sample artifacts only through
`generate_all_samples` and prove the full deterministic inventory.

### F-X004, Fix the shared temp path in the test suite

**Sprint.** S36
**Completed.** 2026-08-13
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The rdocx file round-trip integration uses an output name
containing the test process ID, so concurrent test processes do not share one
fixed path.

**Non-obvious choices.** The regression asserts the exact process identity in
the filename. No production helper or new dependency was introduced for a
single test-isolation correction.

**Deviations from the design plan.** None. Microscope pass 1 was clean.

**Spec sections touched.** None.

**Tests.** The exact test failed under the former fixed-name mutation, two
concurrent invocations passed, the complete rdocx suite passed, and the
integrated full gate passed.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Test files that can be created concurrently
must include process identity or use an isolated workspace.

### F-X006, Tag the expanded rpptx family

**Sprint.** S37
**Completed.** 2026-08-14
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete 14-package shared and PowerPoint family is
published on crates.io at 0.1.3 through `rpptx-v0.1.3`. The release adds
`oxml-cli-support` and `rpptx-cli` to the earlier 12-package family while
keeping unpublished `rpptx-wasm` in the 0.1.3 local preparation group.

**Non-obvious choices.** The release used the smallest fresh patch version
above immutable 0.1.2. The annotated tag peels to reviewed sprint SHA
`805680ab8a6dadd4d4247471a81cbb21b88a3196`. The workflow published only the
14-package incubating allowlist and created the matching GitHub release. No
npm package was published. The user gave separate final approval at that
reviewed SHA immediately before the sprint branch push, annotated tag
creation, tag push, and publication workflow.

**Deviations from the design plan.** Full verification exposed stale release
prose that required font assets in both `rdocx-layout` and `oxml-layout`.
The reviewed correction names `oxml-layout` as the sole owner of 20 TTF files
and four legal files, forbids duplication in `rdocx-layout`, and retains the
`rpptx` default presentation asset check.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Full verification passed at the reviewed release SHA. GitHub Actions
run `31762653847` completed successfully for tag `rpptx-v0.1.3`, including the
incubating publication and GitHub release jobs. All 14 packages resolve from
crates.io at 0.1.3, and every owner check reports `mantissaman (Atul Sharma)`.
The remote annotated tag peels to the reviewed SHA.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve the immutable 0.1.2 and 0.1.3 tags.
Future Rust-family publication must use a fresh version and a separately
approved `/release` invocation. npm publication remains unauthorized.

### F-X007, Integrate PR 25 and stable crate documentation

**Sprint.** S38
**Completed.** 2026-08-14
**Size.** L, estimated 4 days, actual 1 day

**What was built.** PR 25's custom-list, hyperlink, hard-break, and fixed-table
authoring APIs are integrated with side-effect-free rejection and synchronized
table geometry. All seven stable crates now carry package-specific README
documentation with twelve compile-checked Rust examples. Numbering mutation
preserves unmodelled XML through namespace-aware parsing, typed property
overlays, deterministic prefix allocation, and tab-stop occurrence provenance.

**Non-obvious choices.** The contributor's commits and GitHub credit remain
intact. The public numbering preservation fields use the approved breaking
pre-1.0 boundary for v0.5.0. Stable archive checks use the complete local patch
graph so every package is evaluated against the integrated workspace rather
than an older immutable registry dependency.

**Deviations from the design plan.** The reviewed remediation expanded from
the two initial hardening fixes and README work to a complete raw-numbering
preservation model. This was required to uphold the repository's unmodelled XML
contract, and the user approved the breaking pre-1.0 v0.5.0 boundary.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Canonical non-fast verification passed. Focused list rejection,
table geometry, raw-numbering round-trip, namespace projection, and bounded
tab-overlay gates passed. A safe table-width mutation made the named gate fail
before byte-identical restoration. The 21-package workspace dry run passed,
all seven stable archives included their intended README, and every archive
was below 10 MiB.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** F-X008 owns v0.5.0 preparation and publication
through the separate release workflow. No stable crate was published by this
story. Preserve Jon Stokes's `@jonstokes` credit in the PR and merge record.

### F-X008, Tag v0.5.0

**Sprint.** S38
**Completed.** 2026-08-14
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The exact seven-package stable rdocx family is published
on crates.io at 0.5.0 through `v0.5.0`. The eleven-package shared-version group
is coherent at 0.5.0 while unpublished WASM, Python, and support packages
remain outside the crates.io allowlist. The 15-package incubating preparation
group remains at 0.1.3, with its 14 published crates unchanged.

**Non-obvious choices.** The user gave separate final approval at reviewed SHA
`01bd2379097344120f5e1dba0c36882d95af88a6`. Annotated tag object
`5cbf51479ba0f8ae383684b57b2e7ca68eca01d4` peels to that exact SHA. Workflow
run `31815290384` published only the stable family. Stable publication job
`94815375298` succeeded, the incubating step was skipped, and GitHub Release
job `94817628637` succeeded. No incubating, WASM, Python, or npm package was
published.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/11-migration-plan.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The stable and incubating metadata preflights, release workflow
contract, README examples, exact patched 21-package dry run, archive inventory,
WASM checks, dependency graph, and `cargo deny` passed at the reviewed release
SHA. All seven 0.5.0 packages were downloaded independently from crates.io,
and every owner check reports `mantissaman`. The matching
[GitHub release](https://github.com/tensorbee/rdocx/releases/tag/v0.5.0)
targets the reviewed SHA. PR 25 contributor credit and its merge note remain
intact.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve the immutable 0.4.1 and 0.5.0 tags.
Future stable publication must use a fresh version and a separately approved
`/release` invocation. PyPI and npm publication remain unauthorized.

### F-X009, README coverage for every workspace crate

**Sprint.** S39
**Completed.** 2026-08-14
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Every one of the 26 Cargo workspace packages now explicitly
declares a distinct README. Eighteen focused crate-local documents were added,
and the existing package guides were audited and strengthened. Every README
states package purpose, direct-use guidance, neighbouring package relationships,
publication status, and a concrete example for its Rust, CLI, Python, or
JavaScript surface.

**Non-obvious choices.** The root README remains the high-level `rdocx` package
guide. The runner derives package and publication inventories from Cargo
metadata, then obtains primary and companion libraries from one Cargo build
graph. This keeps the `oxml-pdf` example bound to the exact `oxml-layout`
instance used by the renderer. Existing crates.io releases are immutable, so
new README pages appear there only when the affected crate receives a new
published version.

**Deviations from the design plan.** None. Microscope pass 1 found that eleven
initial examples showed dependency installation without demonstrating use.
The final implementation replaced them with real public API examples and
strengthened the exact gate before clean pass 2.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, README example
correctness. `docs/hld/14-development-backlog.md`, F-X009. `docs/hld/15-build-and-toolchain.md`,
workspace package READMEs in the docs job.

**Tests.** `python3 scripts/readme_doctests.py` validates 26 distinct declared
README sources, compiles 26 Rust examples across 20 library READMEs, validates
six CLI, Python, and JavaScript examples, and byte-compares the README in all
21 publishable archives with its declared source. A package-specific API
mutation failed the exact gate before byte-identical restoration. Canonical
non-fast verification passed, including changed-package tests, workspace tests,
WASM checks, warnings-denied rustdoc, and the README archive gate.

**Hash harness.** Unchanged. All 28 entries match.

**Notes for future sessions.** A dependency declaration is installation, not a
usage example. The README gate deliberately requires package-specific surface
text and compiles every applicable Rust block. Unpublished packages have
documentation but gain no publication authority.

### F-X010, Tag v0.6.0

**Sprint.** S39
**Completed.** 2026-08-14
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The eleven-package shared-version train moved coherently
to 0.6.0. The exact seven-package stable family, `rdocx-opc`, `rdocx-oxml`,
`rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and `rdocx-cli`, is
published on crates.io at 0.6.0. The four other train members remain
unpublished, and the incubating family remains at 0.1.3.

**Non-obvious choices.** The user gave separate immediate approval at reviewed
SHA `96cac2a9256351ad03ab3f9499fcc9ed5d48adf2`. Annotated tag object
`2279fd3b4a9183e458c2b7449e5714536c305dfd` peels to that exact SHA. Workflow
run `31830892682` published only the stable allowlist. Publication job
`94866033898` and GitHub Release job `94868199553` succeeded. No incubating,
WASM, Python, npm, or PyPI package was published.

**Deviations from the design plan.** None. Microscope pass 1 strengthened the
README archive gate to require the exact local patch set. Pass 2 reconciled the
two compatibility shims with the coherent stable release train. Pass 3 was
clean.

**Spec sections touched.** `docs/hld/11-migration-plan.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Full verification passed at the reviewed SHA, including all 38
workflow tests, 26 README sources, 26 Rust examples, 21 archive README checks,
the exact 21-package dry run, WASM checks, archive assets, and `cargo deny`.
All seven 0.6.0 packages download independently from crates.io under sole owner
`mantissaman`. Every crates.io README endpoint returns non-empty rendered HTML,
and the matching
[GitHub release](https://github.com/tensorbee/rdocx/releases/tag/v0.6.0)
targets the annotated tag at the reviewed SHA.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve the immutable `v0.6.0` tag. F-X011
owns the separate incubating 0.2.0 preparation and must obtain its own clean
review and immediate `/release rpptx-v0.2.0` approval. PyPI and npm publication
remain unauthorized.

### F-X011, Tag rpptx-v0.2.0

**Sprint.** S39
**Completed.** 2026-08-14
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The fifteen-package incubating preparation train moved
coherently to 0.2.0. The exact fourteen-package crates.io family is published
at 0.2.0. `rpptx-wasm` moved with the local train and remains unpublished. The
stable family remains at 0.6.0.

**Non-obvious choices.** The user gave separate immediate approval at reviewed
SHA `1b13dbe4a5454f1d1629ff8915287b26daa10ed0`. Annotated tag object
`0d9ce33258988377751d7f10fec43e0096f014d0` peels to that exact SHA. Workflow
run `31836554504` published only the incubating allowlist. Publication job
`94884015713` and GitHub Release job `94887859113` succeeded. No stable, WASM,
Python, npm, or PyPI package was published.

**Deviations from the design plan.** None. Microscope pass 1 and all three
sprint-review passes were clean.

**Spec sections touched.** `docs/hld/03-architecture.md`, versioning.
`docs/hld/14-development-backlog.md`, F-X011.
`docs/hld/15-build-and-toolchain.md`, the incubating release family and release
process.

**Tests.** Full verification passed at the reviewed SHA, including all 38
workflow tests, 26 README sources, 26 Rust examples, 21 archive README checks,
the exact 21-package dry run, WASM checks, archive assets, and `cargo deny`.
All fourteen 0.2.0 packages download independently from crates.io under sole
owner `mantissaman`. Every crates.io README endpoint returns non-empty rendered
HTML, and the matching
[GitHub release](https://github.com/tensorbee/rdocx/releases/tag/rpptx-v0.2.0)
uses the annotated tag that targets the reviewed SHA.

**Hash harness.** Unchanged. All 28 integrated entries match.

**Notes for future sessions.** Preserve the immutable `rpptx-v0.1.3` and
`rpptx-v0.2.0` tags. Future incubating publication requires a fresh version and
a separately approved `/release` invocation. PyPI and npm publication remain
unauthorized.

### F-X012, Restore pinned CI toolchains

**Sprint.** S40
**Completed.** 2026-08-15
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Hosted CI now provisions checksum-bound Poppler 26.01.0
and LibreOffice 26.2.5.2 through shared installers with bounded streaming
extraction and exact runtime identity checks. The broad Test and MSRV jobs pin
uv 0.10.2, isolate its cache, use an 8 MiB test-thread stack, and run on Ubuntu
24.04 with the complete LibreOffice runtime package set. The WASM job validates
the official Binaryen 125 Linux identity after checksum verification.

**Non-obvious choices.** Each installer refuses an already populated target
prefix, so a version-looking binary cannot bypass provenance checks. Poppler
builds only the three required tools from reviewed source. LibreOffice installs
the checksum-bound core and Impress packages together with thirteen explicit
Ubuntu runtime packages. Both paths fail closed on unsafe archives, resource
ceilings, missing package members, or wrong executable identities.

**Deviations from the design plan.** Hosted validation exposed two additional
clean-runner requirements after the original Poppler and Binaryen correction.
The approved plan was extended to pin uv and the stack budget, then to install
the exact LibreOffice build and its Ubuntu runtime libraries. Nine microscope
passes hardened the installer and workflow mutation contracts before the final
clean review.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, pinned rendering
oracles and hosted gates. `docs/hld/14-development-backlog.md`, F-X012.
`docs/hld/15-build-and-toolchain.md`, deterministic CI tool installation.

**Tests.** Six focused workflow tests and all 46 workflow regressions pass.
They exercise installer provenance, checksums, streaming member and byte
ceilings, unsafe entries, exact runtime identities, required packages, job
ordering, failure propagation, and successful short-circuit mutations. Hosted
pull-request run `31853529961` passed all 14 jobs at reviewed commit
`e96217f88b9dfd4612913787bc736f3627f73092`, including all 421 presentation
fidelity slides and the LibreOffice viewer gates. Canonical `/verify --full`
passes from the clean sprint tree, including exact 21-package dry-run archives,
README examples, WASM targets, and supply-chain checks.

**Hash harness.** Unchanged. All 28 entries match.

**Notes for future sessions.** Keep the external tool versions, source URLs,
checksums, archive bounds, runtime identities, and consumer-job assertions in
one reviewed contract. A moving package-manager binary or a preinstalled tool
is not equivalent evidence. The temporary hosted-validation pull request was
closed without merge and its remote branch was deleted.

### F-X013a, Footnote line advance

**Sprint.** S41
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Footnote and endnote text drawn at the bottom of a page now
advances across the line instead of drawing every segment at the same x. A note
built from more than one run, which is what any note carrying mixed formatting
or a hyperlink produces, was previously an unreadable stack of overprinted
words. It now reads as a line of text. The same change made a second defect
visible and fixed it: notes were line-broken at the full content width but drawn
one marker indent to the right, so every note line overran the right margin by
exactly that indent. A single `FOOTNOTE_INDENT` constant now feeds both the
break width and the draw position, so the two cannot disagree.

**Non-obvious choices.** The advance covers all four `LineItem` variants, not
just the two that draw. A tab or an inline image inside a note is still not
rendered, but it occupies width, and skipping its advance would pull everything
after it to the left. The match is exhaustive, so a new variant fails to compile
here rather than silently reintroducing the defect.

The right-margin fix was taken into this story rather than deferred. It is one
constant and one subtraction, it shares a root cause with the defect the story
exists to fix, and deferring it would have shipped a story whose stated outcome
is legible notes while leaving them running off the page.

**Deviations from the design plan.** Two. The plan asserted that non-text line
items do not advance in the body path, and that was wrong. `paginator.rs`
advances for tabs and images, and the first microscope pass caught the claim.
The plan also predicted a hash harness delta and there is none, for a reason
worth recording: no corpus document contains a footnote at all, so the harness
never exercises this code path.

**Spec sections touched.** None. The story fixes a defect in rendering a
construct the spec set already describes, and adds no surface.

**Tests.** The gate is the pair of named regressions,
`a_multi_segment_footnote_does_not_stack_its_segments_at_one_x` and
`a_single_segment_footnote_keeps_its_original_position`. Three more were added:
`footnote_segment_advance_matches_body_segment_advance`,
`a_tab_inside_a_footnote_still_advances_the_text_after_it` for the defect the
first review pass found, and `a_long_footnote_does_not_overrun_the_right_margin`
for the width fix. Each was proven to fail against its own reverted code. The
single-segment test is an intentional guard that passes both before and after.

**Hash harness.** Unchanged. All 28 entries match, and that result carries no
information about this story. None of the seven corpus documents contains a
footnote, so `render_page_footnotes` is entirely unexercised by the harness. The
evidence for this story is its regression tests plus an end-to-end render of
`sample1.docx`, the document the external contribution used for its own before
and after screenshots. Exactly one of that document's eight pages changed, the
one carrying the footnote.

**Notes for future sessions.** The harness blind spot matters beyond this
story. F-X013b and F-X013c will both report a flat 28 of 28 for the same reason,
and that must not be read as those stories having no output effect. Closing the
gap means adding a corpus document with notes, which changes the baseline set
and is its own decision. The remaining visible defect on that sample page is
body text overlapping the note area, which is exactly what F-X013b addresses.
The note path still ignores its paragraph's own indent and justification, which
the body path honours. Two placement routines that drift apart is what produced
this defect in the first place, and one shared routine is the durable answer.

### F-X013b, Footnote reservation and splitting

**Sprint.** S41
**Completed.** 2026-08-16
**Size.** L, estimated 3 days, actual 1 day

**What was built.** Notes are laid out once, before pagination, into a new
`NoteRegistry`, and the paginator reserves, splits and draws them from that
single source. Pagination previously filled a page with body text knowing
nothing about the note area, which a post-pagination pass then drew straight
over the top of it. Body text and notes no longer collide. A note too tall for
the room left continues on the following page without repeating its marker, and
a page that opens with carried note content draws the full-width continuation
rule rather than the short one. The post-pagination pass is gone, so the note
placement that is reserved and the note placement that is drawn are now the same
computation rather than two that could disagree.

The note stream model changed to support it. `CT_Footnote` carries a
`NoteType` read from `w:type`, separators are retained rather than dropped, and
`w:type` is written back. Opening and saving a document previously deleted its
separator definitions outright.

**Non-obvious choices.** Note markers are shaped in the registry, before
pagination, because the paginator holds only `&FontManager` and shaping needs
`&mut`. Pre-shaping is what lets note placement live in the paginator at all.

A note's cost is priced without being claimed. A paragraph is measured before
anyone knows which page it lands on, so `available_height_for` prices its notes
and `claim_notes` runs only where lines are actually placed. Claiming during
measurement stranded notes on the page before their own reference.

The note area is measured from `ink_bottom`, where the body's last mark sits,
not from `cursor_y`, which includes trailing paragraph spacing that collapses at
a page break.

Separator identity follows `w:type`, with an untyped id of 0 or below still
read as a separator. The ids separators conventionally use are a convention, not
a rule, and `sample1.docx` puts its `continuationSeparator` at id 1, where the
old id-based test read it as note number one.

**Deviations from the design plan.** Two plan claims were wrong and were
corrected in the plan. The plan said this story fixes notes being positioned
against the final section's geometry: positioning is fixed, line breaking is
not, and the remainder is filed as F-X017. The plan also did not anticipate
retaining separators in `CT_Footnotes::footnotes`, which changes what that field
contains. `Document::footnotes()` gained a filter so its public behaviour is
unchanged.

**Spec sections touched.** `docs/hld/03-architecture.md`, "What stays put",
updated to say note placement belongs to the paginator rather than to a
post-pagination pass.

**Tests.** The gate is the three named regressions:
`a_page_whose_body_fills_the_text_area_does_not_overlap_its_notes`,
`a_note_taller_than_its_remaining_space_continues_on_the_next_page`, and
`a_page_referencing_one_note_twice_reserves_it_once`. Also added
`a_continued_note_draws_the_continuation_separator`,
`an_oversized_note_still_leaves_room_for_body_text`,
`a_note_is_drawn_on_the_page_that_carries_its_reference`, and in `rdocx-oxml`
`a_separator_definition_survives_open_and_save`,
`get_by_id_does_not_return_a_separator`,
`note_types_are_read_through_a_foreign_prefix` and
`an_unknown_note_type_reads_as_a_normal_note`. Each was proven to fail against
its own reverted change.

**Hash harness.** Unchanged. All 28 entries match, and that result carries no
information about this story, for the reason F-X013a recorded: no corpus
document contains a note. A delta here would have been a genuine surprise. The
evidence is the regression set plus an end-to-end render of `sample1.docx`,
where page 5 stops overprinting its table of contents.

**Notes for future sessions.** Two defects in this work were invisible to the
tests as first written and only surfaced by sweeping the reference across every
paragraph position and comparing the note's page against the reference's page.
A note drifting one page from its reference is not something a fixed-position
test finds. That sweep is now
`a_note_is_drawn_on_the_page_that_carries_its_reference` and it is the single
most valuable test in the set. The end-to-end render also caught a third
regression that no unit test saw: keying the registry by note id alone let an
endnote overwrite a footnote sharing its number, silently swapping the rendered
text. Telling the two streams apart is F-X013c.

### F-X013c, Endnotes at the document end

**Sprint.** S41
**Completed.** 2026-08-16
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Endnotes stop rendering at the foot of the page carrying
their reference and are emitted after the last body page, flowing from the top
of their own pages. Footnotes keep the page foot. An endnote reference now costs
its page no height at all.

Underneath, a note reference carries which stream it came from. `TextSegment`
and `GlyphRun` swap `footnote_id: Option<i32>` for `note: Option<NoteRef>`,
where a `NoteRef` is a stream and a number. `NoteRegistry` keys on that pair, so
a document numbering a footnote and an endnote alike keeps both. It previously
kept one and silently dropped the other, which `sample1.docx` triggers with a
footnote 2 and an endnote 2.

**Non-obvious choices.** Endnotes begin on a fresh page rather than continuing
on the last body page, which is what Word does when there is room. An endnote
flowing onto a page that also owes footnotes would put two note regions on one
page competing for the same height, and that interaction is not worth its
complexity here. Recorded in the design plan so the choice is legible rather
than accidental.

`draw_note` was extracted so the page foot and the document end share one
drawing routine. Two placement routines that drift apart is exactly what
produced the F-X013a defect, and this story would otherwise have created a
second pair.

Endnote markers keep the raw id. Word defaults endnotes to lower roman numerals
through `w:endnotePr/w:numFmt`, which is a numbering-format concern rather than
a placement one and was not taken.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`, "What stays put",
extended to describe the two note streams being placed differently and keyed
apart.

**Tests.** The gate is
`a_footnote_and_an_endnote_sharing_a_number_render_their_own_text` and
`endnotes_render_after_the_last_body_page`. Also added
`footnotes_and_endnotes_keep_their_own_regions` and
`an_endnote_reference_does_not_reserve_space_at_the_page_foot`, which pins that
an endnote changes body pagination not at all. Each was proven to fail against
its own reverted change: three against the stream split, and the shared-number
regression against id-only registry keying, where the endnote text vanishes
entirely.

**Hash harness.** Unchanged. All 28 entries match, for the reason F-X013a
recorded: no corpus document contains a note. The evidence is the regression set
plus an end-to-end render of `sample1.docx`, where all eight body pages stay
byte identical, the footnote keeps its own text at the foot of page 5, and a
ninth page appears carrying the endnote's own distinct text.

**Notes for future sessions.** All eight body pages being byte identical is the
useful signal here, and it is worth reaching for whenever a story carves one
behaviour out of another. It says the change was additive far more directly than
any assertion about the new behaviour does. The public surface of `oxml-layout`
changed: `footnote_id` became `note`. That crate is incubating at 0.2.0 with no
consumer outside this workspace, and `rdocx`'s own public API is untouched,
since `RunRef::footnote_id()` reads the oxml model rather than a layout segment.

### F-X014, Kashida justification values

**Sprint.** S41
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `ST_Jc` accepts `lowKashida`, `mediumKashida` and
`highKashida`, mapping each to justified alignment. The symptom was larger than
the backlog entry described and the entry was corrected before implementing:
`CT_PPr::from_xml` propagates a rejected justification with `?`, and that error
reaches `Document::open`, so a document carrying any of the three failed to open
at all rather than losing one property.

**Non-obvious choices.** The three values join the existing `both | justify`
arm rather than gaining a variant of their own. Kashida justification stretches
Arabic text by elongating the connecting stroke rather than by widening spaces,
which needs shaping this crate does not do, so a distinct variant would behave
identically to `Both` at every site that matches on `ST_Jc` while adding a case
to each. `distribute` was rejected because it spreads the last line and kashida
justification does not.

A kashida value round trips as `both`. That is a deliberate normalisation
recorded in the design plan, not an oversight.

**Deviations from the design plan.** None. The backlog story itself was
corrected during design, once the failure was reproduced and turned out to be a
load failure rather than a layout inaccuracy.

**Spec sections touched.** None.

**Tests.** The gate is
`a_document_using_kashida_justification_still_opens`, which loads a document for
each of the three values and asserts the paragraph keeps both its justification
and a sibling property. Plus `kashida_justification_maps_to_both` and
`an_unknown_justification_is_still_rejected`, the latter pinning that the check
was widened rather than removed. All three fail against the unwidened parser.

**Hash harness.** Unchanged. All 28 entries match. No corpus document carries a
kashida value, and no existing behaviour moved, since affected documents
previously failed to open rather than rendering differently.

**Notes for future sessions.** This is one instance of a wider problem, filed as
F-X018. Nine value parsers in `shared.rs` and `styles.rs` reject any string they
do not enumerate, and several are reached through `?` from property parsing, so
a document using a spec-valid value the model has not yet listed fails to open.
Fixing all nine means deciding a general rule, which is that an unmodelled value
falls back to the element's default and its siblings survive. That is a story of
its own rather than something to change in passing here.

### F-X015, Anchored drawing wrap and alignment model

**Sprint.** S41
**Completed.** 2026-08-16
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `WrapType` gains `Square`, `TopAndBottom`, `Tight` and
`Through` alongside `None`, and each wrapping element parses to its own mode in
both the empty and the expanded spelling. `CT_Anchor` reads the four text
distances and the `wp:align` child of `positionH` and `positionV`, and
`AnchoredDrawing` carries all of it into layout in points. Nothing reads the new
fields yet, so placement and rendering are unchanged. F-X016 consumes them.

Before this, `wrap` was parsed-but-dead: set to `None` at both construction
sites and read nowhere, while the serialiser wrote `wrapNone` unconditionally.
An alignment-positioned drawing landed at offset zero, because only the offset
was read.

**Non-obvious choices.** `Tight` and `Through` parse to their own variants
rather than collapsing into `Square`. They wrap to the drawing's outline rather
than its frame, F-X016 will approximate them as `Square`, and approximating is
the renderer's job. Collapsing at parse time would throw away information the
model cannot recover.

A zero text distance is not written. That began as a bug, see below, and the
resolution is right on its own terms: an absent attribute and a zero attribute
mean the same thing.

**Deviations from the design plan.** One, and the harness caught it. The plan
said the serialiser path runs only for a programmatically built anchor, which is
true, and missed that the sample generators are exactly that. Writing all four
distances unconditionally changed `report:word/document.xml` on the first
harness run. Corrected by omitting zero distances, and the plan now records both
the mistake and the fix.

**Spec sections touched.** None. F-X016 carries the HLD update for wrapping,
since that is where the behaviour appears.

**Tests.** The gate is the round-trip pair,
`an_anchor_round_trips_its_wrap_distances_and_alignments` and
`a_parsed_anchor_re_emits_its_original_bytes`. Plus
`every_wrap_element_parses_to_its_own_mode`,
`anchor_alignments_and_distances_are_read` and
`an_unknown_alignment_reads_as_no_alignment`. Proven against two separate
reverts, one of wrap parsing and one of the distance and alignment reads.

**Hash harness.** Unchanged, 28 of 28, which is this story's proof rather than a
formality. It did not start that way, and the delta is described above.

**Notes for future sessions.** The sample generators build anchors
programmatically, so a change to `CT_Anchor::to_xml` reaches the harness even
though a parsed anchor re-emits its captured `raw_xml` and never touches that
code. Any future change to an anchor serialiser should expect the same. More
generally, this is the story where declaring "expected unchanged" and being
wrong was useful: the prediction is what turned a silent byte change into a
question worth answering.

### F-X016, Floating drawing placement and text wrapping

**Sprint.** S41
**Completed.** 2026-08-16
**Size.** L, estimated 3 days, actual 1 day

**What was built.** Two behaviours the model has described since F-X015 but
nothing performed.

An anchored drawing positioned by an alignment now resolves against its
`relativeFrom` frame instead of landing at offset zero. A right-aligned drawing
sits at the right of its frame, a centred one at the midpoint.

Body text flows around a wrapping drawing. `wrapSquare` keeps text clear of the
frame plus its text distance on the lines the drawing spans, on the side the
drawing sits, and `wrapTopAndBottom` pushes the paragraph's content below the
drawing. `wrapTight` and `wrapThrough` are approximated as square, since
wrapping to an outline needs the `wp:wrapPolygon` the model does not carry, and
reserving the frame beats not wrapping at all.

**Non-obvious choices.** The line breaker gained per-line prefix and suffix
reservations rather than a wrapping concept of its own. An empty vector, the
default, reproduces existing behaviour exactly, which is what let every other
caller stay untouched and the harness stay flat.

Re-breaking needs the line breaking inputs to survive past layout, and those
hold the same shaped glyphs the laid-out lines hold. They are moved rather than
cloned, since `inline_items` is finished with at that point and would otherwise
be dropped, and `Engine::layout` drops them again unless the document holds a
drawing that wraps. A document without one carries nothing.

The reflow runs before the paragraph is measured, because a reflow changes its
height and measuring first would measure the wrong thing. Two passes, not a
loop: the second settles a drawing that only overlaps once the text has moved,
and a fixed count cannot fail to terminate.

**Deviations from the design plan.** One, and rendering the sample is what
caught it. The plan limited wrapping to drawings anchored to the current
paragraph or already placed, on the grounds that a later paragraph's position is
unknown, and accepted that as a documented limitation. `sample1.docx` shows the
limitation failing on the contribution's own headline page: its two arrows
flank one paragraph, but the right-hand arrow is anchored to paragraph 282 while
the text is in paragraph 280, so the left arrow wrapped and the right one kept
printing over the text. A bounded look-ahead now collects wrapping drawings from
following blocks whose vertical frame is the page or a margin, which have a
position independent of where their own paragraph lands. The residual case,
paragraph-relative anchors in later blocks, is filed as F-X019.

**Spec sections touched.** `docs/hld/03-architecture.md`, "What stays put",
extended to say the paginator reflows around wrapping drawings and why the
reflow inputs are carried conditionally.

**Tests.** The gate is the three golden tests,
`text_wraps_beside_a_left_aligned_square_drawing`,
`text_wraps_beside_a_right_aligned_square_drawing` and
`a_top_and_bottom_drawing_pushes_text_below_it`. Plus
`a_drawing_anchored_to_a_later_paragraph_still_pushes_text_aside` for the
look-ahead, `a_wrap_none_drawing_leaves_text_untouched` as the identity guard,
and two placement unit tests. Every one except the identity guard was proven to
fail against its own reverted change.

**Hash harness.** Unchanged. All 28 entries match, and here that is a real
result rather than the blind spot the note stories carried: the corpus does
contain floating drawings, they simply all use `wrapNone`, and every new path is
gated on a wrap mode other than `None`. The flat harness is what proves the
gating holds.

**Notes for future sessions.** Rendering the contributor's own document is what
turned an accepted limitation into a fixed defect. The unit tests all passed
with the limitation in place, because they anchor the drawing to the paragraph
it affects, which is the case the limitation covers. Real documents do not.
Where a story is motivated by a specific document, that document belongs in the
loop, not just the tests derived from it.

### F-X020, Refresh the dependency lockfile

**Sprint.** S42
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Sixteen semver-compatible dependency updates taken with
`cargo update`. `Cargo.lock` is the only product change: no manifest moved, so
no crate gained or lost a dependency and no API surface changed. None of the
sixteen was a security fix. `cargo audit` reports zero vulnerabilities across
152 dependencies before and after, and `cargo deny check` passes all four
sections, with `ttf-parser` RUSTSEC-2026-0192 remaining the single documented
exception rather than something this story cleared.

**Non-obvious choices.** The updates were taken together rather than
individually, because the point was to measure their combined effect once while
the isolation tooling was to hand, not to avoid measuring them.

**Deviations from the design plan.** The plan listed `zlib-rs` among the crates
with no path to rendered output. It does have one, through `flate2`, `png` and
`tiny-skia`. The conclusion survived, since PDF stream compression uses
`miniz_oxide` and that did not update, but the reasoning was wrong and is
corrected here rather than edited out of the plan.

**Spec sections touched.** None.

**Tests.** No new test, deliberately: there is no new behaviour to pin, and a
test asserting a version number would pin the lockfile rather than the
behaviour. The gate is the existing instrument, the full workspace suite at 53
binaries and zero failures, plus the 28-entry hash harness.

**Hash harness.** Unchanged, 28 of 28, **and that is not the whole answer.**

The refresh changed all seven sample PDFs. Every sample PNG stayed
byte-identical, which is why the harness stayed flat: it records `page1.png` and
three `word/*.xml` parts per sample and no PDF at all.

The delta was traced as the plan required, by reverting the lockfile and
applying suspects alone. `font-types 0.12.2 to 0.12.3` on its own moves all
seven PDFs, reaching the text shaper through `read-fonts 0.41.0` and `harfrust`.
It was then characterised with the repository's own pinned Poppler oracle before
being accepted: extracted text identical in 7 of 7 samples under `pdftotext`,
`pdfinfo` identical apart from the file size line, sizes moving by single-digit
bytes, and every PNG byte-identical. A serialisation-level difference in numbers
written to the content stream, with no semantic effect.

No baseline was re-recorded, because no recorded baseline moved.

**Notes for future sessions.** The durable finding is not the delta but that a
gate reported green while a first-class output changed across every sample. The
harness has no PDF coverage, so the `oxml-pdf` writer, its glyph positions,
embedded font subsets and compressed streams, can drift with nothing watching.
Filed as F-X021, which also has to decide what a stable PDF fingerprint is,
since raw PDF bytes carry a creation date and object ordering that need not be
reproducible. Until that lands, a dependency refresh should compare sample PDFs
by hand the way this one did.

### F-X024, Move the theme adapter into rdocx-oxml

**Sprint.** S42
**Completed.** 2026-08-16
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `impl From<&CT_OfficeStyleSheet> for Theme` moved from
`oxml-drawing` to `rdocx-oxml`, which owns `Theme`. The orphan rule permits it,
and the effect is that the dependency between the families now runs one way,
from the format crate to the shared crate, like every other cross-family edge.
`docs/hld/03-architecture.md` no longer documents an exception, because there is
none, and `no_shared_crate_depends_on_a_format_crate` keeps it that way.

**Why it was needed.** Scoping the two release stories exposed a cycle between
the publication trains. `rdocx-layout` depends on `oxml-layout` and
`oxml-drawing` depended on `rdocx-oxml`, and `publish.yml` publishes one train
per tag. That works only when the train going first depends on an
already-published version of the other, which S39 satisfied because only one
train moved. S41 broke both APIs, so both had to bump, and neither could go
first. Stable first will not compile, since `rdocx-layout` needs `oxml-layout`
0.3.0. Incubating first would have shipped an adapter bound to `rdocx-oxml`
0.6.0 while `rdocx-layout` 0.7.0 expected 0.7.0's `Theme`, putting two
semver-incompatible copies of the Word model in one graph and breaking the one
cross-family integration point.

**Non-obvious choices.** Moving beat deleting. The adapter has no caller in the
workspace today, so deleting it would have been the smaller diff and cost
nothing immediately. It is the documented bridge for when PresentationML themes
reach Word layout, and removing it would only have to be undone later on the
other side. The accepted cost is that `rdocx-oxml` now pulls `oxml-drawing`, so
a Word-only consumer compiles DrawingML.

`OFFICE_DEFAULT_XML` became public in `oxml-drawing` so the moved regression can
compare a projected `Theme` against one parsed from the same source. The
alternative, a dev-dependency from `oxml-drawing` back on `rdocx-oxml`, would
have rebuilt the edge the story exists to remove, since a dev-dependency still
has to resolve at publish time.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`, "The dependency rule"
and the adapter paragraph, plus the diagram edge. `CLAUDE.md`, the layout table.

**Tests.** The gate is the conversion regression, which moved with the impl and
still compares the projection against `Theme::from_xml` on the same source, plus
`no_shared_crate_depends_on_a_format_crate`. The invariant test discriminates:
adding `rdocx-oxml.workspace = true` to `oxml-layout` fails it with the
offending line named.

**Hash harness.** Unchanged. All 28 entries match, which is the expected result
for the same conversion code in a different crate.

**Notes for future sessions.** The invariant test cannot be exercised against
`oxml-drawing` itself. Reintroducing that exact edge now produces
`rdocx-oxml -> oxml-drawing -> rdocx-oxml`, a cargo cycle that fails to resolve
before any test runs. That is a stronger guarantee than the test, but a reader
proving the test works should edit a different `oxml-*` crate or they will get a
confusing resolver error instead of a clean assertion failure.

### F-X022, Tag rpptx-v0.3.0

**Sprint.** S42
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The incubating train moved 0.2.0 to 0.3.0 and was published.
S41 broke its public API rather than extending it: `oxml-layout` renamed
`TextSegment::footnote_id` and `GlyphRun::footnote_id` to `note`, changing the
type from `Option<i32>` to `Option<NoteRef>`, and added two `LineBreakParams`
fields. A 0.x minor bump is the correct response.

Fifteen packages were prepared and exactly fourteen published. `rpptx-wasm`
moved to 0.3.0 and remains `publish = false`.

**Release evidence.** All fourteen resolve from crates.io at 0.3.0 under owner
`mantissaman`: `oxml-core`, `oxml-opc`, `oxml-media`, `oxml-layout`,
`oxml-drawing`, `oxml-pdf`, `oxml-sml`, `oxml-cli-support`, `rpptx-oxml`,
`rpptx-chart`, `rpptx-layout`, `rpptx-render`, `rpptx`, `rpptx-cli`. The
annotated tag `rpptx-v0.3.0` dereferences to `ab52cd2`, the reviewed SHA.

**Non-obvious choices.** The incubating train published first, and after F-X024
that order is permanent rather than incidental. The stable crates depend on
`oxml-layout`, so 0.3.0 had to resolve on crates.io before the stable train
could publish. S39 released stable first because only one train moved that
sprint.

**Deviations from the design plan.** One, and it mattered. The first pass moved
every version carrier under `crates/` and stopped there, missing the
release-family preflight in `scripts/test_sprint_workflow.py` that
`publish.yml` invokes by name as its gate, and the `ci.yml` WASM literal.
Neither `cargo test` nor `/verify` runs the Python suite, so the gap passed
every local gate and would have failed in CI at publication. Fixed before
release and filed as F-X025.

**Spec sections touched.** None.

**Tests.** All 46 release regressions pass, including
`test_incubating_release_family_is_prepared_at_0_3_0`. Full workspace suite at
53 binaries and zero failures, README doctests, `cargo deny`, and the patched
21-package dry run with every archive under 10 MiB.

**Hash harness.** Unchanged, 28 of 28. A version string reaches no rendered
byte.

**Notes for future sessions.** The publication order is now fixed by the
dependency graph rather than by convention: incubating, then stable. F-X024 is
what makes that true, and reintroducing an `oxml-*` dependency on a format crate
would break it again.

### F-X023, Tag v0.7.0

**Sprint.** S42
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The stable train moved 0.6.0 to 0.7.0 and was published.
S41 broke its public API: `rdocx-oxml` added `note_type` to `CT_Footnote`, six
fields to `CT_Anchor` and four variants to `WrapType`, each of which breaks an
exhaustive match or a struct literal, and `rdocx-layout` added fields to
`ParagraphBlock` and `AnchoredDrawing`.

The `rdocx` facade's own API is unchanged. `Document::footnotes()` still returns
`Vec<(i32, String)>` and `RunRef::footnote_id()` is untouched, so a consumer of
the facade alone sees no break. Eleven packages were prepared and exactly seven
published.

**Release evidence.** All seven resolve from crates.io at 0.7.0 under owner
`mantissaman`: `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`,
`rdocx-pdf`, `rdocx`, `rdocx-cli`. The annotated tag `v0.7.0` dereferences to
`ab52cd2`, the same reviewed SHA as `rpptx-v0.3.0`. The four unpublished
packages, `oxml-py-support`, `rdocx-py`, `rdocx-wasm` and `rpptx-py`, inherited
0.7.0 without gaining publication authority.

**Non-obvious choices.** The stable train published second, because
`rdocx-layout 0.7.0` declares a dependency on `oxml-layout 0.3.0` and could not
have resolved before the incubating train landed.

**Deviations from the design plan.** The story was implemented before its design
plan was written, which is a workflow violation. The plan was written
afterwards and records what was done and the inventory that was taken.

**Spec sections touched.** None.

**Tests.** All 46 release regressions pass, including
`test_stable_release_family_is_prepared_at_0_7_0`. Full workspace suite, README
doctests, `cargo deny`, and the patched 21-package dry run.

**Hash harness.** Unchanged, 28 of 28.

**Notes for future sessions.** Both trains now sit one minor version apart from
where S41 left them, and the two tags share a SHA. A future release that moves
only one train is the normal case again, and only a sprint that breaks both
needs the ordering care this one did.

### F-X018, Unknown enumerated values must not fail a document open

**Sprint.** S43
**Completed.** 2026-08-16
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Twelve call sites across six files in `rdocx-oxml` stopped
propagating a rejected enumerated value out of document parsing. An unmodelled
value is now read as if the attribute were absent, which in OOXML means the
element's default, which is usually inheritance from the style chain. The
document opens and every sibling property survives.

Before this, nine value parsers returned an error for any string they did not
list, and those errors travelled through `?` out of `CT_Document::from_xml` to
`Document::open`. A document using a spec-valid value the model had not yet
enumerated did not open at all.

**Non-obvious choices.** The parsers stay fallible. `from_str` still returns
`Result`, so the tolerance is an explicit decision at each call site rather than
a property of the type, and a caller that wants strictness keeps it. That is
the same shape as `ST_OnOff::from_str_or_default`, which already existed.

`Option`-typed fields become `None` rather than a guessed variant. `None` means
"not specified" and lets the style chain supply the value. Falling back to a
concrete variant such as `ST_Jc::Left` would override a style that does specify
alignment, turning a missing value into an actively wrong one. The three
`borders.rs` locals are not `Option`, so they keep the default they were
initialised with.

No enum gained a variant. Guessing which unmodelled values matter is what
F-X014 did for the one case that was reachable, and the shape was the defect.

**Deviations from the design plan.** The plan cited
`docs/hld/04-opc-and-packaging.md` as the home of the prefix-tolerant read rule.
It lives in `docs/hld/03-architecture.md`. Corrected in the plan and the right
file updated.

**Spec sections touched.** `docs/hld/03-architecture.md`, the domain
conventions list, which gains a bullet stating that an unmodelled enumerated
value reads as an absent attribute, including the round-trip cost.

**Tests.** The gate is
`a_document_with_an_unmodelled_enumerated_value_still_opens`, which loads a
document for each of eight enumerations reachable from `document.xml` and
asserts a sibling property survives. Plus
`an_unmodelled_value_leaves_the_property_unset`, pinning `None` rather than a
guess, and `the_parsers_still_reject_an_unknown_value`, pinning that the check
was moved rather than removed. The first two fail against a single reverted call
site, naming the offending value.

**Hash harness.** Unchanged, 28 of 28. Every corpus document already opens, so
none carries an unmodelled value.

**Notes for future sessions.** Two things are worth knowing. `StyleType` is the
ninth enumeration and is not in the document-level regression, because it is
reached from `styles.xml` rather than `document.xml`, so a `CT_Document`
fixture cannot carry it. It is covered by the strictness unit test only.

And an unmodelled value is now silently lost on save, since the field is `None`
and the serialiser writes nothing. That is the accepted cost of opening the
document at all. Preserving it would need the `raw_xml` capture machinery that
unmodelled elements already use, extended to attributes.

### F-X017, Notes broken to their own section's width

**Sprint.** S43
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The note registry lays each footnote and endnote out once
per distinct section content width rather than once per document, and the
paginator looks a note up by the width of the section drawing it. A document
whose sections differ in page width now breaks each note to the measure of the
section holding its reference. Before this, every note was broken to the final
section's width and then drawn against whichever section carried the reference,
so the two agreed only when the sections shared a page size.

**Non-obvious choices.** The map is keyed on the note plus the content width in
raw bits, through `f64::to_bits`. Both the key and every lookup come from
`PageGeometry::content_width()` over the same `sectPr`, so this is exact
equality on a value computed the same way twice rather than a comparison
needing a tolerance. Repeated widths collapse, so the common single-geometry
document still lays each note out exactly once and no fast path is needed.

The once-before-pagination rule is kept. Laying notes out lazily during
pagination would need a mutable font manager inside the paginator, which
`notes.rs` deliberately does not have, and that is a much larger change than
this defect earns.

Endnotes are looked up at the final section's width, because they are emitted
after the last body page and drawn against that section's geometry wherever
their reference sits.

`NumberingState` gained `Clone` so a note laid out at several widths consumes
its list numbers once rather than once per width. Numbering does not depend on
width, so the state left behind is the state a single layout would have left.

**Deviations from the design plan.** The plan's risk routing recorded only the
layout row. Microscope pass 1 found the **Public API of a published crate** row
also matched, since `pub mod notes` makes both changed signatures public
surface of `rdocx-layout`. The plan now records the semver impact.

**Spec sections touched.** `docs/hld/03-architecture.md`, the `NoteRegistry`
paragraph, which gains the per-width rule and the endnote measure.

**Tests.** The gate is `a_note_is_broken_to_the_width_of_its_own_section`,
confirmed to fail against reverted code: registering only the final width makes
it report the same line count for both sections. Plus
`a_single_section_document_lays_notes_out_exactly_as_before`,
`an_endnote_is_broken_to_the_final_sections_width`, and three registry unit
tests covering distinct widths, repeated widths and an unregistered width.

**Hash harness.** Unchanged, 28 of 28 on the worker tree. No sample defines a
section break or a note, so no sample reaches either path.

**Notes for future sessions.** This is a **breaking change to a published
crate**. `NoteRegistry::build` takes `&[f64]` where it took `f64`, and
`NoteRegistry::get` takes the width as a second parameter. No caller outside
`rdocx-layout` exists in this workspace, which is exactly why it compiled
cleanly and had to be declared rather than observed. Under 0.x it is a minor
bump for the next `/release` to state.

A lookup at an unregistered width returns `None` and the note is not drawn. The
engine registers every width it paginates, so this cannot happen by
construction, and `an_unregistered_width_has_no_layout` pins the deliberate
choice not to silently substitute another width.

### F-X019, Paragraph-relative drawings in later blocks should wrap

**Sprint.** S43
**Completed.** 2026-08-16
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Text now flows around a wrapping drawing anchored to a
later paragraph even when the drawing is positioned relative to its own
paragraph. F-X016 did this only for drawings framed by the page or a margin,
because a paragraph-relative drawing has no vertical position until its own
paragraph is placed, and that cannot happen before the text above it is laid
out. A section holding such a drawing now paginates twice: the first pass
records where each one landed and on which page, and the second offers those
rectangles to the text above them.

**Non-obvious choices.** Two passes, and deliberately not a fixed point. The
second pass reflows earlier text, which can move the drawing's own paragraph, so
the rectangle it flowed around may be slightly stale. Iterating is not
guaranteed to terminate, because growing a paragraph can push a drawing to the
next page, which shrinks the paragraph, which pulls the drawing back. Two passes
give one answer, always, and `a_second_pass_is_stable_for_the_document_that_
earns_it` pins that the answer is the same answer every time.

The second pass is gated on a predicate over the blocks, so a document without
such a drawing paginates once and through code that is unchanged. That is every
sample and every corpus document today.

Drawings are keyed by block index and their index within that block, which is
stable across passes because both walk the same slice. The look-ahead offers a
resolved rectangle only when the recorded page matches the page being built, so
a drawing that landed overleaf does not push this page's text aside.

`PassContext` holds the six values both passes share, so the two calls differ in
one argument rather than eight.

**Deviations from the design plan.** The plan described the driver as taking
eight arguments. Microscope pass 1 recorded that as a smell, since the
alternative was an argument-count lint silenced with an `allow`, and the
remediation introduced `PassContext` instead. The plan's test table also grew
from four rows to seven during implementation.

**Spec sections touched.** `docs/hld/03-architecture.md`, the paragraph on
reflowing around floating drawings, which gains the two-pass rule and states the
limit that the passes are not iterated to a fixed point.

**Tests.** The gate is
`a_paragraph_relative_wrapping_drawing_pushes_earlier_text_aside`, confirmed to
fail against the unfixed look-ahead: with the second pass disabled the earlier
paragraph takes 18 lines against 18 and the assertion fires. Plus
`a_page_relative_drawing_in_a_later_block_still_wraps` guarding F-X016's case,
the stability regression, and four paginator unit tests covering the predicate,
the empty-map first pass, the page scoping and the recording.

**Hash harness.** Unchanged, 49 of 49 on the worker tree. No sample anchors a
wrapping drawing to its own paragraph, so the predicate is false for all seven.

**Notes for future sessions.** The vertical offset in the test helper means
different things per frame, which is why the paragraph case passes `-120.0` and
the page case `150.0`. A drawing that lands below every line of the paragraph
before it pushes nothing aside and proves nothing, which is how the first draft
of the page-relative control failed.

### F-X021, The hash harness should cover PDF output

**Sprint.** S43
**Completed.** 2026-08-16
**Size.** L, estimated 2 days, actual 1 day. Sized M at design time and revised
to L when the story found the PDF writer was not deterministic

**What was built.** The output-stability harness records three entries per
sample for the deterministic PDF, taking the manifest from 28 entries to 49.
`pdf/pages` covers the page count, each page's `/MediaBox` and each page's
inflated content stream in `/Kids` order. `pdf/resources` covers every other
inflated stream, which is the CID font subsets, the ToUnicode CMaps and the
image XObjects. `pdf/bytes` is the file digest. Before this the harness recorded
three `word/*.xml` parts and a page-one PNG per sample and no PDF at all, so the
`oxml-pdf` writer could drift with nothing watching, which is what F-X020
demonstrated.

**And the PDF writer became deterministic.** Recording the first fingerprint
proved that `to_pdf_deterministic` was not. Two runs of the same binary on the
same input produced different bytes for all seven samples. Three hashed maps
were iterated to write the file: `glyph_to_unicode` for the ToUnicode CMap
pairs, `prepared_fonts` and `font_refs` for the font objects and each page's
`/Font` dictionary, and `image_map` for a page's image XObject names. All three
are now ordered, and `FontId` gained `PartialOrd` and `Ord` to allow it.

**Non-obvious choices.** The structural pair and the byte digest are both
recorded, and they do different jobs. The structural pair hashes inflated bytes,
so it says **what** moved and survives a change of Deflate implementation or
level. The byte digest says **that** something moved and cannot be evaded,
including by a compression-only change the structural pair is blind to by
construction. A fingerprint of extracted text and page geometry alone was
rejected because it would have reported green on F-X020, whose `pdftotext`
output was identical in 7 of 7.

The writer fix was absorbed here rather than split into its own F-ID, because
two of the three entries cannot be recorded against output that disagrees with
itself. Normalising the ordering inside the harness instead was rejected: it
would have made the new gate blind to the defect it had just found.

Ordered containers rather than a sort at each point of use, because the property
wanted is "this map is iterated to produce output", and a type states that once
rather than every reader having to notice it three times.

The scanner reads the object syntax with the standard library alone. It takes a
stream payload by its declared `/Length` rather than searching compressed bytes
for `endobj`, reads `/Root` from the trailer rather than from anywhere in the
file, and compares `/Filter` as a parsed value so a chain is refused rather than
inflated. All three came from microscope pass 1. It raises on anything it does
not understand, because a harness that silently skips an object reports green
for the wrong reason.

**Deviations from the design plan.** The plan's `## Risk routing` read `none`,
correctly, for a diff that touched no Rust crate. Absorbing the writer fix made
the **Public API of a published crate** row match, additively, and the plan
records that. The size moved from M to L for the same reason.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, "The hash
harness", for the entry count and what each PDF entry covers.
`docs/hld/08-rendering-spec.md`, "The PDF backend", which gains the rule that
the writer's output is reproducible.

**Tests.** The gates are
`test_a_changed_content_stream_moves_the_pdf_entries_and_no_other` and
`test_refingerprinting_identical_bytes_reproduces_every_entry`, plus
`two_identical_documents_produce_identical_deterministic_pdfs` for the writer
fix, which was confirmed to fail against the unfixed writer by reverting all
three source files. Eleven tests in the harness in total, covering the resource
mirror case, the compression-level pair, a geometry change, a hostile payload,
a filter chain, three unparseable files and a missing PDF.

**Hash harness.** **Expected delta, and it is the story.** Twenty-one added
entries, 0 changed, 0 removed, taking the manifest from 28 to 49. Re-recorded
with `--update --reason` in its own labelled commit, separate from the code that
causes it. Every `word/*.xml` and `page1.png` digest holds the value it held
before the sprint, which is what the separation of the raster path from the
writer predicts.

**Notes for future sessions.** The manual demonstration the backlog's gate asks
for was run twice against the recorded baseline. Perturbing the TJ adjustment in
`emit_glyphs` by one thousandth of an em moved 14 entries, `pdf/pages` and
`pdf/bytes` for all seven samples, and left every `pdf/resources`, `page1.png`
and `word/*.xml` entry untouched. Perturbing the `/Producer` string alone, which
lives in the Info dictionary and in no stream, moved only the seven `pdf/bytes`
entries. That second case is the one the byte digest exists for.

F-X020's by-hand characterisation was therefore comparing against a moving
target. Its conclusion that the dependency refresh was benign is not undermined,
since `pdftotext` and `pdfinfo` agreed, but some of the byte movement it
attributed to `font-types` was the writer disagreeing with itself.

`scripts/golden_png_harness.py` exists and is referenced only by
`docs/hld/12-testing-strategy.md`. It is wired into neither `/verify` nor CI.
That is out of scope here and worth a look.

### F-X025, /verify must run the release regressions

**Sprint.** S43
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `/verify` step 6 runs
`python3 -m unittest scripts.test_sprint_workflow`, the module holding the
release family preflights that `.github/workflows/publish.yml` invokes by name
as the publication gate, plus the pinned CI toolchain assertions. Before this
those preflights ran for the first time on a tag, after the sprint was closed.

**Non-obvious choices.** The whole module rather than the two tests
`publish.yml` names, because naming them here would reproduce the coupling that
caused the problem: a third preflight added later would sit unrun until someone
updated two places. The module takes about four seconds, so nothing about the
omission was a cost decision.

It joins step 6 rather than becoming a twelfth step. Step 6 is already the
standard-library checks that keep the process documents honest, and a step per
script would make the gate a list rather than a shape.

The wiring test asserts the step is present **and** that a copy of `verify.md`
with the line removed fails the same assertion, so the gate defends its own
wiring rather than trusting prose nobody checks.

**Deviations from the design plan.** The plan's third test row, exercising the
preflights against mutated version carriers, was dropped as redundant.
`test_release_preparation_metadata_rejects_wasm_tag_and_version_mutations`
already mutates a version literal and asserts the contract rejects it, through
an injectable helper. Writing a second one would pin the same behaviour twice.

**Spec sections touched.** `docs/hld/15-build-and-toolchain.md`, the
`publish.yml` paragraph, which gains the statement that the same regressions run
in the canonical local gate and has its stale figures corrected to workspace
0.7.0 and incubating 0.3.0. `docs/hld/12-testing-strategy.md`, the
README-inventory paragraph, whose stale stable figure becomes 0.7.0.

**Tests.** The gate is `test_verify_runs_the_release_regressions`, plus
`test_every_test_publish_yml_names_resolves_to_a_real_test`, which resolves every
dotted path `publish.yml` invokes to a real class and method so a rename fails
locally rather than at publication. 48 in the module.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Both halves of the backlog's gate were
demonstrated end to end rather than only asserted, because they are statements
about a tree that does not exist in the repository. Moving
`crates/rpptx/Cargo.toml` to 0.3.1 fails both preflights. Putting `ci.yml`'s
`@tensorbee/rpptx-wasm` literal back to 0.2.0, which is exactly the S42 defect,
fails three tests including both WASM job assertions.

The spec set carried the stale release figures for a whole sprint before anyone
noticed, and it was noticed here only because this story had to read that
paragraph. `/realign-docs` is the command that owns that class of drift, and it
has not run recently.

### F-X026, CI must run the release regressions too

**Sprint.** S44
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Pull-request CI now has a named `Release regressions` job
that runs the complete `scripts.test_sprint_workflow` module with the Python
standard library. It catches stale release metadata before publication and
reports the failure independently of the prose checks.

**Non-obvious choices.** The job runs the whole module rather than naming the
two current publication preflights. That keeps later release contract tests in
the gate automatically. It remains separate from the path-filtered prose job,
whose inputs are Markdown only.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, "What CI runs",
and `docs/hld/15-build-and-toolchain.md`, "Publishing" and "CI job matrix".

**Tests.** `test_ci_runs_release_regressions_in_a_named_job` and
`test_ci_release_regression_job_rejects_wiring_mutations`, plus both existing
release-family preflights and the stale-version mutation regression.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep the whole-module command unconditional.
Narrowing it to named methods recreates the coupling this story removed.

### F-X027, Wire the golden-PNG gate into something

**Sprint.** S44
**Completed.** 2026-08-16
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The CI `test` job runs the golden-PNG check after the full
workspace suite. It reuses the job's checksum-pinned Poppler 26.01.0 build and
compiled sample generators, then compares all seven decoded page-one pixel
buffers at 150 DPI.

**Non-obvious choices.** Reusing `test` avoids a second Poppler source build and
keeps the raster oracle beside the workspace artifacts it consumes. The
portable hash harness remains a separate job because it does not require the
external rasterizer.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, "The golden-PNG
gate" and "What CI runs", plus `docs/hld/15-build-and-toolchain.md`, "CI job
matrix".

**Tests.** `test_ci_runs_the_golden_png_gate_in_the_pinned_poppler_environment`
asserts placement, ordering, uniqueness, exact command, and ordinary failure
propagation. The harness self-test rejects a one-pixel offset. The integrated
oracle matched seven of seven clean buffers and rejected an injected pixel in
`proposal` under Poppler 26.01.0.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The pixel manifest and rasterizer version are a
single reviewed contract. A Poppler change needs a deliberate oracle review,
not a baseline refresh hidden inside CI maintenance.

### F-X028, Repair the agent-facing documentation drift

**Sprint.** S44
**Completed.** 2026-08-16
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `CLAUDE.md`, the canonical verify command, and the two
affected HLD sections now name repository paths, package versions, font
ownership, features, and known defects that match the current tree. A
structured regression resolves repository claims from both governed documents,
and the generated Codex adapter was refreshed from the canonical command.

**Non-obvious choices.** The claim checker distinguishes concrete rooted paths,
rooted globs, placeholders, numeric line suffixes, generated outputs, and
tracked standalone filenames. This lets it reject stale documentation without
requiring generated package archives to exist in a fresh checkout.

**Deviations from the design plan.** Integration review found that the first
helper covered only backticked crate paths. It was expanded to both governed
documents and every documented path shape. Microscope pass 2 then found that
the standalone generated `*.crate` glob was missing from the exemption set.
That case was fixed, and pass 3 was clean.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`, "Packaging", and
`docs/hld/15-build-and-toolchain.md`, "Release process".

**Tests.** `test_agent_facing_repository_claims_resolve_against_the_workspace`
and `test_agent_facing_claim_contract_rejects_stale_mutations` cover stale
crate and HLD paths, a missing workflow, version and feature drift, package
outputs, globs, placeholders, and standalone filenames. The packaged
`oxml-layout` inventory contains exactly 20 TTFs and four family legal files.
The no-default-features suite passes.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Edit `.claude/commands/verify.md`, then regenerate
the adapter. Never edit `.agents/skills/verify/SKILL.md` directly.

### F-X029, Path-filtered CI jobs

**Sprint.** S44
**Completed.** 2026-08-16
**Size.** M, estimated 2 days, actual 1 day

**What was built.** CI detects changed path families once, routes eight costly
jobs only when their inputs can affect their result, and reports one stable
`CI gate` status. Docs-only changes run the documentation checks without
scheduling the workspace, MSRV, WASM, bindings, fidelity, hash, or supply-chain
jobs. Scheduled supply-chain checks still run.

**Non-obvious choices.** `dorny/paths-filter` v4.0.3 is pinned to immutable
commit `ceb8a2b8f2d89434be7ff52d3de7ec3738c5cc9d`. The detector alone receives
`pull-requests: read`. The aggregate job treats a selected failure as failure
and an unselected skip as success, which avoids the required-status trap of
job-level native path filters.

**Deviations from the design plan.** None. Repository-side routing landed, but
branch protection remains deliberately external. F-X031 now owns that setting
and is scheduled for S62 by user direction.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, "What CI runs",
and `docs/hld/15-build-and-toolchain.md`, "CI job matrix".

**Tests.** The three CI-filter contract regressions cover every routed job's
must-trigger and must-not-trigger paths, docs-only routing, scheduled
supply-chain selection, least privilege, immutable action provenance, a
fail-safe `ci.yml` route, and aggregate-gate result mutations.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** F-X031 must make only `CI gate` required after
hosted runs prove its exact reported name. Do not require the routed jobs
individually.

### F-156, Extract oxml-chart

**Sprint.** S45
**Completed.** 2026-08-17
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The complete typed ChartML model, renderer, and tests now
live in the shared `oxml-chart` crate. Active consumers depend on that shared
crate, while `rpptx-chart` remains as a deprecated exact re-export for source
and type compatibility.

**Non-obvious choices.** The extraction was a mechanical ownership move. The
compatibility crate contains no forwarding implementation, and the release,
packaging, README, doctest, and architecture assertions all name the new shared
crate explicitly.

**Deviations from the design plan.** Pre-implementation review added HLD 11 to
the impact list because its current incubating publication allowlist named
`rpptx-chart` without `oxml-chart`.

**Spec sections touched.** `docs/hld/01-glossary.md`,
`docs/hld/02-scope-and-non-goals.md`, `docs/hld/03-architecture.md`,
`docs/hld/07-inheritance-and-resolution.md`, `docs/hld/09-charts-spec.md`,
`docs/hld/11-migration-plan.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `legacy_shim_retains_shared_chart_type`, the 80 shared chart tests,
the shared-crate dependency assertion, both chart package dry-runs, and the
workspace verification gate.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** New chart implementation belongs in
`oxml-chart`. Keep `rpptx-chart` as the compatibility identity until a reviewed
migration removes it.

### F-157, Word chart part and embedded workbook

**Sprint.** S45
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Word packages can now contain typed inline and anchored
chart drawings, collision-safe chart parts, document relationships, content
types, and editable embedded workbooks. Package changes are staged atomically,
and opened producer drawing XML remains the sole round-trip source.

**Non-obvious choices.** Duplicate `externalData` detection uses a
namespace-aware reader because ChartML permits arbitrary prefix aliases and
foreign elements may share the same local name. This avoids rewriting the raw
chart model solely to expose one package guard.

**Deviations from the design plan.** The plan was revised to permit the direct
workspace `quick-xml` dependency in `rdocx` for the namespace-aware duplicate
guard. No public API or shared-crate model expansion was required.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`, native chart
parts and atomic package mutation, and `docs/hld/09-charts-spec.md`, Word chart
relationships and editable workbooks.

**Tests.** `word_chart_part_and_workbook_round_trip`,
`invalid_chart_package_assembly_is_atomic`, sparse suffix allocation, producer
XML preservation, and the SHA-bound Microsoft Word 16.104 native gate. Word
opened without repair and Edit Data successfully changed the embedded values.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve producer chart drawing XML verbatim.
Package assembly may inspect namespaces, but it must not become a second
ChartML parser.

### F-158, Document::add_chart

**Sprint.** S45
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `ChartKind`, `ChartData`, validation, ChartML construction,
and workbook construction now share one `oxml-chart` authoring path.
`Document::add_chart` uses that path to append an atomic inline Word chart,
while the existing PowerPoint public paths remain source-compatible re-exports.

**Non-obvious choices.** ChartML formulas, caches, workbook headers,
categories, and numeric cells are asserted from one validated source. Word
placement follows flow layout through width and height rather than slide-style
absolute coordinates.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/09-charts-spec.md`, shared authoring data,
validation, atomic facade mutation, and Word inline placement.

**Tests.** `added_bar_line_and_pie_charts_keep_source_data`,
`word_add_chart_writes_cache_and_workbook_from_one_source`,
`word_add_chart_rejects_invalid_data_without_mutation`, and
`word_add_chart_uses_inline_flow_placement`, plus the existing PowerPoint chart
authoring suite.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep cache and workbook serialization together
in `oxml-chart`. Facades own their package relationship scopes and should not
reimplement the data projection.

### F-159, Chart rendering in the Word paginator

**Sprint.** S45
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Word layout resolves internal chart and theme
relationships, carries backend-neutral groups through line breaking and
pagination, and renders inline and anchored charts through `oxml-chart`.
Missing, external, malformed, and unsupported targets produce stable
diagnostics and visible placeholders.

**Non-obvious choices.** A generic group item gives the line engine the same
width, ascent, descent, wrapping, and placement behavior as an image without
adding a chart dependency to `oxml-layout`. Theme lookup follows the document
relationship target instead of assuming a conventional part name.

**Deviations from the design plan.** The plan added a dev-only `rpptx`
dependency to `rdocx` so one existing test could author both golden artifacts
from the same `ChartData`. There is no production dependency edge.

**Spec sections touched.** `docs/hld/03-architecture.md`, shared dependency
direction, `docs/hld/08-rendering-spec.md`, generic group transport and Word
pagination, `docs/hld/09-charts-spec.md`, chart and theme resolution, and
`docs/hld/12-testing-strategy.md`, the exact cross-family golden gate.

**Tests.** `word_and_powerpoint_chart_pixels_are_identical` produced 750 by
450 pixel crops at 150 DPI with bundled fonts and `pdftoppm 26.01.0`, with zero
differing RGBA pixels. The Word artifact SHA-256 is
`e50845637449e2af4b8e2dbf16f5f6f53e5f598a00401fcc34c13f5d5716a1c4`, and
the PowerPoint artifact SHA-256 is
`7525e9a088c5fbf58fa1ed98cdfa0ec2fabf998662112ced7a6b6521f2c4edfc`.
Inline, anchored, theme, color-map, and visible-fallback regressions also pass.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep chart geometry child-local until the
paginator applies the inline or anchor transform. Deterministic comparisons
must use bundled fonts and the pinned rasterizer.

### F-147, Comment model and part

**Sprint.** S46
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Word comment parts now parse and serialize through typed
`CT_Comments` and comment values. Paragraphs retain typed comment range starts,
range ends, and reference runs in their insertion-aware content sequence.
Relationship discovery follows the package target, including noncanonical
targets, and documents without a comments part remain unchanged.

**Non-obvious choices.** Comment anchors keep their exact positions among raw
producer XML and runs. The model rejects malformed identifiers while preserving
unmodelled children and namespace aliases around the typed content.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`, comment ownership and
typed anchors, and `docs/hld/04-opc-and-packaging.md`, comment relationships and
part preservation.

**Tests.** `three_comments_and_cross_paragraph_anchors_round_trip_byte_identically`,
the comments parser and writer unit tests, noncanonical relationship target,
malformed identifier, and absent-part regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep body anchors insertion-aware. Do not infer
the comments part from a conventional filename when a relationship exists.

### F-148, Comment API

**Sprint.** S46
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The native facade now exposes stable half-open `RunRange`
coordinates, comment views, ranged comment creation, replies, resolved state,
and thread removal. Mutations stage body, comments, comments-extended metadata,
relationships, and content types together before committing.

**Non-obvious choices.** Reply and resolved state use paragraph identifiers in
the comments-extended part. Range insertion can split hyperlinks without moving
the half-open boundary, while removal preserves unrelated empty runs and
producer metadata.

**Deviations from the design plan.** None. The plan permitted SHA-bound Word
acceptance to be classified as a human action when it could not be observed
scriptably.

**Spec sections touched.** `docs/hld/03-architecture.md`, facade ownership and
the range contract, `docs/hld/04-opc-and-packaging.md`, the comments-extended
graph, and `docs/hld/10-bindings-spec.md`, the additive native API.

**Tests.** `a_ranged_comment_reply_and_resolution_keep_one_intact_thread`,
`removing_a_comment_removes_only_its_anchors_and_thread_metadata`, invalid range
regressions, and the comments relationship and content-type integration gate.
The candidate SHA-256 is
`a5ad0e8eb2d1a676daa07431deb2a0f11ee32e8bb92d099d14d5d16d43708adb`.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Microsoft Word 16.104 build 16.104.25121423 is
installed, but no-repair opening, reply visibility, and resolved-thread UI
acceptance were not observed. That remains an explicit human action.

### F-152, Content control model

**Sprint.** S46
**Completed.** 2026-08-17
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Word content controls now have a recursive typed model for
properties, binding metadata, bounded type markers, and ordered content at
block, row, cell, paragraph, and run placement. Ordinary body, table, cell,
paragraph, and run traversal sees wrapped content exactly once.

**Non-obvious choices.** Each placement owns one ordered content enum instead
of hiding controls behind raw XML. Empty paragraphs, hyperlinks, comment
anchors, and controls sharing the same run boundary retain producer order.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`, the recursive
content-control model and traversal boundary.

**Tests.** `controls_at_all_five_levels_round_trip_without_losing_content`,
`table_traversal_sees_rows_cells_and_paragraphs_inside_controls_once`,
`run_control_keeps_comment_anchor_and_hyperlink_boundaries`, and opaque
property and child preservation regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Recursive traversal must expose ordinary wrapped
content once while preserving the control itself as a mutation boundary.

### F-153, Content control binding

**Sprint.** S46
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The facade can list and mutate content controls in document
order by tag or alias, bind matching controls from a map, and update a related
custom XML part and every display value atomically. Store resolution follows
custom XML properties rather than filenames.

**Non-obvious choices.** Data bindings accept only namespace-aware absolute
child paths with optional one-based indices. Prefix mappings are parsed
strictly, ambiguous or overlapping mutations are rejected, and custom XML is
updated by byte splices so untouched producer bytes remain identical.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`, facade lookup and
mutation, `docs/hld/04-opc-and-packaging.md`, custom XML store resolution and
atomic package changes, and `docs/hld/10-bindings-spec.md`, the bounded binding
contract.

**Tests.** `a_control_map_updates_every_matching_display_value`,
`a_bound_custom_xml_value_updates_the_part_and_display_text_atomically`,
namespace-shadowed indexed binding, byte-preservation, nested-control, invalid
binding, and wrong-namespace item identifier regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep XPath support bounded to the documented
absolute child form. Unsupported expressions must fail before any staged edit
is committed.

### F-154, Bookmarks and cross-references

**Sprint.** S46
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Paragraphs now retain typed bookmark starts and ends, and
the facade correlates, lists, reads, and atomically inserts bookmarks over
half-open run ranges. Structured `REF` and `PAGEREF` fields retain instructions
and cached display values. `REF` resolves bookmark text before shaping, while
`PAGEREF` resolves the target page through the existing pagination result.

**Non-obvious choices.** The format-neutral layout layer carries generic target
markers and target-bearing fields. Word recursively indexes bookmark starts in
ordinary content, tables, and content controls, then substitutes pages without
a second pagination path. TOC bookmark allocation is collision-safe and handles
numeric overflow without panicking.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`, bookmark ownership and
range insertion, `docs/hld/08-rendering-spec.md`, target indexing and single-pass
page substitution, and `docs/hld/10-bindings-spec.md`, the additive bookmark
and cross-reference API.

**Tests.** `a_bookmark_inserted_over_a_range_is_listed_with_its_text`,
`ref_and_pageref_resolve_to_the_bookmark_text_and_final_page`, marker order and
raw-neighbour round-trip, malformed marker reporting, atomic failure, nested
table target, hidden boundary, empty fallback, namespace, and TOC allocation
regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep page targets in the existing layout result.
Do not add a Word-specific type to `oxml-layout` or a second pagination pass.

### F-149, Revision model

**Sprint.** S47
**Completed.** 2026-08-17
**Size.** L, estimated 4 days, actual 1 day

**What was built.** WordprocessingML insertions, deletions, moves, deleted text,
contextual change markers, and run, paragraph, table, and section property
changes now have a typed revision model. The native facade reports revision
identity, author, timestamp, and kind in main-body document order, including
content nested in tables and content controls.

**Non-obvious choices.** The captured revision subtree remains the sole
serialization source while the typed content is a read-only projection. This
keeps untouched producer prefixes, whitespace, namespace bindings, and unknown
descendants byte-identical. Malformed revisions remain raw and unreported so a
previously readable document still opens.

**Deviations from the design plan.** The approved scope included every
reachable main-document placement of the listed insertion and deletion
elements, including contextual paragraph, row, and numbering markers. Sprint
review corrected the original additive semver classification. The new
`RunContent` variant and required preservation fields on public low-level
WordprocessingML structs are an intentional breaking pre-1.0 boundary. The
workspace stays at 0.7 during development, and its next published family must
use 0.8.0. The native `Document::revisions` facade remains additive.

**Spec sections touched.** `docs/hld/03-architecture.md`, revision ownership,
preservation, traversal, and the 0.8.0 low-level boundary, and
`docs/hld/10-bindings-spec.md`, the additive native revision metadata API and
the exact breaking low-level Rust surface.

**Tests.** `revision_elements_round_trip_unchanged_and_report_metadata`,
`revision_attributes_are_prefix_tolerant_and_namespace_checked`,
`property_changes_write_in_their_schema_final_slots`, and
`nested_revisions_are_reported_once_in_document_order`, plus namespace
collision, schema-order, raw-preservation, duplicate-emission, and
`hyperlink_and_nested_content_revisions_round_trip_and_report_in_order` and
`modeled_hyperlinks_preserve_unreported_raw_children_and_foreign_owners`
regressions. The extended review also covers aliased hyperlink runs under a
locally shadowed canonical Word prefix, schema-positioned raw run properties,
parsed namespace repair, live raw boundaries after run mutations, and comment
removal from direct and content-control runs.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep raw revision XML as the write source until
an explicit resolution operation replaces it. Revision discovery remains
bounded to the main document tree in this story.

### F-150, Accept and reject revisions

**Sprint.** S47
**Completed.** 2026-08-17
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The native Word facade can accept or reject every modeled
revision, or scope the operation by exact author, inclusive RFC 3339 date
range, or revision id. Resolution covers content wrappers, moves, deleted
text, all four property-change forms, contextual paragraph markers, numbering
markers, and row markers.

**Non-obvious choices.** Operations transform a cloned document from the
inside out, promote namespace bindings needed by retained content, serialize
and reparse the complete candidate, and commit only after every selected
revision validates. Paragraph-mark deletion chains merge adjacent paragraphs
while retaining the final paragraph's formatting. Layout caches are
invalidated once after a successful commit.

**Deviations from the design plan.** None. Microscope review strengthened the
approved contract with malformed nested-selection checks, strict RFC 3339
edge cases, owner namespace recovery, and chained paragraph merge coverage.

**Spec sections touched.** `docs/hld/03-architecture.md`, placement-aware
document mutation, `docs/hld/04-opc-and-packaging.md`, atomic staged package
integrity, and `docs/hld/10-bindings-spec.md`, the eight additive native
accept and reject methods.

**Tests.** `accepting_every_revision_matches_word_normalized_body_xml`,
`rejecting_insertions_and_deletions_restores_the_recorded_content`,
`scoped_revision_actions_change_only_matching_revisions`,
`contextual_paragraph_markers_merge_the_adjacent_paragraphs`,
`rejected_property_changes_keep_owner_namespace_bindings`, and
`malformed_selected_property_changes_fail_atomically`, plus
`hyperlink_nested_revisions_resolve_inside_out_when_scoped` and
`targetless_revision_only_hyperlinks_keep_sibling_order_when_resolved`, and
`resolving_a_modeled_hyperlink_keeps_unreported_raw_children`. Extended sprint
review also added opaque malformed-wrapper, comment-boundary remapping, and
relationship-namespace collision regressions. The normalized oracle is pinned
to Microsoft Word 16.104 build 16.104.25121423.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Date scoping compares instants rather than
lexical timestamp strings. Shared revision ids intentionally select every
modeled element carrying that id, including paired move placements.

### F-151, Revision display in the renderer

**Sprint.** S48
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Native render options now select an accepted or tracked
revision view, with accepted as the compatibility default. The layout engine
projects revision-wrapped runs in preserved order. Tracked insertions are
underlined, tracked deletions are struck through, and changed paragraphs draw
an outside-margin bar on every page portion they occupy. PDF, PNG, and layout
entry points accept the same concrete options value.

**Non-obvious choices.** Only default accepted layouts are cached. Tracked
layouts are computed for the request so the existing font-mode cache does not
gain another state dimension. Revision projection feeds headings, fields,
bookmarks, hyperlinks, floating anchors, headers, footers, and notes so derived
content follows the selected view consistently.

**Deviations from the design plan.** None. Nine microscope passes extended the
planned coverage to nested-only wrappers, shared-boundary field ordering,
revision-only hyperlinks, note decorations, empty wrappers, and typed
hyperlink-owner serialization.

**Spec sections touched.** `docs/hld/03-architecture.md`, revision projection
ownership, `docs/hld/08-rendering-spec.md`, revision views and tracked
decorations, `docs/hld/10-bindings-spec.md`, additive native render options,
and `docs/hld/12-testing-strategy.md`, the deterministic two-view gate.

**Tests.** `both_revision_views_render_and_accepted_matches_resolved_document`,
`revision_views_project_wrapped_runs_in_document_order`,
`tracked_revision_decorations_override_only_underline_and_strike`,
`a_split_changed_paragraph_draws_one_margin_bar_on_each_page`, and
`default_render_methods_keep_the_accepted_view`, plus focused regressions for
headers, footers, notes, fields, bookmarks, anchors, hyperlinks, nested and
empty wrappers, and serialization ordering.

**Hash harness.** Unchanged, 49 of 49. The deterministic golden PNG gate was
pixel-identical for all 7 page-one baselines.

**Notes for future sessions.** Add new revision-sensitive derived text to the
ordered projection path. Do not parse preserved revision XML again in a
renderer backend.

### F-155, Document protection

**Sprint.** S48
**Completed.** 2026-08-17
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A typed settings root now reports read-only,
comments-only, tracked-changes-forced, and forms-only protection intent. The
native document facade exposes the recorded enforcement, formatting,
cryptographic provider, algorithm, spin-count, hash, and salt metadata through
a borrowed accessor.

**Non-obvious choices.** The settings relationship target is resolved from the
package rather than assumed. Opened producer XML remains the serialization
source. Unsupported modes and malformed numeric metadata stay opaque and
unreported so callers never receive a partial policy that looks authoritative.

**Deviations from the design plan.** None. The separately owned settings module
was explicitly approved before implementation. Microscope review removed an
unnecessary generic parsing helper before the clean second pass.

**Spec sections touched.** `docs/hld/03-architecture.md`, settings ownership,
`docs/hld/04-opc-and-packaging.md`, relationship-resolved loading and opaque
preservation, and `docs/hld/10-bindings-spec.md`, the additive native
protection accessor and unchanged binding surfaces.

**Tests.** `each_document_protection_mode_is_reported_with_its_recorded_hash`,
`document_protection_modes_and_metadata_parse_through_aliases`,
`settings_keep_document_protection_and_unmodelled_children_byte_identical`,
`malformed_document_protection_remains_opaque_and_unreported`, and
`settings_relationship_target_is_resolved_instead_of_assumed`.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Document protection records author intent. It is
not an access-control boundary, and mutation methods must not treat it as one.

### F-160, Field instruction parser

**Sprint.** S49
**Completed.** 2026-08-20
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Simple and complex Word fields now share one recursive
instruction model with normalized names, positional and switch arguments,
nested fields, cached display segments, dirty state, and a private captured
source used for package-preserving serialization. Instructions split across
runs and quoted, escaped, field-specific operands parse through the same
grammar.

**Non-obvious choices.** Untouched producer XML remains the write source.
Changed fields choose the effective public raw or structured instruction and
rewrite only the field-owned content, retaining run formatting, controls,
comments, processing instructions, namespace aliases, and unmodelled XML.
Layout, HTML, and Markdown consume cached display segments instead of a second
field classifier.

**Deviations from the design plan.** None. Microscope review extended the
planned preservation coverage to expanded controls, empty display runs,
multi-run formatting, hyperlink-owned fields, and raw-only public edits.

**Spec sections touched.** `docs/hld/03-architecture.md`, recursive field
ownership and source preservation, and `docs/hld/10-bindings-spec.md`, the
intentional 0.8 low-level Rust field-model break.

**Tests.** `field_instruction_corpus_parses_every_simple_complex_split_and_nested_form`,
`malformed_complex_fields_remain_untyped_and_preserved`,
`unchanged_complex_fields_keep_source_runs_and_unmodelled_neighbours`, and
focused cache, mutation, prefix, malformed, layout, HTML, and Markdown
regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Field serialization must keep the parser's
source identity aligned with the public raw and structured edit rules. A text
replacement that ignores run-level controls or producer trivia can change the
display or invalidate otherwise readable OOXML.

### F-161, Field evaluation engine

**Sprint.** S49
**Completed.** 2026-08-20
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The native `Document` facade now evaluates the supported
Word field families in deterministic source order and returns resolved,
deferred-pagination, or cached-fallback outcomes with stable diagnostics.
Explicit context supplies dates, filenames, merge records, and included text.
Core and custom properties, document variables, bookmarks, styles, headers,
footers, footnotes, and endnotes are read from their relationship-resolved
package sources.

**Non-obvious choices.** Results are reported in source preorder even when a
parent requires child outcomes first. Each effective outer instruction owns a
fresh nested-outcome frame so nested SEQ and IF operands evaluate exactly once
without pointer identity leaking between cloned trees. Missing, malformed, or
ambiguous inputs retain the stored display rather than producing blank text.

**Deviations from the design plan.** None. Microscope review tightened lexical
arity, overflow handling, namespace checks, structured and raw edit identity,
STYLEREF numbering semantics, switch formatting, and nested source ordering.

**Spec sections touched.** `docs/hld/03-architecture.md`, evaluator ownership,
`docs/hld/08-rendering-spec.md`, pagination deferral and stored fallback,
`docs/hld/10-bindings-spec.md`, the additive native evaluation API, and
`docs/hld/12-testing-strategy.md`, the pinned Word oracle matrix.

**Tests.** `every_supported_field_matches_the_pinned_word_result`,
`nested_if_and_comparison_operators_evaluate_recursively`,
`sequence_state_is_scoped_and_reset_by_supported_switches`,
`formatting_switches_match_the_pinned_word_matrix`,
`document_properties_variables_and_author_use_package_values`, and the
existing REF and PAGEREF pagination regression.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Ambient time, filesystem access, and package
iteration order are outside the evaluator contract. New data-backed fields
need an explicit context or relationship-resolved source and a deterministic
fallback diagnostic.

### F-162, Field update policy

**Sprint.** S49
**Completed.** 2026-08-20
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Document::update_fields`,
`Document::save_with_field_updates`, and
`Document::to_bytes_with_field_updates` now materialize F-161 outcomes into
field caches and field-local dirty flags. Existing save and byte methods remain
leave alone. Updates cover the complete typed story order and invalidate both
layout caches once after a nonempty successful batch.

**Non-obvious choices.** Evaluation and mutation run against cloned state, and
the live document commits only after traversal counts and serialized XML
validate. Package-backed header, footer, and endnote changes use anchored
field-local patches rather than whole-part serialization. Nested changes use
marker-level spans so producer run wrappers, properties, namespace scope,
foreign children, whitespace, comments, and processing instructions survive.

**Deviations from the design plan.** None. Eight microscope passes expanded
the approved source-preservation proof across opaque lookalikes, identical
nested siblings, shared boundary runs, hyperlink trivia, raw-only edits, stale
descendants, and multi-run formatting scaffolds.

**Spec sections touched.** `docs/hld/03-architecture.md`, atomic facade update
ownership, and `docs/hld/10-bindings-spec.md`, the three additive native-only
methods and unchanged binding surfaces.

**Tests.** `field_update_policies_produce_the_expected_result_cache_and_dirty_flag`,
`unsupported_fields_keep_their_cached_result_when_updates_run`,
`ordinary_save_leaves_cached_field_results_and_dirty_flags_alone`,
`field_update_failure_leaves_document_bytes_unchanged`, and the simple,
complex, nested, package-story, dirty-alias, layout-cache, and save-reopen
preservation regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Do not replace package stories wholesale when
updating a field. Exact producer preservation depends on typed placement
anchors and on retaining physical run scaffolding outside the field-owned
children.

### F-203, Reader compatibility corrections

**Sprint.** S49
**Completed.** 2026-08-20
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Table-cell properties now recognize Word children and
attributes through expanded names, preserve foreign same-local-name and
unsupported children as raw XML, retain namespace bindings declared on the
property or cell owner, and serialize every modeled and preserved child in its
absolute `CT_TcPr` schema slot. Numbering levels retain raw `w:isLgl` before a
typed suffix.

**Non-obvious choices.** The raw sidecar maps the complete eighteen-child
`CT_TcPr` sequence rather than assigning a relative boundary during parsing.
That keeps later typed mutations valid even when standard unmodelled children
appear before or after them. Content-control cells carry the same owner-binding
rules as direct row cells.

**Deviations from the design plan.** The plan was revised before completion to
record the `CT_TcPr` preservation sidecar as part of the intentional pre-1.0
0.8 low-level Rust source break and to add the owner-local namespace gate.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`, the intentional 0.8
low-level table-property preservation boundary.

**Tests.** `foreign_cell_width_remains_raw_and_unmodelled`,
`aliased_cell_width_uses_in_scope_word_bindings`,
`cell_property_preserves_child_binding_declared_on_owner`,
`content_control_cell_preserves_child_binding_declared_on_cell`,
`unmodelled_standard_cell_properties_keep_absolute_slots_after_typed_mutation`,
and `level_raw_is_lgl_stays_before_suffix`.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** New typed `CT_TcPr` properties must be placed in
the existing absolute sequence. Namespace ownership can live above a preserved
child, so raw serialization must carry the complete external binding scope.

### F-163, Template syntax

**Sprint.** S50
**Completed.** 2026-08-21
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The native `Document` facade now renders scalar
`{{ path.to.value }}` tags from `serde_json::Value`. Tags may cross formatted
Word runs in body content, tables, headers, footers, text boxes, and chart
labels. Strings, numbers, booleans, and null have documented conversions.

**Non-obvious choices.** Rendering evaluates against a staged document clone
and commits only after every tag and value validates. Collision-free sentinels
keep replacement text containing template syntax from being evaluated again.
The first matched run owns replacement formatting while unmatched run content
and unmodelled XML remain in place.

**Deviations from the design plan.** None. Microscope review added direct
table-cell coverage for the shared cross-run replacement path.

**Spec sections touched.** `docs/hld/03-architecture.md`, facade ownership,
`docs/hld/04-opc-and-packaging.md`, staged preservation,
`docs/hld/10-bindings-spec.md`, the additive native API, and
`docs/hld/12-testing-strategy.md`, scalar template gates.

**Tests.** `a_tag_split_across_five_formatted_runs_preserves_surrounding_formatting`,
`dotted_scalar_paths_render_supported_json_leaves`,
`invalid_template_input_leaves_the_document_unchanged`, and
`template_render_preserves_unmodelled_paragraph_xml`.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Scalar rendering is deliberately non-recursive.
Object and array leaves are invalid scalar values, and any invalid input must
leave both typed content and package state unchanged.

### F-164, Loops and conditionals

**Sprint.** S50
**Completed.** 2026-08-21
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Dedicated marker paragraphs and rows now define nested
`for` and `if` blocks over body entries and table rows. Loop variables use
lexical scopes, conditionals use documented JSON truthiness, and section-ending
paragraphs retain their section properties when cloned.

**Non-obvious choices.** A container-aware stack parser rejects mismatched,
crossed, or cross-container markers before mutation. Row markers are resolved
against their owning table depth, so nested table markers cannot be mistaken
for outer row controls. Preflight covers scalar paths even when a conditional
branch will not render.

**Deviations from the design plan.** None. Microscope review corrected the
structural-only commit path, nested-table marker ownership, false-branch
preflight, and populated nested-table rejection for an outer marker row.

**Spec sections touched.** `docs/hld/03-architecture.md`, structural evaluator
ownership, `docs/hld/04-opc-and-packaging.md`, atomic container evaluation,
`docs/hld/10-bindings-spec.md`, block syntax and scope, and
`docs/hld/12-testing-strategy.md`, nested structural gates.

**Tests.** `a_nested_loop_and_conditional_generate_the_expected_document`,
`mismatched_or_cross_container_blocks_fail_without_mutation`,
`loop_scopes_shadow_root_values_and_restore_after_exit`, and
`structural_generation_preserves_schema_order_and_raw_xml`, plus focused
regressions for row-only output, nested tables, false branches, and marker-row
content.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Structural controls remain limited to the main
body and its tables. Other relationship-resolved stories retain scalar-only
rendering through the shared replacement path.

### F-165, Repeating table rows and lists

**Sprint.** S50
**Completed.** 2026-08-21
**Size.** M, estimated 2 days, actual 1 day

**What was built.** One row loop may now repeat several adjacent template rows
for each record. Deep clones retain table banding, grid spans, vertical merge
state, content controls, ordered row and cell XML, and source numbering
identity. Repeated numbered paragraphs continue one list sequence.

**Non-obvious choices.** A recursive preflight validates every numbering
reference in repeated body entries, rows, nested tables, and content controls
before the staged candidate can commit. Valid repetition keeps the existing
`numbering.xml` definitions unchanged instead of synthesizing new identities.

**Deviations from the design plan.** None. Microscope review added direct
typed content-control coverage to the repeated row round trip.

**Spec sections touched.** `docs/hld/03-architecture.md`, multi-row clone
ownership, `docs/hld/04-opc-and-packaging.md`, numbering and raw XML
preservation, `docs/hld/10-bindings-spec.md`, repeated structure semantics,
and `docs/hld/12-testing-strategy.md`, thirty-row and continuous-list gates.

**Tests.** `three_template_rows_over_ten_records_produce_thirty_preserved_rows`,
`repeated_numbered_items_keep_one_continuous_sequence`, and
`repeated_rows_and_lists_preserve_properties_and_raw_xml`, including atomic
rejection of an invalid repeated numbering reference.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Repeated list items preserve their source
`numId` and level. A missing definition is a render error rather than a reason
to create or renumber package state.

### F-166, Mail merge

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Document` now produces one complete document per flat
record or one document with a section per record. Merge fields use a private
missing-as-empty policy, while the ordinary field evaluator keeps its cached
fallback behavior. All record candidates are staged before output is exposed.

**Non-obvious choices.** Section assembly retains final section properties and
remaps bookmark, content-control, drawing, and reference identities across
records. Record-varying non-body stories are rejected rather than silently
reusing one record. Footnote updates patch the relationship-resolved source
part while preserving unrelated raw XML.

**Deviations from the design plan.** None. Microscope remediation expanded the
identity and story scanners to namespace-aware preserved XML and unified the
staging clone used by template and merge operations.

**Spec sections touched.** `docs/hld/03-architecture.md`, mail-merge ownership,
`docs/hld/04-opc-and-packaging.md`, staged package and section assembly,
`docs/hld/10-bindings-spec.md`, native merge APIs, and
`docs/hld/12-testing-strategy.md`, merge fixtures and atomicity.

**Tests.** `a_fixture_record_set_produces_separate_and_sectioned_documents`,
plus missing-field, boundary, non-body story, identity remapping, footnote raw
preservation, and source-atomicity regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Combined-section mode is deliberately bounded
to record-varying main-body content. A future story that varies headers,
footers, or notes per record must clone their parts and relationships rather
than weakening this rejection.

### F-167, Document comparison

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `Document::compare` creates tracked insertions, deletions,
and property changes through deterministic hierarchical alignment of body
paragraphs, tables, rows, cells, runs, and existing content-control shells.
Accepting the result reproduces the edited document and rejecting it reproduces
the original. Formatting-only changes also produce diagnostics.

**Non-obvious choices.** Comparison uses the existing revision grammar and
accept or reject resolver as postcondition oracles. It preserves paragraph,
table, cell, field, control, and unsupported raw XML ownership rather than
canonicalising the whole document. Unsupported shell changes fail atomically.

**Deviations from the design plan.** None. Five microscope remediation rounds
strengthened raw whitespace, field ownership, revision marker placement,
numbering owner cleanup, nested control and row ownership, and attributed
producer property restoration.

**Spec sections touched.** `docs/hld/03-architecture.md`, comparison ownership,
`docs/hld/04-opc-and-packaging.md`, preservation and atomic failure,
`docs/hld/10-bindings-spec.md`, native comparison API, and
`docs/hld/12-testing-strategy.md`, exact accept and reject gates.

**Tests.** The regression gate compares body text, lists, tables, nested
tables, and content controls, then proves both accept and reject postconditions.
Focused regressions cover repeated rows, raw fields and controls, numbering
addition and removal, namespace ownership, and formatting diagnostics.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Header, footer, note, comment, text-box, and
character-granularity comparison remain future scope. Exactness is defined by
the typed and preserved body representation rather than rendered appearance.

### F-168, Watermarks

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Native text and image watermark setters author
header-scoped VML, round-trip recognized shapes without rewriting unrelated
header XML, and render deterministic rotated watermark groups behind body text
on every applicable page and section.

**Non-obvious choices.** Opened VML remains raw serialization authority while a
conservative typed projection drives layout. Generated first, even, and default
header variants match Word fallback behavior. Image relationships stay local
to their header and use the collision-safe media registry.

**Deviations from the design plan.** None. Microscope remediation added native
section fallbacks, inherited-header preservation, page-number restart parity,
canonical VML shape types, namespace-safe scans, named-color handling, and
margin-relative placement.

**Spec sections touched.** `docs/hld/03-architecture.md`, watermark ownership,
`docs/hld/04-opc-and-packaging.md`, VML and header media relationships,
`docs/hld/08-rendering-spec.md`, per-page behind-text lowering,
`docs/hld/10-bindings-spec.md`, native setters, and
`docs/hld/12-testing-strategy.md`, deterministic golden evidence.

**Tests.** `watermark_renders_behind_body_text_on_every_page`, with exact
five-page PNG digests, plus VML preservation, fixed-prefix writing, header
inheritance, first and even variants, media collisions, section parity, and
atomic setter regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Only recognized API-owned text and image VML
shapes are replaced or rendered. Other `w:pict` content remains opaque and
byte-preserved.

### F-X037, Trace Word glyphs to source paragraphs

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `SourceNodeId` and exclusive Unicode-scalar `SourceSpan`
metadata now flow through shaping and both line-splitting stages. New Word
layout entry points return `WordLayoutResult`, whose result-local side table
resolves glyph runs to body, nested-table, header, footer, footnote, and
endnote paragraph paths.

**Non-obvious choices.** Generated markers, evaluated fields, note labels, and
non-bijective text transformations remain truthfully unattributed. Existing
low-level layout functions still return `LayoutResult` and discard provenance.
Field scalar offsets and displayed projection now share one ownership function.

**Deviations from the design plan.** None. Microscope review found and fixed an
ambiguous repeated-text field offset by making `Field::projected_text` the one
source of truth for both run text and source starts.

**Spec sections touched.** `docs/hld/03-architecture.md`, provenance ownership,
`docs/hld/08-rendering-spec.md`, shaping and revision projections,
`docs/hld/10-bindings-spec.md`, low-level source boundary,
`docs/hld/12-testing-strategy.md`, exact path and range gates, and
`docs/hld/14-development-backlog.md`, the issue 38 contract.

**Tests.** `every_sourced_glyph_run_resolves_to_its_exact_word_text`, plus
Unicode split, repeated story, revision view, generated text, caller-font,
deterministic, field ownership, and legacy result parity regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Source ids are one-based and local to one
result. F-X038 must rebind cached scalar ranges to current ids rather than
retaining a prior layout's identity.

### F-X032, Expose complete Word layout results

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Four native `Document` accessors now expose complete
`WordLayoutResult` bundles. Accepted normal-font layouts share the existing
`Arc` cache, tracked options remain uncached, and caller-font layouts return
owned bundles with the exact font bytes and Word source map used by shaping.

**Non-obvious choices.** Cached results use `Arc` so external renderers do not
duplicate positioned pages, fonts, or provenance maps. Caller-provided font
sets remain uncached because borrowed font inputs have no stable cache key.
Existing PDF, raster, page, and caller-font PDF paths consume the same bundle
paths as the new public accessors.

**Deviations from the design plan.** None. Microscope review strengthened the
tests with a distinct in-memory font family, an already-populated accepted
cache, and a non-default tracked revision view.

**Spec sections touched.** `docs/hld/03-architecture.md`, facade ownership,
`docs/hld/08-rendering-spec.md`, layout result and cache behavior,
`docs/hld/10-bindings-spec.md`, native public accessors,
`docs/hld/12-testing-strategy.md`, public traversal and cache-isolation gates,
and `docs/hld/14-development-backlog.md`, the issue 37 contract.

**Tests.** `full_layout_exposes_resolvable_font_data_and_reuses_the_cache`,
plus caller-font byte ownership, accepted and tracked cache isolation, public
downstream traversal, WASM compilation, package dry-run, and archive-size
checks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** F-X038 may reuse shaping and safe paragraph
work, but it must preserve the accepted cache identity contract and rebuild
result-local provenance ids for every returned layout.

### F-X034, Reviewed release notes for every release

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** S, estimated 1 day, actual 1 day

**What was built.** A canonical `/release-notes TAG` ceremony now derives one
reviewed GitHub release body from `CHANGELOG.md`. It requires meaningful
Highlights, Added, Fixed, Compatibility, and Contributors sections for either
release family. Publication validates the notes before crates.io commands and
creates the GitHub release from the exact reviewed bytes.

**Non-obvious choices.** The validator parses visible Markdown structure
conservatively. Headings hidden in comments, fences, or raw HTML do not count,
and syntax-only links, references, code markers, HTML, or invisible Unicode do
not satisfy a required section. Rendering preserves the accepted source body
byte for byte. The generated Codex adapter points to the canonical command.

**Deviations from the design plan.** None. Eleven microscope passes tightened
CommonMark boundaries, semantic emptiness, canonical SemVer, pre-publication
ordering, exact executable checks, and artifact immutability.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, mutation-sensitive
workflow evidence, `docs/hld/14-development-backlog.md`, the permanent notes
ceremony, and `docs/hld/15-build-and-toolchain.md`, pre-publication validation
and exact GitHub body publication.

**Tests.** `test_release_notes_require_complete_reviewed_changelog_sections`
and four adjacent release-workflow tests, plus the 62-test workflow suite,
generated-skill validation, publication-order mutation tests, and the full
integrated verification gate.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Every future release story must run the notes
ceremony, review its rendered body, and verify the published body byte for byte.
The ceremony never replaces the separate final release approval.

### F-X038, Cache relayout work across document edits

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Normal layouts now reuse a process-lifetime system-font
snapshot, bounded file-byte and exact-key shaping caches, one lazy synchronized
`Document` engine, and a bounded cache for context-independent body paragraphs.
Warm layouts remain byte-equivalent to cold layouts while rebuilding only
changed safe paragraphs.

**Non-obvious choices.** Deterministic and caller-font paths remain isolated.
Paragraph entries publish only after the whole layout succeeds, replay exact
font-resolution traces and diagnostics, and rebind scalar ranges to the current
result-local source ids. Context-sensitive paragraphs bypass reuse. Every
persistent and transaction-local cache has an entry or retained-byte ceiling,
including reflow buffers and active font faces.

**Deviations from the design plan.** None. Four microscope passes strengthened
font-table identity, late-failure rollback, caller and tracked isolation,
AlternateContent safety, TTC sharing, poison reuse, active-face correctness,
transaction staging, trace capacity release, and exact retained-memory
accounting.

**Spec sections touched.** `docs/hld/03-architecture.md`, persistent engine
ownership, `docs/hld/08-rendering-spec.md`, safe paragraph reuse and provenance,
`docs/hld/10-bindings-spec.md`, unchanged public facade behavior,
`docs/hld/12-testing-strategy.md`, warm and cold equality gates,
`docs/hld/14-development-backlog.md`, the issue 39 contract, and
`docs/hld/15-build-and-toolchain.md`, process-lifetime font discovery.

**Tests.** `warm_relayout_matches_cold_and_rebuilds_only_changed_safe_paragraphs`,
plus exact font order, 257 active families, TTC byte sharing, context mutation,
late failure, diagnostics, current provenance, poison recovery, transaction and
retained-byte bounds, no-default, WASM, package, and threading regressions.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Normal system-font discovery is a
process-lifetime snapshot, so installing or replacing host fonts requires a
process restart. Release notes for 0.4.0 and 0.8.0 must credit
`@emptinessform` for the issue 39 measurements and cache proposal.

### F-X033, Integrate PR 36 ordered body items

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Pedro Assumpcao's additive `Document::body_items` reader
now exposes direct Word body paragraphs, tables, body content controls, and
preserved unsupported XML in source order. The contribution landed through
GitHub PR 36 with its original commit and merge record intact.

**Non-obvious choices.** Existing `paragraphs()` and `tables()` accessors stay
recursive, while `body_items()` is deliberately direct. Self-closing Word
paragraphs and tables normalize to typed empty values, and a self-closing
section-properties child remains schema-final state rather than an unsupported
body item. Every foreign or unsupported empty child remains raw XML.

**Deviations from the design plan.** The plan was revised after microscope
review exposed the self-closing parser boundary. Three passes added expanded
name checks and unconditional raw fallback without broadening the public API.

**Spec sections touched.** `docs/hld/03-architecture.md`, direct body ownership,
`docs/hld/10-bindings-spec.md`, the native-only reader,
`docs/hld/12-testing-strategy.md`, public opened-package evidence, and
`docs/hld/14-development-backlog.md`, the PR 36 integration contract.

**Tests.** `public_body_items_preserve_opened_document_order`,
`body_items_preserve_paragraph_table_control_and_raw_order`, and
`self_closing_modeled_body_children_are_typed_by_namespace`, plus fresh
current-base GitHub CI run 32516942671, full workspace verification, package
dry-run, and archive inventory.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Direct body order is now public, but nested
content-control and table traversal remains owned by the existing recursive
accessors. Preserve contributor credit through the PR 36 merge record.

### F-X035, Tag rpptx-v0.4.0

**Sprint.** S51
**Completed.** 2026-08-21
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete 15-package shared OOXML and PowerPoint family
is published on crates.io at 0.4.0. This is the first release of `oxml-chart`.
The local `rpptx-wasm` package moved with the family but remains unpublished.
GitHub release `rpptx-v0.4.0` contains the reviewed changelog body unchanged.

**Non-obvious choices.** The release used one annotated tag at reviewed SHA
`9dee4335c531ca24abbdc995294edbb48c00183f`. Workflow run 32527109236 skipped
the stable allowlist, published the incubating crates in dependency order, and
created the GitHub release only after archive and release-note validation.
Independent checks downloaded every selected registry version and confirmed
`mantissaman` as owner. The remote tag peels to the reviewed SHA, and the
published release body is byte-identical to a fresh notes render.

**Deviations from the design plan.** None. The separately approved release
completed every deferred publication checklist item.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`, the published
incubating source boundary, and `docs/hld/15-build-and-toolchain.md`, the
published package family and reviewed tag.

**Tests.** `test_incubating_release_family_is_prepared_at_0_4_0`, the 63-test
workflow suite, full workspace verification, patched 22-package dry run,
archive inventories, both WASM checks, `cargo info` and owner checks for all 15
published crates, remote tag verification, and byte-exact GitHub release-note
comparison.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Stable 0.8.0 may now be prepared against the
published 0.4.0 dependency family. It requires its own reviewed SHA and a new
explicit `/release v0.8.0` approval before any stable tag or publication.

### F-X036, Tag v0.8.0

**Sprint.** S51
**Completed.** 2026-08-22
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The exact seven-package stable Word family is published on
crates.io at 0.8.0. The four workspace-version binding and support packages
remain unpublished. GitHub release `v0.8.0` contains the reviewed changelog
body unchanged.

**Non-obvious choices.** The release used one annotated tag at reviewed SHA
`0cc47eb8632de184ba758fe0929d9f749ab4fcb0`. Workflow run 32536705662
published only the stable allowlist in dependency order, skipped the
incubating allowlist, and created the GitHub release after output, metadata,
notes, and archive verification. Independent checks downloaded all seven
registry versions and confirmed `mantissaman` as owner. The remote tag peels
to the reviewed SHA, and the 9,291-character published release body is
byte-identical to a fresh notes render.

**Deviations from the design plan.** The default third sprint-review pass found
one missing Issue 37 credit in the stable notes. The remediation added the
verified reporter attribution and mutation-sensitive coverage. An explicitly
bounded fourth pass was clean before release approval. No product code or
release carrier changed during that remediation.

**Spec sections touched.** `docs/hld/03-architecture.md`, the published stable
family boundary, `docs/hld/10-bindings-spec.md`, the shipped native and
low-level compatibility surface, `docs/hld/12-testing-strategy.md`, the
published README endpoint, and `docs/hld/15-build-and-toolchain.md`, the
verified release tag and package family.

**Tests.** `test_stable_release_family_is_prepared_at_0_8_0`, the 66-test
workflow suite, full workspace verification, all 49 unchanged hashes, the
patched 22-package dry run, archive inventories, both WASM checks, no-default
layout, docs, README tests, supply-chain checks, seven `cargo info` and owner
checks, remote tag verification, and byte-exact GitHub release-note
comparison.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Issue 37 and PR 36 have verified stable release
comments, and Issue 37 is closed. The later Issue 39 proposals for shared
`FontData` bytes and public engine handoff did not ship in 0.8.0. Review
those as separate follow-up changes against the post-release code.

### F-169, Agile encryption, read

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** L, estimated 4 days, actual 1 day

**What was built.** A default-off `agile-encryption` feature now opens
Microsoft Agile encrypted OOXML packages through `OpcPackage` and the native
Word facade. The reader validates AES-128, AES-192, AES-256 and the declared
SHA family, verifies the password, authenticates the complete encrypted
package before ZIP parsing, and decrypts bounded segments.

**Non-obvious choices.** Package authentication precedes ZIP construction, and
every wrong-password, malformed, or tampered input fails without publishing a
partial package. The Microsoft Word reference is encoded in source rather than
stored as an opaque fixture.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and `docs/hld/15-build-and-toolchain.md`.

**Tests.** `word_agile_document_opens_only_with_its_password`,
`agile_parameters_reject_unknown_or_inconsistent_algorithms`, and
`tampered_agile_package_fails_before_zip_parsing`, plus no-default, WASM,
dependency-direction, package, and supply-chain gates.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The feature stays out of default and WASM
graphs. Authentication is part of the package boundary, not a facade policy.

### F-171, Digital signature verification

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** L, estimated 4 days, actual 1 day

**What was built.** A default-off `digital-signatures` feature discovers OPC
signature origins through relationships, applies exclusive canonicalization
and the OPC relationship transform, verifies RSA-SHA256 and X.509 material,
and reports exact declared part and relationship coverage through `OpcPackage`
and `Document`.

**Non-obvious choices.** Cryptographic validity and certificate trust are
separate results. The verifier uses `ring` rather than the advisory-blocked
RustCrypto RSA crate, and it fails closed for external, duplicate, missing,
or partially covered references.

**Deviations from the design plan.** The implementation changed the planned
RSA dependency to `ring` after the supply-chain gate rejected the former.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and `docs/hld/15-build-and-toolchain.md`.

**Tests.** `valid_signature_reports_complete_declared_coverage`,
`signature_parser_is_prefix_tolerant_and_algorithm_strict`,
`partial_or_malformed_coverage_never_reports_success`, and
`verification_does_not_change_package_bytes`. Microsoft Word 16.104 opened the
generated signed document.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Verification proves signed bytes and declared
coverage. Certificate-chain trust remains caller policy.

### F-X039, Share layout payloads and transfer reusable engines

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Layout font bytes now use `Arc<[u8]>`, pages use
`Arc<PageFrame>`, and the Word facade can transfer reusable normal-layout work
between documents only after exact context compatibility succeeds.

**Non-obvious choices.** Owned single-page facade methods remain source
compatible. The transfer API moves no public `Engine`, preserves both sides on
rejection, and includes every layout-sensitive context component, including
the wrapping-drawing predicate.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, and `docs/hld/10-bindings-spec.md`.

**Tests.** Arc pointer-sharing, compatible and incompatible engine transfer,
staged failure preservation, poison recovery, completed-cache preservation,
and PDF, raster, diagnostics, outline, font, and provenance equivalence.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** PR 40 and PR 41 proposed raw engine take and set
methods. The checked transfer boundary intentionally makes incompatible cache
ownership unrepresentable.

### F-X041, Remove duplicated glyphs at break opportunities

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Word line breaking now reshapes each final run exactly
once, removing duplicated glyph vectors at break opportunities while keeping
the shared line-breaking layer format-neutral.

**Non-obvious choices.** Deterministic bundled caller fonts drive the regression
and every baseline. The behavior commit owns the exact expected baseline
movement rather than folding it into another layout change.

**Deviations from the design plan.** The reviewed hash declaration was expanded
to match the mechanically affected PDF streams and embedded subsets.

**Spec sections touched.** `docs/hld/03-architecture.md` and
`docs/hld/08-rendering-spec.md`.

**Tests.** `break_opportunities_emit_every_scalar_and_glyph_once`,
`reported_words_do_not_duplicate_boundary_glyphs`, and
`fixed_break_runs_match_pdf_and_raster_backends`, plus the pinned golden-PNG
gate.

**Hash harness.** Intentional 26-entry delta. Five page-one PNGs changed, and
the pages, resources, and bytes PDF fingerprints changed for all seven samples.
All 21 XML entries stayed unchanged. The resulting 49-entry baseline and all
seven 150 DPI golden pixels pass.

**Notes for future sessions.** Corrected glyph vectors can change embedded font
subsets even when a sample's first-page pixels stay identical.

### F-X042, Prove headers and footers in PDF output

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Readable in-code DOCX packages now prove first, even,
default, blank, and inherited header and footer variants through public layout,
save and reopen, and deterministic PDF text extraction. The proof exposed and
fixed same-type footer inheritance.

**Non-obvious choices.** A dev-only Flate decoder reads this repository's
deterministic PDF object shape without adding a production API or relying on an
unpinned external text extractor. Explicit blank variants never borrow a
default.

**Deviations from the design plan.** The planned test-only story gained one
narrow production correction after the regression exposed footer inheritance
being omitted from effective section state.

**Spec sections touched.** None. Existing HLD intent already required the
correct behavior.

**Tests.** `authored_reopened_headers_and_footers_reach_pdf`,
`blank_first_and_even_variants_do_not_borrow_defaults`, and
`header_footer_pdf_fixture_preserves_unrelated_package_state`.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Header and footer inheritance must remain
separate and same-type. Unrelated package parts and raw XML are preservation
authority.

### F-170, Agile encryption, write

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `OpcPackage` and `Document` now stage Word-compatible Agile
encrypted output using AES-256-CBC, SHA-512, 100,000 spins, fresh secrets, CFB
version 3, complete DataSpaces streams, and authenticated segmented package
ciphertext.

**Non-obvious choices.** Caller-owned `Vec<u8>` output reserves the complete
append before publication and rolls back injected partial failures. File saves
use same-volume replacement, including `MoveFileExW` on Windows. DataSpaces
bytes are decoded independently in tests.

**Deviations from the design plan.** The generic writer parameter was narrowed
to a caller-owned byte buffer because arbitrary `Write` cannot promise
failure-atomic publication.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `agile_writer_emits_word_profile_parameters`, DataSpaces decoding,
fresh plaintext secret checks, output reserve failure, existing-destination
replacement, and native round trip. Microsoft Word 16.104 build
16.104.25121423 opened password `rdocx-f170` and rejected an incorrect
password, as observed by the user.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Word's DataSpaces transform block-size field is
zero in this envelope, while the outer compound file must be CFB version 3.

### F-X040, Restart pagination and cache table blocks

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The reusable Word engine now restarts only from exact safe
single-section page boundaries and caches safe table blocks transactionally.
Warm edits rebuild a bounded changed region and restore only byte-equal final
page Arcs.

**Non-obvious choices.** Notes, fields, floats, tables in restart regions,
multi-section state, backgrounds, unsupported content, keep constraints, and
provenance-changing insertions conservatively use a full layout. A
traversal-sensitive block disables later retained reads for that layout.

**Deviations from the design plan.** Microscope review added the document-wide
wrapping predicate and deletion and provenance fallbacks to the exact context.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` and
`docs/hld/12-testing-strategy.md`.

**Tests.** `warm_restart_rebuilds_only_the_bounded_changed_region`,
`unsafe_pagination_state_falls_back_to_full_layout`,
`earlier_note_insertion_invalidates_later_cached_markers`, and
`safe_tables_reuse_transactionally_and_with_bounds`.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Exact state and typed equality are correctness
authority. Fingerprints may only prefilter candidates.

### F-X043, Reuse bundled fallback caller-font layouts

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The Word facade now combines caller fonts at highest
priority with deterministic bundled fallbacks and retains a private reusable
engine across edits. A checked transfer moves that engine only for an exact
caller-font and document context match.

**Non-obvious choices.** The strict caller-only path remains isolated and still
fails for incomplete font sets. No system font can enter the bundled-only path,
and staged mutations and poison recovery preserve the private engine safely.

**Deviations from the design plan.** The result remains an owned
`WordLayoutResult` whose heavy page and font payloads already share Arcs.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** Caller override with bundled fallback, strict isolation, compatible
and incompatible checked transfer, warm versus fresh equality, staged failure,
poison recovery, WASM, package, and hash gates.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** This is the safe remaining behavior from PR 40
and PR 41 by `emptinessform`. Raw engine access was intentionally not adopted.

### F-X044, Scale paragraph-cache lookup for editors

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Paragraph cache lookup now uses a borrowed deterministic
fingerprint as a prefilter, keeps typed `CT_P` equality authoritative, avoids a
key clone on hits, and no longer performs an ordered remove and reinsert. The
paragraph partition holds 4,096 entries and 56 MiB inside the shared 64 MiB
layout envelope.

**Non-obvious choices.** Hits retain FIFO position. Unsafe traversal still
disables later reads, and late failures publish no staged paragraph or table
work.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` and
`docs/hld/12-testing-strategy.md`.

**Tests.** `paragraph_fingerprint_collision_requires_typed_equality`,
`editor_scale_paragraph_cache_avoids_warm_thrash`,
`unsafe_prefix_still_disables_later_paragraph_hits`,
`scaled_paragraph_cache_warm_equals_cold`, and bounds and failure publication
gates. The 700-paragraph edit records 699 hits and one build.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** This safely incorporates the editor workload
from PR 41 by `emptinessform` without making a 64-bit hash authoritative.

### F-X045, Cache headers and footers transactionally

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Safe header and footer variants now reuse exact typed
blocks through a transactional 64-entry, 4 MiB cache. Hits rebind current
source ids and replay diagnostics and font traces.

**Non-obvious choices.** Identity includes complete section geometry,
relationships, resolved parts, media, revision, fonts, provenance, and the
outer reusable context. Opaque XML bypasses reuse unless it is the supported
namespace-resolved watermark projection.

**Deviations from the design plan.** Microscope review expanded retained-byte
accounting, tightened opaque XML namespace checks, and added inherited-variant
hit evidence.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` and
`docs/hld/12-testing-strategy.md`.

**Tests.** Exact first, even, default, header, footer, inherited, image,
watermark, same-width geometry, context, provenance, late-failure, oversized,
and combined-bound regressions, plus full warm and fresh PDF equality.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** This replaces PR 41's hash-keyed immediate
publication cache with typed equality and whole-layout publication.

### F-X046, Reuse substituted pages exactly

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Restart records now retain aligned pristine and substituted
page pairs so unchanged PAGE, NUMPAGES, and PAGEREF pages can reuse the exact
prior output Arc without enabling pagination restart for field-bearing blocks.

**Non-obvious choices.** Page number, total pages, bookmark targets, revision,
font trace, pristine identity, and returned canonical font space are exact key
material. Field-free pages in the same record retain pointer identity too, and
all pair and vector capacity stays within the existing 32-entry, 2 MiB budget.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` and
`docs/hld/12-testing-strategy.md`.

**Tests.** `unchanged_page_fields_reuse_substituted_frames`,
`changed_substitution_context_reshapes_pages`,
`substituted_page_reuse_is_bounded_and_complete_equal`, and complete PDF and
raster backend equality.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** This is the bounded exact substitute for PR
41's unbounded pristine and substituted page map.

### F-X047, Attribute empty Word paragraphs

**Sprint.** S52
**Completed.** 2026-08-22
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Every otherwise empty Word paragraph now emits one
zero-width empty text segment with resolved paragraph-mark font metrics. In
provenance mode it carries the paragraph source node and scalar range `0..0`.
Body, table, header, footer, footnote, and endnote stories all participate.

**Non-obvious choices.** The carrier shapes no glyph, preserves the legacy
empty-line box, does not perturb first glyph-use font ordering, and is ignored
by PDF font, alpha, ordinal, and paint emission. Non-Word empty-glyph runs keep
their former behavior.

**Deviations from the design plan.** A hidden additive metrics-only resolver on
the pre-1.0 `oxml-layout` surface earned and passed the public API package
rider.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, and `docs/hld/12-testing-strategy.md`.

**Tests.** `empty_word_stories_emit_one_attributed_zero_width_segment`,
`empty_paragraph_uses_resolved_default_metrics`, and
`empty_segment_is_backend_invisible_and_layout_compatible`, including literal
carrier removal for PDF and raster comparison.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** This completes the safe remaining visible
behavior from PR 41 by `emptinessform` without synthesizing a space or glyph.

### F-172, Digital signature creation

**Sprint.** S53
**Completed.** 2026-08-23
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The feature-gated OPC package and native Word facade can
create deterministic RSA-SHA256 digital signatures from PKCS#8 key material
and an X.509 certificate. Creation stages the complete signature graph,
verifies every resulting signature report, and publishes only after complete
cryptographic and relationship coverage succeeds.

**Non-obvious choices.** Certificate trust remains caller policy. The unsigned
relationship graph is validated before part allocation so dangling targets,
misplaced signature relationship types, and invalid orphan signature parts
cannot be repaired or hidden by signing.

**Deviations from the design plan.** The external oracle was revised to the
available Word for Mac 16.104 evidence. Word recognized the signature and
protected the document, while the exact serialized and reopened bytes passed
local cryptographic and coverage verification. No Windows trust verdict is
claimed. Microscope review added the unsigned-graph and all-report checks.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `signature_creation_uses_schema_order_and_complete_canonical_references`,
`signature_creation_rejects_mismatched_or_unsupported_key_material`,
`signed_package_verifies_with_complete_coverage`,
`every_signature_creation_failure_leaves_live_package_unchanged`, and
`word_for_mac_recognizes_and_protects_the_created_signature`.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The Mac oracle proves recognition and document
protection. It does not replace Windows certificate-chain trust evidence.

### F-173, Tagged PDF structure tree

**Sprint.** S53
**Completed.** 2026-08-23
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Word pagination now carries one backend-neutral semantic
tree into deterministic PDF marked content. The writer emits headings, nested
lists, table roles, figures with alternate text, page parent arrays,
`StructTreeRoot`, `MarkInfo`, language, titles, links, and destinations with
exact MCID ownership.

**Non-obvious choices.** Decorative paint is always an artifact. Invalid
public structure graphs fall back without orphan MCIDs. Presentation output
keeps `structure: None`. A PDF containing a shown `.notdef` glyph remains
tagged but truthfully omits the PDF/UA identification claim.

**Deviations from the design plan.** The real `feature_showcase` sample exposed
its pre-existing glyph-zero content, so it cannot claim PDF/UA yet. The other
six claiming samples and the in-code fixture pass veraPDF 1.30.2. Audit passes
also strengthened source-order, malformed-graph, artifact, and multipage
ParentTree evidence.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `marked_content_is_backend_neutral_and_non_drawing`,
`tagged_pdf_preserves_heading_and_nested_list_structure`,
`tagged_pdf_marks_table_headers_and_cells`,
`tagged_pdf_carries_figure_alternate_text`, and
`tagging_preserves_visible_pdf_and_raster_output`, plus veraPDF 1.30.2 `ua1`
validation.

**Hash harness.** Fourteen declared changes, `pdf/bytes` and `pdf/pages` for
all seven samples. PNG, PDF resource, and OOXML entries are unchanged.

**Notes for future sessions.** Fix the source glyph-zero content before making
`feature_showcase.pdf` advertise PDF/UA conformance.

### F-174, PDF/A conformance

**Sprint.** S53
**Completed.** 2026-08-23
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The shared PDF writer and native Word and Presentation
facades now expose explicit deterministic PDF/A-2b and PDF/A-3b paths. They
preflight fonts, structure, links, paint, and colour before allocation, then
emit matching XMP, deterministic identifiers, and the bundled sRGB2014 output
intent while retaining tagged structure.

**Non-obvious choices.** Ordinary PDF entry points remain byte-identical.
Archival link annotations set the Print flag. Unsupported tile paint fails with
a named error before output instead of being silently dropped.

**Deviations from the design plan.** Microscope review added the annotation
flag and tile-paint preflight. The PDF/UA XMP extension declaration was also
required by the pinned validator.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `pdfa_profiles_emit_matching_xmp_and_output_intent`,
`pdfa_rejects_prohibited_or_incomplete_features_before_output`,
`ordinary_pdf_api_remains_byte_identical`,
`pdfa_retains_tagged_structure_tree`, and
`pdfa_2b_and_3b_pass_verapdf` under veraPDF 1.30.2.

**Hash harness.** Unchanged by this story, 49 of 49 against the reviewed
integrated baseline.

**Notes for future sessions.** The packaged ICC profile digest is
`384b832de3412066743b52a75ee906b6fb9fb8d9e09e936fc2c43223815c6e0a`.

### F-175, Redaction

**Sprint.** S53
**Completed.** 2026-08-23
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Document::redact_text` performs exact native redaction
across visible and revision Word text, comments, notes, metadata, ChartML
caches, and relationship-resolved embedded workbooks. It handles UTF-8 and
BOM-marked UTF-16 XML, nested packages, entity-normalized content, and
fixed-point matches before a raw residual scan.

**Non-obvious choices.** The operation flushes to a staged package, rewrites
only approved expanded names and raw value spans, reparses and validates every
sensitive part, and publishes only after the outer and nested scans pass.
Python, WASM, and CLI surfaces remain unchanged.

**Deviations from the design plan.** Eleven microscope passes expanded the XML
lexical validator, revision projections, flow boundaries, UTF-16 support,
duplicate relationship checks, fixed-point behavior, and atomic cache and
engine evidence. The public contract stayed within the approved native API.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/09-charts-spec.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** `redaction_rewrites_only_approved_xml_text_and_attributes`,
`redaction_removes_body_comments_revisions_and_metadata_traces`,
`redaction_removes_chart_cache_and_embedded_workbook_traces`,
`redaction_failure_is_atomic`,
`redacted_package_preserves_unrelated_parts_and_relationships`, and
`raw_zip_scan_finds_no_redacted_value`.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep redaction native-only unless a later plan
explicitly designs the policy and atomicity boundary for another adapter.

### F-X048, Dense form table fidelity

**Sprint.** S53
**Completed.** 2026-08-23
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Word table cells now retain source-ordered recursive
paragraph and nested-table blocks. Layout implements grid-span-aware vertical
merges, exact and minimum row rules, style inheritance and conditional layers,
outer nil-border fallback, cell-relative anchors, paragraph-mark metrics, and
bounded transactional cache accounting with provenance rebinding.

**Non-obvious choices.** Exact rows clip cell content while leaving borders
outside the clip. Merged content grows the final eligible non-exact row. The
native paragraph facade adds text that inherits direct paragraph-mark run
properties without synthesizing a glyph for an empty mark.

**Deviations from the design plan.** Microsoft Word was unavailable on the
worker host, so no external Word geometry observation is claimed. The readable
one-page deterministic PDF and raster fixture remains authoritative. Review
added exact clipping, terminal merge borders, conditional shading, character
anchor indent, and mutation-safe preserved style output.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** `nested_tables_remain_recursive_cell_blocks`,
`vertical_merges_and_row_height_rules_share_the_exact_grid_span`,
`table_style_cascade_resolves_borders_and_paragraph_spacing`,
`cell_anchors_use_cell_coordinates_and_page_behind_order`,
`outer_nil_border_matches_word_without_changing_interior_nil`,
`empty_form_paragraphs_use_mark_metrics_and_new_runs_inherit_them`,
`dense_form_matches_reviewed_one_page_geometry`, and
`dense_form_caches_are_transactional_bounded_and_exact`.

**Hash harness.** Two declared changes within the F-173 PDF category,
`feature_showcase:pdf/bytes` and `feature_showcase:pdf/pages`, caused by the
recursive nested table and corrected vertical-merge borders. All other entries
are unchanged from the preceding reviewed baseline.

**Notes for future sessions.** The `CellBlock` recursion is part of cache keys,
retained-byte limits, source mapping, semantic ownership, and painting. Treat
all five paths as one contract when changing table layout.

### F-X049, Tag rpptx-v0.5.0

**Sprint.** S53
**Completed.** 2026-08-23
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete fifteen-package shared OOXML and PowerPoint
family was published at 0.5.0 from reviewed SHA
`343388e19bce21b3d83f17e8cc0e5418861a94cb`. The release contains package
encryption, digital-signature creation and verification, tagged PDF, PDF/A,
and shared immutable font and page ownership. `rpptx-wasm` is prepared at the
same version but remains unpublished.

**Release evidence.** GitHub Actions run
[32654116819](https://github.com/tensorbee/rdocx/actions/runs/32654116819)
passed output stability, metadata, release-note, archive, fifteen-crate
publication, and GitHub Release jobs. Every 0.5.0 registry entry resolved and
listed `mantissaman (Atul Sharma)` as owner. The annotated
[`rpptx-v0.5.0`](https://github.com/tensorbee/rdocx/releases/tag/rpptx-v0.5.0)
tag dereferenced to the reviewed SHA, and the published body was byte-identical
to the committed changelog render.

**Contribution inventory.** Authenticated contributor `@emptinessform`
reported [Issue 39](https://github.com/tensorbee/rdocx/issues/39) and authored
[PR 40](https://github.com/tensorbee/rdocx/pull/40) and
[PR 41](https://github.com/tensorbee/rdocx/pull/41). Their profiling and
reference implementations shaped the shared `FontData` and `PageFrame`
ownership boundary that landed as a hardened equivalent. Format-specific
transfer, pagination, and cache work remains on the stable release train.

**Notifications.** The reviewed release-bound comments were posted and
verified at [Issue 39 comment](https://github.com/tensorbee/rdocx/issues/39#issuecomment-5387476283),
[PR 40 comment](https://github.com/tensorbee/rdocx/pull/40#issuecomment-5387476368),
and [PR 41 comment](https://github.com/tensorbee/rdocx/pull/41#issuecomment-5387476474).

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The incubating 0.5.0 metadata regression, all 66 workflow tests,
full verification, the exact patched 22-package dry run, archive and asset
inventory, both WASM targets, and supply-chain checks passed at the reviewed
source. The release workflow then passed the real publication gate.

**Hash harness.** Unchanged by the release preparation, 49 of 49 against the
integrated reviewed baseline. The sprint baseline retains the fourteen
declared PDF byte and page changes from F-173 and F-X048.

**Notes for future sessions.** Release approval for the incubating family did
not authorize stable `v0.9.0`. F-X050 retains its own exact-SHA approval and
publication boundary.

### F-X050, Tag v0.9.0

**Sprint.** S53
**Completed.** 2026-08-23
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The exact seven-package stable Word family was published
at 0.9.0 from reviewed SHA
`e27e519c94c90cd5be340fe5bf8e431cf542ac51`. The release includes the S52 and
S53 stable layout, package, security, accessibility, conformance, redaction,
and dense-form outcomes. Every Python and WASM package remains unpublished.

**Release evidence.** GitHub Actions run
[32658680024](https://github.com/tensorbee/rdocx/actions/runs/32658680024)
passed output stability, metadata, release-note, archive, seven-crate stable
publication, and GitHub Release jobs. The incubating allowlist was skipped.
Every 0.9.0 registry entry resolved as non-yanked and listed
`mantissaman (Atul Sharma)` as owner. The annotated
[`v0.9.0`](https://github.com/tensorbee/rdocx/releases/tag/v0.9.0) tag
dereferenced to the reviewed SHA, and the published body was byte-identical to
the committed changelog render. Every selected crates.io README endpoint
returned non-empty rendered HTML.

**Contribution inventory.** Issues
[15](https://github.com/tensorbee/rdocx/issues/15) and
[23](https://github.com/tensorbee/rdocx/issues/23) reported by authenticated
`@mantissaman` landed directly. Authenticated `@emptinessform` supplied the
Issue 23 break-opportunity diagnosis and reported or authored
[Issue 39](https://github.com/tensorbee/rdocx/issues/39),
[Issue 42](https://github.com/tensorbee/rdocx/issues/42),
[PR 40](https://github.com/tensorbee/rdocx/pull/40),
[PR 41](https://github.com/tensorbee/rdocx/pull/41), and
[PR 43](https://github.com/tensorbee/rdocx/pull/43). Those five records landed
through reviewed hardened equivalents. PR 43 remains closed and unmerged, with
the exact F-X048 implementation recorded in its closure comment.

**Notifications.** The reviewed release-bound comments were posted and
verified at [Issue 15 comment](https://github.com/tensorbee/rdocx/issues/15#issuecomment-5387845923),
[Issue 23 comment](https://github.com/tensorbee/rdocx/issues/23#issuecomment-5387846237),
[Issue 39 comment](https://github.com/tensorbee/rdocx/issues/39#issuecomment-5387846542),
[PR 40 comment](https://github.com/tensorbee/rdocx/pull/40#issuecomment-5387846918),
[PR 41 comment](https://github.com/tensorbee/rdocx/pull/41#issuecomment-5387847279),
[Issue 42 comment](https://github.com/tensorbee/rdocx/issues/42#issuecomment-5387847592),
and [PR 43 comment](https://github.com/tensorbee/rdocx/pull/43#issuecomment-5387847840).

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The stable 0.9.0 and incubating 0.5.0 metadata regressions, all 66
workflow tests, full verification at the reviewed SHA, the exact clean patched
22-package dry run, archive and asset inventory, both WASM targets,
supply-chain checks, registry and owner queries, release target and body
comparison, README endpoint checks, and exact notification comparisons passed.

**Hash harness.** Unchanged by the stable release preparation, 49 of 49 against
the integrated reviewed baseline. The sprint baseline retains the fourteen
declared PDF byte and page changes from F-173 and F-X048.

**Notes for future sessions.** Stable 0.9.0 is an intentional pre-1.0 Rust
source boundary. The release does not authorize Python, WASM, npm, PyPI, or
later stable-family publication.

### F-176, RTF reader

**Sprint.** S54
**Completed.** 2026-08-24
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `Document::from_rtf_bytes` and `Document::open_rtf` now
read a bounded Word-compatible RTF subset into the existing typed document
tree. The reader covers Unicode and legacy code pages, paragraph and run
formatting, tables, lists, PNG and JPEG pictures, destinations, and stable
loss diagnostics.

**Non-obvious choices.** The scanner uses bounded parser-owned buffers and
group state rather than cloning accumulated content. A pinned Microsoft Word
16.104 conversion defines the structural differential boundary, while the
live regeneration route remains an ignored human oracle check.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** `rtf_reader_matches_the_pinned_word_docx_structure`, the malformed
and resource-bound parser regressions, the code-page and destination tests,
and the full `rdocx` suite passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The reader and writer deliberately share one
RTF module and one typed projection. Unsupported content must remain explicit
in diagnostics rather than becoming silent loss.

### F-177, RTF writer

**Sprint.** S54
**Completed.** 2026-08-24
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `Document::to_rtf_bytes` and `Document::save_rtf` now emit
deterministic bounded RTF for the reader-supported text, formatting, table,
list, and image subset. The result returns stable diagnostics for each lossy
source item, and path writes are atomic.

**Non-obvious choices.** A discovery pass allocates stable font, colour, and
list identifiers before the header is written. Signed UTF-16 controls preserve
supplementary Unicode, formatting groups reset state explicitly, and list
identity is kept through RTF list tables and overrides.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** `rtf_writer_round_trip_preserves_supported_document_content`, the
escaping, deterministic-header, table, list, image, diagnostic, resource-bound,
and atomic-write regressions, and the full `rdocx` suite passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Fidelity is intentionally symmetric with the
reader subset. Extending either side requires extending the shared projection
and the structural round-trip gate together.

### F-183, Image export options

**Sprint.** S54
**Completed.** 2026-08-24
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The shared raster backend now exports selected pages as
opaque or transparent PNG, quality-controlled JPEG, or one multi-page TIFF.
Native Word, Python, and both general CLI export paths expose the options while
the existing PNG entry points retain byte-identical defaults.

**Non-obvious choices.** Separate-page formats retain caller order and stage
outputs before publication. Multi-page TIFF has explicit output cardinality,
JPEG composites alpha over white, and authored page backgrounds still paint
over transparent PNG canvases.

**Deviations from the design plan.** None. Presentation thumbnails remain the
specialized fixed-PNG path approved during design.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** `image_export_options_produce_the_declared_formats_and_exact_pages`,
the transparency, JPEG quality, legacy PNG parity, range, binding, and CLI
regressions, both WASM checks, and the golden PNG gate passed.

**Hash harness.** Unchanged, 49 of 49. Golden page-one pixels also remained 7
of 7.

**Notes for future sessions.** TIFF and JPEG encoding belongs in `oxml-pdf` so
Word and PresentationML share validation, page selection, and memory behavior.

### F-X051, Honor caller-supplied font family aliases

**Sprint.** S54
**Completed.** 2026-08-24
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Caller-supplied font bytes can now carry bounded family
aliases. Resolution tries the embedded face and its aliases before existing
mapped and generic fallbacks, and reusable layout contexts preserve exact
warm and cold output identity.

**Non-obvious choices.** Alias state is private, byte-free, deduplicated, and
bounded by entry and byte ceilings. Cache identity includes the normalized
alias mapping, so unchanged aliases reuse work while changed aliases invalidate
only the affected resolution state.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** `document_facing_aliases_share_one_caller_font`, the alias-bound,
font-priority, changed-context, and warm-cold identity regressions, the
no-default-features path, and both WASM checks passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Alias labels never duplicate font bytes and
must remain part of reusable-engine cache identity whenever their normalized
mapping changes.

### F-178, HTML import

**Sprint.** S55
**Completed.** 2026-08-24
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `Document::from_html` and `Document::open_html` now import
a bounded HTML5 and CSS subset into the existing Word document tree. The
importer preserves source order across text,
paragraphs, nested inline formatting, lists, spanned tables, whitespace, and
hard breaks, with stable diagnostics for unsupported visible content and CSS.

**Non-obvious choices.** One private module owns HTML parsing, CSS resolution,
and Word projection. The browser-grade parser performs HTML5 repair, while
explicit construction and projection bounds prevent repaired or hostile input
from expanding without limit. External resources are diagnosed rather than
fetched.

**Deviations from the design plan.** None. Microscope review found seven
boundedness, preservation, and compatibility defects across two passes. All
were remediated before the clean third pass.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The six source-built HTML gates cover malformed input, supported
formatting, nested lists, spanned tables, unsupported CSS with retained
siblings, file bounds, save, and reopen. The full `rdocx` suite, Rust 1.93,
WASM, dependency, package, and supply-chain gates passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The importer materializes effective formatting
instead of retaining a second HTML model. New supported CSS must extend the
diagnostic and source-order gates at the same time.

### F-179, ODT reader

**Sprint.** S55
**Completed.** 2026-08-24
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `Document::from_odt_bytes`,
`Document::from_odt_bytes_with_limits`, and `Document::open_odt` now read a
bounded OpenDocument Text subset into the existing Word document tree. The
reader covers effective text and paragraph formatting, nested lists, spanned
tables, and inline package images with stable loss diagnostics.

**Non-obvious choices.** ODT archive handling remains separate from OPC
because ODT is ZIP-based but not an OOXML package. The complete archive index,
paths, compression, encryption state, XML depth, projected content, and media
are validated before publication. A source-built fixture is compared against
the exact pinned LibreOffice conversion structurally, not byte for byte.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The source-built ODT gates cover archive and XML bounds, effective
styles, lists, table spans, images, diagnostics, save, and reopen. The exact
LibreOffice 26.2.5.2 structural differential, full `rdocx` suite, WASM,
package, and supply-chain gates passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Unknown ODT content is diagnosed and consumed
without dropping supported siblings. It is not smuggled into the DOCX package
as unowned foreign XML.

### F-X052, Restore interactive relayout performance

**Sprint.** S55
**Completed.** 2026-08-24
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Reusable Word layout now avoids eager context clones,
whole-body debug serialization, repeated cache-hit block copies, retained page
deep copies, and linear exact shaping comparisons for most candidates. Private
shared paragraph and table blocks carry result-local provenance and semantic
structure through an overlay, while cache-safe tables participate in restart
pagination.

**Non-obvious choices.** Fingerprints are prefilters only. Exact typed equality
still decides every cache hit and transfer. Public layout block APIs remain
unchanged, cache publication remains transactional, and the 64 MiB aggregate
budget is preserved with 50, 2, 4, and 8 MiB partitions for paragraph, table,
header and footer, and restart state.

**Deviations from the design plan.** None. The approved plan already includes
the evidence-driven 50, 2, 4, and 8 MiB cache partition and the collision-safe
fingerprint prefilter. Neither changes the aggregate bound or public API.

**Spec sections touched.** `docs/hld/08-rendering-spec.md` and
`docs/hld/12-testing-strategy.md`.

**Tests.** The mixed 700-paragraph and 14-table editor gate proves exact warm
and cold structure, provenance, semantic content, fonts, diagnostics,
outlines, collision safety, checked transfer, transactionality, and memory
bounds. The full layout and facade suites, no-default layout, both WASM targets,
and full workspace verification passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Fresh integrated A/B runs produced 58 pages in
every case. Native cold and warm ratios topped out at 1.04 and 0.89. Bundled
fallback cold, typing, checked undo, and table mutation each remained at or
below 1.14, inside the required 1.25 budget.

### F-X053, Complete layout migration and contribution records

**Sprint.** S55
**Completed.** 2026-08-24
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The v0.9.0 compatibility record and layout package README
now tell external backends to recurse through `MarkedContent::children` or use
`oxml_layout::walk`. The next-stable contribution inventory records Issues 44
and 46 and PR 45 with authenticated contributor credit.

**Non-obvious choices.** The published v0.9.0 release body was replaced only
after proving byte equality with the tracked 4,111-byte render. Its tag target
and empty asset set remained unchanged. Issues 39 and 42 received no duplicate
comments because their existing acceptance evidence already closed their
scope.

**Deviations from the design plan.** None. GitHub item 46 is an issue, so the
related contributor pull request is PR 45 under GitHub's shared numbering.

**Spec sections touched.** `docs/hld/10-bindings-spec.md`.

**Tests.** Release-note rendering, package documentation, three workflow
regressions, exact published-body equality, authenticated GitHub state, the
69-test workflow suite, full workspace verification, and both WASM targets
passed. Issue 44, PR 45, and Issue 46 were closed with exact implementation
evidence and maintainer-authenticated comments.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The closure comments are retained at
[Issue 44](https://github.com/tensorbee/rdocx/issues/44#issuecomment-5395173425),
[PR 45](https://github.com/tensorbee/rdocx/pull/45#issuecomment-5395175132), and
[Issue 46](https://github.com/tensorbee/rdocx/issues/46#issuecomment-5395176825).

### F-180, ODT writer

**Sprint.** S56
**Completed.** 2026-08-25
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `Document::to_odt_bytes` and `Document::save_odt` now write
a deterministic bounded ODF 1.3 package from the native Word tree. The writer
projects supported effective paragraph and text formatting, nested lists,
table spans, and inline images, and returns stable path-aware diagnostics for
every unsupported or simplified source fact.

**Non-obvious choices.** ODT remains a private `rdocx` facade concern rather
than an OPC mode. The package writes an uncompressed first `mimetype` entry,
schema-ordered XML, deterministic automatic styles and media names, and an
exact manifest. Byte serialization completes before failure-atomic path
replacement, and write limits are expressed in the units the F-179 reader
will observe after reopening.

**Deviations from the design plan.** None. Review expanded the conformance and
loss matrices, limit proofs, inherited numbering coverage, and exact F-179
piece-to-run accounting without changing the approved public boundary.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Twenty-nine focused writer tests and the public structural round-trip
gate passed. The exact LibreOffice 26.2.5.2 differential, full workspace suite,
WASM, package, documentation, and supply-chain gates also passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** F-179 and F-180 share one declared fidelity
boundary. Extending either direction requires extending the structural record,
loss diagnostics, and bounded package checks together.

### F-181, EPUB export

**Sprint.** S56
**Completed.** 2026-08-25
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The native Word facade now exports deterministic EPUB 3
bytes and files with ordered spine and navigation documents, structured
headings, nested lists, tables, safe hyperlinks, validated raster media, and
bounded metadata. Unsupported or simplified document facts remain visible as
stable diagnostics.

**Non-obvious choices.** EPUB projection is private to `rdocx` and reuses the
existing document tree rather than introducing a second retained model. Media
must pass byte-level PNG, JPEG, or GIF structure validation before packaging.
Heading anchors remain source-correlated, list continuation remains explicit,
and every XHTML, navigation, media, and diagnostic budget is checked before
unbounded expansion.

**Deviations from the design plan.** None. Review strengthened bounded
projection, URI and media validation, style and numbering loss reporting,
table-cell lists, metadata diagnostics, and exact source correlation.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Thirty-three focused exporter tests passed. The source-built
publication passed the exact checksum-pinned EPUBCheck 5.3.0 oracle. Full
workspace, WASM, package, documentation, and supply-chain gates also passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** EPUB output owns document semantics and reading
order rather than fixed-page geometry. New media or XHTML support must extend
the strict validator, bounded preflight, and EPUBCheck fixture together.

### F-182, SVG page export

**Sprint.** S56
**Completed.** 2026-08-25
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The native Word facade now exports one deterministic,
self-contained SVG string from a selected `PageFrame`. Text stays searchable,
fonts and raster images are embedded, safe links remain active, and recursive
geometry, gradients, clipping, opacity, marked content, and supported shadows
retain shared layout order.

**Non-obvious choices.** The backend lives privately in stable `rdocx` because
the shared PDF family remains at its published incubating boundary. Definition
IDs and diagnostics follow deterministic first use. Complex text remains text
with an approximation diagnostic when exact scalar positioning is impossible,
and an unprovable singular-transform effect is omitted without dropping its
source content or siblings.

**Deviations from the design plan.** None. Review replaced the initial narrow
golden with representative text, image, gradient, clip, shadow, link, and
three-level noncommuting transform coverage, and tightened font, XML, gradient,
and filter behavior.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Nineteen focused SVG tests passed, including the calibrated 150 dpi
resvg 0.48.1 SSIM gate and its perturbation control. Full workspace,
no-default-font, WASM, package, documentation, and supply-chain gates passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** SVG consumes immutable layout output. It must
not grow Word format knowledge or silently flatten searchable text into
outlines.

### F-X054, Integrate PRs 47 through 52

**Sprint.** S56
**Completed.** 2026-08-25
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Borrowed ordered readers now expose direct body, cell,
paragraph, hyperlink, and run children while established flattened accessors
remain unchanged. Unsupported XML facts retain qualified names, resolved
namespaces, child facts, and optional exact bytes. Producer-defined numbering
values survive in `ST_NumberFormat::Other(String)`, and undecodable ordinary or
deleted Word text fails closed.

**Non-obvious choices.** The six external proposals landed as hardened
equivalents rather than direct merges. Public item enums are non-exhaustive,
modeled unsupported facts do not fabricate raw bytes, namespace facts come
from parser scope, and namespace replay uses declaration-dependent logical
owner identity with safe unchanged-byte preservation and fail-closed mutation.
Exporters never invent decimal markers for producer-defined numbering.

**Deviations from the design plan.** All six outcomes are hardened equivalents.
The implementation adds current namespace, non-exhaustive API, bounded
allocation, exact raw-boundary, and diagnostic protections beyond the proposed
patches. The S55 Python `Html` and `Odt` error reconciliation was recorded
separately from the numbering outcome.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** The complete ordered-reader and namespace regression suite passed
168 tests before integration, and the integrated suite passed 169. The pinned
Python 3.12.9 and maturin 1.13.3 binding environment passed all 37 tests. Full
workspace, package, documentation, WASM, and supply-chain gates passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The included outcomes came from
[PR 47](https://github.com/tensorbee/rdocx/pull/47),
[PR 48](https://github.com/tensorbee/rdocx/pull/48),
[PR 49](https://github.com/tensorbee/rdocx/pull/49),
[PR 50](https://github.com/tensorbee/rdocx/pull/50),
[PR 51](https://github.com/tensorbee/rdocx/pull/51), and
[PR 52](https://github.com/tensorbee/rdocx/pull/52), with specific credit to
`@pedroassumpcao`. F-X055 owns their release-bound comments and closure after
v0.10.0 publication verifies.

### F-X056, Tag rpptx-v0.6.0

**Sprint.** S56
**Completed.** 2026-08-25
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete fifteen-package shared OOXML and PowerPoint
family was published at 0.6.0 from reviewed SHA
`55fb2f54caf91d7dedc8936b4c7b116354590628`. The release includes the shared
font-alias and bounded relayout outcomes, raster export additions, and current
Presentation CLI output selection. `rpptx-wasm` remains unpublished.

**Release evidence.** GitHub Actions run
[32866396976](https://github.com/tensorbee/rdocx/actions/runs/32866396976)
passed output stability, metadata, reviewed-note, archive, fifteen-crate
publication, and GitHub Release jobs. Every selected 0.6.0 registry entry
resolved under sole owner `mantissaman (Atul Sharma)`. The annotated
[`rpptx-v0.6.0`](https://github.com/tensorbee/rdocx/releases/tag/rpptx-v0.6.0)
tag dereferenced to the reviewed SHA, and its 2,502-byte body was
byte-identical to the committed changelog render.

**Contribution inventory.** Authenticated contributor `@emptinessform`
reported [Issue 44](https://github.com/tensorbee/rdocx/issues/44), authored
[PR 45](https://github.com/tensorbee/rdocx/pull/45), and reported
[Issue 46](https://github.com/tensorbee/rdocx/issues/46). F-X051 and F-X052
landed all three selected outcomes as hardened equivalents. No record state
changed during this release.

**Notifications.** The reviewed comments were posted and verified at
[Issue 44 comment](https://github.com/tensorbee/rdocx/issues/44#issuecomment-5413050520),
[PR 45 comment](https://github.com/tensorbee/rdocx/pull/45#issuecomment-5413050758),
and [Issue 46 comment](https://github.com/tensorbee/rdocx/issues/46#issuecomment-5413050988).

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The full workspace gate, 70 workflow tests, 49-entry hash harness,
no-default-font tests, both WASM targets, warning-free docs, README checks,
exact 22-package dry run, archive inventory and size checks, and cargo-deny
passed at the reviewed source. The publication workflow and independent
registry, owner, tag, release-body, and notification checks then passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve the immutable `rpptx-v0.6.0` tag.
F-X057 owns stable 0.10.1 and retains its separate exact-SHA final approval.

### F-X057, Tag v0.10.1

**Sprint.** S56
**Completed.** 2026-08-25
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete seven-package stable Word family was
published at 0.10.1 from reviewed SHA
`ae0dcb162a7805e59e5890464b226765645ad547`. The release recovers the full S56
stable outcome after the immutable v0.10.0 attempt published only `rdocx-opc`
and `rdocx-oxml`.

**Release evidence.** GitHub Actions run
[32879293813](https://github.com/tensorbee/rdocx/actions/runs/32879293813)
passed the exact seven-package publication and GitHub Release jobs. Every
selected 0.10.1 registry entry resolved under sole owner
`mantissaman (Atul Sharma)`. The annotated
[`v0.10.1`](https://github.com/tensorbee/rdocx/releases/tag/v0.10.1) tag
dereferenced to the reviewed SHA, and its 5,542-byte body was byte-identical to
the committed changelog render.

**Contribution inventory.** Authenticated contributor `@emptinessform`
reported [Issue 44](https://github.com/tensorbee/rdocx/issues/44), authored
[PR 45](https://github.com/tensorbee/rdocx/pull/45), and reported
[Issue 46](https://github.com/tensorbee/rdocx/issues/46). Authenticated
contributor `@pedroassumpcao` authored
[PR 47](https://github.com/tensorbee/rdocx/pull/47) through
[PR 52](https://github.com/tensorbee/rdocx/pull/52). F-X051, F-X052, and F-X054
landed all nine selected outcomes as hardened equivalents.

**Notifications.** The exact reviewed comments were posted by `mantissaman`
and verified at
[Issue 44 comment](https://github.com/tensorbee/rdocx/issues/44#issuecomment-5414486370),
[PR 45 comment](https://github.com/tensorbee/rdocx/pull/45#issuecomment-5414486381),
[Issue 46 comment](https://github.com/tensorbee/rdocx/issues/46#issuecomment-5414486391),
[PR 47 comment](https://github.com/tensorbee/rdocx/pull/47#issuecomment-5414486414),
[PR 48 comment](https://github.com/tensorbee/rdocx/pull/48#issuecomment-5414486360),
[PR 49 comment](https://github.com/tensorbee/rdocx/pull/49#issuecomment-5414486396),
[PR 50 comment](https://github.com/tensorbee/rdocx/pull/50#issuecomment-5414486392),
[PR 51 comment](https://github.com/tensorbee/rdocx/pull/51#issuecomment-5414486351),
and [PR 52 comment](https://github.com/tensorbee/rdocx/pull/52#issuecomment-5414486404).
PRs 47 through 52 were then closed unmerged as authorized.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The full workspace gate, 74 workflow tests, 49-entry hash harness,
no-default-font tests, both WASM targets, pinned Python binding tests,
warning-free docs, README checks, exact 22-package dry run, archive inventory
and size checks, and cargo-deny passed at the reviewed source. Publication and
independent registry, owner, tag, release-body, notification, and closure
checks then passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve the immutable v0.10.0 partial release
and the complete `v0.10.1` release. Binding, WASM, npm, and PyPI publication
authority remains unchanged.

### F-196, Word corpus

**Sprint.** S57
**Completed.** 2026-08-25
**Size.** M, estimated 2 days, actual 1 day

**What was built.** A strict fetched Word corpus now covers a business letter,
report, form, legal document with revisions, and multi-script text. The
manifest binds each input to an immutable source revision, SHA-256 digest,
category, SPDX licence, and immutable licence URL. The fetcher stages downloads
atomically and its read-only check rejects missing, extra, changed, unsafe,
misclassified, or unlicensed inputs.

**Non-obvious choices.** Four fixtures come from one pinned Apache POI revision
under Apache-2.0. The legal redline comes from a pinned docx-mcp revision under
MIT because it proves real contract text with both insertions and deletions.
The corpus stays ignored and outside every published crate archive. Test and
MSRV CI fetch and verify it before Cargo runs.

**Deviations from the design plan.** None. Review tightened the manifest and
workflow mutation checks without changing the approved five-category boundary.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`, the external Word
corpus contract, and `docs/hld/15-build-and-toolchain.md`, corpus acquisition
and CI ownership.

**Tests.** `test_word_corpus_fetcher_verifies_every_checksum_and_refuses_a_mismatch`,
`test_word_corpus_fetcher_refuses_missing_extra_and_unlicensed_inputs`,
`test_workspace_test_jobs_fetch_the_pinned_word_corpus`, and the live fetcher
`--check` all passed for exactly five documents.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Extend the manifest and all category-count
assertions together. Do not replace immutable source or licence URLs with
branch links, and never place fetched binaries in a crate archive.

### F-201, Large document performance

**Sprint.** S57
**Completed.** 2026-08-25
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The existing `rdocx` regression binary now source-builds
an exact thousand-page document, lays it out with deterministic bundled fonts,
and renders the returned layout directly to PDF. The ignored release gate
asserts page count, nonempty output, layout and PDF throughput floors, and
separate peak live heap ceilings.

**Non-obvious choices.** A generation-tagged test allocator measures only
allocations owned by the active stage, so deallocations from setup or a prior
stage cannot corrupt the result. The reviewed limits are 64 MiB and 250 pages
per second for layout, then 16 MiB additional peak and 1,000 pages per second
for PDF. CI runs the exact locked release test serially on Ubuntu 24.04.

**Deviations from the design plan.** None. Calibration confirmed that no
production optimization was needed. The consolidated run measured 64,528
layout pages per second at 22.32 MiB and 76,831 PDF pages per second at 1.73
MiB.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`, the numeric
large-document contract, `docs/hld/12-testing-strategy.md`, the named release
regression, and `docs/hld/15-build-and-toolchain.md`, its CI execution.

**Tests.** `a_thousand_page_document_paginates_and_renders_within_the_declared_limits`
passed in locked release mode with one test thread. Workflow mutation tests
also proved that CI cannot drop, weaken, parallelize, or swallow the gate.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The PDF stage consumes the already returned
layout, so its allocator number is the additional render peak rather than total
document retention. Any production optimization prompted by a future failure
needs fresh layout and PDF risk routing.

### F-197, Word SSIM harness

**Sprint.** S57
**Completed.** 2026-08-25
**Size.** L, estimated 4 days, actual 1 day

**What was built.** A path-filtered Word fidelity job now renders every pinned
document through the deterministic `rdocx-cli` production path and compares it
at 150 dpi with accepted-view LibreOffice Writer output rasterized by Poppler.
The harness records every page in the union, per-page luminance SSIM, source
dimensions, normalization actions, page-count differences, summary statistics,
and exact tool identities in retained TSV and JSON evidence.

**Non-obvious choices.** Unequal dimensions are placed on one shared white
canvas, and a page produced by only one renderer is compared with a blank white
counterpart. Those differences remain scored evidence instead of disappearing
as orchestration failures. A temporary locked and offline helper uses the
existing `Document::accept_all()` API, saves, reopens, and verifies the legal
fixture before Writer sees it. SSIM at least 0.95 on 80 percent of pages is an
advisory trend. Corpus, tool, render, oracle, zero-output, and evidence failures
remain hard failures.

**Deviations from the design plan.** The user approved a material plan revision
after the first calibration showed 18 Rust pages and 16 oracle pages with
one-pixel dimension differences. Review then found that a fresh Writer profile
did not guarantee accepted revision view, so the final plan and implementation
made acceptance explicit through the existing Rust API.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`, complete union-page
comparison, `docs/hld/12-testing-strategy.md`, Word SSIM evidence and hard-gate
policy, and `docs/hld/15-build-and-toolchain.md`, pinned `word-fidelity` CI.

**Tests.** Seventeen harness self-tests covered exact SSIM, alpha compositing,
accepted-view preparation, tool pins, union coverage, white-canvas
normalization, required artifacts, trend classification, and a one-pixel
perturbation. Eighty workflow regressions passed. The consolidated live gate
covered five documents and 18 union pages with exact LibreOffice 26.2.5.2 and
Poppler 26.01.0 identities.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The consolidated baseline records 18 Rust pages,
16 Writer pages, 14 dimension normalizations, two Rust-only pages, and one of
18 pages at SSIM 0.95 or greater. The 5.56 percent trend is deliberately
visible but advisory. Future shaping stories should improve or intentionally
explain this evidence without weakening complete-union coverage.

### F-202, Incremental layout

**Sprint.** S58
**Completed.** 2026-08-26
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Restart pagination now retains up to 1,024 page and
checkpoint entries under the existing 8 MiB restart and 64 MiB aggregate byte
ceilings. Editing one paragraph in the source-built thousand-page document
rebuilds at most two pages through both the engine and public facade paths.

**Non-obvious choices.** The retained partition keeps exact body, provenance,
font, substitution, tail-context, and suffix equality. Oversized or unsafe
state still falls back to full pagination. Substituted-page reuse stays bounded
at the 1,024 and 1,025 page boundary, and retained page frames use shared
`Arc` identity without changing the public API.

**Deviations from the design plan.** None. Microscope tightened the invocation
gate so it proves 1,000 initial page layouts and only one or two warm page
layouts instead of accepting a zero-work false positive.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, and `docs/hld/15-build-and-toolchain.md`.

**Tests.** The engine and facade thousand-page restart regressions,
substituted-page boundary equality, the existing locked F-201 release gate,
and the dependency-prefix full workspace gate passed. The latter included
no-default fonts, both WASM targets, warning-free docs, 27 README inventories,
the exact 22-package dry run, archive size checks, and cargo-deny.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** The 1,024 entry ceiling is an explicit bounded
editor contract. F-X062 may widen restart eligibility for unchanged related
stories, but it must preserve the same byte ceilings and exact fresh equality.

### F-X061, Support staged dependency checkpoints in run-sprint

**Sprint.** S58
**Completed.** 2026-08-26
**Size.** S, estimated 1 day, actual 1 day

**What was built.** `/run-sprint` now finalises an integrated dependency prefix
before claiming its consumer, then resumes the same sprint state. Release
dependencies extend that route with prepared-release evidence, separate final
approval, verified publication, and post-publication delivery records.

**Non-obvious choices.** A clean review file is committed first, the clean
verdict is recorded at that resulting HEAD, and full verification is rerun at
the same HEAD. No self-confirming review is added for the evidence-only commit.
Each scheduled boundary keeps its own remediation bound while global pass
numbers remain monotonic. `init --resume` refreshes canonical title and size
metadata while preserving state, ownership, worker, review, and verification
facts.

**Deviations from the design plan.** The approved release-only repair was
generalised before integration after an independent A to B to C forward test
proved that ordinary dependencies had the same finalisation deadlock.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Eighty-four workflow regressions and an independent forward test
covered ordinary A to B to C chains, release R to D chains, review commit
ordering, and resume preservation. The dependency-prefix full workspace gate
also passed with all package, documentation, WASM, hash, and supply-chain
checks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Use the dependency-prefix route before every
consumer whose formal prerequisite is reviewed but not completed. Earlier
HEAD-bound evidence remains historical and cannot satisfy a later boundary.

### F-X062, Reuse restart pagination with notes and headers

**Sprint.** S58
**Completed.** 2026-08-26
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Restart pagination now reuses unchanged footnote, endnote,
header, and footer context at note-clean page boundaries. The reporter-scale
700-paragraph facade workload retains all but at most two pages while remaining
exactly equal to a fresh layout.

**Non-obvious choices.** Exact related-story and body note-reference equality
is required before reuse. Note-bearing tables and other traversal-sensitive
content keep conservative full pagination. Restarted completion appends
endnotes exactly once with final page numbers offset by the retained prefix,
while exact cached tails keep their attached endnote pages.

**Deviations from the design plan.** Microscope pass 1 found that restarted
completion appended endnotes before accounting for the retained prefix. The
remediation added a no-cached-tail regression and supplies the full selected
note sequence plus retained-prefix page offset only on that completion path.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Five planned engine and facade regressions, the restarted-completion
endnote regression, 175 `rdocx-layout` tests, the full `rdocx` and workspace
suites, the F-201 release performance gate, no-default fonts, both WASM
targets, documentation, README, package, archive-size, and dependency-policy
gates passed. Microscope pass 2 reported zero defects and zero smells.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Restart checkpoints remain note-clean. Changed
related stories, changed note-reference sequences, and note-bearing tables are
intentional full-pagination boundaries.

### F-X063, Avoid duplicate caller-font byte comparisons

**Sprint.** S58
**Completed.** 2026-08-26
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Normal warm layout now skips the second retained-context
font byte equality after `FontManager::load_additional_fonts` has already
proved the ordered caller-font family and byte set unchanged. The measured
duplicate work for the 22 MiB reporter workload falls from 23,068,672 bytes to
zero.

**Non-obvious choices.** The optimization is private and applies only after the
font manager's authoritative exact comparison. Checked engine transfer still
uses the complete ordered family and byte comparison, and equal-length changed
font bytes still invalidate reusable work.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Three exact byte-accounting controls and the public five-font,
40-alias regression passed with complete warm and fresh equality. All 178
`rdocx-layout` tests, full `rdocx` and workspace suites, the F-201 release
performance gate, no-default fonts, both WASM targets, documentation, README,
package, archive-size, and dependency-policy gates passed. Microscope pass 1
reported zero defects and zero smells.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Do not replace the exact comparison with family
and length checks or a full-input hash. Any new warm path must prove the font
manager has already accepted the exact bytes before using the font-elided
retained-context comparison.

### F-X058, Shared multilingual text substrate

**Sprint.** S58
**Completed.** 2026-08-26
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The shared layout family now provides additive multilingual
shaping, cluster-safe line breaking, conditional hyphenation, paragraph
direction, and line-local bidi ordering. DrawingML direction reaches the rich
PowerPoint path, and PDF, raster, and SVG consume validated positioned glyphs
while preserving logical searchable text. Deterministic Noto Arabic,
Devanagari, Thai, and reproducibly subset CJK fonts include complete licence
and provenance records.

**Non-obvious choices.** Existing exhaustive public structs and legacy Latin
entry points keep their exact shapes and behavior. New direction information
travels through a sibling sidecar. Rich shaping is paragraph-wide across
forced breaks, line layout applies UAX 9 L1 before L2, and malformed public
rich runs fail safely in every backend.

**Deviations from the design plan.** Six microscope passes tightened the
approved implementation without changing scope. Remediation preserved legacy
struct literals, moved direction to an additive carrier, made bidi resolution
paragraph-wide, applied line-local whitespace resets at cluster granularity,
and aligned SVG validation with PDF and raster. The final pass reported zero
defects and zero smells.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/05-drawingml-model.md`, `docs/hld/08-rendering-spec.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/11-migration-plan.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The conditional-hyphen, Arabic, Indic, Thai, CJK, bidi,
DrawingML-direction, rich PowerPoint output, legacy source-compatibility, and
font-inventory regressions pass. Full workspace tests, no-default layout, both
WASM targets, rustdoc, 27 README inventories, the exact 22-package dry run,
archive limits, and cargo-deny also pass.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** F-X059 must publish the complete 0.7.0 shared
family before stable Word consumers opt into this rich path. F-198 owns the
first declared rendering-baseline movement.

### F-X059, Tag rpptx-v0.7.0

**Sprint.** S58
**Completed.** 2026-08-27
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete fifteen-package shared OOXML and PowerPoint
family was published at 0.7.0 from reviewed SHA
`1b076c16fb494fe47b054d761e061181a1ea0b15`. The release publishes the shared
multilingual shaping, breaking, direction, deterministic font, and rich
backend substrate. `rpptx-wasm` remains unpublished.

**Release evidence.** GitHub Actions run
[33049354630](https://github.com/tensorbee/rdocx/actions/runs/33049354630)
passed output stability, release metadata, reviewed notes, archive
verification, the exact fifteen-crate publication, and GitHub Release jobs.
Every selected 0.7.0 registry entry resolved under owner
`mantissaman (Atul Sharma)`. The annotated
[`rpptx-v0.7.0`](https://github.com/tensorbee/rdocx/releases/tag/rpptx-v0.7.0)
tag dereferenced to the reviewed SHA, and its 2,529-byte body was
byte-identical to the committed changelog render.

**Contribution inventory.** The selected F-X058 substrate added no new
authenticated external issue or pull-request record after `rpptx-v0.6.0`.
PRs 55 through 57 remain open and belong to later F-X064 through F-X066
stable-reader stories, so they are not part of this release inventory.

**Notifications.** None. The reviewed contribution inventory was empty, and
no issue or pull-request state changed during this release.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Full verification passed at the reviewed SHA with all workspace,
no-default, WASM, documentation, README, exact 22-package dry-run, archive,
and cargo-deny gates. The hash harness remained unchanged at 49 of 49. The
publication workflow and independent registry, owner, tag, release-target,
and byte-for-byte release-body checks then passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve the immutable `rpptx-v0.7.0` tag.
F-X064 through F-X066 may now land the three open contributor PR outcomes,
and F-198 through F-200 may consume the published shared 0.7.0 boundary.

### F-X064, Accept whole-valued decimal table measurements

**Sprint.** S58
**Completed.** 2026-08-27
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Word table widths, cell widths, table indents, and default
cell margins now accept whole-valued decimal lexical forms such as `9345.0`
through one exact string parser. Fractional decimals, exponent forms, overflow,
percentages, universal measures, and malformed values return explicit errors
instead of becoming zero. Serialization retains the canonical integer form.

**Contribution evidence.** This is a hardened equivalent of
[PR 55](https://github.com/tensorbee/rdocx/pull/55) from authenticated
contributor `@pedroassumpcao`, reviewed at source SHA
`056d48fdf23f35e3538ef3d6ff78cf9e3863e3a5`. The original pull request remains
open and unchanged.

**Non-obvious choices.** The existing signed integer public projection remains
unchanged. Valid percentage and universal-measure union arms are reported as
unsupported until a separate lossless model is designed. Parsing is
namespace-aware and does not use floating point.

**Deviations from the design plan.** None. Microscope pass 1 reported zero
defects, zero smells, and zero nitpicks.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** Four focused lexical, namespace, negative, and round-trip
regressions passed. The full workspace, no-default, WASM, rustdoc, README,
22-package dry-run, archive-size, and supply-chain gates passed. The pinned
five-document Word corpus produced all 18 expected evidence pages. Its
advisory SSIM trend remained 1 of 18 pages at or above 0.95.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep unsupported union arms explicit. Do not
restore malformed-to-zero behavior or replace exact lexical parsing with
floating-point conversion.

### F-X067, Prime Word fidelity Cargo dependencies

**Sprint.** S58
**Completed.** 2026-08-27
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The Word fidelity job now runs one named
`cargo fetch --locked` step immediately after the pinned Rust cache restore and
before corpus setup or the locked offline harness. Existing workflow tests
prove the step is present exactly once, locked, and in the correct job and
position.

**Contribution evidence.** This is a hardened equivalent of
[PR 58](https://github.com/tensorbee/rdocx/pull/58) from authenticated
contributor `@pedroassumpcao`, reviewed at source SHA
`c8fed1d1268fd765d602bac2da6524900c1c1cfd`. The original pull request remains
open and unchanged.

**Non-obvious choices.** Dependency priming is the only network boundary. The
fidelity helper remains locked and offline, so a cold hosted runner cannot
silently resolve a different graph while producing oracle evidence.

**Deviations from the design plan.** None. Microscope pass 1 reported zero
defects, zero smells, and zero nitpicks.

**Spec sections touched.** `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The complete 86-test workflow module passed, including missing,
unlocked, duplicate, misplaced, and wrong-job mutations. The full workspace,
no-default, WASM, rustdoc, README, 22-package dry-run, archive-size, and
supply-chain gates passed. The integrated pinned Word gate fetched the locked
graph first and produced nonempty JSON and TSV evidence for five documents and
18 union pages.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve the single explicit network boundary
before the offline helper. The integrated hosted Word job remains a sprint
completion rider.

### F-X065, Expose tracked table grid changes

**Sprint.** S58
**Completed.** 2026-08-27
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Table-grid parsing now recognizes `w:tblGrid`,
`w:gridCol`, and `w:tblGridChange` by namespace URI. One historical grid
subtree is preserved exactly and written after active columns, duplicates fail
closed, and foreign same-local children remain raw. The public table facade can
report the presence of tracked grid history, while layout continues to use
only active columns.

**Contribution evidence.** This is a hardened equivalent of
[PR 56](https://github.com/tensorbee/rdocx/pull/56) from authenticated
contributor `@pedroassumpcao`, reviewed at source SHA
`8b79c4cd0452defafe0a58e86b332c98e7fe52d7`. The original pull request remains
open and unchanged.

**Non-obvious choices.** Historical grid XML remains a preservation sidecar
and never changes active layout widths. Captured historical and foreign
subtrees receive the ancestor namespace bindings they need to serialize as
self-contained XML. Adding the optional historical field to the pre-1.0 public
`CT_TblGrid` literal is an intentional source compatibility change for v0.11.0.

**Deviations from the design plan.** Microscope pass 1 found that raw grid
subtrees could lose namespace bindings declared only by ancestors. The
remediation reused the existing owner-binding normalization and added a
save, reparse, and repeated-serialization regression. Microscope pass 2
reported zero defects, zero smells, and zero nitpicks.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** Six focused namespace, preservation, duplicate, facade, and
active-layout regressions passed. The full `rdocx-oxml`, `rdocx-layout`,
`rdocx`, and workspace suites passed. The no-default, WASM, rustdoc, README,
22-package dry-run, archive-size, and supply-chain gates passed. The pinned
five-document Word corpus produced nonempty evidence for all 18 union pages.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep historical grid data inert for layout and
preserve its namespace context. Do not accept a second modeled grid change by
discarding bytes.

### F-X066, Classify legacy VML horizontal rules

**Sprint.** S58
**Completed.** 2026-08-28
**Size.** S, estimated 1 day, actual 1 day

**What was built.** Native run inspection now classifies the narrow legacy
horizontal-rule form as `RunItemRef::LegacyHorizontalRule` and exposes its
exact raw XML. Classification resolves WordprocessingML, VML, and Office names
by namespace URI, accepts only the documented true lexical forms, and leaves
all foreign, malformed, ambiguous, or visible-content cases unsupported. No
layout or renderer draws the rule.

**Contribution evidence.** This is a hardened equivalent of
[PR 57](https://github.com/tensorbee/rdocx/pull/57) from authenticated
contributor `@pedroassumpcao`, reviewed at source SHA
`44498f042a2290ef40c7a6c26025f38e38e9ce2a`. The original pull request remains
open and unchanged.

**Non-obvious choices.** Classification occurs once at the OXML parse boundary
and is encoded in the existing raw-position sidecar. This preserves the
published `CT_R` struct-literal shape, makes equality reflect observable
classification, and avoids retaining namespace-scope allocations on ordinary
runs. `RunItemRef` was already non-exhaustive, so the new variant is additive.

**Deviations from the design plan.** Microscope pass 1 found that reparsing a
raw child lost namespace bindings declared by ancestors. The first remediation
threaded namespace scope into `CT_R`, but microscope pass 2 found a published
struct-literal break, equality drift, and eager scope cloning. The final
parse-boundary sidecar design remediated all three findings. Microscope pass 3
reported zero defects, zero smells, and zero nitpicks.

**Spec sections touched.** `docs/hld/04-opc-and-packaging.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** URI-aware positive, negative, raw-order, equality, allocation, and
package save-reopen regressions passed. The full workspace, no-default, WASM,
rustdoc, README, 22-package dry-run, archive-size, and supply-chain gates
passed. The pinned five-document Word corpus produced nonempty evidence for
all 18 union pages. Its advisory SSIM trend remained 1 of 18 pages at or above
0.95.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep the classification reader-only and
namespace-aware. Rendering legacy VML horizontal rules requires a separate
story with its own output contract.

### F-198, Hyphenation

**Sprint.** S58
**Completed.** 2026-08-28
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Word settings and run properties now project, preserve,
and author automatic hyphenation plus all three modeled language attributes,
while retaining foreign attributes separately. The layout
path uses the effective run language to select the farthest fitting Liang
opportunity and emits a generated hyphen without claiming a source character.
The page-one showcase enables English hyphenation and renders
`representation` as consecutive `repre-` and `sentation` lines.

**Non-obvious choices.** Automatic hyphenation remains off unless the document
enables it, and unsupported languages remain unchanged. Restart pagination
retains a compact canonical XML identity plus current-view note references,
then performs exact byte equality after its fingerprint accelerator. This keeps
all four 700-paragraph related-story gates within the unchanged 8 MiB budget
without weakening retained-context identity.

**Deviations from the design plan.** Recovery discarded the obsolete local
shared-layout implementation and consumed the published `oxml-layout` 0.7.0
contract from F-X058. Microscope passes 3 and 4 found settings preservation,
target allocation, malformed-language ordering, source-compatibility prose,
and retained-state budget defects. Each received a red regression and a
separate remediation. Microscope pass 5 reported zero defects, zero smells,
and zero nitpicks.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Parser, round-trip, authoring, language inheritance, conditional
hyphen, source-span, no-wrap, field exclusion, cache-identity, and all four
700-paragraph related-story regressions passed. The deterministic golden gate
matched 7 of 7 page-one buffers. LibreOffice Writer 26.2.5.2 rendered exact
consecutive `repre-` and `sentation` lines. The pinned Word corpus produced
complete evidence for five documents and 18 union pages. The 1,000-page release
gate, current 0.7.0 registry carrier, immutable 0.6.0 historical carrier, full
workspace, no-default, WASM, documentation, 22-package dry-run, archive-size,
and supply-chain gates passed.

**Hash harness.** The reviewed delta changes exactly five
`feature_showcase` keys: page-one PNG, PDF bytes, PDF pages, PDF resources, and
Word document XML. The reason is the explicit page-one English hyphenation
example. All 49 baseline entries match after accepting that isolated delta.

**Notes for future sessions.** Preserve generated hyphens as `source: None`,
keep regional English tags mapped to the approved language data, and retain
the exact 700-paragraph restart boundary and 8 MiB budget. The recovery stash
and patch remain available until sprint close for auditability.

### F-199, Complex script shaping

**Sprint.** S58
**Completed.** 2026-08-28
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Word layout now projects Arabic, Devanagari, Thai, and
Simplified Chinese text through the shared multilingual shaping and line
breaking path published in `oxml-layout` 0.7.0. Script-local language values,
glyph clusters, offsets, logical source spans, conditional hyphens, field
formatting, note references, and cached-story provenance survive pagination,
drawing reflow, PDF, raster, and searchable SVG output. Latin-only content
continues to use the byte-identical legacy path.

**Non-obvious choices.** Exact-spaced rich Word lines use Word's 0.8em
baseline instead of each fallback font's `hhea` ascent. The Simplified Chinese
oracle uses a static Thin instance derived reproducibly from the approved Noto
Sans SC variable subset because LibreOffice selects its Regular instance while
the product consumes the file's Thin default. The fixture is oracle-only,
licenced and hash-pinned, and excluded from published crates. A bounded POSIX
lock serialises the complete macOS CoreText registration, Writer conversion,
and unregistration lifetime.

**Deviations from the design plan.** The approved oracle amendment added the
static Thin fixture after exact-font calibration proved variation-axis
selection was the only remaining CJK mismatch. Microscope passes 1 and 2 found
interaction defects involving hyphenation, mixed language slots, field
metadata, drawing reflow, cached source rebinding, hybrid bidirectional order,
and concurrent oracle font registration. Each received a focused red
regression and separate remediation. Microscope pass 3 reported zero defects,
zero smells, and zero nitpicks.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The shared Arabic joining, Indic cluster, Thai break, CJK
punctuation, mixed fallback, positioned-cluster, searchable-output, source,
hyphenation, field, drawing, cache, and paragraph-wide UAX 9 regressions
passed. The pinned multi-script hard gate passed all 4 pages at raw SSIM 0.95
or better: Arabic 0.985079311, Devanagari 0.972241230, Thai 0.997558968, and
Simplified Chinese 0.997294132. The five-document corpus retained complete
18-page evidence with its aggregate SSIM trend advisory. The full workspace,
no-default, WASM, documentation, 22-package dry-run, archive-size, and
supply-chain gates passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep pure Latin on the legacy identity path and
resolve bidirectional order across both legacy hyphenatable items and rich
complex-script items as one paragraph. The static CJK oracle font is evidence
for the approved variable source bytes, not a product font or a new public
variation-axis contract.

### F-200, Vertical and bidirectional text

**Sprint.** S58
**Completed.** 2026-08-28
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Word paragraph `w:bidi` and run `w:rtl` are typed without
losing unsupported attributes, duplicate occurrences, namespace identity, or
schema order. Paragraph and run direction now reach the shared UAX 9 layout
path across body text, tables, headers, footers, footnotes, endnotes, fields,
numbering, tabs, conditional hyphens, and drawing reflow. Visual painting order
and direction-sensitive alignment remain separate from logical PDF and SVG
extraction order and exact source attribution.

**Non-obvious choices.** Absent Word direction uses one inferred paragraph base
instead of forcing left-to-right, while explicit run overrides preserve their
internal digit and whitespace levels. Private reflow and cache sidecars carry
direction and logical provenance without changing the public
`ParagraphReflow` shape. Source-less fields, markers, generated hyphens, and tab
leaders use distinct private provenance so visually reordered lines remain
searchable in logical order. Existing whole-group quarter-turn vertical text
remains the documented approximation.

**Deviations from the design plan.** The approved bidi-only scope and five-file
HLD impact were retained. Eleven microscope passes found parser occurrence,
direction inference, hybrid line, field, object, note, table, header, cache,
reflow, and source-less provenance interactions. Each defect received an
existing-file red regression and a separate remediation turn. Microscope pass
11 reported zero defects, zero smells, and zero nitpicks.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/05-drawingml-model.md`, `docs/hld/08-rendering-spec.md`,
`docs/hld/10-bindings-spec.md`, and `docs/hld/12-testing-strategy.md`.

**Tests.** Typed direction round-trip, raw replay, paragraph and run override,
line-local L1 and L2, logical extraction, drawing reflow, table and story cache,
field, numbering, note, conditional-hyphen, tab-leader, and quarter-turn
regressions passed. The pinned raw oracle passed all five pages at SSIM 0.95 or
better: Arabic 0.956809869, Devanagari 0.972241230, Thai 0.997558968,
Simplified Chinese 0.997294132, and bidirectional 0.992810907. Full workspace,
no-default, WASM, documentation, 22-package dry-run, archive-size, and
supply-chain gates passed at the integrated SHA.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve logical source order independently of
visual positions, including for source-less generated content. Keep direction
transport private where the existing public aggregate is exhaustive, and do
not expand the quarter-turn approximation into upright vertical layout without
an explicit scope and HLD revision.

### F-X068, Tag rpptx-v0.8.0

**Sprint.** S58
**Completed.** 2026-08-29
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete fifteen-package shared OOXML and PowerPoint
family was published at 0.8.0 from reviewed SHA
`7f4414b0aeef1ec2cbae75fcb5aa96ab6dee6d70`. The release publishes the shared
text-direction carrier required by current stable Word source while preserving
the multilingual shaping, deterministic fonts, rich output, and logical text
contracts from 0.7.0. `rpptx-wasm` remains unpublished.

**Release evidence.** GitHub Actions run
[33258210706](https://github.com/tensorbee/rdocx/actions/runs/33258210706)
passed output stability, release metadata, reviewed notes, archive
verification, the exact fifteen-crate publication, and GitHub Release jobs.
Every selected 0.8.0 registry entry resolved under owner
`mantissaman (Atul Sharma)`. The annotated
[`rpptx-v0.8.0`](https://github.com/tensorbee/rdocx/releases/tag/rpptx-v0.8.0)
tag dereferenced to the reviewed SHA, and its 2,016-byte body was
byte-identical to the committed changelog render. `rpptx-wasm@0.8.0` remained
absent from crates.io, and the prepared stable layout graph resolved the
published `oxml-layout@0.8.0` boundary.

**Contribution inventory.** This shared-family carrier release added no
authenticated external issue or pull-request record. Issues 53 and 54 and PRs
55 through 58 remain assigned to the later stable 0.11.1 recovery inventory.

**Notifications.** None. The reviewed contribution inventory was empty, and
no issue or pull-request state changed during this release.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Full verification passed at the reviewed SHA with all workspace,
no-default, WASM, documentation, README, exact 22-package dry-run, archive,
and cargo-deny gates. The hash harness remained unchanged at 49 of 49. The
publication workflow and independent registry, owner, tag, release-target,
stable-graph, unpublished-WASM, and byte-for-byte release-body checks then
passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve the immutable `rpptx-v0.8.0` tag.
F-X069 may now prepare and publish the coherent stable 0.11.1 recovery against
the complete published shared 0.8.0 family.

### F-X069, Tag v0.11.1

**Sprint.** S58
**Completed.** 2026-08-29
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The complete seven-package stable Word family was
published at 0.11.1 from reviewed SHA
`5a850ce9ae6c31f8365594ed2970193266f8b2a6`. The release publishes
`rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`,
and `rdocx-cli` against the complete shared 0.8.0 registry family. Python,
WASM, npm, PyPI, shared, and PowerPoint publication authority remains
unchanged.

**Release evidence.** GitHub Actions run
[33266482507](https://github.com/tensorbee/rdocx/actions/runs/33266482507)
passed output stability, release metadata, reviewed notes, archive
verification, the exact seven-crate publication, and GitHub Release jobs.
Every selected 0.11.1 registry entry resolved unyanked under sole owner
`mantissaman (Atul Sharma)`. The annotated
[`v0.11.1`](https://github.com/tensorbee/rdocx/releases/tag/v0.11.1) tag
dereferenced to the reviewed SHA. Its 6,102-byte body was byte-identical to the
committed changelog render with SHA-256
`a5111e521f1adcb5ca856b54bfb2c69c6cccdd855a608e75737fce74a8f5de47`.

**Contribution inventory.** Issues
[53](https://github.com/tensorbee/rdocx/issues/53) and
[54](https://github.com/tensorbee/rdocx/issues/54) credit authenticated
`@emptinessform`. Pull requests
[55](https://github.com/tensorbee/rdocx/pull/55),
[56](https://github.com/tensorbee/rdocx/pull/56),
[57](https://github.com/tensorbee/rdocx/pull/57), and
[58](https://github.com/tensorbee/rdocx/pull/58) credit authenticated
`@pedroassumpcao`. Each outcome landed through a reviewed hardened equivalent.

**Notifications.** The six release-bound comments are
[Issue 53](https://github.com/tensorbee/rdocx/issues/53#issuecomment-5463995347),
[Issue 54](https://github.com/tensorbee/rdocx/issues/54#issuecomment-5463995659),
[PR 55](https://github.com/tensorbee/rdocx/pull/55#issuecomment-5463995914),
[PR 56](https://github.com/tensorbee/rdocx/pull/56#issuecomment-5463996200),
[PR 57](https://github.com/tensorbee/rdocx/pull/57#issuecomment-5463996504), and
[PR 58](https://github.com/tensorbee/rdocx/pull/58#issuecomment-5463996746).
All six records remain open.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Full verification passed at the reviewed SHA with all workspace,
no-default, WASM, documentation, README, exact 22-package dry-run, archive,
and cargo-deny gates. The hash harness remained unchanged at 49 of 49. The
publication workflow and independent registry, owner, tag, release-body,
shared dependency, and notification checks then passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve the immutable `v0.11.1` tag and all
seven live packages. F-X070 may separately prepare the cleanup of exactly
`rdocx-opc@0.11.0` and `rdocx-oxml@0.11.0`, but the yanks still require a new
final approval at its reviewed SHA.

### F-X070, Yank incomplete v0.11.0 packages

**Sprint.** S58
**Completed.** 2026-08-29
**Size.** S, estimated 1 day, actual 1 day

**What was built.** After the complete stable 0.11.1 recovery verified, the
separately approved cleanup yanked exactly `rdocx-opc@0.11.0` and
`rdocx-oxml@0.11.0`. Both entries remain available as immutable package bytes
but no longer participate in ordinary dependency selection. All seven 0.11.1
packages remain live and unyanked under sole owner
`mantissaman (Atul Sharma)`.

**Non-obvious choices.** The cleanup is limited to an incomplete family, not a
general policy for older versions. Complete coherent releases remain live. The
annotated v0.11.0 tag still peels to
`25350d000ed7ed96bf4f6e371f01f8fbc8e2cec4`, no v0.11.0 GitHub release exists,
and no issue, pull request, notification, tag, release, package byte, or other
registry version changed.

**Deviations from the design plan.** HLD 11 joined the impact list after the
preimplementation audit found its former blanket no-yank rule. The user
approved the narrow incomplete-family exception. Four microscope passes
hardened the exact command allowlist, corrected current release-family prose,
clarified local delivery records, and verified the post-yank evidence. Pass 4
reported zero defects and zero smells.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/11-migration-plan.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The cleanup contract regression permits exactly the two approved
version-specific yank commands and rejects wrappers, commands elsewhere in the
plan, tag or release mutation, comments, closures, publication, and every
other package or version. Independent readback reported both 0.11.0 targets as
yanked, the other five stable 0.11.0 package endpoints as absent, all seven
0.11.1 entries as live and unyanked under the sole authenticated owner, the
tag target unchanged, and the GitHub release absent. Full integrated
verification passed with the workspace, deterministic hash, no-default, WASM,
documentation, README, exact 22-package dry-run, archive-size, and supply-chain
gates.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve both yanked flags and the immutable
v0.11.0 tag. Do not generalize this incomplete-family cleanup into authority to
yank a complete coherent release.

### F-X031, Require the CI gate in branch protection

**Sprint.** S58
**Completed.** 2026-08-29
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The default branch now has active repository ruleset
`21823007`, named `Require CI gate on default branch`. It targets
`~DEFAULT_BRANCH` and requires the exact aggregate check `CI gate`. The sole
bypass is repository role 5, `admin`, in `always` mode so the reviewed
`/close-sprint` direct-push workflow remains executable.

**Non-obvious choices.** A repository ruleset preserves the pre-mutation state
without replacing a classic protection document. The bypass is limited to the
repository-administrator role. No team, app, user, or broader repository role
can bypass the gate. Both proof pull requests were closed without merging and
their disposable branches were deleted after their exact refs were checked.

**Deviations from the design plan.** None. The user separately approved the
persistent ruleset and its narrow administrator bypass before mutation.

**Spec sections touched.** `docs/hld/12-testing-strategy.md` and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The docs-only
[PR 59](https://github.com/tensorbee/rdocx/pull/59) reached CLEAN and MERGEABLE
after run
[33275852961](https://github.com/tensorbee/rdocx/actions/runs/33275852961)
reported a successful required `CI gate` while eight filtered product jobs
stayed skipped. The deliberately failing
[PR 60](https://github.com/tensorbee/rdocx/pull/60) became BLOCKED after run
[33276064981](https://github.com/tensorbee/rdocx/actions/runs/33276064981)
reported a failed `CI gate`, while `viewerCanMergeAsAdmin=true` proved the
approved close-sprint bypass. Independent readback verified the active ruleset,
sole bypass, exact check, closed-unmerged pull requests, and absent disposable
refs. Full integrated verification passed with workspace, no-default, WASM,
documentation, README, exact 22-package dry-run, archive-size, and supply-chain
gates.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve ruleset `21823007` and the exact
`CI gate` check identity. A future workflow rename must update the ruleset in a
separately reviewed operational story. Use `/close-sprint` for the reviewed
administrator-bypass push to `main`.

### F-217, Presentation collaboration and navigation model

**Sprint.** S59
**Completed.** 2026-08-30
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The presentation facade now owns relationship-safe modern
comment authors, comments, replies, ordered sections, slide membership, typed
notes header and footer settings, and typed handout header and footer settings.
Ordered mutations allocate collision-free package identities, validate the
complete graph before commit, and preserve unsupported XML and package content.

**Non-obvious choices.** Modern threaded comments are typed while legacy
comments remain opaque. Mutable collaboration operations use caller-supplied
identities and timestamps. Notes and handout roots retain byte-exact sidecars,
namespace bindings, lexical attributes, and schema-order boundaries.

**Deviations from the design plan.** None. Twelve microscope passes hardened
namespace shadow replay, raw-event placement, relationship ownership, collision
handling, and atomic failure paths before the final clean review.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** The named gate
`modern_comments_replies_sections_and_handout_settings_survive_ordered_mutation_save_and_reopen`
passed. The affected `oxml-opc`, `rpptx-oxml`, and `rpptx` suites passed,
including 116 PresentationML integration tests and the 50-deck corpus gates.
Full integrated verification passed with the workspace, no-default, WASM,
documentation, packaging, and supply-chain gates.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve typed collaboration graph ownership and
fail closed when fixed-prefix namespace shadows cannot be replayed safely. Do
not reinterpret legacy comments as the modern threaded model.

### F-221, Presentation encryption and signatures

**Sprint.** S59
**Completed.** 2026-08-30
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Native `rpptx::Presentation` callers can open and write
Agile-encrypted presentations, inspect and verify package signatures, and sign
the current presentation state. Both security features forward to the shared
`oxml-opc` implementation and remain disabled by default.

**Non-obvious choices.** Signature parts remain inspectable after mutation,
but verification stages current typed state and reports the retained signature
as invalid. Untouched signed content preserves producer bytes. Ordinary unsigned
serialization retains its established canonical modeled-part output, while
security operations use selective source-byte staging.

**Deviations from the design plan.** Integrated verification found that the
first selective-staging implementation also changed ordinary notes output. A
separate amendment split ordinary and security staging policies and added a
source-built notes regression before the full gate was rerun.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Encryption round trip, atomic failure, trusted-certificate signing,
complete coverage, stale-signature invalidation, untouched producer-signature
validity, ordinary notes staging, and feature-isolation gates passed. PowerPoint
16.104 build 16.104.25121423 opened candidate SHA-256
`a0d33171c63ec084231daeef3b35718f5a2d709a5c92c9c0e2017ccaf9fa52d6`
with the correct password and rejected a wrong password. Full integrated
verification passed, and all package archives remained below 10 MiB.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep `agile-encryption` and
`digital-signatures` absent from default, Python, WASM, and CLI graphs. Any new
mutable presentation handle must participate in current-state signature
invalidation without canonicalizing untouched signed producer bytes.

### F-213, Animation and transition timing model

**Sprint.** S60
**Completed.** 2026-08-30
**Size.** L, estimated 4 days, actual 1 day

**What was built.** PresentationML slides, layouts, and masters now expose
typed timing trees and slide transitions. The model covers parallel and
sequence containers, conditions and targets, supported animation behaviors,
set values, motion paths, build metadata, transition parameters, and morph
metadata while retaining unsupported XML in its original raw slots.

**Non-obvious choices.** Namespace-expanded parsing and capability-aware
AlternateContent selection keep producer aliases and fallbacks correct.
Supported mutations replace only the owning lexical range, leaving unsupported
siblings and descendants byte-identical. A narrow read exception accepts the
PowerPoint-produced empty layout transition marker immediately before `p:hf`,
then writes the two typed children in schema order. Private slide-master parse
staging boxes the large text-style value to keep nested producer groups within
the default test stack without changing the public model.

**Deviations from the design plan.** None. Nine microscope passes hardened
schema-choice cardinality, direct-child boundaries, compatibility selection,
lexical preservation, and valid XML character-data handling before the clean
review.

**Spec sections touched.** `docs/hld/06-presentationml-model.md`,
`docs/hld/10-bindings-spec.md`, and `docs/hld/12-testing-strategy.md`.

**Tests.** The named corpus gate parsed, serialized, and reparsed all 50 pinned
decks across 421 slides, 766 layouts, and 76 masters. It exercised 272 timing
roots, 372 transitions, 732 typed nodes, 401 conditions, 25 builds, 82 set
values, 186 effect parameters, 10 compatibility transitions, and 5 unsupported
nodes. Full integrated verification passed the workspace, no-default, WASM,
documentation, README, 22-package dry-run, archive-size, and supply-chain
gates. The `rpptx-oxml` suite passed 148 tests and the `rpptx` suite passed 97
tests with 7 ignored.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep timing evaluation outside `rpptx-oxml`.
Preserve expanded-name parsing, direct-child schema choices, selected
AlternateContent branches, and surgical raw-byte mutation when extending the
typed subset.

### F-214, Timeline evaluation and transition rendering

**Sprint.** S60
**Completed.** 2026-08-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The presentation layout layer now evaluates supported
timing trees at explicit slide-local timestamps and click counts. An additive
resolver path applies entrance, exit, emphasis, motion, group-target, and fill
state before freezing each frame. The renderer composes cut, fade, wipe, push,
zoom, and bounded explicit-name morph transitions through the existing page
frame model. The native facade exposes deterministic timestamped rendering
while every ordinary static entry point remains unchanged.

**Non-obvious choices.** Shape targets remain keyed by PresentationML
non-visual ids, and bounded morph pairs only explicit `!!` names with compatible
resolved geometry. Group animation and clipping use shared page-space geometry
so nested and descendant transforms cannot move a parent boundary. Fade and
compatible morph use source-over composition with an opaque outgoing layer.
Two narrowly approved OXML queries expose non-visual names and whether a timing
condition had an explicit target without changing the completed F-213 value
projection or reparsing XML in the layout crate.

**Deviations from the design plan.** The fixed PowerPoint movie could not
provide an observable zoom differential because the adjacent source slides
were visually identical throughout the zoom interval. The external matrix
therefore fails closed without a zoom row. Deterministic Rust regressions still
cover both zoom directions. The reviewed nine-case oracle retains evaluator,
fade, morph, and push evidence.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/06-presentationml-model.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** The named differential gate passed nine exact PowerPoint 16.104
samples. Maximum geometry error was 0.96 point and minimum luminance SSIM was
0.997866 against declared limits of 1 point and 0.99. The F-213 corpus gate
reparsed all 50 pinned decks across 421 slides, 766 layouts, and 76 masters.
The affected suites passed 151 `rpptx-oxml` tests, 119 `rpptx-layout` tests, 97
`rpptx-render` tests, and 136 `rpptx` tests with 8 expected ignored tests. Full
integrated verification passed the workspace, no-default, WASM, rustdoc,
README, publish dry-run, archive-size, and supply-chain gates. The 50-deck
static SSIM rider passed both corpus checks across 421 slides and recorded 25
slides at SSIM 0.95 or higher, a 0.512539 median, and `target_met=false` as
trend evidence rather than a quality claim.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve the independent static path, explicit
slide-local time and click convention, page-space group clipping, exact `!!`
morph opt-in, and fail-closed external-oracle bindings. Do not claim external
zoom parity until a source deck makes the effect observable.

### F-215, Audio and video package model

**Sprint.** S61
**Completed.** 2026-08-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Presentation pictures and timing trees now expose typed
embedded and linked audio and video attachments, poster ownership, playback
settings, triggers, and commands while retaining unmodelled XML. The native
facade can inspect, add, replace, extract, and remove media atomically. Package
mutation preserves exact external targets, opaque unsupported payloads,
content types, relationship ownership, shared parts, timing siblings, and
schema child order.

**Non-obvious choices.** The picture shape id is the stable public identity.
Trim remains owned by the Office 2010 picture media extension rather than the
timing node. Relationship replacement is structurally scoped to the direct
standard attachment and exact Office media extension, while removal prunes a
candidate part only after no package relationship reaches it. Format-neutral
MIME and container signature checks live in the dependency-free `oxml-media`
leaf.

**Deviations from the design plan.** Two concrete raw-XML and ownership helpers
were added to the existing picture and shape-tree types. The misleading timing
trim arguments and permanently empty timing trim fields were removed before
integration, leaving the Office media extension as the single truthful owner.
No new module, file, dependency, trait, generic, or feature flag was added.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** The named
`embedded_audio_and_video_corpus_media_round_trip_without_duplication` gate
passed against the pinned `EmbeddedAudio.pptx` and `EmbeddedVideo.pptx` decks
with exact payload, relationship, content type, poster, settings, and retained
metadata assertions. Source-built regressions cover embedded and linked
lifecycle operations, atomic failure, shared relationship and part ownership,
unsupported codecs, namespace aliases and shadows, schema order, raw timing
boundaries, duplicate-slide id remapping, and independent transition and
timing-effect parameter coverage across all 50 decks. Full integrated
verification, dependency checks, WASM, documentation, supply-chain checks, and
22 package dry runs passed. Every package archive remained below 10 MiB.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep media payloads opaque to the package model.
Preserve the exact structural ownership paths for picture and timing mutation,
and keep shared relationship records until all retained XML references are
gone. F-216 should consume the shape-id identity, picture-owned trim, timing
trigger context, poster relationship, and format-neutral diagnostics directly
rather than build a second media model.

### F-216, Media poster and playback rendering

**Sprint.** S61
**Completed.** 2026-08-31
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The presentation timeline now evaluates synchronized audio
and video playback state beside one deterministic page result. Valid local PNG
and JPEG posters continue through the existing resolved-image path. Missing,
linked, malformed, or unsupported posters can become deterministic labelled
Audio or Video groups, remain as poster policy output, or fail explicitly.
Audio and video payload bytes never enter renderer image admission and no codec
decoder was added.

**Non-obvious choices.** Playback commands fold in source order against the
existing timestamp and click convention. Stopped seek positions survive an
offset-free play, finite non-loop media stops at its exact end, and unknown
loop duration remains monotonic with a diagnostic rather than inventing a
wrap. Poster fallback diagnostics use private source-scoped slide, layout, or
master identity, while the existing static and timeline facade methods retain
their exact diagnostic text and behavior.

**Deviations from the design plan.** The approved fallback policy was added as
an explicit parameter to `render_media_timeline_deterministic`. This narrow
signature correction gives every policy variant a present consumer and lets
F-227 pass its selected export policy directly. No new file, module,
dependency, trait, generic, feature flag, renderer variant, or codec path was
added.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/07-inheritance-and-resolution.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** The named
`static_poster_output_and_timestamped_playback_state_match_source_built_oracle_fixtures`
gate passed at 150 dpi with deterministic fonts. The exact Audio fallback RGBA
hash is
`79e8da66b2cedf6a8b37b1eb723d4e278e22076aff5cdcff0616e0e7f2decec5`,
and the Video fallback hash is
`bca557835b53eb1ac671297c1d641c4f2e335fa51aa5515651ca95cc104e6917`.
Normalized playback rows cover automatic and click triggers, seek, trim,
volume, finite and unknown-duration loops, pause, resume, stop, linked media,
opaque codecs, and exact boundaries. Regressions prove source-scoped fallback
identity and byte-for-byte compatibility for the legacy static and timeline
entry points. Full integrated verification and patched publication dry runs
passed. The `rpptx-layout` and `rpptx` archives were 76,325 and 165,887 bytes,
both below 10 MiB.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** F-227 should consume
`DeterministicMediaTimelineFrame`, `EvaluatedMediaState`, and
`MediaFallbackPolicy` directly. Preserve the single shared assembly, the
source-scoped diagnostic identity, deterministic font mode, fixed 150 dpi
goldens, and the unchanged legacy facade paths.

### F-227, Animated GIF and video export

**Sprint.** S61
**Completed.** 2026-08-31
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The native presentation facade now exports bounded
deterministic animated GIF and Motion JPEG AVI bytes from explicit slide
segments. Each segment carries its own timeline position, click count, and
transition source, while one media fallback policy applies to the export. GIF
export preserves cumulative
centisecond timing and explicit loop behavior. AVI export writes an MJPEG
stream with exact RIFF headers, frame chunks, indexes, duration, dimensions,
and diagnostics.

**Non-obvious choices.** One prepared package, resolver, media, layout, and
font context is reused for the whole export. Each timestamp resolves and
renders one frame, feeds it directly to a capped encoder, and drops it before
the next sample. GIF and JPEG writes fail at the configured byte cap. AVI
patches sizes in one seekable capped buffer and retains only bounded index
metadata. No subprocess or external codec runtime is used.

**Deviations from the design plan.** None. The approved private
`crates/rpptx/src/animation.rs` module contains the concrete exporter. The
workspace adds `gif` 0.14.2 and enables the existing `jpeg-encoder` standard
library writer at the `rpptx` facade edge. No new crate, trait, generic,
feature flag, integration binary, or binary fixture was added.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/08-rendering-spec.md`,
`docs/hld/10-bindings-spec.md`, and `docs/hld/12-testing-strategy.md`.

**Tests.** The named
`animated_gif_and_motion_jpeg_avi_match_the_reviewed_two_machine_manifest`
gate passed unchanged on macOS and Linux arm64 with Rust 1.97.1. Its six
timestamps, decoded GIF frame hashes, loop count, dimensions, ordered
diagnostics, and GIF container hash `1682901777930996407` match the reviewed
manifest. The AVI gate independently parses every RIFF and LIST boundary,
header, stream rate, frame chunk, index entry, JPEG payload, decoded frame,
duration, dimension, and ordered diagnostic. Its exact container hash is
`6525351511319371367`. Negative mutations cover every encoded and decoded AVI
identity plus diagnostic contents and order. Regressions prove real
two-slide fade, click-triggered shape and media state, bounded 50-frame
streaming, immediate output-cap rejection, and quality sensitivity. Full
integrated verification passed all changed crates, the workspace, no-default,
WASM, rustdoc, README, dependency, corpus, supply-chain, and 22-package dry-run
gates. Every archive remained below 10 MiB, with `rpptx` at 180,361 bytes.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Preserve explicit segment ownership, exact
integer timestamp sampling, cumulative GIF delays, one-frame resolution, and
the native bounded encoder path. Treat the macOS and Linux manifest as one
identity contract. Platform-specific container, payload, decoded-pixel, or
diagnostic constants are not acceptable.

### F-219, SmartArt typed model

**Sprint.** S62
**Completed.** 2026-09-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Presentation packages now expose bounded typed SmartArt
data, layout, quick-style, colour, and cached-drawing parts. Callers can inspect
diagram identities and supported content, edit node text atomically, duplicate
complete diagram graphs, and transfer placeholder-free SmartArt slides between
presentations through an explicitly selected destination layout.

**Non-obvious choices.** Raw part bytes remain authoritative until a supported
field is edited. Checked edits and transfers validate exact relationship roles,
duplicate graph identities, namespace safety, schema positions, placeholder
absence, and complete owned graphs before committing. Cross-presentation
transfer accepts only the five diagram relationship types plus
relationship-free internal images, rejects unsupported dependencies, and uses
one 128-part ceiling in preflight and copy.

**Deviations from the design plan.** None. The approved additive pre-1.0 API
includes `Presentation::transfer_smartart_slide_from` and the bounded
`CT_Slide::contains_placeholder` helper. The approved diagram model is the one
new source module. No dependency, feature flag, trait, integration binary, or
binary fixture was added.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** The named round-trip gate proved supported nodes remain editable
after save and reopen while unrelated mutations leave unsupported diagram parts
byte-exact. Focused coverage also exercises all five part roots, alias prefixes,
schema order, producing relationship scopes, duplicate identities, atomic
failures, slide duplication, cross-presentation transfer, nested graph cycles,
the 128-part ceiling, and the pinned 50-deck SmartArt corpus. Full integrated
verification and all parser, dependency, API, package, and archive-size riders
passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** F-220 should consume the typed layout family,
point graph, DrawingML node text, style and colour projections directly. Keep
unsupported algorithms and producer extensions raw, preserve schema-owned
projection boundaries, and keep complete graph validation atomic.

### F-218, Embedded object and macro inventory

**Sprint.** S62
**Completed.** 2026-09-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The presentation facade now inventories relationship-owned
OLE objects, ActiveX controls, and VBA projects without executing them. Each
entry reports stable package coordinates, content type, byte length, SHA-256,
and signature state. Callers can extract bytes exactly and replace or remove
content through atomic package transactions.

**Non-obvious choices.** Ownership follows namespace-aware XML references and
complete relationship graphs across slides, layouts, masters, control parts,
and the presentation. Mutations preserve shared targets and unrelated orphans.
Signature evidence is either retained as invalidated or removed through an
explicit policy. Private relationship markers make invalidation observable
after save and reopen without adding a default cryptographic dependency.

**Deviations from the design plan.** None. The approved private
`crates/rpptx/src/embedded.rs` module contains the graph and transaction logic.
The native pre-1.0 facade additions are intentional. Existing workspace
dependencies provide SHA-256 and XML parsing. No feature flag, trait, generic,
integration binary, or binary fixture was added.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/10-bindings-spec.md`, and
`docs/hld/12-testing-strategy.md`.

**Tests.** Source-built regressions cover exact inventory and extraction,
shared and orphan retention, transactional replacement and removal, safe
internal targets, schema-position ownership, duplicate identities, complete
signature topology, and both signature policies. The tracked corpus proves two
layout-owned and one master-owned OLE frame. Fresh-reopen tests cover ordinary
package and VBA signature invalidation with and without the optional digital
signature verifier. Full workspace verification, corpus, LibreOffice,
no-default, WASM, rustdoc, dependency, supply-chain, publish dry-run, and
archive-size gates passed. The final microscope pass reported zero defects,
zero smells, and zero nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep executable payloads opaque and retain the
source-part plus relationship-id identity. Validate the complete ownership and
signature graph before mutation. Never infer ownership from filenames or file
extensions, and never delete an unreachable producer orphan that the requested
operation does not own.

### F-X071, Integrate PRs 61 through 64

**Sprint.** S62
**Completed.** 2026-09-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The reviewed reader outcomes from PRs 61 through 64 are
integrated as one hardened Word reader boundary. Hyperlink and drawing facts,
table and row properties, numbering metadata and effective formatting, nested
revisions, and complex-field display segments are available without weakening
raw XML preservation. The contribution is credited to `@pedroassumpcao`.

**Non-obvious choices.** PRs 61 and 64 supplied directly usable outcomes. PR
62 required namespace propagation through every table owner and expanded-name
revision-slot handling. PR 63 required default-style numbering association and
a narrower producer-content contract that still reports retained raw paragraph
and run properties. Later live heads were audited and intentionally not adopted
where they duplicated the hardened result or reintroduced those defects.

**Deviations from the design plan.** None. The pinned contribution inputs were
PR 61 at `7c40c2e`, PR 62 at `fa48a39`, PR 63 at `60bc663`, and PR 64 at
`5cb5cba`. The maintained result includes subsequent hardening without a new
crate, module, feature flag, trait, generic, or dependency.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** Regressions cover expanded-name drawing relationships, inherited
table namespace bindings, malformed revision schema slots, direct and
default-style numbering, nested revision depth and source order, complex field
segments, producer raw property reporting, and repeated save and reopen.
Complete `rdocx-oxml`, `rdocx`, `rdocx-html`, and `rdocx-layout` suites passed,
including the LibreOffice boundary. Full workspace verification, no-default,
WASM, rustdoc, dependency, supply-chain, publish dry-run, and archive-size
gates passed. Microscope pass 10 reported zero defects, zero smells, and zero
nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep namespace identity and schema position
authoritative when exposing preserved reader facts. Preserve exact producer
content unless the typed projection owns it. Any future live PR changes must be
audited against the hardened canonical behavior rather than assumed newer.

### F-X072, Keep paragraph caching across note references

**Sprint.** S63
**Completed.** 2026-09-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** Direct body paragraphs containing a footnote or endnote
reference now remain eligible for paragraph-cache reuse. A later edit in a
700-paragraph document reuses 699 exact cached paragraphs and rebuilds one,
instead of the early reference disabling every later cache read.

**Non-obvious choices.** The complete typed paragraph remains the key, so a
changed note ID misses its paragraph. Exact footnote and endnote parts remain
in the retained-work context, so changed note content invalidates reuse for the
transaction. Note-bearing table, header, and footer paragraphs remain
conservative because their retained payloads do not own body note placement.

**Deviations from the design plan.** None. The implementation changes one
private safety predicate and adds one private helper with table and header or
footer consumers. It adds no public API, dependency, feature, module, file,
trait, or generic.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** The footnote and endnote gates each require 699 hits and one rebuild
after a late edit. Additional regressions cover changed reference IDs, changed
note parts, exact warm and fresh equality, and conservative fields, numbering,
drawings, raw children, tables, headers, and footers. Full workspace
verification, the 219-test `rdocx-layout` suite, no-default fonts, WASM,
rustdoc, README inventories, and the release 1,000-page performance gate all
passed. Microscope pass 2 reported zero defects, zero smells, and zero
nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Direct body note references are safe only
because both the complete paragraph and exact note-part context are compared.
Do not reduce either identity, and do not widen table or related-story caching
without making body note placement part of that payload's ownership.

### F-X073, Restart ordinary-prose pagination within the aggregate cache

**Sprint.** S63
**Completed.** 2026-09-01
**Size.** L, estimated 4 days, actual 1 day

**What was built.** Ordinary multi-line prose, headings, and paragraphs using
`keepNext` or `keepLines` can now publish restart checkpoints at complete block
boundaries. Restart candidates use available capacity in the existing 64 MiB
aggregate cache budget instead of failing at an independent 8 MiB ceiling.
Late edit, insert, delete, and undo operations can therefore reuse the exact
complete prefix while remaining byte-for-byte equal to fresh layout.

**Non-obvious choices.** Record representability is separate from checkpoint
placement. A paragraph split across pages falls back through the normal
paginator and discards that restart candidate. Source-level fields, numbering,
drawings, multilingual state, raw children, anchored empty paragraphs, and
unsupported line items remain fail-closed. Bookmark safety compares complete
namespace-aware raw elements, exact expanded-name attribute cardinality, and
the retained `raw_before` position rather than trusting rendered projections.

**Deviations from the design plan.** None. The implementation changes private
layout and paginator logic only. It adds no public API, dependency, feature,
module, file, trait, or generic.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Source-built regressions cover multi-line prose, headings,
`keepNext`, `keepLines`, aggregate-budget admission and rejection, exact warm
and fresh edit, insert, delete, and undo equality, bounded recomputation, split
paragraph fallback, unsafe fields, and exact bookmark raw XML. Full integrated
workspace verification passed, including the 225-test `rdocx-layout` suite,
LibreOffice and corpus gates, no-default fonts, both WASM facades, rustdoc,
README inventories, publish dry-runs, archive limits, and supply-chain checks.
The release-mode 1,000-page performance rider passed. Microscope pass 4
reported zero defects, zero smells, and zero nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Publish restart records only at complete block
boundaries, and account for every current and pending cache class under the one
aggregate limit with checked arithmetic. Keep source-level safety conservative.
Rendered output alone is not proof that a restart record owns every effect.

### F-220, SmartArt layout and rendering

**Sprint.** S63
**Completed.** 2026-09-01
**Size.** L, estimated 5 days, actual 1 day

**What was built.** Native presentation rendering now supports the six pinned
authentic list, hierarchy, cycle, relationship, matrix, and pyramid SmartArt
programs. Namespace-aware OXML projections retain bounded instruction, style,
and colour semantics. A private evaluator validates exact ownership, order,
cardinality, attributes, parameters, and total work before transiently lowering
the diagram into the existing DrawingML geometry, paint, and text paths.

**Non-obvious choices.** Five families use readable exact semantic profiles
after the field and ownership validators. The three-node PowerPoint 16.104
cycle1 resource has a private identity, raw SHA-256, instruction, and node-count
compatibility boundary because its producer curve solve is not specified by
OOXML. Unsupported or changed programs fail closed and remain preserved. The
same resolved group serves static, timeline, media, and animation output.

**Deviations from the design plan.** None. The implementation remains private
to the presentation facade and doc-hidden OXML projections. It adds no facade,
binding, layout, or renderer public API.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/06-presentationml-model.md`,
`docs/hld/07-inheritance-and-resolution.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** The common-source PowerPoint 16.104 corpus proves exact ownership,
diagnostics, and provenance, shape edges within 1 point, ordered owner-centred
text metrics within 3 points, and symmetric text-masked non-text SSIM of at
least 0.90 for every family. Sensitivity rejects a 1.01-point displacement and
a calibrated decorative mutation. Adversarial tests cover unsupported trees,
excess work, changed identities and hashes, node counts, same-kind reorders,
tuple substitutions, duplicate semantic replacements, and wrong owners. The
integrated full presentation test binary, workspace tests, Clippy, no-default
fonts, WASM, rustdoc, README inventories, package dry-run, archive limit, and
supply-chain checks passed. Microscope pass 12 reported zero defects, zero
smells, and zero nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep the OXML projection as the only diagram
model. Any newly supported authentic program needs an explicit bounded semantic
profile and external common-source evidence. Do not infer plausible geometry
from an unknown program or weaken the cycle1 identity boundary.

### F-222, ODP read and write

**Sprint.** S63
**Completed.** 2026-09-01
**Size.** L, estimated 5 days, actual 1 day

**What was built.** The native presentation facade now reads bounded ODP
packages and writes deterministic ODP packages. The editable subset includes
slides, ordinary rectangles and text boxes, tables, embedded images, slide
names, and notes. Unsupported safe content is retained as stable, capped,
location-aware diagnostics rather than silently disappearing.

**Non-obvious choices.** The reader requires the stored ODP mimetype first,
validates safe unique archive paths, parses namespace-expanded XML at exact
schema positions, and applies part, total, depth, node, text, and diagnostic
limits. The writer charges escaped text and media before growth, writes a
complete manifest, and publishes files atomically. It does not claim parity for
charts, transitions, animation, media timing, SmartArt, or unsupported visual
appearance.

**Deviations from the design plan.** None. The approved private
`crates/rpptx/src/odp.rs` module contains the package and projection logic. The
native pre-1.0 facade additions are intentional. The existing workspace ZIP
dependency is used directly. No feature flag, trait, generic, integration
binary, or binary fixture was added.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/08-rendering-spec.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Source-built regressions cover supported round trips, deterministic
archive construction, namespace aliases and lookalikes, duplicate expanded
attributes, unsafe entries, every declared resource limit, diagnostic order
and ceiling, malformed images, and atomic read and save failures. The pinned
LibreOffice 26.2.5.2 gate passed in both ODP-to-PPTX and PPTX-to-ODP directions.
The full workspace, presentation, Clippy, no-default-font, WASM, rustdoc,
README, publish dry-run, archive-size, prose, workflow, and skill-sync gates
passed. Microscope pass 2 reported zero defects, zero smells, and zero
nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep ODP import bounded before allocation and
classify nodes by expanded name and schema position. Add a construct to the
editable subset only when both directions have a pinned differential. Every
safe unsupported direct object must produce one stable diagnostic.

### F-223, Modern presentation package variants

**Sprint.** S63
**Completed.** 2026-09-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The native presentation facade now identifies and writes
ordinary presentations, macro-enabled presentations, ordinary templates,
macro-enabled templates, ordinary slide shows, and macro-enabled slide shows.
Ordinary saves retain the opened class. Output-only conversion changes only
the presentation main-part content type on a staged package copy.

**Non-obvious choices.** Package identity comes from the exact main-part
content type rather than the filename extension. Class conversion preserves
VBA, OLE, ActiveX, relationships, signatures, and unrelated content types. A
class change records package-signature invalidation because it changes the
signed content-type table. Unknown main-part content types fail closed.

**Deviations from the design plan.** None. The implementation stays in the
existing presentation facade, adds no module or dependency, and keeps the
existing `save_as_show` API as a delegating compatibility method. The native
pre-1.0 enum and methods are intentional additive public API.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Source-built regressions cover all five newly supported classes,
ordinary save retention, all-class conversion, exact executable payload and
relationship preservation, package-signature invalidation, unknown-class
rejection, and the established slide-show method. Full workspace tests,
Clippy, no-default fonts, both WASM facades, rustdoc, README inventories,
dependency direction, package dry-runs, and archive limits passed. Microscope
pass 1 reported zero defects, zero smells, and zero nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Treat the main-part content type as the only
authoritative package class. Keep conversions output-only and preserve every
opaque executable part and relationship. Any additional class requires an
exact registered content type and a source-built round-trip.

### F-226, Notes and handout export

**Sprint.** S63
**Completed.** 2026-09-01
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The native presentation facade now exports one
deterministic notes page per source slide and all six audience handout layouts
to PDF and PNG. Notes pages include vector slide thumbnails, speaker text,
stored metadata, and master content. Handouts preserve the master below
aspect-fitted, clipped, bordered, and numbered thumbnails, with writing rules
for the three-slide layout.

**Non-obvious choices.** Notes-master and notes-slide relationships are remapped
into one collision-free transient owner scope before the shared resolver runs.
This preserves distinct media when both source scopes use the same relationship
id. Placeholder overlay matches explicit index first and compatible type
second, consumes each owner once, and fails on missing or ambiguous ownership.

**Deviations from the design plan.** None. The implementation remains in the
existing facade, reuses the existing deterministic layout and rendering stack,
and adds no module, dependency, feature, binding, or renderer API.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/04-opc-and-packaging.md`, `docs/hld/06-presentationml-model.md`,
`docs/hld/07-inheritance-and-resolution.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, `docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The six named source-built integration gates cover notes geometry,
all handout layouts, deterministic PDF and PNG dimensions, noncanonical
relationship targets, malformed graphs, invalid DPI, and package-byte
preservation. Additional regressions cover exact placeholder ambiguity,
cross-scope relationship id collisions, master z-order, and rejection of a
1.01-point geometry displacement. Full workspace verification and public
package riders passed. Microscope pass 4 reported zero defects, zero smells,
and zero nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep notes and handout composition in the
facade-owned package assembly stage. Do not resolve notes-slide media in the
notes-master relationship scope, and do not replace vector thumbnails with a
raster intermediate.

### F-224, HTML slide content import

**Sprint.** S64
**Completed.** 2026-09-02
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The native presentation facade now imports bounded HTML5
and CSS into a fresh presentation. Top-level slide sections or one document
body project explicit absolute geometry into editable shapes, formatted text,
tables, caller-supplied images, and slide-owned links. Stable ordered
diagnostics retain safe unsupported content, and the candidate validates,
saves, and reopens before publication.

**Non-obvious choices.** The private importer follows the existing
`default-template` boundary and keeps `scraper` as an optional direct `rpptx`
dependency. Images resolve only from the caller-provided resource slice, with
no network or filesystem fetch. Aggregate selector work has its own bound.
Unsupported empty semantic elements remain diagnostic, and diagnostics are
published in document order across collection phases.

**Deviations from the design plan.** None. The completed plan includes the
review-driven selector-work bound, empty semantic-element handling, and
cross-phase diagnostic ordering.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/08-rendering-spec.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The ten source-built unit, regression, integration, round-trip, and
differential tests cover exact CSS geometry, the bounded cascade, every
declared resource limit, unsupported-content diagnostics, document-order
diagnostics, editable content after reopen, opaque XML and schema order, and
geometry, text, and pixel sensitivity. The gate
`source_built_html_matches_pinned_chrome_after_save_and_reopen` passed against
Google Chrome 152.0.7977.65 at 0.984628993 full-image luminance SSIM. Full
non-fast verification and all routed riders passed. Microscope pass 4 reported
zero defects, zero smells, and zero nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep browser semantics limited to the declared
absolute-position and CSS subset. New layout behavior needs an explicit
contract, a hard work bound, and a pinned differential. Keep image resolution
caller-owned and publication transactional.

### F-225, PDF page content import

**Sprint.** S64
**Completed.** 2026-09-02
**Size.** L, estimated 4 days, actual 1 day

**What was built.** The native presentation facade now imports bounded PDF
pages in preserved and editable modes. Preserved mode emits one deterministic
full-slide PNG per page. Editable mode projects supported text, paths, raster
images, and URI links into ordinary slide content. Strict parsing, stable
diagnostics, checked geometry, and candidate save and reopen keep publication
transactional.

**Non-obvious choices.** The private importer follows the existing `render`
boundary and uses no-default `lopdf` only for PDF syntax and object decoding.
Both modes share one normalized layout path. Resource traversal is iterative,
cycle checked, depth bounded, and charged through one cached reference
resolver. Unsupported state is isolated until reset. Active content and
malformed URI actions fail closed. Dash phases are accepted only at zero or an
exact representable boundary.

**Deviations from the design plan.** None. The completed plan records the exact
dash boundary, affine behavior, resource accounting, active-content policy, and
strict URI decoding established during review.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/06-presentationml-model.md`, `docs/hld/08-rendering-spec.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Source-built unit and regression coverage proves exact point-to-EMU
geometry, affine text and image placement, dash isolation, font decoding and
widths, image typing, iterative resource traversal, shared reference-work
accounting, strict page and annotation graphs, active-content rejection,
fallible URI decoding, every declared limit, diagnostics, and transactional
publication. Both modes and editable content pass save and reopen coverage.
The gate
`pdf_page_import_matches_pinned_poppler_geometry_pixels_text_and_links` passed
against Poppler 26.01.0 at raw full-image luminance SSIM 0.999557935 for
preserved mode and 0.998561398 for editable mode. Its geometry, text, link, and
pixel perturbation regression also passed. Full canonical non-fast
verification, all routed riders, and all 22 package dry-runs passed. Microscope
pass 11 reported zero defects, zero smells, and zero nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep PDF engines out of the production graph.
Charge every XObject member and every previously unseen indirect-reference hop
against the shared work budget. Preserve the common normalized layout path
between modes, and add editable operators only with explicit fail-closed state
semantics and differential evidence.

### F-X075, Preserve restart pagination across page-spanning paragraphs

**Sprint.** S64
**Completed.** 2026-09-02
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The Word layout engine now keeps the completed recorded
pagination pass when otherwise safe ordinary prose spans page boundaries. A
split continuation still emits no checkpoint. Restart state is published only
at later complete block boundaries after note, wrap, and resolved state is
clean. The obsolete document-wide split veto and its private flag were
removed.

**Non-obvious choices.** The existing source-safety predicates, exact retained
context, aggregate cache budget, checkpoint limit, field substitution rules,
and transactional publication remain authoritative. The performance harness
authenticates each historical HEAD and hashes the complete measured production
source plus the exact injected harness. It rejects production, harness,
same-line suffix, duplicate-pin, and untracked-source mutations before timing.

**Deviations from the design plan.** None. Review expanded the planned exact
result proof to cover metadata, logical structure, and result-local provenance,
and strengthened the planned pinned performance evidence with source-content
authentication.

**Spec sections touched.** `docs/hld/08-rendering-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** The deterministic Issue 67 fixture contains 175 four-line
paragraphs across exactly 16 pages. Ten sourced middle edits retain 174
paragraph-cache hits, rebuild one paragraph, paginate at most two pages, and
match a fresh result across pages, fonts, diagnostics, outlines, metadata,
structure, and provenance. Late edit, insert, delete, undo, note-bearing split,
displayed page-number footer, unsafe-state, aggregate-bound, and 1,000-page
gates remain green. The authenticated 48-run release comparison measured
current-to-v0.11.1 ratios from 0.341 to 0.389 and
current-to-`0582da0` ratios from 0.175 to 0.203 across the 175 and 700 paragraph
native and bundled-fallback paths. Full integrated non-fast verification and
all routed riders passed. Microscope pass 4 reported zero defects, zero smells,
and zero nitpicks.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep restart checkpoints at complete block
boundaries. A split page is not itself unsafe, but any state absent from the
checkpoint identity remains a reason to fall back. Performance evidence must
authenticate measured source content rather than trusting a checkout label.

### F-X074, Tag rpptx-v0.9.0

**Sprint.** S64
**Completed.** 2026-09-03
**Size.** S, estimated 1 day, actual 2 days

**What was built.** The reviewed `rpptx-v0.9.0` release published the exact
15-package shared OOXML and PowerPoint family at 0.9.0. The annotated tag
dereferences to reviewed sprint SHA
`45b4f277ff5fd6d1b032e929c5dcee7fb9d2c550`. The GitHub `Publish` workflow
completed successfully at
https://github.com/tensorbee/rdocx/actions/runs/33719233249, and the release is
available at https://github.com/tensorbee/rdocx/releases/tag/rpptx-v0.9.0.

**Non-obvious choices.** Publication remained limited to `oxml-core`,
`oxml-opc`, `oxml-media`, `oxml-layout`, `oxml-drawing`, `oxml-pdf`,
`oxml-sml`, `oxml-cli-support`, `oxml-chart`, `rpptx-oxml`, `rpptx-chart`,
`rpptx-layout`, `rpptx-render`, `rpptx`, and `rpptx-cli`. Every registry entry
reports version 0.9.0 and owner `mantissaman (Atul Sharma)`. The stable Word
family stayed at 0.11.1, and `rpptx-wasm@0.9.0` remained absent from crates.io.

**Deviations from the design plan.** None. The release used the separately
approved exact reviewed SHA, pushed the sprint branch first, created one
annotated tag, pushed only that tag, and retained the selected-family boundary.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Full non-fast verification and the clean sprint review passed at
the released SHA with 49 of 49 deterministic hashes unchanged. All 22 local
package dry-runs passed below 10 MiB. The publication workflow passed output,
metadata, release-note, archive, selected-family, and GitHub release gates.
Independent post-publication checks downloaded all 15 exact registry versions,
verified their owner, verified the annotated tag target, and matched the
3,119-byte GitHub release body byte for byte to the reviewed changelog render.
Both bodies have SHA-256
`31f83a5d629c1cbd837176d5993d8d98469dbd0d9fd9fac733055ae9083c4c2e`.

**Contribution inventory.** The selected-family inventory is empty. Pull
requests https://github.com/tensorbee/rdocx/pull/59 and
https://github.com/tensorbee/rdocx/pull/60 were disposable internal CI proofs.
Pull requests https://github.com/tensorbee/rdocx/pull/61 through
https://github.com/tensorbee/rdocx/pull/64 and Issues
https://github.com/tensorbee/rdocx/issues/65 through
https://github.com/tensorbee/rdocx/issues/67 belong to the stable Word family.
No release notification comment was required or posted for this PowerPoint
release.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep release-family selection exact. Verify
registry ownership, the dereferenced annotated tag, release-body bytes, and
explicitly excluded binding packages before completing a release story.

### F-X076, Tag v0.12.0

**Sprint.** S64
**Completed.** 2026-09-03
**Size.** S, estimated 1 day, actual 1 day

**What was built.** The reviewed `v0.12.0` release published the exact stable
Word family: `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`,
`rdocx-pdf`, `rdocx`, and `rdocx-cli`. The annotated tag dereferences to the
reviewed sprint SHA `19adaacfcf82e3918bba4f8c3648747f1969b746`. The GitHub
`Publish` workflow completed successfully at
https://github.com/tensorbee/rdocx/actions/runs/33728011369, and the release is
available at https://github.com/tensorbee/rdocx/releases/tag/v0.12.0.

**Non-obvious choices.** Publication remained limited to the seven stable
crates at 0.12.0. Every registry entry reports owner `mantissaman (Atul
Sharma)`. Shared OOXML and PowerPoint packages stayed at their published 0.9.0
boundary. `rdocx-wasm@0.12.0` remained absent from crates.io, and no Python,
npm, WASM, or PyPI package was published.

**Deviations from the design plan.** None. The release used the separately
approved exact reviewed SHA, pushed the sprint branch first, created one
annotated tag, pushed only that tag, and retained the selected-family boundary.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** Full non-fast verification and the clean sprint review passed at
the released SHA with 49 of 49 deterministic hashes unchanged. All 22 local
package dry-runs passed below 10 MiB. The publication workflow passed output,
metadata, shared-family, release-note, archive, stable allowlist, and GitHub
release gates. Independent post-publication checks downloaded all seven exact
registry versions, verified their owner, verified the annotated tag target,
and matched the 3,525-byte GitHub release body byte for byte to the reviewed
changelog render. Both bodies have SHA-256
`37e3fbc29e0d7425a4ca5559e8fd73db2fc25db31a000183343521ac83f68c07`.

**Contribution inventory.** Pull requests
https://github.com/tensorbee/rdocx/pull/61,
https://github.com/tensorbee/rdocx/pull/62,
https://github.com/tensorbee/rdocx/pull/63, and
https://github.com/tensorbee/rdocx/pull/64 by authenticated
`@pedroassumpcao` contributed reader outcomes through reviewed hardened
equivalents. Issues https://github.com/tensorbee/rdocx/issues/65,
https://github.com/tensorbee/rdocx/issues/66, and
https://github.com/tensorbee/rdocx/issues/67 by authenticated
`@emptinessform` reported cache and pagination defects fixed through reviewed
hardened equivalents. The pull requests remain closed and unmerged. Issues 65
and 66 remain closed, and Issue 67 remains open.

**Notifications.** Release-bound thank-you comments were posted at
https://github.com/tensorbee/rdocx/pull/61#issuecomment-5522283440,
https://github.com/tensorbee/rdocx/pull/62#issuecomment-5522283669,
https://github.com/tensorbee/rdocx/pull/63#issuecomment-5522283889,
https://github.com/tensorbee/rdocx/pull/64#issuecomment-5522284147,
https://github.com/tensorbee/rdocx/issues/65#issuecomment-5522284354,
https://github.com/tensorbee/rdocx/issues/66#issuecomment-5522284700, and
https://github.com/tensorbee/rdocx/issues/67#issuecomment-5522284889.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep stable publication limited to the exact
seven-package family. Verify each downloaded registry package, owner, annotated
tag target, release-body bytes, unpublished carrier, contribution state, and
notification URL before completing a release story.

### F-228, OfficeMath model and authoring

**Sprint.** S65
**Completed.** 2026-09-03
**Size.** L, estimated 4 days, actual 1 day

**What was built.** `rdocx-oxml` now owns one bounded, recursive Transitional
OfficeMath model for inline and display equations, runs, fractions, scripts,
radicals, matrices, limits, n-ary operators, delimiters, accents, and
document-wide math properties. The native `rdocx` facade exposes equations in
paragraph source order with borrowed inspection, indexed mutation, and bounded
authoring.

**Non-obvious choices.** Supported children are projected into typed schema
slots while unsupported siblings, attributes, and property content retain
their original XML. Namespace aliases resolve by expanded name, authored math
uses the fixed `m:` prefix, unsafe prefix collisions fail closed, and legacy
Equation Editor objects remain opaque.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** The round-trip gate
`officemath_corpus_parses_mutates_saves_and_reopens_without_losing_supported_or_raw_siblings`
covers every supported expression, raw sibling preservation, mutation, and
reopen. Public facade tests cover inline and display authoring, ordered access,
and mutation. Six microscope passes ended clean, and full integrated
verification passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Layout and conversion must consume this public
tree directly. They must not introduce a second equation model or weaken the
fail-closed namespace and raw-preservation behavior.

### F-229, OfficeMath layout and PDF rendering

**Sprint.** S65
**Completed.** 2026-09-03
**Size.** M, estimated 2 days, actual 1 day

**What was built.** `rdocx-layout` now measures and lowers the shared typed
OfficeMath tree into backend-neutral groups for inline and display equations.
Fractions, scripts, radicals, matrices, limits, n-ary operators, delimiters,
and accents participate in normal line breaking and pagination with concrete
ascent, descent, and baseline geometry. The document facade carries optional
document-wide math properties into layout.

**Non-obvious choices.** Shared inline and line groups gained an optional
baseline instead of a math-specific backend element. Existing drawing groups
retain top alignment through `None`. Equation glyphs use the existing font
manager and deterministic bundled Caladea fallback, so PDF and raster backends
consume only `LayoutResult` and do not learn Word grammar.

**Deviations from the design plan.** None.

**Spec sections touched.** `docs/hld/03-architecture.md`,
`docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`,
`docs/hld/12-testing-strategy.md`, and
`docs/hld/14-development-backlog.md`.

**Tests.** The golden gate
`officemath_baselines_and_glyph_geometry_match_the_pinned_word_pdf_oracle`
matched the source-built deterministic rendering against Microsoft Word
16.104 and Poppler 26.01.0 within the declared 1.0 point tolerance. Focused
layout, pagination, preservation, and mutation tests passed. Five worker
microscope passes ended clean, and full integrated verification passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Keep equation rendering behind shared layout
primitives. Preserve the deterministic font boundary and require intentional,
reviewed evidence for any Word-oracle geometry change.

### F-230, MathML and LaTeX conversion

**Sprint.** S65
**Completed.** 2026-09-03
**Size.** M, estimated 2 days, actual 1 day

**What was built.** The native `rdocx` facade now imports and exports the
supported equation subset as MathML and LaTeX through four free functions.
Both readers normalize directly into the F-228 `MathArgument` tree, both
writers emit canonical text, and every declared loss boundary returns an
ordered path-aware diagnostic. Input, tree, matrix, text, recursion, event,
token, and diagnostic work is bounded.

**Non-obvious choices.** Conversion owns no second equation model. MathML uses
expanded names through the existing XML dependency, while LaTeX uses a local
recursive-descent parser with explicit grouping and script attachment. Pandoc
3.10 is an exact-version structural test oracle only and is absent from the
published runtime dependency graph. Document-wide and display properties stay
with their owners when a bare equation argument crosses the conversion API.

**Deviations from the design plan.** None. The integration reconciliation kept
the F-229 layout contract and the F-230 conversion contract as adjacent owners
of the same normalized tree.

**Spec sections touched.** `docs/hld/02-scope-and-non-goals.md`,
`docs/hld/03-architecture.md`, `docs/hld/04-opc-and-packaging.md`,
`docs/hld/10-bindings-spec.md`, `docs/hld/12-testing-strategy.md`,
`docs/hld/14-development-backlog.md`, and
`docs/hld/15-build-and-toolchain.md`.

**Tests.** The differential gate
`mathml_and_latex_conversion_matches_pinned_pandoc_texmath_trees` matched
source-built cases against live Pandoc 3.10 in both directions. Unit,
round-trip, bounds, loss-diagnostic, and perturbation tests passed. Six worker
microscope passes and the integration reconciliation pass ended clean, and
full integrated verification passed.

**Hash harness.** Unchanged, 49 of 49.

**Notes for future sessions.** Compare normalized expression trees rather than
serialized bytes when extending the conversion subset. Keep Pandoc outside the
runtime graph and add a mutation-sensitive case for each new grammar rule.
