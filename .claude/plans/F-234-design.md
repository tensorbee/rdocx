# F-234, Full-story document comparison

**Status**: approved
**Sprint**: S67
**Size**: L
**Depends on**: F-167

## Problem

The existing comparison is body-only. `Document::compare` constructs one
tracked `CT_Document`, validates normalized body postconditions, and commits
only `self.document` (`crates/rdocx/src/comparison.rs:89`). It rejects modeled
field changes (`crates/rdocx/src/comparison.rs:1583`), reports most formatting
changes as diagnostics (`crates/rdocx/src/comparison.rs:593`), and emits only
deletion and insertion wrappers for replacements (`crates/rdocx/src/comparison.rs:393`).
It therefore lacks the required related stories, moves, and tracked formatting.

Revision resolution is also main-document-only
(`crates/rdocx/src/revision.rs:101`). The facade already owns typed footnotes
and comments (`crates/rdocx/src/document.rs:1383`), deterministic
relationship-resolved story traversal (`crates/rdocx/src/field.rs:197`), and raw
text-box paragraph discovery (`crates/rdocx/src/template.rs:1259`). F-234 must
reuse those owners and preserve their source mappings rather than create a
second document model.

## Spec reference

- `docs/hld/03-architecture.md`, revision ownership, deterministic comparison,
  and Word source provenance.
- `docs/hld/04-opc-and-packaging.md`, package integrity and staged comparison.
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy" and "The hash harness".
- `docs/hld/14-development-backlog.md`, "F-234, Full-story document comparison".

## Approach

Keep the existing public surface source-compatible:

```rust
pub fn compare(
    &mut self,
    edited: &Document,
    author: &str,
    timestamp: &str,
) -> Result<Vec<ComparisonDiagnostic>>;
```

Inside `comparison.rs`, build one private prefix-aware source index for main,
deduplicated headers and footers in section-reference order, comments in stored
part order, normal footnotes and endnotes in part order, and text boxes nested
at their host part and source ordinal. Fields remain inline at their paragraph
and run owner. Use concrete story identities and child paths, not a new trait or
public result model. Reject dangling, external, wrong-type, duplicate, or
structurally changed story shells before mutation.

Reuse the hierarchical paragraph, table, row, cell, control, and run comparator
for every story. Compare a field as one owned source. A changed cached result
under an unchanged instruction revises only that result. An instruction or form
change replaces the complete owned field source. Pair unmatched identical
owners deterministically after LCS and emit paired `w:moveFrom` and `w:moveTo`
only within one story. Cross-story moves and whole-story shell creation or
removal are outside the declared tracked-change boundary and fail atomically.

For aligned equal content, write supported current properties plus prior
properties as fixed-prefix `w:rPrChange`, `w:pPrChange`, `w:tblPrChange`, or
`w:sectPrChange`. Unsupported row, cell, and unmodelled formatting stays a
diagnostic or a rejected shell difference. Seed one collision-free revision id
allocator from every compared part in both documents.

Patch only owned spans in candidate part bytes and preserve every unowned byte,
prefix binding, raw child, processing instruction, and relationship target.
Reopen the complete staged package, then commit with the existing atomic facade
boundary. Extend all revision-resolution selectors across the same source index
and commit every selected part once. Keep `Document::revisions()` main-story-only
because no story requests a new all-story inspection surface.

Independently accept and reject staged copies. Compare supported story structure
and modeled formatting to the edited and original inputs, verify unchanged story
identities, and preserve exact existing `WordSourcePath` results for body,
headers, footers, footnotes, and endnotes in tracked, accepted, and rejected
views. Keep the public diagnostic shape unchanged. Python, WASM, and CLI remain
unchanged. Add no module, file, dependency, feature flag, trait, or generic.

## Rejected alternatives

- Extend `rdocx_layout::WordStory` with comments and text boxes. That type owns
  rendered provenance and the new variants would be an unrelated source break.
- Add a second comparison result surface. The existing method can cover F-234,
  while F-235 owns policy configuration.
- Serialize whole non-body parts from partial typed models. Source-span patching
  is required for verbatim preservation.
- Treat moves as delete plus insert. The differential gate names moves
  explicitly.
- Compare Word XML bytes. The oracle contract compares normalized records and
  trees, while prefixes and whitespace remain our serializer decisions.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| differential, gate | `full_story_comparison_matches_pinned_word_records` | Source-built pairs match Word 16.104 insertion, deletion, move, formatting, field-owner, and story-placement records. |
| regression | `full_story_comparison_differential_rejects_kind_order_and_story_perturbations` | Every kind, order, move-pair, field-owner, and formatting perturbation fails the predicate. |
| regression | `accepting_and_rejecting_full_story_comparison_reproduces_each_supported_story` | Accept matches edited and reject matches original across every supported story. |
| regression | `scoped_revision_resolution_visits_every_compared_story_once` | All selectors visit each story once and preserve unselected revisions. |
| regression | `comparison_preserves_word_source_paths_in_every_revision_view` | Tracked, accepted, and rejected layout retains exact source paths and scalar ranges. |
| round-trip | `comparison_preserves_unmodelled_story_xml_and_relationships_byte_for_byte` | Raw story XML, shells, fields, namespace bindings, and relationships survive reopen. |
| integration | `a_failed_full_story_comparison_leaves_package_typed_state_and_caches_unchanged` | Metadata, shell, relationship, source-span, and postcondition failures are atomic. |
| unit | `repeated_story_content_and_moves_use_deterministic_matches_and_ids` | Repeated content receives stable matches, move pairs, locations, and ids. |

The **test gate**, from the backlog, is differential. Pinned document pairs
produce the same insertions, deletions, moves, and story placement as Word at
the declared boundary.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- **Any parser or serialiser**. Add prefix-tolerant discovery and reparse,
  fixed-prefix schema-positioned revisions, focused `rdocx-oxml` tests, and a
  round trip proving byte-identical unmodelled story subtrees.
- **Public API of a published crate**. The existing method and resolution
  selectors gain source-compatible behavior. Run rustdoc with warnings denied,
  the patched workspace publish dry-run, and every archive size assertion.
- **An external oracle comparison**. Pin Microsoft Word 16.104 build
  16.104.25121423 plus locale in the harness, compare normalized records rather
  than bytes, record every intentional divergence, use source-built inputs, and
  prove mutation sensitivity.

## Hash harness

Expected unchanged across all 49 entries. No sample invokes comparison. Any
delta is unexplained and blocks integration.

## Implementation checklist

- [ ] Approve the exact full-story, move, and formatting boundary.
- [ ] Build one relationship-resolved story index in `comparison.rs`.
- [ ] Seed one global revision allocator from every source part.
- [ ] Compare related stories, fields, and nested text boxes through the
  existing hierarchical owner logic.
- [ ] Emit stable same-story moves and supported property revisions.
- [ ] Patch only owned spans and reopen the complete staged package.
- [ ] Extend revision resolution across the same stories.
- [ ] Prove accepted and rejected package-wide postconditions and provenance.
- [ ] Add the source-built pinned Word differential and focused regressions.
- [ ] Run focused checks, every risk rider, and the unchanged hash harness.
- [ ] Update exactly the four listed HLD files.

## Open questions

None. Identical story shells, same-story moves, supported modeled property
changes, the diagnostic-only unsupported formatting boundary, and the pinned
Microsoft Word 16.104 differential evidence are approved.
