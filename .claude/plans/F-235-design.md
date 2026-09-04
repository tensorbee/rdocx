# F-235, Comparison granularity and ignore policy

**Status**: completed
**Sprint**: S67
**Size**: M
**Depends on**: F-234

## Problem

The current native comparison API accepts no policy
(`crates/rdocx/src/comparison.rs:89`). It aligns complete run signatures and
replaces a changed run as one deletion plus one insertion
(`crates/rdocx/src/comparison.rs:382`). Formatting differences always produce
diagnostics (`crates/rdocx/src/comparison.rs:1210`), changed modeled fields are
rejected (`crates/rdocx/src/comparison.rs:1583`), and current postcondition
normalization is limited to the main body.

F-234 first replaces that body-only boundary with one ordered, source-mapped
story traversal. F-235 must refine that engine and cannot build a second
traversal or source-map model.

## Spec reference

- `docs/hld/03-architecture.md`, revision resolution and deterministic
  comparison ownership.
- `docs/hld/04-opc-and-packaging.md`, staged atomic comparison and verbatim
  preservation.
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy", document comparison, and
  "The hash harness".
- `docs/hld/14-development-backlog.md`, "F-234, Full-story document comparison"
  and "F-235, Comparison granularity and ignore policy".

## Approach

Promote F-234's private ordered story category to an additive native public
type, then add concrete options:

```rust
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ComparisonStoryKind {
    Main,
    Header,
    Footer,
    Comment,
    TextBox,
    Footnote,
    Endnote,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ComparisonGranularity {
    #[default]
    Run,
    Word,
    Character,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ComparisonOptions {
    pub granularity: ComparisonGranularity,
    pub ignore_formatting: bool,
    pub ignore_whitespace: bool,
    pub ignore_fields: bool,
    pub ignore_comments: bool,
    pub ignored_stories: Vec<ComparisonStoryKind>,
}

pub fn compare_with_options(
    &mut self,
    edited: &Document,
    author: &str,
    timestamp: &str,
    options: &ComparisonOptions,
) -> Result<Vec<ComparisonDiagnostic>>;
```

Keep `Document::compare` source-compatible and behavior-compatible by
delegating to default options. Pass one policy through F-234's traversal,
signatures, diagnostics, and postcondition projection. Apply it in this order:

1. Skip selected story categories before shell comparison or revision
   allocation, preserving original story bytes.
2. Ignore comment story content plus comment anchors and references while
   preserving the original graph.
3. Treat an ignored modeled field as one opaque original leaf.
4. Remove ignored textual whitespace from comparison projection only.
5. Suppress formatting participation and diagnostics when requested.
6. Apply granularity to remaining visible text, keeping non-text content atomic.

Character mode uses Unicode scalar values. Word mode emits maximal Unicode
alphanumeric or underscore sequences, maximal Unicode whitespace sequences,
and maximal punctuation or symbol sequences. `ignore_whitespace` applies only
to Unicode whitespace in `w:t` and `w:delText`. Tabs, breaks, drawings, fields,
and raw run content remain significant.

Build private attributed text units that retain the story, source path, run
properties, run-content position, and raw-child boundary. Align with the
existing deterministic LCS tie-break. Coalesce adjacent results with identical
ownership before writing the smallest valid fixed-prefix revision fragments.
Every non-text or preserved raw child is emitted exactly once.

Ignored differences are left-biased. They create no revision or diagnostic and
retain the original bytes. Acceptance reproduces the edited policy projection,
while rejection reproduces the original supported structure. F-234's staged
package, source mapping, accept and reject proofs, and atomic commit remain the
single boundary. Python, WASM, and CLI stay unchanged. Add no crate, dependency,
trait, generic, feature flag, module, file, or test binary.

## Rejected alternatives

- Change `Document::compare` or its defaults. That would break callers or
  silently change revision grouping.
- Add a second granular comparator. It would duplicate story ordering and
  source mapping.
- Reuse `rdocx_layout::WordStory`. It omits comments and text boxes and belongs
  to renderer provenance.
- Copy ignored content from the edited document. Ignore means preserve the
  original, not mutate silently.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| unit | `comparison_defaults_preserve_run_granularity` | Legacy and default option calls return identical diagnostics and bytes. |
| unit | `word_and_character_granularity_split_only_text_content` | Unicode boundaries are exact and structural run content stays atomic. |
| regression, gate | `comparison_policy_matrix_changes_only_declared_records_and_is_deterministic` | Every policy changes only declared records and repeated runs are identical. |
| regression | `comparison_granularity_preserves_accept_and_reject_postconditions` | All granularities accept to the edited projection and reject to the original. |
| regression | `ignored_formatting_suppresses_only_formatting_diagnostics` | Formatting records disappear while content revisions remain. |
| regression | `ignored_whitespace_preserves_source_whitespace_without_hiding_structural_controls` | Text whitespace stays original while tabs and breaks remain significant. |
| regression | `ignored_fields_preserve_field_sources_and_compare_neighbouring_text` | Field XML stays original and adjacent text is still compared. |
| regression | `ignored_comments_preserve_comment_parts_and_anchors` | Comment content and anchors stay original while nearby stories are compared. |
| regression | `ignored_story_kinds_skip_only_selected_story_categories` | Each category can be skipped independently. |
| integration | `invalid_comparison_policy_leaves_package_and_caches_unchanged` | Validation and postcondition failure is atomic. |
| round-trip | `granular_comparison_preserves_unmodelled_xml_byte_for_byte` | Splitting and ignores never duplicate, move, normalize, or drop raw XML. |

The **test gate**, from the backlog, is regression. Each policy changes only
the declared comparison records and remains deterministic under repeated runs.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- **Any parser or serialiser**. Verify schema-ordered fixed-prefix fragments,
  prefix-tolerant reparse, exact captured-subtree preservation, and no raw-child
  duplication after run splitting.
- **Public API of a published crate**. The enum, options value, re-exports, and
  method are additive native Rust API. Run rustdoc with warnings denied, the
  patched workspace publish dry-run, and every archive size assertion.

## Hash harness

Expected unchanged across all 49 entries. Comparison is opt-in and no sample
invokes it. Any delta is unexplained and blocks integration.

## Implementation checklist

- [x] Complete the F-234 dependency-prefix checkpoint before starting F-235.
- [x] Promote the ordered story category and add the granular policy types.
- [x] Add `compare_with_options` and preserve legacy defaults.
- [x] Thread one policy through story traversal and postconditions.
- [x] Implement deterministic attributed text units.
- [x] Serialize minimal run fragments without duplicating raw content.
- [x] Apply ignore policy in the declared precedence.
- [x] Preserve atomic staging, source maps, ids, and accept and reject proofs.
- [x] Add the named source-built tests to existing files.
- [x] Run focused checks and every risk rider.
- [x] Update exactly the four listed HLD files.

## Open questions

None. Unicode scalar character mode, the three-class word tokenizer, and
text-element-only whitespace ignore are approved.
