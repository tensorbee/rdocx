# F-237, all, pass 13

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the twelve earlier review artifacts, 17 files and 7,618 changed lines, comprising 7,532 additions and 86 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, form-kind containers accept known children from the wrong grammar
`crates/rdocx-oxml/src/text.rs:1849`
`crates/rdocx-oxml/src/text.rs:1980`

Kind-child order validation runs only when the direct child's local name maps
to a slot for the selected form kind. The following match likewise ignores
every nonmatching pair. Consequently a `w:textInput` can contain `w:checked`,
a `w:checkBox` can contain `w:listEntry`, or a `w:ddList` can contain
`w:format` and still be projected as a supported form. A typed value edit
retains the cross-kind WordprocessingML leaf, and staged reopen accepts the
same malformed owner. The known form grammar therefore remains incomplete
despite the corrected per-kind facets.

### D2, non-whitespace character data is accepted in form element-only containers
`crates/rdocx-oxml/src/text.rs:1672`
`crates/rdocx-oxml/src/text.rs:1679`

Character data is rejected only while `leaf_depth` is set. Non-whitespace
text, CDATA, or a general reference directly inside `w:ffData`,
`w:textInput`, `w:checkBox`, or `w:ddList` falls through and leaves the owner
selectable. Mutation replays that data and the same parser accepts it on
reopen, so invalid element-only form content can be committed through the
typed API.

### D3, form rewriting can select a nested retained ffData before the typed owner
`crates/rdocx-oxml/src/text.rs:4145`
`crates/rdocx-oxml/src/text.rs:4147`

The reader projects only the direct `w:ffData` child of the begin field
character, but the rewriter assigns `ff_depth` for every descendant named
`w:ffData`. A retained producer or compatibility subtree placed before the
real owner can therefore contain a nested `w:ffData` and matching form-kind
element that receives the requested value edit first. Closing that nested
owner also clears `ff_depth`, so the actual direct owner is not changed.
Staged reopen observes the old typed value and rejects the operation. This
makes an otherwise supported field uneditable solely because of an unrelated
retained subtree and breaks the reader-to-writer ownership identity.

### D4, glossary element-only containers accept non-whitespace character data
`crates/rdocx-oxml/src/glossary.rs:416`
`crates/rdocx-oxml/src/glossary.rs:435`

The glossary root scanner rejects text, CDATA, and references only when the
element stack is empty. A document with one valid entry plus non-whitespace
character data inside `w:glossaryDocument` or `w:docParts` therefore opens as
a supported glossary. Selected-entry replacement patches only the entry span,
retains the invalid container content, and the same permissive reopen accepts
it. The malformed-root graph gate does not cover the element-only character
grammar it publishes.

## Smells

None.

## Nitpicks

None.

## Not found

All six pass-12 findings are closed. Form names and formats now use the 20 and
64 character maxima. Help and status values use their independent 255 and 140
character bounds. Drop-down defaults are range-checked independently. Glossary
types use the seven-value enumeration, GUIDs require the braced uppercase
lexical form at parser and facade boundaries, and checkboxes require exactly
one size choice. The exact pass-1 through pass-11 cases remain concretely
closed.

No additional findings were found in glossary relationship ownership and
content types, normalized package identities, story root and note scope,
source-order form identity, nested field and content-control traversal,
cached-display updates, selected-entry structural replacement, namespace and
raw-subtree preservation beyond the defects above, staged failure atomicity,
panic safety, public API structure, dependency direction, HLD file scope, or
repository structure. All 427 `rdocx-oxml` unit tests and its doc test pass,
including all 34 focused glossary tests and 11 focused legacy-form tests. All
36 focused F-237 integration tests pass. `cargo check -p rdocx --all-targets`,
`cargo fmt --all --check`, and `git diff --check 4ba8b6b` pass.
