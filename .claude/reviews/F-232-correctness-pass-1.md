# F-232, correctness, pass 1

**Reviewed**: Uncommitted working diff from worker HEAD `9109a9a`, 10 files and 1,486 changed lines, with 1,457 insertions and 29 deletions. The four added `toc_rebuild` regression tests pass.
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, page placeholder substitution can rewrite unowned source XML
`crates/rdocx/src/field.rs:346`

Each generated page token is replaced across the complete serialized main
document, not only in the generated TOC result that owns it. The first token is
the predictable `__RDOCX_TOC_PAGE_0__` value created at
`crates/rdocx/src/field.rs:1515`. If a heading, unrelated field cache,
attribute, or retained raw subtree already contains that text, a successful
rebuild changes it to the resolved page number. Source content and unowned XML
must remain unchanged, so the sentinel must be collision-safe and replacement
must be confined to its owned result location.

### D2, the raw ownership scan accepts foreign ancestors and non-direct field markers
`crates/rdocx/src/field.rs:647`

The scanner treats any Word-namespaced `fldChar` below a recognized paragraph
as a live marker. It does not require the marker to be a direct child of a
Word-namespaced run. In addition, `DynamicXmlElement` discards ancestor
namespaces and `dynamic_paragraph_parent` accepts `body`, `tbl`, `tr`, `tc`,
`sdt`, and `sdtContent` by local name alone at
`crates/rdocx/src/field.rs:820`. A Word `fldChar` retained inside an unmodelled
run child, or a Word paragraph below a foreign same-local-name container, can
therefore become a TOC span or shift the paragraph indexes used against the
typed model. The result is a false diagnostic, rejection, or mutation of
opaque producer XML instead of the required direct-run, expanded-name,
fail-closed behavior.

### D3, nested TOC spans are edited with stale overlapping byte ranges
`crates/rdocx/src/field.rs:324`

Nested complex TOC fields produce the inner span before the outer span because
spans are recorded when their end marker closes at
`crates/rdocx/src/field.rs:944`. Reversing that list applies the outer edit
first, which removes or resizes the byte range still referenced by the inner
edit. The second unchecked `Vec::splice` can panic when the stale range is now
out of bounds, or it can splice into newly generated outer content when the
range still fits. Nested TOC content must either be assigned to one owner or be
rejected before any edit ranges are applied.

### D4, facade and layout enumerate PAGEREF targets in different document orders
`crates/rdocx/src/field.rs:1584`

The facade maps layout target indexes back to bookmark names with
`collect_body_paragraphs` and all `paragraph.runs()`, while layout assigns those
indexes through `visit_document_paragraphs` and revision-projected runs at
`crates/rdocx-layout/src/engine.rs:5668`. The orders can diverge. For example,
the facade interleaves table content controls at their stored row boundary at
`crates/rdocx/src/field.rs:4247`, but layout visits every table-level content
control before every row at `crates/rdocx-layout/src/engine.rs:5713`. A
PAGEREF in deleted revision content also appears only in the facade list.
Subsequent target ids then map to the wrong bookmark, so a rebuilt TOC can
materialize another entry's page number or report a target missing.

The same unordered helper is newly used for bookmark text at
`crates/rdocx-layout/src/engine.rs:5760`, so a valid bookmark that crosses a
positioned table or row content-control boundary can also be observed in
reversed order and incorrectly treated as invalid.

### D5, the declared Word differential gate does not compare the required output
`crates/rdocx/tests/regression_test.rs:1279`

The test pins strings naming the Word build and input, but its result checks use
membership and aggregate counts at
`crates/rdocx/tests/regression_test.rs:1314`. It does not encode or compare the
Word result's exact entry order, entry-to-level mapping, hyperlink targets, or
displayed page values. Reordering all four entries, assigning links to the
wrong bookmarks, or emitting arbitrary cached page values can still satisfy
these assertions. The second matrix also exercises sequence separators without
asserting the resulting page display. This is not mutation-sensitive to the
behaviors that the approved differential gate requires.

### D6, bookmark allocation rejects a legal final id and fails when no allocation is needed
`crates/rdocx/src/field.rs:1134`

Allocator construction immediately computes `max_id + 1`, even when every
source reuses a bookmark or no source needs one. A document with an existing
bookmark id of `i32::MAX` therefore rejects an otherwise allocation-free
rebuild. There is also an off-by-one at
`crates/rdocx/src/field.rs:1148`: when `next_id` is `i32::MAX`, `allocate`
increments before returning, so the last representable collision-free id can
never be used. These valid boundary cases should succeed without partial
mutation.

## Smells

None.

## Nitpicks

None.

## Not found

- Structure: no new trait, generic parameter, forwarding wrapper, module,
  feature flag, or crate was introduced.
- Public surface: the additive native report export stays within the approved
  Rust facade, with Python, WASM, and CLI unchanged.
- Dependencies: no new runtime or oracle dependency was added.
- Generated OOXML child order: the generated paragraph properties, hyperlinks,
  runs, and simple PAGEREF fields are emitted in a schema-compatible order.
- Atomic commit sequencing: apart from the panic path in D3, fallible staging,
  parsing, reopening, layout, and final commit keep the live document outside
  the edit path until success.
