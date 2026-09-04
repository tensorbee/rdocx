# F-232, correctness, pass 2

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 11 files and 1,833 changed lines, with 1,790 insertions and 43 deletions. All 238 `rdocx-layout` unit tests, its doc test, and all eight `toc_` regression tests pass.
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, opaque descendants can still be treated as live TOC fields
`crates/rdocx/src/field.rs:893`

The new direct-owner check proves only that a marker's immediate parent is a
Word `r`. It does not prove that the run itself belongs to a typed paragraph
path. A preserved `<x:wrapper><w:r><w:fldChar .../></w:r></x:wrapper>` inside a
recognized paragraph therefore inherits that paragraph index and is accepted
as a live marker. Repeating that shape for begin, instruction, separator, and
end makes the raw scanner discover a TOC that the typed paragraph parser cannot
project, so `rebuild_toc` returns an error instead of leaving the opaque field
alone. The simple-field diagnostic scan has the same ownership gap at
`crates/rdocx/src/field.rs:820`, where any descendant Word `fldSimple` inherits
the paragraph and increments the diagnostic count. Pass-1 D2 is therefore only
partially closed.

### D2, reuse of multiple valid whole-paragraph bookmarks is nondeterministic
`crates/rdocx/src/field.rs:1189`

Bookmark starts are consumed from a randomized `HashMap`, and every valid
whole-paragraph range overwrites the prior value for the same paragraph at
`crates/rdocx/src/field.rs:1201`. A paragraph may legally carry two distinct,
properly nested bookmarks that both cover its full run range. Which name is
then reused for the generated hyperlink and PAGEREF depends on hash iteration
order, so identical input can produce different output across processes. The
deterministic rebuild contract needs a stable document-order selection.

### D3, TOC sequence matching disagrees with the evaluator's identifier semantics
`crates/rdocx/src/field.rs:1311`

Source discovery compares the TOC `\s` identifier to the preceding SEQ
identifier with case-sensitive string equality. The shared SEQ evaluator puts
identifiers into one case-insensitive sequence namespace by lowercasing them at
`crates/rdocx/src/field.rs:3077`. With `TOC \s chapter` and a preceding `SEQ
Chapter`, the SEQ field resolves normally but the rebuild fails to retain its
value as the page prefix. Matching the same field identity differently at the
rebuild boundary yields a page display without the required sequence prefix.

### D4, the page-number gate still cannot detect target misassociation
`crates/rdocx/tests/regression_test.rs:1364`

The new exact tuples correctly bind order, level, links, and cached displays,
but every entry in both differential matrices has displayed page value `1`.
The separate pagination regression contains only one TOC target at
`crates/rdocx/tests/regression_test.rs:1468`. A mutation that maps every
bookmark to the first resolved target page would therefore pass both exact
matrices, the single-target page-2 regression, and the layout name-order unit
test while producing wrong page numbers for a multi-entry document. The
differential gate needs at least two entries with distinct final pages in one
rebuild to prove target-to-entry association. Pass-1 D5 is stronger but not yet
fully mutation-sensitive.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-1 D1: collision-safe tokens and unique owned-range substitution are now
  enforced at `crates/rdocx/src/field.rs:356` and
  `crates/rdocx/src/field.rs:1623`.
- Pass-1 D3: TOC spans are sorted and overlapping ranges are rejected before
  edit construction at `crates/rdocx/src/field.rs:788`.
- Pass-1 D4: layout now owns the result-local target mapping, accepted-revision
  projection, and positioned control order at
  `crates/rdocx-layout/src/lib.rs:58` and
  `crates/rdocx-layout/src/engine.rs:5718`.
- Pass-1 D6: bookmark IDs are allocated lazily and the final representable ID
  is returned before exhaustion at `crates/rdocx/src/field.rs:1220`.
- Bounds and panics: no additional reachable panic, stale slice, splice, or
  arithmetic-overflow path was found.
- OOXML schema order and verbatim preservation: no additional generated-child
  order, whitespace, prefix-write, or unowned-byte loss was found beyond D1.
- Structure and surface: no unjustified trait, generic, wrapper, module,
  feature flag, crate, dependency, Python, WASM, or CLI surface was introduced.
- Atomicity: staged parse, reopen, layout, and substitution failures remain
  outside the live document commit path.
