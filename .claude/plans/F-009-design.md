# F-009, Cache the layout result

**Status**: completed
**Sprint**: S02
**Size**: M
**Depends on**: none

## Problem

`Document::render_page_to_png` builds a fresh layout input and calls
`layout_document` on every page at `crates/rdocx/src/document.rs:2111`.
Rendering all pages one at a time is therefore O(n squared). Mutable facade
wrappers also hold direct references into document content, so cache
invalidation must cover both direct `Document` mutations and mutable-accessor
entry points.

## Spec reference

- `docs/hld/08-rendering-spec.md`, "Performance".
- `docs/hld/10-bindings-spec.md`, "Threading".
- `docs/hld/13-risks-and-open-questions.md`, "render_page_to_png is O(n squared)".

## Approach

Add a `Mutex`-protected cache with separate
`Option<Arc<rdocx_layout::LayoutResult>>` entries for normal and deterministic
font modes. This deliberately replaces the backlog's `RefCell<Option<Rc<_>>>`,
which would break the `Document: Send + Sync` contract required by the Python
binding design. Private layout helpers compute each mode once, while a private
`invalidate_layout()` clears both entries. `render_page_to_png`,
`render_page_to_png_deterministic`, `render_all_pages`, and `to_pdf` reuse the
matching result. `to_pdf_with_fonts` keeps its one-off layout because its
caller-owned font inputs are not part of the cache key.

Invalidate before every public mutation and before returning any mutable
paragraph, table, row, cell, section, or run access that can mutate through a
borrow. Conservative invalidation on mutable access is preferred to threading a
forwarding cache handle through all facade wrappers.

Add `pub fn layout_page(&self, page_index: usize) -> Result<Option<PageFrame>>`,
returning a clone of the selected cached page. This is additive public API and
keeps the internal `Rc` private.

## Rejected alternatives

- Put cache handles into every mutable facade wrapper. That increases the
  number of places a reader must inspect and creates forwarding state whose
  concrete owner is already known.
- Cache all font modes in one slot. Their font resolution inputs differ, so
  sharing would return the wrong render.
- Use the backlog's `RefCell<Option<Rc<_>>>` literally. It would make
  `Document` neither `Send` nor `Sync`, contradicting the bindings contract.
- Cache by a mutation revision counter. A single invalidated `Option` is enough
  for the one current implementation.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression | `rendering_all_pages_performs_one_layout` | Rendering every page of a 20-page document through page entry points calls layout exactly once |
| unit | `document_mutation_invalidates_cached_layout` | A direct mutation after layout causes exactly one new layout |
| unit | `mutable_accessor_invalidates_cached_layout` | Obtaining and changing a paragraph or table wrapper cannot reuse stale layout |
| unit | `font_modes_use_isolated_layout_caches` | Normal and deterministic entry points reuse only their matching cache, while custom fonts remain uncached |
| compile-time regression | `document_remains_send_and_sync` | The cache does not break the threading contract needed by Python rendering |

The backlog test gate is `rendering_all_pages_performs_one_layout`, observed
through a test-only layout invocation counter in the existing implementation
file rather than a new trait or injectable abstraction.

## HLD impact

- `docs/hld/08-rendering-spec.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/13-risks-and-open-questions.md`
- `docs/hld/14-development-backlog.md`

## Risk routing

- Layout and pagination. Use deterministic font mode for baseline checks and
  re-record no baseline incidentally.
- Public API of a published crate. `layout_page` is additive and story-required.
  Run rustdoc, `cargo publish --workspace --dry-run`, and the 10 MiB archive
  assertion.

## Hash harness

Expected to remain unchanged. Caching reuses the same layout result and must not
alter any selected XML or PNG bytes.

## Implementation checklist

- [x] Add the two-mode thread-safe cache and private compute/invalidate helpers.
- [x] Route normal and deterministic page rendering through the matching cache.
- [x] Add the public cloned-page entry point.
- [x] Invalidate every direct mutation and mutable-accessor path.
- [x] Add one-layout, invalidation, and font-mode isolation tests without a new
      abstraction or test binary.
- [x] Run focused rdocx, layout, PDF, rustdoc, packaging, and harness checks.

## Open questions

None. Use the approved thread-safe two-mode cache. Caller-supplied font
rendering remains uncached because its inputs require a separate key.
