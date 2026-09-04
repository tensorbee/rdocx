# Current Sprint, S67

**Milestone**: M22 Word depth.

**Goal**: deepen the two flagship document-automation workflows. Mail merge
must expand nested and named data into rich document content, while comparison
must cover the complete supported Word story set with explicit granularity and
ignore policies.

## Spec references

- `docs/hld/03-architecture.md`, for staged mail merge, deterministic
  hierarchical comparison, story traversal, revision ownership, and atomic
  facade commits.
- `docs/hld/04-opc-and-packaging.md`, for relationship-resolved story and
  fragment handling, collision-free identity allocation, source mapping, and
  verbatim preservation of unmodelled package content.
- `docs/hld/10-bindings-spec.md`, for the native Word automation and comparison
  surfaces and the explicit Python, WASM, and CLI exposure boundaries.
- `docs/hld/12-testing-strategy.md`, for source-built rich merge records,
  pinned Word comparison differentials, policy matrices, save and reopen, and
  unchanged hash-harness expectations.
- `docs/hld/14-development-backlog.md`, for the F-233 through F-235 contracts,
  dependency order, acceptance gates, and the remaining M22 boundary.

## The wave

| F-ID | Title | Size | Status | Owner |
|------|-------|------|--------|-------|
| F-233 | Advanced mail merge | L | done | - |
| F-234 | Full-story document comparison | L | done | - |
| F-235 | Comparison granularity and ignore policy | M | done | - |

## Sequencing note

Rows are listed in dependency order, not F-ID order.

F-233 and F-234 depend on separate completed foundations and may proceed
independently. F-235 follows F-234 because its character, word, and ignore
policies refine the full-story comparison model and must reuse that single
alignment and source-mapping boundary.

## Definition of done for this sprint

- Nested merge regions and records resolve from multiple named data sources in
  deterministic order without leaving stale merge fields.
- Merge images, document fragments, paragraphs, lists, tables, and formatting
  hooks produce the expected rich output while preserving unrelated package
  parts and unmodelled XML.
- Comparison covers headers, footers, comments, fields, text boxes, footnotes,
  endnotes, formatting, and the main story in one stable source order.
- Accepting or rejecting generated comparison revisions reproduces the edited
  or original supported story content with source mappings intact.
- Character and word granularity plus each ignore policy change only the
  declared comparison records and remain deterministic across repeated runs.
- Invalid data, unsupported shells, relationship failures, and policy errors
  reject atomically without changing the source documents or their caches.
- Full verification passes with every deterministic hash explained, every
  package archive below 10 MiB, and the bounded sprint review clean.
