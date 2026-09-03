# Current Sprint, S66

**Milestone**: M22 Word depth.

**Goal**: update the field-driven structures that real reports depend on.
Extended evaluation must preserve unavailable instructions and intentional
cached results, while TOC rebuild uses headings, styles, entries, bookmarks,
and final page numbers without replacing unrelated field formatting.

## Spec references

- `docs/hld/03-architecture.md`, for recursive field grammar, pure facade
  evaluation, story traversal, bookmark ownership, cache updates, and the
  separation between Word field semantics and format-neutral page targets.
- `docs/hld/04-opc-and-packaging.md`, for namespace-aware parsing,
  schema-ordered mutation, fail-closed owner identity, and verbatim retention
  of unmodelled field and paragraph XML.
- `docs/hld/08-rendering-spec.md`, for displayed page numbers, bookmark target
  resolution, field substitution after pagination, and exact invalidation of
  cached field-bearing pages.
- `docs/hld/10-bindings-spec.md`, for the additive native Word facade boundary
  and explicit decisions about Python, WASM, and CLI exposure.
- `docs/hld/12-testing-strategy.md`, for the pinned Word 16.104 field matrix,
  source-built fixtures, deterministic pagination evidence, and external
  differential rules.
- `docs/hld/14-development-backlog.md`, for the F-231 and F-232 contracts,
  their dependency chain, acceptance gates, and the M22 completion boundary.

## The wave

| F-ID | Title | Size | Status | Owner |
|------|-------|------|--------|-------|
| F-231 | Extended field evaluation | L | in-progress | codex |
| F-232 | Dynamic table of contents rebuild | L | pending | - |

## Sequencing note

Rows are listed in dependency order, not F-ID order.

F-231 must complete first because F-232 depends on its TOC and TC field
semantics as well as the existing bookmark, update-policy, and pagination
foundations. The reviewed F-231 result then becomes the single evaluation
boundary used by TOC discovery, entry construction, and final page-number
substitution in F-232.

## Definition of done for this sprint

- TOC, TC, formula, mail-merge control, and barcode fields evaluate to the
  pinned Word results across supported story locations and formatting forms.
- Unavailable or unsupported instructions retain their original instruction
  text and cached display, with stable diagnostics instead of guessed values.
- An existing TOC rebuilds from headings, custom styles, outline levels, TC
  entries, bookmarks, and final displayed page numbers.
- Heading, style, and TC mutations produce the same entries, links, levels,
  and page numbers as the pinned Word update.
- TOC rebuild preserves unrelated field formatting, surrounding unmodelled
  XML, and package content through save and reopen.
- Full verification passes with every deterministic hash explained, every
  package archive below 10 MiB, and the bounded sprint review clean.
