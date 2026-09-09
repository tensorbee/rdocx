# Current Sprint, S72

**Milestone**: M23 From-scratch business documents.

**Goal**: expose every document section in stable order and provide one
relationship-safe content model for the body and every related story required
by from-scratch generation. Public mutations must preserve schema order,
unmodelled XML, story ownership, and cross-part relationships while committing
only complete validated candidates.

## Spec references

- `docs/hld/02-scope-and-non-goals.md`, for the DOCX-015 through DOCX-021
  capability boundaries and the precise partial surfaces S72 completes.
- `docs/hld/03-architecture.md`, for WordprocessingML grammar ownership, one
  facade-owned story model, source identity, and staged package mutation.
- `docs/hld/04-opc-and-packaging.md`, for relationship-scoped story parts,
  header and footer ownership, fragment dependency closure, and atomic commit.
- `docs/hld/05-drawingml-model.md`, for schema-ordered picture and DrawingML
  payload semantics that asset movement and fragment import must preserve.
- `docs/hld/08-rendering-spec.md`, for section geometry, story isolation,
  header and footer selection, and source-preserving layout behavior.
- `docs/hld/09-charts-spec.md`, for Word chart, editable-workbook, drawing, and
  relationship closure ownership that import must remap atomically.
- `docs/hld/10-bindings-spec.md`, for the native Rust facade boundary and the
  public section, story, relationship, and fragment-import surface.
- `docs/hld/12-testing-strategy.md`, for synthetic round-trip, differential,
  relationship-scope, fragment-remapping, and atomic-failure gates.
- `docs/hld/13-risks-and-open-questions.md`, for the rule that cross-part Word
  invariants are staged and validated as one operation.
- `docs/hld/14-development-backlog.md`, for the F-250 through F-256 acceptance
  contracts, dependencies, sizes, and M23 sequencing.

## The wave

| F-ID | Title | Size | Status | Owner |
|------|-------|------|--------|-------|
| F-253 | Container-neutral story editing | L | pending | - |
| F-250 | Ordered mutable section facade | L | pending | - |
| F-251 | Complete section and page geometry | L | pending | - |
| F-254 | Generic insert, move, clone, and remove operations | L | pending | - |
| F-255 | Part-scoped assets, links, and relationships | M | pending | - |
| F-252 | Rich per-section headers and footers | L | pending | - |
| F-256 | Transactional cross-document fragment import | L | pending | - |

## Sequencing note

Rows are listed in dependency order, not F-ID order.

F-253 and F-250 establish the independent story and section ownership
foundations. F-251 may proceed once ordered section mutation exists, while
F-254 and F-255 build generic mutation and part-scoped relationships on the
container-neutral story model. F-252 converges the section and story tracks so
each header and footer variant can own rich related content safely. F-256 lands
last because fragment import must remap the complete package, style, numbering,
section, story, asset, and relationship surface delivered through F-255.

## Definition of done for this sprint

- Ordered lookup, insertion, removal, and mutation preserve portrait,
  landscape, portrait section order and never orphan a related story.
- Authored size, orientation, margins, gutter, columns, page numbering, header
  and footer distance, title-page state, and break type match pinned Word page
  geometry and numbering while unsupported children remain visible.
- Default, first, and even headers and footers can be created, linked,
  inherited, replaced, and removed per section with rich content that survives
  save, reopen, replacement, and inheritance changes.
- One public location and traversal model visits supported body, cell, header,
  footer, note, comment, and text-box content with identical mutation errors
  and without introducing a second document tree.
- Arbitrary-position insert, move, clone, and remove operations preserve exact
  content order, references, and untouched raw XML, with invalid operations
  leaving the document unchanged.
- Images, hyperlinks, charts, and other related content resolve only through
  the relationships of their owning body or related story part.
- A dependency-rich cross-document fragment can be imported repeatedly with
  deterministic remapping and no collisions, while any unsupported dependency
  aborts without changing the destination.
- Every operation preserves unmodelled XML and schema child order, publishes
  only a complete staged document and package candidate, remains deterministic,
  and leaves the hash harness unchanged unless a separately labelled and
  reviewed behavior change declares the expected delta.
