# Current Sprint, S68

**Milestone**: M22 Word depth.

**Goal**: expose modern Word package content that is currently preserved but
opaque. Executable objects and macros remain non-executing inventory surfaces,
while forms, glossary entries, AutoText, and building blocks gain bounded typed
access only inside modern OOXML packages.

## Spec references

- `docs/hld/02-scope-and-non-goals.md`, for the permanent exclusion of binary
  `.doc` and execution of embedded OLE, ActiveX, or VBA payloads.
- `docs/hld/03-architecture.md`, for relationship-owned payload identity,
  facade ownership, bounded staged mutation, and preservation of unsupported
  document content.
- `docs/hld/04-opc-and-packaging.md`, for normalized internal relationships,
  content-type validation, signature infrastructure, fail-closed package
  graphs, and atomic publication.
- `docs/hld/10-bindings-spec.md`, for additive native Rust inventory and
  mutation surfaces without implicit Python, WASM, or CLI exposure.
- `docs/hld/12-testing-strategy.md`, for stable payload hashes, relationship
  paths, source-built malformed graphs, save and reopen, and byte-exact
  unsupported-subtree preservation.
- `docs/hld/14-development-backlog.md`, for the F-236 and F-237 contracts,
  dependency boundary, acceptance gates, and the remaining M22 work.

## The wave

| F-ID | Title | Size | Status | Owner |
|------|-------|------|--------|-------|
| F-236 | Embedded object and macro inventory | L | in-progress | codex |
| F-237 | Forms, glossary, and building blocks | L | pending | - |

## Sequencing note

Rows are listed in dependency order, not F-ID order.

F-236 depends on the completed digital-signature verification and creation
foundations, F-171 and F-172. F-237 has no unfinished dependency. The two S68
stories may proceed independently because one owns relationship-backed opaque
payloads and the other owns typed WordprocessingML forms and reusable content.

## Definition of done for this sprint

- Embedded objects, ActiveX controls, VBA projects, and their signatures have
  a stable relationship-owned inventory with exact hashes and package paths.
- Extraction and bounded replacement return or retain exact payload bytes
  without decoding or executing them.
- Safe removal leaves a valid document, applies the declared signature policy,
  and preserves every unrelated part, relationship, and owner subtree.
- Legacy form fields inside modern OOXML, glossary entries, AutoText, and
  building blocks are typed and editable without adding a binary `.doc`
  reader.
- Supported entries survive save and reopen, and every unsupported subtree
  remains byte-identical through unrelated edits.
- Invalid identities, external or malformed relationship graphs, wrong content
  types, signature inconsistencies, and failed staged mutations reject
  atomically.
- Full verification passes with every deterministic hash explained, every
  package archive below 10 MiB, and the bounded sprint review clean.
