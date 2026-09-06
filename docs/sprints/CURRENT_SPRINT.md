# Current Sprint, S69

**Milestone**: M22 Word depth.

**Goal**: close Word depth without opening a legacy `.doc` programme. Flat OPC,
DOCM, DOTX, DOTM, and bounded MHTML share the current document model, while the
strict XML lexical checks exposed by S68 move to one existing shared layer
without weakening any owner-specific fail-closed contract. After the M22 end
gate is clean, recover the immutable partial v0.13.0 attempt through a coherent
shared 0.11.0 family and the exact stable Rust family at v0.13.1.

## Spec references

- `docs/hld/02-scope-and-non-goals.md`, for modern OOXML package-class support,
  bounded interchange, and the permanent exclusion of binary `.doc`, Word 2003
  XML, and executable payload execution.
- `docs/hld/03-architecture.md`, for the lowest shared crate boundary, facade
  ownership, source-preserving mutation, and avoiding duplicate parser policy.
- `docs/hld/04-opc-and-packaging.md`, for normalized relationships, content-type
  ownership, package-class identity, safe resource resolution, and atomic
  package publication.
- `docs/hld/10-bindings-spec.md`, for additive native Rust surfaces and explicit
  binding boundaries for new package and interchange operations.
- `docs/hld/12-testing-strategy.md`, for source-built malformed inputs,
  differential import and export evidence, save and reopen, byte preservation,
  and deterministic harness expectations.
- `docs/hld/14-development-backlog.md`, for the F-238, F-239, F-X077, F-X078,
  F-X079, and F-X080 contracts, completed prerequisites, acceptance gates, and
  the M22 end gate.

## The wave

| F-ID | Title | Size | Status | Owner |
|------|-------|------|--------|-------|
| F-X077 | Share strict XML lexical validation | M | done | - |
| F-239 | MHTML import and export | M | done | - |
| F-X080 | Restore CI release readiness | S | done | - |
| F-X079 | Tag rpptx-v0.10.0 | S | done | - |
| F-238 | Flat OPC and modern Word package variants | M | done | - |
| F-X078 | Tag v0.13.0 | S | archived | - |
| F-X081 | Tag rpptx-v0.11.0 | S | pending | - |
| F-X082 | Tag v0.13.1 | S | pending | - |

## Sequencing note

Rows are listed in dependency order. F-X077 builds on the completed F-236 and
F-237 scanners and establishes the shared lexical boundary first. F-X080
restores the hosted CI release gates after F-X077 and F-239 have settled their
binding surface. F-X079 then publishes the new `oxml-core` API as the
incubating family at 0.10.0 only after CI is locally reconstructed and clean.
F-238 builds on F-236, F-X077, and the published F-X079 graph, while F-239
builds on F-178's HTML import foundation.

F-X078 ran after the representative M22 end gate passed. Its immutable
v0.13.0 tag published five low-level stable crates, then registry verification
proved that F-238's new `oxml-opc` constants were newer than the published
shared 0.10.0 archive. F-X081 therefore publishes the complete shared and
PowerPoint family at 0.11.0. F-X082 then publishes the complete stable family
at 0.13.1 against that registry boundary. Each recovery release pauses for
separate final approval at its exact reviewed SHA.

## Definition of done for this sprint

- Flat OPC, DOCM, DOTX, and DOTM read and write through the current document
  model while retaining package identity, macros, template semantics,
  relationships, content types, and unsupported XML.
- Every modern package class saves, reopens, and passes its no-repair structural
  gate without changing executable payload bytes.
- Bounded MHTML import and export preserve body order, formatting, tables,
  lists, images, links, and stable declared loss records through source-built
  differential fixtures.
- MHTML resource resolution rejects unsafe, malformed, external, or over-limit
  inputs before publishing partial output.
- Embedded, glossary, and package-story malformed XML matrices execute through
  one shared lexical validator while retaining their existing error surfaces
  and byte-identical mutation rollback.
- Hosted CI package inventory, Pandoc installation, and Python binding jobs
  have mutation-sensitive local regressions and pass their reconstructed
  release-readiness gates.
- Binary `.doc`, Word 2003 XML, executable payload interpretation, and
  permissive XML recovery remain out of scope.
- The representative modern M22 document authors and renders equations,
  rebuilds fields and a table of contents, performs advanced merge and
  comparison, inventories embedded content, and round-trips its modern package
  variant without losing unsupported XML or executable payloads.
- The immutable v0.13.0 attempt and its five published low-level packages are
  recorded accurately without moving or deleting its tag.
- The complete shared and PowerPoint family publishes at 0.11.0 before the
  exact seven-package stable Rust family publishes at 0.13.1. Both receive
  separate final release approval and have their registry entries, owners,
  annotated tags, GitHub release bodies, selected-family exclusions, and
  applicable contribution notifications verified.
- Full verification passes with every deterministic hash explained, every
  package archive below 10 MiB, and the bounded sprint review clean.
