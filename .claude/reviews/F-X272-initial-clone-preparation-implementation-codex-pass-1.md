# F-X272 initial clone preparation, Codex implementation pass 1

**Scope**: Only the sealed seven-path initial-clone/preparation increment.
Source citations refer to `final/` under
`/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/initial-clone-preparation-increment`.
The existing Field preparation helper and ordinary preflight are prerequisite
source, not new API implementations in this delta.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. This is bounded local acceptance,
not whole C1, Native lineage or provider completion.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Checked boundaries

- Physical paragraph paths are constructed from the actual original main,
  table/control, note, comment and glossary graph at
  `crates/rdocx/src/template.rs:324`. The original-pointer lookup is confined
  to that clone producer at `crates/rdocx/src/template.rs:354`. Its stored
  result contains physical paths and frozen borrows, not candidate addresses,
  serialized equality or a persistent provenance catalog. Separate owner paths
  keep equal-content paragraph/Field occurrences distinct.
- The preparation list follows the actual existing preparation visitor order
  at `crates/rdocx/src/template.rs:365`, including unsupported None entries.
  The corresponding mutable sequence remains unchanged at
  `crates/rdocx/src/document.rs:25479`. Notes and comments retain their actual
  direct-paragraph preparation policy. Comment roots are now enrolled before
  operation creation at `crates/rdocx/src/template.rs:1582`, so a captured
  comment Field is not routed into an operation lacking that frozen owner.
- The initial clone occurs before the first package flush at
  `crates/rdocx/src/template.rs:1610`. The first template_programs flush does
  not replace typed main paragraphs: flush_document_to_package updates bytes,
  namespace and derived state, while its substantive canonical-source read
  borrows the model. No current initial-pass mutation changes the recorded
  paragraph traversal before it is consumed.
- Direct and nested-control run paths resolve through the actual frozen
  paragraph at `crates/rdocx/src/template.rs:425`. Any path containing a
  Revision is explicitly rejected at `crates/rdocx/src/template.rs:410`,
  leaving that reparse boundary on its prior route. Current setters for the
  admitted direct/control branches assign run slots without splicing the
  outer paragraph/control traversal. They do not shift later preparation
  ordinals.
- The existing preflight/raw-bypass decision runs before the shared helper at
  `crates/rdocx/src/document.rs:25441`. The lower sticky guard at
  `crates/rdocx-oxml/src/document.rs:5790` now precedes frozen-member lookup,
  allowance registration and cloning, including normalize false. The existing
  membership check at `crates/rdocx-oxml/src/document.rs:5795` still refuses
  foreign frozen Fields. Unsupported owners retain their old route at
  `crates/rdocx/src/document.rs:25443`. No raw XML is parsed or sanitized to
  discover correspondence, and no ordinary classification is inferred.
- Eq-equal prepared values remain discarded at
  `crates/rdocx/src/document.rs:25448`. Changed slots retire only after the
  replacement succeeds at `crates/rdocx/src/document.rs:25466`. Other run
  slots retain their actual clone correspondence. The wrapper disables all
  initial correspondence after the pass, including an error return, at
  `crates/rdocx/src/document.rs:25613`. Later refreshes, writers and reparses
  cannot reuse these initial claims.
- The three Native cases at `crates/rdocx/src/template.rs:3049` cover physical
  equal-content distinction, per-slot retirement and first-pass disable,
  ordinary model/package parity and actual exhausted-operation refusal for
  direct, nested-control and table paragraph owners. The lower raw-bypass
  control at `crates/rdocx-oxml/src/document.rs:25799` challenges the specific
  sticky-guard repair. These are partial-path tests, not proof of private
  original SDT models, revision reparses, slab coverage or later writer lineage.

## Evidence and limits

Authenticated SOURCE_BOUNDARY.json SHA256
21c4be89886523e9be9a9c506ae7c8e0362d83117131c895a01f8ea9b909934f,
693-line increment.patch SHA256
342189d79d45312530bbe4728ff45e4402915dc7066b46b6590bdb58dd425489,
all seven before/final source hashes, both retained executables and gate logs.
Before files are reconstructed accepted immediate per-path predecessors, not a
fresh pre-execution snapshot. No historical Base RED is claimed.

Root's actual retained results are45Native template PASS, including the3new
cases,30lower Word PASS and12import PASS. The one interleaved trace following
an actual named ok remains intact in the raw log. Sticky controlled-failure
and the earlier unproved writer-fixture attempt remain separately qualified.
Strict lower/Native all-target/all-feature Clippy completed in27.79seconds.
The retained metadata receipt records format, prose, adapter drift and diff
checks exiting0. This reviewer ran no builds or tests and changed only this
review record. Full lineage, remaining consumer closure and all final riders
remain due.
