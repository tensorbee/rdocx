# F-X272, nested consumer functional support, pass 1

**Scope**: Bounded independent support review of the immutable `ROOT_TAKEOVER/root-nested-consumers-review` packet. This is not the whole F-X272 implementation pass and does not reset its existing peer-review lineage or caps. Root remains the sole working-source author.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in the reviewed nested-consumer delta.

## Source and contract

The packet is under `/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/returned-reader-functional-batch/`. All ten source snapshots match their recorded SHA-256 values in `SOURCE.json`. The retained lower executable matches `3e6d4996b8c34e565538666065d9d486c0d0fa08c24eff92768cd38d63a3ef9f`.

Compared lower `document.rs` and `text.rs` with `root-source-closure-review`, and the other five lower paths with `root-returned-private-repair`. The resulting delta has 161 added and 18 removed lines across six changed files. `footnotes.rs` is unchanged against its specified predecessor. The three Native snapshots are authenticated context, not a newly reviewed Native implementation delta.

The canonical plan's returned-reader contract governs this review. `.claude/plans/F-X272-design.md:6292` requires the same successful factory to retain actual physical Text paths and origins. `.claude/plans/F-X272-design.md:6306` limits borrowed binding to the paired returned model and Source without enrolling another input. The current callable contract remains 42 operations and seven borrowed accessors at `.claude/plans/F-X272-design.md:6459`.

Source citations below refer to the immutable packet's `source/` tree, not the moving worker.

## Defects

0 defects.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Reviewed mechanisms

- `crates/rdocx-oxml/src/table.rs:2576` and `crates/rdocx-oxml/src/content_control.rs:3207` collect nested receipts through the same paragraph factory. The receipts move into the successful paragraph's own facts, while its extra XML and ordinary-property outputs remain separately retained. A failed paragraph does not publish a successful child record.
- `crates/rdocx-oxml/src/document.rs:5190` resolves recorded paths against the actual supplied owner before recursively resolving that paragraph's nested receipts. Direct run/content slots retain their existing selected-model checks. The dictionary uses addresses only for invocation-borrowed objects already selected by those producer paths, rather than discovering provenance from an arbitrary model address or equal value.
- `crates/rdocx-oxml/src/document.rs:5078` consumes the synthetic `TableRow(0)` prefix only when the actual supplied root is a Row. Other owner/path combinations retain their existing validation. The row reader's real table boundary supplies this prefix, so the change does not flatten nested table paths generally.
- `crates/rdocx-oxml/src/placeholder.rs:797` prepares one dictionary from the actual returned row before either existing direct/all view consumes its Text objects. `crates/rdocx-oxml/src/placeholder.rs:850` and `crates/rdocx-oxml/src/placeholder.rs:879` borrow that same dictionary for selection and aggregation. Repeated views reach the same origin/decoded interval and retain the existing deduplication and copy-demand behavior.
- `crates/rdocx-oxml/src/placeholder.rs:539` resolves additional raw-wrapper models through their own successful reader facts before selecting their nested Text objects. The wrapper contents remain alive and unchanged during selection. `crates/rdocx-oxml/src/placeholder.rs:3294` similarly binds the actual textbox callback model. The same paragraph factory receives optional outputs at `crates/rdocx-oxml/src/placeholder.rs:3706`, rather than adding another qualification parse.
- `crates/rdocx-oxml/src/text.rs:20759` keeps the ordinary paragraph entrypoint on the all-None route. Optional nested outputs do not enroll a source or add semantic/current Field inputs. Empty/default companions add no positive facts. Existing failed/raw-fallback constructor staging remains authoritative.

## Evidence assessment and limits

The two new tests at `crates/rdocx-oxml/src/document.rs:31273` and `crates/rdocx-oxml/src/document.rs:31297` use entity-bearing Text under revisions and controls, assert the decoded nine-byte source allowance, and reach a charged copy. The textbox test includes direct paragraph, table, block control and custom XML wrapper carriers. These assertions reject unchanged displayed text accompanied by lost input facts.

The preserved initial RED log records both actual allowance mismatches, zero versus 1116. The intermediate `root-nested-consumers-green-1.log` is not an all-green receipt, it records the subsequently repaired Row-root path mismatch. The final `root-nested-consumers-scoped-batch-2.log` records 200 passed and zero failed. This review consumed those existing observations and inspected the test bodies. It ran no Cargo jobs, tests, native applications or oracles.

Other changed or rebuilt revision paths, comment/glossary producers, non-emitting paths and final provider closure remain pending their own actual evidence and reviews. Native cohorts and strict gates reported as underway are not accepted here as completed receipts. Final integrated verification, external compile compatibility and mandatory riders remain due. No whole-provider or feature completion follows from this support pass.
