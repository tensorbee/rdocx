# F-X272, Root glossary omission borrow repair, pass 2

**Scope**: Independent review of Root's targeted S1 repair only. The reviewer authored the preceding glossary omission increment, so this pass does not recertify that increment. Root independently reviewed it in pass1. This is not a whole F-X272 peer pass or a review-cap reset.

**Verdict**: S1 closed. 0 new defects, 0 smells, 0 nitpicks in the repair.

## Exact source

Reviewed the immutable `ROOT_TAKEOVER/root-glossary-borrow-repair/source` snapshots under `/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/returned-reader-functional-batch/`. Both source hashes authenticate against `SOURCE.json`, SHA-256 `494aa2c62e8225416f523d72144c9d1b104d37f2a0f6ad96cb3d9ea74c9e17ce`. The small two-file patch authenticates as `52cb17f5f77f11f42fbd7cf67fb95b901eacf7c8d75a2c3bcb30cfeace0f79b3`.

Source citations below refer to those immutable snapshots.

## S1 disposition

The private selector now borrows its Source at `crates/rdocx/src/document.rs:14693`. Its returned Source remains owned. All selected slices, namespace closures and final fragment assembly continue through the same checked producers at `crates/rdocx/src/document.rs:14711`, `crates/rdocx/src/document.rs:14718`, `crates/rdocx/src/document.rs:14745` and `crates/rdocx/src/document.rs:14763`. Thus the repair removes an ownership-only complete-parent allocation without exempting an actual selected output copy from its reservation.

The glossary omission caller passes its actual retained parent directly at `crates/rdocx/src/document.rs:36207`, replacing the unnecessary complete-main repeat. The repeated-fragment importer similarly borrows the actual returned shared Source at `crates/rdocx/src/document.rs:47838`, removing its ownership-only whole-source composition. Its subsequent repeat and final byte delivery remain at `crates/rdocx/src/document.rs:47844`.

The template replay caller borrows its actual wrapped Source at `crates/rdocx/src/template.rs:1901`. The real wrapper composition at `crates/rdocx/src/template.rs:1896` is retained because it produces the document envelope consumed by this selector. It is not incorrectly removed with the two ownership-only copies.

All three actual callers keep the supplied Source alive throughout selection and receive an independently owned output. Their actual Source identity and selected ranges remain the producer arguments. No source-by-equality lookup, new ingress, public callable, model cache or quota is introduced. The public contract remains42 carriers and7 borrowed accessors.

## Defects

0 new defects.

## Smells

0 remaining smells in the targeted repair. S1 is closed.

## Nitpicks

0 nitpicks.

## Evidence and boundaries

The retained logs record2new omission tests,21scoped template tests and18versioned template tests passing. `RESULTS.json` now records both related cohorts and strict lower/Native all-targets/all-features Clippy at exit0, with strict duration23.21s. These are author-run observations inspected by this reviewer, not fresh reviewer executions. No Cargo job, native application or oracle was run during review.

The original increment's broader semantic assessment remains Root's pass1 review. Comments dirty emission, rebuilt revisions, non-emitting paths, whole-provider closure, canonical/HLD reconciliation, external compatibility and all final integrated riders remain pending. This report closes only S1 and does not mark F-X272 or its provider complete.
