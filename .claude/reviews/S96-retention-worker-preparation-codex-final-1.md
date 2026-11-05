# S96, retention phase/history preparation, Root final1

Reviewed frozen increment2 in memory-reconstruction/variant-preparation-20261104,
including native_worker_body.rs, its composition patch, both collectors,
synthetic controls, custody and protocol handoff. No product build or run occurred.
All ten frozen custody hashes were independently authenticated. The reviewed
worker body hash is
`907ecd4e11b92761e2c3b12bb79452bf24da191cc710e2408317af3b81793840`.

Verdict: one defect, zero smells and zero nitpicks.

## Defects

### D1, a cached namespace proof failure counts as successful activation

`/private/tmp/rdocx-s96-final-runner-recovery/memory-reconstruction/variant-preparation-20261104/native_worker_body.rs:136`

The namespace activation counter uses
`empty_namespace_owners.get().is_some() || namespace_spans.get().is_some()`.
Both actual slots are OnceLock of Option proof. Their documented None means
failed, ineligible or unproved, rather than successful empty or span evidence,
as shown in `crates/rdocx/src/document.rs:8382`. Once either slot is initialized
to None, the outer get is Some and this predicate reports positive namespace
activation. The resulting history receipt can therefore satisfy its positive
activation condition without an actual successful namespace proof payload.

Require a populated inner proof for this successful activation metric, or
record attempted/negative memo initialization separately from successful proof
activation. Do not initialize a proof in the observer. Add a discriminator for
unset, initialized None and actual Some proof. Preserve the frozen first packet
and distinguish source/predicate controls from product runtime evidence.

## Not found

No additional concrete correctness or structural finding was established in
the proposed phase collector and bounded observer. Actual targets and edit
values remain explicitly proposed. Current/peak RSS, native diagnostic builds
and ordinary shipping artifacts remain distinct. The Weak observer does not
claim complete history coverage or invent a live-node budget.

## Pending

The exact phase mapping and companion activating history fixture still require
technical review. The minimal fixed RSS fixture need not activate the required
namespace/decoder/scoped proofs. Individual omission-history, inline/shared
ownership, complete All, final SourceH, types/builds, semantic parity and actual
paired-process measurements remain unqualified. Existing fake collector tests
do not close those requirements. This report ends review and returns to
orchestration for separate TEMP remediation.
