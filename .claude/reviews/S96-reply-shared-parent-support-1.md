# S96, shared reply-parent read, bounded support pass 1

**Reviewed**: The frozen reply-only increment over the retained pre-reply source at H10 `39287e8cb9c4a014adf0b3191cbe3cdf99c20569`. Three tracked paths, 561 additions and 188 deletions. The separately reviewed cold-owner increment is excluded.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Identity and scope

The handoff at `/private/tmp/rdocx-s96-prepared-integration-readiness/reply-shared-parent-read-proposal-1/implementation-1/IMPLEMENTATION_HANDOFF.json` authenticates as SHA `96908f0f6b57c742eb4fcfe1b9a7a3708c6ba9a362a8a3bd640a5131391c4a35`. Its reply-only patch authenticates as `af4c2ee4ffba1cf1464882c907c1c6fc3ae8793469f4c9a646302290aa49d42c`. Current comments source matches `d0c2d20f3f916c50e9517444fcac05ff946894a190704cc022bde4d97fb428c3` and the final complex-control source snapshot. Plan, HLD and progress pins match. All five prior cold-source/review preservation pins also match.

All eight execution receipt hashes, 34 raw log hashes and both binary-receipt hashes authenticate. The final four-control binary authenticates as `c64efc33bea25c1074942d8a47db418054cf55da2f2dd91597ec1a08adfdde47`. Earlier broader results retain their earlier source identity. They are not relabeled as executions of the final test refinements.

This review follows the existing workflow and microscope aspects. It is bounded F268 implementation support, not a new feature peer identity, cap reset, timing acceptance or final sprint review.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Source assessment

- `crates/rdocx/src/comments.rs:2026` flushes the actual private reply candidate before constructing the shared read. Projection and reference qualification consume that same owned read. The read and test-only phase guard are dropped before comment identifier reservation at `crates/rdocx/src/comments.rs:2036`. Later anchor publication remains on the existing path.
- `crates/rdocx/src/comments.rs:3083` performs the existing complete ownership graph read, selected identity check and all-paragraph source/factory collection before projection. The paragraph vector retains actual location, parsed paragraph, marker IDs, original source bytes, flow and physical offset. Unrelated paragraphs are not skipped. The ownership graph still uses the existing complete marker inventory at `crates/rdocx/src/comments.rs:3904`.
- `crates/rdocx/src/comments.rs:3140` preserves the former projection rules for orphan/reference-only comments, marker cardinality, physical flow order, accepted source projection, newline composition and ordered endpoints. Flow indexing and the registered-paragraph expectation follow the same construction invariant as before the extraction.
- `crates/rdocx/src/comments.rs:3228` retains the former reply-specific checks: exact start/end/reference counts, same part and story, physical end before reference, exact accepted endpoint location and run index, unique reference, and metadata-only reference payload. It now reads the already admitted paragraphs instead of repeating their factory pass. It does not substitute IDs, text equality or a route guess for actual source correspondence.
- Public snapshot behavior remains clone, flush, complete read and selected/all projection at `crates/rdocx/src/comments.rs:3524`. The outer public reply transaction at `crates/rdocx/src/comments.rs:1920` still validates operands, stages a candidate, flushes dirty related models and commits only after success. Unknown-parent checks and date validation still precede parent-read work. No public signature, Source ingress, Work allowance, namespace policy or retained cache changes.
- The approved finite scope is stated in `.claude/plans/F-X268-design.md:763`. The required HLD description is present at `docs/hld/03-architecture.md:2616`. The private five-field reader owns real data used by both consumers. No forwarding-only wrapper, trait, generic, feature, module or file was introduced.

## Evidence assessment

The actual predecessor test first checks successful reply semantics and reopened anchor state, then fails the mechanism assertion with two parent all-paragraph passes and seven parsed paragraphs. Its exit101 is retained separately from initial test compilation failures.

The final four controls at `crates/rdocx/src/comments.rs:4923`, `crates/rdocx/src/comments.rs:4954`, `crates/rdocx/src/comments.rs:5059` and `crates/rdocx/src/comments.rs:5132` have an authenticated four-pass result. They distinguish one actual parent factory pass, successful crossing reply and reopen, first-flush versus former cloned second-flush source/projection/reference equivalence, unknown-parent precedence, unrelated malformed-header admission, missing references, nonmetadata reference payload and rollback of package/model/identifier state. The equivalence cases include unsaved current content, simple and actual parsed complex styled/tab Fields, tracked content, aliases and a header parent.

The broader comments log records 39 passes and one existing ignored Word oracle on unchanged production logic before the final test-only refinements. It includes duplicate/split-story, multi-paragraph/header/cell, crossing/repeated replies, imported metadata, reference ordering and late relationship-exhaustion rollback. The final strict lint, format, prose and diff receipts record exit0. These are actual local observations, not native Word admission or performance evidence.

The parent counter is test-only and bracketed before identifier allocation. It counts calls to the real all-paragraph factory sites. Accepted-comment projection probes and later anchor parsing remain possible and are not claimed absent. The actual parsed-complex fixture correction preserves its rich Field premise without requiring an unrelated optional stored-owner cache. Failed compiler, lint and fixture attempts remain retained.

## Limits and handback

Correctness, contract, panic/index boundaries, OOXML preservation, test discrimination and structure each produced zero findings in this increment. Source pins were rechecked before handback. No source edit, build, test execution or product import was performed during this review.

Root final review, a fresh ordinary artifact, original reply and other flat-cost thresholds, cold407 acceptance, native evidence and final integrated memory/release riders remain pending. No threshold, native criterion, status, ledger or prior finding is waived by this report.
