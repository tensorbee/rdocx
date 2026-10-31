# F-X273, all aspects, pass 10

**Reviewed**: D2 Owned paragraph structure-assignment remediation only, on `sprint/s96`. The real-before incremental patch changes2tracked files,123additions and1deletion across151patch lines. Production behavior changes only the existing Owned assignment arm in engine.rs. The remaining Rust additions extend an existing test body. HLD08 documents the invariant. Unrelated F-X275 dirty changes and proposed F-X271 A8 are excluded.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Pass9D2 is addressed for the examined publisher and consumers. This is a bounded remediation review, not final feature, PDF sample, raster, native or sprint acceptance.

## Contract and source identity

Read the canonical workflow and microscope command, current F273 plan, relevant HLD08 contract, existing progress and pass9. The approved plan requires an Owned paragraph to carry its exact companion at `.claude/plans/F-X273-design.md:549` and to preserve structure ownership at line606. The current plan SHA256 is `f39cf22966952420d61a45b8d88e26179442b3b2dc2ffb418b4ba55ef77f130c`.

Current engine.rs SHA256 is `23fe6c9669e01ea50f9431e971d68b8a69588e8a37f76a11b1f4692152cdf353`. HLD08 SHA256 is `656064fcef1e130bf507b2463b40921548054e39b89e7bc3254cfd81b1058bcf`. Both match the actual qualified packet. The existing provider-done record is the workflow's bounded dependency-prefix checkpoint, with the approved plan and final acceptance debt retained. This review does not grant final acceptance or change tracking.

## D2 publisher and consumer examination

At `crates/rdocx-layout/src/engine.rs:10602`, the Owned arm now binds its actual companion. After `assign_owned_paragraph_structure` completes physical paragraph/list allocation, line10608copies that exact block identity into the present companion. Both non-cache-safe/excluded-source and unretained-font-trace Owned constructors converge on this arm. Their initial captures at engine.rs:6414 and engine.rs:6523 are refreshed by the same publisher before pagination.

List and ordinary paragraph allocation remain in the existing function at engine.rs:10635. Both paths assign a concrete paragraph identity, with list paragraphs beneath their ListItem and ordinary paragraphs beneath the document or their existing heading role. Figure assignment still follows the physical paragraph assignment. No structure node is allocated twice by the repair, and list-frame lifetime is unchanged.

The companion remains authoritative. `crates/rdocx-layout/src/block.rs:1674` still selects its exact optional value, including None, and the Owned view still installs the actual companion at block.rs:1780. There is no fallback from a present None, no final-wrapper repair, and no change to neutral compatibility behavior. The new optional match does not conceal a new missing-companion producer. Both production Owned paragraph constructors supply Some. Owned tables and the existing table-only None test fixture use their separate plain table path.

Traced shared paragraphs at engine.rs:10592, shared table-cell paragraphs at engine.rs:10723, Owned tables, plain compatibility assignment, and the tbRl Owned-to-shared conversion at engine.rs:4830. The conversion precedes document structure assignment at engine.rs:5062 and subsequently enters the already synchronized shared arm. No equivalent stale-companion assignment remains established in these examined routes.

Painting consumes `para.structure_id()` at `crates/rdocx-layout/src/paginator.rs:6462`. Continued blocks copy that same selected identity at paginator.rs:5574. Refreshing the producer therefore repairs both first-fragment and continuation ownership rather than changing only one PDF wrapping site. HLD08 at `docs/hld/08-rendering-spec.md:1867` accurately records this synchronization and retains authoritative companion semantics elsewhere.

## Meaningful tests and actual retained evidence

The added assertions in the existing `shared_cached_blocks_keep_result_local_semantics_exact` body begin at engine.rs:22194. A normal shared control, a bookmarked non-cache-safe Owned body paragraph, and a numbered non-cache-safe Owned list paragraph are actually laid out. The latter two are long enough to paint across more than one physical page. The fixture asserts the intended admission boundaries and retains the earlier shared/table cache positives.

The test walks actual MarkedContent and Group paint trees, selecting the relevant Text leaves rather than merely checking for any semantic mark or node count. Every selected leaf must carry an identity. Across all pages each physical paragraph has exactly one identity, which must reference a real Paragraph node with the expected Document or ListItem parent. Missing text, absent continuation text, an Artifact wrapper or unrelated non-paragraph identity fails these checks. Complete cold, warm and fresh layout/provenance results are compared before the per-page assertions. The ASCII fixture intentionally exercises these actual Text leaves. It does not claim new native or universal script coverage.

Independently authenticated all127sealed packet members with zero hash or size mismatches. Packet is `/private/tmp/rdocx-s96-run-refresh/fx273-owned-structure-remediation-v1`, seal SHA256 `017baf350895dec53d91c37ab5c23e95b5381a296ebe085509445c98d0933e6a`, handoff SHA256 `3b537ffbc55b90775ebeafe62291f09bf4cbb123c889f66b18448d5e1f0e4bed`.

The actual production-unfixed test executable fails at the intended `ownedbody body paint lost its assigned structure` assertion. The retained RED engine snapshot matches its acquisition source hash `0ea5f1c261b0c16ffae57ac924d45408805624faf3d2f50d192736af60d1e54b`. This is genuine semantic RED, distinct from the retained initial missing-import compile failure. Shared positives and cold/warm/fresh comparisons execute before the RED assertion. The later list assertions were not reached on RED and are not credited as an observed predecessor list failure.

The actual fixed batch has five named executions, each1passed,0failed and0ignored: extended shared-cache semantics, plain document-order structure, behind-document figure ownership, body-only layout/context transfer, and font-trace request ownership. Their receipts and raw stdout/stderr authenticate. The coherent actual Cargo artifact records opt1, debug assertions and overflow checks, with the8MiB test-stack launch setting. Frozen fixed executable SHA256 is `a841bedcf9e55443eff1129a62a2780416fdd2c8eb6d0a85ba50384ad588c668`.

Affected all-target check, strict rdocx-layout all-target/all-feature clippy with denied warnings, workspace fmt, prose, adapter drift and diff checks have actual exit0receipts. For every examined deciding fixed command, before/after source maps equal the final3851entry inventory and retained streams authenticate. Comparing packet before/after inventories establishes exactly engine.rs and HLD08 changed, with3849other entries unchanged. Root's subsequent separately declared F271 planning/state metadata changes are not represented as executions at that later metadata identity. The current F273 plan and two deciding source pins remain exact.

## Defects

Zero. Pass9D2 is addressed within this bounded source and local paint/continuation qualification.

## Smells

Zero.

## Nitpicks

Zero.

## Not found

**Correctness**: zero additional findings after tracing both Owned producers, assignment, selected view, paint and continuation consumers.

**Contract**: zero findings. The repair preserves the exact present-companion rule, existing public contracts and the declared HLD impact.

**Panics**: zero findings. The production change adds only optional binding and assignment, with no new indexing, allocation, arithmetic or unwrap. New test unwraps validate fixture and painted-output expectations.

**OOXML**: zero findings. No parser, serializer, namespace, schema order or raw source behavior changes.

**Tests**: zero findings. The actual RED distinguishes the stale producer, and fixed assertions inspect meaningful painted ownership, list parents, continued pages and cache lifecycle.

**Structure**: zero findings. No new type, public API, module, file, generic, trait, framework or forwarding wrapper was introduced. The small assignment remains at the existing shared publisher.

## Handback and pending obligations

No final sample PDF reproduction,49hash/7pixel disposition, baseline approval, full workspace/risk union, installed Python/typing, WASM/API, current hosted CI, pinned external oracle or native Word qualification is supplied by this review. Those obligations retain their existing status. Proposed script-fixture amendments, F272/F274 Python failures and F275/F271 remediation are outside this pass. No prior failed output is relabelled fixed-source PASS.

Only this review file was written. No source, progress, configuration, TEMP or other file was modified. No product import, build, test, network or native action ran. End the review pass and return the canonical review-file writer lease to root for a separate phase.
