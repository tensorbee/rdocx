# F-X272 physical simple-header final audit

Reviewer: Codex, delegated by Root.

Disposition: retain this bounded increment. Functional defects: 0. Structural smells: 0. Nits: 0. This does not approve complete C1, establish global original-source credit or add another implementation peer pass.

## Frozen scope

Reviewed the operative physical simple-header section at `.claude/plans/F-X272-design.md:4109` under the existing repository workflow and microscope rules. The packet is `/private/tmp/rdocx-s96-root-c1-implementation/simple-header-origin/REVIEW_MANIFEST.json`.

Current document SHA-256 is `b065626cda96902d6262a309ed9bc13386ec1a93a26c9084e903ccd26f1dd2fb`. Current text SHA-256 is `5551b26f48a7d757408850bf7b16d6e5136987c90f8d25284f90c14fadcdb198`. Both current source hashes and both patch hashes match the manifest. Applying each exact patch to its captured before snapshot in memory reconstructs the current file. No source edits, builds, tests or network operations were performed during review.

## Findings and observations

No findings.

The header lookup at `crates/rdocx-oxml/src/document.rs:1634` requires the same live borrowed input backing. It verifies both remaining length plus consumed position and the remaining slice address before consulting an exact opening range. The caller captures the event's start position before reading and requests the receipt before element capture at `crates/rdocx-oxml/src/text.rs:15248`. A copied allocation or reader over an unmapped subslice cannot qualify by equal content.

Receipts are appended in source order at `crates/rdocx-oxml/src/document.rs:2328`, supporting binary lookup. Original opening identity remains separate from lookup positions. The insertion transform at `crates/rdocx-oxml/src/document.rs:1585` shifts lookup spans, while lookup also requires unchanged original opening length. Prepared root closure and document-wrapper insertion therefore remap the existing occurrence rather than matching text or an ordinal.

All four production `controls_for` callers bind the matching actual consumed bytes: prepared closed paragraph at `crates/rdocx-oxml/src/document.rs:2895`, document parser at `crates/rdocx-oxml/src/document.rs:7661`, ordinary fragment at `crates/rdocx-oxml/src/text.rs:14915` and checked fragment at `crates/rdocx-oxml/src/text.rs:14961`. Empty and child control cursors retain no input binding. Unknown header projection and unstable primitive context do not publish a reusable header, and the projected-alias guard retains fallback. The cloned header carries its own syntax and dirty flag. It is consumed only by the matching started simple-field factory, retaining the existing empty dispatch and early/late result behavior.

The classifier stores its deferred instruction-only sum at `crates/rdocx-oxml/src/text.rs:21941`. `initial_work_bytes` adds the actual current result-cache length and scope at `crates/rdocx-oxml/src/text.rs:21812`. The prior formula consists entirely of checked nonnegative additions, so separating its instruction terms preserves the numerical result and overflow outcome. The Option retains deferred overflow instead of rejecting unused headers. The private syntax fields are not mutated between classification and admission. Materialization consumes them after admission. No new allowance or public failure behavior was identified.

The physical receipt control at `crates/rdocx-oxml/src/text.rs:22242` checks Unicode uppercase expansion, quoted operands and switches, distinct dirty flags and caches, and a skipped raw-control header between equal instructions. Copy and subslice fallback controls are at `crates/rdocx-oxml/src/text.rs:22286`. Explicit closure and wrapper mapping controls are at `crates/rdocx-oxml/src/text.rs:22332`.

## Saved evidence

The RED text snapshot, saved binary and actual log match their receipt hashes. The actual failure reports header-read counts `[5, 5]` instead of `[3, 3]` after the field instruction, cache, flag and skipped-control assertions pass.

The GREEN receipt pins both reviewed sources before and after execution. The actual log matches hash `a484c59f19b226529a9d6382f9d30e2db52eed6a43214c6b431d059de2dad933` and reports 303 passed, zero failed, child exit 0. Passing the physical receipt assertion establishes `[3, 3]` for the current fixture. The lower-crate all-target, all-feature Clippy with warnings denied, workspace formatting and prose logs match their recorded hashes and report exit 0. Prose reports zero violations.

Complete original OPC-part mapping, complex-field receipt integration, derived/captured correspondence, aggregate generated-work scheduling, remaining constructor modes, Issue 430 completion and final integrated native, schema, hash, package and CI gates remain open. These bounded receipts and intrinsic scalar facts do not establish those broader obligations.

Review ended after this report. Writer and build control return to Root.
