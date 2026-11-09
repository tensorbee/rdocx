# F-X272, Codex final follow-up 2, constructor and planning investigation

REJECTED INCREMENT. This is a distinct Root read-only follow-up after final1 returned to orchestration. Final1 ended when its output was written. This pass investigates additional constructor-selection and planning-work evidence on the unchanged frozen source. No repair occurred between the passes. It is not a fourth model peer, a confirmation after a clean review, all-pass-12 or feature acceptance.

## Identity and provenance

Final1 remains immutable at .claude/reviews/F-X272-deepseek-handoff-codex-final-1.md, SHA-256 69c72b31d2688818d60d28b89c3ce5ea9822b41c32d0ac2829a95e17b688786a. Actual Grok source peer3 remains b49a1b3d67d5209f165abbd97a5dafabe3dd062d11c1e4349aa4bfeef4cdced6. Source branch sprint/s96, HEAD e2a31e58a9b0af20f0ec374e3561c2e599abb0cd and plan e37d2ded remain unchanged. All 923 original review source pins were reauthenticated with zero mismatches. No source edits or tests in this pass.

The coordinator initially appended these observations to final1 and labelled them a continuation. That was the wrong process label: AGENTS requires a review to end once its report is written. Final1 has been restored byte-for-byte. This file is the actual subsequent Root investigation report, and the intermediate transport is preserved in TEMP for provenance. No author's work or peer review is relabelled. The in-flight DeepSeek read-only reassessment has final1's four-finding input. Its next implementer must receive this follow-up before any Source write lease.

## Carried findings

C1 through C4 remain open from final1. The additional C1 cases below extend the same constructor-classification invariant, rather than replacing the original empty-inline discriminator. C5 is an additional independent planning-copy defect.

## C1, high, constructor selection includes opaque boundaries and local names

C1-B, missed Word-namespace raw child. crates/rdocx-oxml/src/document.rs:670 marks only its enumerated raw classes, and :694 treats an otherwise unclassified element as raw only when its namespace is not in is_modeled_capture_namespace. W_NS is considered modeled at :488 regardless of local name. CT_R actually captures unknown Start children verbatim at crates/rdocx-oxml/src/text.rs:4858 and unknown Empty children at :4933. For example w:contentPart is not modeled by this run parser. An empty w:contentPart with an inherited r:id is preserved raw, and its r binding is stored by external_bindings_used_by_raw at :4964. The preflight selects neither raw nor typed for that W_NS child and misses the dependency. A repeated-run boundary must compare predicted record storage with actual parsed records. This is source tracing, not an executed RED. Namespace membership alone is insufficient to determine whether a local name is modeled.

C1-C, invented nested owner inside an opaque or dropped subtree. crates/rdocx-oxml/src/document.rs:677 opens an r/hyperlink owner without checking parent_raw or parent_dropped. General opaque foreign children are also absent from the raw_here list at :670. The run parser's capture_element at text.rs:4858 consumes an unknown child's entire subtree without recursively constructing runs inside it. Likewise the drawing parser discards an unknown sibling subtree at drawing.rs:2231. Word-shaped descendants inside those subtrees therefore do not create the root-attribute model records the preflight at document.rs:678 counts. Use an opaque run child and a discarded drawing sibling each containing a w:r with a long inherited producer attribute, paired with a genuine modeled nested owner. Do not turn arbitrary opaque payloads into stronger import refusals.


C1 requires Start/Empty, modeled local-name selection, opaque-boundary and dropped-subtree agreement with the actual parser. A single empty-element predicate patch is not closure. Pair the opaque/dropped cases with genuine modeled owners. Keep arbitrary valid opaque payloads supported and preserve raw bytes.

### C5, high: capture planning itself repeats the full URI copy before admission

At crates/rdocx-oxml/src/document.rs:724 every selected owner dependency executes namespace.to_owned() into owner.needed. A later owner with the same long ancestor URI repeats that allocation. finish_original_capture_owner deduplicates only within that owner at :573, and only afterward discovers whether needed_scope already contains the pair at :578. Its global pair can be moved rather than copied once, but the next owner has already allocated another full URI. The final rejection at :785 therefore follows the planning copies. This is separate from whether expanded metadata construction counters stay zero.

The approved .claude/plans/F-X272-design.md:3653 explicitly requires physical declarations pooled once with stack indices and prohibits the URI-cloning discovery helper during planning. It also requires aggregate refusal before allocation for the new metadata work at :3666. Current owner.needed is an owned (String, String) list, not a borrowed dependency index. For N short runs using one long ancestor declaration, the planner performs N URI copies even when it ultimately rejects the expanded records. The cumulative copy-work mechanism is statically visible. No OOM, allocation counter value or executed runtime witness is claimed here.

Use the existing original-capture success/refusal pair with increasing owner count and a shared long URI, and measure planning URI-copy work separately from expanded record construction. A valid refusal with zero expanded records is insufficient if planning already performs the forbidden copies. Pool physical facts once, retain lightweight scope/dependency indices, and resolve lengths without per-owner URI materialization. Preserve actual namespace shadowing, value-only dependencies and current public interfaces. Do not repair this by adding headroom, ignoring needed dependencies, or claiming every original parsing allocation must be zero.


## Outcome and handoff

Five actionable finding IDs remain C1 through C5. Zero additional smells and nits. C4 accounts the copies selected constructors actually store. C5 bounds planning work itself. Static discriminators are not executed RED, allocation counts or observed OOM. No builds, tests, native actions or network calls in this audit.

The source model peer counter remains exhausted at three. No dependency is unblocked. All prior final-union, Issue430, all-pass-12, D8, format/clippy and unknown-result debt remains pending. Supply both final reports and the expanded C1/C5 evidence to the next implementer after the current read-only author handoff. Record the actual recipient and investigate/implement roles in the existing progress record when a write lease is granted. Author fixes then affected Root final recheck. This pass ends here and returns to orchestration.
