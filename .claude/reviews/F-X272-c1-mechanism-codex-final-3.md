# F-X272 C1 mechanism, Codex targeted review 3

**Reviewed**: completed DeepSeek V11 proposal, 125203 bytes, SHA-256 ccc7b56ec65c09116d1eb61d212ac8725b8323e643b180c66f642bea06576b0c. Author terminal 91142 ended naturally with exit 0. Proposal, handoff and seven probe artifacts are frozen in /private/tmp/rdocx-cross-review/c1-v11-final-frozen. This reviews a proposed mechanism, not product implementation.

**Verdict**: 2 defects, 1 smell, 0 nitpicks. C1 remains unapproved. No fourth design peer is consumed. Root supplied factual triage to the same unfinished TEMP author phase in V12 before writing this record because F283 held the sole canonical writer lease. V12 is not reviewed or accepted here. No product source changed.

## Defects

### C1-M1, nested instruction fields are not permanently empty

/private/tmp/rdocx-deepseek-harness/fx272-design-rescue/DEEPSEEK_MECHANISM_V11.md:427 derives a vacuous read-time model admission from the initial empty fields in Field::parsed. That derivation stops before publication. crates/rdocx-oxml/src/text.rs:8615 and :8631 populate parsed_cached_runs after initialization. The same parsed field is moved into the parent instruction at :8661. The tokenizer retains the actual Field in a boxed nested operand at :20018. resource_owner_children at :2372 traverses populated run caches as well as instruction operands.

A complex field used in an outer instruction can therefore contain retained result runs and fields in those runs. V11 incorrectly excludes every such run, drawing and model descendant from its proof. The required outcome is exact preservation of the existing admit_word_model check at :1322 before construction, either through a complete proof that the global guard dominates its actual populated model or through explicit model-shape facts. Empty initialization alone proves neither. No new refusal or model-limit countercase is asserted by this review. The earlier 257 ordered-result example was unproved and was repeated in Root's task. Correcting that inherited premise does not establish V11's empty-model conclusion.

### C1-M1, the exact scalar work formula differs from the constructor

/private/tmp/rdocx-deepseek-harness/fx272-design-rescue/DEEPSEEK_MECHANISM_V11.md:409 omits text argument lengths from the claimed exact initial_field_work_bytes sum. crates/rdocx-oxml/src/text.rs:1279 adds 64 per argument and :1281 additionally adds text.len() for a Text argument. The proposal also equates concatenated instruction text with FieldInstruction.raw, while :20090 stores raw.trim().to_owned(). An instruction with surrounding spaces and a quoted text argument distinguishes both errors. Parsed name length must likewise use the actual token transformation.

The proposal reuses these facts for normalization_work_budget at V11:411, so this is not merely inaccurate commentary. The proposed checked arithmetic cannot be called equivalent to the constructor until all actual normalized lengths are derived. This review establishes a design mismatch, not an executed overflow or resource refusal from unimplemented code. Preserve the existing arithmetic and exact acceptance rather than adding a cutoff.

## Smells

### C1-M5, per-field rescan scheduling and work bound remain unclosed

/private/tmp/rdocx-deepseek-harness/fx272-design-rescue/DEEPSEEK_MECHANISM_V11.md:302 permits a per-simple-field capture scan, while :326 still says every close is O(1), no qualifier rescans a span, and all modeled nesting is bounded by the listed fixed guards. Pure fldSimple chains do not receive the ordered-owner cost unless the guard encounters the qualifying control or revision at crates/rdocx-oxml/src/document.rs:1451. Source-copy totals are computed at :1611 and returned at :1648. They are not a fixed nesting cutoff for that chain.

V11:688 requires capture accumulation over the field's own span without specifying whether that accumulation re-enters the current root's intrinsic close check, how nested checks are scheduled without recursive requalification, or how total repeated work fits the unchanged admitted operation envelope. The design must state that scheduling and a concrete work bound. A correct bound need not be linear if the existing contract permits it. The conflicting linear and constant-close claims must be reconciled. Do not substitute termination on smaller spans, the old constructor's repeated scans or a conservative new depth limit.

## Retained progress and limits

The source-derived caller modes, three-way outcomes, iterative guard-first direction, lexical per-site scope correction and simple-field independent versus parent-admitted distinction remain useful. The inline-span probe executed six cases against the pinned existing library with compiler and child exit 0. Its scope is constructor observation only. The frozen manifest authenticates its actual source, logs and pins. It is not an execution of a proposed qualifier or resource bound.

No additional defect is asserted for those retained sections in this targeted pass. No panics, schema or product-source finding is manufactured from unimplemented pseudocode. The prior C1-M1 and C1-M5 invariants remain open in the concrete respects above. The unchanged factor 124, declaration pooling, duplicate charges, unused-URI and zero-growth controls, exact fallback policy and narrow direct legacy-read exception remain mandatory. Source lease, coordinated canonical adoption, scoped product validation, native, schema, hash and final integrated qualification remain due.

Return to the author phase. Root final acceptance follows actual closure. STOP.
