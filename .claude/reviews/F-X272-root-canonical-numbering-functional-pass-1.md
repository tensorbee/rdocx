# F-X272, canonical numbering integration, pass 1

**Reviewed**: The frozen joined R2 main-source publisher and setter, F277b same-Document numbering contract, and the actual callback-only post-render ID observation. The separate distinct-start diagnostic has a corrected premise classification below. No post-render start change is established by its initial run.

**Verdict**: 1 defect, 0 smells, 0 nitpicks.

## Defects

### D1, published body numbering is reparsed as live numbering

`crates/rdocx/src/document.rs:16125`, `crates/rdocx/src/document.rs:16183`, `crates/rdocx/src/document.rs:16231`, `crates/rdocx/src/document.rs:48329`

The scoped template publisher flushes a canonical package projection, then its main-source setter reparses that body's published numPr values into the live document. The numbering model was restored to live IDs after flush. Subsequent publication can now interpret an already-published body ID as a live ID, losing the original association before a later inverse reopen.

The callback-only fixture separately observed a paragraph selecting instance1 after scalar rendering instead of its original selected instance2. F277b's approved contract at `.claude/plans/F-X277b-design.md:353` and `:554` preserves same-Document live numbering handles and restored body bindings. Template rendering finishes through same-document staged reopen and commit. It is not process reopen or an authorized newly imported-list allocation. Therefore that original-association assertion cannot be dismissed as an unsupported new blanket ID policy. The fixture's two instances share their definition, so this observation alone does not demonstrate changed visible numbering.

Preserve the actual producer's ID space through the existing private scoped handoff. Do not restore eager numbering mutation or blindly invert every source read. Existing canonical_source/import producers can emit live IDs outside the publication window. The proposed repair retains a genuine live-model main Source separately from the canonical published package Source. Runtime and final review of that repair remain pending.

## Diagnostic correction and evidence limits

The first draft of this report incorrectly adopted semantic-numbering/RESULT.json's manually summarized rendered-start values as actual observations. Root authenticated its binary and restored source but did not inspect the assertion phase before making that claim. That summary is withdrawn.

Direct inspection of diagnostic-document.rs shows its first assertion calls `numbering_level(selected, 0).start` before rendering. That reader returns the abstract base start rather than the instance override. The raw failure at line75030 therefore proves an invalid fixture premise, not a rendered7to1 result. No rendering occurred in that run. The original incorrect RESULT.json and raw log remain preserved, with an additive PREMISE_CORRECTION.json explicitly identifying the error. No production reader policy changes.

The failing diagnostic binary hash remains `949950542f4827913011ae34cccba676759d7d7f821ebc0058c5dae27e6803dc`. Its authentication is not semantic qualification. A corrected discriminator inspecting the actual instance override is pending. Existing post-render ID-binding evidence and the source/contract finding above remain separate and must not be conflated with that unexecuted semantic step.

## Smells

Zero smells.

## Nitpicks

Zero nitpicks.

## Not found

The original F277b publication-only remap and restoration window remain correct. This finding concerns subsequent scoped body handback. No unrelated list policy, new public API, security investigation or release waiver is required. Provider completion awaits actual repair, focused checks and review. This report changes only review documentation and ends before remediation.

## Corrected post-render observation

The successor CORRECTED_RESULT.json, corrected-diagnostic-document-2.rs and
corrected-run-2.log now establish the later phase. Root read the actual assertion
sequence and authenticated corrected-native.bin SHA256
`8d023647e3463566898b8b62a68162f3b1daf084a04d6e5a45767b32a0279b97`.
The before-render instance2 override assertion Some7 passes. Rendering
completes. The actual paragraph then references instance1 with no start override,
and its post-render Some7 assertion fails. This confirms loss of the selected
override in the callback-only predecessor. It is not a native/rendered-glyph
observation or a relabeling of the initial invalid diagnostic. The corrected
fixture uses actual NumberingInstance override data. A subsequent package
reopen and final repaired-source acceptance remain required local controls.
