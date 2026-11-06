# F-X272, note discovery caller closure, pass 1

**Reviewed**: The actual template note-discovery callers and private helper at sampled Native document SHA256 `c1288bd403d2133d2c4eba4a0cc1bad9ecef902881a4b3c2c6f523c4ee194647`. This bounded source audit supplements the separate numbering finding. The comments worker still owns the mutable source. No runtime observation or whole-provider verdict is claimed.

**Verdict**: 1 defect, 0 smells, 0 nitpicks.

## Defects

### D1, allocated-note discovery reparses outside the actual operation

`crates/rdocx/src/document.rs:47588`, `crates/rdocx/src/document.rs:47589`, `crates/rdocx/src/document.rs:47595`, `crates/rdocx/src/template.rs:2612`, `crates/rdocx/src/template.rs:2654`

Both template modes call `template_note_identities` after scoped program preparation, and again after a repeated-note allocation. The helper calls ordinary `CT_Footnotes::from_xml` to construct complete note paragraphs and Fields while discovering IDs. It discards the actual paired part Source and uses a fresh parser work context. Its earlier `self.stories()` also reaches ordinary main-part serialization through `story_sources_and_main_layout` and `main_part_with_body_spans`.

This is reached by admitted templates with a normal footnote or endnote, with the second path requiring a repeated reference. The approved plan's related-reader closure explicitly names Native allocated-note discovery as an actual paired-reader consumer at `.claude/plans/F-X272-design.md:6322`. It is not a direct legacy infallible helper invocation outside a checked operation.

Use the existing operation, actual main/related Source graph and substantive `source_notes` reader. Preserve complete story discovery before note validation, stable footnote/endnote and owner ordering, normal-note classification, missing-part errors, duplicate-ID/count ambiguity refusal and later writer handback. No new public API, input allowance, quota or general XML validator is required. This finding is source-derived, with no invented historical runtime RED.

## Smells

Zero smells.

## Nitpicks

Zero nitpicks.

## Not found

No separate note classification, package corruption or story ordering defect is asserted. The related source readers elsewhere already have the necessary callable contract. Deeper deferred security assessment remains deferred. Comments, final provider microscope, scoped gate, canonical reconciliation and final sprint riders remain pending. Only this review report was written. Remediation begins in a separate source phase.
