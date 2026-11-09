# F-X272, provider caller closure, pass 1

**Reviewed**: Bounded Native template caller audit at sampled `crates/rdocx/src/template.rs` SHA256 `a7b265ff99397e4e134f399080d310126f6696c20efab9b141f1b4f6d5971e29`. The comments worker owns the mutable recovery tree. This is a concrete caller finding, not a frozen whole-provider verdict or a new full peer cycle.

**Verdict**: 1 defect, 0 smells, 0 nitpicks.

## Defects

### D1, repeated-paragraph numbering validation leaves the scoped operation

`crates/rdocx/src/template.rs:1388`, `crates/rdocx/src/template.rs:1391`, `crates/rdocx/src/template.rs:2568`

`render_grammar` first parses the actual program with its scoped part Source, then unconditionally calls `validate_program_numbering`. Every paragraph under a loop is closed and reparsed by ordinary `CT_P::from_xml_fragment` in that validation. The latter constructs the whole paragraph and any Fields with an independent parser work context, although the checked template operation already owns the actual source. The same route is reached for an empty loop before evaluation. It is not the approved exception for direct legacy infallible read helpers on caller-assembled models.

An admitted program containing a for marker, a paragraph with a scalar or numbered paragraph, and an endfor marker reaches this path through both public template modes. Preserve the existing numbering validation and refusal behavior, but carry the actual paired parent Source and operation through this recursive validation. Existing composition, namespace closure and paragraph-reader methods suffice. No new public API, input allowance, resource quota or conservative refusal is warranted. This is source-derived evidence, with no new runtime RED claim.

## Smells

Zero smells.

## Nitpicks

Zero nitpicks.

## Not found

No separate numbering-policy defect is asserted. The finding concerns the actual caller handoff and checked-operation context. Deferred deeper adversarial security assessments remain deferred. Comments carrier review, complete provider microscope, scoped gate, canonical reconciliation and final sprint riders remain pending. This review modifies only this report and ends before remediation.
