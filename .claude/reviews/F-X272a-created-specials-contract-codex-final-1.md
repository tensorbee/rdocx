# F-X272a created specials contract, Root affected final pass 1

Reviewed child plan SHA256
`73cb4e592e341518b04436a6835cdbf30ee127f3f3fb54a33988ab689bd25ccf`
at claim commitbda1042c, with original line citations below.
Reviewed the actual creation-path contradiction in the approved child contract
and existing clean-Base producers. This is bounded Root contract adjudication,
not a new peer allowance. Verdict: one defect, zero smells, zero nitpicks.

## Defects

### D1, unconditional specials synthesis contradicts exact destination inventory

`.claude/plans/F-X272a-design.md:684` requires conventional specials on every
comparison-created part. The creation row at303 and checklist at578 also
assume synthesis. However, the exact destination requirement at344 requires
acceptance to preserve the edited input special inventory. The existing
`crates/rdocx/src/document.rs:33883` add_footnote producer appends only a normal
note before flushing. It does not add separator definitions. F-X277b owns the
later authored defaults and is not an implemented prerequisite of this child.

For absent original versus a present edited part containing a referenced normal
note and no specials, unconditional synthesis publishes extra definitions on
accept. This violates the exact destination contract. A newly invented blanket
refusal, dependence on future authored defaults or unrecorded normalization
would not repair the contradiction. The same unconditional fixed ids can also
conflict with the explicitly supported normal owners at ids0andnegative1.

Remediation must retain the actual destination inventory. Comparison creation
copies the edited input specials when present and does not manufacture specials
when absent. Existing examples whose edited input already contains authored
specials keep those exact records and the existing reject-to-absent cleanup.
Both-present imported-special preservation and existing refusal policy remain.
F-X277b retains authored defaults. No extra persisted synthesis-provenance field
or public API is required. No code or source gate is approved by this review.

## Other aspects

No additional correctness, panic, XML placement, structure or test-contract
finding was established in this bounded read. Existing native/schema/final
riders remain pending and the exact Q-A430-28 exception is unchanged. Root
modified only this review record and returns to separate contract remediation.
