# F-X277 current-source preparation, Codex targeted final 4

The stopped GLM preparation is useful but not accepted for automatic adoption.
This is a Root targeted review of new preparation material, not a fourth design
peer, formal design approval or provider completion. All existing caps remain.

## Defect

D1. Two claimed current source anchors identify different constructs. The report
calls text.rs:9820-9848 the CT_P payload, but that range is ComplexFieldProjection.
CT_P is at crates/rdocx-oxml/src/text.rs:11417. It also identifies
text.rs:15595 as a run emitter, but that is the paragraph content emitter.
These errors appear in /private/tmp/rdocx-glm-f277-revalidation/F277_REVALIDATION.md:38
and :25, and in PATCH.txt's corresponding recommendations. Its assertion that
every citation was verified overstates this evidence (report:84). Re-pin by
looking up actual definitions and reading the relevant body, without inferred
line-offset arithmetic. Validate all proposed changed anchors before adoption.
Root source changes in this pass cannot explain the CT_P discrepancy.

## Smells and nitpicks

No additional smells or nitpicks established in this targeted inspection.
Five local body-level checks are openly pending in the report. Complete those
where current source is available, while retaining unavailable donor-tree
inspection as pending. Uncertainty disclosed as pending is not a fabricated
product defect. The report does not authorize F277 implementation.

## Evidence and stop

Actual child0 and1051.119seconds are recorded by the GLM wrapper. Only the two
allowed temporary outputs were produced. All eight A277 rows and the approved
normal-note policy remain in its report, with the actual F272 dependency intact.
No source, plan or test edit is adopted by this review. A focused author repair
follows separately. This review changed only this record and ends here.
