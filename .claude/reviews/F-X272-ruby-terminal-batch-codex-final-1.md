# F-X272, Ruby terminal family, Root final pass 1

Reviewed two production files and progress,356patch lines, frozen patch
`e1e175a7e2c93b59ca5673cee3aee80767179fddb8324d7c19a48b84f79212b8`.
The canonical contract remains
`f7ea5be35e25827971fe96243ef32f9b82621843e45ffb6501dcd62967654db8`.
This bounded Root review does not reset the exhausted implementation peer cap.

Verdict: one defect, zero smells and zero nitpicks. Source review only. No
product edit, test or build occurred during this pass.

## Defects

### D1, Started property decoding precedes its actual capture boundary

`crates/rdocx-oxml/src/document.rs:2644`

The shared Started/Empty source branch immediately calls RubyPropertyChild::read
at the opening event. Its language and scalar decoders can allocate and fail
before the Started child's subtree has been traversed. The actual property
reader captures that entire child first, then calls its scalar decoder. The
canonical plan's Ruby prefix contract explicitly requires this ordering and
forbids decoding a Started scalar before successful complete capture.

For a Started lid with a large value and a later malformed captured descendant,
the source qualifier materializes a language String that the actual constructor
never reaches. A malformed capture paired with a malformed scalar also does not
exercise the claimed capture-before-scalar ordering. The new late-None controls
cover successful captures and do not discriminate this boundary. Keeping broader
decoder work open does not justify introducing a new earlier decoder call.

Retain the actual child opening coordinates/context and apply the shared scalar
outcome only after its real capture boundary is established. Keep reached raw
capture work and failure precedence, without a second sizing parser, generated
input credit or changed acceptance quota. Add a discriminating malformed-capture
control and retain the existing successful-capture late-None positives.

## Not found

No additional correctness, namespace, child-order, public compatibility or
structural defect was established. Shared Run/Reject selection has actual
readers. The multi-run controls are correctly labelled existing parity rather
than historical failure. Successful raw completion and whole-Ruby success remain
explicitly unproved. The recorded241case cohort and scoped checks do not supply
whole-C1, provider, native or final integrated acceptance.

The review ends here and returns to orchestration for separate remediation.
