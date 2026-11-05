# S96 text-history build custody, Root focused final pass 1

New sibling-consumer applicability review after the concrete COW D1 repair.
This is not a confirmation of the earlier clean liveness repair or another
history-source peer. One evidence defect, zero smells and zero nitpicks.

## CUSTODY-1, the text path also lacks the full build-manifest join

`text-history-consumer-v3-20261105/consume.py:284` compares Git SourceH while
line290checks the child manifest independently. Neither links that manifest
to the complete build inventory. The sibling execution preflight at
`text-history-controller-v2-20261105/run_text_child.py:121` has the same gap.
Native document and Git identity do not cover all private build inputs.

Apply the same finite association already repaired for COW: actual build
source_before_sha256must equal execution_identity.source_manifest_sha256.
That field exists in the actual build producer. Preserve the exact inventory
bytes across build and child. Add missing, wrong and re-pinned non-Native
source-inventory discriminators for both preflight and recorded consumption.
Keep existing text semantics, liveness repairs, protocol, source candidate and
all prior controls. This is a two-site evidence fix, not a new product gate,
public API or framework. No product or Native process ran during this review.
