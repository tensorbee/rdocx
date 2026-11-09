# S96 styles diagnostic expectation, Root final pass 1

Reviewed the combined private v6 diagnostic patch
`5486a21b6601aa8c6f3204c5124e5c477e506ee915a1a308ce5aa164cd3aadaf`.
Root authenticated the successor packet and read the exact assertion changes,
actual failure analysis, source-preserving defaults writer and styles publisher.
No product edit, build or test occurred during this review.

Verdict: zero defects, zero smells and zero nitpicks in this repair proposal.

The actual Normal styles case failed at archived document.rs:50890. Complete
captured Debug analysis isolates one difference: the original docDefaults source
snapshot in the live model versus its newly written snapshot after reopen.
Both models already have the requested four typed font values. The actual
defaults writer deliberately merges changed typed properties into the retained
original wrapper. Whole-model equality therefore incorrectly includes source
history in this diagnostic's semantic expectation. The omitted case remains
unrun, and this diagnosis does not establish its parity.

The replacement checks public default-font values on both Documents, complete
live/fresh stylesheet writer bytes and actual live/reopened saved part bytes.
The saved-part expectation uses the same existing root-attribute preservation
producer as flush. It preserves all ordinary modeled and opaque output rather
than stripping source fields or changing product equality. Other error/refusal,
cache, All and29field mandatory-inventory assertions remain unchanged.
The previously reviewed lexical xmlns repair is the only other change.

No additional correctness, structural, namespace or evidence defect was found.
Both historical failures remain failed. Corrected compilation and all changed
Normal/omitted root/styles cases remain pending. Root may install this reviewed
private successor and batch one relink. No final SourceH, shipping, performance,
binding, whole-C1 or provider acceptance follows from this preparation.
