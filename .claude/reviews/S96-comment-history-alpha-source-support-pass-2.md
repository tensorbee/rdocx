# S96 comment-history oracle remediation review

Verdict: 0 defects, 0 smells, 0 nitpicks. Previous bounded D1 is closed by source inspection. No source changes or Cargo were performed.

At crates/rdocx/tests/regression_test.rs:4646-4652 each build captures its actual shared original comment definitions before either image history. At :4686-4704 each resulting owner is selected by its original comment id and its complete original paragraph_ids prefix is asserted equal, including absent identities. This applies independently to both build(true) and build(false). Thus the later alpha map cannot waive preservation of any shared preexisting paragraph identity.

At :4613-4616 the actual physical primary ID map must be unique. Normalization remains restricted to the three named primary/extended attributes in the two comment parts. The existing per-history repeated serialization, exact save/reopen, image order and raw producer assertions remain. Fresh allocated IDs can differ across histories while their parent references must map to the same physical owner. No new history-independent numeric allocation requirement is introduced.

This closes only the bounded test-oracle finding. Actual regression execution and the remaining sprint gates are separate.

Reviewed regression source SHA256: `ca9aa87450ff5fa5c9b550b5846a8ead3d472fed927dea3f04f93daeb4200693`.
