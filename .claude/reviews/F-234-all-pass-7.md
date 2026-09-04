# F-234, all aspects, pass 7

**Reviewed**: complete working-tree diff across 3 implementation and test files, 2,940 inserted lines and 120 deleted lines
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, related-story comparison rewrites unowned inner bytes

`crates/rdocx/src/comparison.rs:695`

For headers, footers, and each comment or note owner, `compare_story_inner`
parses the complete inner source into a synthetic typed document and returns a
fresh serialization of that document. The typed body parser does not retain
inter-element whitespace or processing-instruction events, so a supported edit
beside either unowned construct silently drops or normalizes those bytes. This
violates the approved source-span boundary, which requires every unowned byte
and processing instruction to remain intact. The round-trip regression at
`crates/rdocx/tests/regression_test.rs:11150` checks only that one raw element
substring remains and does not gate whitespace, processing instructions, or
byte identity of the unowned story regions.

## Smells

None.

## Nitpicks

None.

## Not found

All three pass-6 findings are remediated. The legacy main-body formatting tests
now assert tracked modeled properties and diagnostic-only cell formatting. New
paragraph and table property owners diagnose edited unmodelled children,
exclude those children from tracked and accepted output, and reject back to the
original absent owner. Public `accept_all` and `reject_all` documentation now
states their related-story scope.

`cargo test -p rdocx --test regression_test` passes with 261 tests passed and 3
ignored. The focused differential, provenance, raw-property, package-wide
resolution, and accepted and rejected view coverage is green. `cargo check -p
rdocx --all-targets`, formatting, and diff checks pass. No additional
correctness, contract, panic, OOXML ordering or namespace, test, or structural
findings were found.
