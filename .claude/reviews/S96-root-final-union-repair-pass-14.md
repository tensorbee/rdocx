# S96, Root final replacement selector repair, pass 14

## Scope and verdict

Reviewed the frozen selector increment `04641efc14f8e65d24d268a9a38b5df43019d71fd754c4408ee17f33fcd2921f`, source boundary `209f06c12aa6c1c12914c37ce3f49c50683cdd2e2e437985bb60133761ffa6c8`, and actual independent support review `2eb8b1c88a41c9c10f5ade69927d4b0f59afeb7bdd93546e5b18e546979f7da8`. Previously reviewed cold, reply, picture and CI-test changes are separate increments.

Defects: 0. Smells: 0. Nitpicks: 0.

## Assessment

Root inspected the implementation, approved amendment, actual test assertions, execution boundary and independent report. The binding dispatch at `crates/rdocx-py/src/document.rs:17050` is confined to bound, supported, one-index selectors. It validates item and story ownership before Native work. Constructed, numeric, unknown and nested cases retain the original resolver.

The new support operation at `crates/rdocx/src/document.rs:47880` performs complete actual source discovery before selecting a story. It preserves unrelated-source errors and missing-story precedence. Replacement validation precedes item qualification. The selected fingerprint comes from the current source, not the stale Cell inventory hypothesis. The extracted span helper retains the original bounds and kind policy.

Only that same owned operation constructs the private checked target after qualification. No mutation or transfer to another candidate intervenes. The common pipeline at `crates/rdocx/src/document.rs:48120` still flushes and qualifies canonical source, physical ownership, namespace context, Fields and copy reservations. Public existing APIs, publication order, error atomicity and resource allowances remain unchanged. The new branch removes an initial duplicate read rather than bypassing the postflush checks.

The genuine predecessor completes successful body and header replacement, saved/reopened package equality and neighbor checks before failing with two preflush preparations. The successor requires one. Actual final evidence includes three Native controls with 42 fixture/action comparisons and successful replacement in seven related story kinds, 14 binding legacy/successor pairs, and five existing transaction controls. Assertions cover Fields, aliases, retained raw content, tables, controls, held formatting, stale text snapshots, malformed unrelated stories, replacement errors, no-match, mismatch, nested fallback and held-view/package atomicity. Strict lint, format, prose and diff checks passed. The test-only unsupported nested-call assumption and compilation failure remain separately recorded.

## Pending

The original flat-cost threshold requires a fresh ordinary artifact. Other performance, native, memory, installed artifact, hosted and final integrated gates remain pending. This bounded source review supplies no timing result or feature completion, and does not reset a peer-review cap. Archive metadata adoption is separate from this source review.
