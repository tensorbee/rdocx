# S96, Root final shared reply-parent repair, pass 11

## Scope and verdict

Reviewed the reply-only increment authenticated by handoff `96908f0f6b57c742eb4fcfe1b9a7a3708c6ba9a362a8a3bd640a5131391c4a35` and patch `af4c2ee4ffba1cf1464882c907c1c6fc3ae8793469f4c9a646302290aa49d42c`. The independent support review is `.claude/reviews/S96-reply-shared-parent-support-1.md`, SHA `d7c84b18b9d85cb424eff35fec1ab1295852985c27e39b1fca2b10b32f538765`. This is a bounded final repair review, not final sprint acceptance or a review-cap reset.

Defects: 0. Smells: 0. Nitpicks: 0.

## Assessment

`crates/rdocx/src/comments.rs:2026` prepares the actual private candidate, shares its complete parent read between projection and reply-reference qualification, and drops the read before identifier reservation. The five-field reader at `crates/rdocx/src/comments.rs:3067` owns real source facts. It preserves complete graph and paragraph validation, physical ordering, original XML and accepted paragraph projection. The extracted projection at `crates/rdocx/src/comments.rs:3140` retains its construction-backed flow indexing. The reference checks at `crates/rdocx/src/comments.rs:3228` retain exact cardinality, same story and part, ordering, endpoint and metadata-only payload requirements. Public snapshots at `crates/rdocx/src/comments.rs:3524` retain staging and flushing. The outer transaction, later anchoring and public APIs remain intact.

Root read the implementation, contract amendment, HLD description, independent review, actual receipt and final test assertions. The predecessor mechanism test reached successful reply and reopen before failing with two parent factory passes. The successor records four passing new controls, including actual complex rich Field content, current unsaved content, tracked content, aliases, header ownership, malformed unrelated content, exact error precedence and held package/style/identifier atomicity. The broader 39 passing comments tests and one existing ignored native oracle retain their earlier test-source boundary on unchanged production logic. Final lint, format, prose and diff checks are actual exit0. Test compilation, lint and fixture-assumption failures remain preserved separately.

The measured counter covers the parent all-paragraph factory phase only. It does not claim that projection probes or later anchor parsing disappear. The separately reviewed cold-owner batch remains outside this increment. No namespace, resource limit, Source ingress, public signature or persistent cache was added.

## Pending

The original reply flat-cost threshold and the other four failing flat gates still require actual ordinary-artifact measurements. Cold407 acceptance, native qualification, memory checks, final integrated verification and final sprint review remain pending. This review supplies no timing result, no native pass and no release permission.
