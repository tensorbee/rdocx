# S96, Root final ordinary-frame offset repair, pass 15

## Scope and verdict

Reviewed the four-path F273 A3 increment `e80c833227cfeae73107ad486120d12253f27fa4ec4a023a7d0162f4a30d4067`, frozen boundary `f08922035cb71302a9932f16ebdc721be90e0d55143d9e37eb1a52d3ec2bf10b`, actual independent support review and the earlier Root technical design approval. This is final-rider remediation. Historical provider completion, feature states and review caps are unchanged.

Defects: 0. Smells: 0. Nitpicks: 0.

## Assessment

Root inspected the complete proposal and actual implementation, contract reconciliation, test assertions and independent review. Actual Word controls preserve the anchor bytes while removing or shifting only the enclosing frame. They establish frame-origin causation for the measured offsets. Native saved XML also retains the original offsets. These observations justify reconciling the historical framed 120-point assertion rather than changing the native comparison tolerance.

At `crates/rdocx-layout/src/paginator.rs:7097`, the ordinary selected furniture-frame emitter retains its existing prepared frame and paragraph origins. Each Page-offset axis receives the actual frame origin independently. Alignment and other relative axes retain their existing resolution. Global, unframed and true-cell helpers are unchanged. The matched source paint count is checked before indexing. Actual ordinal, typed paint ownership, behind/front ordering, errors, media and rotation remain with the existing emitter. No XML, model, API, parser, cache or resource allowance changes.

The actual predecessor passes frame grouping, body placement and exactly-once invariants before failing on page coordinates 80,120 against the new 140,300 expectation. The previous 120-point green assertion remains historical evidence. The successor passes 31 Issue424 tests and eight additional cell/page/layer/cache controls. The new matrix covers header/footer, moved and absent frames, independent alignment axes and true-cell Page offsets. Default scoped checking, strict lint, format, prose and diff checks passed. The independent review authenticated the frozen source, evidence, binary and native causal files.

The approved amendment and `docs/hld/08-rendering-spec.md:2066` now state the bounded distinction. It does not infer native Page-alignment or footer parity from the offset controls. The measured 0.0501 pt residual is not compensated with a fitted value.

## Pending

Fresh current-code rendering must compare the exact original, no-frame and shifted sources to their retained native PDFs, with unchanged 0.5 pt measured-edge bounds and exact paint counts. The Word-generated VML fallback SDK error stays separate. Archive, current artifact, final native, performance, memory, hosted and full-union gates remain pending where affected. This source review is neither a native pass nor sprint completion.
