# S96 current bookmark projection bounded repair review

Verdict: 0 defects, 0 smells, 0 nitpicks in the adopted successor repair. This is a bounded independent repair review of f280's provider. It does not reset prior feature review caps or complete final acceptance. No source changes or Cargo were performed.

The reviewed immutable proposal is current-bookmark-projection-candidate-20261107-successor-1/PROVIDER.patch, SHA256d56de1fdbce5aaf8be516851df24ab40b9d0a8631a55d1561f216e3c80f0f527. The explicit Root amendment in .claude/plans/F-X272-design.md:6695-6728 authorizes this concrete doc-hidden accessor and bounded marker classifier while retaining raw source, all identities and the legacy infallible-read exception.

## Provider and identities

At crates/rdocx-oxml/src/text.rs:18555 the read accessor returns a fresh vector. At :28949-28970 assignment updates only projected and tracked indexes on a clone of the original marker occurrences. The assigned bitmap consumes each retained occurrence at most once. Identity, name, physical run slot, raw_before, namespace scope, child-content and unmatched or duplicate evidence remain in the original order. A failed raw match leaves the original evidence present rather than dropping it. No source mutation, Field construction, paragraph parse, new input or output allowance is introduced.

At :28993-29172 the traversal follows current typed controls and revisions at actual direct boundaries, using the existing AcceptedOwnerOrder and raw slots. It accounts for raw-position and hyperlink revision ordinals separately, retaining physical ordering when several owners share a direct boundary. The raw classifier at :28862-28903 recognizes only the admitted qualified marker child with its actual retained scope and attributes. It does not discover runs or Fields in opaque children.

At :29176-29204 accepted insertion/moveTo counts derive from the current typed paragraph. Deleted/moveFrom owners contribute only tracked width and do not expose deleted marker identities in the accepted set. The unmaterialized legacy fallback mirrors append_tracked_revision_runs: non-Runs content contributes zero, and nested wrappers count only at real boundaries through runs.len(). At :29206-29309 controls, typed table/row/cell collections and their boundary controls recurse in the existing physical order. There is no persistent projection cache or new state/type.

## Consumers

Native bookmarks at crates/rdocx/src/comments.rs:1184 and TOC inspection at crates/rdocx/src/field.rs:10364 consume the new current indexes. Layout source registration, boundary collection, range targeting and bookmark targets at crates/rdocx-layout/src/engine.rs:2100, :11484, :13151 and :13785 use the same projection. The comparison boundary reader at crates/rdocx/src/comparison.rs:16715 uses current accepted/tracked indexes without rewriting stored physical metadata. Name-only and physical mutation/custody reads remain on their existing stored metadata. Pairing and duplicate validation remain separate.

## Actual validation limits

Root reports actual all-target/all-feature rdocx check0 and13 lower controls0 on stable source, including four selected bookmark cases. The fifth legacy-revision fallback test and affected Native TOC/bookmark consumers were pending the next batch when this source review was requested. These source checks do not manufacture those runtime outcomes. Full integrated suites, current wheel/runtime, native and performance riders remain separate. Existing raw/save and malformed marker controls remain required, without changed assertions.

Reviewed adopted-source SHA256 pins:

```json
{
  "crates/rdocx-oxml/src/text.rs": "6873a90b91a5935d3d2cb746a6587097ae1d029f6277b410c57be56637d65878",
  "crates/rdocx/src/comments.rs": "19dff2dae6b6152c98cb2f9cbe6e37e2cdb96483404f3cc6d503c6dafd983689",
  "crates/rdocx/src/comparison.rs": "1bb2e316c7f658e48c9b773697405b4b5b12df05151fe9ae78b6e9e58bede694",
  "crates/rdocx/src/field.rs": "5b1dbefd42073b81ea231a8708626d8f47f1760cf1d0ec042638acb9eb849ba1",
  "crates/rdocx-layout/src/engine.rs": "d42eb0186cb92e05fb42ede2d56817c595aafb8d35eabad9b9fd842e18394932",
  "docs/hld/04-opc-and-packaging.md": "5da5527fb122dc7f64c656de0945c203ae748ab6782eb7dafb42e4433afb5ae9"
}
```
