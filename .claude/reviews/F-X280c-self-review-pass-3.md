# F-X280c, GLM self-review, all aspects, pass 3

**Reviewer**: GLM-5.3-Flash-EXL3 (dgx/GLM-5.3-Flash-EXL3#high), implementing
author self-review of the Root final1 remediation. This is explicitly not the
independent review. Root's final2 supplies that.

**Reviewed**: the Root final1 repair delta against the frozen final1
candidate (SOURCE_FREEZE_1.json, diff a1924899...): the iterative
StoredFrame scanner replacing recursive materialization, the
stored_role_roots ownership qualification, the restored direct physical
cell-count equality with the classified corpus case, the position-sensitive
root fault, the identity-based metadata recovery, the S1 erratum and help
text, plus the new discriminating tests. 10 product paths plus both
carriers.

**Verdict**: 0 defects, 0 smells, 3 nitpicks

## Defects

None found. Each Root finding was traced through its repair:

- D1: the row rule is a plain count of direct qualified `a:tc` children
  against that table's own direct `a:tblGrid` gridCol children, reported in
  both directions with the original message text. No span field survives in
  the scanner state. The classified corpus case asserts the exact stored
  diagnostic and the unchanged parsed-model agreement, and the other 49
  pinned decks stay clean.
- D2: structural frames are only created after stored_role_roots qualifies
  the part through an owning role relationship, the role's expected content
  type and the expected qualified root, and the main part through the
  officeDocument relationship plus a presentation package class. The unused
  lookalike controls receive lexical checks only, and the paired owned
  controls still report. Exact target-to-part matching matches the model's
  own required_part behavior.
- D3: one iterative read loop with a flat heap `Vec<StoredFrame>`. No
  function recursion, no owned nested tree, no recursive Drop anywhere in
  the stored path. Opaque islands push a single frame kind that resolves no
  children. The deep fixtures run through the CLI child process, which now
  validates clean at 60000 levels, and the fx214 model-depth controls pass.
  No new quota and no previously admitted payload is refused.
- D4: outside-root CDATA, general references, post-root declarations and
  post-root document type declarations are faulted with distinct messages.
  in-root CDATA and references and outside-root whitespace, comments and
  processing instructions stay clean, over ordinary parts and original
  metadata alike.
- D5: content-types recovery matches the root identity case-insensitively
  under the loader's normalization guards, so binary and nested-XML suffix
  decoys stay under their MIME frontiers while genuine metadata (including
  a case variant holding a forbidden character) is still checked. The
  `.rels` inventory stays aligned with the loader's own relationship set.
- S1: the erratum is appended to the pass 2 review file preserving the
  original text, the false green and saturation claims are corrected, no
  overflow test was invented, and the user-facing help no longer names the
  internal story while keeping the subset and disclaimer.

## Smells

None found.

## Nitpicks

- `crates/rpptx/src/lib.rs:3928`, a root content-types entry spelled with
  dot segments, which the loader would normalize to the root identity, is
  excluded from metadata recovery by the `/.` guard and therefore only
  receives ordinary-part treatment or none. Producers writing such names
  are pathological and name policy stays with B, but the guard is stricter
  than the loader's normalization.
- `crates/rpptx/src/lib.rs:3221`, a part targeted by two different role
  relationship types resolves to whichever role was collected last. The
  finite scan then follows that role only. No plausible deck does this and
  both candidate scans are bounded subsets, but the map is last-writer.
- `crates/rpptx/tests/integration.rs:38193`, the tree-root missing-name
  discriminator asserts the path names nvGrpSpPr but not that the member
  shapes in the same deck stay clean. The six-fault clean deck and the
  continuation control cover that side.

## Not found

- correctness: frame dispatch re-traced against every preserved fixture
  including exact issue paths and ordinals (bogus[3], sldIdLst[3], tr[4],
  cNvPr routes), Empty-element push/pop finalization, self-closing roots,
  unbalanced and stray-end guards, and the Opaque-unreachable invariant.
- panics: no new unwrap, expect, indexing or arithmetic on untrusted input.
  the gridSpan value is no longer parsed at all. Counts are usize element
  tallies.
- contract: no new file, module, crate, dependency, public API, quota or
  feature flag. Python surface unchanged. HLD claims now match the
  implemented equality rule. The corpus classification is explicit rather
  than a weakened predicate.
- ooxml: namespace resolution still reuses PptXmlScope and PptXmlNames.
  attribute identity stays unqualified. Original bytes are never written.
- structure: StoredScan is a real context bundle with three used fields,
  not a forwarding wrapper. Every helper has its caller set.
- tests: every new case failed against the pre-repair candidate for the
  discriminating reason (red1-native, red2-native, red3-cli, red4-cli) and
  passes now. Known chart failures were not rediscovered and stay with the
  F281 continuation.
