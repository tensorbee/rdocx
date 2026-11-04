# F-X280c, GLM self-review, all aspects, pass 2

**Erratum (remediation phase, appended after Root final1)**

Two claims in this review exceeded the actual observations. They are
corrected here without altering the original text above:

1. The full native carrier command reviewed was NOT a passing command. Its
   retained child result is 418 passed, 2 failed, 25 ignored, child exit 101
   (.claude/scratch/logs/native-remediation-full.log). The two failures,
   add_chart_rejects_invalid_data_without_mutation and
   supported_and_fallback_charts_render_deterministically, are reproduced at
   the clean Base (chart-base3.log) and belong to a separate F281
   continuation. They were never fixed in this story. Claims naming the
   broad native pass as green were wrong.
2. The saturation claim was wrong: with u64 saturation an overflowing
   gridSpan value clamps coverage to u64::MAX, so the shortfall comparison
   `coverage < columns` becomes false and the row is NOT reported. No
   executed huge-span regression existed, and none was invented afterward.

Additionally, the D1 disposition supersedes the gridSpan coverage policy
this review examined: the approved direct qualified physical cell-count
equality is restored, the classified pinned-corpus case is asserted
explicitly with its exact diagnostic while the parsed-model agreement is
preserved, and the two chart failures stay with the F281 continuation.

---

**Reviewer**: GLM-5.3-Flash-EXL3 (dgx/GLM-5.3-Flash-EXL3#high), implementing
author self-review after the pass 1 remediation. This is explicitly not the
independent review. Root supplies that separately.

**Reviewed**: the pass 1 remediation delta (saturating coverage sum, the
stored_required_cnvpr_issues split plus tree-root spTree wiring, the new
tree-root discriminator test, the corrected CLI preservation comparison) and
its effect on the whole F-X280c diff. Product diff unchanged elsewhere from
pass 1.

**Verdict**: 0 defects, 0 smells, 2 nitpicks (carried from pass 1)

## Defects

None. Both pass 1 defects are addressed with evidence:

- D1 `crates/rpptx/src/lib.rs:3729`: the coverage sum now saturates. The
  verdict is unchanged for every legitimate value because the comparison
  target is the grid column count, and an overflowing span now reports a
  short-row issue instead of panicking.
- D2 `crates/rpptx/src/lib.rs:3521` and `crates/rpptx/src/lib.rs:3598`: the
  cNvPr requirement is factored into stored_required_cnvpr_issues with two
  existing callers (physical shape members and the tree-root spTree), and
  issue435_stored_tree_root_cnvpr_requires_id_and_name discriminates the
  repaired invariant on a structurally different input (the tree root, not a
  member shape). The full rpptx carrier and the pinned-corpus CLI carrier
  pass after the change (native-remediation-full.log, cli-remediation.log),
  so the new check produces no false positive on admitted decks.

## Smells

None. The pass 1 smell is fixed: the CLI fault loop now captures the fixture
bytes before the CLI runs and compares against them
(`crates/rpptx-cli/tests/integration.rs:4519`), so the assertion is real
preservation evidence.

## Nitpicks

- `crates/rpptx/src/lib.rs:3684`, a DOCTYPE is still not reported because
  the declared shared lexical frontier does not classify it. Carried, out of
  the approved A4 scope.
- `crates/rpptx/src/lib.rs:3639`, text-outside-root still precedes the root
  count fault. Carried, deterministic.

## Not found

- correctness: the remediated paths re-traced. No regression in the original
  six fault detections, the kinsoku/kiosk discriminator, choice members,
  precedence or dedup.
- panics: the only untrusted arithmetic is now saturating. No new unwrap,
  expect or indexing on untrusted input.
- ooxml: the tree-root check uses the same qualified nvGrpSpPr/cNvPr
  identity and unqualified attribute rules as the member path. Original
  bytes remain untouched.
- structure: the split helper has two real callers and adds no indirection
  beyond one hop.
- tests: the new discriminator fails the pre-remediation behavior by
  construction (the tree-root cNvPr was uninspected) and passes now. Both
  carriers and the pinned corpus are green. The stale Python extension was
  rebuilt (py-build3, sha256 229ccb66a8373876f662b9b4ec71a7216e7d95ff7fd4ec69a234a109d6acc05f)
  and the focused Python case passes (py-validate4).
