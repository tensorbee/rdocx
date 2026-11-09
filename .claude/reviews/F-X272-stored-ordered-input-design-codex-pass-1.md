# F-X272 stored ordered input visibility, Codex design pass 1

**Scope**: Only `Borrow stored ordered Field input without materialization` at
`.claude/plans/F-X272-design.md:6056`. Source citations below refer to the
sole-writer recovery worktree `/private/tmp/rdocx-s96-f272-codex-recovery`.
This is bounded design review, not implementation acceptance.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. The amendment is consistent with
the existing accessor and the approved input boundary.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Checked boundaries

- The signature and inert behavior match the actual existing implementation at
  `crates/rdocx-oxml/src/text.rs:6892`. It borrows the stored paragraph through
  `Option` without allocation, normalization, parsing or a work grant. The
  visibility-only authorization at `.claude/plans/F-X272-design.md:6058`
  preserves that body and signature. It introduces no new public type.
- The alternatives do not supply the same inert view. The ordered branch of
  `cached_result_runs` constructs display text at
  `crates/rdocx-oxml/src/text.rs:6903`, and
  `cached_result_physical_runs` delegates to it at
  `crates/rdocx-oxml/src/text.rs:6929`. The owned projection admits work and
  normalizes or clones at `crates/rdocx-oxml/src/text.rs:6247`. The concrete
  pre-clone input-freezing purpose at `.claude/plans/F-X272-design.md:6066`
  therefore justifies exposing the existing borrow independently of those
  selection and projection methods.
- Stored membership remains distinct from current writer eligibility at
  `.claude/plans/F-X272-design.md:6070`. Actual producer correspondence is
  required at `.claude/plans/F-X272-design.md:6074`, without equality or
  reconstructed ordinal authority. The explicit Sdt Run and absolute-index
  requirements at `.claude/plans/F-X272-design.md:6076` preserve a real visitor
  boundary: `visit_paragraph` only recurses through controls at
  `crates/rdocx/src/document.rs:21420`, while `visit_sdt` skips Run entries at
  `crates/rdocx/src/document.rs:21475`. This amendment does not claim that
  exposing one borrow completes Native membership or ordinary certification.
- The accessor is inventoried separately from the35operation carriers at
  `.claude/plans/F-X272-design.md:6082`, giving five separately counted borrowed
  accessors. Rust0.17.0 additive SemVer inventory, exact external lifetime and
  optional-result calls, existing regression carriers and HLD03/HLD10 impact
  remain explicit at `.claude/plans/F-X272-design.md:6078` and
  `.claude/plans/F-X272-design.md:6083`. No Python/WASM change, extra test
  binary, model field, trait, generic or production module is authorized.
- Public-API, shared-consumer and resource riders remain due at
  `.claude/plans/F-X272-design.md:6086`. The amendment neither grants provider
  completion nor waives implementation, independent review or final union
  validation. The sole writer must acknowledge it before the visibility edit
  at `.claude/plans/F-X272-design.md:6088`.

## Verification limits

Read-only source and design inspection only. No builds, tests or runtime proofs
were performed. Only this review record was written. Implementation and the
external compile client still require their actual checks on final source.
