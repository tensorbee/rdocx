# Current Sprint, S02

**Milestone**: M1 Preparation and safety net.

**Goal**: Put every remaining prerequisite for extraction in place while the
current rdocx behaviour is still stable. Resolve the carried packaging defect,
prepare the Rust and Python-facing APIs, pin layout and unit behaviour, reserve
the future crate names, then publish and tag v0.4.1 as the known-good state
immediately before structural churn begins.

## Spec references

- `docs/hld/04-opc-and-packaging.md`, for relationship-based core-properties
  lookup and the package invariants F-007 must preserve.
- `docs/hld/10-bindings-spec.md`, for the non-consuming setter surface F-008
  adds so Rust builders can back Python properties.
- `docs/hld/08-rendering-spec.md`, for the cached `LayoutResult`, mutation
  invalidation, and one-layout-per-document requirement in F-009.
- `docs/hld/11-migration-plan.md`, for pinning truncation before extraction and
  separating behaviour preservation from structural moves.
- `docs/hld/15-build-and-toolchain.md`, for crate-name reservation, publishing
  order, packaging verification, and the pre-churn release process.
- `docs/hld/12-testing-strategy.md`, for the workspace, hash-harness,
  packaging, and supply-chain gates that the v0.4.1 tag must pass.

## The wave

| F-ID | Title | Size | Status | Owner |
|------|-------|------|--------|-------|
| F-007 | Resolve core properties through the rel | S | done | - |
| F-008 | Non-consuming setter twins | M | done | - |
| F-009 | Cache the layout result | M | done | - |
| F-010 | Reserve crate names | S | done | - |
| F-011 | Pin unit truncation behaviour | S | done | - |
| F-012 | Tag v0.4.1 | S | done | - |

## Sequencing note

Rows are listed in dependency order, not F-ID order.

F-007, F-008, F-009, F-010 and F-011 are independent and may proceed in
parallel when their touched files and external publishing actions do not
overlap. F-012 is last. It depends on F-003 through F-011 and serves as the
known-good boundary after every S02 prerequisite is integrated and verified.
After S02 closes into `main`, merge that updated `main` into
`feature/release-0.5.0` before the next release branch continues.

## Definition of done for this sprint

- Core properties resolve through their package relationship at a non-standard
  part path and round-trip with metadata intact.
- The consuming builders delegate to non-consuming setter twins, with the
  setter test gate green.
- Rendering every page of a 20-page document performs exactly one layout and
  every mutation invalidates the cache.
- Every reserved `oxml-*` and `rpptx*` crate name resolves through `cargo info`.
- The reserved `oxml-*` and `rpptx*` crates remain at 0.0.0 until PowerPoint
  development is complete.
- Unit constructors retain their current `as i64` truncation, pinned by tests.
- The full workspace, hash harness, packaging, and supply-chain gates pass from
  a clean clone, the baseline reproduces on a second machine, and v0.4.1 is
  published and tagged.
