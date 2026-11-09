# F-X219, Qualify full-range text watermark interoperability

**Status**: draft
**Sprint**: S98
**Size**: M
**Depends on**: F-X201

## Problem

`crates/rdocx/src/document.rs:27106` refuses every text watermark opacity except0.5 because previous native Word and LibreOffice endpoint controls disagreed with authored semantic opacity. The half-opacity representation and its strict source fingerprint are qualified, but deleting the guard would expose an unqualified native result. F-X201 is done in BACKLOG. No existing F-X219 plan, progress or review was found.

## Spec reference

- HLD14, F-X219, full zero-to-one real text opacity and retention of failed endpoint evidence.
- HLD03, Owned publication and causal paths, staged source and retained authority.
- HLD04, Checked Word operation source ownership, selected header and complete physical owner preservation.
- HLD08, Word watermarks and Authored Word page appearance and note references, actual selected paint, true text, alpha and rotation.
- HLD10, Python Word document operations and section-aware page authoring, unchanged public section API.
- HLD12, Word live-story and page-authoring controls, actual native and independent package evidence.
- HLD15, Deterministic rendering, stable bundled-font baselines.

## Approach

Keep the existing signatures and checked staging path. Investigate the actual producer representation rather than inventing opacity semantics. The candidate must encode genuine text fill at the requested alpha under truthful supported capability selection, retain styled WPS text and an equivalent genuine VML fallback, and project exactly one selected branch. Compare schema-defined solid fill and actual Word-produced source before selecting any minimal existing-carrier correction. Do not reverse opacity operands, flatten colour, rasterize text or hide broken source behind fabricated capabilities.

Any supported correction stays in existing `rdocx/src/document.rs`, existing Word run-properties/drawing projection and layout engine only if the chosen source earns it, and existing regression and Python test carriers. Preserve header variant selection, placement, dimensions, rotation, reserved identity, source fingerprints and strict checked replacement. Keep legacy VML and image paths unchanged. Full-range admission is an intentional labelled behavior change, not a cleanup.

Implementation may prepare source-qualified candidate packages and local regression controls before native acceptance. Final Word/LibreOffice producer experiments and endpoint qualification run once in the integrated sprint union. If no faithful representation qualifies, retain the named restriction and record the specific failed controls. Do not mark this contract completed or silently substitute a weaker acceptance decision. A materially different representation needs a plan revision before product implementation.

## Rejected alternatives

- Deleting only the guard leaves the demonstrated native endpoint risk unresolved.
- Complementing schema alpha, raster text or flattened RGB violates the requested source semantics.
- Repeating the half-opacity positive does not distinguish correct alpha from its complement.

## Test plan

**Test gate**: regression.

| Timing | Category | Test | Asserts |
|---|---|---|---|
| local scoped | regression | Existing text watermark regression carrier extended with `fx219_text_watermark_full_range_source_and_atomicity` | Actual Base refuses0/0.25/0.75/1, fixed source admits only the faithful selected representation, reads identical requested alpha after save/reopen |
| local scoped | regression | Source ownership and invalid input controls in existing native/Python carriers | Nonfinite/out-of-range, invalid section/variant/font, changed owner and allocation failures preserve package, ids and live handles. Exact no-op, repeated replacement/removal, inactive opaque branches and unrelated header/body/borders remain exact |
| final sprint | differential | Native Word and pinned LibreOffice open, local printing PDF, native save, known close and normal disk reopen at0/0.25/0.5/0.75/1 | Exact app/build and source/font/tool hashes, genuine text with requested alpha, asymmetric values distinguish complement error, one paint, correct rotation/geometry and retained body/header/border regions |
| final sprint | golden | Deterministic own renderer and saved-package alpha inspection | API/getters/source agree, endpoints invisible/opaque correctly, no system-font baseline |

Retain `/private/tmp/S91-F-X201-LO-opacity-endpoints-root-1/failed-endpoint-visual-receipt.json` and the old Word endpoint record in the completed F201 design. Historical receipts remain historical. A native investigation failure cannot become local GREEN or completed interoperability.

## HLD impact

- docs/hld/04-opc-and-packaging.md
- docs/hld/08-rendering-spec.md
- docs/hld/10-bindings-spec.md
- docs/hld/12-testing-strategy.md
- docs/hld/14-development-backlog.md

Update implemented intent before microscope/final freeze, preserving pending native qualification explicitly and correcting current scheduling references to S98. Do not rewrite historical delivery entries.

## Risk routing

Parser/serializer: read HLD04 checked source ownership and HLD06 preservation. Final checks retain schema order, namespace scope, exact opaque branches and no-op/refusal byte identity. Layout/font changes: HLD08 and HLD15 deterministic font mode, unchanged bundled inputs and actual native app/build identity. Published API and binding changes, if earned: HLD10, actual affected verified archives below10MiB, both WASM graphs, fresh complete installed family suites and strict typing/stubtest in the final sprint union. External oracles: canonical differential-testing skill, pinned Word and LibreOffice26.2.5.2 with Poppler26.01.0. Native application PDFs are separate from deterministic product baselines. No new dependency, module, crate or tracked font input is planned.

## Hash harness

Expected unchanged for the49 hash and7 deterministic pixel cases. A discovered intentional paint delta must be separately labelled and independently reviewed before the sole baseline writer changes anything. Native outputs are never product baselines.

## Implementation checklist

- [ ] Authenticate Base, prior failure artifacts, source and actual candidate representation authority.
- [ ] Add actual Base RED and minimal faithful correction plus fixed-source GREEN controls.
- [ ] Preserve all ownership, inactive-source and invalid-operation contracts.
- [ ] Update exactly the HLD impact list before source freeze.
- [ ] Pass scoped local checks and independent microscope, with final native debt explicit.
- [ ] Pass the final integrated Word/LibreOffice opacity matrix and own-renderer agreement.
- [ ] Complete delivery records only after actual full-range acceptance.

## Open questions

The user approved carrying this unfinished contract from S95.11 to S96. The bounded source investigation established no faithful shared representation. Repeated deferral triggers a design reset under WORKFLOW.md. This draft preserves the previous contract as input, not an approved implementation candidate. The 2026-10-28 reorganization schedules recovery in S98. Before S98 implementation, identify a concrete evidence-supported approach and revise the test plan or obtain an explicit acceptance disposition. No product implementation, full-range support or native acceptance is claimed. Preserve the existing0.5-only restriction and failed endpoint evidence.
