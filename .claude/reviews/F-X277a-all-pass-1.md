# F-X277a, all aspects, implementation pass 1

**Reviewed**: Working diff on work/f-x277a-codex at Base e6fec886557f819fc28f5ff1e0196edd458b1371. Six files,171 added lines and7 removed lines. Five changed product/test/HLD paths plus the child checklist. The sole production change is the authored constructor's degree-presence boolean.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Clean bounded implementation microscope. Final integrated acceptance remains pending.

## Defects

Zero findings.

## Smells

Zero findings.

## Nitpicks

Zero findings.

## Contract and consumer audit

- `crates/rdocx-oxml/src/math.rs:1031` sets authored degree presence without changing the base, empty degree or hidden-degree value. The independent import parser at:1045 still starts with absent degree and records actual imported presence. The writer at:1087 is unchanged, so the fix does not force degrees into imported missing-degree radicals or rebase preserved child slots.
- Native MathML msqrt/mroot at `crates/rdocx/src/math.rs:747` and LaTeX sqrt at:2201 select the actual shared constructor and indexed constructor. `crates/rdocx/src/paragraph.rs:1041` publishes their actual XML and existing equation sidecar. Python `crates/rdocx-py/src/paragraph.rs:836` uses those same conversions and staged paragraph publication for inline/display modes. No public signature, error policy, dependency, module, feature or helper abstraction changed.
- The inspected official SDK3.0.1 source, SHA2564e182e11492f4dc80fab470438e8bd1d5683fe69dbde591f9a8fe1d49c6ecdf2, records Sequence(radPr0..1, deg1..1, e1..1) at Radical.ConfigureMetadata:1648-1652. The unchanged writer's new authored output follows that order. This static primary-source comparison is not an executed SDK validation.
- `docs/hld/03-architecture.md:839` and `docs/hld/10-bindings-spec.md:1949` describe the implemented authored/imported distinction without claiming imported invalid input became valid or native rendering qualified. Both files are in the approved HLD impact list. The checklist explicitly distinguishes Rust runtime evidence from Python source/AST coverage and installed-runtime debt.

## Tests and evidence challenged

`crates/rdocx-oxml/src/math.rs:3647` combines an empty base, an ordinary square root and an indexed root. Counts and ordering reject a converter-only fix, omitted/duplicate degrees and invented nonempty square-root degrees. It checks hidden-degree values and exact round-trip output. The imported control at:3674 combines an alias, missing degree with and without properties, explicit empty degree, escaped opaque attributes and an opaque child in its retained slot. It checks absence, the existing degHide0 normalization, raw bytes, ordering and a second serialization. Existing optional-argument and raw-slot mutation controls remain unchanged and appear in the retained passing batch. Existing unaffected equation and conversion cases cover the zero-radical path.

`crates/rdocx/tests/integration_test.rs:4605` reaches both real converters, square/indexed forms and inline/display publication, then saves/reopens and checks degree content/order, visibility, equation count and surrounding text. It accumulates missing-degree failures so all four square-root format/display failures are visible on old production. `crates/rdocx-py/tests/test_core.py:6115` uses the actual public Python method and inspects expanded XML names in saved bytes, then reopens. Its execution is pending, not inferred from AST parsing.

The retained evidence under `/private/tmp/rdocx-s96-f277a-implementation` was inspected without running builds or tests:

- All five FIXED_SOURCE source/HLD hashes, the SCOPED_EVIDENCE plan hash, both actual pre-repair binary hashes, both authenticated RED log hashes, all three GREEN log hashes and the Clippy log hash match their files.
- Reversing only the production boolean in memory reconstructs the exact BASELINE_SOURCE_3 lower-source hash. Reconstructing the baseline public carrier from its retained patch matches its recorded hash. Its difference from the fixed carrier is only formatting of the same constructor call and assertion, with no altered test input or assertion semantics.
- Actual lower RED is the missing-degree count1versus3. Actual public RED lists four missing square-root variants, while indexed controls pass. Earlier type/prefix-fixture setup errors are retained separately and are not counted as semantic RED.
- Retained GREEN logs report30 lower math tests,3 public equation tests and17 conversion tests passed. Two existing Pandoc oracle tests are ignored and unperformed. Recorded scoped check and all-target/all-feature warnings-denied Clippy finish successfully. Formatting/prose/whitespace outcomes are recorded separately. No new execution or current binary qualification is attributed to this review.

## Not found

Correctness, contract, panic/resource exposure, OOXML namespace/order/preservation, test discrimination, structural rules and scope each produced zero findings. Production adds no indexing, arithmetic, allocation path or panic. No unfinished canonical Python field-validation carry is present in this child diff. A277-1a through A277-1d have the stated scoped Rust coverage. A277-1e has the specified actual public-call test source, with its runtime requirement still pending under the approved plan.

## Acceptance boundary and handback

Installed Python3.9/3.12 and typing, final authored-DOCX SDK checks, native Word/LibreOffice rendering residual, deterministic-font hash attribution, packaging/archive, normal hosted CI and the final whole-sprint gate remain outstanding. The author does not claim those passed, and this review does not waive them or complete the parent/F272 dependencies.

This is the first actual independent child implementation microscope, within the original maximum of three implementation passes. It is not another design peer, a cap reset, implementer self-review or model evaluation. Only this review file was written. Return the source/build lease to Root for adjudication and the authorized preparation lifecycle. No remediation was performed.
