# S96 F273 A3 frame-offset causal repair, support pass 1

**Reviewed**: The frozen four-path increment against cec87f9501eb13d0d5d60572472b227483262620, 175 added and 19 removed tracked lines, plus the existing ignored progress checkpoint. This independent implementation review covers correctness, contract, panics, OOXML, tests and structure. It does not reset the original feature review cap.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Source and evidence authentication

SOURCE_BOUNDARY.json at /private/tmp/rdocx-s96-F273-A3-frame-offset-implementation-1 has SHA256 f08922035cb71302a9932f16ebdc721be90e0d55143d9e37eb1a52d3ec2bf10b. Its increment.patch has SHA256 e80c833227cfeae73107ad486120d12253f27fa4ec4a023a7d0162f4a30d4067. All five current source/progress pins and all 46 retained evidence-file hashes matched during review. The source pins are:

- paginator.rs, 6730b92008d9c7599c6dc8d4526c376b604f49d77d6f357684e971f7817a5597
- engine.rs, efcfbe13f28b72de0461b3b56dd2e4aebd70bd4022bb9dfd7e3fbe4e904479d1
- F-X273-design.md, f69e8c268cb085a95d96216538a63f770f3f819fe28699d4f7b41c3c1c654fc4
- HLD08, 6837abd3bee79bdb5cf40e4b2b4ab96a565a6beab86bc7fa96bb7ad25d70003a
- Existing F273 progress, 3a0710d1ab45d4dd43b0639ec6aa168608d4f86dbf68fdd30e6899c333611f9b

The actual Root design approval has SHA256 717726832590a9411d02bb51c9c0687413bb486d90065ad2508dc07df33d771b. Its finite implementation and native requirements match the appended approved contract at .claude/plans/F-X273-design.md:1012 and HLD reconciliation at docs/hld/08-rendering-spec.md:2066.

The retained causal ACTUAL_HANDOFF.json has SHA256 3e39fcb4591ac03749106b5186533a86f75a12ad4ebefc913f312364bc61c05e. Its observer, raw logs, result and interpretation hashes match. All nine result source/tool/PDF pins match. The actual original, no-frame and shifted native PDFs show Page origins 139.9499,299.9499, then 80,120, then 169.9499,359.9499. Each has exactly two ten-point rectangles and the actual source anchors remain unchanged. This is evidence for explicit offsets, not alignment or a footer-native claim.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Aspects checked

Correctness and contract: crates/rdocx-layout/src/paginator.rs:7081 remains reachable from the actual ordinary header/footer Frame branch at 7197. Its existing prepared frame geometry remains authoritative. At 7108 the implementation obtains the ordinary anchored_rect first, then adds only the actual frame origin to explicit Page-offset axes independently. Alignments keep the existing section-page basis. Paragraph and all other axes retain the old resolver. Global anchored_rect at 7780 and true-cell place_cell_anchored at 7813 are unchanged. Body frames do not enter this furniture-only emitter. No field or source metadata is rederived from matching coordinates.

Panics and paint ownership: paginator.rs:7100 validates the physical anchor count before indexing its paired paint owner at 7120. The new loop preserves physical enumeration order. It calls the same anchored_elements owner/type validator as the old cell helper, including the existing mismatch error at 3818. Behind/front publication at 7125 preserves the old branching and fallible append behavior. Rotation, images, grouped policies, shape-text production and media lookups remain in the unchanged typed emitter. There is no new unchecked owner lookup or swallowed error.

OOXML and structure: no XML reader, writer, namespace, model, source-part or raw-attribute behavior changes. The implementation modifies an ephemeral resolved rectangle only. It adds no API, persistent state, trait, generic, wrapper, module, crate or feature flag. The short local loop is justified by its actual ordinary-frame offset distinction, while the general cell/global implementation remains unchanged.

Tests: engine.rs:32698 retains frame grouping, body position and exactly-once physical shape checks before the new 140,300 assertion. The retained test-only predecessor actually fails on 80,120 with exit101 after those invariant checks. Its unchanged product and retained predecessor binary hashes match. The historical 120 assertion was previously green and is expressly preserved as such, rather than represented as historical product RED. The new engine.rs:32767 matrix checks header and footer, original and shifted frames, unframed placement, both independently mixed alignment axes and both aligned axes. It preserves dimensions and paragraph-origin checks. paginator.rs:10677 adds a true-cell Page control that retains section-page positioning.

Actual local evidence: the frozen issue424 cohort selected and passed 31 tests, including the 12 matrix subcases. Eight separately named current-binary controls passed for actual cells, offsets, alignment, layer placement and cache identity/invalidation/transactionality. The retained current binary has SHA256 dc66f87a7acd6a8512f882898193a904bed72ad25719b496f9c6cd192319ed89. Default all-target check, strict all-target/all-feature Clippy, final format, prose and diff checks have actual child exit0 and stable source receipts. These existing observations were authenticated, not rerun by this review.

## Qualification limits and handback

This review finds no blocking source issue in the bounded increment. Root final review remains separate. Fresh qualified current-code rendering against the exact original, no-frame and shifted native inputs is still mandatory at the unchanged 0.5 pt measured-edge criterion and exact paint counts. No footer-native or Page-alignment parity is inferred. The separate Word-generated VML duplicate-ID SDK failure remains retained. Whole-sprint native, full-suite, installed shipping, performance, memory, archive and hosted acceptance are not granted here. No source, status, ledger, baseline or artifact was changed by this review.
