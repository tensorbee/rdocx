# S96 bounded semantic TOC and normalization source review

Verdict: 0 new source defects, 0 smells, 0 nitpicks in the assigned b823 provider changes. The existing accepted-revision SEQ runtime failure remains open and is not accepted by this source review. No source changes or Cargo were performed.

## Actual source transport

At crates/rdocx-oxml/src/revision.rs:642-655 the successful revision reader retains its actual input projection map with publication roles frozen and a weak handle to the same registry. At :718-731 semantic preparation constructs WordFieldInput::produced from that map and the upgraded live registry. The produced constructor at crates/rdocx-oxml/src/document.rs:2902-2912 sets initial_delivery false. It does not register synthetic output as a new original input.

The selected physical revision body is located by revision_content_bounds at crates/rdocx-oxml/src/revision.rs:1776-1786. The actual synthetic paragraph pieces undergo the existing admission before Vec growth at :1803-1825, and its map slices only that original content range then adds generated opening/closing shell coordinates. WordFieldInput::slice and with_shell at crates/rdocx-oxml/src/document.rs:2986-3005 preserve initial-delivery state and registry, with generated shell maps. The parser consumes that actual paragraph cursor before recursive semantic preparation. The retained raw revision XML stays the serializer authority. Clearing read_context occurs only after successful preparation. Ordinary unscoped None behavior is retained, without inventing scoped provenance when no live registry is available.

## Normalization and TOC consumers

At crates/rdocx-oxml/src/text.rs:11906-11917 the conditional retainer prepares the actual freshly parsed current paragraph only when accepted paths differ. It then recomputes and compares the same path sequence at :11926-11989. The subsequent content-slot, form, instruction, display, nested cached-child and actual current-map proofs remain at :11990-12117. This preparation does not authorize a mismatch, swap old identity in place of a current map, or change the final refusal.

At crates/rdocx/src/field.rs:3422-3433 rebuild_toc prepares its owned staged candidate before inspect_toc_bookmarks and discover_toc_sources. It no longer assumes layout's separate clone prepared the authoritative discovery body. The staged operation and unchanged source XML remain available to the established qualified discovery and atomic publication. F-X273-design.md:31 retains existing rebuild_toc API and atomicity, and its local gate at :90-94 remains an actual-runtime obligation.

## Explicit unresolved observation

Root reports lower semantic/normalization controls12 passed and most TOC cases passed. The actual accepted_revision_sequence_sources_keep_deleted_physical_ids_through_update_and_reopen still fails in /private/tmp/rdocx-s96-h6-regression-repair-2/affected70-and-comment-controls.stderr with normalized field owner changed its accepted source paths, from the unchanged final identity guard in crates/rdocx-oxml/src/text.rs:11987. The reason for that remaining path divergence is pending diagnostic analysis. No weakening of that proof, successful SEQ lifecycle claim or whole-provider completion follows from this bounded source verdict. The temporary conditional diagnostic helper is excluded from shipped logic and must be removed before the final source freeze.

Reviewed source pins:

```json
{
  "crates/rdocx-oxml/src/revision.rs": "ee36efddff7ff78c74fb797ab8519157d31fc2d6fbe5164e67113024713787e8",
  "crates/rdocx-oxml/src/text.rs": "7d6882ccfcd8167985a6e415474d392e1afe9cadb39d668433aee6a77e335a3e",
  "crates/rdocx-oxml/src/document.rs": "c063c9e8a93238b25c7e01484fb7ba1f1d9273da97192d4b379bfb23855cd3c6",
  "crates/rdocx/src/field.rs": "6f138eb30d0a09936e22f24c51af68a7a579a3d8ee46368a712363c144e25e67",
  "crates/rdocx/src/document.rs": "e4e4f77fd2c6e7946b44e73cb56910c5bb020fc742cefd7a60945d0bb67f1c8b"
}
```
