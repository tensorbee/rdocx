# S96 H9 namespace and table cross-review support

Verdict: 0 defects, 1 smell, 0 nitpicks. This is bounded independent support for the six-file H9 candidate, not another formal feature peer pass or whole-provider acceptance. No source edits, Cargo commands or runtime probes were performed.

## Source and contract

The reviewed base is H8 `924b348c7d019a41b80c2a955284b2d83c519d10`. The working diff has six files, 383 insertions and 22 deletions. Its binary diff SHA256 is `1c2592dff36117f00069e823efa902faaaccf9e2e6cb02ad825bf9134fb578b9`.

| Path | Reviewed SHA256 |
|---|---|
| crates/rdocx-oxml/src/numbering.rs | bf2b5ad5862933fbe633d40b103443ba2bc0f446df7ad0a2858e992cd65bd497 |
| crates/rdocx-oxml/src/paragraph_properties.rs | acdff79eab58f183efe5856875d407c12cc9fff2c45e7ab3f5edd853ad8ea251 |
| crates/rdocx-oxml/src/run_properties.rs | bde9f5ace7a614cd79e873a0e9d283ecdc1186443ddaa2ef3ad98c281a93d1fe |
| crates/rdocx-oxml/src/table.rs | 7a48261e4335c675e690d59bf0045794b1dc8177d9d257d9b2074bc9967e3da2 |
| crates/rdocx-oxml/src/text.rs | cc4fa5bbd3ab3332dead67b12c120d827ac8e6437b24e6bd6002ff71189ebff6 |
| crates/rpptx/src/html.rs | 9380b3dd7621b4047fe6b6620d6d39595249bca13b5a58cc094c247917d67f3d |

The review used the canonical microscope command, F272 namespace and Source versus Publication requirements, the corresponding HLD03 owner-retention contract, and F281's distinct empty-format authoring refusal and imported painter fallback. The scope covers the complete six-file diff and adjacent parser, namespace, positioned raw replay and emission helpers. It excludes final native application, performance, memory and integrated acceptance riders.

## Findings

### S1: A single-prefix property query repeatedly builds the complete namespace catalog

The new Publication checks at `crates/rdocx-oxml/src/paragraph_properties.rs:1399` and `crates/rdocx-oxml/src/paragraph_properties.rs:1914`, and `crates/rdocx-oxml/src/run_properties.rs:1551` and `crates/rdocx-oxml/src/run_properties.rs:2109`, call `emitted_bindings().bindings()` merely to determine the effective `w` binding. `crates/rdocx-oxml/src/text.rs:535` allocates a vector for all effective prefixes and checks every candidate against its accumulated vector at `crates/rdocx-oxml/src/text.rs:541`. For K distinct inherited prefixes this is quadratic catalog work, newly repeated for each emitted property record, including empty records under an ordinary canonical Word destination. An existing root scope's catalog construction does not remove this additional per-property work.

Use a shared private effective-prefix query that walks nearest frames and reversed local declarations until `w` is found, preserving the existing canonical default when absent. The four sites need one binding fact, not full catalog construction. Keep restoration and emission accounting unchanged. This is a source-confirmed unnecessary allocation and complexity smell, not a measured runtime regression or a claim that an acceptance threshold failed.

## Aspect results and evidence limits

Correctness: 0 defects. The four property-root paths restore canonical Word only for Publication. Source qualification remains on the existing route. The modeled paths keep a temporary destination scope until child emission finishes. Guard drop restores the enclosing context on ordinary return and error. Empty property records have no child scope to maintain.

Contract and OOXML: 0 defects. At `crates/rdocx-oxml/src/numbering.rs:1413` and `crates/rdocx-oxml/src/numbering.rs:1483`, inherited foreign `w` is captured at the actual property reader. Existing owner and then local overrides take precedence through `merged_owner_bindings`, so retained opaque names keep their original expanded identity when Publication restores Word on the generated property root. Positioned raw replay preserves the generated root declaration and retained child order. Added declarations use the existing FieldXmlStart reservation and write path, without new input authority or an alternate unmetered emitter.

The qualified Empty table branch at `crates/rdocx-oxml/src/table.rs:3687` retains actual present empty properties as Some(Default). It does not synthesize an absent owner. Local aliases and shadowed foreign names still pass through the existing expanded-name classifier and raw fallback. The existing strengthened test distinguishes present empty, paired empty, absent, local Word restoration and foreign raw preservation, and checks schema order and reopen stability.

Panics: 0 findings in the delta. No new unchecked input indexing, unwrap or arithmetic was introduced in the production changes reviewed.

Tests: 0 findings. The property tests independently resolve emitted expanded names under canonical and foreign destinations, test modeled and empty roots, preserve opaque foreign child identity and local overrides, and verify restoration of the enclosing destination. The test-only PPT change at `crates/rpptx/src/html.rs:3323` exercises actual empty-code authoring refusal, then constructs an imported empty-code package through the existing relationship and checks the unchanged painter diagnostic contract. It does not relax the public authoring rule or change product code. The precise malformed modeled property error and foreign opaque counterpart remain distinguished with atomicity assertions.

Structure: 1 smell, S1 above. No new public API, module, trait, feature, parser or budget was introduced. Nitpicks: 0.

Root's actual `/private/tmp/rdocx-s96-H9-integration-checkpoint-1/RESULT.json` records source-stable fmt, scoped all-target/all-feature Clippy, lower library, targeted PPT and hash-harness child exits of 0. Root reports 1,289 lower tests, one targeted PPT test and all 49 hashes unchanged. This review consumes those receipts without presenting them as new reviewer executions. The independent Markdown untimed prerequisite work is separate and terminal, with 84 chart controls and 426 composite registry rows qualified. No clocked benchmark, RSS result, native application result or final gate pass is inferred here.
