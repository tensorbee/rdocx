# F-X272, Root import Source functional support, pass 1

**Reviewed**: The immutable five-path snapshot at `native-template-support/returned-reader-functional-batch/ROOT_TAKEOVER/root-main-note-import-batch/source`, compared with the previously reviewed `root-returned-reader-review-freeze/source`. Three Native paths change, with 1,718 additions and 299 deletions. Both lower snapshots are unchanged. This is bounded independent support for Root disposition, not whole F-X272 acceptance or a reset of an earlier review cap.

**Verdict**: 1 defect, 0 smells, 0 nitpicks.

## Defects

### D1, the frozen increment fails its required strict check

`crates/rdocx/src/document.rs:1391` and `crates/rdocx/src/document.rs:28197`

The exact archived strict command exits 101. `TemplateFragmentSelection` introduces a 456-byte Source variant alongside a 24-byte byte-buffer variant, triggering `clippy::large_enum_variant`. The additive-note append branch also calls `Arc::from(source)` where source is already the same Arc type, triggering `clippy::useless_conversion`. With warnings denied, neither the library nor its test target completes the required strict check. Resolve both diagnostics and record the affected successor check before accepting this snapshot as locally gated. The passing functional cohorts do not establish a strict-check success.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Source and evidence authentication

All five source hashes authenticate against the frozen `RESULT.json`.

| Path | SHA256 |
|---|---|
| crates/rdocx/src/document.rs | 16057c9435f9122f4ceb22232d368333114ff0a94e2dbc159a56089da4c55c2f |
| crates/rdocx/src/template.rs | 90c4aca1cbed47fe91360c2237a786829e2eecfeb7ebc8cb54b2ad8492b5401b |
| crates/rdocx/src/field.rs | 2e9e10e931e03b263cb3dd3297428b0c159f4a15ad58cbb7fed979104c81285e |
| crates/rdocx-oxml/src/document.rs | 53bcc06a158951f9bb7fb3a46310bf961e450aab79a7b7bdd9ef356db74e26b3 |
| crates/rdocx-oxml/src/text.rs | 772df33c5dd4738a1ba1591c5ed3f4e859beb0f39412c3c48e0e7735ae26807b |

The archived Native binary authenticates as `adaed65b307868eb63b90c7f17838ac07f92a895e25b70ae4d51906888b5382a`. Actual logs record 21 scoped passes, 18 versioned passes and 10 fragment passes. These are separate named cohorts, with potential name overlap, not 49 unique controls or a full suite. The versioned and fragment commands exit 0. The strict command is all-targets with warnings denied, exits 101 in 10.40 seconds, and is not an all-features success. No reviewer compilation, test, endpoint or native application execution was performed.

## Functional assessment

The main Source and returned-reader Source now share actual immutable Source ownership through Arc. `TemplateMainSources::read_document` records the paired successful factory result and disables initial model lineage. `published_main` binds the actual package Arc to its producer Source. Retained-main selection checks that exact storage identity, with the frozen original-part path as its distinct fallback. It does not infer provenance from equal output bytes. The retained-main edit branch applies the real relationship/drawing plans and reads the resulting Source before the next writer. Ordinary None-context cached behavior remains separate.

The scoped fragment route carries the selected Source through `DocumentFragment::from_part_source`, the existing ZIP writer/loader and the actual same-operation whole-document factory. Companion Sources follow their actual part names into the loaded package. The main raw input remains distinct from the prepared model used for dependency inspection. Shared edit planners preserve existing scanner behavior while assembling copied intervals and generated replacements through the operation. The ordinary callers still consume the same planner results as byte edits. Glossary-specific fallback remains explicitly outside claimed Source closure.

The copied-note path selects the physical owner at `fragment_note_dependency_in`, slices that exact current Source window and performs the existing namespace closure. The note-copy record retains its Source through note-ID, store, relationship, identity, numbering and marker edits. Generated companion framing is sliced away through the actual wrapper geometry. Publication appends through the existing root-position planner or replaces the already selected owner range. It installs the Source beside the final package Arc. Scoped final preparation retains that handback instead of silently entering the ordinary preparation path. Source-aware style remapping also preserves or updates the actual related-part Arc, rather than attaching facts to an obsolete buffer.

Initial clone paths now include retained paragraph/control revisions and Ruby raw children. Revision retained bytes come from the actual accessor, with content paragraphs, physical runs and nested revisions traversed through their real clone paths. Revision branches are withheld after a recorded revision Field replacement. Ruby base runs already use ordinary paragraph Text facts, and phonetic text is not added to visible selection. Transient pair deduplication removes repeated visits to the same actual borrowed payload and rejects conflicting frozen pairs. It is not used to discover origins after mutation.

## Controls examined

The retained-picture control compares the scoped and ordinary package/model outputs, checks final package Arc identity and then exercises a charged Text consumer on A&B. The allocated-note control requires a real old-to-new note allocation, a Source on the imported main program and the exact final related-part Source. It parses that note family, selects the actual returned physical note index, emits the selected note and exercises the next charged Text consumer. This distinguishes final-byte parity from preserved reader/writer correspondence.

The initial revision/Ruby control combines an insertion containing A&B, Ruby base BASE, phonetic TOKEN and an opaque retained child. It requires visible A&BBASE, retained opaque data and a successful charged consumer. It does not qualify changed revision reparses or arbitrary returned inline-control contexts.

## Explicit boundaries

Changed or rebuilt revision returned-model facts, paragraph-inline control returned maps, complete non-emitting expression evaluation, remaining related-comment/glossary paths and final context joins remain pending. This report does not certify a whole ordinary inventory, all importer contexts, full C1, provider readiness or feature completion. Same-operation Source/model pairing retains the adopted caller responsibility. No new quota, API, parser, input grant or security assessment is approved here.

Final source integration, complete callable inventory, SemVer, full local and integrated checks, hash/native/oracle riders and final acceptance remain pending. The strict failure above must remain distinct from any later repaired snapshot.

## Not found

No additional concrete functional correctness, panic, namespace/schema-order, XML preservation or producer-origin defect was found in the bounded mechanisms examined. No independent structural smell or nitpick was found. These zero categories do not close the explicit pending contexts.
