# F-X272, namespace investigation, pass 4

**Role**: independent Codex design escalation after full microscope pass1. Read-only outside this report.
**Reviewed**: dirty `work/f-x272-codex`, Base `c8acc502b74f29ed7e9546939f17007709a9f22b`, 14 tracked files, 3056 added and 505 removed lines. Tracked binary diff SHA-256 `8406a9aacbcae05c3963e6adef9112c1edf5b72f3f1dd0e1e534d16773066381` remained unchanged during review.
**Outcome**: D2 remains unresolved at property source authority, and N1 is reopened at equivalent generated-container publishers. D1 has actual local fixed evidence. There are 2 focused defects, 0 smells and 0 nitpicks. This is not full microscope pass2 or an integration-readiness verdict.

## Actual evidence and its limits

- `fx272-d1-cli-red.log` actually reproduced body1/text-box1 misclassification and human2stories while total/main2 and preserved input/text-box bytes were correct. The current stable key includes kind, part and owner index, excluding fingerprint, at `crates/rdocx-cli/src/commands.rs:1935`. `fx272-d1-cli-green.log` executed both CLI Issue425 cases, 2 passed, 0 failed, 100 filtered, 4.56 seconds. D1 is locally repaired, with final independent feature review still required.
- `fx272-d2-native-red.log` recorded the actual four no-pPr/existing-rPr and insertion/deletion token-parse failures. The current reader-boundary implementation fixes those lexical insertion sites. Its successor `fx272-d1d2-native-green.log` actually ran 24 named tests: 23 passed and 1 failed, 4.22 seconds. The first strengthened existing-rPr resolution case loses x:note entirely at `crates/rdocx/tests/regression_test.rs:409713`. Assertions stop at that failure, so later loop combinations are not independently credited as passing. This is a property-preservation failure, not merely different escaping of the quoted greater-than.
- `fx272-foreign-w-row-diagnostic.py` and `.json` authenticate native hash `7a36134634571349e8de2ba91eefc947eec98d65249ae0e34bff7cdbda81e353`. That is the PREDECESSOR wheel, not the current D1/D2 candidate. Canonical Word and q-alias without foreign w passed. The q=Word/w=producer row-insertion case failed accepted body item1. This is a genuine predecessor observation, while current-candidate applicability below is source tracing. A new current-source RED is still needed before claiming a repaired GREEN.

No product tests, builds, wheel installations, broad suites or oracles were executed by this investigator. Source, tests, plan and progress remained frozen. Earlier successful native/Python observations remain attributed to their respective producers and do not qualify this partial repair.

## Findings

### D2, final property reconstruction discards the raw property owner's authority

`crates/rdocx/src/comparison.rs:4214`
`crates/rdocx/src/comparison.rs:4321`
`crates/rdocx/src/comparison.rs:4405`
`crates/rdocx/src/comparison.rs:6228`
`crates/rdocx/src/comparison.rs:6233`
`crates/rdocx-oxml/src/numbering.rs:1389`
`crates/rdocx-oxml/src/run_properties.rs:606`
`crates/rdocx-oxml/src/run_properties.rs:1189`

The reader-token repair allows the original supported source through comparison, exposing a later loss in the same D2 preservation contract. A predecessor's q:rPr can contain opaque root metadata such as x:note. `parse_raw_rpr_with_owner_bindings` consumes that start element, computes namespace context, then advances the reader into a CT_RPr parser that starts from default and reads children. CT_RPr does not retain those opaque root attributes, and its serializer writes a new bare w:rPr root. The typed value cannot recover the original root shell later.

`EmittedOwner::Paragraph` retains only modeled accepted/rejected CT_PPr references. `mark_previous_paragraph` returns one of those references after marking the raw predecessor. `final_paragraph` then reconstructs current and prior properties from typed values through `final_paragraph_properties_xml`, `property_xml` and `run_property_xml`. The raw predecessor contains x:note before resolution, but the property side needed when the final paragraph mark merges is regenerated without it. The actual failing output contains a new pPr/rPr with bold formatting and no x:note.

This is source authority loss, not a reason to parse x:note as a special modeled attribute. Merely putting the attribute text on a fresh w:rPr is also insufficient: if the source q:rPr means Word while an opaque w:note attribute means a producer namespace, a local canonical w binding on that new root would change the attribute's expanded name. Root QName, effective context, unknown attributes, raw children and property precedence travel together.

#### Concrete private repair boundary

Carry raw property-source snapshots beside the existing typed accepted/rejected property sides. Both ORIGINAL and EDITED source views are required. Main and related comparison entrypoints already have both raw bodies, spans and bindings at `comparison.rs:966` and `comparison.rs:2312`, but only the original view reaches compare_body at `comparison.rs:995` and `comparison.rs:2330`. Preserve that information privately before the typed parser loses opaque property-root state.

For each emitted paragraph side, retain the qualified direct pPr/rPr source shell and content, including actual inherited/local namespace authority. Capture the pre-marker source, not a later tracked result that already contains the new mark. Carry the selected source with its typed value through ordinary, trailing and moved predecessor handling. The current precedence remains: deletion/moveFrom uses merged predecessor properties on acceptance and own final properties on rejection, while insertion/moveTo uses own properties on acceptance and merged predecessor properties on rejection. The final paragraph's own section-break precedence remains unchanged at `comparison.rs:6227`.

Generate only the actual owned change nodes and modeled edits, then patch their qualified slots into the correct preserved source template. Preserve the source root QName and raw opaque metadata. Strip or replace only the intended DIRECT Word rPr, section and property-change children according to the existing schema contract. Foreign same-local lookalikes remain opaque. Use the existing reader-boundary and namespace-closure helpers rather than searching attribute text or inventing a generic XML equivalence rule. The shared resolution renderer already transports actual source context and must receive complete property snapshots.

Unknown-formatting policy remains left-biased where it was previously left-biased. Do not start accepting opaque metadata edits merely because both raw input sides are now available. Do not weaken revised-owner refusal, old-revision byte checks, accepted/rejected reconstruction, marker ordering or atomic publication. An absent property owner and an empty owner carrying opaque metadata are different source states, even when both project to the typed default.

This can stay in private comparison paths and existing files. No new public model field, public type, parser for unknown attributes, trait, module, crate or source file is shown to be necessary. Record the source snapshot/precedence clarification and all reached callers in the existing plan before implementation. If implementation discovers an actual need to broaden a public model, return with the contradiction rather than silently doing so.

### N1, generated Word containers still bypass local qualification at publication

`crates/rdocx/src/comparison.rs:6619`
`crates/rdocx/src/comparison.rs:6774`
`crates/rdocx/src/comparison.rs:6793`
`crates/rdocx/src/comparison.rs:6804`
`crates/rdocx/src/comparison.rs:4198`
`crates/rdocx/src/comparison.rs:650`

This reopens the existing generated/retained namespace authority invariant. Markers now use IdAllocator.generated, but a qualified marker does not qualify its containing row, row properties, cell or paragraph. `marked_row` obtains fresh canonical row XML, `marked_row_xml` may add a fresh w:trPr, and the completed output returns without qualification. `compare_table` inserts that result under the retained source table. Under q=Word/w=producer, those new canonical w containers can become producer elements. The new marker's own correct local binding cannot repair its ancestors. `moved_row` and table/control composition expose equivalent publication paths.

The predecessor diagnostic is consistent with this source defect, but it does not isolate rows as the ONLY cause. The same table path also replaces the retained source properties with table_properties_xml. Its equal and ignore-formatting branches return fresh canonical serialization without the qualifier at `comparison.rs:6657` and `comparison.rs:6674`. A single generated() call around one new row would leave those property branches and recursive publishers inconsistent.

The policy owner already exists in IdAllocator.generated. Classify each output as fresh canonical Word markup, actual retained source, or their composition. Qualify each newly generated Word container at the boundary where it is published beneath retained source, including early returns that are semantically unchanged but freshly serialized. Retained raw fragments first retain their ACTUAL source context. Do not globally rebind producer w or blindly declare Word w on an opaque source root. Both properties and complete containers need the same consumer-map closure.

## Required publisher inventory

These are concrete callers to reconcile. Rows without an actual observation are source-audit obligations, not additional claimed executions.

| Boundary | Required authority and proving path | Citation |
|---|---|---|
| Final predecessor properties | Typed precedence plus raw source snapshots for both sides, before new markers | `comparison.rs:4214`, `comparison.rs:4303`, `comparison.rs:4368` |
| Final insertion/deletion/move | Select own/merged accepted/rejected source, preserve own section precedence | `comparison.rs:3884`, `comparison.rs:3911`, `comparison.rs:3958`, `comparison.rs:4405` |
| Block-control final paragraphs | Same snapshot policy inside existing sdtContent, no alternate typed-only merge | `comparison.rs:7253`, `comparison.rs:7261`, `comparison.rs:7291`, `comparison.rs:7299` |
| Ordinary paragraph property regeneration | Audit raw pPr/rPr authority when text or modeled properties change and for ignored formatting | `comparison.rs:6056`, `comparison.rs:6078`, `comparison.rs:6134`, `comparison.rs:6143` |
| Ordinary run-property revisions | Audit preserved source shell and prior property source, without parsing opaque root attributes | `comparison.rs:7681`, `comparison.rs:7691`, `comparison.rs:7952` |
| Inserted/deleted rows | Fresh row/trPr/cell/paragraph containers close locally, retained content keeps its scope | `comparison.rs:6619`, `comparison.rs:6774`, `comparison.rs:6807` |
| Moved tables/rows | Completed fresh roots and property containers use the same qualifier | `comparison.rs:4181`, `comparison.rs:4191` |
| Marked whole tables and controlled rows | Composition cannot leave a fresh table/row/paragraph/cell under a foreign-w retained owner | `comparison.rs:6725`, `comparison.rs:6750`, `comparison.rs:6760`, `comparison.rs:6762` |
| Matched/unmatched control children | Existing paragraph/table/row publishers share the rule, while unsupported unmatched cells/nested controls still refuse | `comparison.rs:7628`, `comparison.rs:7650` |
| Generated property early returns | Equal or ignored typed values are not proof that freshly serialized names retain source meaning | `comparison.rs:6078`, `comparison.rs:6331`, `comparison.rs:6340`, `comparison.rs:6657`, `comparison.rs:6674` |
| Generated field fragments | Audit preserved field payload versus fresh Word run/field containers at direct publication and early return | `comparison.rs:7707`, `comparison.rs:7709`, `comparison.rs:7731` |
| Already qualified paragraph publishers | Existing moved/inserted/deleted/final paragraph completion is the pattern, not proof for other publishers | `comparison.rs:4167`, `comparison.rs:4441`, `comparison.rs:6373`, `comparison.rs:6389` |

## Bounded discriminating checks

### D2, retain the exact reproduction and prove source precedence

Keep the original four quoted-opening contexts in `issue425_final_marks_respect_quoted_greater_than_attributes`, not a replacement with ordinary text. The current attribute-loss observation is an actual pre-preservation-repair RED. Collect individual insertion/deletion and resolution-side outcomes if the test currently short-circuits, without weakening its original byte/value assertions.

Add a structurally different raw-property discriminator within the existing carrier: q:rPr is Word, w is a producer binding, and opaque w:note contains a quoted greater-than. Include an unknown child and source-local alias or value-only namespace binding. Verify root QName/context and expanded opaque element/attribute identity, original raw value/children, intended typed formatting and selected old revision bytes through new-author accept/reject and save/reopen. Merely asserting that x:note text exists is insufficient.

Use distinct own and merged property values so the original/edited and accepted/rejected sides cannot be swapped. Cover both absent and empty-with-metadata owners. Include one existing supported block-control or related-story final paragraph path as the cross-caller discriminator. Retain all existing property-rejection, legacy mark, schema-order and direct relocation controls. A source-bearing predecessor that itself holds an old revision must still refuse atomically. Existing non-final property regeneration needs either a focused proving case under the shared source helper or an evidence-backed exclusion, rather than another typed-only implementation hidden beside the repaired final path.

### N1, prove the current candidate and shared publication rule

First execute the exact row diagnostic's three controls on current source: canonical, q alias with ordinary w, and q=Word with foreign w. Keep the predecessor receipt unchanged. The current source must earn its own RED before a repaired GREEN is credited.

Use a bounded matrix that distinguishes container and property publishers: inserted and deleted ordinary row under a retained table, an unchanged nonempty tblPr/paragraph property owner, and an existing supported table move or controlled-row/table path. Prefer combining row/property and recursive/control dimensions in a few meaningful inputs. Include a neighboring retained opaque producer-w element/attribute and an unchanged old revision owner. Verify expanded names of every actual generated row/trPr/tc/p/mark, intended row count/text and property values, exact old revision/opaque source bytes, and accept/reject/reopen outcomes. A cell-text-only alias test cannot prove row insertion or deletion.

Preserve atomic refusal for unsupported unmatched cell/nested-control topology and any revised whole table/control edit. Do not add a new row-move capability solely for testing: exercise the table/move publishers that exist today. Any new related or controlled context labelled RED requires actual pre-fix execution. No Cartesian product, broad crate suite, native oracle or per-worker hash run is requested here.

After both shared repairs, refresh affected focused native/CLI and scoped check/clippy/fmt observations. A successor Python result needs a successful freshly built/authenticated artifact for the changed source. Full feature microscope pass2 follows actual local qualification. All final integrated native/hosted/installed/typing/WASM/hash/pixel/archive/README/risk gates remain pending.

## Not found and handback

D1's stable-coordinate repair and actual2/2CLI result are accepted as local facts, without substituting this investigation for full review. N3 equality-only metadata handling was not implicated by the new observations and is not being reopened. Zero new independent panic, structure, smell or nitpick findings are reported.

No material user product choice is required. The smallest coherent plan clarification is private both-input property-source transport plus complete generated-container publication coverage, preserving existing opaque metadata and namespace authority. It does not authorize unrestricted edits inside prior revision-bearing owners or a new public OOXML model surface.

Only this report was written. No remediation occurred. Return control to root now with D2 and N1 unresolved.
