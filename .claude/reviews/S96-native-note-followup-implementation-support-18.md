# S96 F275 A6 implementation support18

Independent bounded follow-up of Support17 D1 and D2. Both defects are closed for this reviewed source. Zero new defects, zero smells and zero nitpicks. This is support for Root final union reconciliation, not formal review of completed F272/F272a, a peer-cap reset or release acceptance.

## Boundary and authenticated evidence

Branch sprint/s96 at e430071d75ed33e8831752d644aec3c0a865d3a3. Scope is comparison.rs, its integration regressions, HLD03/04/10 and the operative F275 final corrective section. Unrelated pacing, CI, HLD15 and workflow-script changes are excluded. The exact source, dependency and receipt hashes are in BOUNDARY.json, with the scoped working delta in REVIEWED_INCREMENT.patch.

Independent current hashes match GREEN3 source pins: comparison.rs SHA256699a35e69b1cdc362eb74e0a19899cf76ccfdcaa0b9bcb0a0fac1afab30fed0a and integration_test.rs SHA25699338d27290af646d82fde0a5116f4a65abe2d8fbfe306367f3d566abf498790. ISSUE430_NATIVE_REFRESH_GREEN3.json records exit0 and source_unchanged true. Its raw log line66 reports59 passed,0 failed,0 ignored, with all three current terminal-marker/property regressions named at lines51,52,55.

SUPPORT17_PRESERVATION_RED.json records exit101 and source_unchanged true. Its comparison pin exactly matches the previously reviewed Support17 source, and its integration pin exactly matches GREEN3. Raw log lines7-14 retain both property-preservation failures and line22 reports0 passed,2 failed,0 ignored. The old output visibly loses the property-root binding/attribute and the real foreign lookalike. This is a genuine old-source RED/new-source GREEN comparison with identical regression source. Earlier RED2, compile and intermediate-format failures are not replaced or relabelled. Current comparison source has no B_ONLY_TRACE.

## D1 closed: preserve root metadata and actual namespace bindings

The new code no longer calls set_footnote_properties or regenerates the property root. At crates/rdocx/src/comparison.rs:5208 it takes complete source spans only for the selected reference elements. It copies the actual settings source and drains those spans in reverse order at comparison.rs:5212. Consequently the complete property opening tag, local declarations, producer attributes, whitespace, retained child bytes, other policy and endnote bytes remain untouched.

The D1 regression now puts x:flag and the only x binding on footnotePr, retains <x:keep/>, includes a nonempty true Word q alias and a locally rebound foreign q:footnote at crates/rdocx/tests/integration_test.rs:37910. The exact opening and foreign child are asserted after public resolution and reopen at integration_test.rs:37941. Both accept-retained and reject-absent branches remain in this case. The alias removal is resolved by the actual XmlTree element scope, rather than by an inferred wrapper scope. No additional namespace serialization is introduced.

## D2 closed: exact top-level property ownership

The selector now requires the parsed root itself to be Word settings at crates/rdocx/src/comparison.rs:5197. The candidate element must itself be Word footnote and a direct child of a Word footnotePr whose parent is tree.root at comparison.rs:5199. An earlier opaque nested settings/property pair cannot qualify, nor can a foreign-root or foreign-property lookalike. XmlTree uses each element's actual namespace scope at crates/rdocx/src/revision.rs:2260 and derives its Word identity at revision.rs:3380. Local aliases/default namespaces and local foreign rebindings therefore follow actual expanded names.

The D2 regression retains the earlier opaque subtree with its conflicting Word q binding, while the real top-level property binds q to a foreign URI at crates/rdocx/tests/integration_test.rs:37957. Its assertions preserve the opaque subtree, foreign child and actual property opening while removing the actual direct Word special edge at integration_test.rs:37984. The test failed on Support17 and passes on GREEN3.

## Model synchronization, pending state and atomicity

The relationship lookup at crates/rdocx/src/comparison.rs:5174 uses the same first internal settings relationship and resolved actual target as document loading at crates/rdocx/src/document.rs:27353. The code parses the complete updated bytes before installing either package or model at comparison.rs:5216, then writes those same bytes and the parsed model at comparison.rs:5217. No distinct settings-dirty flag or ownership change is needed. The parser retains its source bytes at crates/rdocx-oxml/src/settings.rs:827, and to_xml returns them directly at settings.rs:2163. The existing flush compares a parse of the actual package settings with that same model at crates/rdocx/src/document.rs:29514, so it cannot recreate the removed edges or erase the retained namespace context. Owned-model carry through reopening also carries this updated model at document.rs:28455 and document.rs:28495.

Pending public settings changes are flushed by prepare_staged_package at document.rs:28343 before comparison inputs are read at crates/rdocx/src/comparison.rs:3442 and before revision settlement at crates/rdocx/src/revision.rs:485. No other settings mutation occurs between that preparation and this cleanup in the reviewed comparison/revision callers. Thus raw surgery starts from the current committed-to-candidate settings representation, rather than discarding pending typed changes.

The absence predicate and pending-unit guard are unchanged at comparison.rs:5013. The helper still runs only after actual part absence is established. A retained part does not enter this cleanup. Removing multiple matching siblings is safe: each selected reference is a direct child of a top-level property, so selected element ranges cannot contain one another. Ranges come from a successful parse of the same immutable source. Descending removal preserves all earlier indices. Empty and start/end forms are both complete XmlTree ranges.

Missing settings bytes, XML parse errors or updated-model parse errors return through the existing unpublished candidate. Resolution clones/prepares at revision.rs:485, settles at revision.rs:657, reopens at revision.rs:664 and only then commits at revision.rs:675. Comparison likewise stages input/candidate work and commits only at comparison.rs:1770. Existing missing-owned-target and partial-selection atomic controls remain in the actual59-case cohort at crates/rdocx/tests/integration_test.rs:38819 and integration_test.rs:38934. No new settings-specific malformed-input execution was performed in this review. Error atomicity for this repair is established by the unchanged candidate publication boundary and parse-before-install ordering, not claimed as a new executed negative test.

No additional completed B-only finding. Its exact token/bits/pair/part/full-paragraph/relative-boundary discrimination is unchanged from Support17, including no lost-owner inference and the existing inside-revision guard. No API, callback, cache, admission policy, Source credit, status or ownership model is added. The raw repair is smaller and keeps the implementation local.

## Contract and pending riders

The operative F275 HLD list now explicitly includes HLD03 and04 at .claude/plans/F-X275-design.md:270. The final corrective section at plan:891 preserves historical done-provider status, the existing G09/G12 and A430-35/A430-41 debt, and independent/current-source installed/archive/hosted requirements. HLD03/04/10 match the bounded current intent. No preservation exception or gate waiver is introduced.

GREEN3 is filtered local Native integration evidence only. Fresh installed qualification of the exact actual Word pending/resolved inputs, complete applicable SDK2019/2021 output validation, strict relevant checks, archives, hosted authority and final union reconciliation remain required. Historical Word16.114.4 saves and SDK passes remain bound to their original bytes. Neither Support18 nor59 passing local tests constitutes full acceptance, publication authorization or release approval.

Only this private/tmp support packet was written. No source, tracked documentation, Cargo, runtime, SDK, tests, UI or network action occurred. End support18 and return control to Root.
