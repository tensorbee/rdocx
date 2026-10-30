# F-X272, all aspects, pass 5

**Reviewer**: independent Codex reviewer, not an implementer of this repair.
**Reviewed**: the working result on sprint/s96 at6506fa660372b2cfe7dd620cfaaf0959a62ba640,18tracked changed files,33263added and8800removed lines. Tracked binary diff SHA25628504a1f827131122e93a23906c30032c4b04c54597220d8f4bfe88990336dae stayed unchanged throughout inspection and immediately before this report. The count includes the separate F-X273planning amendment. This pass assesses the F-X272prior-repair checkpoint and its composed consumers, not F-X273completion.
**Contract**: read repository/workflow/microscope guidance, current approved prior-repair plan, progress, cited HLD03/04/10/12/14and prior full/focused reviews. F-X272is in-progress. Issue430is separately admitted draft and unimplemented additional scope. Its open contract and native evidence are not treated as implemented acceptance.
**Verdict**:1defect,0smells,0nitpicks. D3blocks microscope readiness. Passing exact256does not prove the remaining schema ancestry. No feature or sprint completion approval.

## Defects

### D3, carried whole-paragraph fields still put fldSimple directly under run revisions

`crates/rdocx/src/comparison.rs:13887`
`crates/rdocx/src/comparison.rs:13905`
`crates/rdocx/src/comparison.rs:13958`
`crates/rdocx/src/comparison.rs:17703`
`crates/rdocx/src/comparison.rs:13770`
`crates/rdocx/src/comparison.rs:13783`
`crates/rdocx/tests/regression_test.rs:25073`
`crates/rdocx/tests/regression_test.rs:25106`

The whole-paragraph publisher has a remaining schema-invalid sibling of the repaired ordinary-field path. wrapped_children moves a direct hyperlink outside the revision and recurses into its children, but a direct fldSimple falls through the ordinary grouping branch. The final wrap then makes that simple field a direct child of ins, del or the selected move wrapper. Enclosing complex-field begin/separate/end markers are semantic field boundaries, not XML ancestors that change this child grammar.

The existing public regression supplies the concrete valid input: a cross-paragraph complex TOC containing a direct PAGEREF fldSimple in its first paragraph. Its whole-field-inserted and whole-field-deleted rows call compare_and_resolve in both Run and Word modes. The carried-paragraph sets bypass refuse_uncarried_field_owners, which otherwise recognizes that simple fields cannot sit inside run revisions. Thus the positive deliberately reaches the problematic publisher. The existing helper checks successful public comparison, saved tracked output, modeled accept/reject and another comparison. It does not validate the tracked candidate against the schema. Its current exact256PASS cannot discharge this gap.

This grammar is independently established from the retained primary Transitional wml.xsd, SHA2560f7784b36ddc42950e5ff354cd608b83b60e4de39325d5568913f1225365029b, authenticated against the existing schema-provenance record. CT_RunTrackChange at995uses EG_ContentRunContent and math. EG_ContentRunContent at2082does not admit direct fldSimple. EG_PContent at2228separately adds fldSimple, and CT_P at2236uses that paragraph grammar. The retained Microsoft Open XML SDK3.0.1InsertedRun Children/Particle independently omit direct fldSimple and hyperlink. Its complete schema is pinned atb29e60a07afc3e0a4695eecca6c2a7a35070eb81f53a2f11033c6df6f946212b. This is a schema fact, not an inference from the Q-S96-02instruction-edit choice or a new user preference.

The finite publisher inventory is moved_paragraph at8899, final_paragraph at10565, deleted_paragraph at13771, inserted_paragraph at13784and the hyperlink recursion at13916. The final path also writes selected move wrappers. All use the same wrapping helper. The existing carried TOC fixture establishes the addition/removal reachability. Move and other composed callers share the defective publisher, but this reviewer did not execute an additional move input. No invented runtime observation is credited.

The fix must reconcile that actual publisher grammar and the existing supported positive, retaining all asserted content, fields, marks, source metadata and lifecycle semantics. Do not turn the positive into an unexplained refusal or accept a parser round trip as schema validation. Keep the finite existing case and inspect complete saved candidate and resolved parts. Historical ownership of the behavior does not make the invalid output suitable for completion. No repair is performed in this pass.

## Actual evidence and limits

Root released this report write only after the source acquisitions finished. This reviewer independently authenticated all802members of the V7execution seal and all256exact named outcomes against membership and raw outcome receipts. Every child exited0, ran its exact name, reported1passed/0failed/0ignored, and recorded no panic. Source before/after agrees. The four cohorts are lower61, layout14, facade65and regression116.

| Evidence | Actual observation |
|---|---|
| V7source binding | ab4b3dfbeb1f803010574d355ea73dc3a0f9d2d100fc9abebf34e3ba5cab1fdc |
| V7semantic receipt | b0845cc096ea42e1634c03ff5b4f700e01967d6157d6f3573ea39b5ee10a3a04 |
| V7execution seal |76a59110c2c6daa37a350ca3a307047d3288efab3ab93b2fc00dd3fde2de60b3 |
| Scoped clippy v6 | Four affected crates, all targets/features, locked/offline, warnings denied, exit0in41.83seconds, source stable |
| Fresh opt2debug wheel build | Actual maturin build exit0in205.28seconds, Cargo native and wheel member agree |
| Python installation and exact9 | Installation exited2because the UVcache path was outside writable roots. No nine-test execution or installed qualification follows. Sealed packet remains NONPASS |
| Fresh CLI exact2and sibling successor | Prepared, held pending D3, not executed |
| Schema/native coverage of D3 | Source and primary-schema finding. No reviewer-created runtime or native observation |

The new wheel is a successful build artifact, not a passed installed test suite. Its SHA256isa1fa0e6d17383268e2d0e388029d1bc7388209626bb10202d5445103caba7f18. Actual Cargo native and wheel member hash5cb5124373cf316d60d3b33561954157b8273d462d71cc009ae301c5bb47f720. No stale import is substituted after installation failure. All older RED packets remain historical, including V1throughV5semantic failures and the separate scoped lint failures.

## Prior findings and current shared invariants

D1, stable new-only CLI identity, remains addressed on examined source. The current fresh CLI follow-on is pending. The stable story coordinate and metadata multiset policy is not replaced by mutable story text identity.

D2, protected-owner source identity, remains addressed on the examined repaired paths. Both typed meaning and positive source residuals precede protected raw emission. Actual required binding checks distinguish an unchanged packed owner from a source-meaning change without treating an unused extra declaration as a changed owner. Property-role removal follows admitted current Word parents and leaves foreign wrappers, unknown Word owners, fldData, prior revisions and drawing descendants opaque. The12opaque-parent/control/format-option negatives inspect normalized inequality and public refusal, alongside compatible source-preservation positives. D3is a different physical child-grammar gap, not a reopening of the protected residual rule.

N1namespace publication and N3revision parser-context identity retain their examined repairs. The new deletion leaf mapping proves actual source/generated Word roles and renames the actual qualified root while retaining its prefix, attributes and inner bytes. Physical run grouping and source slots remain distinct from generated typed event partitions. An initially suspected raw-instruction rename rollback was not retained as a finding: all examined deletion publishers use the same shell/content source, whose complete root closure preserves scope equality. No reachable differing-role deletion counterexample was established.

The operation-local occurrence maps retain actual input role, physical part extent, owner path and field slot. Checked part-origin propagation subtracts only proved opening growth and synthetic offsets. Clone bindings validate actual Field facts/slot correspondence and roll back in reverse order on success/error before temporary owners disappear. The independent normalized package graph still examines current field result owners and relationships, rather than granting conversion witnesses the power to fabricate missing typed facts.

The lower ordered-result representation uses one current COWparagraph. Current/pending projections, normalization, borrowed revision wrappers, accepted paths and source locators remain distinct as specified. Comparison result preparation and graph identity borrow actual physical cache facts independently of legacy scalar display, while the semantic getter retains its consistency/significance policy. Qualified hyperlink results use the canonical owner. Current layout, field evaluation, relationship validation, revision listing and checked save consumers were traced. Preflight and source/copy work admission precede relevant owner copying and normalization. The lint aliases are private transparent spellings, not new nominal exported types. No additional concrete correctness, panic or structure defect was established in those paths.

## Acceptance coverage and remaining work

A1throughA7retain the bounded prior-revision preservation, new-id resolution, precise atomic refusals, comments/ignores, stable counts/live handles and checked workaround contracts. The current Rust selection exercises those bodies, but fresh CLI and installed Python execution remain pending. No broader editing of revision-bearing owners is inferred.

A8throughA14cover ordinary result edits, instruction/form dispatch, actual paired-host policy, nested revision ownership, pending/current lifecycle and resource/graph custody. Source dispatch precedes global Field equality, which continues omitting form. Required legal owner parsing, complete complex conversion witnesses and independent package postconditions were examined. D3prevents treating all affected whole-owner publishers as schema-clean even while their current modeled tests pass.

Issue430rows remain open: added/removed footnotes, note-text control, coherent ignored-note topology or atomic refusal, and the separately unreproduced following-paragraph rewrite. Published-producer reproduction is diagnostic evidence. No current fix, selected note cleanup or Word parity is claimed. The native control pipe remains unavailable before app inspection, so no current custody or native export is inferred.

Final acceptance still owes independent schema/native saved-artifact qualification, complete workspace/hosted gates, full installed Python3.9/3.12and typing/stub checks, WASM,49hashes/7pixels, affected archives/size, actual README measurements, producer oracles and all plan risk/metadata riders. F-X273/F-X274and the Issue429release/SemVer work retain their own debt. None is waived by this scoped pass.

## Smells

Zero independent smells.

## Nitpicks

Zero.

## Not found and handback

Zero additional independent correctness, contract, panic, test, structure or OOXML findings beyond D3. D3includes its missing candidate-schema discriminator rather than manufacturing a second test finding. Public comparison signatures/defaults and the Python binding publication boundary show no additional export/default mismatch in the inspected diff. No extra unsupported authoring API or new file/module/crate was introduced by this repair.

Only this report was written after the explicit release signal. No source, tests, plans, HLD, progress or other file was changed. No build, semantic diagnostic, native call, oracle, commit or push ran during review. Return control to root for a separate remediation phase with D3blocking.
