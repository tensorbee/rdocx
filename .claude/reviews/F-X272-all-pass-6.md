# F-X272, all aspects, pass 6

**Reviewer**: independent Codex reviewer, not an implementer of this repair.
**Scope**: D3same-form whole-paragraph repair checkpoint against actual pre-lease copies in fx272-d3-implementation-v1/before. Three changed source/test files,769added and97removed lines. Comparison SHA2562a6bafeaba9005983d55069392f5036d854fa8f25796a66396c9ec5c052ac8b7, revision SHA256d2aa2d0517be264d550685c601648fa48270db2291c3952fe7f3868a09547e56, regression SHA25611392a915d2ba6a6276cb924c50f6d6a826ca17aaea5e7c04bc8f488f05ad2b6. Current bound source was independently checked against the focused receipt before writing. F-X273prefix and Issue430implementation remain outside this repair.
**Contract**: current approved D3same-form, per-occurrence selected custody, source-qualified hyperlink projection and iterative-work amendments, canonical microscope/workflow and pass5. Main and shared related/transport resolver entry points are included. This is a remediation review, not overall feature or sprint acceptance.
**Verdict**:2defects,0smells,0nitpicks. D3remains blocking. D4identifies a distinct opaque-custody regression in the new publisher. No clean readiness attestation.

## Defects

### D3, the new publisher renames only the opening of paired deletion leaves

`crates/rdocx/src/revision.rs:1006`
`crates/rdocx/src/revision.rs:1010`
`crates/rdocx/src/revision.rs:2839`
`crates/rdocx/src/revision.rs:3111`
`crates/rdocx/src/comparison.rs:13771`

tracked_paragraph_children supplies an entire actual paired t or instrText leaf to rename_element. That existing helper replaces only the first occurrence of the qualified name. It serves opening/closing fragments in the ordinary renderer and does not rename both ends of a whole element. The new call therefore emits an opening delText or delInstrText with the original closing t or instrText.

This is actual failed successor evidence, not a hypothetical schema concern. The new revision unit exited101at3111with expected closing w:delText but actual w:t. The retained whole-field paragraph regression separately exited101with expected closing w:delInstrText but actual w:instrText. Both test bodies stopped at those errors, so later assertions in them are unexecuted. The first focused repair result is13PASS/2NONPASS, not GREEN.

All moved, final, inserted/deleted and nested hyperlink paragraph routes now share this publisher. The intended shell-outside/run-inside representation addresses the original direct fldSimple ancestry, but paired deletion source must also remain well formed. Keep the actual qualified prefix, attributes and inner bytes while applying the complete leaf transformation. Do not globally rename arbitrary subtree strings or change the old helper's fragment callers without reconciling their actual input shape. Refresh the affected source-bound controls and complete saved-part schema evidence after repair. D3does not close on a successful build or on the other13passing bodies.

### D4, the D3publisher descends through an opaque drawing into textbox text

`crates/rdocx/src/revision.rs:979`
`crates/rdocx/src/revision.rs:994`
`crates/rdocx/src/revision.rs:2418`
`crates/rdocx/src/revision.rs:889`
`crates/rdocx/src/comparison.rs:13770`
`.claude/plans/F-X272-design.md:2510`

The deletion scan iterates every descendant of an ordinary raw child and qualifies text by its parent being a modeled Word r. It does not use the local admitted inline-source route computed above it. The shared XmlTree deliberately re-enters modeled ownership at any qualified txbxContent, because the existing all-story source scanner must recognize actual text boxes even below a preserved drawing host.

Consequently a deleted paragraph containing r/pict/foreign-shape/textbox/txbxContent/p/r/t reaches that nested t in the new scan. Its parent r becomes modeled after the textbox re-entry, so the publisher renames the nested textbox payload while wrapping the outer run. The same applies to nested instruction leaves. This crosses the D3source ownership boundary and changes retained opaque drawing content, rather than transforming only eligible leaves of the current paragraph-result owner. The D3contract expressly excludes opaque/drawing descendant descent. A separate textbox story becoming recognizable to the shared scanner does not grant the enclosing paragraph publisher ownership of its payload.

This finding is source-established. No new textbox runtime probe was executed by the reviewer. D3's paired-name error can currently make the transformed leaf malformed as well. Even after that paired rename is corrected, the ownership violation remains. No failed rejection or native behavior is invented: the concrete effect is the new publisher changing opaque-host source bytes before resolution.

Preserve the shared scanner's legitimate textbox/listing domain while enforcing the D3publisher's actual local grammar and custody boundary. Inspect the complete current subtree/publisher inventory, including the source carrier inventory that also consumes the shared tree, rather than broadening a global modeled flag or repairing only one t spelling. A potential cross-domain carrier association was observed at the global active-field lookup, but no separate downstream refusal/retirement outcome was established in this pass, so it is not manufactured as a third finding.

## Actual focused evidence

Independently authenticated the focused receipt against all15raw exact child records/log hashes, both original and frozen executables and actual Cargo opt1/debug-assertion/overflow profiles. Every selected child started its exact named test. Source before/after agrees, and the current tracked inputs still matched that frozen source before this report.

| Observation | Actual result |
|---|---|
| Focused acquisition |13PASS/2NONPASS, no setup/source invalidation |
| Failed unit | revision::carried_simple_fields_preserve_schema_and_selected_source_custody, exit101 |
| Failed public regression | comparison_carries_bookmarks_and_whole_fields_in_whole_paragraph_changes, exit101 |
| Focused receipt SHA256 |b93a59afbfa38b2beef7b3bca82b9572d27c20d3435cfd076b4b5797d681ae66 |
| Root raw authentication | fx272-d3-focused-root-review-v1.json, SHA256477f186c9c488e543938de2c70860e0ebf66d1b4eb2171723374489121fccf7f |

Successful ordinary field conversion/form/source controls, hyperlink/control consumers and inherited drawing namespace controls remain actual supplied-input observations. They do not execute the failed bodies' later branches or prove the missing opaque textbox boundary. No scoped CLI, installed Python, schema-success or native result follows this NONPASSbatch.

## Other examined invariants

The new private renderer preparation is called at both public main-document resolution and render_with_source_ranges. Retirement invokes selected-descendant validation/counting before a simple shell disappears. Material and marker-removal facts use actual scope.matches/removes_content policy, preserving cache-only empty simple shells and refusing supplied partial-marker and unselected-custody cases. Qualified source listing can enter simple fields under real hyperlink grammar without promoting the lower opaque hyperlink payload into a borrowed typed Field owner. Unknown and prior-property barriers remain in modeled_child, with D4showing why the new publisher still needs its narrower local ownership predicate.

The publisher uses explicit frames and a final output buffer rather than copying each nested output through every ancestor. The source tree, marker/domain inventory, prefix material budget and bottom-up survival calculation are concrete private carriers. No new public API/model or general persisted exchange mechanism is added. These observations are source review, not a claim that every later branch in the two failed bodies passed.

D1stable CLI identity, D2protected residual/source identity and N3parser-context identity are not newly reopened by this finite D3review. D4concerns a new whole-paragraph ownership regression, distinct from their earlier supplied contexts. The existing current-source consumer/native/full-union debt remains.

## Smells

Zero independent smells.

## Nitpicks

Zero.

## Not found and handback

No additional independent correctness, contract, panic, test, structure or OOXML defect beyond D3andD4was established in this repair scope. The missing opaque discriminator belongs to D4rather than a duplicate test finding. Issue430remains separately draft/unimplemented, and F-X273main K1/R4/cache integration and every final schema/native/workspace/installed/hosted/hash/risk gate remain pending.

Only this review file was written after the worker stopped and root released report writes. No source, tests, plans, HLD, progress, TEMP or other file was modified. No build, semantic diagnostic, native call, oracle, commit or push ran during this review. Return control to root for a distinct remediation phase. Source-changing fixes must not overwrite the preserved failed packet or turn this report into an attestation on later source.
