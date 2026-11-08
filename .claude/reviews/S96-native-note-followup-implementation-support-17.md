# S96 F275 A6 implementation support17

Independent bounded support for Root final union reconciliation. Outcome: two correctness defects block approval of this frozen settings repair. Zero additional smells and zero nitpicks. This is not a formal review of the completed F272/F272a provider, a peer-cap reset, a release gate or full acceptance.

## Reviewed boundary and evidence

Branch sprint/s96, HEAD e430071d75ed33e8831752d644aec3c0a865d3a3. The six scoped files and exact SHA-256 values are in BOUNDARY.json. REVIEWED_INCREMENT.patch is the exact scoped working delta against that HEAD. Unrelated workflow pacing/CI/HLD15 changes are excluded. Root remains sole source author and Cargo coordinator.

The source pins in ISSUE430_NATIVE_REFRESH_GREEN2.json match the independently hashed current comparison.rs and integration_test.rs. They are respectively 1b8e66250b20d7f2f393fdab6984b6503e20626cb9f1050bf7f81bc37e46402d and 183364d1c6afed71dac70b80cd6df34f8c3f4c83acc411465ac27fe064d1fac9. GREEN2 records actual exit0 and source_unchanged true. Its raw log line65 reports58 passed,0 failed,0 ignored. RED2 records actual exit101 and source_unchanged true with different predecessor pins. Its log lines9-12 and27-29 retain the dangling-settings and terminal-B inside-revision failures, and line78 reports56 passed,2 failed,0 ignored. These are actual filtered Native integration observations, not full-workspace or installed/native follow-up acceptance. Initial compile and intermediate settings-format failures remain in the separately hashed red/green1/check logs. Instrumented source/trace packets remain separate. The current comparison source contains no B_ONLY_TRACE.

## Defects

### D1, blocking: cleanup replaces the property root and loses unmodelled attributes and namespace context

The new cleanup invokes settings.set_footnote_properties when references are removed at crates/rdocx/src/comparison.rs:5242. That setter regenerates a new w:footnotePr root at crates/rdocx-oxml/src/settings.rs:1132 and replaces the whole original top-level element through finish_set at settings.rs:1190 and rewrite_ordered_child at settings.rs:3485. CT_NoteProperties stores typed values, special references and retained children, but no root attributes or owner declarations at crates/rdocx-oxml/src/document.rs:10383. Its parser receives no property-root attributes at settings.rs:766. The replacement writer creates a fresh root at document.rs:12385.

Therefore an admitted source such as a footnotePr with xmlns:x="urn:producer", x:flag="exact", a typed special reference and an opaque child <x:keep/> loses x:flag and the local x binding when the last note part is removed. The child bytes remain, but its expanded-name context is lost. This is newly exercised by automatic absence cleanup even though the general setter predates this increment. It violates the operative preservation contract at .claude/plans/F-X275-design.md:895 and docs/hld/04-opc-and-packaging.md:1837. The added test at crates/rdocx/tests/integration_test.rs:37909 declares producer namespaces on each retained child and gives the property root no unmodelled attribute, so it does not distinguish this failure. Source-established finding only, no counterexample execution in this pass.

Required regression discriminator: an inherited foreign child binding declared only on the actual property root, plus a producer root attribute. Verify the root metadata, child expanded names and unrelated policy remain exact after absence cleanup. This is a finding, not a patch or new permission condition.

### D2, blocking: namespace scope can come from a nested opaque settings lookalike

The property lookup at crates/rdocx/src/comparison.rs:5191 requires only a Word footnotePr whose immediate parent is a Word settings element at comparison.rs:5194. It never requires that parent to equal tree.root. XmlTree parses and retains all elements, including opaque descendants, at crates/rdocx/src/revision.rs:2241 and revision.rs:3440. The typed settings model, however, selects only depth1 properties at crates/rdocx-oxml/src/settings.rs:759.

An earlier opaque subtree may contain nested w:settings/w:footnotePr with q bound to the Word namespace, followed by the actual top-level footnotePr where q is bound to urn:producer. A nonempty <q:footnote q:id="99"></q:footnote> in the actual property is retained as opaque by document.rs:11697. The new lookup takes the earlier nested Word q scope. The raw-child classifier at comparison.rs:5228 then interprets the actual foreign q:footnote as a Word edge and deletes it. This independently violates foreign-lookalike preservation and the actual-property-owner scope contract. Fixing only root serialization does not fix this lookup.

Required regression discriminator: an opaque earlier nested settings/property pair with a conflicting alias, alongside the actual top-level property and its foreign lookalike. Confirm only actual direct Word reference roots are removed. Source-established finding only, no counterexample execution in this pass.

## Shared invariants and controls

No additional defect found in the completed B-only reconciliation. The new paragraph range is derived from the same positively admitted parsed paragraph as the marker at crates/rdocx/src/comparison.rs:4718. The ownerNone match at comparison.rs:5363 requires exact token, bits, pair id, part, full paragraph source bytes and relative marker boundary at comparison.rs:5367. It does not infer a note owner from adjacency. Range slicing/subtraction uses ranges from the current package parse. The left package stays immutable and right markers are reparsed before use after rewrites. Nonmatching collisions still enter the existing renumber path at comparison.rs:5408. Existing inside-revision refusal remains at comparison.rs:4643. The terminal B tests exercise both completed bit patterns, actual Word paragraph metadata, public scoped replacement and Word-granularity comparison at crates/rdocx/tests/integration_test.rs:37849.

The settings cleanup is reached only under the existing absence predicate after custody settlement at comparison.rs:5018 and comparison.rs:5026. It does not broaden that predicate. Comparison inputs and candidate publication remain staged at comparison.rs:1136, comparison.rs:1447 and comparison.rs:1770. Revision resolution clones a candidate at crates/rdocx/src/revision.rs:485, settles it at revision.rs:657, reopens it at revision.rs:664 and commits at revision.rs:676. Errors introduced during settings classification therefore do not publish partial package removals to the public document. No new callback, public signature, cache, type abstraction, owner inference or admission bypass was introduced. Zero additional atomicity or internal-index findings in this bounded increment.

HLD03/04/10 and the F275 final section express the intended repair and preserve the historical child status. Their preservation claims are not yet met because of D1/D2. The new typed/raw alias test and retained-part branch are meaningful but incomplete for root metadata and opaque nested-owner selection.

## Pending riders and limits

Root must retain the actual failed native G09/G12 observations. Source GREEN2 does not discharge them. After corrected source and independent review, exact actual Word pending/resolved input follow-ups, fresh ordinary installed output, complete SDK2019/2021 views, strict applicable checks, archive/hosted and final union reconciliation remain pending under the existing plan. The earlier native saved files and SDK passes belong to their authenticated producer bytes. No transfer of those passes to this Rust repair, no gate waiver, timing claim, publication or release approval is made.

This pass performed read-only source/evidence inspection and wrote only this private/tmp support packet. No remediation or additional execution was performed. End support17.
