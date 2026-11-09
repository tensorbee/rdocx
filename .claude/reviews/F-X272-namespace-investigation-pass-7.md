# F-X272, namespace investigation, pass 7

**Role**: independent focused design reassessment after protected-source remediation. Review-only outside this report.
**Source**: dirty `work/f-x272-codex`, Base `c8acc502b74f29ed7e9546939f17007709a9f22b`. Tracked binary diff SHA-256 `1bf9d8d7c50359eae8e7e1eae31067ddcca29f5cb5d6da30b0a6c01ccb775cb9` checked immediately before writing.
**Outcome**: D2 remains open. The current supplement catches meaningful omitted source state, but also rejects an actually supported known-format canonicalization. Zero new independent smells or nitpicks. This is not full microscope pass3 or feature readiness.

## Actual evidence

Read latest actor progress, current protected-source implementation and the annotated contract. Consumed `fx272-protected-known-canonicalization-current.log` and `fx272-known-format-canonicalization-predecessor.json`. No tests, builds, source edits or artifacts were produced by this review.

The predecessor receipt authenticates native b301c71fddfd1c62ed8b8ede223594739bc8c7961effe2087ac8de5e5fd6d1c5. A paragraph holds old insertion7/A, paragraph spacing before=00120, mark bold val=true, and ordinary run bold val=true. Public add_comment targets the OTHER paragraph, then save/reopen rewrites spacing to120 and both bold elements to bare b. The old insertion bytes remain unchanged. That authenticated predecessor comparison succeeds and carries the comment.

The current source test at `crates/rdocx/tests/regression_test.rs:410371` checks that actual public route and the lexical normalization before comparing. `fx272-protected-known-canonicalization-current.log` executed1test after a successful70secondbuild. It fails with the location-qualified protected paragraph refusal, 0passed1failed,2496filtered,2.97seconds. Its17protected-source refusal contexts and the existing declaration-order/redundant-closure, ordinary-left-bias and ignored-header positives proceed without recorded failure. Its subsequent comment resolution and save/reopen assertions are not executed because the positive comparison refuses. Do not credit those later assertions as current GREEN.

## D2, all-source-attribute identity is stronger than unchanged modeled meaning

`crates/rdocx/src/comparison.rs:2805`
`crates/rdocx/src/comparison.rs:2846`
`crates/rdocx/src/comparison.rs:2886`
`crates/rdocx/src/comparison.rs:4372`
`crates/rdocx/tests/regression_test.rs:410371`

`protected_source_metadata` records every non-declaration attribute's encoded and decoded value. It also records all property nodes at raw element ordinal paths, including nodes with no source-only metadata. The guard requires this supplement to match alongside typed equality. Consequently true versus omitted val and00120versus120 differ even though the typed known formatting is equal and this is a real ordinary-save normalization. Default known container/node serialization can similarly affect this inventory. Those are source counterfactual risks beyond the actually executed bool/numeric case, not additional runtime observations.

The supplement's name says it supplies metadata absent from typed owners, but its current output includes metadata already represented by typed values. Its successful17refusals do not justify rejecting the compatible comment route. Removing the supplement would reintroduce full microscope pass2's meaningful metadata bypass. Ignoring every Word attribute would instead miss unknown Word metadata on modeled elements. Neither is a coherent correction.

## Smallest coherent private invariant

Require both existing modeled meaning and positively preserved source residual to agree for a protected owner. The residual is what that owner's current parser/writer does NOT represent as modeled semantics. It is not every source attribute, and it is not the result of serializing away unknown data.

Known supported attributes may use their existing typed normalization, including boolean spellings, integer spelling, accepted default/container serialization and prefix aliases. Opaque root metadata, unknown attributes INCLUDING attributes in the Word namespace on a known Word node, opaque payloads/events, their owner association and meaningful namespace context remain positively checked. Preserve actual original raw owners on successful publication, so permitting a known lexical normalization in the edited input does not rewrite the original protected owner or old revision. This is consistent with the existing raw-emission boundary at `comparison.rs:4384` and the contract's bounded unchanged-owner policy.

The concrete carriers already exist: HeldRevisions.source_metadata/section_source at `comparison.rs:2740`, physical source spans at `comparison.rs:2752`, actual source namespace closure at `comparison.rs:2778`, and typed BodyContent identity at `comparison.rs:4384`. Keep those scopes and ignored-story boundaries. Replace the supplement's all-attribute rule with an authority-backed source residual, before the held-owner fast path. Ordinary non-held owners keep their existing left bias.

Residual association should use physical owner and recognized property slot/occurrence, retaining opaque multiplicity/order. A raw ordinal for EVERY known-only node is not stable authority when the writer can omit a default node. Existing slot/carrier machinery is relevant here. Do not allow a new default-node reconciliation to erase attributed or nonempty source owners. Whole-table/block-control owners require the same residual recursively, not a paragraph-only exception. Standalone revised sectPr and Empty pgMar remain included.

## Existing authorities and their actual limits

| Existing carrier/helper | Useful authority | Limit that must remain explicit |
|---|---|---|
| `rdocx-oxml/src/numbering.rs:1031`, property_attributes | Existing modeled attribute sets include paragraph spacing before/after/line and run b val. Unknown attributes in Word or producer namespaces remain distinguishable. | Not an exhaustive schema for every parser or owner family. Do not copy the list into comparison.rs as a second authority. |
| `numbering.rs:1137`, projected_word_start | Resolves actual Word/relationship namespace identity before classifying an attribute. Root attributes and unsupported attributes are producer state. | Produces a projected property start, not the complete exact opaque residual needed by comparison. Its has_producer flag alone cannot distinguish different opaque values. |
| `numbering.rs:1638`, merge_property_child_attributes, and `:1734`, `:1806`, `:1976` consumers | Already uses that attribute authority to keep actual producer attributes beside generated modeled properties. | Preserved values and serialization serve overlays. Do not repurpose their output as proof that opaque original bytes can be discarded. Changing the shared table can affect existing numbering overlay behavior. |
| `numbering.rs:179`, PropertyKind, with PPR_MODELLED_CHILDREN/RPR_MODELLED_CHILDREN at `:102`, `:121` | Concrete paragraph/run distinction and supported numbering-overlay child slots. | Only Paragraph/Run, with subset lists. These do not classify tblPr/trPr/tcPr/sdtPr, arbitrary content attributes or all full-parser paragraph/run children. |
| `numbering.rs:1108`, nested_property_children | Existing numPr, border, tabs, rPr, selected sectPr and cols relationships. | Selected subset, not authority to drop every unlisted Word child or attribute. |
| `numbering.rs:1501`, PropertyOverlay | Source start, modeled child slots and raw extra slots are concrete reusable concepts. | Its reader trims text at `:1530` and ignores non-element events in the overlay dispatch. Reusing it wholesale would lose the text/comment/PI/CDATA requirements already established for this feature. Keep the complete-event source scanner. |
| `numbering.rs:668`, typed_leaf_requires_raw, and property raw-carrier readers | Distinguish a plain typed leaf from one requiring raw preservation. Existing parser dispatch and carrier positions resolve partial modeling. | Crate-private and context-specific. A checked XML validator proves accepted grammar, not necessarily that every accepted attribute is modeled. |
| `rdocx/src/document.rs:10491`, scoped_children_correspond, and `:10622`, normalize_scoped_parsed_child | Existing actual-source/parser/writer correspondence is explicitly scoped to real source bytes and writer frames, with normalization separated from source authority. | Canonical writer equality alone can hide omitted metadata. Do not replace the residual with canonical serialization. Do not repurpose its caches or source AST as a global comparison registry. |

The existing property attribute authority is a strong starting point for the exact current failure. It should be reused at its owning layer rather than reconstructed as a broad local-name list. Coverage must be reconciled against actual current parser/writer support, including attributes partly modeled beside raw carriers. This is not permission to broaden numbering serialization semantics incidentally. A conservative residual must continue retaining unknown cases until actual modeled ownership is established.

The required all-family consumer inventory is concrete: pPr/rPr roots and children, ordinary run properties, nested cell/table/control paragraphs, tblPr/trPr/tcPr/tblGrid, sdtPr, revised sectPr/pgMar, and source content nodes such as w:t. The current17cases prove meaningful omitted edits at those reached boundaries. A partial paragraph/run normalization exception must not silently remove their coverage. Foreign lookalikes must remain opaque, and value-only prefix dependencies must retain their actual context.

## Cross-crate access and technical contract clarification

`PropertyKind`, property_attributes, projected_word_start and PropertyOverlay are private inside rdocx-oxml. Most related parser classifiers are pub(crate), which still does not make them callable from the separate rdocx crate. A recommendation to call them directly from comparison.rs would not compile. No existing complete public residual-classification entrypoint was found in this inspection.

The product policy is unchanged, so no new user product choice is needed. A technical plan clarification is sufficient to name the modeled-versus-residual invariant, actual coverage and the selected ownership boundary. If the implementation needs an internal cross-crate bridge in an existing file, explicitly reconcile its visibility with the current no-new-public-API rule before coding. A doc-hidden public export is still a new exported symbol and cannot silently be called private. This report does not authorize it. Keep any necessary bridge concrete and read-only, with no public authoring operation, public model field/type, new source file, trait or generic framework. Alternatively show how the same authority stays at an already authorized callable boundary. Do not claim the cross-crate access issue is already solved.

## Bounded proving cases

Retain the exact authenticated predecessor and current failing public comment route. Its corrected GREEN must preserve the original00120and explicit true spellings in the emitted protected owner, exact old revision bytes, the actual new comment graph and selective accept/reject/save/reopen behavior. Add independent known numeric/toggle variants through existing callers rather than a special-case true/00120 rewrite.

Pair this positive with a mixed node: b val=true plus an unknown Word attribute and a producer attribute. The known value may canonicalize, but changing either opaque attribute still refuses. Include the analogous spacing child, property root and existing w:t metadata negative. This distinguishes actual modeled ownership from blanket Word exclusion. Keep encoded opaque value and changed namespace/value-only-binding negatives.

Retain all17protected-source contexts and existing comment/PI refusal controls, not a reduced replacement test. Preserve benign reordered/redundant namespaces, ordinary non-held left bias and ignored-story positives. Include a default/Empty/paired representation positive with no opaque loss, alongside metadata-bearing and nonempty default-owner negatives. Keep intentional duplicate opaque siblings, raw event order and nested owner association.

After the shared correction, refresh applicable scoped native/CLI/Python evidence and obtain full independent microscope pass3. Historical b301 successes are predecessor evidence only. D1/N1/N3 are not reopened by this diagnosis. Final integrated native/hosted/full installed/typing/WASM/hash/pixel/archive/README/oracle/risk riders remain pending.

Only this report was written. No source/test/build mutation or remediation occurred. Return control to root for a distinct implementation phase with D2 unresolved.
