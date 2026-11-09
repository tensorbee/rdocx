# F-X272, namespace investigation, pass 5

**Role**: independent focused reassessment after the first pass4 repair. Review-only outside this report.
**Source**: dirty `work/f-x272-codex`, Base `c8acc502b74f29ed7e9546939f17007709a9f22b`. Tracked binary diff SHA-256 `dc651f282d594c21a33aac0a19766d1547426f7652409109e15e8fc444c64030` was checked immediately before writing this report.
**Outcome**: D2 remains open at property-source composition. N1 remains open at unnecessary closure of retained run children. Two continuing findings, zero new independent findings, zero smells and zero nitpicks. This is not full microscope pass2 or an integration verdict.

## Evidence boundary

Read current source, latest progress, the current plan including its pass4 source-frame contract, and investigation pass4. Consume `fx272-pass4-native-first-attempt.log` as actual current-source execution: 26 tests ran, 24 passed and 2 failed. No tests, builds, wheels or oracles ran during this review.

The original four quoted-opening contexts passed. Canonical, alias and producer-w inserted/deleted/controlled row cases passed. These results close the particular pass4 reproductions locally. They do not prove every publisher or the complete feature ready. D1 CLI2/2 remains separately earned evidence. The historical bb501a12 wheel and native7a361346 artifact do not qualify this changed comparison source.

The property matrix actually reports acceptance reconstruction failures for left-biased final deletion, metadata-empty final deletion/insertion, and ordinary property edit. Contrasting, absent and controlled cases complete without recorded failures. The ignored-box failure actually reaches header1 item0 in its producer-host collision fixture. The successful earlier ordinary/main-box controls remain evidence.

## D2, property patching loses source order and still lacks a proven empty-owner role result

Citations: `crates/rdocx/src/comparison.rs:4530`, `crates/rdocx/src/comparison.rs:4551`, `crates/rdocx/src/comparison.rs:4580`, `crates/rdocx/src/comparison.rs:4596`, `crates/rdocx/src/comparison.rs:4600`, `crates/rdocx/src/comparison.rs:4615`. Reproduction: `crates/rdocx/tests/regression_test.rs:409884`, `crates/rdocx/tests/regression_test.rs:409913`.

`property_children` inventories only direct Word children. When a modeled child disappears, the patch deletes its span. A newly introduced child is placed before a later surviving Word key, or appended after every remaining child. Retained foreign children are absent from this insertion inventory. In the actual ordinary and left-biased fixtures, changing q:b to q:i deletes b, retains the following producer w:opaque, then appends i after that opaque child. The source child order is reversed. This also changes the opaque child's parsed property slot, so byte-preserving retention of its fragment alone is insufficient. This is a source-proven shared defect in the patch algorithm, consistent with those two actual reconstruction failures. It is not an attribute-specific defect.

The smallest correction is to apply modeled deltas against the complete source child sequence, retaining opaque children and their intended source-relative positions. Replacement of one modeled property by another needs an explicit source/schema insertion position. A missing later Word anchor is not authority to move a newly introduced property after trailing opaque children. Respect existing property schema order and raw slot conventions rather than sorting all children or comparing unknown XML semantically. Use the same correction for nested rPr and ordinary/final paragraph paths.

The raw snapshot architecture itself is appropriate. The stack is operation-local and the owned snapshots precede marker insertion at `comparison.rs:3916`, `comparison.rs:3926`, `comparison.rs:4506`, `comparison.rs:4744`. This review found no new evidence requiring a registry, public model field or parser for unknown attributes.

### Source selection and payload roles must remain distinct

Citations: `comparison.rs:4762`, `comparison.rs:4776`, `comparison.rs:4791`, `comparison.rs:4796`, `comparison.rs:6706`, `comparison.rs:6725`, `comparison.rs:6839`, `comparison.rs:6880`, `comparison.rs:6894`, `comparison.rs:6920`, `comparison.rs:6932`.

`PropertySide::matched` deliberately computes accepted raw properties by applying the supported typed delta to the original source. Unsupported opaque changes remain left-biased. Do not replace this with edited raw XML merely because BOTH input frames now exist. Its typed identity and selected raw snapshot must describe the same supported projection when used by final merge decisions.

Current prior-source substitution already distinguishes pPr from other owners at `comparison.rs:4634`: it strips direct rPr, sectPr and pPrChange for a prior base pPr, and uses the complete source for rPr. Do not report that this implementation blindly copies a complete pPr as the base payload. However, a complete source owner is not a universal replacement template for every prior property role. Preserve root QName, effective bindings and opaque metadata while applying the role's actual child selection. A prior base pPr excludes the independently resolved mark and section. A prior paragraph-mark rPr must represent the rejected mark formatting and cooperate with retained mark markers. Own-section precedence remains explicit at `comparison.rs:6903`.

The resolver establishes these contracts at `crates/rdocx/src/revision.rs:1084` and `crates/rdocx/src/revision.rs:1098`. Rejected base properties retain independently resolved mark/section children. A rejected mark change replaces mark formatting, including unmodeled children, while preserving its revision markers. Legacy payloads containing a mark have a separate path at `revision.rs:1020`. A repair must not accidentally manufacture that legacy payload shape or duplicate opaque children across prior and retained roles.

The two metadata-empty final failures are actual unresolved D2 evidence, but this receipt contains no resolved XML or first differing signature. The ordering defect above does not establish their exact dynamic cause. Here the selected pPr contains metadata but no modeled children, and the opposite final owner has real base/mark formatting. The current branch compares optional raw XML and constructs empty fallback owners at `comparison.rs:6894`, `comparison.rs:6921`, `comparison.rs:6938`. Namespace/attribute-bearing empty shells also differ from removable empty shells in the resolver at `revision.rs:773`, `revision.rs:794`, `revision.rs:973`. These are the concrete source/role boundaries to instrument, not evidence that a namespace-only shell or the metadata attribute itself is the cause. Preserve absent versus metadata-bearing empty source states. Do not change that policy to make the test pass.

Before another speculative empty-owner patch, capture the first differing actual resolved property projection for these two existing cases in a bounded diagnostic. Compare the requested supported projection, selected source snapshot, emitted prior base/mark payload, and accept/reject result. That will establish which role is wrong. Keep this diagnostic separate from evidence that the ordering correction fixes the ordinary and left-biased cases.

## N1, retained run children are closed as though every child relocated

Citations: `crates/rdocx/src/comparison.rs:4657`, `crates/rdocx/src/comparison.rs:4680`, `crates/rdocx/src/comparison.rs:4698`, `crates/rdocx/src/comparison.rs:3106`, `crates/rdocx/src/comparison.rs:3141`, `crates/rdocx/src/comparison.rs:8494`. Reproduction and exact bytes: `crates/rdocx/tests/regression_test.rs:34759`, `crates/rdocx/tests/regression_test.rs:34778`, `crates/rdocx/tests/regression_test.rs:34783`.

`preserve_run_source` retains the actual source run shell, but sends every non-rPr child from generated serialization through `source_child`. That helper closes a child under its scanned scope. It is valid for a fragment actually detached from that scope. It is not byte-neutral for a retained opaque child remaining under an equivalent source owner. In the collision fixture the opaque x:textBoxHost has its own x binding and exact bytes, while Word scope comes from its run ancestry. This path can add a Word declaration to the opaque root unnecessarily. The fixture explicitly requires the original producer fragment once, before and after acceptance.

The signature does not ignore this difference: raw run payloads and their positions remain part of `semantic_run_raw` at `comparison.rs:9380`, and paragraph tokens preserve boundary identity at `comparison.rs:9768`. The dynamic header reconstruction failure is consistent with this concrete publication defect. The log alone does not expose the final changed fragment, so it does not independently prove the declaration delta was the sole mismatch. Do not weaken these signatures, skip the collision fixture, or classify its producer host as an internal marker to suppress the failure.

Keep the retained-versus-fresh distinction inside the existing private publication architecture. Fresh canonical Word children need the existing generated qualifier. Actual retained children should retain raw bytes when their effective destination context agrees. Actual relocation needs source-context closure, including local shadows and value-only namespace dependencies. Retaining the run shell is useful evidence of context, but not permission to skip closure for a child genuinely coming from another context. The publishers at `comparison.rs:7103` and `comparison.rs:8494` must share this distinction. Do not globally rebind producer w or close all unknown children preemptively.

## Bounded proving cases and handback

Keep all 26 current named controls and the separate D1 CLI controls. The next correction should earn the existing matrix GREEN, including both metadata-empty directions, without deleting the actual RED receipt. Preserve exact q root identity, producer-w expanded attribute/element names, quoted greater-than values, opaque child order and parsed raw slots, intended modeled formatting, own-section precedence, old revision bytes, author-selective accept/reject and save/reopen.

For the ordering correction, retain b-to-i with trailing opaque content and include one insertion before an existing later Word property so both insertion branches are discriminated. For payload roles, keep a real nonempty mark, base-only prior pPr, metadata-bearing empty and absent pPr, and the existing controlled context. Check current and prior snapshots separately rather than substituting a complete source frame into every slot. Retain the legacy mark and property-rejection controls already passing in this batch.

For N1, keep the exact header producer collision plus w:tab, earlier ordinary/main-box cases and a genuine cross-context relocation control. Verify exact retained producer bytes under unchanged context and expanded namespace identity after relocation. The successful foreign-w row matrix must remain green. No new row or owner-edit capability is authorized.

No new public model/API, trait, module, source file or unknown-attribute parser is justified. No new product choice is pending. D1 and N3 are not reopened. Full independent feature microscope, fresh authenticated successor Python qualification, and the final integrated native/hosted/installed/typing/WASM/hash/pixel/archive/README/risk gates remain pending.

Only this report was written. Return control to root for a distinct implementation phase. D2 and N1 remain unresolved.
