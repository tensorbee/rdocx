# S96 Native story frame names, design investigation 16

## Disposition

The four-name Cow storage policy is technically feasible for XmlElementFrame and all ten actual constructors. A single substantive private same-file mapping helper is justified by those ten real consumers. No Source, private-reader, lifetime, owned-sink or structural blocker was found for the exact finite contract below.

Actual current predecessor construction/copy observations remain required before adoption. No production change is recommended until Root retains those observations and explicitly adopts the contract/HLD scope. The existing certificate facts are not evidence for this allocation seam. This report predicts neither elapsed sufficiency nor a fix for the small actual Document.add_picture miss or remaining reply failure.

Zero existing correctness defects, zero smells and zero nitpicks were established. Only this report was written. No source mutation, product import, compiler, test, runtime, diagnostic or benchmark was executed. A failed source-reading script was corrected in memory and created no file or product state. Root remains sole product/Cargo writer. Pass numbering, ordinary failures and final riders are unchanged.

## Immutable source and complete field scope

The reviewed source is `.claude/scratch/S96-final-evidence/f268-direct-cell-value-implementation-20261108-1/before/crates/rdocx/src/document.rs`, independently matching SHA-25688abd485c574432bdc296a4c3ace85973c42c1061be157270df2a6a9a0c40bd2. Native document.rs line citations below refer to this snapshot. Root's concurrent certificate edits do not define the frame policy audited here.

XmlElementFrame at document.rs:13114has a Vec<u8> local_name at13116. It is private, has no Clone implementation and appears only in this file. Exactly ten actual push constructors exist. The frame also owns namespace/is_word, offsets, owner/item kind, direct-owner and SDT context, shared prefix scope, opaque and typed-content state. All of those fields and their producers remain unchanged.

The name's only ongoing consumers are byte equality, slice matching and helpers accepting &[u8]. The field is never mutated, hashed as an identity, returned as an owned name, installed in a model/cache, or used for its capacity/allocation pointer. Frame pop moves other metadata into owner/item/link/text results, then drops the name. There is no shared-reference semantic mutation or custom destructor to preserve in that Vec. Unrelated local_name fields, tuple stacks and retained owner names elsewhere in the file are outside this proposal. In particular, the rebased-owner pointer test at56981concerns modeled owner names, not this frame.

## Exhaustive constructor and control matrix

The selection/projection column is part of the contract. Do not substitute a new QName/local-name extractor merely because its normal output looks equal.

| Actual consumer | Old name selection/allocation and event scope | Required local control |
|---|---|---|
| scoped_alternate_text_box_copies17690, constructor17753 | reader.resolver().resolve_element selected local, borrowed during the event. local.to_vec only at non-Empty push17755. | Paired MC Choice/Fallback, Word drawing/pict and VML text bodies, qualified versus opaque descendants, Empty/expanded branch forms and ranges. |
| paired_text_box_projection17790, constructor17967 | Same resolver-selected local. Copy only at non-Empty push17969, after existing projection/qualification checks. | Actual paired signature/slot projection, start range activation, duplicate/opaque/nested branches, actual namespace scope and errors. |
| template_direct_table_children18014, constructor18065 | Resolver-selected local. Copy only at non-Empty push18067. | Direct rows versus nested table/control/preserved children, qualified parent table, Empty/expanded rows and exact output ranges/flags. |
| scan_story_owners_with_special19078, constructor19144 | element.local_name().as_ref().to_vec at19114in Start before opaque/owner/note/prefix checks, then moved into frame. | Complete sorted owner kind/index/range/fingerprint, special-note true/false, editable separator notes, cells/textboxes, opaque/foreign cases and unmatched/malformed errors. |
| scan_story_control_paragraphs19993, constructor20085 | element local copy20033for both Start and Empty before selected-root/opaque/direct-content checks. Only non-Empty moves it into a frame. | Qualified selected control, direct versus nested control paragraphs, SDT content rules, Empty paths, namespace closure and exact root/refusal behavior. |
| scan_story_items20142, constructor20297 | element local copy20210in Start before SDT/opaque/nested-owner/item/field classification, then moved. | Complete item kinds/ranges/scan/direct-owner/complex ancestors, table/textbox boundaries, controls/revisions/fields and Start/Empty variants. |
| scan_story_item_links20490, constructor20556 | element local copy20522in Start before opaque/nested-owner/link/prefix work, then moved. | Actual hyperlink relationship/anchor/ranges, nested owner exclusions, imported aliases, opaque controls and exact attribute/errors. |
| bookmark_marker_spans20945, constructor21027 | element local copy20960for both Start and Empty before modeled/visibility/marker validation. Only non-Empty moves it into a frame. | Marker start/end/ID/range, Word versus foreign, duplicate owned IDs, child payload refusal, visible versus deleted/moved text and expanded markers. |
| story_item_text21233, constructor21388 | element local copy21293in Start before qualified control/opaque/field/text/prefix work, then moved. | Accepted text and field context, controls admitted from actual inventory, nested owners, revision visibility, run versus property tabs/breaks, symbols and exact decode errors. |
| replace_story_item_text21657, constructor21779 | element local copy21709in Start after the existing root qualified-name UTF8/prefix extraction, before opaque/field/removal work, then moved. | Actual replacement XML, root prefix, preservation/removal ranges, fields/tabs/breaks/controls/revisions, invalid root UTF8 and refusal atomicity. |

Three constructors copy only at actual push, five allocate an eager temporary in Start, and two allocate an eager temporary in Start/Empty. That distinction affects actual materialization counts. Do not describe every removed temporary as a removed frame allocation or pretend an Empty frame was popped.

No current owned-name output sink was found for these ten selected values. If implementation reveals one, it needs explicit into_owned at that real boundary and its actual work must remain visible. Do not add blanket into_owned conversions at borrowed byte-comparison callers or duplicate an unknown-name Vec that was previously moved once.

## Minimal representation and helper

Change only this field to std::borrow::Cow<'static,[u8]>. A private concrete helper adjacent to the frame can own the four-name policy:

```rust
fn story_frame_local_name(value: &[u8]) -> Cow<'static, [u8]> {
    match value {
        b"p" => Cow::Borrowed(b"p".as_slice()),
        b"r" => Cow::Borrowed(b"r".as_slice()),
        b"rPr" => Cow::Borrowed(b"rPr".as_slice()),
        b"t" => Cow::Borrowed(b"t".as_slice()),
        _ => Cow::Owned(value.to_vec()),
    }
}
```

Cow is already used in this file. The helper implements actual storage selection and the real owned fallback, reducing ten copies of that policy to one. It is not a forwarding-only wrapper. It needs no struct lifetime parameter, generic, trait, module, file, feature flag, public API, dynamic dispatch or speculative argument. Both result forms have current consumers.

For the first three constructors, call it on the original resolver-selected local exactly where local.to_vec occurred. For the seven eager local variables, replace the original `.to_vec()` expression with this helper at that original point and carry the Cow through the same local decisions into the frame. Do not leave the eager Vec and invoke the helper later at push, which would retain common-name copies and potentially duplicate unknown-name copies. Do not move those eager projections after guards or suppress Empty temporaries by changing branch structure.

All four Borrowed arms reference static literal slices. Unknown/Unicode/malformed/longer/differently cased bytes retain exactly one owned byte copy, never an incoming Reader buffer or caller reference. A foreign p or r can use static storage without acquiring Word namespace or owner classification. Empty selected bytes retain the original empty owned-value behavior. There is no UTF8 conversion, case folding, prefix normalization, interning or expanded static dictionary.

Convert frame/Cow byte comparisons and patterns mechanically to as_ref where trait resolution requires it. Preserve every match arm, Boolean grouping, short-circuit and helper argument. Borrowed Vec callers already accepted slices, so std Cow dereference or explicit as_ref supplies the same bytes. Do not change unrelated event LocalName views, retained String/Arc names or owned tuple stacks.

The core invariant is that each converted local and stored frame exposes exactly the bytes of the old projection at every use and after event-buffer reuse. Allocation ownership alone changes. The resolver-selected versus element.local_name selection remains caller-owned. In replace_story_item_text, preserve qualified-root UTF8/prefix validation before the selected local at21709. No new fallback substitutes for that error.

## Shared readers, fields and first causes

The principal shared field consumers are story_sdt_context13340, story_element_sdt_context13358and modeled_story_child13393. They combine exact name bytes with independent namespace/context state. Their MC/VML/drawing/WordGroup branches remain identical. admit_story_sdt_content_slot17544mutates only typed_sdt_content_seen and keeps Empty/duplicate-content refusal. opaque_story_event/modeled_story_event keep actual namespace and payload validation.

Run/Field readers also consult these frames: story_enclosing_run_source18800, complex_field_scan_context18815, push_complex_field_instruction18885, accepted_story_text_visible19039and story_run_child21548. They still distinguish run content from properties, simple versus complex Field contexts, deleted/moved content and actual run source boundaries. No instruction/result/private Field read is removed or inferred from a static name.

Keep all actual NsReader/Reader events, resolver calls, attribute iteration/normalization, captured scope, namespace declarations, source bytes/offsets, fingerprints, SDT admission, work/resource guards and errors in their current schedule. Name comparison has no Source allowance, input receipt or owner identity. The helper cannot certify a subtree, change parser role or select output. Removing a pure temporary Vec copy does not authorize an XML scan, Field read, qualification or validation omission.

The original Vec owns only bytes and has no fallible result-producing destructor. New static storage adds no semantic error path. Unknown names still copy as before. Preserve malformed QName/attribute/entity/namespace/closing/EOF first causes and publication atomicity through actual old/new outcomes. Any unrelated scanner simplification, fallback sharing, clone removal or cache policy is outside this increment.

## Required actual predecessor work and parity

Before adoption, Root must retain a source-bound actual current predecessor with positive real name materializations/copied payload in the intended public states. Instrument the ten literal old copy sites, including the eager temporary sites, not every XML event or an always-incremented frame marker. Label consumer/copy point and selected name bucket. Count values separately from actual UTF8/raw byte payload lengths and actual pushed frames. Empty values count materializations with0payload, without a malloc claim.

The successor meter observes the helper's real Borrowed arms and Owned construction, preserving the actual calling-site/consumer distinction. Common-name old owned values must be positive and replaced by corresponding new static views. Unknown-name values/copied bytes, total selected projections and actual push/pop sequences remain equal. A new result constructed from an already allocated Vec is not an elision. All persistent output names, prefix scopes, buffers, stack storage, strings/signatures and any real owned sinks remain outside the removed-work claim.

Use the exact current public reply states and picture states with original setup/prior actions, without added getters between them. Include Document picture action1andaction2at both sizes as previously required, and the passing Cell control. The current table-certificate candidate or earlier OXML frame observations cannot establish Native frame work. If the certificate change composes first, name that actual predecessor source and observe this seam on it rather than relabelling older invocation counts.

Full public semantics, held/fresh/refusal behavior and reopen precede work assertions. Same original input ZIPs must produce identical whole ZIPs and every member. Count actual events/reader attempts/consumed bytes and actual attribute/capture/qualification work at the real owners before claiming equality. Existing debug reference work must be explicitly identified and cannot be silently excluded from actual diagnostic totals. None of these observations is elapsed, total allocation or RSS evidence.

## Focused correctness and scoped gates

Exercise every matrix row through actual consumers, with explicit unavailable-path disclosures rather than assuming one public action reaches all ten. The output comparison should include complete owner/item/link/marker/range/text/replacement results and exact typed errors. Reuse existing applicable native controls, adding focused discriminators in the existing test module rather than a new test binary.

Common cross-row cases are zero/one/many/deep frames, mixed Start/Empty, expanded equivalents, long intervening text and buffer growth/reuse, all four exact common spellings, longer/case-different/Unicode/unknown names, aliases/default/shadowed/foreign namespaces, and malformed qualified names. A cfgtest actual selected-name/push/pop trace can prove byte stability, while remaining separate from performance work counts. Do not merely test that a helper returns the constants it was written to return.

Include block/table/row/inline SDT contexts, duplicate/empty content slots, nested and opaque controls, MC first Choice/Fallback, VML/textbox/anchor versus opaque inline shape, revisions and special notes. Field and hyperlink cases must preserve physical parent/run/ancestor/range values and errors, not only visible text. Distinguish run tabs/breaks from tab stops in properties. Bookmark cases must retain duplicate-owned-ID and expanded-child-payload refusal. Replacement cases must compare full unchanged output on failure.

After implementation run affected Native and binding all-target checks and strict all-target/all-feature lint, scoped actual consumer/risk controls, fresh authenticated uninstrumented installed controls and cheap gates. No unchanged full suite is requested. Obtain independent implementation and Root reviews over the final composed union before an ordinary final-source qualification. HLD impact is plausibly03-architecture.mdand04-opc-and-packaging.md, naming pure Native scanner storage and unchanged Source authority. HLD10or API changes are not earned by this private field alone.

These are temporary scanner frames. Cow may alter frame size/alignment, stack Vec capacity costs and transient peak behavior. Do not infer total heap or RSS savings from copied-byte counters, add a retained-cache omission mode for this temporary field, or treat static storage as zero process cost. Final ordinary/diagnostic memory calibration and existing lifetime/RSS riders consume the actual final representation with no new ceiling or policy relaxation.

The source-level contract is closed. Actual current old-positive work, explicit adoption, implementation correctness/mechanism evidence and ordinary elapsed sufficiency remain separate pending steps. No Source/finitelanguage/cache/AST/history/public API/workload/threshold change, state/ledger update, commit, push, tag or publication is authorized by this report. The investigation ends here.
