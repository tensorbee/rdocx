# F-X280b implementation, Root final pass 1

Reviewed the thirteen-path frozen worker delta from2e3c0780,1997added and
87removed lines, against operative A280-1/2/3/5/6 and the approved carrier
reconciliation. Patch SHA256 is
`3d5c03ee410217556f3edaf9aa4adcc9be3bc9af9366adea9ab8771f055c6330`.
Native document SHA256 is
`8865a99fadb3aff9176e6709b983c8685bcff610f17f5d1a26912f824fac6ade`.
Root reviewed independently of the implementer. Citations name frozen worker
source. Verdict: four defects, zero smells, zero nitpicks.

## Defects

### D1, known picture containers disable their drawing extent check

`crates/rdocx/src/document.rs:82840`
`crates/rdocx/src/document.rs:82969`
`crates/rdocx/src/document.rs:82991`

SchemaNamespace recognizes DrawingML main and Word drawing namespaces, but
classifies the actual picture namespace as Foreign. The opaque flag then
persists through all its descendants, so the promised a:ext check is skipped
inside pic:pic/pic:spPr/a:xfrm. This is an ordinary supported publisher path,
not unknown extension XML. The existing picture writer emits precisely those
containers at `crates/rdocx-oxml/src/drawing.rs:1998` and
`crates/rdocx-oxml/src/drawing.rs:2034`.

The existing negative fixture places wp:extent directly under w:drawing. It
cannot distinguish a validator that misses a:ext inside real picture XML.
Exercise a genuine authored picture with valid wp:extent and only the picture
a:ext made negative. Require the specified part-named subset finding, with a
valid picture and unknown foreign payload countercontrols. Reconcile actual
known Drawing container paths required by this finite check without making
unknown namespaces transparent or expanding into a full schema validator.
This is source-traced undercoverage, not an executed RED claim or a claim that
the later open predicate necessarily accepts every such mutation.

### D2, related UTF-16 auxiliaries are reparsed through the raw style reader

`crates/rdocx/src/document.rs:83512`
`crates/rdocx/src/document.rs:83524`
`crates/rdocx/src/document.rs:82591`

The new lexical helper losslessly decodes supported UTF-16 auxiliaries and
can correctly report no fault. The related-part loop then always passes the
original encoded bytes to document_style_references, including non-story XML.
That reader has no corresponding lossless bridge. For a BOM-encoded opaque
auxiliary, it can report raw encoded text outside a root or another parsing
fault after the lexical helper has already accepted the XML. Being related
to main therefore changes a valid opaque auxiliary's validation result.

The added UTF-16 control replaces docProps/app.xml, which is not traversed by
the main-related loop, so it misses this route. Use an actual open-admitted
opaque UTF-16 XML part with a main-part custom relationship and compare it
with the same unconnected auxiliary. Both must validate and retain bytes
through save/reopen. Avoid reparsing non-story bytes through a UTF-8-only
semantic helper or supply the proved decoded view to the appropriate existing
consumer. Preserve the actual main/story open predicate, unsupported-input
channels and lexical validation. This records a contract gap in the promised
auxiliary compatibility, not a claimed introduced historical failure.

### D3, first relationship wins the story role for a shared target

`crates/rdocx/src/document.rs:83492`
`crates/rdocx/src/document.rs:83509`
`crates/rdocx/src/document.rs:83516`

Target deduplication retains only the first relationship's story boolean.
A valid custom relationship to header1.xml listed before the actual header
relationship records that part as non-story. The later header relationship
is discarded, and document_story_findings is not called for the actual header.
Qualified child-order and other named story faults can consequently lose
their required findings. Case aliases can also be deduplicated later through
checked_xml without merging their role evidence.

Merge story eligibility across all relationships resolving to the same actual
normalized part identity before scanning once. Exercise a valid header plus
a custom relationship to that same target, both relationship orders, then a
header-only named schema fault. Keep part-local relationship scopes and the
existing content-type/open checks. This is source-traced undercoverage, not
an executed historical RED claim.

### D4, the required fragment or relocation rejection control is absent

`.claude/plans/F-X280b-design.md:105`
`crates/rdocx/src/document.rs:64108`

The operative acceptance contract explicitly requires representative raw XML
and fragment/relocation publisher controls with rollback and retained state.
The private new test injects bytes directly into candidate.package and calls
reopen_staged. It proves the common guard runs in both reopen branches, but
does not exercise a real fragment/relocation publisher or its earlier
allocation, remapping and commit behavior. The retained fx211 test covers an
actual raw paragraph collision and atomic bytes, which is useful applicable
evidence. The new public integration tests cover read/render/allocation/save,
not the other prescribed negative publication path.

Add the missing actual fragment or relocation control under the new imported
duplicate tolerance. Prove the applicable refusal, source and destination
preservation, and retained/fresh identity or location state. Record the actual
refusal stage, without treating a private injected candidate as that public
observation. Preserve positive imported-repeat behavior and the independent
pair/custody checks. This is an explicit local acceptance gap, not a claim of
a demonstrated product failure on that path.

## Smells

None found in the reviewed delta. The concrete scanner state serves actual
Word elements and known drawing/math checks. The new ZIP scan has existing
File and Cursor consumers and reuses the indexed ZIP name decoder rather than
adding an encoding policy. No new public API, crate, module, trait or feature
flag is introduced.

## Nitpicks

None.

## Not found

No additional finding in the examined physical duplicate guard, XML-character
read/write refusal, metadata identity repair, signed-note root/type lookup,
finding cap or bookmark counter/guard ordering. Physical duplicate scans run
before part publication on bounded and unbounded readers. Actual raw names
join existing metadata decoding and normalized identity. The XML carve-out is
removed, with valid binary/non-XML and literal escape controls preserved.
The new read fault uses existing validation and Python error channels.

Bookmark fresh and replay scans use the same counter. Both actual reopen
branches compare candidate facts before reconciliation copies the explicitly
approved retained source allowance. That allowance is not current occurrence
identity or pair/custody proof. The existing checks remain separate. Namespace
scope resolution, default attribute namespace behavior and opaque semantic
stops were examined, subject to D1 and D3's missing known-path coverage.

## Evidence and exclusions

Mechanical authentication SHA256 is
`1cea999ebf65d54e1651bb8c62f70d7a52aa215ac95a840778477d8a56be1365`.
All thirteen frozen paths reconstruct exactly from the claim Base and patch.
Copied executable hashes and sizes authenticate. Actual named11Word,2private,
2Word CLI,5PPT CLI,56Issue430 and1fx211 outcomes are retained alongside scoped
OPC and metadata controls. Strict lint, formatting, prose, drift, whitespace
and Python AST outcomes are recorded. They do not cover the findings above.

The broader OPC75pass/1old-policy failure and PPT18pass/1metadata failure remain
separate observations. Their named updated/corrected controls pass without a
whole-suite confirmation rerun. The prepared RED test hash differs from the
actual receipt's test hash. The actual copied RED binaries and logs remain,
but no full RED source/restore snapshot or unstored per-command seal is
invented. Compiler, fixture and lint errors are not semantic RED.

Actual installed Python and typing, native Word/PowerPoint, SDK/schema,
LibreOffice, final C1 composition, hash/pixel, full workspace, packaging and
hosted CI remain pending. No source, test, plan, HLD, baseline or delivery
record was changed during this review. Return the findings for one batched
remediation and affected checks, without resetting the review lineage.
