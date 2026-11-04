# F-X280c Root source aspect support 1

**Role**: Read-only OOXML/source aspect support for Root's actual independent final1. This is not another formal peer, independent aggregate verdict, design approval or repair cycle. Root owns CLI/Python evidence, table-policy adjudication and the aggregate final review.

**Reviewed**: Stopped worker HEAD72b1f8cbd0fcbcf1b0fefb75c22a3147bc8a2c01, ten-path working diff2001insertions12deletions, frozen diff SHA256 a1924899fbb98fdb1bb9dfb971fa4ffc2e28c217bfec8cf0b054ad7ddfeef304. All ten current path hashes match SOURCE_FREEZE_1.json. Focus was the new stored-file engine in crates/rpptx/src/lib.rs, its actual OPC/lexical/scope helpers and approved finite structural contract. Primary SDK3.0.1 metadata excerpts and source identities were read from .claude/scratch/F-X280c-inputs.

**Source aspect findings**: 4 defects, 0 smells, 0 nitpicks. All consequences below are static source traces, not executed reproductions. No build/test/native action was performed.

## Defects

### SOURCE-D1, structural role selection treats arbitrary custom XML roots as presentation owners

`crates/rpptx/src/lib.rs:3216`, `crates/rpptx/src/lib.rs:3252`, `crates/rpptx/src/lib.rs:3366`, `crates/rpptx/src/lib.rs:3369`

Every ordinary part with an XML MIME type is passed to the same structural scanner. The scanner receives no part role or actual owning relationship and selects presentation/tree rules solely from the root expanded name. Thus an unused application/xml custom part containing `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:opaque/></p:presentation>` receives an InvalidXmlChild presentation-membership issue. An otherwise well-formed custom part rooted p:sld with a PML-looking subtree is likewise treated as a slide role. The model loader does not need to consume that part for this new diagnostic to occur.

The contract's all-XML frontier is lexical, while finite tree/nonvisual/table checks are confined to actual presentation role trees and physical owners. Namespace identity proves vocabulary, not OPC part ownership. Carry actual role/part qualification into structural dispatch. Preserve lexical checking of every XML source without promoting custom/opaque source to a deck role. This is an input classification defect, not permission to demand a general required-root validator or new schema engine.

### SOURCE-D2, valid opaque XML is recursively materialized before the opacity or role decision

`crates/rpptx/src/lib.rs:3359`, `crates/rpptx/src/lib.rs:3366`, `crates/rpptx/src/lib.rs:3392`, `crates/rpptx/src/lib.rs:3495`

stored_next_element invokes stored_subtree for the complete source before the foreign-root check. stored_subtree recursively calls itself on every Start event and retains every descendant in Vec<StoredNode>. This happens even for unused application/xml with a foreign root and for semantically opaque extension/contentPart/foreign Choice payloads. A small valid deeply nested `<x><x>...` custom part therefore consumes one native call frame per XML nesting level, with another recursively owned tree to drop. Sufficiently deep input can exhaust the thread stack and abort validation before the scanner reaches its return-opaque decision. The inherited strict lexical helper is iterative, and the ordinary OPC loader retains this unused source rather than constructing that tree.

Use iterative finite owner/path frames and preserve opaque traversal as lexical-only work. A new depth quota or blanket refusal would change the approved support boundary and is not a remedy authorized by this review. No threshold or executed crash is claimed. The unbounded native recursion on caller-controlled admitted XML is directly present in the cited call path.

### SOURCE-D3, root completion misses non-Text character events outside the root

`crates/rpptx/src/lib.rs:3332`, `crates/rpptx/src/lib.rs:3338`, `crates/oxml-core/src/xml.rs:105`, `crates/oxml-core/src/xml.rs:113`

The added root checker records outside-root content only for Event::Text. It ignores Event::CData and Event::GeneralRef. The shared lexical validator validates a GeneralRef token independently and has no document-position constraint for it, while CData falls through its accepted event branch. quick-xml0.42.0's reader/state.rs:148-157 emits CData without an outside-root admission check. Consequently an unused XML part such as `<![CDATA[POISON]]><x/>` or `<x/>&amp;` passes the new root count/depth check although neither is a well-formed single XML document. Its foreign x root then contributes no structural issue, and the unused part does not force the live deck model to fail.

The promised whole-source lexical/root frontier must reject these outside-root event variants. An XML declaration after the root is another currently ignored position-sensitive event worth a paired control, but the concrete CDATA/reference cases alone establish this finding. Keep legitimate in-root CDATA/reference and outside-root XML whitespace/comments/PIs controls, without treating this as a full OOXML schema requirement.

### SOURCE-D4, metadata suffix matching classifies nonmetadata binary parts as original content types

`crates/rpptx/src/lib.rs:3228`, `crates/rpptx/src/lib.rs:3810`, `crates/rpptx/src/lib.rs:3824`, `crates/oxml-opc/src/package.rs:207`

The ordinary scan correctly uses the MIME classification, but original-metadata recovery selects any raw archive name ending `[content_types].xml`. The OPC loader recognizes only the normalized root `/[Content_Types].xml` identity as its content-type metadata. An additional admitted nonmetadata binary part such as `/customXml/payload[Content_Types].xml`, explicitly mapped to application/octet-stream, is skipped by stored_ordinary_xml_part and then decoded/parsed by stored_metadata_issues solely because of this suffix. Non-XML binary bytes now produce InvalidXmlChild despite the approved rule that extension-only binary parts must not become XML. This static example asserts the current loader's admitted name/classification behavior, not full OPC URI-schema validity of the illustrative name.

Recover the actual root metadata identity under the loader's normalization/case policy. Relationship suffix behavior should remain aligned with the existing loader's relationship inventory. Do not expand C into a new ZIP duplicate/name policy or discard genuine original metadata bytes.

## Smells

None found independently of the concrete defects above.

## Nitpicks

None.

## Checked without additional findings

- The selected fourteen-slot presentation sequence matches pinned SDK3.0.1, including kinsoku and excluding kiosk/smartTags. Singleton rank and order checks use URI identity.
- New immediate Choice and Fallback members reuse the same finite shape membership/nonvisual/table dispatch. Foreign branch payload and extLst/contentPart are not structurally walked by those checks, although SOURCE-D2 records their premature materialization.
- Required cNvPr id/name use direct qualified owner paths and unqualified attribute identity. No new numeric range or empty-name policy was found.
- The shared five-kind nonvisual-parent extraction is concrete and preserves its existing caller match. Existing PptXmlScope/name/leave is reused, rather than replaced with another namespace resolver.
- Lossless UTF8/BOM and UTF16LE/BE decoding preserves original file bytes and rejects odd byte counts/unpaired surrogates. No universal encoding claim is made.
- Role content-type comparison resolves each internal relationship through its actual owning source and compares the expected MIME constant. External targets are skipped. This does not independently prove missing physical targets, which remain existing live validation behavior.
- API ordering preserves initial I/O/ZIP/OPC failure propagation, stored issues before parsed-model issues, deduplication, stored-issue return when model opening fails, and propagation when no stored issue exists. The existing InvalidXmlCharacter mapping is finite and does not assert B's future open behavior already exists.

## Limits and handback

Root separately owns the changed table coverage/less-than policy finding and full contract/evidence adjudication. This support does not approve source, waive any rider or reset review caps. Test assertions were read only as navigation support, not rerun. No source, canonical metadata, Git, native app or external service mutation occurred.

One review-process slip occurred before source inspection: a read command wrote a temporary concatenation of CLAUDE.md and WORKFLOW.md to /private/tmp/f280c-review-workflow-read.txt instead of stdout. Root was notified immediately. That file was not changed or deleted afterward. The requested report is the sole subsequent write and no product bytes changed.

Source aspect support ends here. Return control to Root before remediation.
