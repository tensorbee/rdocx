# F-X272 ordinary root events implementation review, pass 1

Result: zero defects, zero smells and zero nitpicks in the bounded combined implementation. This is partial positive disposition metadata, not whole ordinary coverage or C1/provider qualification.

## Scope and custody

Reviewed ordinary-root-events-increment v1 plus its named-reader-context-successor against canonical4e88519e. The v1 boundary SHA256 is `fe21a77737783cbd3c4f44a043cd460e24d42aa528c6e5bc29edb8c43c642b09`, with patch `4baac7248970cb306b3e2f93db65a164cdac1f87f51852e72786ed0e3d44a7fe`. The successor boundary is `db286be3a27e5e9d928106326b8299e8c6b8c390cd0e7afb8f280780f14dc6e2`, with patch `03f81aea5cab07a7ed93a05a9e7684f8abd7ad45cd38e47f43201964c7f432b9`.

Authenticated all combined five final source pins, before snapshots, the exact v1-to-successor joins and both retained GREEN binaries. Strictly applied both complete patches in memory and checked each context/deletion and full reconstructed final snapshot. Authenticated the indexed prior evidence and successor logs. Final document.rs is `0e006d533294ee9fcf766bff3eea9f7ca54b65e853ecc506c599110a55dbd20b`. Code citations below refer to that frozen successor, with unchanged supporting code from the accepted Base. Read CLAUDE.md and WORKFLOW.md. No moving worker source was treated as authority for the changed implementation.

## Implementation checks

The actual same-reader root branches at `crates/rdocx-oxml/src/document.rs:11943`,11976,12012 and12030 record only qualified document opening, body opening or Empty body, and document closing events. Body opening coordinates are captured at11980 before child parsing at11985. They do not include body content or the consumed body closing event. Unknown background, skipped children, text gaps and the complement of Field ranges are not classified. This matches the exact positive-root scope in `.claude/plans/F-X272-design.md:6131`.

Ordinary bounds accumulate locally at11923 and commit only after the complete successful factory and required body construction at12062 to12075. Existing failure returns publish no new root facts. Ordinary from_xml callers still supply None. The successor's private concrete ScopedDocumentInput at11843 owns only the existing registry/publication Arc handles and delivery flag. It replaces the private nested tuple without adding a public type, signature, reader, model clone or allocation-enabled ordinary None path.

The new reader_input_origin at `crates/rdocx-oxml/src/document.rs:2772` derives displacement from the actual live input and remaining-slice identity, checked lengths and logical parser position. It permits only origin0 or the actual UTF8 BOM3. Equal foreign bytes and unsupported suffix displacement do not establish the relationship. ControlSourceCursor::source_range at3137 translates its actual captured range before checking unchanged transport. This is not a content search or a blanket shift of constructor/decoded-text domains. Namespace closure uses its own physical consumed suffix at4873, preserving the BOM and placing insertion at the actual opening boundary.

WordPublicationCoverage at102 to160 keeps ordinary facts separate from Field intervals. Actual Field facts take priority by subtraction. The monotone interval walk does not fill unknown gaps. Map slice/extend, freeze/materialization, composition, namespace insertion and repeat transport retain exact copied images, including predecessor-independent numeric ownership. Inserted bytes remain unproved until another actual producer observes them. The existing has_publication_owned transport test preserves these numeric facts through the existing map route, without turning them into Word input.

The repaired composition at5975 and namespace closure at6078 preserve ordered copied Field segments. They do not normalize adjacent copied occurrences into one Field demand boundary. Ordinary priority is applied without changing that representation. repeat_source at6483 counts only coverage.fields for reserve_publication_copy, then carries ordinary metadata separately. Root ordinary facts therefore add no Original, current-model, Word, namespace or output allowance. Existing operation identity and sticky resource checks remain in place.

HLD03 at3326, HLD04 at2410 and HLD10 at5169 accurately limit the new root facts and BOM translation. They do not promise ordinary body contents, imported-payload completion, actual Native lineage or a new binding surface.

## Actual evidence and applicability

Authenticated78 distinct successful named tests across the v1 selected logs. Successor logs contain11 successful ordinary-root cases plus6 source_document cases. The other61 v1 results remain earlier-source evidence, not fresh successor executions. The successor differs only in the private context representation and progress text. Its retained Clippy log completes successfully after22.93seconds. Root reports successor format success, with unchanged HLD prose, adapter drift and whitespace checks previously successful. No Native runtime result is inferred from Clippy.

Preserved evidence distinguishes the actual wrong-body-bound observation failure, the three genuine BOM failures plus already-GREEN Word overlap, and the first invalid copy expectation. The raw-fallback attempt incorrectly expected fallback for a direct typed bad-size Run. Final tests instead exercise an actual inline-control fallback and separately preserve direct typed-error parity. These are fixture corrections, not manufactured product failures.

The introduced copied-Field seam regression retains23PASS/1FAIL and its actual source/binary. The existing seam discriminator is unchanged and later passes with its24-case cohort after removal of the new normalization. The BOM controls check exact root and Field coordinates, physical namespace insertion, live-backing rejection and unchanged input accounting. No failure was relabeled as historical claim-Base RED.

The two imported-payload methods, relevant-consumer completeness, actual Native lineage, final installed/external client, resource/history and all final integrated source/oracle gates remain pending. Neither positive ordinary root metadata nor an empty Field list supplies a whole-source certificate.

## Prose and review boundary

Checked this review prose before writing. It contains no em dash, en dash or prose semicolon. Only this review file was written. No source, plan, private packet, build, test, Native operation or oracle was changed or executed. Review pass ends here and returns control to Root.
