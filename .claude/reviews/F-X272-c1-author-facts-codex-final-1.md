# F-X272 C1 focused author facts, Codex final 1

USEFUL INCOMPLETE CHECKPOINTS. No complete mechanism, source lease, fourth peer, counter reset or provider completion. Actual artifact pins are retained in /private/tmp/rdocx-cross-review/C1_FOCUSED_AUTHOR_FACT_PINS.json.

## Grok projection checkpoint

The final stdout checkpoint is an actual same-session response after the focused author hit its 16-turn cap without a file. It correctly states that no structural witness is designed. Its claimed source-origin bug is refuted by the actual reader semantics.

At /private/tmp/rdocx-grok-harness/fx272-c1-v6-author/PROJECTION_CHECKPOINT_STDOUT.md:5 the claim treats Reader::get_ref as the original full input and infers that relative spans have the wrong base. In fact crates/rdocx-oxml/src/content_control.rs:3012 saves physical_source from the current underlying slice before parsing. The following source_origin records the absolute counter, and :3018 and :3104 subtract that counter to obtain offsets into the saved remaining slice. The pinned quick-xml 0.41.0 reader/mod.rs:858 returns the underlying reader, not an immutable original-input archive. Its buffered_reader.rs:82 consumes that reader, and slice_reader.rs:231 explicitly notes that self.reader changes and saves the current reference for the same reason. Absolute buffer_position and an advancing remaining slice are compatible. No base-mismatch defect is established by the checkpoint.

Other projection refusal and source-identity conditions remain open. This refutation does not prove admitted input always projects or that the span comparison is unreachable. No runtime test was performed. Root requests a corrected structural-witness analysis without repairing a nonexistent base bug.

## DeepSeek semantic error inventory

The actual completed inventory distinguishes propagated hyperlink errors from sdt/revision raw fallback and names the author-required versus author-optional revision difference. These are useful source facts. It explicitly excludes untraced CT_RPr, drawing, synthetic paragraph and table/row/cell errors. It is not an exhaustive witness for the full approved selection contract.

One proposed helper direction is invalid as written. ERROR_INVENTORY.md:118 proposes run_retention_facts reusing capture_root_attribute_record. The latter is the manufactured record writer at crates/rdocx-oxml/src/text.rs:692, including record construction after :735. A preflight witness must not call it before admission. Extract or reuse pure qualification/attribute facts without allocating the retained record. A list of unresolved subparsers cannot qualify the exact-selection promise until those paths are closed or an explicit contract amendment is approved.

Zero new C4 findings. No parser, test, native, schema, hash or source qualification was executed. Root retains the existing source contract. The unfinished author work may continue in focused TEMP outputs and returns directly to Root. STOP.
