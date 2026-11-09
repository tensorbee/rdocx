# F-X272 DeepSeek source increment, Grok peer pass 3 of 3

Stdout review only. No file was written. Prior reports stay immutable. This is not feature acceptance and not a model verdict.

## Identity

Reviewer Grok, session `01a19364-efef-7bc1-97a0-a7ec4c726991`. Author DeepSeek after peer 2. Codex is coordinator only. Branch `sprint/s96`, HEAD `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd`. Packet kind `source-after-peer2-frozen-no-builds-until-audit`. Handoff `fc11ff69008b17458953f79f6d8f2189a23542f12a3fa0316f3042c02176d3ec`. Plan `.claude/plans/F-X272-design.md` `e37d2ded5174b69d322a053fe85b85afbedddd7448aacfb5033b5922454d9949`, unchanged. HLD 03 `e6cc502f9431dc3bda4f72c4a623ad5c78d87127fe7009343f165b96354dc5c8`, unchanged. HLD 10 `c09cf197a66315737897aa597361ff2d32dfa1a109d8db67930924e985410ade`, unchanged.

Authentication before and after the read: 923 manifest files, packet copy mismatches 0, canonical tree mismatches 0, including 4 metadata candidates. Code identity matches the packet. The author delta against the peer 2 30 pins is oxml `document.rs` `fccc968e6429427580fccc342d68f565bcd449c29eeae1be921ade391b1d91f6`, oxml `text.rs` `ab317d23825bfe321b3116ddabe0c5cdb2b35678bce0c79250527159e394dd18`, oxml `placeholder.rs` `98e6689c` as named in progress, progress, `CURRENT_SPRINT.md`, and `SPRINT_PLAN.md`. `drawing.rs`, facade `document.rs` `756c7632`, facade `field.rs` `a7efc10f`, `revision.rs`, the plan, and HLD 03 were not in that delta. F-X272 is in progress. Owner column remains codex. `working-tree.diff` was not used as authorship.

## Prior dispositions

Peer 1 field-cache D1 and debug-span S1 stay closed. Their facade files match the peer 2 pins and were not re-read body by body.

Peer 2 D1 is closed for the retained non-empty cases the new tests actually store. A non-empty direct `w:drawing` is the drawing root (`document.rs:652`). A Start inline or anchor is selected by local name (`drawing.rs:2146`, `drawing.rs:2177`). Used run-frame bindings, including explicit canonical `w`, are charged at `depth <= owner.depth`. Used drawing-scope bindings are charged at `depth == owner.depth + 1` (`document.rs:704`). `raw_with_external_bindings` copies only bindings the raw uses, including MC and `xsi:type` value prefixes through `root_attribute_value_prefixes` (`text.rs:19529`, `text.rs:19658`). Unused xmlns declarations are not copied and are not charged. One parsed owner stores ancestor `w` and document-scope `x` on `rdocxRootAttributes`, copies drawing-scope `x` onto the inline raw, stores local-name `<w:inline>` with consumed `wp`, drops `x:opaque` to an empty drawing, and keeps an empty `w:drawing` verbatim. `mc:AlternateContent`, `w:pict`, and `w:object` stay in the raw class in code. Those three were not re-executed.

Peer 2 D2 is closed on the public removers and the counted collapsed total. `collapsed_hyperlink_source_total` counts escaped attribute values (`text.rs:1346`, `text.rs:12269`), the start-tag `>` as `2 + name_len` (`text.rs:12244`), the end tag, and `retained_source_len` for in-range hyperlink revisions (`text.rs:12278`). Source `write_xml_role` writes `raw_xml` unchanged (`revision.rs:325`, `revision.rs:808`). Admission runs before mutation and the reservation is the needed total (`text.rs:12332`). `remove_runs` and `remove_comment_anchors` clone and install only on success (`text.rs:12039`, `text.rs:12095`). Direct and revision `remove_accepted_run_segments` propagate. Field `set_cached_runs` and `normalize_cached_result_owner` clone before ordered replacement (`text.rs:1639`, `text.rs:1712`, `text.rs:1868`). The backstop at `text.rs:12584` still runs after the candidate is mutated. The public install happens only after `Ok`. The atomic log is labeled fault injection in the test at `text.rs:28092` and is not old-code RED. The total RED is a real pre-fix reservation miss.

## Defects

### D1. Empty inline and anchor are still charged

Severity: high.

`typed_here` is true for every direct `inline` or `anchor` under a drawing root, with no empty check (`document.rs:662`). `dropped_here` is only the complement (`document.rs:666`). The comment at `document.rs:664` says an empty inline or anchor creates no record. The parser handles only `Event::Start` (`drawing.rs:2144`). `Event::Empty` falls through at `drawing.rs:2238` and creates no inline, no anchor, and no `raw_xml`. A prefix used by that empty element, including the element prefix itself, is still appended when its depth is the run frame or the drawing scope (`document.rs:704`). `finish_original_capture_owner` then counts that record (`document.rs:569`).

Consumer: `admit_word_xml` before `CT_Document::from_xml`.

Discriminator: many runs whose only drawing child is an empty `wp:inline` or `wp:anchor` using a long ancestor or drawing-scope prefix. Admission returns `original source capture exceeds admitted source work`. One parse of the same shape stores no binding record. This input was not executed. The green batch has no empty-inline test.

Remedy: classify an empty inline or anchor as dropped, charge only a Start the parser retains, and keep the comparison before allocation.

### D2. `record_count * 22` raises the capture allowance

Severity: high.

`finish_original_capture_owner` adds 22 to `total_needed` per stored record (`document.rs:569`). The gate then adds `record_count * 22` on top of `retained_owner_work` (`document.rs:773`). The wrapper cancels. Declaration bytes are what remain against the envelope. Plan `e37d2ded` forbids a raised limit and an expansion-derived allowance, and it says projected records are not new input budget. Copies above the unchanged allowance must refuse. No budget waiver exists. The 2000-box rows in `fx211_weighted_source_boundary_and_actual_anchor_routes` are the separate weighted-owner gate that expects `weighted limit 256` (`document.rs:5803`). The author log reports that test ok. That gate does not authorize this add. An intermediate format log showed the same add at 16 bytes per owner. The frozen code uses 22.

Consumer: the same capture preflight, before model allocation.

Discriminator: N retained-binding records whose declaration bytes fit `retained_owner_work` and whose declaration bytes plus `22 * N` do not. The current comparison admits that input. The unchanged formula refuses it. Not executed as its own probe. The 200-run and 400-run refusal tests still refuse in `fx272-peer2-d1-post-05.log`, so this headroom did not flip those named cases.

Remedy: compare `total_needed`, wrappers included, with `retained_owner_work` alone.

### D3. Placeholder replacement publishes an ignored removal

Severity: high.

`replace_in_paragraph` edits the caller's paragraph, then calls `remove_emptied_runs` (`placeholder.rs:305`). `remove_paragraph_emptied_runs` and `remove_paragraph_run` discard `remove_runs` with `let _` (`placeholder.rs:787`, `placeholder.rs:796`). `remove_runs` clones and installs only on success (`text.rs:12095`), and it calls the core with source context `None` (`text.rs:12097`). The recovered wrapper context is used for `write_wrapper_back` (`placeholder.rs:709`) and is not given to the removal. On a refused collapse the removal clone is dropped, the earlier text edit remains, and the public replace returns a match count. The progress note that the staged paragraph keeps its prior state describes only the removal clone.

Consumer: `replace_in_paragraph`, `replace_prepared_occurrences`, and wrapper write-back.

Discriminator: a hyperlink whose replacement empties every run, with retained raw or a revision, whose collapsed total exceeds the allowance. The replace count is positive, the text has changed, the emptied runs remain, and no collapsed element is installed. Not executed.

Remedy: apply the replacement and the removal on one uninstalled candidate, pass the recovered source context into the private core, and leave the caller unchanged when removal refuses. Keep the public signature infallible by withholding that candidate.

## Smells

0.

## Nits

0.

## Evidence and limits

No cargo, test, format, git, or network action was run. Named logs were read.

`fx272-peer2-d1-red-02.log`: `test result: FAILED. 0 passed; 2 failed` is not the line. The log says 0 passed and 5 failed, with four `Ok` under-counts and one dropped-sibling refusal. Cargo prints `error: test failed`. No separate EXIT stamp.

`fx272-peer2-d1-post-05.log`: `test result: ok. 14 passed; 0 failed`, 10.36s. Includes the nine drawing tests, the shared and repeated owner tests, the Ruby and VML route, and the 2000-box weighted test.

`fx272-peer2-d2-red-total-01.log`: revision and escaped tests fail with `collapsed hyperlink replacement exceeds its reservation`, 1 passed, 2 failed.

`fx272-peer2-d2-red-atomic-01.log`: fault-injection assertion fails, 0 passed, 1 failed.

`fx272-peer2-d2-post-03.log`: 4 passed. `fx272-peer2-d2-post-04.log`: 1 passed.

`fx272-peer2-oxml-all-post-04.log`: `test result: ok. 804 passed; 0 failed` in 8.22s. This is an observed `rdocx-oxml` lib diagnostic, not a workspace or final-union result.

Facade comment log: 2 passed. Layout comment log: 1 passed. Check logs that were searched show the existing `table.rs` warnings and no `could not compile` line. No EXIT stamp was present. Clippy post-04 reports 8 lib errors and 12 lib-test errors, with help links under `rust-clippy/rust-1.93.0`. Progress says `cargo +1.93`, MSRV 1.93.1, and `RUST_MIN_STACK=33554432`. No `rustc -V` line was read. No exact 1.93.0 credit and no default-stack credit. No new clippy finding is added beyond that log. Format post-04 has no `rdocx-oxml/src/document.rs` hunk and still has `text.rs` and `placeholder.rs` hunks. The 232 count was not recounted. Old format and clippy debt is not waived. `capture-rdocx-all-03.log` stays UNKNOWN. Prose post logs say 0 violations.

HLD 03 lines 2812-2814 still require one collapsed reservation before construction. That matches the counted total. The file was not edited. The plan HLD impact list remains HLD 03, HLD 04, and HLD 10. The 22-byte allowance is not an HLD change and is not a design waiver.

The table-depth assertion after `document.rs:6513` was not finished. Unchanged facade field and document bodies were not re-read. Anchor `source_namespace_bindings` remains a projection clone. Serializer authority is `raw_xml`.

## Pending debt

D8 named union, all-pass-12, Issue 430 callable ownership, collisions, import, cross-document identity, and selective cleanup stay pending. They are not new increment bugs. Full native, schema, CLI, Python, wheel typing, WASM, archive, hosted, hash, raster, and release riders stay pending. The native omitted-type probe remains a qualified design reference only. No dependency is unblocked. This source counter ends at peer 3. There is no peer 4.

## Verdict

REJECTED INCREMENT.

Three defects, zero smells, zero nits. The lease is returned to root for the metadata writer to adopt this text only if the frozen hashes still match, then for root final review and any author fix with an affected root recheck. Source peer 3 is complete.
