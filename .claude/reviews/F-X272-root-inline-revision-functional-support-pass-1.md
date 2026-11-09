# F-X272, Root inline and revision functional support, pass 1

**Reviewed**: Seven immutable lower snapshots authenticated against `ROOT_TAKEOVER/root-returned-private-freeze/SOURCE.json`. The bounded change adds returned paragraph-inline control and revision companions, actual nested Text selection and mapped retained Publication emission. The prior reviewed root-main-note-import snapshot supplies the document.rs/text.rs baseline, with 509 unified-diff lines across those two files. This is Root functional support, not a new whole-feature peer cycle or provider acceptance.

**Verdict**: 2 defects, 0 smells, 0 nitpicks.

## Defects

### D1, ordinary retained Publication creates template maps without a template capture

`crates/rdocx-oxml/src/revision.rs:1095`, `crates/rdocx-oxml/src/content_control.rs:1844`, `crates/rdocx-oxml/src/text.rs:4489`

Both new retained Publication paths unconditionally call template_word_raw_map and pass Some(map) into the mapped retained-root emitter. That helper starts by constructing a generated WordSourceMap even when template_word_capture is None. WordSourceMap::generated allocates its piece vector at document.rs:454. The returned mapped record then installs a mapped_source_record lease and acquires the normalization metadata lock at text.rs:4601. No early sink guard suppresses this work.

Consequently ordinary non-template revision and inline-control serialization now creates template companions, including on an unchanged borrowed-root return. The approved None-context path must remain inert. Select the optional map only when the actual template capture exists and retain the existing unmapped emission path otherwise. This is a source-backed ordinary-path allocation and metadata-work regression, not a measured memory result or a new resource quota.

### D2, successful owner handbacks discard revision raw correspondence

`crates/rdocx-oxml/src/content_control.rs:2543`, `crates/rdocx-oxml/src/table.rs:2739`, `crates/rdocx-oxml/src/table.rs:3183`, `crates/rdocx-oxml/src/table.rs:3715`

The new WordBlockSourceMaps.word_raws collection is preserved by append, but these successful SDT, cell, row and table handbacks manually extend only paragraphs, word_texts, word_runs and ordinary. They omit pending.word_raws. For example, the SDT revision producer records the selected retained revision map at content_control.rs:3743, then its successful enclosing SDT return drops that map at line2543.

returned_block_inputs registers revision.retained_xml only from word_raws at document.rs:5199. Its exact captured raw allocation therefore has no returned-map relationship when the next revision Publication writer calls template_word_raw_map. Typed Text facts do not substitute for that separate raw record's correspondence. The same loss occurs for a paragraph revision inside a returned table through any of the three table-family handbacks.

The combined regression selects A&B, C&D and E&F in one operation. A successful final copy reservation can be funded by the other correctly selected inputs, so that assertion does not prove each revision input survived. Preserve the collection at every successful handback and add a discriminator for the exact selected Word allowance or an isolated surviving revision payload after the real next writer. This finding is source-derived. No new RED execution was performed during review.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Source identities

All seven snapshot hashes authenticate against SOURCE.json.

| Path under crates/rdocx-oxml/src | SHA256 |
|---|---|
| document.rs | 93268bb97a0ca24b2d45c1f7dc9cfefcf3309be223774664b9c1680d81693222 |
| text.rs | 6e3c036545eb4b21bb6aefc9892ba0403b2160080d3a47b63a0b48c2377e4eb0 |
| revision.rs | 16b0d1d7213e245283425966fd7d3ad6dc4e6e02e357e7babcc59c87067b2e91 |
| content_control.rs | 223b858ab7aa9fb240769c6d40d7fcea05dcc7e630a0812c35bca8b4b89e5b21 |
| table.rs | 1a5b1c8e6cbc0ee7d1da7a52da2bfff3d9c8fed127e299cf5b57aacfa638952b |
| footnotes.rs | 58ee89d9641a29232f5ac794b8dcad6e70b493ddddef18b5df093eb4486348f8 |
| placeholder.rs | b704f39f0fe4cb499120068c8ebbf9d7ca8e9317ee8209d6e5048d5af9c59374 |

## Mechanisms and evidence examined

The successful revision constructor stages its optional output and publishes it only on Some(revision). The ordered revision path uses the same mapped synthetic paragraph factory and prefixes its actual RevisionParagraph return. Paragraph and SDT producers append successful children under their actual pushed owner indices. The new TextWalk selection map is built from actual borrowed returned Text objects and remains local to the same live paragraph and walk. It does not search by equal values or enroll returned data as new input.

The retained Publication writers use the existing retained-root planner and same piece-emission transcript to return maps beside bytes. Their copied and generated regions stay distinct. No second parser or new public method is introduced. D1 and D2 concern the optional-context and successful handback boundaries around that mechanism.

The archived strict log records completion in 26.08 seconds. The scoped log records six named scoped_returned_read tests passing. The new combined test exercises body, table, block-control, standalone paragraph, footnotes and endnotes inside one named test. Its byte comparison and charged consumer are useful, subject to D2's per-input discriminator gap.

The broader scoped batch records 195 passes and two failures. Root separately reproduced both failures, scoped_format_paragraph_unchanged_field_retains_actual_word_source and scoped_original_textbox_return_materializes_known_field_roles_before_source_drop, against the archived pre-increment lower binary. They remain actual existing debt, not new defects assigned to this increment or a passing whole gate. No reviewer build, test, native application or security probe was run.

## Explicit boundaries

Other consumers including row grammar and raw-textbox companion selection remain pending. Changed or rebuilt revision returned contexts, comments/glossary, complete non-emitting expression evaluation, final handoff joins and provider closure are not certified. Final callable inventory, integration, SemVer, complete local and integrated gates and hash/native/oracle riders remain pending. This review neither changes the input equations nor approves a new API, quota or acceptance policy.

## Not found

No additional concrete byte-output, namespace/schema-order, panic, operation-identity or successful-return staging defect was found in the bounded new code examined. No independent structural smell or nitpick was found. Those zero categories do not close the explicitly excluded contexts.
