# F-X280c, Root final3 structural source support

**Role**: Codex read-only aspect support for Root's actual final3. This is not a fourth GLM author review, new design peer, independent model acceptance, or permission to implement. Root owns the aggregate disposition. The implementing author's self-review3 remains self-review and all three author passes remain used.

**Reviewed**: SOURCE_FREEZE_3.json at worker branch work/f-x280c-codex, Base72b1f8cbd0fcbcf1b0fefb75c22a3147bc8a2c01, cumulative binary diff SHA256 e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007. Twelve paths,2754insertions and45deletions. All twelve frozen file hashes and the current cumulative diff hash matched during this pass. Read the approved contract including Root takeover, Root final1/final2, author self-review3, source/tests, HLD04/10 changes and the retained execution receipts. This pass performed no build, test, application action, product edit or Git mutation. Only this report was written.

**Aspect verdict**:0 source defects,1 test-evidence smell,0 nitpicks. This aspect report is not the whole-story acceptance decision. Shared strict-XML/legacy field-identity namespace correctness and the D3 closure receive Root's separate support and aggregate review.

## Defects

None found in the assigned structural, metadata identity, stored-source selection and consumer aspects.

## Smells

### S1.2, classified corpus immutability baseline is captured after the consumers

`crates/rpptx-cli/tests/integration.rs:540`, `crates/rpptx-cli/tests/integration.rs:556`, `crates/rpptx-cli/tests/integration.rs:558`, `.claude/plans/F-X280c-design.md:99`

The classified LIBRE_OFFICE-100610-0.pptx branch runs the CLI validation and Presentation::open before reading before. It then immediately compares a second read with that late baseline. This cannot detect a file mutation caused by either consumer, although the nearby comment and assertion present it as input-preservation evidence. The other49corpus entries have no before/after assertion here.

Capture the original bytes before the CLI call in the corpus iteration and compare after the selected consumers. This is the already-promised immutable-input control, not a new product policy or request for another broad run. Existing grouped and normalized-metadata controls correctly capture before validation and compare afterward at crates/rpptx/tests/integration.rs:38525 and38537, and38574 and38590. The production file implementation reads input into memory and does not write its path at crates/rpptx/src/lib.rs:3184-3210. This finding concerns the claimed regression evidence, not an observed product mutation.

## Prior finding dispositions in this aspect

### D5, normalized original metadata identity: supported closure

`crates/rpptx/src/lib.rs:3965`, `crates/oxml-opc/src/package.rs:167`, `crates/oxml-opc/src/package.rs:672`, `crates/rpptx/tests/integration.rs:38540`

Metadata recovery now resolves each raw archive name from the package root through the existing OPC resolver, then selects root content types case-insensitively and .rels by the normalized suffix. This handles both directions of the loader-normalization issue, including a raw .rels name ending in a dot segment. It preserves the binary suffix decoy frontier and reads the original selected archive entry bytes. No normalized metadata serialization is substituted.

The new paired regression renames original content types to ./[Content_Types].xml, x/../[Content_Types].xml and //[Content_Types].xml, and original package relationships to `_rels/.rels/.`. For each it preserves the valid source and adds a second root only to the invalid pair, then compares input bytes before/after. The source string _rels/.rels/. in the actual test has a final slash and dot, so the normalized relationship identity remains /_rels/.rels. The checked native issue435receipt includes this test by its full name and reports14passed. D5's source-selection gap is closed in this frozen candidate.

### D6, group own metadata and descendant context: supported closure

`crates/rpptx/src/lib.rs:3653`, `crates/rpptx/src/lib.rs:3752`, `crates/rpptx/src/lib.rs:3851`, `crates/rpptx/tests/integration.rs:38469`

Physical grpSp now enters a nonroot TreeContainer rather than the leaf Owner state, both directly and through a qualified compatibility branch. Its nvGrpSpPr feeds the same direct qualified cNvPr check, with seen_cnvpr returned to its actual group frame. A missing group owner is diagnosed at group close. The same group frame dispatches direct shapes, graphic frames, nested groups and AlternateContent, so nested nonvisual, membership and table checks remain reachable. Descendant extLst/contentPart and unsupported foreign payload remain opaque. Empty elements finalize through the same push/pop path.

The added group cohort covers valid grouped content, missing child name, missing group owner, unknown group member, short grouped table row, group in MC:Choice with a missing descendant name, and an opaque extension containing a PML lookalike. It captures the bytes before validation. Existing MC controls exercise nested recognized wrappers, both qualified branches and unsupported foreign branches. Actual code reuses the same iterative states for arbitrary nested physical groups, rather than special-casing fixture depth. D6's demonstrated reachability gap is closed.

### Earlier D1, D2, D4: no reopened aspect findings

Physical direct qualified gridCol/tc are counted as one each at crates/rpptx/src/lib.rs:3825-3834, and row equality is checked in both directions at3683-3691. No gridSpan parsing or funding is present. Each reached table owns its count.

Role qualification at3238-3298 requires actual internal role relationships plus the expected content type. The main part comes from the actual officeDocument relationship and recognized presentation package class. Structural scanning at3521-3546 additionally requires the expected qualified root. Unused or foreign-role XML still receives lexical/root checks without gaining a PML structural role by local spelling. Actual-role/content-type controls are retained.

Outside-root CDATA and general-reference checks at3426-3430 remain distinct from valid in-root content. Ordinary and original metadata still receive the common lexical/root checks, and the preserved event controls remain in the issue435cohort. D3's core namespace implementation is intentionally not adjudicated by this aspect report.

## Evidence and consumer review

All fourteen successful receipts listed by SOURCE_FREEZE_3.json had actual_child_exit0 and matching complete log SHA256 values during this pass. The native issue435log enumerates14tests including the new group and normalized-metadata cases. The affected CLI receipt enumerates3tests, including the50entrycorpus loop. The current Python validate receipt reports1passed,1037deselected. No aggregate wrapper exit was substituted for child exits. Earlier failing build/CLI/Python observations remain retained, rather than overwritten by a claimed green run. Receipts before the final metadata-only plan amendment retain their original source_diff hashes, not an invented execution at freeze3.

The one corpus outlier is classified through its actual stored diagnostic, /ppt/slides/slide8.xml, tr[4],10physical cells against12grid columns, while the existing parsed Presentation route remains clean. The other49entries retain their clean CLI expectation. This no longer weakens the physical predicate to preserve a blanket all-clean claim. It is not an SDK/native proof of source validity. The preservation assertion limitation is S1.2 above.

Python's implementation remains self.inner.validate under GIL release at crates/rpptx-py/src/presentation.rs:1839, with the unchanged issue.to_string conversion at197. It does not call the new stored file function. The installed test at crates/rpptx-py/tests/test_documented_examples.py:3563 compares the same duplicate-ID Display sentence with the actual CLI and keeps frozen snake_case snapshots. The CLI calls validate_presentation_file at crates/rpptx-cli/src/commands.rs:666 and retains BrokenPipe behavior. The stored function preserves OPC-first errors, stored-first issues, equality dedup and model-open fallback at crates/rpptx/src/lib.rs:3184-3210.

## Not found

- Structural correctness and OOXML:0 additional defects for the finite root-order, qualified physical owner, mixed group/tree membership, MC or physical-table subset.
- Original-source handling:0 additional defects beyond the evidence smell. Ordinary XML borrows original package parts, metadata reads retained original archive bytes, and no input path is written.
- Public consumer/API scope:0 additional defects. Additive Rust file API and stored CLI route coexist with the unchanged Python/native parsed route.
- Structure:0 new unapproved product file, module, trait, generic framework, quota, source rewrite or full-schema claim found in this aspect.
- Nitpicks:0.

Native PowerPoint/SDK outcomes, final B duplicate/open-policy composition, integrated full verification, hash, installed-version/typing, WASM, package/archive, hosted CI and release gates remain pending as declared. This report does not waive them. End this support review and return to Root for the aggregate disposition and any separate remediation.
