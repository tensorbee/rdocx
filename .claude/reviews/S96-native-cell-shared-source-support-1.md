# S96 Native Cell shared initial source, support pass 1

**Reviewed**: The five-path frozen increment against cec87f9501eb13d0d5d60572472b227483262620, consisting of Native initial Cell-source sharing, three Native regressions, a separate binding test-only cold diagnostic, and the existing F268/HLD03/HLD10 contract additions. This is independent bounded implementation review, not a reset of the original review cap.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Authentication

/private/tmp/rdocx-s96-native-cell-shared-source-implementation-1/SOURCE_BOUNDARY.json has SHA256 f2d5a60be0a94fd61df19e4cbe32b34ae16cdc0f3327b18946fb99004266b4a4. The increment patch matches 11a92cff75507ff6de752ac64543ff850fe0e29807a2ec1b6f6f2a844962f8ac. All five current final source pins, all six protected F273 source/review pins, all three retained binaries and all 21 recorded log hashes match. The final source pins are:

- Native document.rs, 84c07547ef6f28d3b76be9a080f18b58684e843ee71091804cab3c76134fa237
- Binding document.rs, 9dbf158857563e2eb0823be3167f2c68e0775eb66e031c65c72113c91437af75
- F268 plan, 669dff1482fdf72ccac725097fc67acf091a076bf38216c30159bb1a1860735e
- HLD03, 2972fe67133c58c3e7ea0ec5ab74a9c43fe5da13fa383077d87d5166a7d26fbc
- HLD10, 4fc6cedd7dd0af1392cf0db9ad14e55ef5dbe4feb0f3bc7e89eef79011023612

The Root finite design approval and the operative F268 amendment require exact current writer equality, complete related discovery, current typed selection, optional legacy fallback and unchanged postflush checks. The implementation was reviewed against that finite contract and the current canonical microscope command.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Checked implementation and consumers

Correctness and contract: crates/rdocx/src/document.rs:32047 preserves the public direct-body method and delegates its exact old qualifier to direct_body_item_location_in. That substantive qualifier now serves both the old public consumer and the new private Cell preparation, including the same full fallback and debug differential checks. It introduces no public signature or selector.

The new private direct_body_cell_story at document.rs:32109 obtains actual current main bytes and physical body spans. Its guard at 32131 uses offered_canonical_blocks and the existing writer_bytes_match, not a retained Cell snapshot. The latter requires actual document frame, explicit declaration eligibility, current child count, reusable current child facts, no held Field words, current staged model equality and complete exact byte-piece correspondence at document.rs:10263-10297. A proof miss or non-table/end location calls the original story_table_cell method. This is a current writer proof, not scoped input authority or a new budget allowance.

The qualified route builds its main StorySource from those same actual bytes, collects complete related_story_sources before qualification, then reuses story_source_owner_and_layout_in, story_item_span and physical_table_cell_owner. The unchanged source-owner helper checks the actual owner fingerprint at document.rs:30175. The old spanned_story_sources order at 29954 likewise prepares main, then related, then owner/item lookup. Consequently, this sharing neither suppresses unrelated related-part errors nor invents cell ordinals. Typed pointer selection remains first at 48112, the existing optional failure fallback remains at 48133, and the common candidate flush and full postflush Source/Field/physical correspondence pipeline remains at 48202 onward. A live formatting change is qualified from current document bytes, not the previously rejected cached StoryId shortcut.

Panics and errors: the new path uses existing fallible parsing and physical-owner helpers. Its main-body mutex is released before downstream qualification. It adds no unchecked index or persistent borrowed state. Errors remain optional only at the same existing Cell direct-route boundary, with the legacy controlled-content route unchanged. The direct public location method retains its original error behavior. No new parser, resource limit, Source credit or allowance appears.

OOXML and structure: no serialized content or namespace policy changes. The implementation shares an existing producer and reuses the complete physical parser and owner checks. The extracted helper has two actual consumers. No new type, trait, generic, wrapper, module, crate, flag or retained cache is introduced. HLD03 and HLD10 accurately distinguish ordinary writer equality from scoped source authority.

## Acceptance evidence examined

The actual predecessor RED succeeds full owner and package equality before observing two current outputs instead of one. RED.log records exit101 and scoped.normalized_opening_copy two outputs totaling 1168 bytes. This is an executed mechanism RED, distinct from the initial mutable-receiver test compile failure.

The three final Native tests pass. They require exact current StoryId equality with the full legacy lookup, unchanged full package bytes, one output on qualified inputs and two on missing-declaration, proof-disabled and Field cases. Their raw and alias cases, invalid coordinates, live text/bold mutations and missing related-part error priority discriminate against stale selection and suppressed discovery. Seven existing controlled/direct, physical/merged-cell atomicity, foreign namespace, owned-transaction, Field/raw/refusal and related-owner filters pass. The first insufficient-stack abort is retained separately from the successful explicit 8 MiB rerun. Strict Native/binding all-target/all-feature Clippy and metadata checks have recorded exit0. These observations and retained binaries were authenticated without rerunning them in review.

The added binding code at crates/rdocx-py/src/document.rs:20469 is test-only. Its eight 250/1000, held-handle, insert/remove rows report a real prepared plan and zero Native capture output work. It helps distinguish a missing-plan explanation on that semantic rich fixture. Its generated ZIP and quote spelling are not the original timing input. It does not prove the deciding cold timing or a performance improvement, and no binding production route changed here.

## Limits and handback

The bounded increment has no blocking source finding. Root final review remains separate. Fresh installed/current-artifact performance and full integrated gates, all native and memory riders, and the original deciding Cell/cold measurements remain pending. Cell.add_picture is outside this increment. This review performed no compilation, tests, product import, rendering, timing, UI operation, source fix, status update, ledger or baseline change.
