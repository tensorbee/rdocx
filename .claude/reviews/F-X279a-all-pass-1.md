# F-X279a, all aspects, implementation pass 1

**Review kind**: actual independent read-only microscope. This is the first implementation pass, not a design peer or implementer self-review.
**Reviewed**: working diff on work/f-x279a-codex against Base6bea9093818d6b691bf88b2980e9cc17927be015. Eight tracked product/test/HLD files,780 insertions and24 deletions. Exact diff SHA2563fdf1ecfad95a79557e636c339a4ca1e525780d39b8e1f0582b7bae63591d329.
**Verdict**: 2 defects, 0 smells, 0 nitpicks. Do not prepare this candidate. Final integrated qualifications remain pending independently of these findings.

Read CLAUDE.md, the canonical workflow and microscope command, the approved child contract, operative clean-Base correction, progress record, cited HLD text and relevant existing parser, accepted-text, field, revision inventory, Python export and CLI consumers. Only this review file was written. No builds, tests, imported-extension execution, source edits or fixes were performed.

## Defects

### D1, independent small revisions repeatedly scan and replay the whole source

`crates/rdocx/src/revision.rs:1183`
`crates/rdocx/src/revision.rs:1194`
`crates/rdocx/src/revision.rs:774`

Every record calls snapshot_text, which allocates and constructs paragraph_source_domains over the entire tree. For a paragraph-owned record it then scans every tree element again, collects every preceding paragraph in that structural domain and replays their descendants. This occurs even when there is no field in the input and only the selected one-character insertion needs text.

A document containing N consecutive ordinary paragraphs, each with one independent one-character insertion, has N output records and N characters of required output. All those paragraphs share one existing source domain. The new implementation performs N whole-tree domain computations plus a replay of1,2,...,N paragraph prefixes. Its work and cumulative temporary allocation grow quadratically while both source and required result are linear. The same problem applies to disjoint run-property changes. This is unnecessary new work from the read projection, not an unavoidable large-output case or a request to invent a new quota.

The source/domain and field-context facts need reuse across the actual story-part snapshot batch, with each necessary source observation processed a bounded number of times. Preserve the agreed one-tree mechanism, actual prior-field context and source/error policy. Add a discriminator that distinguishes linear independent revisions from genuine overlapping large text outputs, using operation counts or another direct bounded-work observation rather than a fragile elapsed-time threshold. Existing small fixtures do not reject this implementation.

### D2, a modeled property revision can borrow an unrelated control's text

`crates/rdocx/src/revision.rs:1156`
`crates/rdocx/src/revision.rs:2996`

The property-current branch selects the grandparent without checking the actual qualified property/owner chain. Resolver metadata is not sufficient proof of that chain: the unchanged modeled_child admits rPrChange, pPrChange and tblPrChange directly under sdtContent.

For example, a preserved inline control containing a direct rPrChange with id/author, followed by an ordinary run whose text is unrelated, is already inventoried as a modeled revision by the existing tree. The new branch treats sdtContent as the properties and sdt as the current owner, then returns the control's unrelated run text as a run-property snapshot. Equivalent direct pPrChange and tblPrChange records receive that same control text, without an actual current paragraph or table property owner. This contradicts the approved requirement to select the real rPr-to-r, pPr-to-p and tblPr-to-tbl relationships, rather than infer ownership from two parent steps.

Validate the actual qualified chain before attributing current owner text. Preserve existing revision inventory and imported raw retention/error behavior, including the valid paragraph-mark case whose mark has no run text. This finding does not authorize schema-wide rejection or removing previously counted revisions. Add a direct misplaced-property control versus a real property-owner positive, so preserving or refusing an unsupported owner cannot silently fabricate text from neighboring control content.

## Smells

Zero findings.

## Nitpicks

Zero findings.

## Other correctness and contract checks

The owned String and public borrowed text getter are populated by both actual producers. story_part_revisions uses its existing XmlTree and source-index order, retaining innermost StoryId selection. The private comment-package producer explicitly stores empty text. No RevisionRef text API, synthetic owner serialization, derived fragment parse, decoder API, public quota or resolver modeled_child change was added.

The scalar observations reuse the existing parse stream and public core decoders, keep entity-split whitespace and unknown entities, and do not unescape CDATA twice. Text-bearing direct run children are selected through a separate read ancestry. Opaque drawing paths remain closed even where a deeper txbxContent can independently restart resolver modeling. The existing resolver marker still requires a modeled run. New marker observations avoid introducing value-decode failures for otherwise opaque runs.

Current property mode excludes removed revision ancestors and rows. Revision-local mode retains the selected deletion/move-source payload and nested atoms once. Prior property subtrees are not walked. Table traversal adds paragraph tabs and retained outer-row newlines without nested table aggregate duplication. Control child selection uses the physical enclosing table/row/block/inline grammar. Clean-Base transparent-wrapper canonical-w restrictions are explicit, and final C1 coherence remains due. Field phase observations distinguish instruction from stored cache, preserve outer suppression around nested result text, and retain actual story/domain boundaries. Apart from D1 and D2, these aspects produced no additional cited defect.

Python PyRevision retains its frozen/get_all/equality model, four required keyword fields, required nullable timestamp and optional story. The new owned text defaults to String::new with a matching visible empty-string text signature, the stub adds the same optional str parameter and property, and None has no new conversion path. The real native-to-Python producer copies native text once. Existing class registration and public Revision export remain at src/lib.rs:231 and python/rdocx/__init__.py:131/243. Source inspection does not prove the installed constructor signature or equality at runtime.

CLI JSON takes text from the native getter. Both actual revision_kind_label consumers use the same exact eight snake_case labels, while story_kind_label retains its independent spelling. Human output still writes seven columns and adds no unescaped text column. HLD03 and HLD10 stay within the approved projection/snapshot sections and do not claim new comparison/comment/import support.

## Tests and evidence inspected

Authenticated all eight SOURCE_FREEZE_1.json source hashes and the exact current diff hash. All retained JSON receipts that name log_sha256 match their full logs. The actual baseline log shows two semantic old-CLI failures: missing text values and run-property-change versus run_property_change. These are different from the retained PyO3 default conversion setup failure and native-third compile setup failure, which called unwrap on the infallible Document::new. The earlier field-context cases passed. The final native receipt reports12 focused cases passed, the CLI label receipt1 passed, CLI composition/lifecycle2 passed, and the broader revision control receipt13 passed. The scoped all-target/all-feature Clippy receipt for the three changed crates has actual child exit0 with warnings denied. Format, prose and whitespace receipts are separate checked observations.

The native cases exercise all eight kinds, scalar/event/opaque paths, removed ancestor/row state, nested tables/controls, exact separators, source roots/order, stored simple and complex caches, prior-paragraph field context, domain isolation, local nested revisions and the private package producer. The real CLI case additionally saves/reopens native owned snapshots, accepts changes, retains the earlier owned values, checks actual JSON, human columns and unchanged original input. Existing field carriage control remains green. The two defects above identify untested discriminators rather than invalidating these observed outcomes.

The current retained executable identities were read without execution:

| Artifact | SHA256 |
|---|---|
| Native library test rdocx-4b490f76e4e2a26d | efd9315aa82c9349301d945e75d1299bc6d8e1117401db93462eec20cb058590 |
| CLI label test rdocx-33b100aaa3ec767a | e7078ef25c374dc17f3abe934e0107a93a66c26cfe24fc90b4918e0b57ea1a8e |
| CLI integration test integration-c9e904dddc7fee67 | bbe440ac51c143b7d277e1a0273c35d76db85b20bc571c9b346e437f8c048b36 |

These are current artifact hashes corresponding to retained log paths, not retroactively captured hashes of the overwritten old executable. No old artifact identity or installed Python run is invented. Python regression sources cover public constructor omission/explicit values, visible signature, frozen assignment, equality, None refusal, explicit story, real document fixture, held snapshots and save/reopen/accept. They are source/AST evidence only until the actual child-source extension is installed and qualified.

## Not found

Additional panics, OOXML source mutation, namespace decoder divergence, serialization child-order changes, binding/export mismatch and structural abstraction violations produced zero findings. Private helper indexing follows existing tree indices and the actual revision caller. No new module, file, crate, trait, generic, feature, dependency or second source parser was introduced. The new iterative traversal avoids source-depth recursion, although D1 still blocks its work complexity.

## Acceptance boundary and handback

F-X279a remains in-progress, its plan approved and parent/B/F272 barriers unchanged. Installed Python3.9/3.12 runtime, signatures/typing/stubs, final C1 alias/wrapper coherence, integrated full/hash, applicable SDK/native, WASM, archive/dry-run and hosted CI remain mandatory pending checks. Q-A430-28 supplies no comment or publication waiver. Neither current local test counts nor this review complete the story or provider.

This review ends with this file. Return control to Root for a separate remediation phase. No fixes, preparation, confirmation pass or new design peer were started.
