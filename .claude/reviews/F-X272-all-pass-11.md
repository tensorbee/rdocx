# F-X272, all aspects, pass 11

**Reviewed**: the complete unreviewed working union since all-pass-10, against HEAD `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd`. The union contains 21 files, 4642 additions and 508 removals. The approved plan is `7cdbb69628edbf0e671af5ade6a8702682a30a32287c7807692444ee4a7e8c86`, with HLD03/04/10. Final source authority is the 3857-entry `frozen-source-v3.json`, SHA `316b1009eff8697120faa83055616e2b229d4bb1feb9bcf34e92a95fd3e6614f`, in `/private/tmp/rdocx-s96-run-refresh/fx272-writer-context-successor-v1`. Every source entry was checked against current bytes. `working-v3.patch` covers the whole union. `incremental-v3.patch`, SHA `8fe472ec164df2a31f7f86bd93222bb1050363c60a85ab4763f8aafd74adcd30`, isolates the latest successor.

**Verdict**: 3 defects, 0 smells, 0 nitpicks. This is not completion or final qualification.

## Defects

### D6, scalar replacement removes an existing Ruby owner and its opaque property

`crates/rdocx-oxml/src/text.rs:5966`, `crates/rdocx-oxml/src/text.rs:10942`, `crates/rdocx-oxml/src/text.rs:10952`, `crates/rdocx-oxml/src/text.rs:10960`, `crates/rdocx-oxml/src/text.rs:11679`, `crates/rdocx-oxml/src/text.rs:11960`

The text-only ordered-cache replacement sends the selected run through whole-slot insertion and removal. Insertion shifts the Ruby base interval past the replacement. Removing the original base then collapses that interval and `remove_runs` discards the Ruby. The slot replacement explicitly restores hyperlinks and revision slots, but does not preserve this enclosing Ruby owner. The accepted-mutation contract requires replacing selected text while retaining shells, nontext and opaque data. Changing the scalar display does not authorize deleting the Ruby property subtree.

The genuine ordered fixture has a direct Ruby base containing `base`, an opaque no-namespace `producerOnly` property child, and an existing deleted revision with ID40 and author `Prior`. Its initial prepared owner is `Some`, with Ruby1 and the expected accepted text. After assigning `NEW` and normalizing, the canonical-w and q-with-local-w=Word probes already show Ruby1 becoming Ruby0 in the prepared owner, before publication, while the prior revision remains. Publication and reopening succeed with `NEW`, so save success does not qualify source preservation. The current regression reaches `crates/rdocx-oxml/src/text.rs:26063` and fails because the published `producerOnly` expanded-name list is empty. Original field owner ID, original source bytes and prior deletion bytes remain exact, which isolates the lost retained owner from source-identity reassignment.

The normal-library phase probe independently attributes the loss to normalization rather than the paragraph namespace publisher. This is not the earlier Ruby-only fixture's unsupported assumption that every Ruby result has an ordered projection. Preserve the actual selected text replacement and the existing owner/opaque payload together. Do not resolve this by weakening the positive or relabelling the discarded Ruby as intentional retirement.

### D7, ordered source synthesis loses its own original namespace consumer frame

`crates/rdocx-oxml/src/text.rs:1805`, `crates/rdocx-oxml/src/text.rs:1811`, `crates/rdocx-oxml/src/text.rs:1819`, `crates/rdocx-oxml/src/text.rs:5896`, `crates/rdocx-oxml/src/text.rs:5898`, `crates/rdocx-oxml/src/text.rs:13339`, `crates/rdocx-oxml/src/text.rs:13367`

Normalization serializes the ordered paragraph before supplying the parsed field's original `word_prefixes`. `ordered_paragraph_xml` invokes the Source writer, whose paragraph and run cores reset the surrounding binding facts, then removes the generated paragraph wrapper. The generated nodes therefore use a canonical-w fallback rather than the original consumer frame in which the returned child fragment is reparsed. Isolating unrelated publication facts is necessary, but does not replace the source fragment's own namespace authority.

The actual discriminator is an admitted q=Word simple `PRODUCER` field with a local `w=urn:producer` binding, a plain q:r result containing `base`, and retained q:del ID40/author `Prior`. It has a valid initial typed ordered owner. Assigning `NEW` then calling normalization returns `normalized field owner changed its accepted source paths`. Source replacement and publication also refuse. The plain case contains no Ruby, so D6 cannot explain this refusal. The same input with canonical w, or q with local w=Word, completes normalization, publication and reopening with `NEW`. The q/local-foreign-w Ruby counterpart independently refuses in the same phase.

The observed defect is a false refusal of the admitted source, not an observed published malformed fragment. The caller's explicit scalar assignment remains pending, so this evidence does not claim rollback of that assignment. Source-role generation must use its own original contractual consumer frame while remaining isolated from unrelated current-part bindings. Keep the exact accepted-path guard and immutable original source pairs.

### D8, the main cache mapper counts paragraphs in an opaque duplicate SDT content slot

`crates/rdocx-oxml/src/content_control.rs:2231`, `crates/rdocx-oxml/src/content_control.rs:2235`, `crates/rdocx-oxml/src/content_control.rs:2246`, `crates/rdocx-oxml/src/content_control.rs:2248`, `crates/rdocx/src/field.rs:17647`, `crates/rdocx/src/field.rs:17651`, `crates/rdocx/src/field.rs:4123`, `crates/rdocx/src/field.rs:4127`, `crates/rdocx/src/field.rs:18594`

The typed SDT parser selects the first modeled `sdtContent` and retains subsequent content elements as whole raw slots. The modeled paragraph collector traverses only `control.content`. The source grammar, however, recognizes every Word `sdtContent` under a control as modeled content, without the typed parser's first-content selection. A paragraph in a retained duplicate therefore enters only the source inventory. The new public cache transaction refuses at its exact modeled/source inventory check.

The actual normal-facade public probe imports a control whose first content contains one complete complex PAGE field with cache7. With no duplicate, and with an empty retained duplicate, `update_page_fields` returns Ok1 and save/reopen yields cache1. A nonempty second content containing an opaque paragraph also imports, saves and reopens with exactly one typed PAGE/cache7, but the public updater returns `field-cache modeled/source paragraph inventory differs`. After refusal the package is byte-exact the independent pre-update save and reopens with cache7. This establishes an atomic false refusal, not package corruption.

These are imported retained off-schema children. The finding does not authorize duplicate-content authoring or assert schema/native acceptance. Make source association follow the same selected modeled slot as the typed reader while preserving the opaque duplicate exactly. Do not weaken the inventory equality check or turn this into unrestricted descendant traversal.

## Evidence and applicability

The deciding lower phase packet is `/private/tmp/rdocx-s96-run-refresh/fx272-ordered-source-phase-probe-v1`, REPORT SHA `512a8f5eb72e16cc1549f8b07ab853b5e6d2319c622c88d1f2f6bc4cc02318fe`, seal SHA `7df2e64ca5d3ca0a3dbfb214c52ea1d63b90069954073746e13314de6ab64adf`. All 105 members were independently checked. Its before/after source maps equal frozen v3. The actual normal OXML artifact has `profile.test=false`, assertions and overflow checks enabled, and rlib SHA `71d5af6cce8b30bf2db0145d76556cab0266a665f7ab605970942ebf52aa99eb`. Six child exits0 record method outcomes. They are not six semantic passes. Their phase dumps distinguish initial projection, normalization, source pairing, publication and reopening.

The deciding public packet is `/private/tmp/rdocx-s96-run-refresh/fx272-duplicate-sdt-public-cache-diagnostic-v1`, REPORT SHA `5d78c3dc089515f2e4e3f08c1e6f6ae0413b57250900470422f89401d5a6564c`, 70-member seal SHA `b29d3f63034af7176e6e38c0a3cab1cc7f650a81c623cdb8f6f379884f241f20`. Root authenticated all members and the unchanged 3857-source map. The report and qualified result provenance were read. Only `fixtures-v2/actual-results.json` and `public-probe2` qualify the three public cases. The retained initial ZipInfo setup failure and rustc dependency-path E0463 are excluded from product evidence. The actual normal facade rlib has `profile.test=false` and SHA `6ae980f01bfdb1f09338b27b9de6052ec9c4a46fb50598a5ea41cf4862868dcd`. No predecessor product was executed, so no old-runtime regression claim is made.

The final main handoff is `fx272-writer-context-successor-v1/PAUSED-HANDOFF.md`, SHA `abaa40a2ca6651cfc5ad27e422ca367965ff0f975114584fdf8e08f4ef369609`, with root-authenticated 111-member paused seal SHA `30ea1e633053b2543d04bc493401819bb5f4f83b27f2525ba050e33ed5a05a2e`. Both 56-child batches stopped at their first child. The earlier batch failed the Ruby-only fixture's invented `Some` projection expectation. The corrected v3 batch failed the genuine retained-property assertion at text26063. Each batch left its remaining 55 children unrun. The first coherent build passed in 83.05s and the successor build in 15.95s, but build success is not semantic acceptance. Historical lint failure, one-line lint correction, strict lint2 PASS and scoped check receipts retain their actual source applicability. They do not substitute for a final complete gate on repaired source.

Final CLI and fresh wheel successors were prepared, not acquired. Earlier 34-Rust/49-Python and CLI/wheel results qualify their original source only. Current 18/12 successor matrices, seven public ancestor rows, six public saves, 15 cache8 rows and remaining named controls have not completed on this final tree. No current installed-binding, Issue158 full-workflow or full-suite acceptance follows from these diagnostics.

## Smells

Zero additional smells.

## Nitpicks

Zero nitpicks.

## Not found

Correctness and contract have D6 through D8. OOXML and unmodeled-source preservation are implicated by D6 and D7. Tests expose the failures and do not establish completion. No separate test-gate defect is added merely because the run stopped on a genuine assertion.

Panics: zero additional findings from checked indexing, source-range/owner association, borrowed length accounting and refusal paths examined. Structure: zero additional findings. No new trait, generic family, file, module, dependency or feature flag was found in the reviewed union. The one additive doc-hidden ancestor-aware document emitter is explicitly approved. Existing public signatures and Ruby struct shapes remain intact. No separate API-shape finding was found.

The audit covered imported-program versus fresh-authoring guards, ordinary qualified metadata filtering versus held/source strictness, immutable field capture and source pairs, source/publication roles, actual root and ancestor binding transport, conditional w14/default knowledge, borrowed closure preflight, control/property/content publishers, revision/Ruby capture and publishers, main/related-story cache transaction callers and HLD03/04/10. No additional distinct finding is recorded in those paths. This is not a claim that D6 through D8 exhaust every possible input in those families.

Native Word, schema, full verification, hosted CI, current archive/README qualification, deterministic hash/raster gates, typing, Python3.9, WASM and release checks remain pending. Q-S96-04 and Issue430 remain pending and unimplemented. No native parity, schema validity, repaired hosted result, release readiness or issue closure is granted. No baseline, public comment, commit, push or publication action was taken.

The review changed only this report. Return the sole writer lease to root for a separate remediation/design phase.
