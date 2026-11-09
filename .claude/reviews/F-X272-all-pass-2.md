# F-X272, all aspects, pass 2

**Reviewer**: independent Codex reviewer, not an implementer of F-X272. The six focused investigations are historical design escalations, not substitute feature microscopes.
**Reviewed**: complete working diff on `work/f-x272-codex`, Base `c8acc502b74f29ed7e9546939f17007709a9f22b`, 14 tracked files, 5002 added and 763 removed lines. Tracked binary diff SHA-256 `88182fe9d4522b683057e4c26ec6efc49ff6ba0f7cb5fb1558cb341fe34cc4b2` remained unchanged through review. Historical untracked review reports are not product changes.
**Contract**: canonical0e9ae23a as reconciled into the worker plan, acknowledged by the implementer. F-X272 remains in-progress in CURRENT_SPRINT and BACKLOG. Read microscope, repository workflow, approved plan/progress, prior full review/investigations, and the cited current HLD03/HLD10 comparison contracts.
**Verdict**: 1 defect, 0 smells, 0 nitpicks. D2 remains blocking at the protected-owner source guard. No scoped integration-readiness approval. Final integrated qualification remains pending independently.

## Defects

### D2, meaningful opaque property-root edits bypass the protected-owner refusal

`crates/rdocx/src/comparison.rs:4228`
`crates/rdocx/src/comparison.rs:4242`
`crates/rdocx/src/comparison.rs:3952`
`crates/rdocx/src/comparison.rs:3654`
`crates/rdocx-oxml/src/text.rs:10041`
`crates/rdocx-oxml/src/run_properties.rs:601`
`.claude/plans/F-X272-design.md:13`
`.claude/plans/F-X272-design.md:203`
`docs/hld/03-architecture.md:1425`

A protected owner is emitted raw immediately when the two typed BodyContent values compare equal. Its held-revision refusal runs only in the unequal branch. Paired pPr and rPr root metadata can be meaningful source state without appearing in the typed property model, so equality is not sufficient proof that the protected owner is unchanged. The new raw-source frames intentionally skip held owners. The positive default-property projection does not cover a nonempty property root such as pPr containing jc or rPr containing b, and cannot retrospectively enforce the bypassed guard.

Actual bounded public-API diagnostics during this review authenticated the already built native module as `b301c71fddfd1c62ed8b8ede223594739bc8c7961effe2087ac8de5e5fd6d1c5`. Inputs were constructed entirely in memory using public `Document.from_bytes`, `compare` and `to_bytes`. No source, test, package or artifact file was written and no build or suite ran.

The original paragraph contains an old insertion with id7/author A. In each input pair the only edit is producer attribute v:note changing from before to after, with v bound to urn:producer. Five contexts were exercised:

| Context containing the metadata edit | Concrete property shape | Actual compare result |
|---|---|---|
| Protected paragraph pPr | `<w:pPr v:note="before"><w:jc w:val="left"/></w:pPr>` | Success, empty diagnostics |
| Protected paragraph mark rPr | `<w:pPr><w:rPr v:note="before"><w:b/></w:rPr></w:pPr>` | Success, empty diagnostics |
| Ordinary run rPr beside the old revision in that protected paragraph | `<w:r><w:rPr v:note="before"><w:b/></w:rPr><w:t>anchor</w:t></w:r>` | Success, empty diagnostics |
| Whole protected table | First pPr shape inside its cell paragraph, with the old insertion in that paragraph | Success, empty diagnostics |
| Whole protected block control | First pPr shape inside sdtContent, with the old insertion in that paragraph | Success, empty diagnostics |

All five returned `()`, left the original serialized bytes unchanged and silently retained the old metadata value. The three paragraph cases explicitly checked that before remained and after was absent. These observations do not allege corruption of the old revision or lost original metadata. The defect is accepting a meaningful edit inside a protected owner when the approved A4 contract requires a location-qualified atomic refusal. The ordinary non-held unsupported-formatting left bias does not authorize that bypass. This is the continuing source-authority boundary of D2, rather than a reopening of namespace transport or parser-context order.

Add a bounded source-aware held-owner check before typed equality authorizes raw passthrough. Its source must cover actual metadata roots inside paragraph, whole table and block-control ownership. Reuse existing physical spans/source inventories and namespace authority. Do not expand public OOXML fields, parse unknown attributes into a new model, or introduce generic semantic XML equality. Do not simply require identical source bytes, since benign redundant declarations, namespace enumeration order and compatible public comment edits already have valid positive controls. A changed opaque attribute value or binding is meaningful even when modeled formatting and the old revision payload are identical.

Regression guidance: retain the five actual RED contexts above, plus a benign declaration-order/redundant-closure positive and an ordinary non-held left-bias positive. Assert owner-qualified refusal, unchanged bytes, unchanged old revision identity, no output publication and applicable held-handle preservation. Keep the existing new-comment success and N3 equality controls. After repair, refresh affected scoped evidence and obtain another independent microscope.

## Evidence authentication and limits

Consumed root's independent `fx272-pass6-root-native-auth.json` and `fx272-property-successor-root-auth.json`, their underlying receipts/logs, and inspected the actual test assertions. Independently checked all native receipt source hashes against current files and all227product files in the successor source manifest. They match. No broad tests, builds, wheels or oracle runs were repeated during review. The additional observation above was the essential missing bounded read-only diagnostic against the authenticated installed artifact.

| Current local observation | Earned result and source |
|---|---|
| Native provider/comparison controls | 28 passed,0failed,2468filtered,8.90seconds in `fx272-pass6-publisher-final-native.log` |
| Full-event positive source projection | 1 passed,836filtered,0.02seconds in `fx272-pass6-publisher-final-projection-unit.log` |
| CLI comparison/count/workaround controls | 2 passed,100filtered,5.29seconds in `fx272-pass6-publisher-final-cli.log` |
| Lower revision equality | Applicable20-test observation from unchanged lower source, `fx272-identity-unit-green.log` |
| Scoped compile/lint/format | Four changed crates all-target check and all-target/all-feature clippy with denied warnings, plus fmt, each exit0 in `fx272-pass6-publisher-final-native-cli.json` |
| Fresh installed Python selected nodes | Same4nodes passed,643deselected,5.95seconds in `fx272-property-successor-python.json` and its test log |
| Metadata checks | Recorded prose0, adapters26in sync, clean diff whitespace |

The actual wheel SHA-256 is `4b71e9602a0440009c3450d58e053f3da87c2567c954149750eab2efed311a6d`. Wheel member and isolated installed native bytes match `b301c71fddfd1c62ed8b8ede223594739bc8c7961effe2087ac8de5e5fd6d1c5`. Manifest `37651b21a333f0812a4ac7800a9faaffc1d7d4e9a85222557d13aad2321bbe86` covers the unchanged227product inputs before/build/after tests. Root authentication includes shipped stub/__init__, the executed native/CLI test binaries and the actual embedded CLI application, not merely wrapper status. This is dev semantic qualification, not release archive/performance qualification.

Historical Base blanket-refusal RED, namespace failures, Python3-pass/1-fail predecessor, quoted-opening failures, row publication failures, same-owner duplication and setup/compiler/test-expectation failures remain separately attributed. Current successes do not retroactively turn those producers green. The supplied28+1+2+20+4 controls do not cover the new D2 diagnostics.

## Acceptance and lifecycle coverage

| A-ID | Reviewed providers, discriminators and result |
|---|---|
| A1 | Actual two-reviewer native/Python/CLI cases retain the earlier revision bytes/authors/ids and add new edits elsewhere. `regression_test.rs:408574`, `test_core.py:2390`, CLI case at `integration.rs:7402`. Covered locally for those inputs. |
| A2 | HeldRevisions maps physical paragraph/table/control/section ownership at `comparison.rs:2736`. Raw owner passthrough, id seeding, main/related/table/control/alias tests and footer/endnote/text-box controls examined. Supported unchanged cases pass, while D2 shows typed equality does not establish complete unchanged-owner identity. |
| A3 | New-id-only resolution and original-revision byte/order proof at `comparison.rs:1144` and `comparison.rs:3496`. Complete source properties, ordinary/final marks, first-shell/last-properties merge, local/default/value-only namespace controls and selective save/reopen examined. Actual named controls pass. |
| A4 | Held mutation, insert/delete/move, revised section and final predecessor mark guards examined. Existing atomic refusal controls at `regression_test.rs:408670` and `408832` pass. Opaque-root property changes bypass this policy as D2. |
| A5 | Changed revised-comment package replacement refuses. Ignored-story/comment selection, allocator exclusions, compatible public new-comment graph and saved results examined. Existing synthetic comment-comparison record still refuses before comparison at `comparison.rs:765`. No new unrestricted special-record capability is promised. |
| A6 | Stable CLI kind/part/owner coordinates, new-only multiset counts, same-author/no-op and ignored-box collision assertions examined. Native staging commits once after both postconditions. Python held paragraph/run no-op/refusal and checked publication are tested at `test_core.py:15348`, with implementation at `rdocx-py/src/document.rs:13860`. These local cases pass, with D2's missing refusal distinct. |
| A7 | Exact checked author-resolution workaround executed in Rust at `regression_test.rs:409574`, Python at `test_core.py:15348` and CLI at `integration.rs:7540`. Copies remain independent and earlier resolved markup is not falsely promised preserved. The early plan Contract closure sentence saying evidence is still needed is prospective design wording, superseded by the explicit qualified handback and actual receipts. No missing-evidence defect is inferred from it. |

The public signatures and defaults remain unchanged. Rust compare and compare_with_options still stage through the same native operation. Python's run default, false ignore defaults, None ignored-story default, tuple diagnostics and RdocxError route match the stub at `_rdocx.pyi:1215` and runtime at `rdocx-py/src/document.rs:13848`. Actual selected Python execution uses the public import. No new export/factory/property union or WASM authoring method was added. Full installed typing/runtime and WASM qualification remain pending.

README Honest boundaries, CLI README/help, stub docs and HLD03/HLD10 describe bounded preservation and the checked resolve-copy workaround. The D2 defect is implementation nonconformance to that stated boundary, not a request to change it.

## Shared invariant and equivalent publisher audit

**D1, new-only CLI story identity**: locally addressed. `commands.rs:1887` consumes a multiset, while `commands.rs:1934` uses kind, part and owner index plus revision metadata. It does not use the mutable whole-story fingerprint. Main counts and all-story records share this subtraction policy. The exact ignored-box/body collision, human/JSON, same-author, no-op and no-output refusal cases ran. No further D1 finding.

**N1, namespace meaning and relocation**: addressed on examined current paths. `comparison.rs:3054` detects producer w, `source_fragment:3096` and `source_child:3110` retain actual scope, and `IdAllocator::generated:651` qualifies fresh Word roots. Main/body extraction, related and owned stories, table/row/cell/control slices, fields, paragraphs, moved/final marks, equal property early returns and completed row publishers were traced. Default-Run controls now use source-aware publication at `comparison.rs:8395`, matching attributed control/run publishers at `comparison.rs:8691`. Exact retained run children bypass redundant closure under their actual retained shell in `preserve_run_source`.

Shared revision rendering derives source ancestry and actual destination scope at `revision.rs:594`, then uses it for ordinary recursion, wrapper removal, rejected properties and paragraph merging at `revision.rs:1417`. First shell, last pPr, child-local shadowing, absent default namespace and value-only compatibility prefixes have concrete direct controls. The producer-w row insertion/deletion/control tests assert actual generated Word identity, old raw owner bytes and resolution results. No remaining equivalent namespace publisher defect was established. The new D2 diagnostics change opaque values without changing bindings and do not reopen N1.

**D2 token/property sub-boundaries**: the original quoted greater-than corruption is repaired through reader-based `element_opening:5483`, reused by source-consuming insertion/shell/grouping helpers. The generated-only CT_P shell is not a separate untrusted opening scan. Both-input SourceSnapshots at `comparison.rs:3894` are operation-local, register only live immutable input owners, push once and pop after every Result exit at `comparison.rs:4080`. Nested calls isolate their frames. Owned PropertySide snapshots predate marker insertion, avoiding historical pointer/source lookup.

Property patching at `comparison.rs:4704` uses complete child slots and source-authoritative opaque matches, retaining QName/context and unknown order. Ordinary paragraph/run/table and final/control publishers share the path. Prior base pPr excludes independent mark/section fields, own-section precedence remains explicit, and complete prior rPr replaces formatting for ordinary and mark owners at `revision.rs:1111`. Direct role tests retain two intentional identical opaque siblings, another author's marker, a nested prior change, ordering, counts and save/reopen. No XML deduplication is used. Owner-local split-run assertions at `regression_test.rs:409986` distinguish legal formatting copies from same-owner duplication. These repaired sub-boundaries are locally demonstrated, but D2 remains open at the earlier protected-owner guard.

**Positive Empty/Start projection**: `empty_property_source:3654` checks actual direct Word ownership, lexical QName, encoded/decoded attribute values, attribute namespaces and scope before replacing the exact incidental raw boundary. `collect_property_source_events:3747` records text, CDATA, comments, PI and references between child spans. Nonempty or attributed rPr stays outside the special default-container path. The executed unit at `comparison.rs:10962` includes absent/changed metadata, changed namespace, encoded value, unknown child, non-element-event and rPr-shape negatives, plus paired/Empty and generated declaration-only rPr positives. This is a bounded private projection, not generic XML equality. Its intended default-container limitation explains why it cannot cure D2's nonempty protected property cases.

**N3, parser-context equality**: locally addressed at `rdocx-oxml/src/revision.rs:56`. Every other revision field remains compared. Only complete context enumeration order is ignored, retaining binding values/multiplicity, original parser input order, raw revision XML, nested/projected content and parse policy. No shared prefix parser was sorted or weakened. The applicable20-test evidence includes same-binding reordered context, meaningful foreign/Word/default differences, multiplicity, physical edits and raw changes. No further N3 finding.

## Inherited limitation and final evidence still due

The canonical paired-pPr unrelated-authoring probe in `fx272-pass5-property-projection-attempt.log` fails before comparison is invoked. The paired root metadata is absent from the typed model at `text.rs:10041`, and later modified main-part serialization at `document.rs:26204` can drop it. This is attributed inherited exploratory evidence, not a fixed feature, a desirable pinned behavior, or a reason to require an unauthorized global serializer/public-model project. Its exact probe remains in the evidence directory rather than as a deliberately failing normal test. Required compare/selective-resolution/save/reopen metadata controls remain present and passed for their inputs. Any newly introduced loss in that authorized lifecycle would remain blocking. D2 is independently established without invoking the unrelated-authoring edit.

Still pending for the final integrated union: native Office all-story revision and author-filtered artifact inspection, hosted CI, complete workspace tests/clippy, full installed Python3.9/3.12, strict typing/stubtest, WASM,49hashes/7pixels, actual README examples/current measurements, affected release archives/size and every approved producer-oracle/risk/metadata-policy rider. No pending result is labelled green or waived by this review.

## Smells

Zero. The concrete private source/snapshot types carry substantive state. No new trait, generic public surface, crate, module, source file, feature flag, forwarding wrapper or dynamic dispatch was introduced. The existing private count generic has two actual shapes today, and postcondition callbacks have actual accept/reject uses. No speculative second implementer is claimed.

## Nitpicks

Zero.

## Not found

No additional independent correctness, contract or OOXML defect beyond D2. No new panic finding after checking fallible reader boundaries, byte slices, source/span correlation, matched immutable/mutable paragraph traversal, frame lifetime, id allocation and staged publication. No separate test defect is raised because the missing protected-source cases belong to D2's discriminating regression obligation. No new public binding/default/export, cache/live-owner, atomicity or structure finding was established.

Only this review report was written. No remediation, source formatting, commit, push, wheel build, broad test or oracle occurred. Return control to the orchestrator with D2 blocking. Another independent full microscope follows remediation and refreshed applicable local evidence.
