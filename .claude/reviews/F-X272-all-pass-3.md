# F-X272, all aspects, pass 3

**Reviewer**: independent Codex reviewer, not an implementer of F-X272. Prior focused investigations are historical escalation records, not substitute full reviews.
**Reviewed**: complete working diff on `work/f-x272-codex`, Base `c8acc502b74f29ed7e9546939f17007709a9f22b`, 22 tracked files, 6479 added and 764 removed lines. Tracked binary diff SHA-256 `cf667bf5776f247f26329f0780bd6ce1496333d0650536af1dbec71119baef7d` remained unchanged through review. Historical untracked review reports are excluded from product counts.
**Contract**: canonical d3a028e1, acknowledged in the annotated worker plan. Later36a95110 concerns unrelated F-X274 preparation. F-X272 remains in-progress. Reviewed repository/microscope workflow, full contract/progress, HLD03/HLD10 intent, previous full passes and focused reports.
**Verdict**: 1 defect, 0 smells, 0 nitpicks. D2 remains blocking at opaque-root namespace closure in the residual guard. No scoped integration-readiness approval or final acceptance.

## Defects

### D2, opaque residual identity rejects benign source-root namespace closure

`crates/rdocx-oxml/src/properties.rs:849`
`crates/rdocx-oxml/src/properties.rs:850`
`crates/rdocx-oxml/src/properties.rs:852`
`crates/rdocx/src/comparison.rs:2778`
`crates/rdocx/src/comparison.rs:4251`
`crates/rdocx/src/comparison.rs:4262`
`.claude/plans/F-X272-design.md:259`

The new bridge correctly separates modeled attributes from omitted metadata on known nodes. Its UNKNOWN-node branch instead records the entire literal source blob plus effective scope. A raw carrier's public save can move a namespace declaration from its property ancestor onto the opaque root while retaining exactly the same expanded names, effective root scope, opaque attributes and payload. Literal blob comparison treats that supported closure as a meaningful protected-owner edit and refuses.

This is an actual compatibility regression, not an inference from a hypothetical XML rewrite. Essential bounded read-only public-API diagnostics used fresh isolated Python subprocesses with authenticated native paths/hashes:

- Current installed module `/private/tmp/rdocx-s96-fx272-py312/lib/python3.12/site-packages/rdocx/_rdocx.abi3.so`, SHA-256 `32560b83cdfeea04e23910b7b40fa31abe4f2c4ff91ca15f3e5190ccccef105a`.
- Authenticated unpacked predecessor `/private/tmp/rdocx-s96-run-refresh/fx272-b301-predecessor-package/rdocx/_rdocx.abi3.so`, SHA-256 `b301c71fddfd1c62ed8b8ede223594739bc8c7961effe2087ac8de5e5fd6d1c5`. Root's extraction receipt binds it to wheel4b71e960. The predecessor package was loaded through sys.path before import in a separate -I -B process, without modifying the installed current package.

Each input has the canonical document root namespace set, an unchanged insertion id7/author A in paragraph0, and a second ordinary paragraph. Public add_comment targets ONLY paragraph1, then save/reopen materializes the normal raw-carrier closure. The three actual property shapes and observed rewrites were:

| Protected paragraph property source | Edited source after the public comment save |
|---|---|
| `<w:pPr xmlns:v="urn:producer"><v:opaque v:note="same"/></w:pPr>` | `<w:pPr><v:opaque v:note="same" xmlns:v="urn:producer"/></w:pPr>` |
| `<w:pPr><w:rPr xmlns:v="urn:producer"><v:opaque v:note="same"/></w:rPr></w:pPr>` | `<w:pPr><w:rPr><v:opaque v:note="same" xmlns:v="urn:producer"/></w:rPr></w:pPr>` |
| `<w:pPr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:future q:note="same"/></w:pPr>` | `<w:pPr><q:future q:note="same" xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/></w:pPr>` |

All three predecessor comparisons succeed with empty diagnostics. All three current comparisons refuse with `comparison cannot change body/paragraph[0], which holds an existing tracked revision: the edited document changes it`. The old insertion's exact source bytes survive the edited input in every case. Plain and root-inherited opaque-child canonical-root controls succeed on the current artifact. Thus the fault is the source-capable local-binding publication boundary, not every opaque child or the old revision itself.

Earlier stripped-root exploratory inputs also refused, but are not used as regression evidence. The paired diagnostics above use canonical root declarations, exclude unrelated namespace-context changes, and provide actual predecessor success. No suite, build, oracle, package extraction or artifact file was created by this reviewer. All diagnostic documents stayed in memory.

Repair the one opaque-root residual boundary coherently. Positively check actual expanded root/attribute identity and effective source scope, and permit only benign declaration plumbing of the opaque ROOT. Retain exact non-declaration opaque attributes, encoded values, inner payload, event order, sibling order and multiplicity. Preserve descendant-local shadowing and value-only namespace dependencies. Do not drop blobs wholesale merely because modeled raw-carrier equality happens to cover these examples. The residual serves additional opaque branches whose typed equality may omit data. Do not add generic XML semantic equality, arbitrary subtree canonicalization or a blanket namespace exclusion. Root namespace plumbing can be treated more narrowly than the opaque subtree it surrounds.

Keep original protected owner emission and old-revision byte checks exact. A permitted edited-side closure is not permission to rewrite the emitted original. Add lower bridge positive/negative controls for foreign and unknown-Word opaque roots, plus the three actual public comment regressions above. Pair them with changed effective URI, changed opaque attribute/value/encoded bytes, changed inner events, reordered/removed intentional duplicate siblings and descendant-local shadowing. Include whole table/control or related-story use of the same residual branch. Retain the known-attribute normalization and17meaningful-edit refusal controls. D2 closes only after this correction is implemented, qualified and independently reviewed.

## Evidence authentication and applicability

Consumed `fx272-residual-root-auth.json`, underlying `fx272-residual-final-native-cli.json`, successor Python receipt/manifests and actual test assertions. Independently compared all241current source/test inputs in the native receipt and all227product inputs in the successor manifest. No mismatch was found. Source remained frozen. Recorded Rust binaries include the actual embedded CLI application, not just a test wrapper.

| Actual current observation | Result |
|---|---|
| Native provider/comparison selection | 30 passed,2468filtered,8.49seconds |
| Lower selection | 88 passed,679filtered,0.85seconds, including residual units, numbering and unchanged revision identity |
| Positive Empty/Start projection unit | 1 passed,836filtered,0.02seconds |
| CLI comparison/count/workaround selection | 2 passed,100filtered,5.09seconds |
| Four changed crates, all-target check | Passed,39.07seconds |
| Four changed crates, all-target/all-feature clippy with denied warnings | Passed,30.70seconds |
| Formatting | Passed |
| Fresh selected Python3.12.13 | SAME4nodes passed,643deselected,8.47seconds |
| Metadata checks | Recorded prose0, adapters26in sync and clean diff whitespace |

Fresh wheel SHA-256 `541224310d3d7d2bcbc255f3e2d04a797ccc421ccdacbd2a4e5868df61f3ce0b` and native/member/import SHA-256 `32560b83cdfeea04e23910b7b40fa31abe4f2c4ff91ca15f3e5190ccccef105a` are bound to unchanged product manifest `ee557b98818464ca1fb6806e299e0ffdce250a8047042d4de8d12c7ad42d51a2`. Root separately authenticated the shipped stub/__init__, isolated actual import and four passed nodes. This dev artifact qualifies the focused semantics, not release measurements.

The public paragraph bool/numeric and valid table-v2 normalization cases are genuinely current GREEN. `regression_test.rs:410201` exercises the actual adjacent comment route after checking true-to-bare-b and00120to120 normalization. `regression_test.rs:410496` places the table before terminal sectPr, checks four numeric rewrites, then asserts the ORIGINAL whole table remains in compared and selectively resolved saved outputs, with old7/A and the intended comment graph. The earlier unsuffixed table setup after terminal sectPr remains invalid setup, not compatibility evidence.

All historical RED receipts remain attributed: blanket refusal, namespace/quoted-opening failures, lower identity/Python predecessor failure, same-owner duplication, meaningful protected-source edit bypasses, all-attribute canonicalization refusal and unsupported-border value-unit failure. Setup/compiler/test-only failures are not semantic evidence. Passing the supplied selections does not discharge the new D2 opaque-root regressions.

## Full contract and consumer audit

| Acceptance | Reviewed implementation and actual evidence |
|---|---|
| A1, second review beside the first | Native/Python/CLI two-reviewer reproductions retain raw earlier revisions and create new noncolliding revisions elsewhere. Main reproduction, public comment and same-author cases remain present. Covered locally for supplied inputs, with D2 exposing another supported comment input. |
| A2, physical owner/story preservation | HeldRevisions records actual original/edited paragraph, whole table/block-control and standalone section ownership. Its residual supplement runs before typed passthrough in main and equivalent related callers. Aliased related stories, footer/endnote/text-box and nested table/control raw preservation controls remain. D2 prevents complete compatibility closure. |
| A3, new-only resolution and save/reopen | Input-scoped allocator seeds both sides, excluding ignored stories. Postcondition copies resolve newly allocated ids only, then prove accepted/rejected modeled and source-backed projections. Original revision ranges are checked byte for byte and in order at `comparison.rs:3518`. Prior roles, merge relocation and actual selective save/reopen controls examined. |
| A4, protected edits refuse atomically |17meaningful-source contexts cover root/child properties, whole table/control, related properties, Empty section pgMar, omitted content attributes and opaque events/value-only bindings. Extended Rust/Python/CLI refusal assertions preserve bytes/identity/handles/output. Known normalization must remain supported, with D2 now blocking one benign case. |
| A5, comments and ignores | Revised whole-comment replacement guard, special existing comment-comparison-record refusal, compared-story exclusion and id seeding remain conservative. Public comment graph cases and ignored source controls pass. No unsupported special-record editing capability was added. |
| A6, counts/no-op/publication/handles | Stable CLI kind/part/owner identity and new-only multiset counts, human/JSON same-author and ignored-box collision, exact native no-op, late refusal and Python held/fresh publication controls examined. Actual current Rust and Python selections retain these cases. |
| A7, checked workaround | Exact public Rust/Python author-resolution and CLI resolve-both-copies route remain executed, with independent originals preserved. Documentation explicitly states resolved old markup is discarded. Contract closure now cites the actual evidence. |

Public compare signatures/defaults remain unchanged. Existing Python compare still stages a candidate, maps errors to RdocxError, produces tuple diagnostics and publishes through checked owner/cache invalidation only when changed. Current installed tests use the public import. README Honest boundaries, CLI README/help, stub and HLD03/HLD10 retain bounded preservation and the checked workaround. Final typing/runtime qualification remains due.

The sole additional export is the explicitly approved doc-hidden `properties::protected_source_residual(xml, inherited_bindings)` at `properties.rs:547`. Its output is documented as opaque internal identity, not persistent interchange or authoring. No public model fields/types, authoring operations, Python/WASM methods, new source file, crate, trait, dynamic dispatch or feature flag were added. The prior no-new-API wording is explicitly superseded by the narrow canonical technical exception, rather than silently ignored.

## Shared implementation and historical finding disposition

**D1, CLI identity**: addressed on examined paths. Kind/part/owner index plus revision metadata determine subtraction, without mutable story fingerprint. Main tuple and all-story keys remain two concrete instantiations of the existing multiset helper. Current CLI2tests retain same-author, no-op, refused output, ignored-box collision and checked workaround membership. No new D1 finding.

**N1, namespace transport/publication**: addressed on the previously repaired generated/retained paths. Source fragments/children retain real context. Fresh canonical Word containers use the existing qualifier under producer-w bindings. Paragraphs, fields, rows/cells/tables, moved/final marks, default-Run and attributed inline controls, equal property branches and recursive control publishers were rechecked. Retained run children remain exact under their retained shell. Shared revision ancestry/destination transport at `revision.rs:596` reaches ordinary recursion, wrapper removal, properties and first-shell/last-properties paragraph merging at `revision.rs:1415`. Direct controls include child-local shadowing, absent default namespace, legacy property QName and value-only compatibility prefixes. D2 is an overly strict guard on equivalent source, not a newly established corruption in these publishers.

**N3, parser context identity**: addressed. The lower revision equality implementation still compares every other field exactly and ignores only complete-context enumeration order. Binding values/multiplicity, stored parser order, raw revision XML, projected/nested content, physical edits and parse policy remain authoritative. Those units execute within the current lower88selection. No new N3 finding.

**D2, earlier repaired sub-boundaries**: reader-based opening boundaries at `comparison.rs:5523` cover literal quoted greater-than safely. Both-input source frames at `comparison.rs:4102` register live immutable parse owners and pop on every Result exit. Owned pre-marker property sides retain source authority through ordinary/final/move/control paths. Complete source-child slots in `patch_property_source:4775` preserve unknown anchors. `preserve_run_source:4883` and `publish_run_source:8732` keep exact retained children and source property shells. Complete prior rPr replacement at `revision.rs:1100` preserves independently scoped markers, nested prior changes and intentional duplicate opaque siblings without XML deduplication. Owner-local text-split assertions still distinguish legal copies from duplication inside one rPr. These local repairs remain applicable, but D2 stays open at the new residual UNKNOWN-node branch.

**New lower source authority**: inspected owning paragraph/run/section/table/control/content classifications at `paragraph_properties.rs:2221`, `run_properties.rs:2030`, `document.rs:4375`, `table.rs:3031`, `content_control.rs:2296` and `text.rs:15994`. Known attributes and unsupported-value checks follow actual readers rather than merely validator admission. Unknown Word/producer keys remain residual. The existing numbering PropertyKind/property_attributes/on-off parser changes are crate-private visibility only. Their bodies and serialization policy are unchanged, with current numbering tests included. No second upper schema was introduced.

The complete event scan retains opaque text, CDATA, comments, PI and reference bytes. Stable expanded-name occurrences and default-only-container handling avoid counting omitted known defaults against unrelated residuals, while attributed or nonempty owners stay significant. Residual units at `properties.rs:897`, `:917`, `:977`, `:1002` check numeric/toggle normalization across families, mixed unknown keys, unsupported values, encoded bytes, foreign lookalikes, full events/order/multiplicity, scope and default association. The benign namespace unit compares a KNOWN metadata-bearing root. It does not exercise the UNKNOWN opaque-root literal path exposed by D2.

The source-backed Empty/Start projection at `comparison.rs:3677` separately retains positive root/attribute/binding identity and full events through `collect_property_source_events:3769`. Nonempty or attributed rPr stays outside default-container reconciliation. It is not a general XML comparator and does not cure or justify D2.

## Inherited limit and pending final gates

The canonical paired-pPr unrelated-authoring probe fails before comparison is invoked at the existing typed main-part serialization boundary. It remains explicitly attributed in the plan/progress and immutable evidence, not claimed fixed, deliberately pinned as desirable behavior or turned into an unauthorized general serializer project. Required comparison/selective-resolution/save/reopen metadata cases remain actual controls. Any new loss in that authorized lifecycle would remain blocking. D2's predecessor/current diagnostic uses only the already supported public comment lifecycle and is independent of that unrelated-authoring limitation.

Final integrated obligations remain pending: native Office all-story and author-filtered artifact inspection, hosted CI, full workspace gates, complete installed Python3.9/3.12, strict mypy/stubtest, WASM,49hashes/7pixels, actual README examples/current measurements, affected release archive/size, producer oracles and all approved risk/metadata-policy riders. No absent rider is marked green or waived by this review.

## Smells

Zero independent smells. Private residual node/item carriers hold substantive source state. The bridge is concrete and authorized. No speculative generic/trait framework or forwarding-only layer was introduced.

## Nitpicks

Zero.

## Not found

No additional independent correctness, contract or OOXML defect beyond D2. No new panic finding through the admitted facade paths after inspecting reader-event boundaries, byte ranges, source/model correlation, matching immutable/mutable traversal, frame lifetime, allocation and transactional publication. No separate test finding is raised for D2's missing UNKNOWN-root namespace discriminator. No additional public-default/export, live-owner/cache, count, workaround or atomicity finding was established.

Only this report was written. No remediation, source/test formatting, commit, push, build, broad test, oracle or reviewer-created artifact file occurred. Return control to the orchestrator with D2 blocking. Another independent full microscope follows remediation and refreshed applicable evidence.
