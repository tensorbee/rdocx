# F-X272, namespace investigation, pass 6

**Role**: independent focused reassessment, review-only outside this report.
**Source**: dirty `work/f-x272-codex`, Base `c8acc502b74f29ed7e9546939f17007709a9f22b`. Tracked binary diff SHA-256 `91586c6a170fd5a8011756b4cdf35e526ee655a77469a3f4e0af67f79c3545a5` checked immediately before writing.
**Outcome**: D2 remains open at ordinary run-property rejection and the new empty-property projection boundary. N1's exact ignored-box reproduction now passes locally. Zero new independent smells or nitpicks. This is not full microscope pass2 or an integration-readiness verdict.

## Actual observations

Consumed `fx272-pass5-property-projection-attempt.log`. After a successful 63 second build, 3 tests executed, 1 passed and 2 failed, 2492 filtered, 8.18 seconds. No tests, builds, oracles or source edits ran during this review.

The ignored-text-box test passes, including the producer-host collision previously failing at header1 item0. The property matrix now gets through compare/reconstruction, including both self-closing metadata-empty directions and the paired metadata-empty case. Its remaining assertions expose rejection duplicates in ordinary and later-Word cases, plus global-count failures in the text-changing left-biased case. The separate canonical paired-pPr unrelated-edit control fails before comparison is called. Preserve this receipt as partial progress, not GREEN.

## D2, rejecting a complete ordinary run-property prior appends current opaque children again

Citations: `crates/rdocx/src/revision.rs:937`, `crates/rdocx/src/revision.rs:953`, `crates/rdocx/src/revision.rs:1100`, `crates/rdocx/src/revision.rs:1109`, `crates/rdocx/src/revision.rs:1137`.

`render_rejected_property_owner` renders the complete prior owner, then appends retained current children. The exception that treats prior rPr as a full formatting replacement applies only when the rPr parent is pPr. An ordinary run rPr does not satisfy that condition, so its current foreign child is retained and appended even though the newly source-authoritative prior already contains that child. The log shows two producer opaque elements inside the SAME rejected run rPr. This is a real duplication defect, not the legal creation of another run. The complete raw prior payload makes the existing resolver assumption observable on this comparison path.

The smallest shared repair belongs in the existing property-resolution role logic. A complete prior rPr represents rejected run formatting as well as rejected paragraph-mark formatting. Retention must distinguish formatting payload from independently retained revision markers, rather than append all current foreign children to an already complete prior. Preserve selected-scope accounting, nested/unselected revisions and the established marker rules. Do not deduplicate XML by string or namespace name, because two identical opaque siblings can be intentional. Do not remove opaque children from the prior payload to preserve the old typed-only workaround. No public model/API or new file is required.

Keep paragraph pPr base/mark/section precedence separate. The documented base-only pPr policy and legacy mark handling remain at `revision.rs:1020` and `revision.rs:1084`. This finding does not authorize replacing all property-owner rules with one indiscriminate discard rule.

### The left-biased fixture also overconstrains legal run splitting

Citations: `crates/rdocx/src/comparison.rs:6382`, `crates/rdocx/src/comparison.rs:6397`, `crates/rdocx/tests/regression_test.rs:409221`, `crates/rdocx/tests/regression_test.rs:409913`, `crates/rdocx/tests/regression_test.rs:409972`.

The fixture changes Anchor. to Changed. and b to i. Existing granular comparison splits content and clones run properties onto the resulting runs. Acceptance XML in the actual log has Changed and punctuation runs, each with one property owner and one retained opaque child. A document-wide multiset requiring exactly the original one run-property metadata occurrence rejects this legal result. Rejection additionally shows a genuine duplicate inside one of those owners, so fixing the fixture's counting must not conceal the resolver bug.

Use owner-local preservation checks for split runs: every resulting run has the intended modeled formatting and exact opaque source payload once, the runs cover the intended text in order, and adjacent equivalent run splitting does not change meaning. Keep global exact counts for paragraph/mark metadata and the unchanged old revision where those owners did not split. Ordinary unchanged-text and later-Word controls must still reject a second opaque child inside one rPr. Do not blanket-deduplicate the fact list, since that would hide the real failure.

## D2, the new source projection needs a complete empty-container boundary

Citations: `crates/rdocx/src/comparison.rs:3649`, `crates/rdocx/src/comparison.rs:3664`, `crates/rdocx/src/comparison.rs:3688`, `crates/rdocx/src/comparison.rs:3707`, `crates/rdocx/src/comparison.rs:3725`, `crates/rdocx/src/comparison.rs:3732`, `crates/rdocx/src/comparison.rs:10887`.

The private projection correctly addresses a concrete parser asymmetry: attributed Empty pPr is raw paragraph content at `crates/rdocx-oxml/src/text.rs:10237`, while Start pPr becomes typed at `text.rs:10041`. It positively records root QName, attribute QName, resolved namespace, encoded and decoded value, and effective scope. Replacing the parser's incidental raw-token distinction with that positive identity can be bounded to the existing source architecture. It is not authority to remove raw identity wholesale.

However, the helper tests only direct element spans when determining whether the pPr contains anything besides empty rPr shells. Direct text, CDATA, comments and processing instructions do not enter that inventory or the emitted source identity. For example, metadata-bearing paired pPr with one comment and the same pPr with a changed comment can receive the same identity. The original raw container is then excluded from the projected model, so no other check here accounts for that source difference. This is a source-proven preservation gap in the newly added projection, without a claimed runtime reproduction in this pass.

Use existing reader event boundaries to require the exact narrow empty-owner shape. Retain or explicitly account for every non-element event. Reject the special reconciliation when opaque content exists, or carry its exact bytes in the positive source identity. Do not add a generic semantic XML comparator. The existing unit checks Empty/paired agreement, missing metadata, changed value and added foreign element, but does not cover these non-element events or meaningful changed bindings. This 3-test log did not execute that unit test, so source assertions are not runtime GREEN evidence.

Required negatives are changed attribute namespace with the same lexical key/value, changed encoded opaque value, removed attribute, unknown child, changed comment/PI/CDATA/text, and nonempty or attributed rPr. Keep positive Empty/paired agreement and the actual generated declaration-only default rPr case. Preserve normal property and opaque signatures outside this specifically checked representation boundary.

## Paired-property save boundary and scope

Citations: `crates/rdocx/tests/regression_test.rs:410008`, `crates/rdocx/tests/regression_test.rs:410015`, `crates/rdocx/tests/regression_test.rs:410020`, `crates/rdocx-oxml/src/text.rs:10041`, `crates/rdocx-oxml/src/paragraph_properties.rs:1584`, `crates/rdocx/src/document.rs:26122`, `crates/rdocx/src/document.rs:26204`.

The new canonical w paired-pPr probe opens and saves, reopens, adds an unrelated paragraph, then saves. It never invokes comparison. The actual output drops the paired pPr and opaque root attribute. This matches existing parser/writer behavior: Start pPr is typed, the root attribute is not retained there, and the default typed property writer can emit nothing when the document takes the modified serialization path. Source-preserving unchanged saves do not prove modified serialization safe.

This is a real existing serialization boundary, but this new probe alone does not establish a feature regression or justify a global serializer repair. Do not expand F-X272 into a public CT_PPr model change or unknown-attribute parser solely to make it green. Attribute the pre-comparison baseline separately. If necessary, record a distinct follow-up through the orchestrator rather than silently changing this feature's scope or claiming the probe passed.

Conversely, the new source projection cannot substitute for physical preservation in the authorized comparison lifecycle. Keep positive source checks after comparison, selective accept/reject, save and reopen, for paired and Empty metadata owners on both sides. The matrix currently reaches these checks without recorded metadata-empty failure. Verify actual metadata/root namespace and opaque children in saved parts, not only typed equality or the helper's identity. A later unrelated authoring edit is useful differential evidence, but must be compared with the identical pre-comparison baseline. Report an inherited limitation distinctly from any newly introduced loss. A demonstrated new loss within compare/resolve/save/reopen remains D2 and cannot be waived as a baseline problem.

## Bounded next qualification

1. Correct complete-rPr replacement in the shared resolver. Prove ordinary run and paragraph mark rejection, with foreign opaque siblings, intentional duplicate siblings, marker ordering and applicable unselected/nested revision controls. Preserve accepted and rejected saved outputs and counts.
2. Make the text-changing fixture owner-aware while retaining the SAME-owner duplicate discriminator. Keep contrasting, later-Word, metadata-empty, paired and absent controls intact.
3. Tighten the private projection's exact content boundary and execute its focused positive/negative unit. Preserve the actual 3-test receipt and the pre-comparison save failure separately.
4. Refresh the applicable scoped native/CLI qualification and authenticated successor Python only after source is corrected. D1/N3 are not reopened. The successful exact N1 ignored-box and earlier foreign-w row controls must remain green.

No new user product choice, public API/model, source file, trait or broad serializer project is authorized by this report. Full independent feature microscope and final integrated native/hosted/installed/typing/WASM/hash/pixel/archive/risk gates remain pending.

Only this report was written. Return control to root for a distinct implementation phase.
