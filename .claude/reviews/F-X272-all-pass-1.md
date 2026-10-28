# F-X272, all aspects, pass 1

**Reviewer**: independent Codex reviewer, not an implementer of F-X272. Earlier namespace investigations are historical and do not substitute for this pass.
**Reviewed**: complete working diff on `work/f-x272-codex`, Base `c8acc502b74f29ed7e9546939f17007709a9f22b`, 14 tracked files, 2830 added and 417 removed lines. Tracked binary diff SHA-256 `0a141c5f60878b505a0e4d59d78775f5b134e7fd0a288ea526c978fb9d509f51` remained unchanged throughout review.
**Contract**: approved F-X272 plan, including owner acknowledgments of canonical `59dc1f48`, `10d45b11` and `0987372c`. Q-S96-01 remains the bounded unchanged-owner preservation policy and checked resolution workaround.
**Verdict**: 2 defects, 0 smells, 0 nitpicks. Scoped integration readiness is not approved.

## Defects

### D1, new-only CLI story counts use package-part identity instead of story identity

`crates/rdocx-cli/src/commands.rs:1935`
`crates/rdocx-cli/src/commands.rs:1898`
`crates/rdocx-cli/src/commands.rs:1952`
`crates/rdocx/src/document.rs:1226`
`crates/rdocx/src/comparison.rs:2646`

The subtraction key for all-story revisions includes only part name, revision id, author and timestamp. A package part can contain multiple stories. Text boxes share their host part, while their StoryId also distinguishes kind and owner index. `story_part_revisions` attaches the innermost physical story at `crates/rdocx/src/revision.rs:443`, and the CLI later groups the surviving values by the complete returned story at `commands.rs:1957`.

A supported counterexample is an ignored text box with old revision id0, author Ada and timestamp T, plus an ordinary body paragraph before that text box whose text changes. Compare with `--ignore-story text_box`, the same author Ada and timestamp T. Ignored text boxes are deliberately excluded from id seeding, so a new body deletion can receive id0. The old text-box revision must remain unchanged and must not appear in newly created story counts. Because both values share the key, the multiset consumes the earlier new body revision as if it were old, then retains the old text-box revision as if it were new. The aggregate revision total can remain correct while JSON per-story counts and the human number of stories are wrong. With a body deletion/insertion pair, the CLI can report two stories instead of one body story.

This is a source-traced deterministic collision, not a newly executed test in this review. It violates A6 and the consumer-map promise of new-only all-story human/JSON counts, combined with the approved ignored-story id scope. The current CLI same-author case uses only body revisions and cannot discriminate it at `crates/rdocx-cli/tests/integration.rs:7402`.

Use the stable physical story coordinates in the subtraction identity, including kind and owner index as appropriate, rather than only its part name. Do not use the full fingerprinted StoryId blindly across mutation, because its fingerprint is part of its equality at `document.rs:1230`. Add a current-source failing CLI control using the same author AND timestamp and a deliberately colliding ignored-text-box id. Assert correct body-only story counts, total/new main counts, unchanged text-box revision XML and both JSON/human output. Keep the existing same-author/no-op/refusal and checked-workaround controls.

### D2, paragraph-mark insertion still mistakes a quoted greater-than for the opening-tag boundary

`crates/rdocx/src/comparison.rs:4347`
`crates/rdocx/src/comparison.rs:4362`
`crates/rdocx/src/comparison.rs:9575`
`crates/rdocx/src/comparison.rs:9587`

When a final paragraph change needs to mark a preceding raw paragraph without pPr, `marked_paragraph_xml` inserts the generated properties after the first literal `>` character. XML permits that character inside a quoted attribute value. For a supported predecessor such as `<q:p x:note="a > b">...`, it selects the greater-than inside x:note and inserts the property markup into the attribute value. The candidate then fails parsing or its reconstruction condition even though the requested final paragraph change is otherwise supported and the opaque attribute should remain intact.

The equivalent existing-rPr route also bypasses the parser boundary. `with_mark_revision` finds the first literal `>` and uses that position when no preceding marker child is selected. An rPr with an opaque quoted attribute containing `>` and a normal formatting child can therefore insert the mark inside that attribute. This second caller needs the same invariant, not an input-specific fix to the paragraph case.

These are source-traced defects, not newly executed RED cases here. The raw-source paths are actual consumers: the predecessor is closed under its source bindings and sent to the first helper at `comparison.rs:4333`, and an existing property child reaches the second through `comparison.rs:4378`. The repair already changed `paragraph_shell` to obtain its opening boundary from the XML reader at `comparison.rs:4708` and `comparison.rs:4720`, but these copied insertion predicates still use a lexical search. This is a distinct token-boundary issue, not a new loss of namespace meaning in the repaired relocation renderer.

Reuse an actual parsed opening boundary for every source-capable mark insertion. Preserve the original QName, attribute bytes, empty-element handling and marker schema order. In the existing regression carrier, obtain actual pre-repair outcomes for both no-pPr and existing-rPr cases with a valid quoted `>` attribute, then require successful final insertion/deletion, retained opaque value and qualified property/marker identity through accept/reject and save/reopen. Keep the existing refusal when the predecessor itself holds an old revision. No broad XML validator or new public surface is needed.

## Evidence consumed and applicability

The complete implementation, tests and public documentation diff were read. Current hashes of the source files named by `fx272-identity-corrected-frozen-native-cli.json` match its frozen producer. That receipt records native 23/23, CLI 1/1, scoped all-target check and fmt, each with actual exit0 and executed binary identities. The 23-test log explicitly includes all original Issue425 preservation/refusal cases, the public new-comment case, both ignored-source controls, direct relocation/legacy/value-prefix controls and the six applicable pre-existing resolution controls. The unit log records 20/20 lower revision-module tests, including complete-context identity, meaningful binding changes, multiplicity, physical projection and accepted-edit behavior. Four-crate clippy and changed-carrier lint remain applicable to this frozen source as recorded in the handback.

Fresh successor wheel `bb501a120f9f1612786d56f2cb892ab115e82564f9d647e167507f86b0ddfe1a` and actual imported native `7a36134634571349e8de2ba91eefc947eec98d65249ae0e34bff7cdbda81e353` are bound by `fx272-identity-successor-python.json` to unchanged before/after source manifest `40e690f52b64bb9bbdbd7e35aecc128d57962aef096ee6b4b725593c08abdc1d`. The same four Python nodes passed, 643 deselected, 5.04 seconds. The failed authenticated predecessor remains a 3-pass/1-fail observation rather than being relabelled. The Base blanket-refusal RED, intermediate namespace/legacy/compatibility failures and setup/compile failures are separately retained.

The reviewer consumed these applicable observations rather than rebuilding a wheel or rerunning suites. No product tests, compilation, oracle or broad gate was executed during this review. Passing supplied controls does not cover D1 or D2.

## Acceptance and consumer coverage

| Acceptance | Examined implementation and evidence | Result for this pass |
|---|---|---|
| A1, earlier redline plus later edit | New-id-only postconditions, actual Rust/Python/CLI reproduction and old author/id/raw assertions | Covered locally for the named cases |
| A2, owner/story preservation | HeldRevisions physical owner map, section guard, raw copy route, namespace scans and original-byte postcondition. Table/control, header/footer, footnote/endnote, text-box and alias controls | Covered locally for supplied cases, with D2 marking boundary still open |
| A3, new-author accept/reject | Shared source/destination renderer, first shell/last properties, child-local shadowing/default namespace/value-only bindings, author/id paths and save/reopen | N1 addressed on the reviewed paths and actual controls |
| A4, unsupported revised-owner mutations | Changed/deleted/inserted/moved owners and final predecessor mark guard, atomic native/Python/CLI controls | Bounded policy retained, no authorization for broader edits |
| A5, comments and ignores | Physical revised-comment package replacement guard, new comment beside old body revision, ignored scope selection and id seeding | Covered locally for named cases. Pre-existing special comment-package record refusal remains conservative |
| A6, counts/no-op/late refusal/live owners | New-only multiset, zero counts, same-author counts, exact no-op and held Python paragraph/run controls | D1 remains open. Other named controls passed |
| A7, checked workaround | Existing Rust author resolution, exact public Python calls, three CLI commands and independent input preservation | Executed locally, no false claim that resolved old markup survives |

The special `/customXml/rdocxComparisonComments.xml` boundary is still explicitly refused before id seeding at `comparison.rs:765`. The parent also independently observed safe refusal through public current-wheel APIs. A theory of overwriting an older synthetic comment comparison record is refuted by that guard. A2 does not promise unrestricted editing of that special record. This is not an added defect or a reason to expand the user's bounded scope.

## Historical finding disposition

- **N1, source namespace meaning and relocation**: addressed on the examined main/related paragraph, field, table, control, mark, promotion and replacement paths. The renderer now closes each transported child from its actual ancestry against the actual destination. The direct chained, foreign-w, shadow, empty-default, legacy-lookalike and value-only prefix tests assert expanded identities and resolution outcomes. D2 is the separately identified copied opening-boundary predicate, which must also be repaired.
- **N2, lost original alias reproduction**: addressed. The original plain q-root and foreign-body lookalike are separate controls at `regression_test.rs:408885`.
- **N3, ordered parser metadata identity**: addressed. `rdocx-oxml/src/revision.rs:57` compares all nine other stored fields exactly and compares complete context entries without order dependence, preserving multiplicity and stored parser order. Marker/nested/raw/projection/policy controls at `revision.rs:991`, `revision.rs:1020` and `revision.rs:1054`, plus the fresh public comment reproduction, qualify the actual repair.

## Not found

**Contract**: no unapproved public API, dependency, feature flag, crate, module or source file was introduced. The bounded policy, unchanged signatures, ignored-story id scope and final acceptance debt remain explicit. README, CLI README/help, Python stub and HLD03/HLD10 describe current intent. Donor archive measurements and published versions were not adopted.

**Panics**: no new independently established panic finding. Inspected source range arithmetic derives from the same reader-owned byte spans, aligned indices come from current owner arrays, and parent chains are constructed by the tree parser. Existing resource/validation boundaries remain in place.

**OOXML beyond D2**: no additional established schema-order, namespace relocation, opaque-subtree or whitespace defect. Raw revised owners bypass source closure, generated Word markup receives local qualification when required, and old revision preservation is checked before publication. Property marker insertion retains the existing ordered marker policy.

**Structure**: zero smells. The concrete multiset generic has actual tuple and StoryRevision uses, and captured postcondition callbacks have actual accept/reject instantiations. New aliases describe existing complex tuples. Helpers perform real selection/context work rather than forwarding alone. The equality implementation defines no new trait. No new tracked source/test carrier or link binary was added.

**Tests**: no manufactured test finding. The actual input/assertion matrix discriminates the named historical defects, validates public operations and preserves original failing observations. D1 and D2 need their own pre-repair controls. The new native comment test's resolution count is tied to the public inventory while independently fixing old-author counts, bytes and comment-graph outcomes, rather than asserting an invented one-record count.

**Smells**: zero. **Nitpicks**: zero.

## Remaining gates and handback

Fix D1 and D2 in a distinct implementation phase, preserving their IDs, record actual RED/GREEN, refresh affected scoped checks/artifacts and obtain a new independent microscope pass. Do not remediate inside this review. No user scope choice is needed for these concrete supported-input fixes.

Final native Office all-story/author-resolution observations, hosted CI, full workspace and complete installed Python3.9/3.12 suites, strict mypy/stubtest, WASM, 49 hashes/7 pixels, actual README examples and refreshed measurements, archive dry-run/size, supply-chain and all approved risk/metadata riders remain pending on the integrated union. They were neither performed nor waived by this scoped review. No completion, integration, ledger, commit, push, release or native-parity claim follows from the local evidence.

Only this review file was written. Return control to root now.
