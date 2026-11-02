# F-X283 design, Codex final 1

**Reviewed**: the complete existing 766-line canonical draft, SHA-256 `8728a36273ff468f45cb8ed955b0ed554266dfe02eeddaaf956442303568b121`, and its current native/binding/support consumer mechanisms on `sprint/s96` at `aff3f2b8744ed86755db8520c9bf6abeed5f2038` with preserved carry.

**Verdict**: retain draft pending the remaining F9 publisher-inventory repair. One medium defect, zero smells, zero nitpicks. This is the Root-delegated independent final design audit, not a fourth design peer, implementation acceptance or Root approval. All three design peers remain used. No counter was reset.

## Defects

### F9, medium: dormant cache policy still omits direct image publishers that reopen native state

The repaired lifecycle explicitly invalidates dormant facts/pieces after `set_hyperlink_url` at `.claude/plans/F-X283-design.md:436` and after `save`/`to_bytes` at `:515`. It otherwise routes positive mutations through guarded publication at `:414`. It does not specify the same policy, or an exact proof that it is unnecessary, for three existing public image operations that also bypass those guarded routes.

Supported no-handle entry points are `Document.replace_image(relationship_id, data)` at `crates/rdocx-py/src/document.rs:13277`, `replace_image_for_story(story, relationship_id, data)` at `:13289`, and `set_picture_sizes_for_relationship(relationship_id, width, height)` at `:13429`. The story argument can be a plain caller-constructed Story. Its unbound resolution does not establish active authority, as shown by `validate_owned_story` at `:17289`. None of these wrappers calls `prepare_word_revision` or clears binding facts/pieces.

The native image replacement path stages privately, invokes `prepare_and_reopen_staged`, then installs the reopened candidate at `crates/rdocx/src/document.rs:31210`. Native relationship-wide resize likewise changes main-story source, reopens privately and commits at `:36771`. These are native model/package replacement boundaries, rather than simple settings or style-value changes whose Word owner facts can be shown unaffected. They therefore have the same current-owner normalization concern that the plan already recognizes for serialization at `.claude/plans/F-X283-design.md:517`.

Under the proposed dormant representation, open-time validated addresses/projections remain present at plan `:324`. First activation only rebuilds them when they were invalidated at `:369`. Following the current named inventory literally permits old dormant facts to survive these native replacements without any specified requalification. The design has not closed that publication boundary. This is a source-derived contract omission, not a claimed executed stale-handle regression or proof that every image replacement changes an address.

Expected repair: explicitly inventory all three routes. Preserve their current native validation/error order and success semantics. Conservative dormant facts/pieces invalidation after native success is sufficient and need not activate authority or advance its serial. Both native paths stage and reopen before installing, so their returned failures leave live native state unchanged and do not require discarding valid old facts. Preserve active behavior and held-handle semantics. Include fresh-open no-handle success/refusal controls followed by first authority access, native/fresh/reopened observations and appropriate normalization-sensitive input. Do not impose blanket invalidation on unrelated settings/style APIs without tracing whether their owner facts change.

This retains stable F9 because it is the same under-enumerated ingress/publication invariant. The original hyperlink removal counterexample is repaired, rather than re-reported as still failing.

## Prior F9 targeted repair

The unbound-record removal repair at `.claude/plans/F-X283-design.md:426` now matches current call order: resolve the native record, remove on a private candidate, activate before the physical `self.owners` lookup, then publish. Current source does the lookup at `crates/rdocx-py/src/document.rs:15469` before the later publication helper. C15 now exercises the actual plain public constructors before any checked factory. Invalid/ambiguous native record errors remain before activation. The adjacent direct URL setter has its own success-only dormant cache policy and does not invent a revision. This closes the concrete peer3 removal boundary and the named URL-publisher concern.

## Mechanism and acceptance review

- The three additive hidden Word0.17.0 facade signatures are concrete. Local cell preparation uses existing singleton CT_Body resource validation, distinct body slot/table ordinal, one table candidate, checked fallback and one nonfallible table replacement. The binding never needs private Cell storage across crates. The consuming ordered replacement bridge retains actual replace_batch and staged-origin behavior without claiming replace_batch itself reopens.
- Source qualification retains the actual main-part Arc/key/name and namespace facts, with one keyed lookup and identity checks per setter. It does not depend on F282's shared-part getter or claim that a cached None alone proves source safety. Creation/open/source replacement/flush custody and the positive Document.new workload are explicit. Native resource limits and invalid-source fallback remain implementation proof obligations.
- Selected-table owner/index deltas distinguish parent, route and generation memberships. Body and table ordinals are not conflated. Births are reserved, disappearing owner records take fallback, no vector deletion shifts unrelated owners, first tcPr relocation uses direct paragraph ordinals, and text retirement is separate from property publication. Piece replacement occurs after the last fallible commit without a live splice or whole-document install.
- Dormant construction retains the existing eager fallible inventory and related-part checks, storing immutable facts rather than falsely promising zero inventory work. The existing support initializer is explicitly amended for nonzero dormant serial while retaining initialized/unique-epoch guards and allocation-before-state-change. Activation preserves serial/origin, uses reserved maps/vectors and does not invent a commit. A live preparation or held current epoch still blocks initialization.
- The ten factory boundaries, six earlier splice callers, immutable checked descendants and unbound hyperlink records have explicit borrow/activation dispositions. No additional concrete missing activation factory was established. The remaining direct publisher cache issue is F9 above.
- Dormant replacement covers no-expect, expect-bearing and ordered batch callers, one native staging clone, existing caller-specific validation order, zero/mismatch/refusal with no publication and positive counts with one commit even if later pairs undo bytes. Returned Python allocation and immutable validation precede commit. Active retirement remains checked.
- Dormant save/to_bytes accounts for native normalization, PyBytes allocation and stream-write failure without inventing a revision. The general serializer rule is sound as a requirement, but the additional direct image replacement boundaries need the same explicit disposition under F9.

All thirteen A283 rows remain, including correctness-only width/border/margin coverage and the separate cell-paragraph setter observation. Case1 remains bodies0/500/2000, a20by5table, the100reused handles, text then alignment then shading, no warmup, each2000-body cost below1ms and at most1.5times empty plus0.05ms. Case2 remains4000three-run paragraphs and1000matches, inside-process Python timing before open, CLI subprocess time, three paired trials, lexicographic Python tuple selection and minimum CLI command time. Matching Linux x86_64 release wheel/CLI remain deciding. No threshold, timed boundary or denominator was relaxed.

Historical source references and missing-link artifacts remain labelled as historical in the operative closure notes. Actual implementation Base revalidation remains required after the F282 resource predecessor. There is no F272 provider dependency and no new F282 API edge. The original design-peer reports and counts remain preserved. Old preparation availability statements are superseded by the final targeted-repair section, not new peer capacity.

## Not found

No independent new panic, OOXML, numeric, test-design, structural, public-signature or performance-contract defect was established. No separate smell or nitpick was found. Full code correctness and performance cannot be accepted from a design proposal.

## Pending evidence and stop

All implementation RED/GREEN, current missing-link reproduction, schema/native controls, exact release performance, installed Python3.9/3.12, typing/stubs, F26879/252/407 and other retained union riders, hash49, WASM, wheels/archives/publication dry-run and hosted CI remain pending. Word0.17.0 and shared/PowerPoint0.15.0 retain separate final publication approval at the reviewed SHA.

No plan/source/progress edits, test, build, benchmark, network, Git mutation or cleanup ran in this audit. This report is the sole write. Review ended and the canonical review-file writer lease is returned to Root. STOP.
