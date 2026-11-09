# F-X267, Bound PowerPoint mutation cost across public publishers

**Status**: completed
**Sprint**: S95.4
**Size**: L
**Depends on**: F-X260, F-X262

## Execution boundary, 2026-10-27

The API, preservation, refusal and acceptance contracts below remain binding. Problem descriptions, approval-time unrun statements and local provider-prefix checkpoints retain their historical source and scope. They do not describe the final integrated execution state. At this checkpoint, the native union and affected corrective cases, installed Python suites and typing, native Word and PowerPoint qualification with explicit source applicability, A7 measured memory and retained-history review, A9, the original Issue407 gate, hash49, lint, WASM, rustdoc, supply chain and policy observations are available. Actual producer identities remain attached to each observation. No old binary is relabelled as a new build.

The remaining integrated acceptance debt at this checkpoint is Cell.add_picture's required flat-cost gate, the active README example batch, the clean22-package publication dry run and the final full-union record. The original failed attempts remain evidence. Status stays approved until the required final acceptance passes. When this plan becomes completed, its final source-bound delivery entry in docs/sprints/AS_BUILT.md supersedes these historical execution-status statements. Completion changes neither the contract nor any threshold. Build-only hosted wheel qualification belongs to release preparation before final release approval. The subsequent F-X270 intake prepares fresh0.16.4/0.14.4 patches within this sprint because the old0.16.3/0.14.3 candidates were already published. It is not an additional acceptance criterion in these eight sprint designs. Publication remains separately approved at the exact reviewed main SHA.

## Problem

Public PowerPoint mutations still rebuild more of the package than the edited owner. Issue 403 asks for every mutating call in the external catalogue. PR413 `bdf22890de1762d025b89b83b549c56941ef7199` sha256 `03eeced92fef9ef9ecfde8ebd01761f04f6d21da2280ea1175fedce98ef56f28` leaves 107 package-bound calls. PR416 `768f8c109bd43fb9c1f6ecdfe2d913382bcc0368` sha256 `3497dd303faf1f7619c5a5ac096647bf2cadbbc892e8e056f0e111d938eb1592` is incremental on that head only. Incremental diff sha256 `690ca832aa7a26f69102c5f7fed282301ea0164ef52d452646a3bda8c670b80f`. Base is `a57be365`. Patches were not executed. Donor timings are not evidence.

Current storage is `parts: HashMap<String, Vec<u8>>` at `crates/oxml-opc/src/package.rs:75`. `set_part` takes `Vec<u8>` at line 350. `remove_part` returns `Option<Vec<u8>>` at line 358. `get_part` stays `Option<&[u8]>` at line 334.

## Spec reference

- `docs/hld/10-bindings-spec.md`, Packaging, Explicit PowerPoint text fitting, and PowerPoint footer, field, theme and transition authoring.
- `docs/hld/04-opc-and-packaging.md`, Checked PowerPoint XML transactions.
- `docs/hld/03-architecture.md`, PowerPoint checked XML ownership.
- `docs/hld/12-testing-strategy.md`, PowerPoint checked XML qualification, and PowerPoint footer and transition acceptance.
- `docs/hld/06-presentationml-model.md`, Public facade and Preservation strategy. New ownership prose is HLD impact under those existing sections.
- `docs/hld/14-development-backlog.md`, F-X267 story intent. HLD 15 is not in this list. No asset or version mechanism change.

## Approach

Planning estimate is 6 to 10 developer-days, provisional and not a measurement. No new file or trait. Python signatures stay. Q9 allows one bounded Rust break and no other public break. The retained release-version resolution below applies.

### OPC storage (Q9)

Q9 allows the storage and removal break. The parameter is generic. Two instantiations are implemented with this feature. They do not have to already compile on the old `Vec<u8>`-only base. Do not wrap every `Vec` caller only to avoid the generic. Do not keep a concrete `Arc`-only parameter.

```rust
pub parts: HashMap<String, Arc<Vec<u8>>>,
pub fn set_part(&mut self, part_name: &str, data: impl Into<Arc<Vec<u8>>>),
pub fn remove_part(&mut self, part_name: &str) -> Option<Arc<Vec<u8>>>,
```

Instantiation one is the existing `Vec<u8>` caller: `package.set_part(name, mutate(xml).into_bytes())`. Instantiation two is `Arc::clone` of shared media at PR413 diff lines 3230, 3254 and 3268 in `rpptx` embed-media call sites. Both are structural and present in the implemented feature.

`get_part` stays `Option<&[u8]>`.

### Retained release-version resolution

Q9's deferred release-version decision is resolved by the retained-version policy already approved in F-X263 and Root's selection on2026-10-27 under the user's standing instruction to choose recommended options. Keep prepared Word0.16.3 and shared/PowerPoint0.14.3. F-X263's completed plan explicitly superseded the earlier0.17.0/0.15.0 proposals. Disclose the approved Arc storage and removal source break prominently in both family release notes, including the Word OPC re-export and concrete migration examples. This is the existing retained-line exception, not a claim that the source break is semver patch-compatible. No version carrier changes are required. Historical Q9 answer provenance remains intact. Publication stays subject to separate final approval at the reviewed main SHA. The current S95.4 issue/PR contribution inventory is authenticated from preserved intake receipts under /private/tmp/s954-release-notes-assessment/contributions.json. No fresh scope was admitted. Release notes retain earlier unpublished work and add these eight stories, native limitations and source migration. Final acceptance and delivery records remain pending.

### Reread and compare

```rust
#[doc(hidden)]
pub fn reread(mut self, read: &OpcPackage) -> Result<Self>;
#[doc(hidden)]
pub fn same_package_as(&self, other: &OpcPackage) -> bool;
```

`reread` is the hidden public cross-crate provider. The standalone public helper documents this precondition: the baseline is still unmodified and read-equivalent. A package that was read and later mutated is not a baseline. The internal `Parsed.package` is the immutable baseline, and the candidate is edited from it. The trusted flag records that provenance. It does not detect an arbitrary public map edit. The known hand-built fallback is the untrusted flag, not a map scan. No snapshot uses the full ZIP fallback. Unsafe equivalence also uses that ZIP fallback, including normalization, a malformed producer, oversize input, and any other unproved fast-equivalence case. `same_package_as` compares packages that are a normalized read or a reread. It is not an unconditional comparison of an arbitrary hand-built package. Within the proved precondition, write then read matches, including errors, producer bytes, normalized names, backslash, duplicate case, relationship collision, content-type change, and malformed raw bytes.

Debug ZIP attribution, added after pass 1, is one hidden note and one crate-local helper. No new file, module or trait. `oxml-opc` does not depend on `rpptx`.

```rust
#[doc(hidden)]
#[cfg(debug_assertions)]
pub fn set_reread_zip_note(note: Option<fn()>);
```

Precondition: the note is installed by `rpptx` before `OpcPackage::reread`. It runs once immediately before the fallback `write_to`, including when that write returns an error. It does not run on the fast path and it does not run for any other `write_to` caller. `take_commit_work`, `saved_package_read_back` and `reopen` install `count_reread_zip_write`, which increments `zip_writes` through `note_commit_work`. `without_commit_work` pauses that increment, so the debug byte comparison does not count. `Presentation::note_embedded_zip_write` increments the same field once immediately before `commit_embedded_candidate`'s `write_to`, including when that write returns an error, and it does not count the following read. `to_bytes` and the parsed-`None` reopen and read-back paths keep their existing single increment and do not pass through the new notes.

```rust
#[doc(hidden)]
pub fn saves_as(&self, other: &Presentation) -> Result<bool>;
pub fn unchanged_slides_since(&self, earlier: &Presentation) -> Vec<Option<usize>>;
pub fn unchanged_layouts_and_masters_since(&self, earlier: &Presentation) -> bool;
```

`saves_as` is exact saved-byte equality, including pending direct edits, and it keeps the existing error semantics. `unchanged_slides_since` returns a conservative `Arc` allocation-origin map. `unchanged_layouts_and_masters_since` uses the same allocation and order. Both are doc-hidden. A newly allocated record with identical bytes may be `None`. These queries are not saved-byte equality. Do not describe every query as byte equality.

Debug only:

```rust
#[doc(hidden)]
#[cfg(debug_assertions)]
pub fn take_commit_work() -> [usize; 4];
```

Python `_inventory_work` and `_commit_work` stay debug test extensions. They are not a stable export.

### Behaviour commits (Q13)

Three separate declared reviewed behaviour commits, each with targeted evidence.

- Untouched slide and notes producer bytes stay preserved. Word fixture 49 stays unchanged. An intentional M21 digest delta for two handwritten notes is measured in its own reviewed behaviour commit. The oracle is this repository's golden, hash harness and regression. No PowerPoint application oracle is added for the digest. Genuine existing native gates stay.
- `fit_text` on a slide ignores an unrelated broken slide. Layout, master and notes still use every relevant context. Global `fit_report` and refresh stay unchanged.
- Transition and footer owner proofs cover only affected parts. Real structural ambiguity still refuses.

Unchanged owner incarnations stay preserved.

### Publishers and commits

Pointer identity is not enough. Source bytes, relationships, notes, comments, media, field ids and signature facts are fresh after a real edit. Tests go read bytes, then direct mutate, then package edit, then held, fresh, native and reopen, plus a fresh in-memory equivalent. A shared layout with two slide contexts is required. No-op detection compares output bytes, not pointers.

Distinct operations: `commit_candidate`, `commit_owned_parts`, and `commit_embedded_candidate`. Owner-proof preparation finishes every fallible Python result and revision before publish. Maps use allocation identity.

`take_commit_work` is `[usize; 4]`: parse, serialize, hash, ZIP. It is thread-local and resets on read. A separate owner-payload counter is `_inventory_work`. Do not treat the four-slot commit counter as that payload count. A local reopened target edit compares N with 4N. Payload parse, serialization, hash and changed-owner payload work stay equal. ZIP is 0 on the fast path. Global metadata, removal ordinals and reference remap are an explicit bounded linear exception. Signed and embedded fallback timings stay disclosed under Q14.

PR413's historical 107 package calls are contribution accounting only. They are not a product test and not an intermediate bench. The PR416 increment is consumed once. The final gate is the same 248 calls with every mechanism verified together. The removed test name `pr413_residual_stays_107_until_pr416` is not a gate.

Signed and embedded packages keep complete correctness and signature guarantees. Costly safety fallbacks are measured and disclosed. There is no constant-time promise.

### Performance gate (Q14)

External pin `3374348a48164f198ef077af46df485d0a1c469d`. Thresholds in the pinned `test_pptx_scaling.py` lines 26 to 30: N is 5, CALLS is 3, RATIO is 2, FLOOR_MS is 5, REOPEN is 16. The assertion shape is `large < 2 * small + 5` for decks of 5 and 20 slides. CALLS is 3 fresh actual targets, fastest sample. Reopen budget is 16. Large-deck early stop is 0.05 s.

Original growing shapes stay in memory: `shape_calls.py` sizes 10, 40 and 80, with and without media and notes, and `package_calls.py`. Raw outputs stay. Do not relabel them.

Added limit: each 40 and each 80 reported timed edit satisfies `unit_ms < 2 * 10_baseline_ms + 5`. Workload setup and aggregation match the originals. Shape timing is the mean of 10. Package timing is the mean of 5. Reported batch timing converts seconds to milliseconds and does not relabel a per-edit figure. The 40-slide large-image 600 case needs an additional 10-slide baseline with the same 600 payload. The original 40 case is still printed unchanged. No compiling during the release bench.

Whole-deck operations are the original 8. They are excluded from timing only by that meaning. The catalogue is the installed stub union of TIMED and WHOLE_DECK. The 89 names that are already constant stay constant. Original strict xfail and XPASS outcomes stay separate from `--runxfail`.

Static key count, without executing contributor code: 212 literal keys plus 36 dynamic loop keys is 248 unique keys. Core's 15 keys sit inside those 36. The installed stub completeness union keeps context variants. Do not collapse them for timings. The literal 212 figure is not a missing-key defect.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Shared bytes | `set_part(&mut self, part_name: &str, data: impl Into<Arc<Vec<u8>>>)`, `remove_part`, `get_part` | Two instantiations in this feature: `into_bytes` and `Arc::clone`. Python unchanged. Consumers include `oxml-chart`, `oxml-sml`, signatures, Word comments and Flat OPC redaction, both native families, CLI and WASM. Source compatibility and output preservation. No per-call fake test | Read, serialize |
| Trusted reread | `OpcPackage::reread`, `same_package_as` | Baseline stays unmodified and read-equivalent. `Parsed.package` is immutable. The candidate is edited from it. No snapshot and unsafe equivalence use the ZIP fallback | Write, read |
| Slide delta | `unchanged_slides_since`, layout and master query, `saves_as` | Allocation provenance for the unchanged queries. `saves_as` is saved-byte equality, including pending edits | Commit, reopen |
| Fit text | Slide `fit_text` versus global `fit_report` | Unrelated broken slide is ignored only by the slide target | Update |
| Footer and transition | Affected-part owner proof | Structural ambiguity still refuses | Update, serialize |
| Cost | External 248 calls plus 10, 40, 80 shapes | Means and the 40 and 80 budget | Timed public calls |

Publisher map that must be rechecked before a fast path: `slide_mut`, `shapes_mut_at`, `core_properties_mut`, notes and handout `HeaderFooter_mut`, and `TextParagraphMut.default_run_properties_mut`. Each one changes facts that a `Parsed` `Arc` baseline must not keep. Direct mutation, package edit, held, fresh, native and reopen are the observations.

## Behavior contracts and evidence

| Surface or placement | Exact behavior and boundary examples | Primary source identity and location | Local discriminator or pending evidence |
|---|---|---|---|
| Arc storage | `remove_part` returns the same allocation a later `set_part` can share | `package.rs:75`, `350`, `358`. Q9 | `remove_part_returns_shared_bytes` |
| Reread | Write then reread matches a ZIP read, including a malformed raw part | PR416 incremental, unexecuted | `reread_matches_write_and_read` |
| Hand-built baseline | Parsed `None` and an untrusted hand-built package | One ZIP fallback. No equivalence promise | `hand_built_baseline_uses_fallback` |
| Untouched slide | Edit slide 2. Slide 1 producer bytes stay | Q13 commit | `untouched_slide_bytes_stay` |
| Fit text | Broken unrelated slide. Target slide still fits. Layout and notes still see their own context | Q13 commit | `fit_text_ignores_unrelated_broken_slide` |
| Footer | Edit one slide footer. Another slide's footer part is untouched. Ambiguous structure still refuses | Q13 commit | `footer_proof_is_part_local` |
| No-op | Set the same XML the part already has | Output bytes equal. Pointer equality is not the proof | `noop_compares_output_bytes` |
| 248 calls | 212 literal keys plus 36 dynamic keys. Core 15 sits inside the 36 | Pin lines, not a live run | `mutating_call_catalogue_is_248` pending execution |
| 40 and 80 | Same setup as the original scripts | `unit_ms < 2 * baseline_ms + 5` | Pending release bench |

## Dependency readiness

F-X260 and F-X262 are done. F-X267 is pending and unclaimed. F-X264's `zip_entries` contract is a neighbour, not a formal dependency. Implement this story only on a base that already contains any integrated OPC signature this plan assumes. Today the base still has `Vec<u8>`. This draft is not approval.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A1 | Issue 403 catalogue. Static 248 keys. 89 constant names stay constant. 8 whole-deck names excluded from timing by meaning | Catalogue matches the stub union. Strict xfail stays separate from `--runxfail` | External pin and installed stubs | `mutating_call_catalogue_is_248` | Execution pending |
| A2 | N 5 and 20, CALLS 3, reopen 16, early stop 0.05 s. Fresh actual targets, fastest sample | `large < 2 * small + 5` | Public mutating calls | `slide_ratio_matches_pinned_gate` | Pending release bench. No compilation |
| A3 | Original 10, 40, 80 in-memory shapes with and without media and notes. Original package script. Raw lines retained | Printed originals unchanged. 40 and 80 each satisfy `unit_ms < 2 * 10_baseline_ms + 5`. The 600-payload 40 case has a matching 10-slide baseline | `shape_calls.py` and `package_calls.py` workloads, reimplemented in existing tests | `shape_40_and_80_meet_added_budget` | Pending. Means stay 10 and 5 |
| A4 | Q13 untouched slide and notes bytes | Producer bytes preserved. M21 digest delta measured, not copied | Commit | `untouched_slide_bytes_stay` | Separate behaviour commit |
| A5 | Q13 `fit_text` | Unrelated broken slide ignored. A shared master owner with two consuming slide contexts that differ in fit uses `fit_text_at` to choose a size valid in both. A notes owner fits, then refuses when that notes-master theme relationship is removed. Global report unchanged | Fit publishers | `fit_text_ignores_unrelated_broken_slide`, `fit_text_at_selects_size_valid_in_distinct_shared_contexts` | Separate behaviour commit |
| A6 | Q13 transition and footer | Affected parts only. Structural ambiguity refuses. Unchanged owners stay | Those publishers | `footer_proof_is_part_local` | Separate behaviour commit |
| A7 | Direct mutate, package edit, held, fresh, native, reopen, and a second in-memory presentation built with `Presentation::new`. Shared layout, two slides | The common facts are subject, slide number start 5, the collection-marker text and the default-run bold payload. Notes and handout header edits stay on the reopened fixture. Whole-package equality is not required. No-op uses output bytes | The publisher list above | `parsed_baseline_tracks_real_edits` | Local regression |
| A8 | Signed and embedded complete packages | Correctness and signature checks hold. Fallback cost is disclosed. No constant-time claim | Signature readers | Existing `fx211b_native_opaque_root_edit_retains_invalidated_signature`. Embedded ZIP count is `embedded_replace_counts_commit_zip_write` | Local correctness. Timing is separate |
| A9 | `reread` and `same_package_as` on an unmodified read-equivalent baseline, plus parsed `None` and one untrusted hand-built package. Neighbor proof uses a valid unmodified `read_pptx` baseline and one candidate. Ordinary ZIP entry lengths are 65504 and 65505. Relationship entry lengths are the same pair, measured with `part_name_to_rels_path`. One refused ordinary name also carries an invalid XML payload | Proved unmodified baseline matches. No snapshot, normalization, malformed producer, oversize, and the hand-built package use the ZIP fallback. One function covers parsed `None` and the hand-built package. A trusted baseline with one unnormalized ordinary part counts a ZIP fallback in `unnormalized_opaque_part_counts_reread_zip_fallback`. Entry length 65504 agrees on success for the ordinary name and the relationship name. Entry length 65505 agrees on the shared writer name error. The invalid XML payload keeps that XML error ahead of the name error. A long otherwise valid name may use the ZIP fallback | OPC and `saves_as` | `reread_matches_write_and_read`, `hand_built_baseline_uses_fallback`, `reread_matches_shared_zip_name_limits` | Local regression |
| A10 | Historical PR413 residual of 107 | Contribution accounting only. Not a product test, not an intermediate bench, and not delivery credit. The final mechanisms are verified together on the 248 gate | Inventory note | None | Accounting only |
| A11 | Counters for parse, serialize, hash, ZIP and owners | Bounds follow N and 4N. Zero ZIP only where promised. Metadata work is classified | `take_commit_work` | `commit_counters_follow_real_work` | Debug only |

## Test matrix

| Case | Modes or state combined | Expected difference or agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 | Edit one slide. Read the other slide's bytes | Untouched bytes match the producer | A full ZIP rewrite labelled constant fails |
| C2 | Shared layout, two slides, edit one placeholder | Both contexts stay coherent. The other slide's unrelated part stays | Pointer identity as a no-op check fails |
| C3 | Broken other slide during `fit_text`, and a broken layout used by the target | Target ignores only the unrelated slide. Its layout still participates | A global skip fails |
| C4 | Ambiguous footer structure | Refusal, no partial part write | A best-effort footer write fails |
| C5 | 10-slide 600-payload baseline beside the 40-slide 600 case | Same payload class | Comparing 600 against a tiny image fails |
| C6 | Strict xfail run versus `--runxfail` | Outcomes stay in separate reports | Marking the strict run green fails |

Omitted: a Python signature change, a release version bump, and equivalence for a hand-built baseline.

## Rejected alternatives

- A concrete `Arc<Vec<u8>>` parameter that rejects `impl Into`.
- Skipping the ZIP fallback when no authenticated baseline exists, or when fast equivalence is unproved.
- Copying the donor M21 digest instead of measuring it.
- Folding the three behaviour changes into the performance commit.
- Treating the literal 212 as a missing-key defect.

## Test plan

Backlog test gate, verbatim: regression.

| Timing | Category | Test | Asserts |
|---|---|---|---|
| local scoped | regression | `remove_part_returns_shared_bytes` | A concrete Arc round-trip |
| local scoped | regression | `reread_matches_write_and_read` | A9 proved precondition |
| local scoped | regression | `reread_matches_shared_zip_name_limits` | A9 ordinary and relationship ZIP entry lengths 65504 and 65505, plus XML error precedence on a refused name |
| local scoped | regression | `hand_built_baseline_uses_fallback` | A9 parsed `None` and untrusted hand-built |
| local scoped | regression | `untouched_slide_bytes_stay` | A4 |
| local scoped | regression | `fit_text_ignores_unrelated_broken_slide` | A5 unrelated slide, notes owner, and broken notes-master theme |
| local scoped | regression | `fit_text_at_selects_size_valid_in_distinct_shared_contexts` | A5 shared master contexts with different fit results |
| local scoped | regression | `footer_proof_is_part_local` | A6 |
| local scoped | regression | `parsed_baseline_tracks_real_edits` | A7 |
| local scoped | regression | `fx211b_native_opaque_root_edit_retains_invalidated_signature` | A8. Existing test name. The forwarding alias was removed |
| local scoped | regression | `embedded_replace_counts_commit_zip_write` | A8 embedded ZIP write |
| local scoped | regression | `unnormalized_opaque_part_counts_reread_zip_fallback` | A9 trusted baseline with one unnormalized part |
| local scoped | regression | `commit_counters_follow_real_work` | A11 |
| final sprint | performance | Release build. N 5 and 20. Shapes 10, 40, 80. No compilation | A1, A2, A3 pending |
| final sprint | regression | M21 two-note digest measured against golden, harness and regression | A4 pending. No PowerPoint application oracle |
| final sprint | risk rider | `/verify --full`, hash, installed typing, stub, package dry-run, WASM | Union once |

Existing `crates/rpptx-py/tests/test_documented_examples.py` and the native integration entrypoint. No new test binary. Original `shape_calls.py` and `package_calls.py` stay external raw controls. Their workloads are reproduced by tests in the existing files.

## HLD impact

- `docs/hld/03-architecture.md`, PowerPoint checked XML ownership
- `docs/hld/04-opc-and-packaging.md`, Checked PowerPoint XML transactions
- `docs/hld/06-presentationml-model.md`, Public facade and Preservation strategy, plus new ownership prose under those sections
- `docs/hld/08-rendering-spec.md`, native-qualified timeline masks and SmartArt profile behavior
- `docs/hld/10-bindings-spec.md`, public API and bindings
- `docs/hld/12-testing-strategy.md`, tests
- `docs/hld/14-development-backlog.md`, F-X267 story intent
- HLD 15 is not an impact file. No asset or version mechanism changes.

## Risk routing

- Public API. The Arc storage and `remove_part` return are the only approved Rust break. Python unchanged. The retained release-version resolution below applies. Dry-run is final-only.
- Parser or serialiser. `reread` and `write_to` share entry rules. Child order stays.
- PyO3 and WASM. No new Python signature. WASM check is final-only. Debug helpers are not exports.
- Performance. A2 and A3 pin the workload, means, conversion from seconds to milliseconds, and the ban on compilation.
- Cached facts. The publisher list is the fact-changing set. Held, fresh, native and reopen are required. Pointer identity is rejected.
- Compatibility and ordering. Part names, backslash, duplicate case and content types are inside A9.
- Shared consumers. Slide, layout, master, notes, footer, transition, core properties, and text defaults.
- Contributed criteria. The 248 gate and the three behaviour commits stay separate. The historical 107 count is accounting only.
- Digest evidence is the repository golden, the hash harness and regression. No new PowerPoint application oracle. Genuine existing native gates remain.
- Release scripting. This story does not bump a version. The row is a prohibition, not a task.

## Hash harness

The untouched-slide commit may change the M21 deck digest. That delta is measured in its own reviewed behaviour commit. It is not copied from the donor and not folded into a performance or file-move commit. Other fixtures stay unchanged unless a named behaviour commit states the delta. No baseline is recorded in this approval. Rendering baselines stay on deterministic fonts.

## Implementation checklist

- [x] Change `parts` and `remove_part` to `Arc<Vec<u8>>`. Change `set_part` to `impl Into<Arc<Vec<u8>>>` with both instantiations.
- [x] Add `reread`, `same_package_as`, `saves_as`, and the unchanged-since queries.
- [x] Split the three Q13 behaviour commits. Untouched producer bytes, slide fit scope, and part-local footer proofs are separate labelled commits. Independent M21 measurement stays pending.
- [x] Prove publishers against output bytes in the local tests named above. The pass 2 repair checks a checked package commit, a real collection edit, the common in-memory facts, and distinct shared fit contexts.
- [x] Leave the release bench and the digest measurement pending. The M21 constant is the original `411f028c4b51c6aac0a0e2ae11bc63592a6d0947ab1a5f0d15116d19bfe766bc`. The donor digest was copied and has been removed. Measurement is final only.
- [x] Keep Python signatures and the release version unchanged.

## Open questions

| ID | Affected F-IDs and A-IDs | Decision and recommendation | Status and answer/source |
|---|---|---|---|
| Q9 | F-X267 A7 and the OPC storage row | Bounded Rust Arc storage. | Answered 2026-10-23 by user answer. Allow owned `Vec<u8>` to `Arc<Vec<u8>>` for OPC public Rust storage, removal, and that signature break. Python unchanged. Release version later. Not permission for any other public break. Root contract the same day: `set_part(&mut self, part_name: &str, data: impl Into<Arc<Vec<u8>>>)`. The two instantiations are the `Vec` `into_bytes` caller and the `Arc::clone` media caller. They are implemented with this feature. |
| Q13 | F-X267 A4, A5, A6 | Three PR416 behaviour changes. | Answered 2026-10-23 by user answer. Separate declared reviewed behaviour commits. Untouched slide and notes producer bytes stay, with an intentional M21 digest delta measured and not copied. Slide `fit_text` ignores an unrelated broken slide. Layout, master and notes still use relevant context. Global `fit_report` and refresh stay. Transition and footer proofs cover affected parts only. Structural ambiguity still refuses. Unchanged owner incarnations stay. |
| Q14 | F-X267 A1, A2, A3 | 248-call gate and the 40 and 80 budget. | Answered 2026-10-23 by user answer. `large_ms < 2 * small_ms + 5`. N is 5 and 20. CALLS is 3 fresh actual targets, fastest sample. REOPEN is 16. Large early stop is 0.05 s. Original 10, 40 and 80 in-memory `shape_calls.py`, with and without media and notes, and `package_calls.py`, stay with raw outputs. Added limit: each reported 40 and 80 timed edit is `unit_ms < 2 * 10_baseline_ms + 5`. Setup and aggregation match the originals. Means stay 10 for shapes and 5 for package calls. Seconds convert to milliseconds without relabelling a per-edit figure. The 40-slide 600 image needs a 10-slide 600 baseline. The original 40 line stays. No compilation during the release bench. Signed and embedded correctness stays. No constant-time promise. The original 8 whole-deck operations are excluded from timing by meaning. The 89 constant names stay constant. Strict xfail and XPASS stay separate from `--runxfail`. 212 literal keys plus 36 dynamic keys is 248 unique. Core 15 is inside the 36. Context variants stay uncollapsed. |

No unanswered user choice remains. Q16 is answered on F-290 only. Its provenance is the 2026-10-23 user choice "Author Word-compatible canonical slots (Recommended)".

## Contract closure

Codex root semantic validation on 2026-10-23 approved this contract at source `a57be365`, including the shared `Arc<Vec<u8>>` buffer contract. Every timing and digest claim is unexecuted. Implementation and final gates stay pending. This approval does not mark the story completed.

Operations: `set_part` and `remove_part` prove storage. `reread` and `same_package_as` prove the package comparison. `saves_as` and the unchanged-since queries prove presentation comparisons. `fit_text`, footer and transition publishers prove A5 and A6. `take_commit_work` proves counters and is debug-only. No new Python method is added.

`set_part` takes `impl Into<Arc<Vec<u8>>>`. The two instantiations are part of this feature. A concrete Arc-only parameter is rejected.

Trusted reread requires a baseline that is still unmodified and read-equivalent. `Parsed.package` stays the immutable baseline. The candidate is edited from it. Parsed `None` and an untrusted hand-built package share `hand_built_baseline_uses_fallback`. No snapshot and unsafe equivalence keep the ZIP fallback. A7 lists the publishers whose edits must refresh source bytes, relationships, notes, comments, media, field ids and signatures.

A1's 248 count is 212 literal keys plus 36 dynamic keys, with Core 15 inside the 36. It is not a claim that the suite was executed.

HLD impact is 03, 04, 06, 10, 12 and 14. HLD 15 is absent. No version carrier is changed. Release notes apply the retained-version resolution recorded below. The M21 digest is measured here. Word 49 is expected unchanged. A10 is accounting only.

## F-X268 local provider checkpoint

The next paragraph is the 2026-10-24 approval. It is dated historical for the state before the 2026-10-25 local completion recorded below.

Codex root approved this checkpoint on 2026-10-24 for use by F-X268. It does not mark F-X267 done. The story stays in progress and reviewed, owner grok, until the checkpoint is actually satisfied. Final Word, the full suite, the hash harness, the release 248-call run, and the new 65-call, 252-call, and memory riders that belong to F-X268 stay mandatory and unrun. `AS_BUILT.md` and `SPRINT_TRACKER.md` stay unchanged.

The checkpoint consumes all 13 existing local scoped rows: `remove_part_returns_shared_bytes`, `reread_matches_write_and_read`, `reread_matches_shared_zip_name_limits`, `hand_built_baseline_uses_fallback`, `untouched_slide_bytes_stay`, `fit_text_ignores_unrelated_broken_slide`, `fit_text_at_selects_size_valid_in_distinct_shared_contexts`, `footer_proof_is_part_local`, `parsed_baseline_tracks_real_edits`, `fx211b_native_opaque_root_edit_retains_invalidated_signature`, `embedded_replace_counts_commit_zip_write`, `unnormalized_opaque_part_counts_reread_zip_fallback`, and `commit_counters_follow_real_work`. It also consumes genuine RED, GREEN, and negative proof, the scoped checks, clean Root microscopes, and the integrated source ancestor. Map each name to the command and outcome already in the progress record. Do not invent a combined pass. Rerun only an input a later change invalidates. Rebuild the current import where binding evidence is affected.

Source head `70573c91030be7073d29e3451383060eb207207a` is integrated at `d0f17a6ac7ef675d0ed2a86cea641ce21d65bfdd`. The earlier `.so` artifact `1fe91` is earlier fast-path proof. It is not a fresh D2 or D3 qualification of this source head. That limitation stays. This section does not weaken the final performance, hash, or Word riders above.

## Local prefix completed 2026-10-25

Root approved genuine local provider completion for F-X267 on canonical `79047bf193bcd9dfb229a0b37a93d979dc0997a7`. This record is that Root review. It is not a self-review. The plan status stays approved. The story is done for this local prefix only. It is not final acceptance. All 13 scoped rows are covered. The post-counter retained rows were not rerun for this record.

The preferred retained proof for the rows logged after the counter change is `.claude/scratch/F-X267-progress.md` at the post-D2 block. Status 0 there records `reread_matches_write_and_read` as 1 passed and 70 filtered. It records `commit_counters_follow_real_work`, the unnormalized test, `untouched_slide_bytes_stay`, `parsed_baseline_tracks_real_edits`, `hand_built_baseline_uses_fallback`, `footer_proof_is_part_local`, and `fit_text_ignores_unrelated_broken_slide` as each 1 passed and 201 filtered. It records `embedded_replace_counts_commit_zip_write` as 1 passed and 428 filtered. That block is the latest applicable proof for those names. A looser pass-2 description is not substituted for it.

Fresh canonical runs on that HEAD fill the 7 rows that lacked a dedicated GREEN log: `remove_part_returns_shared_bytes`, `hand_built_baseline_uses_fallback`, `footer_proof_is_part_local`, `unnormalized_opaque_part_counts_reread_zip_fallback`, `commit_counters_follow_real_work`, `embedded_replace_counts_commit_zip_write`, and `fx211b_native_opaque_root_edit_retains_invalidated_signature`. Those runs are not a claim that a later change invalidated the rows. `cargo test -p oxml-opc --lib -- remove_part_returns_shared_bytes` passed 1 on `oxml_opc-7aaac1354a37b776`. `RUST_MIN_STACK=33554432 cargo test -p rpptx --lib --features default-template,render` passed 4 for `hand_built_baseline_uses_fallback`, `footer_proof_is_part_local`, `unnormalized_opaque_part_counts_reread_zip_fallback`, and `commit_counters_follow_real_work` on `rpptx-95549ba5a963a2a6`. `embedded_replace_counts_commit_zip_write` passed 1 on `integration-0fd6b67621096d8e`. `fx211b_native_opaque_root_edit_retains_invalidated_signature` passed 1 with `--features digital-signatures` on `integration-c5f9cf8fc0b03efd`. The earlier filter without that feature selected 0 tests. That 0-test run is not a pass.

`fit_text_at_selects_size_valid_in_distinct_shared_contexts` and `reread_matches_shared_zip_name_limits` stay consumed because `package.rs`, `rpptx/src/lib.rs`, and `rpptx/tests/integration.rs` have no commit after `d0f17a6ac7ef675d0ed2a86cea641ce21d65bfdd`. The current artifact `/private/tmp/s95.4-fx267-local-debug-_rpptx.abi3.so` sha256 `9fab4e917aef5ee74457ab95e85ffddaa262698fc7d9cbeaa1e88cc9cd0e56ca`, 81514792 bytes, is a fresh development build and import proof only. It is not the native 13 and it is not a release build. The old archive `1fe91d2dac0ae8fea4fed7bb97978b466d143e1f0bd0958c77a46c0700658c57` is not the current qualifier.

Source head `70573c91030be7073d29e3451383060eb207207a` and integration `d0f17a6ac7ef675d0ed2a86cea641ce21d65bfdd` are ancestors of `79047bf193bcd9dfb229a0b37a93d979dc0997a7`. Clean Root reviews remain the reviews already consumed at integration. Final Word, the full suite, the hash harness, the release 248-call run, the original 10, 40, and 80 timings, the M21 digest, the 252-call catalogue, the 79 flat gates, the 5 linear calls, memory, strict mypy, stubtest, WASM, and packaging stay pending. `AS_BUILT.md` and `SPRINT_TRACKER.md` stay unchanged.


## Q26, missing native PowerPoint reference requalification

The user confirmed on 2026-10-27 that the original reference bundle is unavailable and answered "Requalify on installed PowerPoint (Recommended)". The exact decision is recorded at `/private/tmp/s954-q26-powerpoint-requalification-approved.json`. This authorizes fresh independent native evidence for the four existing missing-reference checks. It does not add a PowerPoint comparison for the M21 notes digest, revoke the consumed local provider checkpoint, complete any final rider or guarantee that the new observations pass.

The missing references pin PowerPoint 16.104, bundle 16.104.25121423, application build 1214. The installed application is 16.114.1, bundle 16.114.26102031. Record the observed version, bundle and About identity before admitting a new manifest. The Q26 profile records About build26102031 and explicitly leaves the Application.Build getter unqueried. Do not infer a short getter value from those observations. Historical1214 remains only on the legacy profile. Keep the historical pins, fixed frame hashes and missing-input failures. Do not overwrite historical artifacts or globally retarget unrelated recorded tests. Explicit new trusted capture pins and provenance require independent review. Runtime self-authentication by an unreviewed manifest is insufficient.

Use the unchanged deterministic SmartArt source generator. The timeline source receives only the authored-owner correction specified below before fresh capture. SmartArt has exactly six families: list, hierarchy, cycle, relationship, matrix and pyramid. Native PDF export uses local Best for printing. Separate native-saved PPTX copies provide independent generated geometry. Rust diagnostic images and bounds are comparison subjects, never native expected values. All source bytes, native copies, PDFs, raster inputs, fonts, application identities and transformation commands are recorded and authenticated. Preserve original source decks. Use the pinned raster tools and the existing bounded comparison transforms. Native timeline evidence must come from an actual PowerPoint movie and exact encoded-time extraction with zero time tolerances. Preserve all nine cases and the terminal outgoing fade binding. No library-rendered replacement baseline, relabelled time, omitted case or threshold relaxation is authorized.

| Existing check | Required requalification | Unchanged acceptance |
|---|---|---|
| `supported_smartart_corpus_matches_pinned_powerpoint_geometry_and_ssim` | Six independently native-rendered families, original source identity, native geometry, ordered text ownership and diagnostics | Geometry at most 1 pt, symmetric text-masked non-text SSIM at least 0.90, exact ordered line counts, text edge and line-width errors at most 3 pt |
| `pinned_timestamps_match_powerpoint_frame_oracle_within_declared_tolerances` | Nine source-bound native movie samples, exact requested and returned encoded time, trusted movie/source/application pins and terminal outgoing fade | 150 DPI comparison, geometry at most 1 pt and SSIM at least 0.99 for every case |
| `powerpoint_timeline_oracle_foreground_cutoff_ignores_fixed_opacity_halo` | New independently captured appear/opacity pins, with old fixed pins retained as history | Existing foreground cutoff 240 keeps bounds equal and legacy cutoff 250 exposes the opacity halo |
| `fade_composition_matches_the_fixed_powerpoint_observation` | New trusted native fade frame, original source and outgoing-state binding | Existing 150 DPI normalized comparison and SSIM at least 0.99 |

Implementation uses existing `crates/rpptx/tests/integration.rs` and the named HLD impact files, particularly `docs/hld/12-testing-strategy.md`. No new tracked file, module, public API, dependency or product behavior is authorized by this reference refresh. Actual new product divergences remain failures and require separately assessed remediation. Label reference and pin changes in a distinct commit with measured expected test evidence and hash-harness outcome. Do not merge them into the Word result repair or performance change.

Risk routing adds external-oracle provenance, deterministic rendering and test compatibility. Before final verification, the source-to-application-to-output chain, reference independence, exact case coverage, refusal controls for substituted sources/pins/times, actual semantic outcomes and final relevant source applicability require review. Successful file creation is not acceptance. Native capture, new trusted pin adoption, four actual checks, affected metadata checks and independent microscope remain pending.


### Q26 authored timeline owner correction

Root accepted the read-only bounded design at `/private/tmp/s954-q26-timeline-target-binding-design.md`. The initial source generator selects children[0] and children[1], but the actual serialized slides have Title 1, Subtitle 2, then the authored Hero and click marker. Consequently its seven evaluator targets select empty Title 1 and its click target selects empty Subtitle 2. Source SHA `1f1a6da66e646ca4b2312fcc0f239e1dce023fdb54554c63b67e843a98bb5d0f`, initial native movie and all nine initial comparisons remain preserved as semantically inadequate evidence, including seven geometry failures. They are not accepted or relabelled as observations of visible Hero animation.

Q26's purpose requires meaningful independent native evidence. Correcting this fixture binding is within the approved reference requalification, without a new product API or feature. It supersedes Root's earlier unchanged-timeline-generator premise, not the user's unchanged thresholds. Select the unique exact authored name in the actual serialized CT_Slide, then use that same ordinary non-placeholder rectangle's unique nonzero ID. Refuse missing or duplicate names/IDs, wrong kind, wrong paint, absent or empty extent and off-page geometry. No positional fallback or hardcoded IDs 4/5. Retain placeholders and every other source property, all six slides, nine logical cases, exact local and encoded times and outgoing binding.

Tests bind evaluation by the actual selected Hero ID and click marker ID, not the first state-map value. They distinguish missing/duplicate names, missing/duplicate IDs, placeholder and wrong paint/extent, and reordered template children. A source retargeted to a placeholder must fail admission even if whole-image SSIM is high. Existing source binding and terminal outgoing fade controls stay. Visible paint and native pose observations must demonstrate that wipe, opacity, transform and exit concern the Hero. Merely unequal compressed frame bytes do not prove animation. The yellow marker remains in the terminal outgoing state, making the fade comparison nontrivial. No native active-click observation or zoom sample is invented.

This intentionally changes only the generated fixture source and its future reference pins. Hash harness 49 is expected unchanged. Preserve the old generator/source hash and all old pins as history. A changed source requires a fresh native movie and nine exact-time extractions. Independently review fixture correction before capture and actual provenance/outcomes before trusted adoption. If the corrected native export cannot satisfy the existing case/time/owner preconditions, retain the failure and escalate before changing samples or rendering semantics. All final outcomes remain pending.

## Root-assessed native remediation after Q26

Root approves the following bounded repair on2026-10-27 under the user's instruction to select recommended options without further questions. This explicitly extends the earlier reference-only boundary. It does not weaken the four native gates, replace independent references with renderer output or revoke the local performance-provider checkpoint. The actual corrected-target330ms failure, shadowed native comparisons and all strict extraction failures remain preserved.

Independent native controls establish the supported numeric property names `ScaleX`, `ScaleY` and `style.rotation`, with fractional opacity. Native directions reveal from the named edge and exits remove from that edge. Native exit fade overwrites a prior opacity hold, so retain that existing behavior. Keep all existing aliases and source-preserving timing serialization. No new timing node variant or public timing-model struct field is required.

The plain-paint width/duration corpus qualifies `alpha_in = clamp(8*p - 7*u,0,1)`, where u is normalized distance from the named edge and p is supported effect progress. Exit uses its complement. The two-width and fourfold-duration controls distinguish a spatial linear band from temporal blur. `/private/tmp/s954-q26-wipe-factor-analysis/REPORT.md` and its manifest retain exact native source, movie, actual timestamp, physical paint and measured fits. The first-slide phase controls require no fitted clock shift. Later-slide declared slots remain unqualified. No timing offset, easing correction or pixel-based band constant is approved.

The canonical animation fixture must use explicit plain solid rectangles with no line or inherited effects. Its appearance state uses the existing native style.visibility set operation, rather than claiming that the historical appear filter is a native-supported filter. The legacy Rust alias remains supported. This isolates the intended animation from unrelated theme shadows. It retains all six slides, nine logical cases, local and encoded sample times, click-marker binding and thresholds. The fixture change and expected new native source identity are separately labelled and independently reviewed before new capture. All shadowed sources, native files and failed comparisons stay available. Those observations exposed a distinct shadow falloff/effect-domain qualification gap, now pending under F-X269. Do not claim that the new mask qualifies those old shadowed samples or fixes general shadow fidelity. Native trust requires fresh exact-source capture and review, not a transformed old frame.

The canonical wipe has fill=remove at node102 so its completed entrance no longer participates in later opacity and exit observations. This is an explicit isolated-lifetime fixture contract, not a product fix for held filter composition. A one-variable independently captured native experiment changed only that literal: source35bbcdb88f1d469fd0f5d7dc4e3df793316487928e78c4892ac8c7dea9a362e5, movie006367184cf06caced6cea58207806013b9221591db51817433307434d9a3a82. All nine original exact samples decoded, and appear/wipe pixels stayed byte-identical while later opacity became visible. The original held source53d7fdc382b3bd0650a7279626800b1d8f64205083c7772dfd1d4befffabca9c and all failures remain preserved. Native held-wipe interaction with later opacity is not qualified by the current evaluator and joins F-X269's effect-composition design debt. No time, threshold, sequence, root duration or other timing property changes. The independent source assessment is /private/tmp/s954-remove-wipe-native-assessment.json. All nine native comparison metrics now meet the existing limits in /private/tmp/s954-q26-debug-native-remediation-v2/comparison.json, but these remain diagnostic until trusted source/pin review and actual gates. The separate fixed halo assertion currently fails at cutoff240 and remains pending without relaxation. Root selects this recommended bounded fixture contract under the user's standing no-questions instruction on2026-10-27.

### Q26 source-specific native halo acceptance amendment

Root selects the independent recommended contract on2026-10-27 under the user's instruction to select recommended options without questions. This is a semantic acceptance amendment, not a claim that the historical exact-equality predicate passed. The user required preserving geometry and SSIM thresholds, which remain1pt and0.99 for the timeline product gate. The native16.104 fixed observation retains its exact240 equality assertion and original pins. The new16.114.1 plain-source35bbc profile instead authenticates source, application observation, movie, all exact samples and full manifest, pins both observed240 bounding boxes exactly, and requires their maximum edge difference in source points to stay within the existing1pt bound. Its250 control must both exceed1pt and exceed the240 edge difference. Keep cutoff240, cutoff250, original sample times, source placement/color and all product geometry/SSIM thresholds unchanged. Current actual240 boxes are143,215,720,648 and144,216,720,648, a0.5pt maximum difference. Actual250 boxes are143,215,720,648 and138,213,725,650, a2.5pt difference. These are new source-specific native observations, not proof of a codec-only or version-only cause. Source, plain paint and effect lifetime also changed. Preserve the old equality failure in /private/tmp/s954-remove-wipe-native-assessment.json. The stricter250 control rejects a useless cutoff without claiming identical native antialias pixels. The independent assessment and exact alternative are supplied by /root/review_shared_proof_batch. Final qualification requires independent review and actual execution of all four gates against the versioned trusted profile. New code must fail on a changed native frame, unknown manifest/source, exceeded1pt bound or absent250 sensitivity.

### Concrete output and API contract

Add one variant to the existing non_exhaustive neutral Effect enum, without changing GroupElement fields:

```rust
Effect::LinearOpacityMask {
    bounds: Rect,
    start: Point,
    end: Point,
    stops: Vec<GradientStop>,
}
```

A doc-hidden inherent query `Effect::linear_opacity_mask_is_valid(&self) -> bool` owns the same geometry and stop validation for raster and SVG. Other effects return true without asserting their validity. Root approved this additive query during implementation on2026-10-27. It introduces no new type or forwarding abstraction.

Coordinates are child-local. Outside bounds alpha is zero. Inside, interpolate only stop color alpha, with pad extension beyond the endpoints. RGB does not contribute. Valid bounds are finite and positive. Endpoints are finite and distinct. At least two stops have finite ordered offsets in0through1 and finite alpha in0through1. Invalid masks make the group's raster/SVG contribution transparent rather than revealing children or panicking. Multiple masks multiply. Masks act on the complete group including its original effects, then group opacity applies once. Keep one child scene and one physical text/font/source owner, including overlapping descendants. No clipped-scene replication, raster replacement of semantic text or new generic framework is allowed.

Raster and SVG support this linear capability. PDF group effects remain unsupported under HLD08's existing contract. The new effect is deliberately ignored by the ordinary PDF writer, with child text emitted once and ordinary clips/transforms/opacity preserved. Checked conformance retains explicit unsupported-effect refusal. Do not promise a soft-wipe PDF or unchanged old hard-clip pixels. The public timeline result exposes a page, so document this limitation and disclose it through the existing frame diagnostic vector for an authored partial soft wipe. No PDF transparency-form, tagging or new PDF capability is in this repair.

Native-qualified timeline soft wipes target direct leaf owners. Group-target timing retains its existing flattened hard-clip fallback and an explicit diagnostic, without a new grouped timeline scene model. The neutral mask independently supports one existing group scene with nested/overlapping descendants. Root approved this consumer boundary on2026-10-27.

EvaluatedShapeState may retain a private source-derived wipe record. Its existing private fields already prevent external struct literals. Keep public clip authority: a caller-modified or cleared clip must not silently keep the previous automatic wipe mask. An additive `local_opacity_mask(size) -> Option<Effect>` projection may accompany the existing local_clip_paths method. For an unchanged evaluator-authored wipe clip, the renderer uses the qualified mask rather than the old hard front. Other explicit and oriented clips remain effective. No existing signature or Python method changes.

### Materialized timeline geometry consistency

Root approved the bounded A13 correction on2026-10-27 after the actual native1980ms comparison exposed correct numeric scale with stale materialized paths. At candidate publication, apply actual signed axis factors to every Custom path MoveTo/LineTo point, cubic control and endpoint, and Custom text_rect x/y/width/height before group clip projection. Keep center-resized bounds, rotation, motion and parent composition. Rectangle and BoundsFallback remain dynamically sized and receive no second geometry scale. Picture extents and custom clipping stay distinct. Factor1 is structurally equivalent. Zero and negative finite factors retain existing policy without abs or clamp.

Coordinate overflow or non-finite geometry declines the entire candidate through the existing diagnostic, leaving the original shape and default resolved leaf state. The qualified repair scales materialized local geometry. It does not scale font sizes, body insets, strokes, shadows or precomputed paint coordinates and does not promise general PowerPoint emphasis fidelity for those domains. No public field, API or whole-object affine replacement is introduced. The frozen canonical fixture generator and alpha-admission investigation remain separate.

### Consumers, acceptance and discriminators

| ID | Input and expected result | Consumers and local evidence | Final evidence |
|---|---|---|---|
| A12 | Unique authored Hero and click marker with explicit plain paint. Missing, duplicate, placeholder or wrong-paint targets refuse. All original cases and clocks remain | Existing generator and integration entrypoint, named owner guards and actual visible paint controls. Separate fixture commit | Fresh native movie, exact nine frames, trusted reviewed source/application/movie/frame identities and the unchanged four Q26 checks |
| A13 | An earlier applicable visibility operation survives a later pending entrance. Initial pending entrance remains hidden under the existing evaluator contract. Native numeric aliases and all four wipe directions follow their concrete values | Existing timeline tests cover both traversal orders, distinct targets, explicit hidden/visible sets, opacity/motion-only entries, click gating, parent endpoints, boundaries, uniform and independent scale, rotation, fractional and legacy numeric values. Preserve held-exit overwrite | New normal native/binding artifact and unchanged public runtime/performance riders |
| A14 | Linear alpha mask composes one complete scene, keeps transforms/clips/shadows/opacity and handles transparent/colored backgrounds and nested/multiple masks | Existing output, raster, SVG, PDF writer and conformance modules. Test exact endpoints, interpolation, invalid values, contrasting background, overlap, one text occurrence and explicit PDF limitations. Native mask tests distinguish spatial band and direction from hard clips | Qualified plain native wipe comparisons at existing1pt/0.99 limits. SVG/raster semantic control is separate. No PDF soft-mask parity claim |
| A15 | Six unchanged SmartArt sources use independent native frame/text owners and exact caller Calibri bytes | Existing diagram implementation and integration tests. Cycle uses native135.703pt square text owners and4.6pt insets. Pyramid uses actual separate txXfrm frames with192/249.6/374.4pt widths and4.4pt insets. List retains outer roundRect and preset text rectangle, with15.24/0/15.24/0 body insets. Apply native90percent line and35percent after defaults independently of vertical centering, preserving explicit data overrides and first/last policy | All six existing geometry, ordered-line, text-edge/width and nontext SSIM gates, unchanged thresholds, independent native ownership mapping and exact physical font provenance |

Use the native geometry and patch map in `/private/tmp/s954-smartart-profile-repair-map.md`. Preserve meaningful invisible text-frame geometry, painted pyramid shapes and connector topology. Whole-circle arrow transforms are not tight visible path bounds. Do not substitute Rust bounds for native expectations or a painted trapezoid for its actual text transform. Existing profile constants are updated from independently captured native/layout authority, not from a higher SSIM search.

### Source, HLD and risk work list

Existing source files only: oxml-layout/src/output.rs, oxml-pdf/src/raster.rs, writer.rs and conformance.rs for the explicit limitation tests, rdocx/src/svg.rs, rpptx-layout/src/timeline.rs and context.rs, rpptx-render/src/lib.rs and text.rs, rpptx/src/diagram.rs and existing integration modules. No new file, module, dependency, trait, generic parameter or feature flag. Use a coherent source batch, then one impacted build/test batch and independent microscope. Fixture, timeline/rendering behavior, SmartArt profile and eventual trusted pin changes have separate labelled commits with measured outcomes. Do not rebuild unaffected production solely for later test/pin metadata.

HLD03 describes the neutral output boundary, HLD06 the bounded animation semantics, HLD08 the linear mask/backend limits and qualified profiles, HLD10 the additive Rust surface and unchanged Python signatures, HLD12 the native and backend discriminators, and HLD14 the current repair/follow-up split. These extend the existing HLD impact list. HLD04's existing package contract stays. No global font policy change is introduced.

Risk routing adds additive public output API, shared raster/SVG consumers, PDF capability/conformance boundaries, layout and native compatibility. The final integrated union includes applicable rendering/backend and API tests, installed Rust/Python runtime, strict typing/stub checks, WASM, packaging and unchanged49-entry Word hash harness. Run final performance and memory checks on authenticated final artifacts. The mask adds no retained document edit history. Any changed neutral Effect storage or archive footprint must be measured and disclosed. No new RSS ceiling or constant-time promise is implied.

## Q-S96-NATIVE-01, PowerPoint 16.114.4 requalification

On 2026-11-07 the user selected "Requalify 16.114.4 (Recommended)". This authorizes fresh independent native evidence for only the four Q26 checks above using PowerPoint 16.114.4, bundle 16.114.26110224. Installed Info.plist identity is preparation evidence. The running application and About build26110224 must be observed before capture. Application.Build remains unqueried unless actually observed, and no short getter value is inferred.

Keep the historical16.104 and accepted16.114.1 profiles, manifests, movie/frame identities, halo observations and outcomes unchanged. The successor uses a separately reviewed reference profile and hardcoded trusted hashes after actual capture. It does not authenticate itself through adjacent manifest hashes or replace old constants globally. The new version is approved for requalification, not already qualified.

Use the existing six SmartArt source families and the current plain-rectangle, uniquely named authored timeline owners with wipe node102 fill=remove. Preserve all six timeline slides, nine logical cases, requested and decoded encoded times, terminal outgoing fade binding, source placement/paint, click binding, foreground cutoffs240/250, 150DPI comparison and bounded normalization. SmartArt retains geometry1pt, ordered text and exact line counts, edge/line-width3pt and symmetric text-masked non-text SSIM0.90. Every timeline case and fixed fade retain geometry1pt and SSIM0.99. Apply the existing source-specific halo contract to new observed boxes, never copy old boxes as observations of the new application. Retain any failure.

Native SmartArt PDFs use local Best for printing and independent native-saved PPTX geometry. Timeline requires an actual PowerPoint movie,1920x1080, with exact-time AVAssetImageGenerator extraction and zero before/after tolerance. Record original sources, native copies, movie, frames, PDFs, raster transformations, selected font files/faces, resources and tool hashes. Rust images/bounds are comparison subjects. Native capture uses the approved computer-use channel. Do not run the existing AppleScript-launching timeline generator.

Actual native capture, trusted successor pin adoption, substitution/refusal controls, all four checks and independent review remain pending. No product behavior, source fixture semantics, threshold or other native profile is changed by this amendment.
