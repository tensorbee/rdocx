# S96 admission frame local allocation, design investigation 12

## Disposition

The exact four-name static-borrow proposal is technically feasible and has a finite contract. Replace only WordAdmissionFrame.local with Cow<'static,[u8]>, borrow exact literal p/r/rPr/t values and retain the original owned Vec for every other selected byte sequence. No lifecycle, private-reader, Source or structural blocker was found for that exact proposal. Actual predecessor mechanism observations, implementation review and new ordinary elapsed qualification remain required. This is technical design closure, not implementation permission or a performance acceptance claim.

Zero existing correctness defects, zero smells and zero nitpicks were established. Root remains sole product/Cargo writer. Only this report was written. No product edit, compiler, runtime import, test, benchmark or instrumentation was executed. Prior design/implementation pass numbering and memory rider obligations remain unchanged.

## Actual source and elapsed boundary

The current OXML document.rs independently hashes to7f25e636789a9c659eeed13aadfb28a6275f5e24cc2a9ccea9383fd8cc934819, matching the frozen support8/Root26 runtime. The branch is sprint/s96 with the existing dirty final union. This remains integrated S96/F-X275 remediation consuming the historical F-X268 contract.

The new ordinary summary at `.claude/scratch/S96-final-evidence/f268-admission-binding-release-qualification-20261108-1/monitored-run-1/ACTUAL_TIMING_SUMMARY.json:1` independently hashes to90cc5de0f4bf3044ee5aa34e8f5a698fbfe6c064bb61b9724ef7a8361bb91e98. It reports the actual compiler-free00:45:10to00:46:37UTC window,84observations,0monitor/compiler errors and77/79flat passes. Reply9.810292/22.950166ms exceeds19.715438ms. Document.add_picture12.672292/24.346583ms exceeds24.008438ms. Cell.add_picture, all5linear and both407gates pass. No successful retry or remaining-failure waiver is inferred.

The qualified ordinary wheel is c2383c5d669633d5df9071235eb6e6b7b94e0bd9cf96554025e0b38ea93a01b6 and Native09397c7f0619b130e3b744c90535ba422304e06be290864910f0238d42859ed5. Those are the summary's producer identities, not new artifacts created here. The prior symbol packet concerns reply and Cell at its original producer. The533-line frame algorithm was unchanged by design11, but Cell's old sampled call graph is not evidence of Document.add_picture's current stack or minimum state.

The retained design11 mechanism observations show classified event reads grow with the document, while temporary namespace-pair copy counts stayed constant between its small/large public fixtures. That distinction motivates observing frame-local copies. It neither identifies their elapsed contribution nor proves that this change will close either current miss.

## Complete constructor and consumer map

`crates/rdocx-oxml/src/document.rs:1129` defines WordAdmissionFrame. Its local field at1145 is currently Vec<u8>. Repository search finds only this definition, the classified-reader stack at9122 and the sole constructor at9566. There is no clone implementation, retained model owner, external user or public signature for this type.

The sole local construction at9594 occurs inside the existing Event::Start branch, after all current namespace, admission, control, owner, work and overflow decisions. The selected bytes are exactly `consumed_owner.filter(|_| virtual_owner).unwrap_or(local.as_ref())`. virtual_owner itself is computed at9198. The field is copied even for unadmitted/opaque Start frames. Empty events do not push a frame and currently do not make this Vec copy. Do not turn the proposal into a general per-event name store or move it earlier in the scan.

All local consumers are in this classified reader:

- sdtContent boundary9144and legacy paragraph/run qualification9204.
- parent-name dispatch9256through admitted_word_child1189, drawing parent9269and MC parent checks9278/9282/9295.
- VML pict/object/group/shape checks9320through9325and direct drawing9361.
- textbox9423and ordered-field backward charging/stopping9441/9445.

No consumer mutates local, interprets its allocation identity, borrows it into a receipt or uses its Vec capacity. The mutable backward frame walk changes cost/ordered-field state, not the name. At End, pop consumes work/copy/control/namespace state at9606, then the ordinary field drops. Reader buffer reuse occurs after that event at9636. The local name is pure owned bytes with no custom destructor or hidden XML work.

Byte equality and slice patterns are the entire local contract. Namespace and admission remain separate fields and computations. For example a foreign x:p still has local bytes p and may use static storage, while its word/admitted flags and foreign binding remain exactly unchanged. Borrowing a common local spelling must never promote its namespace or payload grammar.

## Minimal code shape and invariants

Use the concrete field type `std::borrow::Cow<'static,[u8]>`. No struct lifetime parameter or project generic is needed. At the existing field expression, first select the identical consumed-or-local bytes, then match their complete value:

```rust
local: match consumed_owner
    .filter(|_| virtual_owner)
    .unwrap_or(local.as_ref())
{
    b"p" => std::borrow::Cow::Borrowed(b"p".as_slice()),
    b"r" => std::borrow::Cow::Borrowed(b"r".as_slice()),
    b"rPr" => std::borrow::Cow::Borrowed(b"rPr".as_slice()),
    b"t" => std::borrow::Cow::Borrowed(b"t".as_slice()),
    local => std::borrow::Cow::Owned(local.to_vec()),
},
```

The explicit slice coercions make every Borrowed arm the same concrete byte-slice type. They borrow static literals, never element/Reader buffer storage or the caller's consumed-owner slice. The Owned arm is the original byte copy for all other names. Keep it byte-oriented: no UTF8 conversion, lossy string, interning, lowercase normalization, qualified-name match, prefix stripping beyond the existing LocalName, length shortcut or expanded finite set.

Replace only frame-local `.as_slice()` uses with `.as_ref()`, and express frame-local equality against the returned &[u8] where required by Cow's trait resolution. Do not mechanically rewrite unrelated local variables or owner.local fields elsewhere in the file. Preserve every equality, pattern, Boolean grouping and short-circuit order. Standard Cow has actual Borrowed and Owned consumers today, owns the invariant directly and adds no forwarding construct, trait, helper, module, cache or API.

The value invariant is `new_frame.local.as_ref() == old_selected_local` at every real push and every later read, including after arbitrary buffer clearing/reallocation. Only storage ownership changes. The original dynamic selection runs before matching, so a virtual p/r is not accidentally stored as the synthetic rdocxAdmission spelling. Ordinary document roots and that synthetic root remain Owned unless the selected value is exactly one of the four names.

The existing debug assertion at9115restricts consumed_owner to p/r. Preserve it and the current caller preconditions. The fallback arm still handles arbitrary consumed-owner bytes wherever the old non-debug algorithm is callable. Do not weaken the assertion or add a new owner restriction merely to make a diagnostic easier. Tests must distinguish existing debug preconditions from the release byte-copy fallback.

## Error, Source and private-read assessment

The change adds no fallible interpretation or validation. The same event/attribute decoder, QName, namespace declaration checks, consumed-owner/part-root facts, table/control/weighted depth, copy visits, overflow arithmetic, constructors and source-map offsets remain authoritative. WordAdmissionFrame.local is neither an input source nor a Field/read receipt.

Keep the local construction after the original guards and inside the original Start branch. No parser/read/capture call, constructor action, frame push/pop or namespace undo is removed. Unknown, Unicode, malformed or differently cased names take the original owned-copy behavior whenever the reader reaches this constructor, and retain the original failure if the reader/guard rejects earlier. No new malformed-input permission follows from static storage.

All private Field reads, including the transitive XML work of unsupported model equality, remain outside this field's ownership change. The ordered-field backward walk still sees the same local bytes and booleans at9441. MC/VML/drawing admission still uses namespace and capture-depth state. No source authority, cache hit, input allowance or error cache is created.

Cow may have different stack/Vec element layout from Vec. Do not infer a smaller WordAdmissionFrame or total memory saving. The scan stack and all nonlocal field allocations remain real costs. The final memory/calibration riders must include the actual resulting representation. The proposal adds no retained AST or per-edit history.

## Required observations and risk controls

1. **Actual predecessor construction work.** Before permanent edits, observe the existing selected-value `.to_vec()` at its real construction, after dynamic virtual-owner selection. Count actual frame pushes and local Vec constructions/copied byte payload separately from reader attempts/consumed bytes. Distinguish p, r, rPr, t and the complete other-name bucket. Do not call XML events or loop visits local allocations. Empty events and failures before construction contribute no constructed local. Preserve the original-source capsule and exact restoration.
2. **Actual successor discrimination.** Instrument the actual static Borrowed arms and actual Owned fallback construction. Require positive predecessor common-name Vec copies and positive new static views, with0owned construction for those four names. Fallback construction count and copied bytes must remain equal for the same complete input/action. Total pushes and actual classified/remaining reads must agree. Borrowed view byte lengths are referenced payload, not copied bytes. HashSet, namespace undo/prefix/URI, legacy_run_prefix, event buffer, stack and output copies remain outside this narrow meter. A sampled stack or an always-incremented marker cannot replace this evidence.
3. **Complete actual outcome comparison.** Reuse the actual96-case old/new admission/remaining-reader comparison mechanism from support8, executing the real predecessor and successor. Compare full work, constructor/control facts, projected aliases, source-map correspondence and exact Input/Resource error payloads. This is new source evidence, not a relabelled old pass. The envelope helper and borrowed namespace policy should remain source-identical to support8.
4. **Name, event and lifetime discriminators.** Add focused real-reader cases covering nested/repeated common Start names, mixed Empty/Start forms, p versus r parent dispatch, an actual rPr subtree and text owner, unknown names, Unicode names, longer names sharing a common prefix and differently cased names. Use enough intervening bytes/events to reuse/grow the reader buffer before the parent is consulted. Include the synthetic/actual document root, consumed p/r virtual roots, and foreign/default/alias/shadowed bindings with the same local spellings. Preserve exact bytes in the Owned route, not an assumed string representation. A diagnostic trace of actual stored local bytes may compare real frames, but is not a performance work count.
5. **Context and refusal controls.** Include MC AlternateContent Choice/Fallback, admitted versus opaque VML/textbox and drawing anchors, ruby, revision/control/ordered-field chains and foreign same-local children. Retain inclusive resource boundary/sibling reset, ordinary typed/raw fallback, malformed/mismatched closes, bad/reserved namespace attributes, EOF, piecewise input and first-cause controls. Meaningful existing anchors include resource tests at22956/22969, C1 revision error ordering18623, captured namespace18080/18112, table nesting22556and every-byte piecewise25347. Wrong mapping p to r, truncating unknown names, storing event-buffer slices, treating Empty as a pushed parent or folding namespace qualification into the static match must be distinguishable.
6. **Actual public states and bytes.** Measure diagnostics for original reply action2, Cell picture action1 and Document picture action1andaction2, both sizes and predeclared trials. Document's current small/large minimum states differ, so neither old Cell profiling nor a single favorable Document state supplies its evidence. Use exact original preceding actions/target preparation. Full public semantics, held/fresh/refusal behavior and reopen precede copy assertions. Preserve the actual picture-only paragraph API behavior. Reuse the exact predecessor input ZIPs and compare complete successor ZIPs and every part without normalization. Report each state separately and preserve failures. These are mechanism controls, not elapsed retries.
7. **Implementation and final acceptance.** Obtain affected scoped all-target/strict-lint checks and Native/installed controls over a fresh uninstrumented debug artifact, then independent implementation support and Root review. Only a fresh ordinary final-source artifact and unchanged compiler-free complete79flat/5linear/both407protocol decide performance. Preserve the latest failures and passing Cell/cold controls. No thresholds, selected timings, fixture, font, workload, Source policy or final rider is changed by this proposal.

The proposal is bounded enough for explicit plan/HLD adoption and implementation after its actual predecessor observations are retained. It may still be insufficient to repair elapsed behavior. No wider frame representation, namespace rewrite, table/comment certificate, Source batching, memory policy, status, ledger, commit, push, tag or publication follows from this investigation. The pass ends with this report.
