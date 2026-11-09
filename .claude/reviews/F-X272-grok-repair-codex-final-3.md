# F-X272, Codex final 3, Grok source repair

REJECTED INCREMENT. This is the actual Root affected final after Grok implementation, not peer 4 or all-pass-12. Source peers remain exhausted at three. No source or test edits, builds, tests, native actions or network calls occurred in this audit. The pass ends with this report.

## Identity and preserved work

Branch sprint/s96, HEAD e2a31e58a9b0af20f0ec374e3561c2e599abb0cd. Frozen four-file packet is /private/tmp/rdocx-cross-review/fx272-grok-final-repair-frozen. document.rs SHA-256 16b12ab9773b19c4453880e516342d6417ae204c1128ea41a005192953dc12ea. text.rs 06955c248a0c073995408b96575fe02c25635dece2d08271ca2773a26bf2fef0. placeholder.rs e41ea5c2170abf5a331422eaaea4fce8e8e655dc3b00efe8ce190c7b41c3a92d. Progress c41a276d41a27ec54e2f19e9d72d39a3a405fe11f8b9aa54c3b8edff018f3231. Root authenticated all four against the actual handoff, 3862 other pins unchanged and no new untracked paths. The F276 draft is unchanged at 72508420b52107a5fa5c6efccd9150da092766d9c0dd420c096de8d29a5d865a. Source lease returned before this audit.

## Prior dispositions

C2 is closed for this increment. Record wrappers remain in needed storage and no longer cancel through extra allowance. The same retained_owner_work formula is used. C3 is closed for the reviewed paragraph replacement path: the public literal and regex entries stage the entire paragraph, private removals propagate errors, wrapper write-back refusal reaches the outer candidate. The nested fault-injection witness is labelled honestly. It is not a natural budget-exceed RED or whole-facade acceptance. Broader union and caller evidence remains pending.

C1 fixes Empty inline/anchor, unknown Word-local raw contentPart and the supplied opaque and dropped run-child cases. C4 distinguishes shared, run-local and physical drawing copies on the supplied cases. C5 removes the statically visible per-owner URI clone by keeping pooled indices. These are useful fixes, but the remaining mechanisms below prevent closure.

## C1, high, default-namespace dependencies are still omitted

crates/rdocx-oxml/src/document.rs:659 records an element dependency only when the name has a colon. An unprefixed opaque element therefore contributes no empty-prefix dependency. The actual discovery at crates/rdocx-oxml/src/text.rs:19628 uses qualified_name_prefix(...).unwrap_or("") and records that dependency. CT_R stores its selected binding in retained_bindings at text.rs:4964. For a document with canonical prefixed Word elements, a long inherited default namespace and many short runs containing <opaque/>, each run needs a default-namespace record, while this preflight predicts none. This bypasses the aggregate preconstruction guard required by .claude/plans/F-X272-design.md:3653.

Discriminator: pair many <w:r><opaque/></w:r> owners under one long xmlns with the equivalent prefixed opaque child. Compare actual selected namespace records and needed bytes, with one-owner, locally redeclared and xmlns="" controls. The genuine repeated inherited default must refuse before expanded record construction when its unchanged envelope is exceeded. No runtime RED is claimed here. Include drawing raw and QName-valued controls without applying a default namespace to unprefixed attributes.

## C1, medium, modeled leaf subtrees still invent run owners

crates/rdocx-oxml/src/document.rs:504 marks rPr as modeled but does not mark its descendants as a separately consumed property subtree. At document.rs:521 any Word r or hyperlink with neither raw nor dropped parent becomes an owner, independent of the actual enclosing parser. CT_R consumes the entire started rPr with capture_element and parse_scoped_rpr at crates/rdocx-oxml/src/text.rs:4759. It does not recursively construct CT_R owners inside an opaque property extension. A preserved extension such as <w:rPr><x:box><w:r x:keep="1"/></x:box></w:rPr> can therefore manufacture a root-attribute record in preflight that the run parser never constructs. This is the same C1 constructor-selection invariant, not a request to reject arbitrary imported extension XML.

Discriminator: repeat a property extension containing a Word-shaped run with a long inherited x binding, alongside genuine modeled nested runs and the already repaired opaque direct-run-child control. Trace the actual property parser and preserved bytes. Make capture-owner selection follow the consuming parser boundary, rather than adding another isolated local-name exception. These are static cases, not executed assertions.

## C4, medium, redundant drawing declarations are charged as lost bindings

crates/rdocx-oxml/src/document.rs:719 classifies a binding at owner depth plus one as drawing-local. At :728 it charges one injected copy for each physical raw solely on that depth. The actual drawing parser computes lost_bindings by comparing each prefix and URI with the ancestor frame at crates/rdocx-oxml/src/drawing.rs:2127 and injects only lost bindings at :2164 and :2195. A drawing that redundantly declares the same x URI already inherited from its run does not inject it into either raw. CT_R instead retains the used ancestor binding once on its shared record at text.rs:4973. The planner charges two raw copies and misses the actual shared-record shape. At enough owners this can create a stronger refusal than the unchanged actual-storage envelope permits.

Discriminator: one inherited x binding, each run containing a drawing that repeats exactly the same xmlns:x, with both nonempty inline and anchor using x. Compare prediction with actual retained record bytes, and contrast a genuinely rebound URI and a declaration absent from the run frame. Use counts selected from measured formula values, not an assumed repeat count. Preserve the repaired two-physical-copy refusal for a genuinely lost drawing binding. No runtime observation is claimed.

## C5, medium smell, the planning counter measures a constant

crates/rdocx-oxml/src/document.rs:169 declares ORIGINAL_CAPTURE_PLANNING_URI_COPIES. Its only other uses at :6719 and :6721 reset and read it. No allocation or copy site increments it. The zero assertion at :6929 is therefore tautological, including on the old per-owner-copy implementation if this counter is merely added there. The handoff candidly acknowledges this limitation. It must not be called measured zero planning-copy work.

The indexed production representation is an actual static improvement. Replace the vacuous proof with meaningful instrumentation or a bounded representation/work witness that would fail on the previous copying mechanism. Clearly separate physical declaration decoding, planning dependency storage and expanded record construction. Keep instrumentation local and avoid a speculative allocation framework. A genuine old-code RED is required only when actually obtained, never relabel a fault injection as one.

## N1, low, preserved F276 hash is inaccurate

The TEMP handoff /private/tmp/rdocx-cross-review/fx272-grok-final-repair-handoff.md:41 reports 229fe4a7426fe3966da4ee64ff36bd06c984ae4a59bdb325ee0eb3b522dcf714. Root independently found both the before manifest and current canonical draft at 72508420b52107a5fa5c6efccd9150da092766d9c0dd420c096de8d29a5d865a. There is no protected-file mutation. Correct this provenance in the next handoff without rewriting the immutable prior handoff or attributing a source change that did not happen.

## Evidence and remaining gates

The producer reports six capture discriminators failed on old source, 24 capture/weighted/table cases passed on the revision, 60 placeholder cases passed, the nested candidate control passed and scoped check passed with two existing table warnings. The reviewed source and recorded receipt names support only this scoped handoff. Root has not rerun the tests or certified a full gate. The heavy 2000-box case is honestly classified as capture refusal, with a separate lighter weighted refusal case. No limit or admission order change is accepted.

Zero additional findings in public signature changes or out-of-lease source modifications. Fmt and clippy were not run by this repair. Existing warning/format debt, D8, all-pass-12, Issue430 ownership and selective cleanup, unknown old facade result and final schema/native/hash/CLI/Python/typing/wheels/WASM/archive/hosted CI remain pending. No provider completion, claim, release or sprint closure follows. Return to orchestration for a distinct implementer investigation and repair, followed by affected Root review. No peer 4.
