# S96 F268 remaining performance, design investigation 5

## Scope and disposition

Read-only design investigation5 retains all earlier pass numbers under the user's authorization for continued passes. Root remains the sole product/Cargo writer. The accepted product union is frozen at support3/Root21. This pass ran no Cargo, runtime import, test or benchmark and changes only this review record.

**Cold disposition: the proposed per-slot scheduling contract is technically sound for adoption with the ordering and memory conditions below.** It preserves the same complete equality relations and successful finite admission, rather than weakening either. A wall-time benefit from locality remains unproved and must be decided by the existing ordinary elapsed gates.

**Reply disposition: needs actual frame-refusal/read/cache observations before product adoption.** Eager reconstruction before exact frame refusal is a concrete source fact. A staged/reusable model check alone does not establish that skipping reconstruction is free of shared cache or fresh-reader effects. A narrower already-proved, no-shared-fill refusal can be justified if the actual candidate qualifies.

There are zero newly identified correctness defects, zero structural smells and zero nitpicks in the frozen accepted source. Neither pass4 candidate is selected for implementation by this report.

## Current diagnostic applicability

Packet `.claude/scratch/S96-final-evidence/f268-current-phase-diagnostic-20261107-1` identifies a temporary diagnostic release module1eebfa46a64df53b5b93e42d2d9c5e471c98e945ee30de9ad6e7dc2a1ec19287 and a separately pinned diagnostic patch. OBSERVATION.json records12 successful original semantic/reopen cases. PROFILE_POINTS.json matches the actual observe.stderr SHA-256420574712eb0adbbed1f8cdfa2846ac7b1a4114b737db1aed4eaf89e19e86a83. RESTORED.json records all4309 shipping inputs restored exactly. These clocks perturb operations and are not shipping acceptance.

Logical/marker binding copy payload is constant between original25/100reply sizes:664copies/766910UTF8bytes and251copies/288329UTF8bytes, with roughly0.15ms and0.065ms clone intervals. Cold1000 has1001temporary fact-copy pairs,137345UTF8bytes and6001run paths/segments, around0.04to0.067ms. These findings do not justify selecting pass4's allocation candidates to close the remaining misses.

The current large cold trace instead separates roughly0.56to0.73ms opening proof, another0.53ms staged-model alignment and another0.68to0.74ms finite facts, followed by body-piece and inventory construction. Activation completes around2.95to3.42ms before later splice/native work. These inclusive/cumulative observations locate repeated body traversals. They do not establish memory-locality causality or predict the result of interleaving.

The latest actual ordinary result remains the compiler-free immutable-scope/model release qualification reviewed in investigation4: reply fails its unchanged flat limit, and cold insertion/removal fail3ms. No diagnostic timing supersedes those results or supplies a favorable retry.

## Cold contract: same proofs scheduled per physical slot

### Exact compared state

The current query at `crates/rdocx/src/document.rs:34818` first proves actual opening membership, writer-proof availability, namespace/plain-local context and complete current/opened CT_Document equality at34829. It separately aligns all optional staged models at34851, then performs the finite grammar/property walk and constructs facts at34864. The proposed change interleaves those existing body operations. It does not use staged models as opening authority or substitute pointer identity for the current equality relation.

CT_Document is exhaustively defined at `crates/rdocx-oxml/src/document.rs:13588`: body, extra_namespaces, background_xml, background_extra_xml and root_attributes. CT_Body at12972 contains content and sect_pr. Destructure **both actual borrowed documents** and their bodies without `..`. Compare sect_pr, extra_namespaces, background_xml, background_extra_xml and root_attributes exactly, plus exact content length. Keep the existing opening membership, successful serialization, root/body declaration/binding and plain-local guards before creating facts. Obtain the actual kept main_document through its existing method, preserving its ordinary parse/proof behavior.

BodyContent has derived PartialEq at `crates/rdocx-oxml/src/document.rs:9733`. Use the same complete item equality on corresponding physical slots. Do not replace it with accepted text, partial properties, canonical XML, an address or a hash. In particular, Field equality has established projection semantics at `crates/rdocx-oxml/src/text.rs:8584`. This proposal neither strengthens nor weakens that existing relation. Its cached-result readers return borrowed views, not a new source receipt. The equality and finite readers used here operate on immutable values and introduce no parser/source work. Metadata comparison order can change because those comparisons are value-only and expose no recoverable error.

### Minimal concrete code arrangement

Keep the current public hidden signature and substantive query in the existing file. After the unchanged outer guards and exhaustive metadata/length check:

1. Select the optional writer record through the same offered_canonical_blocks, serialized-first precedence and exact length admission. Hold only its existing temporary lock/borrow. Initialize a model-alignment boolean and a temporary optional facts vector. Keep the call-local borrowed last-successful RPr qualification.
2. Walk the actual current/opened body slices in a single physical-order zip. Compare the complete current/opened item **first**. The first unequal item immediately returns Ok(None), even if a prior finite item was unsupported. There is no resource fallback on an opening mismatch.
3. At that slot, perform the same optional current/staged equality while its record remains aligned. A missing/unequal staged value disables models for the entire result. It does not disable reader facts. Retain the old all-short-circuit relation after the first staged mismatch and never search a later record as an alternative.
4. If finite facts remain eligible, execute the unchanged paragraph/run/text/raw/RPr/shading predicates for this current item and construct its accepted text/run paths. A finite miss drops/disables the temporary facts and stops subsequent finite work, matching the old finite walk. It must **not** break the outer opening-equality walk. Continue all remaining required current/opened comparisons, and optional alignment as required by the chosen unchanged alignment relation.
5. Only after complete opening equality succeeds, return the original full resource error if facts are unavailable by invoking document.validate_resource_limits exactly as at34930. If facts exist, their unchanged finite grammar still proves exact resource maxima(0,0). In a final cheap metadata zip, attach the actual staged model Arcs only when complete model alignment succeeded. Otherwise leave every model entry None. No partial early model loans escape.

One local finite-item closure or equivalent substantive block is sufficient. It must not introduce another source reader, wrapper, trait, generic, cache or representation. Do not duplicate the finite predicate into subtly different branches. Existing empty/one/many body handling and successful fact order remain exact. The final zip needs length admission before indexing, and still attaches no containing record/frame/opening.

This preserves the three logically separate outcomes: failed opening returns None without a model resource error, a fully matching but unsupported body executes the original full resource guard, and a fully matching finite body returns complete facts with all-or-none optional models. The old opening equality stopped before optional model/fact work when it failed. The new code may do pure temporary work for an equal prefix, but must never turn that prefix into authority for an unequal remainder.

### Failure, allocation and memory limits

Finite facts allocated before a later opening mismatch are transient and must be dropped on that return. Every such fact comes from a current item already equal to its corresponding actually admitted opening item. Its size is a prefix of the existing successful-facts result, not another full model or a second retained cache. Once a finite miss occurs, stop allocating further facts and release the prefix. Do not claim zero allocation on all None results or an unchanged failure-path allocation schedule. This is an explicit bounded scheduling tradeoff. No recoverable allocation API or error is replaced by a success.

Avoid cloning either CT_Document or CT_Body to make exhaustive comparison convenient. No source/model mutation, publication, invalidation cause, serial, generation, index or capture changes. The mutex must not be reentered from the per-item logic or the resource fallback. The finite readers/equalities do not call main_body_scans, so the existing temporary record borrow can span the interleaved loop. Release it before invoking the full resource fallback. Contention or locality benefits remain measurements, not semantic proof.

### Required controls and acceptance

Inline Native controls should compare the new query against the exact old query/reference outcomes, complete facts and Arc-loan membership. Cover empty/one/many, both w14 root forms, all five metadata components, body-length changes, first/middle/last opening mismatch, missing/failed opening proof, and package/namespace/plain-local mismatch. Pair a supported equal prefix with a late mismatch. Include an earlier finite miss followed by a later opening mismatch whose current model is over budget: the outcome must be the original None, not a prematurely recovered resource error. For a fully equal unsupported over-budget body, preserve the exact original resource error, using a clearly labeled internal fault setup if it cannot be a valid public opening.

Test absent writer record, wrong length and first/middle/last staged mismatch. These keep facts where eligible but make **all** models None, including entries before the late mismatch. Keep positive actual Arc identity, held-value lifetime and no old package retention. Exercise actual RPr equality reuse and late nondefault/raw/Field/property refusal, rather than only plain strings. Test that no borrowed Arc/fact survives a failed whole proof.

Existing installed cold insert/remove, warm, scope, splice, callback, package, related/raw/Field and refusal controls retain their exact assertions. Returned values, held/fresh/reopened behavior and complete predecessor package comparison remain necessary. No artificial loop-order marker or constant counter test should be invented. This is a scheduling change with the same meaningful work. The actual ordinary predecessor cold3ms failures are performance RED. A new authenticated ordinary successor running the unchanged protocol supplies performance GREEN or remains failed. Diagnostic phase comparisons can explain scheduling, but cannot waive that gate.

## Reply: early exact-frame refusal needs a stricter purity boundary

### Actual source order and callers

`crates/rdocx/src/document.rs:28503` receives an owned MainPartBlocks and first calls proved_document at28509. Only afterward does it reject prefix/head and suffix/tail differences at28520 and28531. proved_document qualifies current partition/frame at10821, reads/clones every child at10858, clones the shell, and calls store_empty_namespace_owners at10876. This reconstruction can be discarded by a later exact frame refusal.

source_blocks invokes this helper on a current source record, a canonical writer-derived candidate, an actual producer source and a whitespace-adjusted candidate at28269,28275,28283 and28332. The actual source setter consumes Some or runs the existing complete parser fallback at16466 and16470. package_main_part_blocks also uses source_blocks at28614. The new current diagnostic identifies source_blocks inside the100reply set-source window as substantial, around20ms. It does not yet record how much of that work occurs on an eventual prefix/tail mismatch.

### Why staged/reusable alone is insufficient

BodyBlockScan.read_child at11975 deliberately fresh-parses lexical Field/private shapes. Otherwise it can initialize the shared proved OnceLock at11979 or fresh-read on a scope mismatch at11988. A staged value is writer provenance, not necessarily an existing actual source parse. Merely finding staged.owner_facts_reusable_by_value does not prove that proved_document would use an already cached reader value.

The record's own OnceLocks are copied into fresh locks by MainPartBlocks::clone at9731. Its blocks, however, are shared Arc values at9761. Consequently proved_document can populate shared child.proved or shared child.empty_namespace_owners before the failed local record is dropped. store_empty_namespace_owners at11213 can call empty_namespace_owner, which fills positive or negative child results at9296 through9324 and can scan declared children. Its record-level empty-owner result is local to the clone, but those child cache effects need explicit treatment.

Skipping these reads could change private/Field allocation order, shared cached success/failure occupancy, later proof availability and observable work. An early refusal does not grant Source authority, but that alone is insufficient to preserve the existing read/cache schedule. This is a concrete design obligation, not a defect in the current source.

### Conservative possible contract

An early exact-frame refusal may be adopted only for a positively identified pure path. Compute the **same exact** current/xml head/prefix and tail/suffix refusal with the same checked lengths. Return None only when the old code must ultimately return None. Leave the established successful reconstruction, current/prospective child reads, replacement parsing, resource validation and partition checks unchanged. A compatible frame executes the old order. This is a refusal optimization, never a new equivalence or source proof.

A sufficient narrow gate would establish that every actual child read is already positive under this exact current frame, not reads_anew, lexical Field-free and owner-facts-reusable, and that the skipped store_empty_namespace_owners path performs no new shared child fill. A staged value or cached parsed value may establish the reusable model predicate, but only an actual existing matching cached parse establishes no fresh source read. A previously populated record empty-owner result can make the store a no-op. Alternatively, already populated child namespace results plus exact existing policy can show no shared fill. Inspect existing locks with get, not get_or_init, when checking this gate. Unknown, missing, cached-failed, mismatched-frame, private and Field records retain the old order.

If the desired current caller would populate shared caches, that case is not covered by this pure gate. Root must either keep its old order or explicitly design and prove the changed cache/read lifecycle. Do not silently discard those effects because the local record is temporary. Do not broadly skip parser/input/resource work or move a prospective proof across mutation.

### Required observation before adoption

For each actual source_blocks_in attempt in original25/100reply, record call-site category, source/current lengths, first prefix/head and suffix/tail mismatch, and eventual outcome. Observe the **would-execute old path**, not a hypothetical marker. Separate current-context qualification, actual borrowed child clones, owned/fresh reads, frame clone and empty-owner work. Record before/after occupancy, success/failure and exact frame of child.proved and child.empty_namespace_owners, plus relevant record frame/namespace/nested-owner locks. Preserve private/Field classification and reader-allocation observations. Count and time only actual work, keeping inclusive clocks separate.

Then show a successful real public predecessor operation whose expensive reconstruction actually precedes the measured exact-frame refusal. If no such attempt or no pure qualified attempt exists, this candidate is not selected. If it exists, require exact old/successor fallback outputs/errors, complete packages and held/fresh/reopened semantics, plus actual removed reconstruction on that same refusal. Existing source normalization, context-change/poisoned-cache controls at63074, producer/source/writer separation and Field/private occurrence controls remain mandatory. No successful reuse path may lose its actual current/prospective reads.

## Handoff

Root may adopt the finite cold scheduling contract in the existing F268 plan/HLD before implementation. Its performance hypothesis is explicitly pending the unchanged ordinary gates. Reply remains needs-observation, with shared child cache effects included in the observation contract. Pass4's constant namespace-copy and small transient-fact-copy candidates remain unselected.

No dormant authority memo, revision coalescing, new cache/history/API, alternate parser, Source credit, resource/namespace relaxation, threshold change or favorable retry is approved. All remaining elapsed and final native/typing/memory/hosted/archive/full-union debts stay explicit. This report ends investigation5 without product changes, ledger/status changes, commit, push or publication.
