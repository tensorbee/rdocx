# S96 admission context weight, design investigation 13

## Disposition

One call-local Option<usize> can safely retain the successful context-weight fold until the namespace map mutates. This exact proposal is technically feasible without changing the input language, Source authority, checked budget or first-cause schedule. It requires complete mutation invalidation, lazy evaluation at the original fold point, successful results only and the same cold fold. No existing correctness defect, smell or nitpick was established.

This candidate is unadopted. Design12did not authorize broader namespace work. Root must explicitly adopt the finite contract and HLD intent before implementation, preserve actual predecessor fold-work observations and obtain new correctness/mechanism review. The report neither grants implementation permission nor predicts elapsed sufficiency. Composing a later approved increment before an ordinary build is an orchestration choice, not acceptance or a timing retry.

Only this report was written. No compiler, product import, runtime, test, benchmark, instrumentation, overlay or product modification occurred. Root remains sole product/Cargo writer. All prior pass numbers, ordinary failures and final riders remain unchanged.

## Source custody

Frame12work is in progress, so this investigation used the immutable Root26source at `.claude/scratch/S96-final-evidence/f268-admission-binding-allocation-implementation-20261107-1/final-source/crates/rdocx-oxml/src/document.rs`. Its admitted binding allocation state is the support8/Root26boundary. Frame12changes only local-name storage and its mechanical reads, not the map/fold policy under review. The references below identify that immutable snapshot rather than falsely pinning an in-flight instrumented source.

The source anchors in that snapshot are:

- `.../final-source/crates/rdocx-oxml/src/document.rs:1124`, WordNamespaceBinding.
- `.../final-source/crates/rdocx-oxml/src/document.rs:1154`, restore_namespace_depths and both mutation arms.
- `.../final-source/crates/rdocx-oxml/src/document.rs:1167`, decoded_namespace_is.
- `.../final-source/crates/rdocx-oxml/src/document.rs:9129`, the call-local namespace map.
- `.../final-source/crates/rdocx-oxml/src/document.rs:9154`, attribute iteration, reserved binding validation and insertion.
- `.../final-source/crates/rdocx-oxml/src/document.rs:9188`, PartWrapper undo clearing.
- `.../final-source/crates/rdocx-oxml/src/document.rs:9496`, the complete checked context fold.
- `.../final-source/crates/rdocx-oxml/src/document.rs:9602` and9632, Empty and End restoration.

Here `...` expands exactly to `.claude/scratch/S96-final-evidence/f268-admission-binding-allocation-implementation-20261107-1`. These are source locations, not claims that a new diagnostic or final artifact was built here. The actual latest ordinary c2383cqualification still fails reply and Document.add_picture. Frame12correctness/mechanism work does not turn those failures into passes. No new elapsed observation was consumed.

## Complete mutable owner and dependency analysis

The namespace_depths value is a local HashMap<Vec<u8>,WordNamespaceBinding> created once per admit_word_input_classified call. WordNamespaceBinding contains only an owned URI Vec and depth usize. This inspected function inserts at the actual namespace attribute at9176. The only mutable escape is restore_namespace_depths. That helper iterates undo in reverse, then either inserts the previous owned binding or removes the newly introduced prefix. No clear, retain, entry/get_mut, mutable URI reference or other in-place update exists for this map. All other map/binding consumers borrow immutably.

The successful fold at9496depends only on these actual entries:

```text
prefix.len + binding.uri.len + 2
+ (prefix.len if decoded_namespace_is(binding.uri, W_NS) else 0)
```

The implementation adds these terms through the existing checked-add chain and returns the existing Resource error on overflow. URI length means the stored raw URI byte length, not the unescaped length. Aliases with different prefixes remain separate entries. An empty/default prefix has its original zero-length prefix/marker terms. The fold does not depend on stack depth, copy depth, word/admitted state, control counts, event offsets, owner kind, namespaces in another operation or an external policy. Binding.depth is used by other classification, but is not a fold term.

decoded_namespace_is at1167is a pure projection of its byte slice and the supplied static expected URI. It first compares raw bytes, then tries UTF8 and ordinary quick_xml::escape::unescape, converting either failure to a false classification through `.ok()`. The quick-xml0.41source was inspected at `/Users/atulsharma/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/quick-xml-0.41.0/src/escape.rs`, functions unescape/unescape_with. The ordinary function uses its fixed predefined resolver and scans the supplied string. It has no actual XML Reader, operation receipt, source cursor, user callback or input allowance. Its temporary decoded string is not a typed source read or identity-bearing model. Existing external namespace validation and every actual XML/constructor/capture read remain separate and mandatory.

Caching this successful scalar can avoid repeated pure classification/unescape work. It cannot remove the other decoded_namespace_is/captured_namespace_is calls that decide actual admission. In particular, captured_namespace_is at1175uses depth-sensitive exact-versus-decoded behavior. Do not change or memoize that different policy under this proposal.

An invalid URI escape that the projection currently converts to false remains false for this calculation. That is not a new cached Input failure. Actual attribute, namespace and parser errors still run and retain their existing verdict/order. Only the successful complete usize result is retained. A failed checked fold is never stored or reused.

## Exact invariant and minimal code

Add one `let mut copy_scope_weight: Option<usize> = None` next to the local map. At the exact original9496fold point, use Some(value) directly. For None, execute the original namespace_depths.iter().try_fold code and exact error conversion, then assign Some(value) only after `?` succeeds. Some(0) is a valid cached result, not an empty/uninitialized signal.

The invariant is: whenever the option is Some(v), v is the actual successful original fold over the current complete map. No mutation has occurred since that successful fold. The cache contains no map snapshot, URI, frame, AST, Source fact, version history or allocation. It dies with this one classified admission call.

Invalidate immediately before every actual map write, after the original attribute checks and in the original mutation order. At the attribute insertion, clear the option before namespace_depths.insert even if the prefix/value is redundant or the URI is byte-identical. Changing only stored depth also invalidates. Do not perform a fold, URI unescape, checked delta arithmetic or error creation at insertion.

The existing restore_namespace_depths helper may take one additional concrete `&mut Option<usize>` argument and clear it inside the undo loop before each insert/remove. Both existing callers pass the local option. The helper still substantively owns actual restoration, so this is not a forwarding wrapper or new helper. Alternatively co-located caller invalidation can be correct, but its completeness must be equally explicit. Do not hide the policy behind a new namespace provider, generic or mutable global.

Clearing undo for an actual PartWrapper is not a map mutation. Its declarations were already inserted and invalidated the option. The clear intentionally keeps those declarations rather than restoring them later. Retain exactly that behavior. An Empty/End restore with no undo entries performs no map mutation and need not invalidate. Every nonempty restore invalidates, including an insert of an equal previous value or a remove that would leave the aggregate unchanged. The conservative rule is mutation-based, not an optimization that compares old and new weights.

There is no cross-scope cached sum to restore from a parallel undo stack. After restoration the option is None and the next original fold point computes the actual current map. Do not restore a previously cached scalar based on assumed prefix/URI equality or infer it from undo length.

Keep WordAdmissionFrame.copy_scope_bytes as the original per-frame value assigned at Start. It describes that frame's entry context. End uses that stored value before namespace restoration. The new local scalar describes the current map and must not replace the frame value on End or cause charging under an already restored parent scope.

## Checked work and first-cause proof

For an unmutated map, all term bytes and the pure Word-URI classification are unchanged. A previously successful checked fold therefore produces the same finite usize at the next point. Reusing it cannot conceal a new overflow because no operand changed. A map mutation always clears it, so the same old checked fold executes at the same next eligible event before any successful result is retained.

Preserve the original order: attribute/declaration errors and owner/control/table guards occur before the context fold. After the fold, next_copy_depth is checked, then Empty copy visits/nodes/bytes and later Start frame installation proceed. On End, the existing frame-copy visits/nodes/bytes checks precede undo restoration. The final duplicate-context multiplication/addition stays after the entire scan. No checked multiply/add, copy-node increment or error is moved, merged or replaced.

The lazy fold must remain unconditional at its original Start/Empty location. It is not moved into only copy_fact, admitted or word branches. Its successful reuse is allowed even for an opaque owner, but those owners still receive the same fold value and subsequent classification/work schedule. The parser/read/attribute/constructor loops, namespace insertions, stack event count and source-map origins stay exact.

HashMap iteration order is not a new policy. The cold fold remains the original iteration and checked-add sequence. With no mutation the map is unchanged. After any mutation a new cold fold runs rather than trying to preserve an iteration-dependent accumulator. Do not substitute a sorted map, incremental wrapping/saturating sum, decoded URI-length sum or a final unchecked aggregate.

The change removes no read with private Field/source identity, no XML capture, no constructor fact and no Source debit. It computes no permission from the scalar. The scalar affects only the same already-existing copy-scope quantity. All acceptance comparisons still need the actual outputs and error/resource results rather than relying solely on this argument.

## Required actual predecessor and successor evidence

Before adoption/implementation, retain a source-bound actual predecessor observation at the real fold expression. Distinguish fold executions, actual iterator entry visits, URI classification invocations and the defined prefix/raw-URI input payload lengths visited. The lengths are inspected input payload, not copied bytes, exact decoder byte reads, heap allocation or time. An empty-map fold executes with0entry visits. Count the actual events/read attempts/consumed bytes separately.

The successor observes actual None/cold-fold executions and Some/reuse branches. Count invalidation at each real insert/remove, distinguishing writes from actual Some-to-None transitions when useful. An entry/loop marker cannot prove reuse. A cached-sum trace or comparison may verify the actual value at every original fold point without adding an unlabelled reference fold to the claimed successor work. If a diagnostic reference fold is run, its actual visits/unescapes are separately named and not hidden from total diagnostic work.

Require a positive old repeated-fold workload and actual new reuses/reduced entry visits, with the same read events/bytes, map mutations and complete work/error outcomes. Do not predeclare a reduction for a path before it is actually reached. Root's source-bound103complete admission outcomes from frame12can supply the reusable comparison machinery, but their prior pass is not an execution of this new policy. Execute actual old/new producers for this increment and compare their complete tagged outcomes, not only message text or count summaries.

The focused discrimination matrix must include:

| Input/state | Required discriminator |
|---|---|
| No namespace entries | Real successful0fold, Some(0) reuse and unchanged work. A truthy/nonzero cache test is rejected. |
| One/many persistent Word and foreign bindings | Reuse across declaration-free Start/Empty siblings, with exact per-entry prefix/URI/marker charges. |
| Different aliases/default prefix for the same URI | Each actual map entry remains charged. No URI-keyed deduplication or normalized-length substitution. |
| Escaped Word URI, raw Word URI and foreign/invalid escape data | Same raw URI lengths, same pure decoded marker and unchanged actual reader/input result. No canonicalization or new validation. |
| Shadowed inner binding with a different URI | Inner fold changes, End restore invalidates and the sibling receives the outer weight. This rejects missing restore invalidation. |
| Empty element with a declaration | Its work uses the temporary inner scope, then restoration invalidates before the next sibling fold. |
| Redundant declaration, equal URI and depth-only replacement | Every actual insert/restore invalidates despite unchanged final numeric weight. Verify real writes/cold folds, not only output equality. |
| Multiple declarations for one prefix and nested restores | Preserve original attribute/undo order, including actual duplicate-attribute failures where the old reader rejects them. No speculative normalized namespace set. |
| Actual PartWrapper declaration retention | Build through the real part-reader constructor facts. Undo clear retains bindings at the old depth/policy. Later folds and nested restoration see the actual map, not the pre-wrapper cache. |
| Invalid attributes/reserved binding plus a competing resource failure | Same original Input/Resource first cause. A warm successful scalar cannot bypass attribute iteration or move the weighted guard. |
| Control/table/ordered Field/MC/VML boundaries, opaque and failed parses | Preserve typed/raw/private routes, inclusive limits, late errors, source windows and sticky refusal. |
| Contiguous/piecewise, malformed close/EOF and unrelated late bad source | Complete outcomes and source-map correspondence stay equal. No previously cached scalar certifies an unread input. |

Keep the exact cold fold source as the overflow proof. Do not invent a public sum override, weaken checked arithmetic or call an artificial entry marker an actual resource RED. The existing real inclusive boundary/resource/error cases remain required, while impossible enormous fixtures are not fabricated as executions.

Run the predeclared original public reply action2, Cell picture action1 and Document picture action1/action2at both sizes, with actual preceding actions/target acquisition. Their semantics, held/fresh/refusal behavior and reopen precede mechanism assertions. Use identical actual predecessor input ZIPs and compare whole ZIPs and every part. Reader attempts/bytes and complete admission work must remain equal. Preserve all failures and unvisited-path disclosures. Old Cell symbols do not establish Document picture's work.

The normal measured product must contain no diagnostic trace or reference-fold overhead. Obtain affected source/check/lint and Native/installed correctness, an independently reviewed final union, and only then a fresh ordinary compiler-free unchanged performance qualification. All existing memory/Office/typing/hosted/archive/full-union riders remain due. No threshold, favorable retry, workload/font change or memory-policy relaxation is part of this candidate.

No concrete technical blocker remains for the exact scalar/invalidation contract. Adoption, actual mechanism RED, implementation/equivalence evidence and elapsed sufficiency are separate pending boundaries. This report ends investigation13.
