# S96 F268 remaining performance, design investigation 4

## Scope and finding

Read-only investigation4 follows the user's explicit authorization for as many passes as required. Earlier design investigations1/2/3, implementation support1/2/3 and Root21 remain retained with their original numbering and findings. No historical cap is reset. Root remains the sole product/Cargo writer, and the product/HLD union is frozen at clean support3/Root21. This pass changed only this review record and ran no Cargo, runtime imports, tests or benchmarks.

There are zero newly identified correctness defects, zero structural smells and zero nitpicks in the accepted repair. This investigation finds two finite allocation contracts worth evaluating. Extending immutable binding scopes to the two remaining scanners is supported by source-level policy equivalence. Consuming temporary plain paragraph facts is also bounded, with an explicit two-sink splice exception. Current-source observations must establish their relevance and size before either is presented as a repair for the remaining elapsed misses. No broader cold representation, guard removal or dormant coalescing is justified.

## Current evidence boundary

The actual ordinary shipping result is `.claude/scratch/S96-final-evidence/f268-immutable-scope-and-model-release-qualification-20261107-1/monitored-run-1/ACTUAL_TIMING_SUMMARY.json`, SHA-256ec7f894730134fdae0e598411e66d570afbe159589648a8a572c1893f2aca6bd. Its eight named evidence pins match actual files. The wheel is69c3213bf4981aa84866b47852c34cc6135356ecd2a43329671338fafc413e0e and Native module67e4f864c0903a43d7760f620cc6dc5ed0fdaba210e1bf8e2d10178b29e3b947. The record has96 compiler-free observations, no monitor errors or contamination, and78/79flat passing. Reply12.916167/32.596292ms exceeds24.3742505ms. Cold insert3.754ms and removal3.718333ms exceed3ms. All five linear checks, historical nongrowing checks, adapted public coverage and original primary407 pass. There is no favorable retry or threshold amendment.

The separate current shipping sample at `.claude/scratch/S96-final-evidence/f268-current-shipping-sampled-investigation-20261107-1/INVESTIGATION.json` records90 successful original100reply operations and two reopened results. Local Rust symbols are stripped, so raw sampled addresses cannot identify logical_owner_snapshot, marker scanning, allocation or another named hotspot. Its deliberately retained90fixtures also cannot supply a per-document memory result. Earlier phase diagnostics precede the accepted immutable repair and are not current causal timing evidence.

## Supported finite namespace contract

### Remaining source work and callers

`crates/rdocx/src/document.rs:3654` constructs each logical owner snapshot. It copies inherited bindings and appends prospective declarations at3662, then deep-copies the entire current binding vector on every Start/Empty at3678 and End at3715. It resolves element and attribute names, normalizes nonnamespace attributes, sorts attribute name/value vectors, and emits ordered structure/semantic events at3680 through3755.

`crates/rdocx/src/document.rs:4017` builds namespace markers. Its seed similarly appends prospective declarations at4034. It deep-copies current bindings at4058 and4132 and copies the pre-element inherited bindings into each expanded marker's pending stack at4096. It also clones shadow sets. The shadow sets are a different policy and allocation class and are outside the proposed binding-vector change.

These are live producers. namespace_owners_from_spans invokes both at4285 and4301, and repeats them for ambiguity/structural/namespace alternatives at4306,4330 and4354. Candidate marker matching calls the marker producer at4223. Namespace replay uses snapshots and candidate markers at4612 and4620. Comment-move source reconciliation invokes them at31360,31377 and31397. Thus a private helper change reaches ordinary namespace preservation, source/replay qualification and comment mutations, not only the timed reply. The helpers must retain complete unrelated source and owner errors.

### Exact implementation shape

Keep both existing function signatures, input XML, Reader loop, output types and callers. Within each function, preserve the current seed construction exactly: copy inherited bindings once, append all prospective declarations in their original order, then wrap that vector in Arc. Do not deduplicate, canonicalize or normalize the seed. inherited_namespace searches from the end at2235, so duplicates and append order can affect prospective lookup.

Use an immutable Arc binding vector in scope_stack. For each Start/Empty, invoke the existing substantive shared_namespace_bindings at3393. The helper first executes the real attribute iterator, shares only a declaration-free scope, and on a declaration clones the inherited vector then invokes the unchanged apply_namespace_declarations. That decoder at3015 remains the sole normalization and ordered retain-then-append policy. This introduces no new forwarding helper, parser, trait, generic, module, cache or source authority.

For logical_owner_snapshot, retain all name resolution, default-namespace versus attribute rules, attribute iteration/normalization, sorting and event production. End may borrow the current immutable scope for name resolution, then pop it. Do not skip or fuse End events. Start/Empty still differ only through their current emitted start/end sequence. Whitespace, CDATA, references, comments and processing instructions keep their existing distinct structure and semantic entries. The existing first-error sequence remains: apply namespace declarations first, resolve QName, then parse/normalize nonnamespace attributes. The helper's declaration-free iterator validates the same attributes that the former first apply pass validated.

For namespace_owner_markers, preserve direct_drawing_child computation from the parent scope before processing the current element at4045. Keep shadow_stack and local-prefix/redundant-binding logic exactly as at4060 and4134. The pending marker tuple must retain an Arc to the **pre-element inherited scope**, matching4096, not the new post-declaration bindings. This scope later supplies marker normalization at4195. That distinction is required when the element shadows a binding. Keep the existing drawing closure/admission checks at4104, actual raw slices and offsets, marker normalization, candidate matching and namespace facts.

Keep empty_markers and expanded_markers separate until their existing append at4213. Their output order is empty markers first and expanded markers in completion order. Changing to one traversal-order collection is outside the contract. namespace_marker_without_owner_declarations at3766 retains its current root vector, drawing validation, normalization and error policy. Do not broaden the change merely because it also contains a copy.

This is a complete finite allocation-only policy contract. Its compatibility follows from using the same ordered values and decoder at the same event boundaries. It is not a claim that those allocations dominate current reply time. Before product adoption, Root should observe actual calls and copied payload on the current source and record that factual selection input in the existing F268 plan.

### Discriminators and coverage

Retain test-only verbatim predecessors of the two substantive scanners, rather than a second production parser. Compare exact LogicalOwnerSnapshot structure/semantic vectors and exact NamespaceMarker raw/fact vectors, including order, for zero/one/many elements, Empty versus expanded pairs, nested siblings, prospective duplicate/overriding bindings, root/default/alias/redundant/shadowed/undeclared prefixes, xml attributes, QName-valued opaque payload, attribute permutations, entities, whitespace, references, CDATA, comments and PI. Compare actual first error family/message for malformed or duplicate attributes, invalid declaration values, malformed QName bytes, truncation and mismatched ends. Preserve any old accepted malformed shape rather than inventing stricter validation.

Exercise both marker modes and the actual drawing closure. Existing `fx211_drawing_marker_closure_preserves_owner_identity` at65755 provides positive inherited-drawing and negative foreign/default/opaque controls, but does not replace the new exact scanner comparison. Complete replay/ambiguity, comment move/reply, raw/Field and unrelated malformed-source controls must retain their old results, bytes and atomicity.

Observe real vector copies by caller and site: initial seed, declaration mutation, logical End, and marker pending inherited scope. Record calls and UTF8 name/URI payload separately from allocator overhead. The helper's existing modeled scanner row must not silently become a differently scoped counter. Add separately named caller observations or explicitly document an additive breakdown. Follow the existing reference pause only. A public predecessor must complete semantics before failing a positive actual copy discriminator. Successor expectations come from the actual remaining seed/declaration sites, not an assumed zero. Growing declaration-free descendants should add no inherited-vector deep copies. Preserve all produced namespace facts, complete packages and reopened/held outcomes before interpreting that mechanism as GREEN.

## Cold path: facts, mandatory ownership and finite candidate

### What the source establishes

The opened query still performs the complete opening/model/context proof, optional exact per-slot model alignment and finite grammar/property admission at `crates/rdocx/src/document.rs:34818`. It constructs fresh accepted text and run paths at34922. The cold binding build borrows each temporary batch entry at `crates/rdocx-py/src/document.rs:5647`, then plain_paragraph_owners copies its String and run-path vector at5443. The temporary batch is subsequently discarded. The direct per-block producer likewise owns fresh direct facts at5374 and copies them into that constructor before discarding them.

The next copy has a different lifetime. WordPieces::inventory copies paragraph text/run facts from retained active pieces into a combined inventory at5819. Both cache and active owner authority survive publication. word_piece_owner_addresses then moves text into Paragraph read facts and moves run paths into Run owner routes at5285 and5292. append_word_owners_with_states moves address/read facts into final owners at8358. Those latter moves are not hidden clones, and the retained-cache versus owner split cannot be removed by simply moving one value.

Activation also builds address indexes at12438, initializes exact scopes at12442 and installs generation lookup at12447. prepare_addresses reserves and indexes actual addresses at4602. These, subsequent splice planning and final Conservative generation/publication work are mandatory under the current contract. A current elapsed breakdown is needed before attributing the remaining0.7ms excess to any of them. No dormant fact memo, coalescing or skipped validation is proposed.

### Narrow consuming constructor option

If current observations show material body-piece cost, revise the existing substantive plain_paragraph_owners constructor to accept owned String and Vec<AcceptedRunPath>, moving them into its existing facts instead of cloning them. Its three callers are local and explicit at5374,5660 and6721. Do not add an owned forwarding variant or generic ownership abstraction.

The direct fact caller can move the fresh returned values. The cold batch can consume each owned triple once, moving its optional model Arc and facts together into the current physical slot. Keep the exact batch length check and cold-only gate at5592. With a nonempty old block list the optional batch remains absent. Do not compact, search by text, skip slots, or change unavailable-model fallback. If the batch is unavailable, execute the same current producer and clone route. All opening/resource/RPr checks have already completed and remain unchanged.

The splice caller has **two live sinks**. At6721 it constructs plan.piece, while6733 moves the original text into an address and6735 moves run paths into Run addresses. Preserve an explicit copy for the piece before those moves. That is required ownership under the existing representation and must remain visible in diagnostics. A consuming constructor must not make this copy disappear from accounting or consume values still needed by the plan.

This candidate removes only the temporary batch/direct-fact-to-piece clones. It retains the later cache-to-inventory copies, locations, physical projections, sizes, scopes, captures, failure order and all stable interfaces. Empty/one/many runs, zero-byte text, absent model records, unsupported/property/raw/Field misses, warm edits and returned splice handles keep their existing behavior. This source-supported ownership transformation is a candidate, not selected performance remediation until current-source cost observations justify it.

## Minimal current-source diagnostic contract

Root can prepare a scratch driver while canonical source remains frozen. After this read-only pass ends, any temporary diagnostic source must be separately pinned, built, installed, authenticated and restored exactly before a product change. A diagnostic build is not an ordinary shipping timing successor. Do not use sampled raw addresses to assign named Rust time.

For original cold100/1000 insert and remove, with the original cold/warm factory and no pre-clock handle acquisition, capture these nested boundaries:

- Opened query: opening/context equality, optional model alignment, finite content/property qualification, accepted-fact construction and any full resource fallback.
- assemble: current related-source discovery, body-piece construction, related-piece construction and inventory reconstruction. Separate direct/opened/fallback hit counts and actual text/run-path copies from final owner counts.
- ensure: validate_word_facts, address/index preparation, initialize_scopes and generation-index installation. Keep the existing splice plan, prepare, native edit, commit and mandatory final indexing boundaries distinct from activation.
- Reply on original25/100: the existing parent read, anchor, refreshed-source and publication windows, with logical snapshot and marker call/input/event counts plus actual copied bindings by site. Target acquisition through comments already activates authority outside reply timing at14998, so do not label this as dormant activation.

Use cumulative/nested timings with explicit inclusive and exclusive meanings. Do not add inclusive children to parents or interpret counter traffic as wall time. Record refusal paths, zero-duration/call cases, original workload/source/artifact identity and all samples, not only favorable medians. Existing debug guarded reference work stays separately identified. Current ordinary elapsed failures remain authoritative.

At the two proposed transient fact-clone sites, count actual clone calls, UTF8 text bytes, run-path records and path-segment payload separately, excluding heap/container overhead. Keep required splice copies and later retained-cache inventory copies in separate named rows. The actual old public operation must succeed semantically and then fail the proposed zero-transient-copy discriminator. A successor must remove those exact site observations while retaining all required output ownership, full facts and refusal controls. Helper-only or constant-marker RED is insufficient.

## Handoff

Namespace scope sharing has a finite source-supported implementation contract, pending factual current caller/copy observations and explicit plan/HLD adoption. The cold consuming-constructor candidate has a finite ownership contract and remains needs-measurement for performance selection. Broad inventory/state/representation changes remain unapproved.

After an adopted increment, require actual predecessor mechanism RED, semantic and package equality GREEN, affected scoped checks and independent implementation review. Only a fresh authenticated ordinary artifact running the unchanged79flat/5linear/primary407/secondary407 protocols can decide the three remaining elapsed misses. All final native, typing, memory, hosted, archive and full-union debts remain explicit. This report ends design investigation4 without product edits, status/ledger changes, commit, push or publication.
