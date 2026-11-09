# S96 F268 reader scopes, design investigation 7

## Scope and recommendation

Read-only investigation7 retains all previous pass numbers under the user's continuing authorization. Root remains the sole product/Cargo writer. The plan's same-proof cold scheduling amendment and current progress record were read. Support4/Root22 are clean, with no new ordinary timing qualification. This reviewer ran no compilers, runtime imports, tests, benchmarks or timings and changes only this report.

**The finite allocation contract is technically supported:** change the two existing word_prefixes_at helpers to borrow declaration-free inherited scopes, preserve each helper's distinct policy, and defer three unused namespace-vector constructions in the actual run/paragraph readers. Root can adopt this exact contract and HLD impact before implementation. Actual mechanism and ordinary elapsed results remain required. There is no claim that this increment alone closes the remaining elapsed misses.

Do not include merged_owner_bindings, the legacy consumed-paragraph namespace predicate, writer-roundtrip reuse or Source batching in this increment. There are zero newly identified correctness defects, zero structural smells and zero nitpicks in the accepted source. The reported allocations are existing work, not correctness faults.

## Profile evidence and limits

The optimized-symbol packet is `.claude/scratch/S96-final-evidence/f268-optimized-symbol-diagnostic-20261107-1`. BUILD.json records child0, unchanged source and the sole release override `profile.release.package.rdocx-py.strip="none"`. The wheel is8010433e970f1fe1796bdfcee7eefef71ba15f9b8104a2421a8d15ebe5ea7234 and Native module7dfbaca9a4df4814c5b88297e74c778e599bff64e58d80424282029f7d2aa061. All ten recorded wheel payload hashes match the actual archive. OBSERVATION.json records90successful original reply operations and two reopened cases. This is a symbol-bearing diagnostic, not an ordinary shipping acceptance retry.

reply.sample.txt shows the successful source_blocks_in/proved_document/read_child/parse_one_child chain, actual CT_Document and CT_Body readers, parse_run_raw_with_observation and scoped RPr parsing. Both numbering::word_prefixes_at and properties::word_prefixes_at appear in those reader stacks. These are named allocations inside actual reading, unlike the rejected frame-refusal hypothesis. Sample counts are inclusive observations, not function elapsed fractions or predicted speedups. No numeric percentage of improvement is established here.

The source-part/frame reads still succeed and actual child-reader facts still start cold. Design6's Source-provider/admission/correspondence obligations remain open. This proposal leaves every actual reader, input, admission, cache fill and roundtrip in place.

## Two distinct borrowed-scope helpers

### Concrete signatures and lazy ownership

Revise each existing crate-private substantive helper to the concrete shape `word_prefixes_at<'scope>(start: &BytesStart<'_>, inherited: &'scope [String]) -> Result<Cow<'scope, [String]>>`. The return lifetime is tied only to the inherited scope, never the event buffer. Both Borrowed and Owned forms have real current consumers. No new forwarding-only owned variant, trait, generic abstraction, cache or scope representation is needed.

Initialize Cow::Borrowed(inherited), then execute the same one-pass attribute iterator in the same order, with the same error conversion, UTF8 interpretation, decoder and namespace detection. Ordinary attributes must still be visited even when no declaration exists. On a declaration, own the inherited vector at the original mutation boundary and execute the exact original mutation sequence. A declaration always retains its old owned path, including a redundant declaration. Do not optimize it by value equality, sort, deduplicate, scan raw substrings or skip malformed attributes.

### Numbering policy

`crates/rdocx-oxml/src/numbering.rs:644` currently clones inherited before iteration. At659 it decodes the declaration prefix, then normalizes the declaration value. At662 it removes both the plain Word alias and encoded namespace binding for that prefix, appends the new encoded binding at667, and appends a plain alias only for the Word URI at668. Preserve that exact order and all encoded entries, including default bindings and foreign namespaces. The new function must return precisely the old vector contents, order and errors when viewed as a slice.

### Properties policy

`crates/rdocx-oxml/src/properties.rs:429` is a different helper. It removes only the plain prefix at445, before decoding the namespace value at447, and appends that prefix only for the Word URI. It does not implement numbering's encoded-binding replacement policy. Preserve its own order and behavior, including inherited encoded values if present. Do not unify the implementations or import numbering policy here merely because both functions return Cow.

The properties helper is an actual named RPr-reader hotspot through `crates/rdocx-oxml/src/run_properties.rs:982`, and is also consumed by paragraph properties. Including these two independent helper bodies is finite. Each preserves its existing semantics and has its own exact predecessor comparison. This is not approval for a broad namespace-policy rewrite.

## Caller ownership and lifetime map

The numbering helper reaches actual run and paragraph parsing through `crates/rdocx-oxml/src/text.rs:11452`, scoped RPr/PPr at `crates/rdocx-oxml/src/numbering.rs:1492` and1421, recursive property projection at1244 and1257, and other existing document/part/table/control/revision/numbering readers. The properties helper supplies scalar/property readers, including the actual CT_RPr parser. Audit every direct caller and its imports, not only the named hot stack.

Immediate parser arguments, attribute tests and nested calls can borrow the resulting slice for the same immutable lexical lifetime. Keep those hot transient results borrowed. A declaration produces an owned vector independent of the event buffer. Nothing may retain a reference into BytesStart or a cleared Reader buffer.

Persistent sinks still require explicit ownership. Examples include validate_typed_leaf_prefixes' Vec scope stack at `crates/rdocx-oxml/src/numbering.rs:841` and877, numbering's cross-event root scope assignment at4598 and4611, OccurrenceOverlay.word_prefixes at2001 and2061, and text's owned source-probe stack at `crates/rdocx-oxml/src/text.rs:12513`. Convert with into_owned only at such actual storage/return/mutation boundaries. If a scope borrows an older stack frame, materialize before pushing or mutating that stack. Do not build self-referential Cow stacks or add unsafe lifetime extension.

Those necessary sink copies remain actual work and must be accounted separately from eliminated transient clones. No eager into_owned at every call site, which would preserve the cost while only changing its location. Conversely, do not change an owned field or private model into a borrowed representation just to avoid a required copy. The compiler's lifetime checks support, but do not replace, this consumer audit.

Namespace vectors remain data, not Source identity. XML buffers, Reader cursor positions, actual Word source maps, callbacks, constructor order and private Field ownership are unchanged. Cow must not be used to suppress admission or infer reader/writer authority. Removing pure String/vector copies can change allocation timing, not the actual Source origins or namespace values supplied to readers.

## Three lazy pure materializations in text readers

### Run external raw scope

At `crates/rdocx-oxml/src/text.rs:10465`, namespace_bindings builds an owned vector, filters canonical w and collects another vector before iterating extra_xml. namespace_bindings at `crates/rdocx-oxml/src/numbering.rs:636` and that filter are pure. Build this external vector only if the existing extra_xml loop has work. Keep the canonical-w exclusion and every external_bindings_used_by_raw call, first error, raw order and retained-binding deduplication at10469 unchanged.

### Run drawing scope

At `crates/rdocx-oxml/src/text.rs:10476`, a distinct unfiltered namespace vector is constructed even with no drawing raw payload. Initialize it lazily on the first actual inline/anchor raw value reached by the existing loop at10481. Reuse it for later payloads in that call. Keep the unfiltered policy separate from the external-raw vector, and preserve every fallible raw scan at10493 in the same order.

Crucially, raw_namespace_declarations(root) at10502 remains unconditional whenever root is Some, even with no drawings or raw payload. Its validation must not move behind drawing eligibility. Retained root-attribute capture at10458, local binding checks and all later writer/capture behavior stay in their original positions.

### Paragraph equation scope

At `crates/rdocx-oxml/src/text.rs:22012`, namespace_bindings is used only by OfficeMath::from_raw inside the existing run-index/raw-position loop at22014. Initialize once on the first actual raw visited there. Do not replace that loop with an XML substring or tag test. Every OfficeMath invocation, error and equation position retains its existing order. Root attribute capture remains after that loop at22027. Retained legacy namespace frame work at22034 remains unchanged.

These three changes omit only pure vector construction when its result has no consumer. No fallible operation, declaration validation or actual source read is omitted. No new helper is required.

## Explicit exclusions

merged_owner_bindings at `crates/rdocx-oxml/src/numbering.rs:1559` can theoretically borrow when local is empty, but it has little sampled representation in this packet. Nonempty equal-valued updates can still reorder bindings. Leave this helper and all its callers unchanged in the first increment.

The legacy_consumed_owner predicate at `crates/rdocx-oxml/src/text.rs:21292` short-circuits on actual roots. Replacing namespace_bindings with a duplicated decoding predicate is not needed for this contract. Preserve it unchanged.

The actual per-child writer roundtrip at `crates/rdocx/src/document.rs:11861` is another possible later cost. Any reuse must follow the actual source read and prove independent exact writer frame, private facts, staged model and byte correspondence. Sampled presence does not establish that it can close the reply gate. No writer-proof shortcut or percentage benefit is approved here.

## Required controls and genuine mechanism evidence

Retain each helper's exact test-only predecessor and compare complete ordered output and exact first errors. Cover empty/one/many inherited scopes, ordinary attributes without declarations, encoded bindings plus plain aliases, default namespace, canonical/alias/foreign bindings, redundant declarations, rebinding/undeclaration, repeated or duplicate declarations, mixed attribute order, malformed attributes before and after declarations, invalid UTF8 prefixes, entity/character-reference normalization and declaration values. Include a redundant declaration that changes vector order, and an inherited encoded binding demonstrating the two helpers' distinct policies.

Test Borrowed identity only as a positive successor ownership observation. A test that fails to compile against the old Vec return is not RED. Real predecessor mechanism RED must execute the old helper/parser and successful semantics, then observe actual inherited deep copies at their real sites. Instrument old inherited.to_vec and new first to_mut materialization, plus real persistent-sink copies, with calls and actual UTF8 String payload bytes. Declaring nodes and persistent sinks retain required copies. Do not hide them behind a borrowed-hit counter.

Instrument the three namespace_bindings materializations by actual reader/site with calls and copied prefix/URI payload bytes, excluding headers/allocator overhead. Ordinary runs with no extras/drawing raw and paragraphs with no visited raw should positively show old work and successor absence after semantic success. Rich raw/drawing/OfficeMath inputs must positively retain the same required work, ordered results and errors. Preserve malformed root declaration refusal even without a drawing, foreign defaults, producer metadata, inline/anchor differences, opaque payload and equation positions.

Use one predeclared old/new installed original25/100reply cohort with the same pinned input bytes, parent acquisition, fixed author/text/date and semantic/reopen protocol. Two observations per size are sufficient mechanism evidence if predeclared and retained in full. Authenticate separately identified diagnostic predecessor/successor source, wheel/module and payload. Both must succeed in parent/text/anchor, stale first-cause and held/fresh neighbor checks before counter assertions. Compare complete output packages and parts. Preserve exact before/diagnostic copies and source restores. Do not claim an absent counter/export or a compile failure as mechanism RED.

Risk-based Native and installed controls must cover actual run/RPr/PPr parsing, property scalar/error order, imported aliases/default/shadowing, root/child scope persistence, numbering/style/part readers, drawing/raw/OfficeMath, Fields/bookmarks/tab-stop identity, unknown metadata, unrelated malformed sources, C1/resource refusal and atomic publication. Existing input spans, read/publication maps, constructor/source-work metrics and held/fresh/reopened results remain exact. No parser-input or resource allowance can be relaxed to satisfy the allocation target.

After implementation, run changed-crate scoped checking, strict lint, format/prose/adapters/diff and independent implementation support5/Root review. Preserve all failed attempts and current performance RED. Only a fresh authenticated ordinary artifact running the unchanged full elapsed protocol can provide performance GREEN. The diagnostic profile, counters and semantic cohort cannot substitute for it.

## Handoff

Root can adopt two separate Cow helpers with exact policy preservation and the three lazy pure conversions, with a caller ownership inventory and HLD current-intent update. This is a finite existing-work allocation increment, not a Source-provider or namespace-equivalence change. Wider ideas above stay unselected.

All three last ordinary elapsed misses remain recorded. Cold scheduling has correctness acceptance but no new ordinary elapsed qualification. Final native, typing, memory, hosted, archive and full-union obligations remain due. This report ends investigation7 without product edits, status/ledger changes, commit, push or publication.
