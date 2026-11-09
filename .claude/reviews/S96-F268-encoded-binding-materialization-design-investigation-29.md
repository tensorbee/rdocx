# S96, encoded binding materialization, investigation29

## Result and source boundary

A narrow direct String builder is semantically feasible in numbering::namespace_binding. It can preserve the exact existing encoded bytes and every caller's reader/namespace policy. Selection remains conditional on actual helper-site observations and complete scope/Source proof. This report adopts and implements nothing, claims no hidden standard-library allocation growth, and promises no elapsed improvement.

Read the current F268 plan/progress, investigation28, actual old28 producer hooks/observations and all namespace_binding call sites. CLAUDE.md and the canonical workflow remain controlling. Use the frozen before28 numbering snapshot at `.claude/scratch/S96-final-evidence/f268-foreign-prefix-storage-implementation-20261108-1/before/crates/rdocx-oxml/src/numbering.rs`, SHA-256925cb0e082696835e7468b38f1844521df7bcf955b96e273e83659499c4b37ef. Its bytes match that packet's BEFORE-inputs.json. Live numbering differed during Root's ongoing work, so this report does not authenticate transient live source as the frozen predecessor. The targeted helper is the original format expression at `crates/rdocx-oxml/src/numbering.rs:628` in that snapshot.

## Exact candidate

Keep the existing function, signature and owned String return. Compute prefix.len() plus namespace.len() plus2 with checked_add and require the result to be at most isize::MAX. When available, allocate String::with_capacity(total), push the NUL character, push_str(prefix), push another NUL and push_str(namespace), then return it. If checked sizing is unavailable or exceeds isize::MAX, return the exact original format!("\0{prefix}\0{namespace}") expression. No new helper, type, API, parser, cache or namespace representation is needed.

The exact result is one NUL byte, all prefix UTF8 bytes, one NUL byte and all namespace UTF8 bytes. The original format uses ordinary Display for two &str values, with no width, precision, alignment, escaping or custom formatter. String push_str copies those same bytes. Rust char NUL has UTF8 length1. Lengths must be byte lengths, not character counts. Empty strings, Unicode, XML-sensitive characters and embedded NUL remain unchanged.

In particular, do not repair the protocol's existing ambiguity for an embedded NUL in prefix. split_namespace_binding strips the leading NUL and splits once at the next NUL at `crates/rdocx-oxml/src/numbering.rs:632`. An arbitrary prefix containing NUL does not necessarily round-trip to the original pair today. The candidate must preserve the exact encoded string and the same split result, rather than introducing escaping or claiming a stronger round-trip invariant.

The overflow fallback preserves the original capacity/error route for unavailable totals. The normal branch adds no Result, new error variant or recoverable allocation policy. It changes allocation strategy, so identical allocator exhaustion points are not promised. All recoverable caller parse/decode/Source errors and their ordering remain untouched. There is no field access, user callback or fallible formatting trait implementation in this helper's two &str operands.

## All helper callers

Repository search found five production call sites, plus the cfg(test) legacy scope helper and test fixtures:

| Production caller | Existing owned consumer and ordering |
|---|---|
| numbering::word_prefixes_at at `crates/rdocx-oxml/src/numbering.rs:668` | Actual attribute iteration, prefix UTF8 validation, URI decode, inherited Cow materialization and exact-prefix retain precede this binding push. Optional bare Word alias follows it. |
| Choice-parent resolver snapshot at `crates/rdocx-oxml/src/drawing.rs:561` | After actual resolver binding iteration, prefix UTF8 and URI unescape, retains owned encoded entries in choice_prefixes. |
| Selected drawing projection at `crates/rdocx-oxml/src/drawing.rs:621` | Encodes already owned resolver pairs into the projection scope after the existing physical source and owner checks. |
| typed_num_pr_scope bare aliases at `crates/rdocx-oxml/src/paragraph_properties.rs:524` | Encodes existing non-sentinel bare aliases against W_NS, preserving iterator order and later exact duplicate suppression. |
| typed_num_pr_scope owner pairs at `crates/rdocx-oxml/src/paragraph_properties.rs:528` | Encodes each complete owner pair after the alias sequence, preserving namespace spellings and exact contains/push behavior. |

The legacy cfg(test) caller at `crates/rdocx-oxml/src/numbering.rs:700` also uses the production namespace_binding helper. Therefore an existing current-versus-legacy scope comparison alone would not discriminate a wrong change to this shared encoder. Keep a literal old format expression as the independent byte oracle for helper tests and diagnostics. Existing direct callers in revision, document and drawing test fixtures remain consumers, not additional production call sites.

No caller may omit a foreign/default/Word binding, reorder or deduplicate more entries, change decoder choice or delay actual attribute validation under this repair. Parent scope borrowing from28 and decoder-Cow retention from26 are separate mechanisms. Their before/after counts must remain meaningful rather than being relabeled as this helper's allocation counts. Properties::word_prefixes_at has a distinct scope policy and is not part of this change.

## Standard-library evidence and capacity tradeoff

Read the installed stable toolchain's library sources under `/Users/atulsharma/.rustup/toolchains/stable-aarch64-apple-darwin/lib/rustlib/src/rust/library/`. alloc/src/fmt.rs:651 uses Arguments::estimated_capacity and write_fmt, with an as_str shortcut. core/src/fmt/mod.rs:754 describes that estimate as neither a lower nor upper bound. Display for str at core/src/fmt/mod.rs:2965 delegates to pad, whose no-width/no-precision path at :1910 directly writes the string slice. This supports the exact byte comparison above.

alloc/src/string.rs:493 implements with_capacity via Vec::with_capacity, :1104 appends string bytes, and :1432 appends a char after reserve. For a checked successful total, the builder requests enough capacity for the entire result before the four appends. No subsequent logical String capacity growth is required to fit those bytes. This statement is about capacity sufficiency, not a measured allocation count or an emitted-code claim.

These local sources are not proof of the exact optimized shipping call sequence. Compiler specialization, inlining, allocator size classes and actual call values remain unobserved here. Do not claim that the old format performs two allocations, reallocates for each operand, starts at a particular capacity in the final artifact or copies all payload bytes repeatedly. A fmt function appearing in samples is not that proof.

The returned encoded String remains owned and can live in inherited scopes, projection records or capture work. Its requested capacity strategy changes. Observe actual old/new lengths and capacities, aggregate retained scope capacity and release lifetimes. Do not shrink later, add a retained cache or claim universal RSS/peak improvement. The new branch's capacity guarantee does not remove the caller's mandatory Vec, URI, inherited-scope or bare-alias storage. ExistingQ22 and final equivalence/calibration still apply, without a new omission endpoint or ceiling.

## Actual old exposure and required site proof

The old28 observation has actual completed semantic/reopen results at all six declared warm states. Its parent hook increments after namespace_binding returns and before the Word-alias branch. Document picture25/100 records7837/15956 completed encoded-binding pushes and Cell picture7963/16082. URI decoder attempts/successes at that same parent match those counts. The real hook is in `.claude/scratch/S96-final-evidence/f268-foreign-prefix-storage-implementation-20261108-1/build_mechanism.py:17`, with actual observations in old-diagnostic/observation/OBSERVATION.json. The old build receipt records child0.

Those numbers are parent-site completed calls, not a namespace_binding entry/completion census across all five production sites, not allocation totals and not observed fmt growth. They establish exposure worth examining. They cannot authenticate this candidate's claimed copy reduction by themselves.

Before adoption, a source-bound old diagnostic must observe the actual format call inside namespace_binding. Count entry and completed result separately, with exact input lengths, full encoded result, result capacity and caller/source context where already available. The successor must observe actual checked-builder or original-format fallback execution and its completed bytes, initial/final capacities and existing scope consumer. Preserve counter meanings. A completed builder count is not an allocator counter. If allocation/reallocation claims are wanted, observe the actual allocator/code path separately and disclose diagnostic overhead, rather than inferring hidden fmt operations from payload sizes.

Compare complete ordered scopes and decoder/Source outcomes before mechanism assertions. Cover empty/default/Word/foreign pairs, Unicode and long prefixes/URIs, extra NUL in either input, XML-sensitive bytes and exact old split results. Use moderate safe strings, not forged huge slices or attempted address-space-sized allocations, to test semantics. Checked-overflow/isize fallback can be reviewed directly without a new arithmetic helper or unsafe test fixture.

Keep the existing full scope value/Cow/error-order matrix at `crates/rdocx-oxml/src/numbering.rs:7301`, malformed UTF8-before-URI controls and actual drawing/typed-numbering consumers. Preserve complete public immutable inputs and original setup/prior actions, full results/ZIPs/every member, owner/staging/source observations, IDs/relationships, held/fresh/refusal and saved/reopened semantics. Retain actual attribute/event/capture/namespace/resource reads, private Field and unrelated malformed-source controls. No new Source allowance, cache proof or broadened finite language is selected.

## Handback

Zero new defects, zero smells and zero nitpicks are established. The checked direct builder is a bounded candidate with an exact value argument and pending helper-site/capacity proof. It is not approved for adoption by this report. Root remains sole product/Cargo writer, and any composition needs explicit plan/HLD adoption, focused source-bound proof, ordinary installed checks and the next independent implementation review before another ordinary qualified artifact.

Only this report was written. No source/plan/test/progress edit, Cargo/build/test, product runtime, timing or network/Git mutation occurred. Existing actual76/79flat failures for reply and both picture routes remain retained. All final memory/equivalence/calibration, Office, installed typing, hosted/archive and integrated-union obligations remain pending. Control returns to Root.
