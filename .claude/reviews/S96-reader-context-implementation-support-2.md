# S96 reader context, implementation support pass 2

## Reviewed scope and verdict

Independent read-only microscope of the finite reader-context increment in the existing S96 integrated final-union remediation. This is implementation support pass2. Implementation support pass1 and design investigations1/2/3 remain consumed and retained. Historical F268 completion, current plan/run status and review caps are not reset. Root remains the sole product and Cargo writer.

Reviewed branch sprint/s96 at HEAD cec87f9501eb13d0d5d60572472b227483262620, with the frozen existing dirty work. The retained six-file increment contains247 added and6 removed lines: Native document source, Word binding source, existing Python test carrier, F268 plan and HLD03/10. Existing F273, scope-allocation, Native Cell/reply and other dirty work is protected, not reimplemented or requalified here.

**Verdict: zero defects, zero smells, zero nitpicks.** The implementation conforms to the adopted finite contract at `.claude/plans/F-X268-design.md:845`. This is correctness and mechanism acceptance only. No performance pass or final-union acceptance is claimed. No Cargo, tests, product imports or benchmarks were run by this reviewer.

## Source and artifact authentication

The evidence packet is `.claude/scratch/S96-final-evidence/f268-reader-context-implementation-20261107-1`. Independently authenticated SOURCE_BOUNDARY.json has SHA-2565891d4da6d855abd33fe58396e88eac609b403e09c397825432421561bb34877. Its14 evidence pins and six current source pins match the actual files. All six BEFORE.json pins match their retained before copies. Regenerating a read-only textual comparison from those copies exactly matches INCREMENT.patch, SHA-25630c500306e0a886ff09ae033d9c69aa3e8bafcfce2d01740da702de2d8408890.

The successor wheel has SHA-25649d266bfed0c486081ce2e645ff302998a6db5133b5725a12e74ce3c33f6ffe7. All ten wheel package payload members match the actual installed files. The installed successor Native module is17eec823f6682911f04576f7b59113eed5bd35b878ae192b888e108edde0d434. The predecessor receipt identifies the retained scope-allocation successor module57df7191ca3701f7019f90b3c22dec3cd1b440e4e5765668c43a291a3a37d3f2. Both final public executions pin the same current test source,5b6f2a8b5b4e4fedefc155e0a110e36f2610c80e2401fd316d4e212f1b1b5712.

The complete build input manifest differs from the final tree only in the explicitly retained Python alias expectation correction and plan clarification. Runtime Rust and HLD inputs match the actual build. The public runner asserts the intended module directory and actual debug producer probe before invoking the tests. Build and corrected installation receipts report actual child exit0.

The separate context diagnostic's retained before/diagnostic source hashes match SOURCE.json. Its temporary change is only a debug observation at the changed-context branch. Its before binding source exactly matches this increment's retained before binding source, confirming restoration before implementation. EXACT_CONTEXT_DELTAS.json authenticates the actual raw stderr SHA and records both original25/100 fixtures: same part and flag, exact collections0/1/3, and unique unchanged pairs reordered only in2/4. This closes the factual input left pending by investigation3. It supplies no shipping timing result.

## Predicate and reader proof

At `crates/rdocx/src/document.rs:1880`, derived PartialEq/Eq remain unchanged. The substantive method at1898 immediately preserves old exact-equality behavior. Otherwise it requires exact part, eligibility flag and collections0/1/3 at1902. It checks unique lexical names in every collection on both sides at1910. For collections2/4, equal length, complete pair membership and unique names establish symmetric permutation equality at1921. Only declaration-shaped names qualify, including the default declaration. Names and values remain exact and case-sensitive. The method neither sorts nor mutates stored context, creates a receipt, grants resource credit, nor interprets a hash as identity.

Preserving the exact-equality branch means an unchanged context containing duplicates retains its previous behavior. Duplicates cannot enable the new order-only branch. This is the explicit approved distinction, covered by the native matrix at83594. Changed membership, default URI, prefix spelling, value, MC/root attributes, body declarations, part or blocked flag cannot qualify through the new branch.

The reader-specific proof is supported by the actual producers. `crates/rdocx-oxml/src/document.rs:14096` serializes a standalone block using root_start and root_binding_scope. At14119, those depend on the exact model extra_namespaces and root_attributes, which remain exact collections0/1. They do not depend on the two reordered retained collections. `crates/rdocx/src/document.rs:34465` closes the actual standalone source. Its closure at34483 obtains the scope from those bytes and the exact model namespaces, including the existing w14 rule. Equal eligible blocks therefore retain the same standalone bytes and parsed read facts under the newly admitted retained-declaration permutation.

The native payload comparison at `crates/rdocx/src/document.rs:83626` invokes the real body_item_owners producer on both contexts and compares snapshots, locations, XML, cells, paragraphs, links, owner counts, routes, tables and field status. Its existing varied-body input at82429 includes formatted and linked paragraphs, simple/complex Fields, bookmarks, controls, revisions, tabs/breaks and nested/spanned tables. This tests actual produced payloads rather than only a boolean marker.

## Consumer, provenance and refusal audit

The only product call of read_facts_compatible is `crates/rdocx-py/src/document.rs:5594`. It controls old body-piece eligibility, related-part eligibility and candidate inventory admission within the existing assemble body. Exact block alignment and qualified_block remain required at5543 and5551. Field/private-value exclusions remain in that qualification, and unqualified blocks still execute word_block_owners. Unsupported facts still make inventory unavailable.

Every assemble call first extracts the current related sources at5568 and rejects duplicate case-equivalent physical names at5573. Reusing a related part still requires an existing successful owner payload, exact current kind, part name and XML at5673. Inventory still validates current hyperlink relationships at5755 and5848, and reconstructs placement, parents, routes and projections from the current piece order. This predicate does not supply any Source or writer authority.

Incompatible existing contexts still force inventory=None at5700 and then the existing complete candidate reader at11249. The debug successor still compares complete piece inventory and projections with an independent guarded full rebuild at5919 and11246. Its body fingerprint exclusion is pre-existing and is not broadened by this diff.

The non-form publishers continue through assemble at11233. FormValues, ModernFormValues and Signing retain the explicit full path at11224. Local cell publication still requires exact context Eq at11658. Native opening/source/writer equality is untouched, including the retained namespace checks at `crates/rdocx/src/document.rs:25072` and26584. No scope preparation, generation budget, retirement, serial, staging clone, resource allowance or C1 enforcement changes.

The additive method is substantive and belongs to the existing doc-hidden concrete native context. It introduces no forwarding wrapper, trait, generic, source file, cache, feature or dependency. The hidden Rust API addition is explicitly adopted at `.claude/plans/F-X268-design.md:849`. Python/WASM/CLI exports and stubs are unchanged. HLD03/10 describe the narrow read-facts exception while retaining exact source/writer identity.

## Discriminators and retained outcomes

The installed public regression at `crates/rdocx-py/tests/test_core.py:15748` uses a byte-opened mixed body with table, picture and header neighbors. Parent acquisition activates authority before the measured reply window. The test completes reply parent/text/anchor, held stale operation label, fresh neighbor values and save/reopen assertions before checking the real full-inventory count at15802.

The final predecessor ordinary case reaches that last assertion and fails with actual count1 versus0. Its alias case passes with count1. The corrected successor executes both cases and reports2PASS,0FAIL,0SKIP, child0 in4.63seconds. This is an actual successful predecessor semantic execution followed by mechanism RED, not a missing symbol or compilation failure. Independent comparison confirms exact ZIP bytes, identical member order and all14 part bytes between predecessor and successor for each case.

The alias fixture adds xmlns:w during fixed-prefix writer staging alongside retained xmlns:q. The diagnostic records changed binding membership, not just order. The final test at15800 consequently requires full fallback1. The retained initial expectation failure is a premise error, not evidence that compatibility should be broadened. The earlier missing add_run argument and failed uv install, whose invalid older artifact produced2SKIP, are retained separately and provide no RED/GREEN acceptance credit.

Native logs report the two new inline cases passing. The native matrix at83534 checks symmetry, unchanged Eq, empty/one/multiple collections, each collection's name/value/count/case changes, duplicate names on either side, and changed part/flag. Scoped Native/binding all-target checking, strict all-target/all-feature Clippy, format, prose, adapter drift and diff checks report actual exit0.

The final additional control receipt pins the authenticated successor module and current test source. Its XML and log agree on61PASS,0FAIL,0SKIP, child0 in273.47seconds. The reviewed assertions include physical duplicate-part fallback at `crates/rdocx-py/tests/test_core.py:15462`, namespace and reply refusal with held-view/byte preservation at16965 and17034, formatted alias/raw reply controls, Field namespace lifetimes at3950, live tab/break text, run/cell/header fact updates at14449, local scopes, conservative retirement, splice rollback and the prior five zero-owner-successor tests. These are targeted controls, not a claim that all final riders ran.

The additional original-fixture debug record uses the authenticated successor and pinned original factory. Both25/100 cases retain semantic/reopen success, report zero general rebuilds and14 piece reads, with88/254 alignments. Five ordinary writer outputs remain. These actual mechanism counts support the intended dispatch change but do not assign elapsed time or predict the flat-cost result.

## Findings and pending acceptance

Defects:0. Smells:0. Nitpicks:0. Correctness, contract, panic/index bounds, OOXML preservation, tests, structure, provenance and refusal review produced no concrete finding in this increment.

The latest ordinary release still has the three recorded timing misses. A fresh ordinary artifact and the unchanged original79flat,5linear and407 protocols remain required. All applicable final-union, native, typing, memory, hosted and archive riders retain their existing pending status. Dormant coalescing and mixed cold batching remain unapproved and are outside this increment. This report ends implementation support pass2 and returns control without remediation.
