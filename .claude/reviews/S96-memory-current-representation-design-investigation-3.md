# S96 current memory representation, design investigation 3

## Disposition and boundary

The finite text, moved Finish children and ephemeral paragraph-loan changes fit
the existing A7/Q18/Q22 memory policy. They change the concrete representation
and lifetime that final equivalence, calibration, activation and RSS must cover.
They introduce no new ceiling, omission endpoint or permission to weaken Source
budgets. This report closes that applicability argument only. It does not
implement or approve an omission overlay, qualify memory, or predict elapsed
acceptance. Zero product defects, zero smells and zero nitpicks established.

Read memory investigations1/2, design investigations18/19/20, the existing plan
and the actual18/19 MECHANISM_COMPARISON. Root was temporarily instrumenting live
source, so source inspection used the frozen five-file tree under
`.claude/scratch/S96-final-evidence/f268-text-and-model-stack-implementation-20261108-1/loan-value-diagnostic/before/`.
Its Native document SHA256 is
`9cbb239777eba5b9311838c3715d6fa8f9baa78713a880e31b74b1dd77f96e60`,
OXML document is
`cbd7ab3ab041f22a57b9cac27bb629b6eff28e3e64462cc9b5d9cad51ebbef14`,
and comments is
`08d9f4f8a0160d87d96b809f8f6777fc1c522fb40374a30b2a3d466702a0ab13`.
All source line references below are relative to that frozen tree, not a claim
about instrumentation currently in live files. No compiler, product import,
runtime, test, measurement, plan/source edit or external action ran. Only this
review was written. Root retains the sole source/Cargo lease.

## Finite text materialization

The frozen Native qualifier at `crates/rdocx/src/document.rs:35464` still applies
the unchanged finite predicate before inspecting Text lengths. Checked byte
accumulation and the isize bound select one preallocated String. Direct pushes
copy the same stored Text values in run/content order. A non-Text or unavailable
length retains the original accepted_text materializer, and accepted_run_paths
is unchanged at `:35489`. No metadata length becomes a source-read receipt.
Investigation18's exact domain and fallback argument remain at
`.claude/reviews/S96-F268-finite-opening-fact-materialization-design-investigation-18.md:29`.

The final text is still an owned returned fact. Its capacity and lifetime must
be followed through the query result into the existing binding facts, not
called an eliminated allocation merely because per-run intermediate Strings
disappear. Native Normal and omissions use the same final materializer when
the existing proof permits it. If an omitted opening/model owner makes that
query unavailable, execute its actual fallback instead of installing another
classifier or retained text cache to recover eligibility.

The actual18/19 comparison reports6001 removed temporary run Strings and137345
removed intermediate push bytes for each1000-paragraph cold edit, with equal
complete ZIPs, parts, staging and normalized resource traversal. See
`MECHANISM_COMPARISON.json:114` and `:158` in the implementation packet above.
These are observed source operations and copied payload, not optimized malloc
totals, retained bytes, peak memory or RSS savings. No new retained omission
mode follows. Final calibration must include the current final String capacity,
ordinary path storage, temporary lengths and their real drop boundaries.

## Moved Finish children and higher observed capacity

At frozen `crates/rdocx-oxml/src/document.rs:1552`, the guard moves the actual
children Vec into the existing Finish step and copies the same borrowed owners
into queued Enter steps. Finish still consumes that Vec at `:1577`. The
summaries map, pending queue and child storage remain operation-local and drop
on success or refusal. No model child, admission or Source work is omitted.

The source Vec can have spare capacity. Retaining it until Finish therefore has
a different capacity schedule from the old length-sized cloned Vec. Actual
MECHANISM_COMPARISON SHA256 is
`8a62b0c1c7b3d608607ae510b331953a525a36c8dd53889b91221850e2a30fdf`.
All eight recorded states have higher new peak live Finish capacity slots.
Representative large states are:

| State | Old peak slots | New peak slots | Old Finish capacity sum | New sum |
|---|---:|---:|---:|---:|
| reply100 | 147 | 272 | 62372 | 119476 |
| picture100 | 149 | 272 | 154668 | 290040 |
| insert1000 | 1008 | 1032 | 480199 | 621876 |
| remove1000 | 1007 | 1032 | 479890 | 621440 |

These values are at comparison lines26,70,114and158. Capacity sum is cumulative
over observed Finish occurrences, not simultaneous live bytes. Peak slots
cover Finish child capacity, not the entire queue, map, model, allocator or
process. Removed duplicate-vector copies and shallow-copy bytes are separate
metrics. None can be subtracted from RSS or used to assert universally lower
peak memory. Preserve the actual higher capacities rather than shrinking a
vector to manufacture a more favorable comparison.

All final-source modes retain this guard representation and every mandatory
traversal. It creates no retained-cache omission endpoint. Ordinary-versus-
instrumented Normal calibration and per-mode current/peak observations must
include its temporary capacity and drop schedule. Existing guard/resource
equivalence remains mandatory even when a fallback enumerates more real work.

## Owned and model-loaned paragraph values

The frozen Native owner/storage at `crates/rdocx/src/document.rs:9202` has inline
Owned(CT_P) and Model(root, selector) cases. The exact existing full-key/live
Weak-root hit creates Model outside the cache lock at `:9409`. Every miss still
executes the actual fixed reader, retains all qualification/error behavior and
moves that parsed value into Owned at `:9521`. Frozen comments stores this
owner in its actual paragraph tuple at `crates/rdocx/src/comments.rs:3073`.

A loan adds one strong Arc edge to an existing immutable root. A cell selector
keeps the entire table allocation alive, including unrelated/private siblings.
Several loans may share one physical root. Count the payload once and ownership
edges separately. The loan owns no BodyBlockScan, writer vector or certificate
cell, and its immutable resolver must not clone, compare or traverse table
siblings. Existing finite selection and actual first reader remain decisive.
See `.claude/reviews/S96-F268-standalone-value-loan-design-investigation-20.md:49`
and `:61`.

Source Document ownership makes these roots ordinarily live during projection,
but the new edge can delay release on other controlled lifetimes. The reply
consumer explicitly drops the read at frozen comments`:2034` before publication.
Snapshot construction/projection at `:3527` retains its own source through the
read. No public comment result or document cache may retain the loan. A cache
Weak can legitimately upgrade while a held loan owns that exact snapshot.
Expiry checks must drop the last loan before requiring root payload expiry.

The inline enum pays its full actual size/alignment in every tuple, including
Model hits and omission-mode Owned values. Keeping Owned inline avoids a new
wrapper allocation but does not establish smaller tuple storage. Record exact
CT_P, value, storage enum and tuple sizes/alignments on the authenticated build,
plus Vec length/capacity and live selected roots. The actual debug receipt
`LOAN_MECHANISM_COMPARISON.json:15` in the implementation packet reports CT_P,
CommentParagraphValue and storage each2448bytes/alignment8, and the comment
tuple2624bytes/alignment8, for both the owned predecessor and loan successor
in all eight states. It also records equal complete output and packed graph/
capacity traces. This is observed layout equality on that compiler/build,
not universal ABI equality or a layout improvement. The older56-byte
certificate,48-byte key enum and64-byte hint observations are different types.
These recorded debug sizes do not establish heap, RSS or optimized stack cost.
Final ordinary/private calibration and representation applicability remain due.

## Exact StandaloneValue omission amendment

Retain report2's complete25-Native registry, group implications and BindingPieces
mode. Extend only its existing StandaloneValue description for the actual
table-key and loan representation already adopted in the plan at1106and1165.
The omission must suppress optional key Vecs/maps/capacities, Weak roots,
cell/lock/control allocations, installs, lazy seeds and clone/fresh-context
carry before their construction. It must also suppress returned Model loans
and optional advisory candidate vectors, retained block/root candidates, table
hints/selectors and associated payload construction. Empty inline containers
may remain, with their actual layout disclosed.

This explicitly updates report2's StandaloneValue row at
`.claude/reviews/S96-memory-omission-endpoint-design-investigation-2.md:53`, whose
earlier sentence retained call-local advisory candidates. A complete omission
must not keep those optional root-retaining candidate payloads merely because
hits are disabled. Preserve actual marker/location/namespace preparation,
constructor/reader and all other mandatory Source and publication work. Each
read must return the genuine parsed Owned result. Do not replace it with a
cloned candidate, empty success, projected leaf or deferred loan. No extra
standalone predicate or shadow AST recovers an omitted shortcut.

WriterModels still implies ChildRead and StandaloneValue, including suppression
of every retained alias/carry owner, as report2 specifies at `:39` and `:148`.
Removing the candidate or loan edge alone does not remove a root independently
held by a Normal writer/opening/current model. NativeAll and PythonAll retain
their exact declared closures at report2`:77`. There is no new mode count or
claimed isolated whole-table saving.

## Precise existing-plan amendment before measurement

The obsolete baseline wording is the two sentences at
`.claude/plans/F-X268-design.md:452` beginning "Pair a pre-cache commit baseline"
and "Record the baseline source". Replace those sentences explicitly, without
changing the phase/sample protocol, with this source-equivalence contract:

> Compare Normal with true per-layer omission variants from the same frozen final
> source and mandatory representation, and execute the complete NativeAll and
> PythonAll scopes directly. Record the exact resolved endpoint manifest,
> source, input, build/artifact, profile, allocator and fresh-process identity
> for every mode. Establish complete semantic, Source/error and ownership
> equivalence before interpreting RSS. These are current-source pre-layer
> variants, not a recovered historical pre-cache commit or historical raw run.

At line454, keep "Use true equivalent pre-layer and final variants" and all
per-layer/combined obligations. Replace "Retain the original fixture,
statistics, raw RSS" with explicit custody wording that preserves all surviving
historical records as dated history, discloses missing original raw/helper/input
custody, and pins the openly reconstructed fixed20,000-by-six fixture and every
new generator/input/main-XML identity, raw observation, statistic and peak.
The actual current fixture size must be reported. Matching an old description
must never be labelled historical byte identity or raw-data recovery.

Those edits formalize report2's argument at `:136`. They do not relax A7, Q18
or Q22, remove an endpoint, adopt the donor6/68MiB claims, or introduce an RSS
pass ceiling. Keep the ordinary final-release versus private Normal calibration
as its separate actual comparison. Do not disable the new text/Finish/inline
representations to manufacture an otherwise-equal historical binary.

## Unchanged closure obligations

Report2 remains operative for all borrowed getters and actual owned/Cow
fallbacks, especially unsavable-body diagnostics, and for BindingPieces active
publication/held-handle behavior. No omitted getter may return None, leak text
or create another retained cache in place of its actual result. Its exhaustive
31-field registry comparison still ignores only optional scanned_parts and
checks all30 mandatory identifier/provenance fields, including nested authored
drawing IDs. Live package/source metadata, numbering, owner states, indexes,
scope generations, trusted provenance and signature invalidation remain.

Retain complete output/part/model/reopen equivalence, real setters, typed first
errors, private Field behavior, callback/atomic refusal and source budgets for
every mode. More fallback work cannot mint another allowance or Source credit.
A changed resource verdict fails equivalence. No font-only or entry counter
substitutes for successful activation of every lazy endpoint.

Keep separate ordinary/private-Normal Native and installed-Python calibrations,
three fresh alternating pairs per deciding category, actual current RSS and
every peak/raw sample, exact phase/buffer lifetimes and a directly executed
combined scope. Activated N/4N companions require equal terminal mandatory
state, held root/loan controls, source/root replacement, partial collection
failure, clone, in-memory and ZIP reopen, explicit drop order and final last-loan
release. Weak observers stay out of deciding RSS and disclose their COW effect.
No graph traversal or correctness warming enters measured phases.

Actual overlay implementation, per-mode equivalence/activation, final-source loan
lifetime and representation applicability, final calibration/RSS, timing readiness, Office,
installed typing, hosted/archive and final full union remain pending. This
report changes no plan, product, gate status or publication authority. End the
bounded pass and return control to Root.
