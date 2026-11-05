# F-X272, first main writer implementation, Codex pass 1

**Reviewed**: Sealed first-Native-writer and checked-copy increment. Seven
before/final paths, 1146 additions and 67 deletions in the full 1551-line patch.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Bounded increment is clean.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Contract and source review

Read repository workflow and microscope requirements, existing progress and
canonical75bfec84. Authenticated its approved plan SHA256
`6a7febb877eb10891a0324e6dc8493942570d55be888aa6ae1a52eb6f0dcce09`.
Reviewed the complete sealed delta and its actual existing model/writer callers.
Moving successor source was not treated as this review boundary.

The existing staging clone now retains its actual main-body reference at
`crates/rdocx/src/template.rs:440`. initial_writer_inputs at451 is used only by
the first substantive writer before preparation, and disables reuse after
preparation or slot retirement. Its physical current/frozen paths come from
that clone producer. The concrete graph follows existing paragraph, table,
row, cell and control edges. Direct control Runs are paired at their actual
content slot. Revision, Ruby and Drawing graphs decline the connection and
retain their legacy path. No equality lookup, candidate-address catalog or
post-preparation ordinal inference creates an origin.

MainPartBlocks::serialize_in at `crates/rdocx/src/document.rs:9674` calls the
existing source_model_body_item for eligible fresh child writes. The same
substantive model writer receives borrowed Text, Field and retained-raw pairs.
The separately owned block bytes are obtained through repeat_source and
finish_source, rather than an unchecked clone or duplicate serialization.
Assembly at `crates/rdocx/src/document.rs:9798` uses actual scoped composition
for head, selected child sources, unproved legacy child pieces and tail. Its
separate legacy output and canonical record copies consume the same operation's
applicable copy demand. Distinct buffers are not merged by shared provenance.
Inspected cache-frame refresh, reused-block exclusion and fallback recursion.
Unproved pieces remain unproved rather than becoming ordinary input.

Exact no-target namespace replay at
`crates/rdocx/src/document.rs:4151` retains a checked copy and moves the actual
prior Source into legacy output. The actual target-writing branch clears the
Source handback before its existing event writer runs. Conditional w14 insertion
also clears it when the old buffer is changed at
`crates/rdocx/src/document.rs:9842`. Thus those pending producers do not return
a posthoc receipt for an equivalent but separately produced buffer. Inspected
the assembly fallback conditions and source-reset boundaries.

template_programs_initial at `crates/rdocx/src/document.rs:45811` consumes this
actual first flush and uses its Source only when the connected producer returns
one. It preserves related-story ordering and Comment exclusion. All selected
initial physical-program raw copies pass through the scoped copy helper at
`crates/rdocx/src/document.rs:16467`. Table namespace closure retains the
existing checked table parse and row-count validation. This is physical program
inventory, not the later grammar TextWalk or second prepared writer. The real
render caller at `crates/rdocx/src/template.rs:1783` preserves that distinction.

## Lower checked-copy boundary

compose_sources at `crates/rdocx-oxml/src/document.rs:6139` prepares anchors,
ranges, checked lengths, maps and constructor admission before reserving the
actual output-coordinate Field union. It combines mapped publication facts
and delivered coverage without double-counting overlap. Two output occurrences
remain two copies even when they share origin. The existing materializer grows
bytes only after reservation. Imported rewrite keeps its separate reservation
and does not pass through this debit again.

close_source_namespaces at
`crates/rdocx-oxml/src/document.rs:6288` prepares the existing lexical insertion,
map offsets, event positions, coverage and admission before copied Field growth.
A genuine zero-insertion path moves the original Vec. Actual insertion pays
for known copied Field intervals, while its generated declarations acquire no
input or ordinary role. Current publication combines both fact stores. Sticky
refusal remains in force, including empty composition after refusal. Invalid
preparation does not consume the copy reservation. No metric, quota or public
signature changes.

## Evidence and authentication

Authenticated SOURCE_BOUNDARY.json SHA256
`221957beec7eacc703e0959691f65e21f8e2b917a50cac62543b2ae47b90b4e5`
and increment.patch SHA256
`c9531cb503eb1434f090ba20336af40cf83b0a5222393921b49341631a40d034`
under `/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/main-writer-implementation-handoff/implementation`.
All 14 source snapshots match their pins. Regenerated the entire unified patch
in memory and matched exact bytes. All eight boundary-listed artifact hashes
and both final executable hashes authenticate.

Native executable SHA256
`3439662e903fccf0a3087d28bbd95d0a04441c450d57eec8d8f5f82fa8c8ba51`
has 50 actual named template PASS results. Lower executable SHA256
`ccc051e11e7bc34b321887780d439d13eb23201c5a79f1687405e8c21ba627f4`
has six new transport PASS results and 78 distinct related PASS results. Checked
all seven related log hashes, their named counts and matching binary receipt.
Read the new assertions, including byte/program parity, exhausted shared Work,
no publication on failed main production, opaque/control preservation and the
explicit pending w14 branch. Lower controls exercise both fact stores, output
multiplicity, preparation refusal, unchanged ingress state, generated gaps,
genuine move and actual New/simple writer bounds.

Authenticated the pre-fix lower source, executable and raw log for the two
actual checked-copy failures, exit101. These are this increment's observed
pre-fix failures, not historical claimBase RED. No Native historical product
RED is asserted. The legacy replay fixture failure, nonexistent CT_R::default
compile attempt, missing xmlns:w fixture, obsolete sticky-compose expectation,
wrong-edition rustfmt and Clippy corrections remain separately classified in
the seal. The obsolete expectation's reconstructed source is explicitly
identified as reconstructed, not a newly discovered pre-run snapshot.

Final strict lower/Native all-target/all-feature Clippy has a captured exit0
and completes in 22.52 seconds. The earlier Finished-only log with a lost shell
exit is not substituted for it. Metadata receipt records formatting, prose,
adapter drift and diff checks at exit0. No review build or test was run.

## Not found and limits

No correctness, contract, panic, OOXML, testing or structural findings were found
in this bounded source. Existing compatibility entries remain ordinary callers
of substantive implementations. No new public type, trait, generic framework,
input allowance or transport exception is introduced.

Actual nonempty namespace target replay, conditional w14 producer transport,
second writer/parser, memo receipts, complex selected publication, private stored
graphs, related/comment writers and refresh lineage remain pending as sealed.
Whole model clone/cache-copy coverage, the canonical hyperlink-style neighbor,
Native import/grammar closure and final integrated riders are not qualified by
these tests. No whole C1, provider or sprint completion is declared.

Review ended after writing this canonical report. No product, progress, HLD or
plan was edited and no remediation occurred.
