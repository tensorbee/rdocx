# F-X272 core capture and raw textbox, Root bounded final pass 2

Reviewed the five-path repair against the immutable pass1 snapshot and the
same approved contract. Manifest SHA256 is
`5df8176a934c0af2a6250c6f1b6480ec49bc68fcfae77dc2011487f88e873bb1`.
Patch SHA256 is
`6f99f28ae6b220ccd1e764dcdf997c0009eed43117d118afb1ed39317fde9f9d`.
Final placeholder SHA256 is
`9f01b168d1ec71923193df947e0216c268f33165d91e0bd837939b5067282922`.
Citations refer to the frozen repair source. Verdict: zero defects, zero
smells, zero nitpicks. D1 is closed locally, not the whole F272 provider.

## Defects

None. At `crates/rdocx-oxml/src/placeholder.rs:3254`, rewrite shares the existing
owned-first aggregation after selecting only the actual walk's Word inputs.
The private moved-slot transcript records the real stretch, run, content slot
and original length. At `crates/rdocx-oxml/src/placeholder.rs:3273`, occurrence
and property ranges are checked against the full aggregated stretches before
restoration. The original String buffers then return to their unchanged slots
before model admission, dependency validation, candidate cloning and editing.
A failed append also restores those moved buffers before propagating the
sticky refusal.

The first nonempty uniquely owned Text moves once per stretch. Appended tails
remain in their original model slots and pay their actual copy demand. Truncate
uses the moved Text's original byte boundary, then moves its buffer back.
Producer-fixed walk paths and visit counts prevent reassociation or moving a
multiply visited slot. No text equality search, extra parse, new input grant,
Arc copy-on-write or no-match validation bypass is introduced.

## Smells

None. The concrete slot record serves actual borrowed editing and restoration.
The existing path resolver now returns the mutable String so both consumers
can take or restore it. Revision traversal at
`crates/rdocx-oxml/src/revision.rs:652` retains its insertion/move-to eligibility
and Arc::get_mut uniqueness requirement. Shared owners do not gain a free copy.
The private edit tuple and extracted validator serve actual ordinary and raw
template callers. No new public API, module, trait or quota is added.

## Nitpicks

None.

## Not found

No additional correctness, contract, panic, OOXML, test or structure finding
in this repair. Existing core transport, recovered-run alignment, glossary
handback and positive publication lifetime materialization remain as reviewed
in pass1. The repair does not change the core capture producer or claim that
unknown publication regions are ordinary.

## Actual evidence and source boundaries

Mechanical authentication SHA256 is
`a9a622ba2e5f40ec07e4da1793a78ff3469d902aa09488191a7bb5a30e7d7e55`.
It authenticates the exact five-path delta, preceding Base, final snapshots,
retained RED source and executable, fixed executable and named raw results.
The actual public generated PAGE carrier, composed textbox and discovery have
zero allowance and zero copy demand. Before repair, legacy no-match returns0
while scoped rewrite refuses. After repair, the same no-match succeeds with
unchanged legacy bytes and zero new input or copies. A separate all-generated
two-piece stretch still refuses the real tail append in both discovery and
rewrite. Matching, joined-range, no-match and invalid-value controls exercise
restoration before the original edit kernel.

Actual11raw,7affected row and14disjoint shared controls pass,32unique, alongside
eight scoped gates and three post-documentation metadata checks. The retained
PRE_GATE file is the first attempt's pin map, not a final before map. Root read
the actual runner: it uses that file only for the path list, computes a fresh
before map on each invocation, checks it after every command and writes the
final map after the successful invocation. The mechanical audit cannot itself
recover each unstored per-row map, so no stronger per-row archival claim is
made. Final product pins match the repair snapshots. HLD03 changed afterward
and has its separate metadata observations, not a fictitious identical gate
pin. Earlier lint failure remains retained.

The first zero-selected filter is not RED. The ordinary same-run multi-Text
reference hit an existing subtraction panic before the scoped call. That
fixture is retained separately, and the supported two-run join supplies the
actual parity control. No unrelated arithmetic fix or new acceptance claim
is folded into this repair.

Historical allocation stdout, driver and binary do not repair the missing
archived dependency identity. The expected earlier rlib was not found among
eight retained candidates. That custody remains unproved and is not final
trim-source allocation qualification. Final source inspection confirms the
disabled transcript path, without inventing a new measurement.

Stop confirmation of this bounded increment. Remaining repeat/import methods,
Native production wiring, complete versus unknown Field classification,
deferred raw/model roles, nested requested Text correspondence, current and
constructor pre-growth, child serial integration and all final integrated
schema/native/installed/hash/performance/provider gates remain open.
