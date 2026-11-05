# F-X272, comments composition, Codex pass 1

**Reviewed**: Sealed C2/C1 reconciliation, seven paths, 690 additions and
17 deletions in the complete 830-line patch. Lower and Native comments,
Python tests, HLD03/04/10 and the existing progress record.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Bounded composition is clean.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Source and contract review

Reviewed the approved F-X279c split and implementation contract, its existing
independent implementation review, C1 source obligations and the complete sealed
reconciliation delta. This review concerns their actual composition, not whole
C1 or final native acceptance.

The lower patch preserves the C1 admission entry at
`crates/rdocx-oxml/src/comments.rs:49`, including admit_reader_source,
controls_for and the actual ControlSourceCursor passed through parse_comment.
Authenticated the parser entry and intervening parser/helper region as byte
identical before and after composition. The paragraph parser still receives
that cursor at `crates/rdocx-oxml/src/comments.rs:222`. No ordinary parser,
Field source or constructor policy was substituted from the child baseline.

The C2 guard at `crates/rdocx-oxml/src/comments.rs:120` inspects original root,
comment and retained paragraph-root declarations before fixed-name emission.
Foreign canonical w/w14 now refuses instead of being rebound. The guard does
not scan arbitrary opaque descendants or create import-read refusal. Compatible
aliases and genuinely unused declarations dropped by the existing paragraph
parser retain their reviewed interpretation.

The existing emitted root-binding scope at
`crates/rdocx-oxml/src/comments.rs:180` is unchanged and follows the actual
preserved root attributes. It remains live across all root raw children,
comment/paragraph writes, final root close and late w14 declaration handling.
Inspected the existing scope helper and its replacement-root semantics. The
new original-binding guard and unchanged lexical scope establish compatible
roles, rather than restoring forced foreign canonical binding. No new scope,
resource exemption or input authority is introduced.

Native comments before composition exactly matches the child claim baseline.
The sealed successor is byte identical to reviewed canonical94775888 Native
comments, SHA256
`e355e4b8971f94f7dfe88487bb2bfd7a12066a928d0d19c1a6f884d3a4bddcac`.
Rechecked the shared add/reply authoring helper, style type prechecks, existing
style-builder use, MC URI/local-name lookup, fresh-prefix selection and staged
publication paths. The lower reconciliation preserves those paths' actual
namespace refusal and rollback. Paragraph styles, first-paragraph annotation
placement, style order/base handling, optional commentsIds and untouched imports
retain the reviewed C2 behavior.

The C1 Python test source is preserved as an exact prefix. Exactly 107 reviewed
C2 lines are appended, and that suffix matches canonical94775888. The old
issue435 append context was absent in the C1 carry, so the retained git apply
check failure occurred before Python writing or Cargo. It is setup evidence,
not a product RED. HLD03/04/10 add the reviewed comment paragraphs without
removing the existing C1 sections.

## Authentication and evidence

Authenticated BOUNDARY.json SHA256
`789073c6846967882a84bb285bc7268257fcdadf424b3f9681a7285a84a06120`
and increment.patch SHA256
`af0aeccbaa14cb96f717f3080ac5e24b96599fdf245e3e0249e176f76b27c7ec`
under `/private/tmp/rdocx-s96-root-c1-implementation/comments-c2-reconciliation`.
All 14 before/after snapshot hashes match. Regenerated the entire unified patch
in memory and matched its exact bytes. All nine sealed log hashes match.

Read the actual eight lower comment PASS names, five Native issue434 PASS names
and one Native all-entry-point PASS. Inspected their assertions in the sealed
source, including original binding refusal, alias/unused-declaration support,
retained producer identity, markup/styles, wrong-style atomicity, MC tokens,
untouched package bytes and optional commentsIds graph preservation. Scoped
lower/Native all-target checks and strict all-target/all-feature Clippy succeed,
with final Clippy completing in 24.93 seconds. Formatting, prose and adapter
logs are clean. No review builds or tests were run.

The appended Python source is reviewed, not a claimed runtime qualification of
this C1 composition. Child-worker extension runs remain evidence for their
recorded source/artifact boundary and are not relabeled as C1 runtime results.

## Not found and limits

No correctness, contract, panic, OOXML, testing or structural findings were found
in this bounded composition. No new public API, module, parser, quota or C1
funding rule is introduced. Existing historical native hash/version constants
are unchanged and are not claimed to pass fresh Word execution.

Fresh Python runtime on the composed source, exact-candidate Word execution,
remaining Native lineage/consumer work, later serial reconciliation and all
integrated final riders remain pending. This report makes no whole C1,
provider, feature-completion or sprint-completion claim.

Review ended after writing this report. Only this canonical review file was
written. No source, progress, plan or HLD changes and no remediation occurred.
