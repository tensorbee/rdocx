# F-X283, C1 diagnostic rebase support, pass 1

**Reviewed**: Source preparation only. The six-file3336line FOUNDATION.patch in
`/private/tmp/rdocx-s96-diagnostic-foundation-c1-rebase-20261106` and the one-file
995line FOUR_OVERLAY.patch in
`/private/tmp/rdocx-s96-four-overlay-c1-rebase-20261106/resolved-four-v2`.
This is bounded support for the existing accepted diagnostics, not another
feature peer cycle, final-H qualification or product acceptance.

**Verdict**: 1 defect, 0 smells, 0 nitpicks.

## Defects

### D1, whole-main omission loses the actual scoped writer

`/private/tmp/rdocx-s96-diagnostic-foundation-c1-rebase-20261106/after-document.rs:28134`
`/private/tmp/rdocx-s96-diagnostic-foundation-c1-rebase-20261106/after-document.rs:28135`
`/private/tmp/rdocx-s96-diagnostic-foundation-c1-rebase-20261106/after-document.rs:10645`
`/private/tmp/rdocx-s96-four-overlay-c1-rebase-20261106/resolved-four-v2/after-document.rs:28317`

The omit9 early branch in serialized_main_part_in calls MainPartBlocks::serialize,
whose compatibility entry always forwards template=None. A live
Some(TemplateMainWriter) is therefore discarded when block_frame_family or all
selects this branch. It uses the ordinary fresh-work serializer and returns
without supplying the current operation's main Source. This loses the actual
C1 Field/Word receipt and Source handback, rather than merely omitting the
selected cache family. A previously stored Source may also remain in the writer
state because this branch never replaces it.

The lower adaptation already provides the correct substantive route:
`/private/tmp/rdocx-s96-diagnostic-foundation-c1-rebase-20261106/after-document.rs:10610`
retains the scoped writer in serialize_whole, and
`/private/tmp/rdocx-s96-diagnostic-foundation-c1-rebase-20261106/after-document.rs:10654`
forwards it from serialize_in. The early whole-main caller must carry that same
actual optional context into this route. Do not create a new operation or infer
Source authority from output equality. This is a source-derived defect. No
runtime failure or compiler result was produced during this review.

## Smells

Zero smells.

## Nitpicks

Zero nitpicks.

## Authentication and examined preservation

All six foundation before hashes equal the actual integrated C1 files sampled
in canonical5a4c5c0b. All six after hashes and the foundation patch hash match
FOUNDATION_PREPARATION.json. Independent in-memory exact-context application
reconstructs every after file. The Native foundation is
`cf6d4b1280bef37941192a87320cd259b3f7c41b41ba2e6f650e79454920b90d`.

All21 accepted input source hashes authenticate. The five pinned four-overlay
Root review hashes authenticate. All six accepted COW, text, font and recorder
source fragments occur byte-exact in the new composition. Independent patch
application reconstructs the four-overlay Native source
`a70fec092ae0f512949686d182a919e3e003660603eace4597d71dc8fc240acf`
from the supplied foundation without fuzzy matching or filesystem application.
The single adapted COW hook remains after the actual ordinary record edit and
before its existing qualification. It does not replace the C1 scoped branch.

Examined the seven foundation conflict resolutions, scoped whole-fallback
handback, preserved seventh loader argument and authoritative scoped-factory
decline, optional package reread, ordinary cache-install counterparts, runtime
policy exports, counter registration and four-overlay hook placement. Other
than D1, no additional concrete correctness, XML-policy, source-custody or
structural rebase defect was established in that bounded examination. Earlier
accepted diagnostic semantics are reused, not newly qualified by this review.

## Remaining evidence boundary

Neither packet was compiled or executed here. Source preparation, manifest
checks and exact accepted fragments do not establish Normal-All parity,
observer neutrality, natural COW activation, memory/timing results, font/history
completeness or final-H applicability. B is still a future source operand.
Final source rebasing, compiler/link custody, actual calibration and every
previously required semantic, history and resource observation remain pending.
No product, preparation, controller, plan or progress file was changed. Only
this review record was written. The pass ends before remediation.
