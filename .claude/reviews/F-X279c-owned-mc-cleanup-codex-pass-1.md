# F-X279c, owned MC cleanup, Codex pass 1

**Reviewed**: Frozen C2-only Native comments cleanup repair. One source file,
93 additions and 4 deletions. No D authoring implementation is in this subset.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Bounded repair is clean.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Mechanism and contract

Read the approved in-progress C2 contract and the frozen repair source/patch,
existing cleanup caller, part retirement helper and original regression. This
is compatibility remediation of C2-authored MC metadata, with no new peer cap,
public API, type or namespace policy.

In the private candidate
`/private/tmp/rdocx-s96-f279d-evidence/C2-NEIGHBOR-source.rs:2074`, cleanup still
requires an empty comment definition list, no opaque root children and actual
library ownership. The new predicate at the same candidate's line2105 admits
namespace-only roots as before. Its only additional admitted non-namespace
attribute is one Ignorable attribute with a genuinely MC-bound prefix, the
single w14 token and a correct canonical w14 binding. The prefix is resolved
by URI/local name, so an existing genuine MC alias works. An opaque attribute,
extra MC token, foreign MC/W14 URI or second such attribute prevents retirement.
Unowned imported roots remain excluded by the unchanged ownership condition.

The existing remove_owned_part helper retains relationship, content-type,
part and authored-identifier retirement. The commentsExtended empty-part policy
is byte unchanged. No opaque children are parsed or rewritten. The repair
changes only eligibility for retiring an already empty owned comments part,
not comment serialization, import admission or reply topology.

Inspected the eight-mode added control at the private candidate's line4126.
It distinguishes the canonical authored root and genuine MC alias from foreign
w14, extra tokens, foreign MC, opaque attributes/children and imported ownership.
The original removing_the_last_owned_comment_retires_its_complete_identifier_bundle
control at line4522 is byte identical to Base. It exercises real add/remove,
complete part retirement, identifier reuse and exact history-versus-direct
saved output. Its assertions were not weakened to accommodate C2.

## Exact source and evidence boundary

Base Native comments at claim ff246c8b8b76ef573e8438b582259ba7045124ba has
SHA256 `e355e4b8971f94f7dfe88487bb2bfd7a12066a928d0d19c1a6f884d3a4bddcac`.
Frozen C2-NEIGHBOR-source.rs has SHA256
`4b45feb1bc158ac4f839c0a9186e7fed78b5801afcbaf06d67424354a6c8a9a7`.
C2-NEIGHBOR.patch has SHA256
`24987463db0b833f0a0d615373c5d1f69338bea2f0f3eecd46b077d85b0b211b`.
Regenerated the complete patch from the git-object Base and frozen candidate
in memory and matched exact bytes. Only the cleanup predicate/helper and its
added test are present. No D reply production change is accepted by this review.

Authenticated the retained baseline source and binary against baseline/RECEIPT.json.
The baseline binary's recorded production source is byte identical to claim
Base. The preexisting-c2-owned-part.log records the actual unchanged bundle
control failing before reply authoring, with zero passes and one failure.
Its SHA256 is
`4c23f7c42faf787632cde4db0382ed84705d7f26729c2e98bc3e6c771249556f`.
This establishes the C2 neighbor regression separately from D's new behavior.

The latest native-final.log has SHA256
`dcfe96d21732f591985a4b92fcd040c75b787da2f74f73590130daa9f9c7c0e8`
and records 30 named passes plus one ignored Word test. It includes the
unchanged bundle control and new cleanup control. That executable also contains
D changes. Its module success is successor evidence, not an isolated C2-only
candidate execution or independent acceptance of D. No review build or test
was run. The frozen C2 subset remains independently identifiable for Root's
separate integration and affected checks.

## Not found and limits

No correctness, contract, panic, OOXML, test or structural findings were found
in this bounded repair. Root/worker source leases and existing namespace/input
contracts remain unchanged. The ignored historical Word test remains ignored
for its stated native-observation requirement, with no new waiver or claimed
Word pass.

D crossing-reply authoring, Python/CLI changes, C1 composition and all fresh
native/installed/final integrated gates remain outside this verdict. No feature,
provider or sprint completion is declared.

Review ended after writing and explicitly prose-checking this canonical report.
No source, progress, plan or HLD edits and no remediation were performed.
