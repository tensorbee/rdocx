# F-X279d final union repair, Codex pass 1

**Reviewed**: seven working paths, 144 additions and 57 deletions, against
19cde21f. Author preparation is the actual Codex worker. Root independently
reviewed the implementation and actual scoped observations.
**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Correctness and contract

The integrated H1 Python gate exposed a real root-move failure after reply
creation. The selected comment proof now filters qualified physical markers
only after preserving their original ordinals at
`crates/rdocx/src/comments.rs:818`. Complete ownership, one range/reference,
root-only eligibility and representable same-owner anchors remain required.
The comment branch uses this proof before generic nesting checks at
`crates/rdocx/src/comments.rs:841`. The existing destination reference helper
uses the same proof at `crates/rdocx/src/document.rs:30821`.

Generic pairing and generic range APIs retain their established strictness.
The new regression at `crates/rdocx/src/comments.rs:4348` challenges one and
three replies, both move entry points, exact definition/extended metadata,
unmoved reply anchors, refused unknown/reply/invalid destinations, byte
atomicity and reopen. Three replies discriminate a residual crossing pair
after the selected root leaves. Source marker ordinals remain physical.

## Other aspects

Panics: zero findings. The selected range remains uniquely checked before use.
OOXML: zero findings. The original reference replacement and namespace replay
are retained. No new XML serializer or public type is introduced.
Structure: zero findings. Only an existing private proof becomes crate-visible.
Tests: zero findings. The actual H1 failing root-move case remains as a binding
consumer. Test-only fixes retain the actual annotation text handle, schema
normal-note default, early malformed-XML refusal, duplicate archive exercise,
known bundled PAGE face and two physically different caller-font controls.
No diagnostics are broadly ignored and no required case is converted to skip.

## Actual observations and limits

Actual scoped all-target check and strict all-feature Clippy exit zero.
The exact new Native regression passes and the complete comments module has
33 passes, zero failures and one existing ignored native generator case.
An initial unqualified exact filter selected zero and is retained as a
selection mistake, never accepted as GREEN. Actual exact execution supersedes
that missing test selection.

The corrected private Word binding cohort has 32 passes on the authenticated
existing development extension. Its earlier six note-loop failures are
retained. This is test-premise evidence only. Fresh release wheels, complete
installed suites, typing, final source hash/pixel checks and native/hosted
acceptance remain due. This review does not close the story or sprint.

Actual logs, source pins, compiled binary digest and child outcomes are in
`/private/tmp/rdocx-ci-triage-42293933089-20261106/python-repair-proposal/ROOT_SCOPED_GREEN.json`.
No source was changed during this review pass.
