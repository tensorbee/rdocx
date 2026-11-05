# F-X272a custody recognition, Root bounded final pass 1

**Reviewed**: frozen review-tranche-v1, ten changed paths, 5272 patch lines.
The bounded source review follows note discovery, custody recognition,
definition and graph cleanup, actual resolver transport, grouped destination
guards and their affected callers. It is not whole-feature acceptance.
**Verdict**: one defect, zero smells, zero nitpicks. Remediation required.

Source manifest SHA256 is
`e6bc68296189308a71a234bd31363751d387eb96474909ec4d0d1f761da342d8`.
Patch SHA256 is
`bc34172d72c27c6f09f0be12e801e36f784c885da256d2ec55431b45635dff1f`.
Frozen comparison.rs SHA256 is
`a21168323ec3d192ff8705c19419deb5a587b978b0d890b1d3c19b26cdacf14b`.
Citations below use that snapshot's repository-relative paths and lines.

## Defects

### D1, custody recognition can consume markers inside an opaque foreign owner

`crates/rdocx/src/comparison.rs:4275`
`crates/rdocx/src/comparison.rs:4350`
`crates/rdocx/src/comparison.rs:4364`
`crates/rdocx/src/comparison.rs:4495`

The decoder walks every lexical element of every related story. It checks
Word namespace and local bookmark name, but does not require a modeled owner.
For N it accepts any enclosing Word footnote with the matching numeric id,
without proving this is a direct actual normal owner in the relationship-
resolved footnotes part. The first descendant paragraph check does not supply
that missing physical ownership proof. Later normal.contains_key checks only
the number in the real part, not the marker's physical owner or part.

A foreign opaque subtree in a header can therefore contain a Word footnote
and first paragraph with an adjacent, otherwise well-formed N bookmark pair.
An unrelated real normal note can have the same id. This opaque pair is then
treated as custody despite being outside the actual note definition. With no
B pair the final orphan-N check refuses ordinary resolution of an unrelated
revision. With a matching B the decoder can associate the opaque marker with
the unrelated real definition and make it eligible for marker retirement.
The source failure is the missing physical-owner predicate, not a claim of an
executed public failure or an assertion that all later guards accept a forged
comparison. No runtime discriminator was run in this review.

Require positive actual note-part and direct normal-owner membership before
recognizing an N pair, and preserve opaque descendants instead of interpreting
their lookalike markup. Apply the same real-owner boundary to B recognition.
Retain a valid owned-pair positive beside a foreign opaque lookalike control.
This is the approved preservation and owner-binding contract, not a new marker
grammar, numeric-id policy, blanket imported-XML refusal or Q-A430-29 decision.

## Smells

None found in the bounded paths examined.

## Nitpicks

None.

## Other observations and acceptance limits

The grouped guard retains actual owner partitions rather than flattening them
before note comparison. Stable positively parsed note-id locations are used
for the dedicated Field witness path, while generic and endnote callers keep
their previous location policy. The private publication observer follows the
actual resolver and validates stage hashes before transporting ranges. No
additional finding is recorded for those source changes in this pass.

No additional panic, XML ordering or public-signature finding is recorded in
the examined diff. This is not a completed audit of every imported context.
Selected Rust, CLI and Python fixtures were read, with installed Python still
pending. Mechanical support authenticated all25initial custody files and
11snapshots and reconstructed all ten changed paths from actual bda1042c Base.
That support supplies no semantic review or current-source runtime claim.

STATUS.json explicitly scopes the older4library,43integration and1regression
passes, current Clippy and pending affected confirmation. Final-source receipts
and the complete45A/14G mapping are still being supplied. Q-A430-29 remains an
unanswered product choice and known held failure. The exact Q-A430-28 native
exception is unchanged. Native, SDK, installed bindings, integrated full gates,
feature completion and parent-provider readiness remain unproved.

Root changed only this review record during the pass and returns to sprint
orchestration for a separate remediation handoff.
