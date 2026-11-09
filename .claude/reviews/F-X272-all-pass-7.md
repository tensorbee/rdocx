# F-X272, all aspects, pass 7

**Reviewer**: independent Codex reviewer, not an implementer of this repair.
**Scope**: current D3/D4 whole-paragraph repair checkpoint against the original D3 before copies and the approved qualified-source, fragment-composition and structural-domain amendments. All affected paragraph publishers, carriage proof, shared balance consumers and resolver preparation were examined. This is a bounded remediation review, not overall F-X272, Issue 430 or sprint acceptance.
**Verdict**: 1 defect, 0 smells, 0 nitpicks. D3 is addressed for the examined checkpoint. D4 remains blocking because carriage validation quantifies owners beyond the publisher's actual framed ownership.

## Reviewed source identity

The tracked inputs were rechecked after root closed all acquisition processes and released this report write.

| Input | SHA256 |
|---|---|
| crates/rdocx/src/revision.rs | be73b579175097f91e98a367ac4bd802a5619a36b4502bdacd00313992b887cc |
| crates/rdocx/src/comparison.rs | 5a049765d70dd0b90f493aa4b6da6fee0d4cfa50d29d77e8ff8a2256a7ebae36 |
| crates/rdocx/tests/regression_test.rs | a69529d931af8f4bd164368802309a1f3313faf5a8d6b23743ec2be19577aa24 |
| .claude/plans/F-X272-design.md | bd80485ba47ddf083cf40f9c2ce992efa675082dea1934ddde7c5966e2cb7215 |

## Defects

### D4, carriage validation refuses owners that remain inside atomic publisher payloads

`crates/rdocx/src/revision.rs:1097`
`crates/rdocx/src/revision.rs:1105`
`crates/rdocx/src/revision.rs:1106`
`crates/rdocx/src/revision.rs:924`
`crates/rdocx/src/revision.rs:980`
`crates/rdocx/src/revision.rs:3340`
`crates/rdocx/src/comparison.rs:17518`

validate_carried_paragraph_fields builds the supplied source envelope and checks every owner returned by the all-story carried_field_sources inventory. An owner with no complete carrier rejects the entire paragraph range. The inventory intentionally includes independent textbox stories and modeled fields inside atomic controls. It is broader than the actual fldSimple/hyperlink publication frames, which recurse only from an existing frame and otherwise retain the child as an atomic payload.

The existing boxed unit supplies a concrete source witness. A complete body complex field contains an opaque drawing with a standalone simple field in its independent textbox, followed by a direct body simple field. The unit establishes that the textbox owner has no carrier, the body owner has one, and the qualified paragraph events remain Begin, Owner, End. Passing those same supplied paragraphs to validate_carried_paragraph_fields inevitably rejects the textbox owner. The public comparison route calls this validator after coarse balancing succeeds. The textbox shell remains inside its opaque run payload and needs no shell-outside carriage proof from the enclosing publisher.

The same quantifier mismatch applies to a qualified simple field inside a direct atomic SDT outside the body carrier. The publisher preserves that SDT as one payload instead of framing the nested simple shell, yet the global inventory still exposes its uncarried owner to validation. Restricting validation only to the body structural domain would therefore leave the shared invariant incomplete. Validation must quantify the actual recursively framed publisher owners, while preserving checks for nested owners that the publisher really does frame.

This is a source-established false refusal, not an executed RED. The boxed test inspects inventory and events but does not invoke the carriage validator on that source. No reviewer runtime or native experiment was performed. Retain D4 because its remaining cause is the repair's inconsistent custody boundary. Use one shared actual publication-frame policy for publisher and validator rather than broadening the all-story scanner or making textbox owners borrow a body carrier. Preserve independent story inventory, local qualified SDT deletion handling and full proof for actual framed shells. Add positive boxed-validator and atomic-SDT discriminators alongside the direct standalone-field negative and empty-input controls before refreshing affected evidence.

## D3 and other examined invariants

The paired deletion-leaf transformation now composes exact opening, untouched inner bytes and exact closing fragments at revision.rs:1019 and revision.rs:1022. Namespace closure is applied only to the actual atomic opening at revision.rs:1039, then appends the transformed remainder. Existing one-tag helpers retain their old fragment contracts. Qualified aliases, empty leaves, quoted greater-than signs, QName-looking attribute values, comments and CDATA are covered by the supplied focused custody body.

The publisher's local inline-source qualification preserves opaque drawing payloads. Per-structural-domain carrier stacks keep independent textbox owners from borrowing body carriers. The crate-private paragraph event adapter uses the existing XmlTree and shared qualified marker/run policy, replacing the former local-name scanner. Its coarse events retain direct Owner semantics and do not substitute for complete per-occurrence carriage proof. The actual framed-owner validation mismatch above remains.

Both resolver preparation entry points retain source-aware retirement and selected-descendant custody checks before a simple shell disappears. Raw simple owners under qualified hyperlink grammar remain source projections without promotion to a borrowed lower typed Field owner. Iterative frames, source ranges and the final output buffer preserve the approved bounded-work approach. No further independent defect was established in this finite repair scope. D1, D2 and N3 are not reopened by these observations.

## Actual acquired evidence

Root closed the current processes and authenticated the source-bound artifacts before releasing this report. The reviewer read the root authentication records and rechecked the tracked source pins. The reviewer did not rerun these acquisitions or claim a separate raw-artifact authentication.

| Current acquisition | Actual result |
|---|---|
| Focused Rust batch | 16 exact tests PASS, source stable, build 376.17 seconds |
| Fresh opt2 Python | 9 exact tests and 27 phases PASS, source stable, Cargo native equals wheel equals installed native |
| Public comparison/schema | 4 comparisons, 8 author lifecycle outputs and 16 main-part XSD validations PASS |

Focused receipt SHA256 is 93d48d61d38617afcb9e70370150fdc95fafea549cae9b7ef389e94d9f3f01bc. Its root authentication is fx272-d3-focused-root-review-v4.json. Python authentication is fx272-python-exact9-root-review-v3.json, with 93 sealed members and receipt SHA256 976d543acdd2da7eed86447ac30ed8ee20d4e7d59baca566341a0ea076f7b051. Public/schema authentication is fx272-d3-public-schema-root-review-v5.json, with 64 sealed members and receipt SHA256 534f3a9abae6d9a52f65b28a798b57a3b0bf01712d2c60fb34bb542249d6b003.

These are supplied-input observations on the pinned checkpoint. They do not execute the missing D4 validator positives, establish every saved package part's schema validity, or provide native Word parity. Earlier failed packets remain failed historical evidence and are not overwritten by these successor results.

## Smells

Zero independent smells.

## Nitpicks

Zero.

## Handback

D3 is addressed within this reviewed repair scope. D4 requires a distinct remediation phase and refreshed affected evidence. No overall feature, Issue 430, release or sprint completion follows. Current CLI, final union, workspace, installed, hosted, hash, risk and native gates remain subject to their existing delivery record. F-X273 main Word, R4 and cache work is outside this checkpoint.

Only this review file was written after root explicitly released the write. No source, tests, plans, HLD, progress, TEMP or other file was modified. No build, test, native operation, oracle, commit or push ran during this pass. Return control to root for remediation.
