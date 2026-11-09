# F-X272, all aspects, pass 8

**Reviewer**: independent Codex reviewer, not an implementer of this repair.
**Scope**: D4 actual publication-frame repair and bounded diagnosis of the two current public custody discriminators. This continues the prior D3/D4 checkpoint. It is not overall feature, Issue 430, release or sprint acceptance.
**Verdict**: 1 defect, 0 smells, 0 nitpicks. D4 is addressed for its examined private publisher/validator invariant. New D5 blocks a clean public readiness attestation.

## Source identity

Root closed all processes and released this report write. The reviewer rechecked these tracked inputs.

| Input | SHA256 |
|---|---|
| crates/rdocx/src/revision.rs | da0a2fd9e3dfb36a8f73ee7e95a09174ce92f7049f33893172acc4d472b6245b |
| crates/rdocx/src/document.rs | 6e2f775923411f110d0fb664d8da73c17886b586044cee18958b0116d64d8a86 |
| crates/rdocx/src/comparison.rs | 5a049765d70dd0b90f493aa4b6da6fee0d4cfa50d29d77e8ff8a2256a7ebae36 |
| crates/rdocx/tests/regression_test.rs | a69529d931af8f4bd164368802309a1f3313faf5a8d6b23743ec2be19577aa24 |
| .claude/plans/F-X272-design.md | 0682bd9cfa3ba7339e7910f77abf5831604207965998e17615a5be1785919290 |

## Defects

### D5, typed complex-field discovery reparses an unbalanced detached run span

`crates/rdocx/src/document.rs:16827`
`crates/rdocx/src/document.rs:17663`
`crates/rdocx/src/document.rs:17685`
`crates/rdocx/src/document.rs:17687`
`crates/rdocx/src/document.rs:15705`
`crates/rdocx/src/document.rs:18190`
`crates/rdocx/src/document.rs:18210`
`crates/rdocx/src/comparison.rs:1293`
`crates/rdocx/src/comparison.rs:1321`

Two actual current installed-Python public comparisons fail with typed story field run scan failed: ill-formed document: close tag </w:ins> does not match any open tag. The first pair retains the same VML textbox and adds a complete body carrier with a simple result owner. The separately executed second pair retains an atomic inline SDT and exercises the corresponding body carrier. Both normalized input main parts in each pair pass XSD validation. The public operations do not complete, and no later lifecycle or output-schema success follows either failure.

The source establishes a detached-fragment context defect. scan_complex_field_marker constructs candidate.scan from the begin marker's enclosing run start through the end marker's enclosing run end. typed_complex_field_candidate_is_admitted then feeds that byte slice directly to an independent Reader before constructing its typed paragraph probe. Runs can occur inside qualified revision wrappers, which the story grammar explicitly admits through CT_Revision::story_raw_is_typed. A first-run to last-run slice does not retain the ancestors that balance intervening wrapper closes and opens. Namespace closure of an owner does not restore missing physical wrapper ancestry inside this detached slice.

The enclosing scan_story_items NsReader loop has already read and closed the owner before the typed-admission loop calls this helper. The actual error therefore comes from the detached run probe after the owner parse, rather than demonstrating malformed owner XML. The helper propagates that fragment parse error instead of classifying whether this occurrence is a supported typed field under its real physical context. The public compare source reaches staged reopening and subsequent selected-revision guards, where checked story discovery is required. Exact runtime candidate offsets, detached bytes, story identity and the particular higher scan subcaller were not captured in these observations. None is invented here, and the observed error plus the cited helper contract is sufficient to establish the public bug.

This is new D5. Existing N3 concerns enumeration order of complete lower CT_Revision parser-context identity and is not reopened by this different structural-fragment failure. Existing D3 publication and D4 custody fixes do not waive this checked-discovery failure.

A bounded repair must establish the actual physical context before passing a run range to a standalone parser. A possible private eligibility discriminator is the immediate physical parent identity of the begin and end enclosing runs, followed by the existing typed isolated probe. That is a technical proposal for the implementing phase, not approval of a particular API or an assertion that every cross-wrapper field is intrinsically invalid. Keep intrinsic complete-field validity and enclosing-field propagation separate from typed-excerpt eligibility. Preserve template field ranges and supported same-context fields. Do not replace arbitrary parse errors with false, suppress real source failures, add a full paragraph rescan/cache or weaken original-revision custody. Verify actual public Box and SDT comparisons and their selected-author lifecycle/schema outcomes after repair.

## D4 repair and finite positive evidence

At revision.rs:1110, the validator now selects only owners in paragraph_publication_frames. The same private map drives the publisher's framed branch. The map admits actual direct body paragraphs and recursively qualified fldSimple/hyperlink children of frames, stopping at atomic SDT, run, drawing and opaque paths. Independent textbox and atomic SDT owners remain in the all-story inventory without being incorrectly required to borrow body carriage. The local qualified inline deletion route and all-story field inventory retain their separate responsibilities.

The existing custody body now exercises boxed validator and atomic SDT positives, a direct standalone field negative, empty input and namespace-alias cases. Root supplied current source-bound 16 Rust tests PASS, fresh 9 Python tests PASS, existing 4 public comparisons with 8 lifecycle outputs and 16 main-part XSD validations PASS, plus four-crate strict all-target/all-feature lint PASS in 52.39 seconds. These are finite supplied-input observations, not proof that the additional Box/SDT public discriminators passed. The reviewer inspected source and the actual public diagnostic result records, and did not rerun or independently reauthenticate every acquisition artifact.

## Actual failing public evidence

The first packet is /private/tmp/rdocx-s96-run-refresh/fx272-d4-public-scope-prepared-v2. Its result records show the Box comparison failure and two valid normalized input main parts. Only its first pair ran. The later SDT branch and lifecycle outputs did not execute.

The separately run packet is /private/tmp/rdocx-s96-run-refresh/fx272-d4-public-sdt-only-prepared-v1. Its receipt SHA256 is 04af339b1fbbf3520b6cffb3278078e986011911e43185471cd36c1fc1cfbe7e and seal SHA256 is 40393c8a827f87e1b98677e5a0dc6f0d9e182b0153bf184559ce611dd5b743d1. Its result records show the same public error and two valid normalized input main parts. The original document's serialized bytes were unchanged after this failed SDT comparison. Atomicity was measured only for this SDT case and is not retroactive proof for the Box case.

Both packets remain NONPASS. Their qualification text describing intended later phases is not evidence that those phases ran. No native Word or package-wide schema outcome follows.

## Smells

Zero independent smells.

## Nitpicks

Zero.

## Handback

D4 is addressed for the reviewed custody quantifier. D5 requires a distinct implementation phase and current-source evidence. Prior D3 and N3 repairs remain scoped observations. Native Word, final union, CLI and every remaining workspace, installed, hosted, hash and risk gate retain their existing delivery-record status. No full F-X272, Issue 430, F-X273, sprint or release acceptance is asserted.

Only this review file was written after root released it. No source, test, plan, HLD, progress, TEMP or other file was modified. No build, runtime experiment, native operation, network action, oracle, commit or push ran during this pass. Return control to root for separate remediation.
