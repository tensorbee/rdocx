# F-X275, all aspects, pass 1

**Reviewed**: Complete working diff on `sprint/s96`, Base and HEAD `492950f553e7e114b9766f4ce5502e27a9a5fdff`.59 tracked files,338 additions and262 deletions before this report. Independently rehashed all59 inputs in `f275-current-candidate-source6.json`, with zero mismatches.

**Verdict**: 0 defects,1 smell,0 nitpicks. The microscope exit condition is not met until S1 is repaired or assigned an F-ID. This review does not complete release preparation or qualify publication.

## Defects

Zero.

## Smells

### S1, current packaging intent still names the old .3 release pair

`docs/hld/15-build-and-toolchain.md:703`

The present-tense Packaging paragraph still says the shared0.14.3 family publishes before Word0.16.3 where the verified internal pins require it. It is presented as the current preparation/publication rule, without historical attribution. The same section now prepares0.14.5/0.16.5, and the actual workspace pins select0.14.5. Readers following this release-order instruction receive an obsolete family pair. The approved F-X275 HLD impact explicitly includes current prepared versions and previous published identity. Update this current instruction to the selected .5 pair while preserving genuinely historical release accounts.

## Nitpicks

Zero.

## Contract and source review

Read repository guidance, canonical WORKFLOW, microscope, release, release-notes, verify and risk-routing instructions, the approved F-X275 design/progress, and applicable HLD10/14/15 sections. F-X275 is in-progress in CURRENT_SPRINT and BACKLOG. The design remains approved with final work unchecked. F-X270 is done and unchanged. No source, formatting, test, build, native application or external mutation was performed in this review. The only write is this report.

Inspected the entire working diff, including explicit manifests, workspace pins, lockfile, both Python projects, Rust current-version assertions, README examples and measurement constants, CI WASM assertion, both new changelog sections, policy tests and status/checklist edits. There is no new runtime algorithm, public operation, trait, generic, source module, source file or release machinery. The unpublished rpptx-wasm carrier remains0.12.1. Existing exact7Word and15shared/PowerPoint publication boundaries remain intact.

Independently parsed old and current Cargo.lock. Exactly26local package versions differ, with no other package-field or external dependency change. Independently compared the entire changelog suffix from `## v0.16.4` against published24af95d3. Bytes are identical, SHA256 `50b3e43c22f44f276537eea506b1e40136945fb8f9228d64a4bf50f5d4c244f8`.

## Claims and acceptance challenge

Compared the new notes with F-X272to274 contracts, current product sources and their latest applicable independent reviews. The Word notes retain bounded revised-owner preservation and atomic refusal, the destructive resolution-on-independent-copies workaround, new-only CLI counts, five supported page formats with Roman/alphabetic bounds, selected furniture/source-navigation scope and the ordinary-frame/vertical-text limits. The amended F-X273 contract separately authorizes explicit markers for newly authored add_footnote content while preserving marker absence on imported source. The notes do not promise unrestricted native Compare or general native layout equivalence. The pulldown-cmark wording names the patch and disabled wikilink option without asserting public-importer exploitation.

The Word contribution inventory is exactly Issues423to426 and PR427/428. The rendered-note assertions require all six links and both authenticated handles. Issue423 credits set2374, and the remaining reports/reference implementations credit hadim. Adapted work is described as hardened equivalents. The PowerPoint section makes no new product capability or fix claim and carries no new external contribution. Historical S95.4 credits stay in their original sections. Read-only intake refresh and independent authentication receipts retain unchanged bodies/heads and zero new comments/reviews. Donor failures/skips are not S96 acceptance. Planned notification text remains unsent.

The current policy tests distinguish stale carrier versions, missing new family notes and exact new contribution sets while retaining .4inventory controls and both historical suffix checks. The focused mutation controls still reject mismatched carriers and invalid build-only/package inventory states. A passing inventory or wrapper is not treated as native, installed or hosted acceptance.

## Evidence examined

Evidence root: `/private/tmp/rdocx-s96-run-refresh/release-prep`.

- `f275-local-red1.log` retains three named tests, three assertion failures plus one missing-section error, with historical .4suffix preservation passing. `f275-local-green2.log` records the same three tests passing. This is actual policy RED/GREEN, not a product regression claim.
- `f275-focused3.log` records eight selected policy tests passing. Assertions and their changed inputs were inspected rather than substituting the aggregate count for coverage.
- `f275-static4.json` and associated logs retain successful prose, generated-adapter, whitespace and formatting checks. Current report prose remains subject to the caller's affected metadata gate.
- `f275-versioned-workspace-clippy6-receipt.json` records actual child exit0 for workspace all-targets/all-features strict clippy, with devopt1, debug assertions and overflow checks enabled. The log finishes successfully in2m42s. Independently verified its SHA256 and unchanged guarded source inputs. Root supplied the observed completed child outcome during this review.
- `f275-archive-preparation1-receipt.json` identifies22actual .5archives. Independently verified every retained archive hash, package-root identity and size below10MiB. These were produced with cargo package allow-dirty/no-verify and local patches. They establish measurement preparation only, not the required clean verified publication dry run. `f275-archive-inventory5.log` records27README and22inventory checks after measurement-row updates.
- Registry receipts retain all24occupied .4queries and all24absent .5queries, plus matching .4origin tags and absent .5tags. These are mutable preparation observations, not a substitute for fresh release preflight.

## Pending acceptance

A1toA3 have local carrier/history/note evidence. A4 still needs the final-source clean verified22-package dry run and complete README qualification. A5/A6 require actual SHA-bound hosted jobs, downloaded authenticated payloads, host CLI identities, installed Python3.9/3.12priority calls and Python3.12typing/stubs. No older artifact or donor CI status supplies these outcomes.

A7 remains the complete integrated union, including outstanding native/source/composition observations, full installed suites, WASM, deterministic hashes/pixels and applicable supply-chain checks. A8 requires final metadata reconciliation, clean sprint review/close provenance, exact main identity and separate final publication approval. The canonical validated consumed F-X271 source-availability checkpoint remains the only proposed source route. This report authorizes no push, dispatch, tag, publication, notification or new exception.

## Not found

Correctness: no established carrier, lock, family selection or current test-literal defect.

Contract: no additional unsupported scope expansion or manufactured acceptance pass. S1 is the remaining current HLD consistency issue.

Panics: no new untrusted-input indexing, unwrap, arithmetic or runtime path introduced.

OOXML: no schema-order, namespace, whitespace or opaque-subtree change introduced by this preparation diff.

Tests: no established dropped historical control or false artifact/native/hosted acceptance claim.

Structure: no new structural abstraction or unauthorized source file.

Return control to the orchestrator for a distinct remediation phase. No remediation occurs in this review.
