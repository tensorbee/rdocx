# F-X271, all, pass 1

**Reviewed**: independent Codex microscope of the uncommitted worker diff on `work/f-x271-codex` against Base `376a8cec`. Five files, 54 added lines and 14 removed lines. Implementation, tests, HLD intent and plan evidence were read. No implementation file was modified by this reviewer.
**Verdict**: 0 defects, 1 smell, 0 nitpicks

## Defects

0 defects found in the retrieval, checksum/version preservation, macOS size bound, Linux safety boundaries or roadmap-policy correction.

## Smells

### S1, active contract still describes an unavailable and unstarted implementation

`.claude/plans/F-X271-design.md:104`
`.claude/plans/F-X271-design.md:108`

The active Open questions text says archive authentication is unavailable, and Contract closure says the bound must be checked against actual bytes and that no source is approved or started. The same approved plan records authenticated platform archives and passing local implementation evidence at lines 121 to 127. A resuming worker consumes this plan as current authority, so its readiness statements now conflict. Reconcile these current-state paragraphs with the actual authenticated local implementation while preserving supported-target installation, hosted CI and final sprint acceptance as pending. The historical baseline and failed Grok attempts remain valid history and should stay attributed.

## Nitpicks

0 nitpicks.

## Evidence inspected

- Read the actual RED and GREEN transcripts at `/private/tmp/rdocx-s96-run-refresh/fx271-local-red.log` and `fx271-local-green.log`. The baseline had two failures, retrieval-route and roadmap placement, with the runtime guard control passing. The successor ran the same three named cases and passed. No full suite was rerun by this reviewer.
- Independently rehashed both retained archive files. Linux is 218075885 bytes with SHA-256 `2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e`. macOS is 297407265 bytes with SHA-256 `c99fb4fe574437fc4cb820a4ca15271bca325920861f7139858b36d7f9df78ad`. Both exactly match their independently retained source pins.
- Independently listed the Linux archive without extracting or running binaries. Its 46 members total 218024484 declared bytes under `LibreOffice_26.2.5.2_Linux_x86-64_deb`. Both required core and Impress packages exist. No negative-size, absolute/traversal or non-file/non-directory entries were found. The actual safe-extraction receipt reports 42 Debian packages. This does not establish installation or executable runtime identity.
- Read each changed caller and relevant guard. Linux retains its original streaming download, member-count and expanded-byte bounds, safe extraction and exact executable identity checks. macOS authenticates the archive before readonly mounting, retains the exact executable identity check and adds a 320 MiB curl bound. No TLS bypass, fallback version or success short-circuit was introduced.
- Inspected the macOS negative mutations for unbounded retrieval, wrong checksum and writable mounting. Inspected the existing Linux hash/size/extraction/version/prefix controls and install-before-expose order. The roadmap change updates only the authorized future set while retaining duplicate, coverage, size, dependency and cross-record assertions.

## Acceptance coverage and pending evidence

A1 and A2 have genuine archive-byte authentication, with actual supported-target installation and hosted runtime outcomes still pending. A3 has focused runtime refusal controls. A4 has observed old-source RED and repaired-source GREEN. A5 remains pending actual source-bound Test, MSRV, Word fidelity, rpptx Python and CI-gate runs. All final full, hash 49, deterministic pixel 7, native/oracle, installed runtime/typing and applicable S96 risk riders remain pending. No final acceptance or dependency completion is granted by this local review.

## Not found

- Correctness: 0 implementation defects found.
- Contract: 0 product scope or API deviations found. S1 records the current-plan readiness inconsistency.
- Panics and unsafe input: 0 newly introduced unbounded or unchecked source paths found.
- OOXML: 0 findings. This diff changes no product parser, serializer or schema placement.
- Tests: 0 additional findings. Assertions and negative mutations were inspected, and actual individual outcomes were consumed rather than inferred from an aggregate count.
- Structure: 0 findings. No new source module, crate, trait, generic, public API or dependency was added.
- HLD intent: the only declared impact file, HLD 15, describes the implemented retrieval route and retains the distinction between authentication and actual oracle execution.

Return to the orchestrator for a separate documentation remediation. A new microscope pass must evaluate S1 after that edit. No source repair occurred in this pass.
