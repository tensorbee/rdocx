# F-X279c, microscope, pass 1

**Reviewed**: Working diff against `b03b206be3d6028ac979323304a3419c7f4b1374`
on `work/f-x279c-codex`. Six leased files, 677 additions and 17 deletions.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Bounded source review is clean.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Contract and consumer review

Read AGENTS.md, CLAUDE.md, WORKFLOW.md, the microscope command, approved
F-X279c design, split final1, progress and the affected HLD sections. The sprint
and backlog identify F-X279c as in-progress. Reviewed the full Native comments,
lower comments, Python test and HLD03/04/10 delta.

Traced add_comment through its dated/staged implementation, story-range add,
text add and the separate reply_to_staged path. The shared authoring setup at
`crates/rdocx/src/comments.rs:1680` checks both existing style types before
adding either missing definition. Existing style builders and staged publication
retain graph validation and rollback. Missing definitions append, and basedOn
is added only for an existing base. Existing style definitions are retained.
Comment paragraphs at `crates/rdocx/src/comments.rs:2189` receive CommentText.
Only the first paragraph receives the CommentReference annotation run, ahead
of its text run and after paragraph properties in serialization.

The authoring root merge uses the MC namespace URI and local Ignorable name,
retains other tokens, reuses a genuine alias and chooses an unused prefix when
canonical mc is foreign. Root w14 is added only when absent. Lower serialization
at `crates/rdocx-oxml/src/comments.rs:115` inspects original root, comment and
retained paragraph-root attributes before writing fixed names. The new guard
at `crates/rdocx-oxml/src/comments.rs:386` refuses foreign canonical w/w14
bindings rather than coercing their URI. Compatible aliases, opaque descendants
and genuinely unused declarations discarded by the existing paragraph parser
were considered separately. No new import parser policy is introduced.

Public add/reply paths publish staged candidates only after the relevant
serialization succeeds. Style graph failure and namespace refusal cannot
publish candidate anchors, comments, styles or package changes. The existing
clone uses separate layout caches. Python refusal controls also retain held
paragraph/run views and exact package bytes. Untouched imported comments keep
the existing non-dirty package path. No commentsIds producer is added, and the
existing-part control checks its bytes, relationship graph and content type.

## Evidence examined

Inspected actual assertions and retained logs under
`/private/tmp/rdocx-s96-f279c-evidence`. Historical Native baseline execution
has four named failures, and lower baseline execution has two named failures.
Successor evidence includes five Native controls, eight lower controls and the
all-entry-point empty-line/paragraph-id control. The latter now checks the
new paragraph styles and annotation placement as well as the existing entry
paths. Python logs retain eight focused passes and twelve related passes.
These are overlapping cohorts, not an additive total. No review builds or tests
were run. Scoped check, final Clippy, formatting, prose and adapter logs were
read as local evidence, not as integrated verification.

The extension on disk matches `python-artifact.json`, SHA256
`c315309642e4219b0600552f416d4e9bfea8ff04ac8e35d93aaf310492a7a419`.
That receipt records Native comments source
`47f962a3da330e02358a0f9a70211e47c0b9052581c6045010bf96b9fa6795d3`,
whereas this reviewed final Native source is
`e355e4b8971f94f7dfe88487bb2bfd7a12066a928d0d19c1a6f884d3a4bddcac`.
The lower comments source matches the receipt. This review does not assume
that the Native difference is test-only or relabel that extension as built from
the final source. Exact before/final source reconciliation, or a fresh build if
production differs, remains an evidence follow-up rather than a source finding.

## Source boundary

The six leased file hashes were unchanged across review reads. Full working
diff SHA256 is
`849cf765c00997d4a65cbc050c124dcca3076b0c389be90beecdb4348e62107c`.

- Native comments: `e355e4b8971f94f7dfe88487bb2bfd7a12066a928d0d19c1a6f884d3a4bddcac`.
- Lower comments: `bf26daa72950d513e923f6302a00a9c545d7709bd69a9cc146527c086f992865`.
- Python tests: `75bea57f508d057e8f8ae7bf8604144c73b2057c6d6d0d1cc8fa30f416a78a5d`.
- HLD03: `b483f7168957b285d4cc863d7123630dc002328e879b7b7fcb1c7edffd6bbd2e`.
- HLD04: `1b53509e4f56bf28ac2ddaaeb8aa77fd381cb3bf11a384089ee5e4e485a45bfb`.
- HLD10: `cee5526895646498a20d56c0588fabcc80450bd7dfae9424a1ce8b450ff2c841`.

## Not found and limits

No additional correctness, contract, panic, OOXML, test or structural findings
were found in this bounded delta. No new public API, module, generic framework
or C1 dependency was introduced.

Fresh Word execution of the exact integrated candidate remains pending. The
historical candidate hash and native-version constants are unchanged and are
not a claimed passing final gate. No guard was ignored. C1 lower comments
parser/emitted-scope reconciliation, child composition, installed typing,
schema/hash checks and full integrated verification remain outside this pass.
No whole-feature, provider or sprint completion is declared.

Review ended after writing this report. No source, progress, plan, HLD or
tracker changes and no remediation were performed.
