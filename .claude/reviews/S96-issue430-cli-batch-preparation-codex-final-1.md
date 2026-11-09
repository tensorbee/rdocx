# S96 Issue430 CLI batch preparation, Root pass 1

Reviewed the frozen608line private batch.py, its eight mocked controls and
actual Native fixture and CLI consumer definitions. Source SHA256 is
`dfa36175f7cc935687d4c07e4ec3d665f615d1558032c6f97114c30d001e32ef`.
Verdict: three defects, zero smells, zero nitpicks in the private helper.
This is final-gate preparation review, not another F-X272 implementation peer.
No product, SDK or Word execution occurred in this review.

## Defects

### D1, Q29 recipe changes a prior revision-bearing paragraph

`/private/tmp/rdocx-s96-f272a-local-evidence/final-h-cli-batch-preparation/batch.py:482`

The first comparison puts note5's revision in the first paragraph. The second
comparison inserts reference6 into that same paragraph by replacing its first
closing tag. The supported prior-revision policy refuses that change. This
cannot generate the promised Q29 history, and the final batch stops before its
selective-resolution evidence. It also does not reproduce the existing Q29
case at `crates/rdocx/tests/integration_test.rs:38088`, which adds references5,
6and7in one comparison and then selectively rejects the two note units before
the independent reference-only insertion. Use that actual fixture and its
actual listed ids for Q29, with its own explicit native substitution limit.
Keep the two-comparison mixed-lifetime recipe separate and do not modify any
prior revision-bearing owner. This is source-derived, not an executed product
failure in this review.

### D2, selector numeric identity accepts a JSON boolean

`/private/tmp/rdocx-s96-f272a-local-evidence/final-h-cli-batch-preparation/batch.py:357`
`/private/tmp/rdocx-s96-f272a-local-evidence/final-h-cli-batch-preparation/batch.py:405`

Selector dictionaries are compared using Python equality. An expected id1and
reported id=true compare equal, although the report violates the numeric id
schema. The outer resolved count is correctly type-checked, but its nested id
is not. Require exact integer type for a reported id selector and preserve the
exact selector shape. Add the boolean corruption alongside the existing wrong
selector mock and ensure producer_complete remains false. This is a collector
counterexample, not a claim that the Rust CLI emits booleans.

### D3, mixed-history terminal check ignores part presence and remaining revisions

`/private/tmp/rdocx-s96-f272a-local-evidence/final-h-cli-batch-preparation/batch.py:499`
`/private/tmp/rdocx-s96-f272a-local-evidence/final-h-cli-batch-preparation/batch.py:502`

The history loop discards the final listing and checks only normal_ids. A result
with no normal notes but an incorrectly retained special-only part passes the
mixed-history check. Remaining revisions can likewise go unnoticed. The
existing mixed-lifetime postcondition requires absent part after both rejections
when originally absent. Q29 retains the original unused part and must resolve
its independent reference revision separately. Check exact expected presence
as well as normal inventory and remaining active revisions at the correct
phase. Keep the full graph/raw-equivalence collector debt explicit. The helper
must not label this limited recipe complete when these actual promised terminal
facts are false.

## Smells

None in the bounded helper review.

## Nitpicks

None.

## Accepted boundaries and retained limits

Default preparation invokes no children. Pinned synthetic inputs, fresh output
paths, exact child status/logs, input baselines, expected refusals, artifact/build
receipt joins and pending expected-invalid SDK rules are explicit. Actual CLI
comparison counts are created counts, so the existing baseline-plus-created
listing check matches commands.rs1998to2039. Exact source_h and artifact hashes
remain joins authenticated by Root, not automatic compiler authority. Word
substitution receipts remain supplied claims requiring actual UI authentication.
No generated package is called Word-admitted or schema-valid by preparation.

The current mock and XML checks are preparation observations only. The retained
30cases/55inputs do not qualify all45A/14Gor replace installed/native/full gates.
Native staged-history resumes, full terminal equivalence and expected-invalid
SDK details remain explicitly pending. Q28's exact shared-reference positive
stays UNPROVED. Root changed only this review record during the pass and returns
to orchestration for private helper remediation.
