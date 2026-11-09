# F-X278, Codex final review, pass 3

Reviewed the Issue433 comment6715174819 increment on `sprint/s96`, HEAD
`cec87f9501eb13d0d5d60572472b227483262620`. This is Root's own bounded review
after final2, not a fourth independent implementation peer. Earlier three peer
passes and integration reviews retain their attribution and scope.

Scope is three tracked paths: 59 added test lines in the existing engine module,
the existing plan's A278-2 follow-up and matrix/test entries, and five HLD08
example lines. Existing unrelated integrated changes are excluded. The current
four-path source/progress hashes and actual test binary hash are retained in
`S96-final-evidence/f278-issue433-numbering-comment-20261108-1/SCOPED_RESULT.json`.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Contract and correctness

Read the approved F-X278 plan, current progress, prior implementation peer3,
integration composition/carry reviews, canonical workflow and affected rendering
specification. The user requested this small follow-up. Numbering-level
properties were already an approved A278-2 consumer. The current actual merge
order at `crates/rdocx-layout/src/engine.rs:11390` and `:11395` applies numbering
before direct paragraph formatting. Pair replacement at `:14600` clears the
stale inherited hanging value. The equivalent model merge at
`crates/rdocx-oxml/src/paragraph_properties.rs:2027` already has the same rule.
No production implementation changes in this increment.

The new regression at `crates/rdocx-layout/src/engine.rs:27501` uses the reporter's
exact level90, left1440/hanging360, direct left0/firstLine720 and repeated text
in three paragraphs. It observes transformed positioned text, all three ordered
numbers, first text on the marker baseline, and wrapped text at the margin.
It rejects a numbering-path merge that leaves hanging360 active. It adds no
source file, module, trait, generic, feature, public API or binary fixture.
All unwraps/assertions are confined to the existing test module. The input is
constructed in code and uses deterministic bundled fonts.

The test retains the existing style, inverse-pair, same-element and vertical
controls. It does not adopt the reporter's sequential assignment literally,
which would clear the stored firstLine when one element contains both values.
The existing preserve-both model contract and its residual calibration debt
remain intact. HLD08 and the existing plan explicitly include the new numbering
example without narrowing any of A278-1 through A278-7.

## Evidence and limits

Actual scoped engine Cargo command exits0, 8 passed, 0 failed, 0 ignored,
414 filtered out. The added test is explicitly named in the successful log.
Actual all-target/all-feature Clippy with denied warnings exits0. Final fmt,
prose and diff checks exit0. Original fmt line-wrap failure is corrected before
these successful checks. No broad suite was repeated.

The exact python-docx1.2.0 reporter fixture on the identified current ordinary
Native `b9143d1924f48784287c6e5e4d46c5722d150cd90333acb108cd4da1dd7a1274`
puts the three numbers at126pt on its90pt margin and wrapped lines at90pt.
Actual save/reopen retains the same coordinates. The initial bbox extraction
misclassified shared-line fonts because their glyph tops differ0.492pt.
The corrected one-point grouping remains well below the14pt next-line distance.
That extractor correction changes no library behavior or acceptance positions.

Actual isolated published0.16.4 Native
`24c70012d10f6f6c2456bc0d10dfc510704cfc087bc28b7d98d8a8e364f446db`
fails the same source ZIP's number-position assertion with72pt instead of126pt.
The current Native passes the identical `to_pdfa_deterministic` bundled-font
observation at126pt. Both source hashes are
`4d7ae4b869f8033bcba4f86bbebe2a16ec81b7656b9330b986f9dae69d88d6bc`.
This is actual prior-release RED/current GREEN rendering evidence. It is not
an execution of the new Rust test on the historical complete source tree.

The4309-input applicability check finds only plan, HLD08 and engine differences.
Removing exactly the new cfg(test) function recovers the release engine hash.
The existing wheel therefore still identifies the unchanged production source.
No new performance run is claimed. Prior77/79 performance failure remains open.

python-docx supplies the fixture, not a rendering oracle. Native Word and
LibreOffice qualification, the other Issue433 final criteria and F-X275 release
riders remain pending. No baseline, acceptance status, ledger, commit, push,
reporter reply, tag or publication was changed. This review edits only this
review record and ends the review phase.
