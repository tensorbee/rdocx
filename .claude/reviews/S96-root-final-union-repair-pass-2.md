# S96 H5 bounded independent repair review

Verdict: 0 defects, 0 smells, 0 nitpicks in the assigned production and integration-test repair diff. This is private support for the existing S96 repair pass 2, not a new peer cycle or sprint acceptance.

## Scope and contracts

Read current CLAUDE.md and .claude/WORKFLOW.md. Reviewed the complete diff of crates/rdocx/src/document.rs, crates/rdocx-layout/src/paginator.rs and crates/rdocx/tests/integration_test.rs against HEAD. Consulted operative F-X273 legacy producer closure, F-X277b A277-2/A277-3, F-X272 scoped/no-op boundaries, F-X272a physical count and complete-unit selection, F-X280b note-reference/duplicate validation and F-X281 A281-7. Other dirty plan/Python/HLD changes are outside this bounded review.

The optional endnote lookup at document.rs:27297 is a best-effort scan for synthesizing a missing Hyperlink style. It no longer promotes an unavailable optional lookup into a failure of unrelated save operations. The actual note identity lookup remains strict at document.rs:37395. Existing duplicate-note relationship refusal coverage remains in the executed batch. This does not weaken scoped reader budgets or change the substantive note reader.

The three paginator matches at paginator.rs:6676, :6714 and :6730 recognize the same legacy MultilingualText carrier produced by F273. They leave nonlegacy rich carriers excluded. Existing source span, generated dash, field/note and direction-adjacent prefix checks are retained. There is no source-node normalization change or diagnostic print in the reviewed diff. Actual diagnostic-2 stderr:227-232 proves the generated dash is a legacy MultilingualText item with no source. line.rs:1847-1853 is its real producer. The legacy rendering arm at paginator.rs:5965 emits positioned Text, so no generic logical-index branch change is necessary. This review does not claim that the preexisting inferred-hyphen mechanism is a universal provenance validator.

Integration fixtures explicitly set the formerly assumed kern-on context, preserving native geometry/raster assertions. The theme test retains full PNG equality at integration_test.rs:2375 and :2401. The restored hybrid case retains logical PDF and SVG ordering plus actual visual origins at :13844 onward. Authored note-space expectation at :24551 matches F277b's marker, explicit space, caller text law.

The Issue430 helper at integration_test.rs:37553 operates only on its named canonical generated fixture. It gathers actual note revision IDs and the actual body reference ancestor, changes their author attributes only, checks complete author inventory, and preserves logical tokens, physical IDs, order and unowned producer content. Updated tests distinguish unique physical records from logical custody units. They retain partial scalar-ID atomic refusal, full-author selection counts, both resolution orders, mixed lifetimes, absent-part cleanup, tombstone/reference-only ordering and invalid unowned content. F272a section 20.14 item 7 at plan:788 explicitly permits these complete author scopes after physical IDs differ.

## Actual evidence and limits

repair-1 records actual 14 PASS / 2 FAIL, retained as historical. Subsequent failed hybrid candidates and diagnostic-2 are also preserved. Current repair-4 compile is exit 0 with source stability. Its actual selected run is 16 PASS / 1 FAIL / 0 ignored. The only failure is the theme LibreOffice subprocess exit at integration_test.rs:2399, not a failed PNG comparison. The hybrid case and stored-field control both pass the current source.

Separate repair-3/THEME_NATIVE_RESULT.json records the exact theme test exit 0 in 7.831 seconds. Its integration test and document source hashes equal repair-4. Its paginator differs only in the bounded conditional-hyphen classification helper, a path not used by the plain single-run theme fixture. Thus the earlier actual strict full-PNG result remains applicable to that fixture, without relabeling it as a repair-4 execution or calling the failed combined run green. No LibreOffice version, Word execution or blanket native result is invented by this review.

All seven repair-4 SOURCE.json entries were authenticated against current files at report writing. Comparison/redaction/Cargo source remain as pinned. No source, test, progress, Cargo, runtime or UI mutation was performed by this reviewer. Final integrated/hosted/hash/packaging/native/performance gates remain separate and pending where unperformed.

## Authenticated source hashes

- crates/rdocx/tests/integration_test.rs: `0e4d93b10d70aa2a34d50473a9081ab6a83f2eb6e0aaee9afde05e3e5ac6b62d`
- crates/rdocx/src/document.rs: `11cf683070cdeb1514295554cb2815fe202c9ffa5877d12ffc55be6d51f56533`
- crates/rdocx-layout/src/paginator.rs: `7713771f6d1e30ed6fdcdc122f031e8a4e2c1c3131d2d6f5544f3968432b8bf0`
- crates/rdocx/src/comparison.rs: `e3fd8e11613206c6c6c010fc67fd27b0ea7343d6458ef4688894a9620776d5ca`
- crates/rdocx/src/redaction.rs: `9a5a49994b594061ee8fefe63758af53e00f49ce6eff39bfeebc1569b6df239f`
- Cargo.toml: `c597660166215d14813213e1f40f12f3f08302be265470a9771a7e52228bf7c5`
- Cargo.lock: `918faaaf6a2068b6233436eadc23dac2aa68a6d8670f1798b822baf88fda3d5a`

## Later actual host observation

HOST_RESULT.json now records the exact 17-case current-source batch exit 0 in 27.653 seconds, source_stable and binary_stable true. Its executable SHA is b12bc53d22dc5c3ce6d413f103171f5b2cc6a33bb7ea6ba1333208a9434ec811. The prior failed subprocess run remains preserved. This supplies the current-source theme full-PNG execution as well as all other selected cases, with no new blanket final acceptance claim.


## Root final adoption and subsequent actual host execution

Root independently inspected the complete bounded diff and the named contracts.
The two actual product repairs restore optional imported-endnote preservation
and normalized legacy conditional-hyphen association. The original logical
PDF and SVG assertions are restored, with no temporary logging or source-node
normalization candidate retained. No public surface or resource budget changes.

After the independent review, the exact guarded current binary and source ran
all16affected cases plus the source-less stored-field control on the normal
native host with pinned LibreOffice and Poppler. The actual command passed
17cases with zero failures in27.62seconds. Source and binary guards remained
stable. HOST_RESULT.json and full stdout/stderr are retained in the repair4
packet. Failed sandbox combined runs remain failed and are not relabeled.

The inherited report above was produced by the actual independent Codex agent
f279c_implementation, SHA256cfeeee5ca3ab7021b98c6a3e84aac7a93b38c638938ecd5340fd43d24c5d1061.
Root records zero defects, zero smells and zero nitpicks in this bounded pass.
Full current-source hosted, packaging, binding and performance acceptance is
separate and remains due where unperformed. This is not the final sprint review.
