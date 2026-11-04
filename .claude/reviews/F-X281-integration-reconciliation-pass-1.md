# F-X281, integration reconciliation, pass 1

**Reviewed**: Only the three canonical files textually merged when preserved carry was restored over reviewed F281 integration commit `b3d6e4f241db985b3a5c421b97083c8ed9faac63`: `crates/rdocx/src/document.rs`, `crates/rdocx-py/tests/test_core.py` and `docs/hld/10-bindings-spec.md`.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this bounded composition. This is not another full implementation peer, confirmation pass, F272 acceptance or final gate.

## Defects

None found.

## Smells

None found.

## Nitpicks

None found.

## Source custody

Canonical HEAD is the stated F281 integration commit. Compared the actual canonical files against both immutable commit objects and the archived before-carry files under `/private/tmp/rdocx-cross-review/canonical-carry-before-f281-integration-20261102T190541Z/files`. The two immutable commit sides are the F281 integration commit and its parent. The carry side is the actual archive, not a reconstructed approximation.

For each file, independently compared the complete ordered edit payloads, including removed and inserted bytes, in both directions. The parent-to-F281 edits equal the archive-to-current edits. The parent-to-archive carry edits also equal the F281-to-current edits. All six comparisons are exact. Line numbers shift, while the actual edit payloads do not. This establishes that both inputs' edits survive the textual composition, rather than relying on absence of conflict markers.

| Path | Archived carry SHA256 | Current merged SHA256 |
|---|---|---|
| `crates/rdocx/src/document.rs` | `1dccc6ca14b970aa1a0350d8ab303273c54639800cd7d99f5343916a0ba1eac6` | `a013f286c467d47b349e1fd5527b3e6150affbd44d861fe89ff716dd3365394d` |
| `crates/rdocx-py/tests/test_core.py` | `478562b717bd1d10c7f46bf060a45d6e2f0dc55cf79511b2d573871e8767c93b` | `cf20aaadc9220c7579da8deb12a26368f51f67ae6999c8ecfa7c232c06d18200` |
| `docs/hld/10-bindings-spec.md` | `a18fb043446cdb17dffacd0428404acb9fca1fcfcb82f92906fe8adc3561b570` | `8e7198be76400de0e65010f3e46ff3e4059bb1d0bd1d6e01bc80426aba328caa` |

## Not found

No lost F281 image annotation, chart/export API, diagnostic or test intent was found. The seven fresh annotation calls remain at `crates/rdocx/src/document.rs:25220`, `25240`, `25311`, `46511`, `46550`, `46591` and `46691`. The relationship/story classifier annotation remains at `crates/rdocx/src/document.rs:50068`. The string methods and diagnostic tuple methods retain the reviewed forwarding behavior at `crates/rdocx/src/document.rs:46766`. The private HTML builder still projects chart title and data from current relationship target bytes, returns the completed chart map and diagnostics, and leaves HtmlInput fields unchanged at `crates/rdocx/src/document.rs:46818`.

No lost carry field/publication scaffolding was found. Exact carry edits retain the namespace ancestor inventory at `crates/rdocx/src/document.rs:13302`, strict incoming-owner field checks at `crates/rdocx/src/document.rs:14960`, current-source body span handling at `crates/rdocx/src/document.rs:25848`, namespace-restored source at `crates/rdocx/src/document.rs:29312`, and original validated paragraph input publication at `crates/rdocx/src/document.rs:32052`. F281's added export readers and annotation callers do not replace those publication operations or grant their field facts additional authority.

No lost test case or assertion was found. The existing binding owner/no-op/atomic-refusal test retains its carry additions at `crates/rdocx-py/tests/test_core.py:8203`. Both F281 Python cases and their helpers remain intact at `crates/rdocx-py/tests/test_core.py:16297` and `16327`. The five F281 native tests remain at `crates/rdocx/src/document.rs:50113` through the test starting at `50346`. The separate carry regression additions remain exact after their shifted locations. This is source membership, not a claim that any test executed in this pass.

No HLD intent was lost. Checked raw-owner field validation and namespace-preserving publication remain at `docs/hld/10-bindings-spec.md:2942` and the final checked-cache paragraphs. F281's Word chart export contract remains at `docs/hld/10-bindings-spec.md:4695`, and its presentation chart projection contract remains at `docs/hld/10-bindings-spec.md:4729`. Neither substitutes chart projection for field cache authoring, changes Python strings into tuples or adds a WASM export method.

## Limits and handback

The current F281 contract retains all eight acceptance rows and final obligations at `.claude/plans/F-X281-design.md:274`. Current F272 text explicitly retains open C1 work at `.claude/plans/F-X272-design.md:3880`. Preserving those carry bytes is not approval of the unreviewed C1 remainder, original scheduling, aggregate allowances, mutator propagation or a provider checkpoint.

No builds, tests, native/oracle runs or final gates were performed here. Root's separate composition commands are not claimed as this review's observations. The other 101 byte-exact restored carry files and the full F272 implementation are outside this bounded three-file semantic reconciliation. No source, plan, tracker or test was changed. Return the report lease to Root and stop.
