# F-X280a, all aspects, pass 2

**Review kind**: identified GLM (opencode host) self-review of its own D1/D2
remediation, using the `/microscope` aspects. This is explicitly NOT an
independent review and does not substitute for Root's pending final2.

**Reviewed**: the remediation delta on top of the pass-1 state, against
`.claude/reviews/F-X280a-codex-final-1.md` as the repair contract.
`git diff` against Base f4616029f19e12edc0b808670cff1db5e6baa9f0 remains four
paths. This remediation changed only `crates/rpptx/tests/integration.rs`
(one appended test), `crates/rpptx-py/tests/test_documented_examples.py`
(one rewritten test) and evidence records. `crates/rpptx/src/lib.rs` and
`docs/hld/10-bindings-spec.md` are byte-identical to their pass-1 state
(lib.rs diff still exactly the helper, three call-site replacements and one
unit test).

**Verdict**: 0 defects, 0 smells, 2 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

- `crates/rpptx-py/tests/test_documented_examples.py:12829`: 
  `held_picture.width is not None` is a weaker liveness proof than the
  member/collection reuse assertions beside it. Acceptable because picture
  authority is not the story's discriminator.
- `crates/rpptx/tests/integration.rs` (populated-group test): 
  `assert_eq!(seeded_ids.len(), 3)` duplicates the preceding
  `assert_eq!(original_count, 3)` in strength. Harmless.

## Defect-by-defect contract verification

### D1, populated-group and companion-preservation cohort

The new native discriminator
`issue435_refusals_preserve_populated_group_companions`
(crates/rpptx/tests/integration.rs:37704 area) seeds all three required
companions before any refusal: an existing ordinary textbox, a
relationship-backed picture with its own media part and relationship
(integration.rs:37711), and an opaque package part `/ppt/opaque.bin`
injected through the public OPC layer and reopened
(integration.rs:37738). The group at :37725 is POPULATED with an existing
member before the refused additions. After every failed call the test
asserts complete package byte equality (integration.rs:37816, :37841 and
the root-level case), the unchanged shape count and an unchanged id vector
parsed from the slide XML, which covers the existing member's preserved
geometry, the picture relationship and the opaque part without relying on
any single proxy. Subsequent successful use of the same collection and
owners is exercised, then save/reopen asserts `validate()` empty, the exact
opaque part bytes (integration.rs:37869), the seeded id prefix, the group
member count of exactly two and the seeded member's exact geometry. The
empty-group, local-append and connector-forced controls in
`issue435_new_shapes_refuse_a_negative_size` are retained unchanged.

The Python fixture (test_documented_examples.py:12686 area) now seeds a held
picture companion and a populated group with a held member and member-text
handle, and asserts `deck.to_bytes() == before` after EACH of the four
refusals (:12804, :12810, :12697 area and the fourth), not once at the end.
Held group-member and text handles are actually exercised after the errors
(:12826 add_run and reads) and the held group collection performs a
successful subsequent add (:12832), followed by a reopen asserting member
text, member width, new member width and absence of refused members.

The focused populated-group discriminator was executed against the actual
old producer, not source-traced. Native: identical final test source run on
a scratch extraction of the claim Base gave checked child exit 101 with the
semantic panic "negative shape extents got -1000 by 2000 EMU were accepted
in a populated group" (receipts/oldsrc-red-populated-group.log, old test
binary sha256 9d28818c...). Python: a genuine old-source extension
(sha256 912e81ae..., distinct from the new 1f8cdc0d...) built with maturin
from the extraction shows the old producer ACCEPTED the populated-group
negative add (probe receipt), and the rewritten test fails against it with
checked exit 1, "DID NOT RAISE RpptxError" (receipts/
oldsrc-red-python-issue435.log). No hidden old error existed to document.
the honest finding is unconditional old acceptance. An earlier uv-cache
attempt that served the new artifact under the old path was discarded when
its hash matched the new build, before any evidence was claimed from it.

### D2, checked exits and provenance

Every final-source check now records a separately captured child status with
the whole log retained under .claude/scratch/receipts/ and no pipe between
producer and filter: fmt exit 0 (after one repaired exit-1 reflow, preserved
in fmt-check.log), clippy exit 0, lib test exit 0 with 1 passed, native
integration exit 0 with 4 passed bound to binary
integration-63d3fc61d54bd197 sha256 03c7b198... (profile dev, default
features), and the installed Python test exit 0 bound to the actual imported
extension path crates/rpptx-py/python/rpptx/_rpptx.abi3.so sha256
1f8cdc0d..., unchanged by this test-only remediation so no rebuild was
claimed. Source identity for all records is HEAD f4616029... plus the dirty
diff hash d14deb41... recorded in the progress file. The historical
pass-1 records are relabelled as wrapper-exit-0 observations with their
visible semantic failures and passes retained verbatim. Nothing was erased
and no wrapper outcome was promoted to a checked child exit. The two
fixture mistakes made during remediation (ppt/media OrphanMedia seeding,
original_count + 2 miscount) are recorded and neither reached a claim.

## Aspect results

### correctness

The refusal predicate, call sites and messages are unchanged from pass 1 and
Root found no production defect there. The new assertions compare complete
package bytes, which subsume relationships, parts and ids. The explicit id
vector and member-geometry reads make the intent inspectable. The group
member adds no top-level child, and the final count assertion now states
`original_count + 1`. No finding.

### contract

The fixtures now match the plan's atomic-publication matrix row and the
Root D1 repair list exactly: populated group, relationship-backed member,
opaque part, per-failure unchanged state, held-handle reuse and an actual
old-producer discriminator. No public API, no source-file or module
addition, no matrix weakening. The retained controls were not removed. No
finding.

### panics

No new production code. Test panics are explicit `panic!` with the refused
input named, and `expect`s sit on seeded state whose absence is a fixture
bug. No finding.

### ooxml

The opaque part survives refusals and the final save verbatim
(OPAQUE-SEED bytes asserted after reopen). Existing member XML, picture
relationship and id assignment are unchanged through refusals. No finding.

### tests

Would the new tests fail if the guard were reverted? Yes, observed, not
inferred: the identical source fails on the old producer natively (exit 101,
populated-group acceptance panic) and through the installed old extension
(exit 1, DID NOT RAISE), and passes on the new source (exit 0 both). The
RED/GREEN pairing now covers the populated-group route that pass 1 lacked.
No finding.

### structure

Test-only delta plus evidence records. No new file, module, trait, generic
or crate in the product. Scratch extraction, venvs and logs live in
gitignored .claude/scratch/ and are labelled not-for-commit. No finding.

## Not found

- No defect in the retained pass-1 tests. They were re-run green post-fmt in
  the same checked batch (exit 0, 4 passed) and were not weakened.
- No missing per-failure assertion in either language.
- No status-masking pipeline remains in any record presented as current
  evidence.
- No hash-harness, workspace-union, final installed Python/typing, WASM,
  archive, SemVer or hosted-CI check was run. All remain pending final
  riders for the integrated union.
