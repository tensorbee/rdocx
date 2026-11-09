# F-X280a, all aspects, pass 1

**Review kind**: identified GLM (opencode host) self-review of its own working
diff, using the `/microscope` aspects. This is explicitly NOT an independent
review and does not substitute for Root's pending independent microscope.

**Reviewed**: uncommitted working tree against Base
`f4616029f19e12edc0b808670cff1db5e6baa9f0`. 4 files, 454 insertions,
5 deletions (`git diff --stat`): crates/rpptx/src/lib.rs, 
crates/rpptx/tests/integration.rs, crates/rpptx-py/tests/
test_documented_examples.py, docs/hld/10-bindings-spec.md. Untracked local
build artifacts excluded.

**Verdict**: 0 defects, 0 smells, 3 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

- `crates/rpptx-py/python/rpptx/_rpptx.abi3.so`, untracked: the local
  editable wheel producer leaves its extension artifact inside the worktree.
  Not for commit, noted in the progress record.
- `crates/rpptx/src/lib.rs:22871`: the `sized_transform` doc comment states
  the refusal but not the exact message contract. The tests pin the exact
  string, so this is taste only.
- `crates/rpptx/tests/integration.rs:37565`: the acceptance test asserts
  exact reopened extents rather than a full byte-equality control against
  old-source output. The transform builder is the same function on the
  accepted path, so byte identity holds by construction.

## Aspect results

### correctness

`crates/rpptx/src/lib.rs:22871` refuses only strictly negative dimensions
(`width.0 < 0 || height.0 < 0`), accepts zero and negative positions, and
formats raw i64 operands exactly as the plan specifies. The three call-site
replacements (lib.rs:22490, 22508, 22529) propagate with `?` before shape
construction. `positioned_transform` is not replaced globally: its remaining
callers are picture, chart and table authoring plus one read-side
normalisation (lib.rs:11766, 13378, 13933, 22584, 26031), all outside the
approved lease. No finding.

### contract

The diff implements exactly the approved approach: one private helper, three
call sites, no public API, no table/connector change, build-before-append
ordering retained. HLD impact list executed as `docs/hld/10-bindings-spec.md`
only, touching the two ShapeCollection authoring/atomic-publication
paragraphs (lines 1022 and 1044 area). Word and OfficeMath sections
untouched. A-ID coverage: 7a `issue435_new_shapes_refuse_a_negative_size`
(integration.rs:37382), 7b lib.rs:43814 on the actual `add_unstyled_shape`
route with a no-style assertion, 7c integration.rs:37565 with save/reopen and
`validate()` empty, 7d test_documented_examples.py:12771 against the actual
installed child-source wheel, 7e integration.rs:37638 plus the existing
table control (integration.rs:16425) and reversed-endpoint connector control
(integration.rs:20782). No finding.

### panics

No new `unwrap`, `expect`, indexing or arithmetic in production code. The
guard does integer comparisons and a `format!` only. Test-side `unwrap`s are
in tests that also assert refusal. No finding.

### ooxml

The accepted path calls the unchanged `positioned_transform`, so serialized
`a:xfrm` bytes are unchanged for zero/positive authoring. Refusal produces no
element, relationship or part. Child order and unmodelled subtrees are
untouched. No finding.

### tests

Old-source RED was executed on the real Base before any production edit and
was semantic, not a compile or setup failure: integration panicked at
"negative textbox width was accepted" (integration.rs:37394 on old source),
lib panicked at "negative unstyled width was accepted" (lib.rs:43803 on old
source), and the installed wheel failed with "DID NOT RAISE RpptxError".
The same commands are green after the change. The two acceptance tests were
green controls on old source by design and remain green, which is what
distinguishes a guard that over-refuses. Atomicity is asserted by complete
package byte equality after every refusal, including on the candidate
transaction path forced by an existing connector and on the populated group
path. Python live-handle checks use actual public handles (collection, shape,
text frame paragraph) and reuse the collection after refusal. No finding.

### structure

One private function beside its consumer in an existing file. No new file,
module, trait, generic parameter, feature flag or crate. No finding.

## Not found

- No missed publisher: no other textbox/preset authoring route exists outside
  `textbox_member`, `preset_member` and `unstyled_preset_member`, and the
  Python wrappers call exactly those through `ShapesMut`.
- No scope broadening into validation, parsing, resizing, layout or Word.
- No defect in the refusal-to-RpptxError mapping: `InvalidShapeMutation`
  falls through the existing `_ => "RpptxError"` arm (rpptx-py lib.rs:128).
- No hash-harness or final-union check was run by this worker. Those remain
  pending final riders.
