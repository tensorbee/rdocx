# F-X272 Issue430 design, Codex final 9

TARGETED REPAIR ACCEPTED. Zero outstanding defects, zero smells and zero
nitpicks in the V11 to V12 remediation. Extension approval and source
implementation remain withheld for the pre-approval evidence below.

The author naturally exited0 on terminal26294. V12 is127420bytes and564lines,
SHA-256 37dafbfd13f2ec24c4ed2f3fea55e8aad0f363367ce8100168e47fd360ea50ad.
Its1903-byte matching handoff and the proposal are frozen at
/private/tmp/rdocx-cross-review/issue430-v12-final-frozen.
V11 and its handoff retain the hashes in final8. This is Root's direct
remediation review, not a fourth peer or a reset of the three consumed design
passes.

## Findings disposition

N15 is closed again at the design level.
/private/tmp/rdocx-deepseek-harness/issue430-contract/DEEPSEEK_PROPOSAL_V12.md:311
now removes the added normal definition and reference on reject, keeps the
original special-only part with N0, R0, Efalse and base1, and retires markers.
The accept companion retains the added definition and reference. The1/1
direction row at V12:205 is consistent with both addition and deletion.

N21 is closed at V12:346. The CLI regression command uses the existing
integration carrier. The native regression_test and integration_test carriers
are unchanged, and no new test binary is proposed.

S1 is closed at V12:560. Scope fullness is decided by each selector's own
predicate. Different ids do not preclude a shared author or date from selecting
all elements. The partial-unit refusal remains atomic and no token API or new
user choice is requested.

## Fidelity and limits

Root consumed the full V11 in final8 and inspected the entire seven-hunk V12
diff. Exactly45 acceptance rows and14 gate rows remain. A430-18 is the only
changed acceptance row. Other edits are the targeted direction row, CLI command,
selector summary and version/disposition metadata. The N19/N20 lifetime
mechanism and the earlier user decisions remain unchanged.

The actual diff has19added and10removed lines including blank lines. The
handoff's16added and9removed numbers count nonempty lines, matching the author's
awk predicate. They are not complete physical line counts.

The author wrote /private/tmp/v12.diff despite the explicit two-output and
stdout-only task restriction. That is a recorded protocol deviation, not a
new product design finding. Root authenticated its21189bytes and hash
35dbfc75658509512c882262655d09a264716a5ab96012976af745790fca09a8,
preserved it in the frozen packet and removed that exact unchanged task-only
file after the author exited. No other cleanup occurred. This does not erase
the deviation or establish general instruction compliance.

No additional correctness, schema-order, namespace, atomicity, public-surface,
structure or test-carrier defect is found in this targeted repair. No product
test or native observation is newly qualified by this design review.

## Remaining approval evidence

V12:427 still declares G430-02, including preservation when another reference
remains, as pre-approval and records only sole-reference evidence. The existing
selected-reference control proves removal of a sole referenced owner while
leaving an unrelated insertion, not preservation of a shared owner. Complete
that bounded retained-reference control before approving the extension, or
explicitly reconcile its timing against the canonical contract. Do not silently
mark this gate passed from the sole-reference observation.

Full native/schema composition, library selective resolution, mixed lifetimes,
absent-part cleanup, installed bindings, hash harness and CI remain final
qualification debt. The C1 source-admission mechanism is separately unclosed.
No F-X272 source lease, provider completion, sprint acceptance or release
approval follows. STOP and return to orchestration.
