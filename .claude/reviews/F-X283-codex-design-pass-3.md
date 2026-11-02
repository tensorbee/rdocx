# F-X283 independent design peer 3, Codex

Verdict: retain draft status. One high design defect remains, stable ID F9.
Counts: 1 defect, 0 smells, 0 nits. This is the third actual design peer after
Grok 1 and DeepSeek 2, using the authorized Codex takeover. No counter is reset.
Root final design review remains required.

## Scope and provenance

Reviewed only `/private/tmp/rdocx-s96-f283-review-pass3`, detached at
`aff3f2b8744ed86755db8520c9bf6abeed5f2038`. All 106 overlays in
`F283_DESIGN_PEER3_SOURCE_PINS.json` authenticated unchanged. Manifest SHA-256:
`7677ff67d702f34bb90249e8dc640382647f646b410cc1e2aceba95327aa6603`.
The complete 735-line draft plan SHA-256 is
`06f89b183d98b7f303ad3af8ee2e24471f45f046c823cab82647aff50c65bcab`.

Inspected the plan, prior design reports and their adjudications, repository
workflow and microscope requirements, and the source mechanisms for local cell
editing, native staging, source certificates, owner/piece preparation, public
binding factories, replacement, scope initialization and serializers. Compared
the retained workload to the pinned Issue445 body. Read the existing historical
missing-link probe manifest and outcomes without rerunning it. Historical rlibs
are not current native or binding evidence. The canonical working tree was not
used. No source edits, tests, builds, benchmark, network, Git mutation, cleanup
or delegation occurred. This report is the sole write.

## Finding

### F9, high defect: activation inventory still misses an earlier public owner lookup

The proposed dormant state has empty live owners/indexes at
`.claude/plans/F-X283-design.md:323`. The activation inventory at
`.claude/plans/F-X283-design.md:414` activates inside `prepare_word_revision`
and explicitly handles the earlier `plan_body_splice` lookup. Its claim that
other scoped mutations reach the activated publication route does not cover
`remove_hyperlink`.

Concrete supported input: open a valid document with one internal-anchor
hyperlink, then construct a plain public `Story` and `Hyperlink` matching that
link and call `Document.remove_hyperlink` before acquiring any checked handle.
This is an actual supported record path, not an invented stale handle:
`crates/rdocx-py/src/document.rs:1095` creates an unbound Story,
`:1337` creates a Hyperlink with no owner, `:17289` accepts the unbound Story,
and `:11598` resolves the uniquely matching link from the native document.

The current removal then reads `self.owners` at
`crates/rdocx-py/src/document.rs:15469`, before the call to
`publish_owner_candidate` at `:15484`. With the proposed dormant empty vector,
a valid removal fails with `hyperlink has no checked physical source owner`
at `:15482`, before the planned activation can run. Current eager construction
supplies that owner. This is a source-derived design counterexample, not an
executed regression against an implementation that does not yet exist.

Expected: preserve the current valid unbound-record removal and its error
policy. Explicitly place activation before this earlier owner lookup, retaining
appropriate input-validation/refusal ordering, and include a fresh-open unbound
record control that performs no factory access first. The broader ingress audit
must also give the adjacent direct native `set_hyperlink_url` publisher at
`crates/rdocx-py/src/document.rs:15448` an explicit dormant-facts policy rather
than assuming all positive edits enter `prepare_word_revision`. No separate
executed stale-cache failure is claimed for that adjacent setter.

This retains F9 because the invariant is the same under-enumerated activation
boundary. It does not revive the prior raw capture count as an ingress count.
The generic instruction to audit more during implementation at plan line 426
does not settle this concrete design boundary.

## Prior finding dispositions

F1 to F8, F10 to F12, F14, F16 and F17 now have concrete proposed mechanisms:
singleton-body validation, distinct body/table ordinals, selected-table child
indexes and deltas, a prepared piece replacement, in-place native commit,
source-bound namespace qualification, consuming replacement dispatch, retained
fallible inventory, whole-cell exclusions, separate text/property retirement,
exact saved-byte controls, optional authority and explicit checked fallbacks.
These address the prior design gaps. Their implementation remains unproved.
F9 is partially repaired by the named mutable factories and six splice callers,
but remains open for the input above.

F13 was not an established duplicate-scope call in the old plan. The revision
explicitly excludes the automatically retired document scope. F15's original
registration citation did not prove its claim. The required F268 79/252/407
union riders are nevertheless explicitly retained with corrected provenance.
S1 to S6 and N1/N2 have concrete dispositions: one consuming clone, separate
release/debug evidence, missing cache/route fallback, existing caller errors,
signed wrap controls and corrected assignment/error contracts. No independent
residual smell or nit was established in this pass.

## Acceptance and mechanism assessment

| A-ID | Design assessment |
|---|---|
| A283-1 | Exact text setter workload and thresholds retained. Local native and binding cost mechanism is specified, pending implementation and measurement. |
| A283-2 | Alignment uses the same selected-table packet and held-handle policy. Exact workload retained. |
| A283-3 | Shading uses the same local preparation, with absent/empty/nonempty properties represented. Exact workload retained. |
| A283-4 | Python open/replace/save versus CLI criterion retained. Consuming replacement avoids the named duplicate clone. No timing success claimed. |
| A283-5 | Actual paragraph setter route and direct-cell ordinal conversion are specified. Separate body-independent measurement retained. |
| A283-6 | Existing negative width, truncation and signed i32 wrap behavior is explicitly pinned. |
| A283-7 | Existing border semantics and errors retained, with candidate-before-publication ordering. |
| A283-8 | Margin refusal and valid controls remain atomic obligations. |
| A283-9 | Property relocation and CellText retirement are separate, with exact route and index deltas. |
| A283-10 | Fallible reservations and packet preparation precede the last fallible revision commit. Injected refusal controls retained. |
| A283-11 | Existing comment, anchor, note and revision outcomes are specified without fictional preservation. Local versus checked package controls retained. |
| A283-12 | Parsed unknown XML, namespace qualification, schema order and exact local/fallback saved-byte checks remain required. |
| A283-13 | Retained eager fallible validation, nonzero-serial initialization, lazy authority and serializer invalidation are concrete. F9 still blocks complete lifecycle closure. |

Case 1 retains 0/500/2000 paragraphs, a 20 by 5 table, 100 reused cells,
text then alignment then shading, no warmup, and the original absolute and
relative limits. Case 2 retains the 4000-paragraph, 1000-match workload, the
inside-process Python timer before open, the CLI subprocess timer, three paired
trials, lexicographic selection of the Python total/stage tuple and minimum CLI
command time. Final artifacts remain matching Linux x86_64 release wheel and
CLI. Neither workloads nor thresholds were relaxed.

The source certificate is tied to actual main-part Arc identity and publisher
invalidation. It does not assume anchors alone prove namespace safety. The
shared initializer amendment reserves and prepares before initialization and
retains outstanding epoch/preparation refusal checks. It does not reset origin
or serial. Dormant save invalidation correctly accounts for native model
normalization and external stream failure. The three hidden published facade
bridges remain an explicit additive API and archive review obligation. No new
F282 shared-part API dependency is required by this proposal.

Correctness and contract have the F9 finding above. No independent panic,
OOXML, test-design or structural defect was established. Candidate-versus-live
publication, selected-cell fallback, XML preservation and budget mechanisms
remain implementation review targets, not passed evidence.

## Pending evidence and boundary

This is design review only. All actual RED/GREEN implementation evidence,
current native/Python missing-link reproduction, installed Python 3.9/3.12,
typing/stubs, exact release performance, native/schema controls, final union
riders, hash49, WASM, archives and hosted CI remain pending. No native parity,
performance success, full feature acceptance or release approval is granted.

Report written. Review stopped and frozen-worktree writer lease returned.
