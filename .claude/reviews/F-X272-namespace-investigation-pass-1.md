# F-X272, namespace investigation, pass 1

**Role**: independent Codex review and design investigation, read-only outside this review record.
**Source**: dirty `work/f-x272-codex` at Base `c8acc502b74f29ed7e9546939f17007709a9f22b`, 13 tracked files with 1323 added and 89 removed lines when inspection began. The source, tests, plan and progress record were not edited by this investigator.
**Outcome**: namespace invariant remains unresolved. This is not a clean feature microscope or final acceptance.

## Observed evidence

The actual first fixed-source batch passed 7 of 8 cases and failed the aliased main-root case with `serialized document has no w:body start`. After the first namespace repair, the second batch again passed 7 of 8. Its alias no-op succeeded, but changing the ordinary aliased paragraph failed with `comparison could not correlate paragraph run owner 0: 1 physical runs from cursor 0, 0 available`. Logs are `/private/tmp/rdocx-s96-run-refresh/fx272-native-green-attempt-1.log` and `fx272-native-green-attempt-2.log`. The prior baseline is retained at `fx272-native-red.log`.

The canonical table/control, footer/endnote/text-box and ignored-story controls in those actual batches passed. That evidence qualifies those cases, not arbitrary aliased descendants. No compilation, new test execution, native Word observation or parity claim was performed during this investigation.

## Findings

### N1, descendant source views discard the namespace facts that qualified their owners

`crates/rdocx/src/comparison.rs:3782`
`crates/rdocx/src/comparison.rs:4380`
`crates/rdocx/src/comparison.rs:9535`
`crates/rdocx/src/comparison.rs:9606`

Main-story parsing now obtains effective document/body bindings and correctly parses the aliased typed owners. The dispatch then passes a bare owner substring through `Option<&str>`. A paragraph such as `<q:p><q:r>...</q:r></q:p>` no longer carries the root declaration that gave q its Word meaning. `paragraph_run_spans` reconstructs a scope containing only canonical w, so its namespace-qualified child check finds zero Word runs. The typed paragraph still has one modeled run. This is the direct explanation of the actual second failure. Replacing only the wrapper's root-body assumption cannot establish namespace correctness for its descendants.

The required invariant is: every source fragment used for physical owner selection or mutation retains its actual effective namespace scope, including local overrides, and selection checks expanded child and parent identity. Source byte ranges remain ranges into the original fragment. Generated markup and preserved producer markup must retain their separate namespace authority.

### N2, the original alias reproduction was replaced by the foreign-body control

`crates/rdocx/tests/regression_test.rs:408758`
`crates/rdocx/tests/regression_test.rs:408762`

Every alias case now inserts the foreign `x:body` before the actual `q:body`. The original plain full-q-root case is no longer independently exercised. Restore that exact original input and add a separately labelled foreign-body control. Preserve both old failure identities rather than treating the changed fixture as an unchanged successor. This agrees with the root handoff's explicit preservation warning.

## Relevant path inventory

| Boundary | Current authority and remaining risk | Citation |
|---|---|---|
| Main document extraction/replacement | `main_body_source` selects the direct expanded Word body and obtains its effective scope. Keep this mechanism and the existing self-closing-body refusal boundary | `crates/rdocx/src/comparison.rs:9333`, `crates/rdocx/src/comparison.rs:9398` |
| Held revision scan | `HeldRevisions::of` uses the namespace-bound `story_wrapper` before the existing revision scanner. This is the right pattern. The low-level revision tree already maintains element scopes, so no new revision parser is needed | `crates/rdocx/src/comparison.rs:2729`, `crates/rdocx/src/comparison.rs:2753`, `crates/rdocx/src/revision.rs:495` |
| Owner-to-descendant dispatch | Bare slices remove the inherited bindings. The same loss recurs when cell children and control payload are sliced | `crates/rdocx/src/comparison.rs:3782`, `crates/rdocx/src/comparison.rs:6886`, `crates/rdocx/src/comparison.rs:7090` |
| Direct child scanners | The canonical helper compares literal `w:local`. The prefix-aware variant uses namespace resolution but has no inherited-scope argument, so it is only sufficient on self-contained source | `crates/rdocx/src/comparison.rs:9491`, `crates/rdocx/src/comparison.rs:3130` |
| Paragraph physical runs | The wrapper binds only w. Parent selection stores only local names, so a Word r under a foreign p or hyperlink lookalike can also be mistaken for a modeled run | `crates/rdocx/src/comparison.rs:9536`, `crates/rdocx/src/comparison.rs:9554` |
| Complex-field physical counts | Both detached source and source-replacement counting create a synthetic w paragraph without carrying the original field source scope | `crates/rdocx/src/comparison.rs:9615`, `crates/rdocx/src/comparison.rs:9640` |
| Paragraph end boundary | Granular insertion at a missing next run uses a literal closing w:p, even for a retained q paragraph shell | `crates/rdocx/src/comparison.rs:5419`, `crates/rdocx/src/comparison.rs:9530` |
| Tables and cells | Raw-source table property/row/cell selection uses canonical helpers. Cell modeled-child selection and its expected-name comparison also require literal names. A trailing inserted row uses a literal table close | `crates/rdocx/src/comparison.rs:6489`, `crates/rdocx/src/comparison.rs:6502`, `crates/rdocx/src/comparison.rs:6520`, `crates/rdocx/src/comparison.rs:6800`, `crates/rdocx/src/comparison.rs:9687`, `crates/rdocx/src/comparison.rs:9721` |
| Controls | Both inline and block paths locate literal sdtContent and replace it through a literal name. The block path seeds inherited context with canonical w instead of its actual parent scope. Run-only slicing then repeats the loss | `crates/rdocx/src/comparison.rs:6981`, `crates/rdocx/src/comparison.rs:7013`, `crates/rdocx/src/comparison.rs:7019`, `crates/rdocx/src/comparison.rs:7084`, `crates/rdocx/src/comparison.rs:7093`, `crates/rdocx/src/comparison.rs:7490` |
| Generated Word names | Whole-part missing-binding closure helps when w is absent. It cannot turn generated w elements into Word elements when the producer already binds w to another namespace, and it must not overwrite that producer binding | `crates/rdocx/src/comparison.rs:982`, `crates/rdocx/src/comparison.rs:3014` |
| Existing scope mechanisms | Reuse the inherited-plus-local declaration merge and existing scoped fragment closure. The latter adds missing declarations only and preserves existing bindings | `crates/rdocx/src/comparison.rs:3026`, `crates/rdocx/src/document.rs:13109`, `crates/rdocx/src/document.rs:13895` |

## Smallest robust repair recommendation

1. Keep one scope-bearing source route through the existing comparison file. Extend the existing source argument with its effective binding slice, or pass that slice explicitly beside the source. The current `compare_body` source tuple already carries bytes and spans, so adding scope there is a small internal contract correction. Thread it through the source-capable paragraph/table/row/cell/control paths. Merge each child's local declarations with its actual parent scope before scanning descendants. Do not guess a Word binding from the root prefix or from a local name.
2. Consolidate source-capable direct-child scanning on expanded names using `NsReader` plus the existing collision-safe `story_wrapper`. Translate offsets back to the untouched fragment, and account for the wrapper's extra depth. Reuse the resulting qualified boundaries for replacement and closing positions. Merely changing a Reader to NsReader without supplying scope will preserve N1. Reuse one direct-child policy rather than keeping canonical and alias implementations with different semantics.
3. Make paragraph physical-run selection verify expanded Word identity for both the child and its p/hyperlink parent. Carry source scope into complex-field span counting and into inline-control run selection. Keep the existing modeled-to-physical count checks, refusal and atomic publication. Do not weaken those checks to accept a missing run inventory.
4. Use source-reader closing offsets for retained paragraph, table and control shells. Keep literal w closing names in functions whose input is exclusively fresh canonical serialization. For example, `paragraph_owned_run_xml` constructs its paragraph through `paragraph_xml` before the literal close at `comparison.rs:9221`. This generated-only path does not justify a broad rewrite of all literal names. Audit source-capable `append_word_child` callers separately at `comparison.rs:9435`.
5. Preserve unchanged revision-bearing owners as the original source bytes. Namespace wrappers used only for analysis must not leak into output or normalize these owners. For newly generated Word fragments, provide a local correct canonical binding when the producing context needs it, while retaining the actual namespace of transported producer children. Do not globally rebind an existing foreign w, strip producer declarations or canonicalize the entire document to hide alias selection defects. The existing whole-part missing-binding closure can remain a missing-only operation.
6. Amend the existing plan's namespace clarification and consumer map with this descendant inventory before repair. These are private source-context and qualified-selection changes within approved A2/C2 preservation scope. No new public API, trait, crate, module, file or unrestricted revision-bearing-owner edit is required. Keep the user's bounded refusal/workaround decision unchanged. No material user scope decision is identified by this investigation.

## Discriminating successor cases

Each new case must obtain an actual pre-repair RED before being credited. The cases below are source-traced recommendations, not executions by this investigator.

- Restore the exact plain q-root reproduction with the unchanged revised table and changed ordinary q paragraph. Keep the foreign x:body lookalike in a separate case with its own label. For both, assert byte-identical no-op, retained old revision bytes, new author/id counts and accept/reject reconstruction.
- Add a canonical document/body with two Word aliases inherited only from an ancestor: q for the changed paragraph and z for its runs. Keep the prior revision in a separate unchanged owner. This rejects a repair that guesses only the paragraph's prefix or transports only the document's preferred prefix.
- Add a separate empty aliased ordinary paragraph that receives text at word granularity, again beside a distinct unchanged revised owner. This reaches the granular end-boundary path even when there is no original run. It distinguishes a correct source closing offset from a repair limited to run discovery.
- Add a self-contained Word q paragraph with an opaque foreign x:p subtree containing a Word q:r, plus one real direct q:r. Change only the real direct run. The parser preserves unknown children through `rdocx-oxml/src/text.rs:10230`, while the current scanner tests only its parent's local name at `comparison.rs:9555`. The opaque nested run must not enter the physical modeled inventory, and its complete source must survive. This is a structurally different parent-identity discriminator for the same namespace invariant.
- Add unrevisioned aliased table and block-control owners whose nested text changes, with old revisions in a separate owner. Exercise inherited aliases through tr/tc/p and sdtContent. Putting the old revision inside the changed table/control would instead test the deliberately unsupported bounded-edit case and cannot qualify descendant alias support.
- Add q-bound Word source with w bound to a producer namespace and an unchanged opaque w child outside new revision wrappers. Verify generated revisions resolve as Word while both the old revision and opaque child retain their original bytes and expanded identity. This rejects globally rebinding w or treating every declared w as the generated Word binding.
- Include one aliased complex-field physical-owner case and a related-story alias case after the main regressions. Retain existing ignored-story/text-box controls and late atomic refusal. These are bounded cross-path checks, not an exhaustive namespace Cartesian product.

## Provenance and remaining work

During investigation the parent reported a fresh PR 427 rebase to head `c1e261ca63ebda5b7da1adec886ee40364a4e3de`, base `24af95d3`, with the Issue 425 body unchanged. The parent is authenticating the exact successor patch separately. This report analyzes only the preserved dirty worker tree and does not claim to have adopted or reviewed that new patch.

Pending: canonical plan/source reconciliation, the actual repaired scoped compile, genuine RED/GREEN for the preserved and distinct discriminators, focused CLI cases, a fresh authenticated Python build with held/fresh owner checks, then an independent complete feature microscope. Actual native Word, installed full suites, typing/stubs, archive/readme and final hash/pixel/risk qualification remain final sprint debt. No source edits, compilation, commit, push, final gate or clean feature verdict occurred here.
