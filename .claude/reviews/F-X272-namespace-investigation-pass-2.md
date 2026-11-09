# F-X272, namespace investigation, pass 2

**Role**: independent Codex design and reviewer escalation. Read-only outside this report.
**Reviewed**: dirty `work/f-x272-codex` at Base `c8acc502b74f29ed7e9546939f17007709a9f22b`, 13 tracked files, 1992 added and 308 removed lines. Tracked binary diff SHA-256 `9f221badb9d98adc154ca8ef8bc790c34c171673bde613bacbf66676e952b926` remained unchanged during this investigation.
**Verdict**: 1 defect, 0 smells, 0 nitpicks in this focused escalation. N1 remains unresolved. This is not a complete feature microscope or acceptance verdict.

## Evidence and contract

F-X272 remains in-progress in CURRENT_SPRINT and BACKLOG. The approved plan's A2/A3 and C2 require qualified source preservation and reconstruction after accepting or rejecting only the new revisions. HLD03 describes transport of source namespace scope and locally qualified generated Word markup at `docs/hld/03-architecture.md:1431`. The baseline reconstruction condition remains explicit at `docs/hld/10-bindings-spec.md:2870`. The bounded revised-owner refusal and checked workaround decision remains unchanged.

Consumed actual execution evidence: `/private/tmp/rdocx-s96-run-refresh/fx272-native-final-scope-green-2.log` records 12 executed tests, 11 passed and 1 failed, in 3.70 seconds after a successful 42.76-second build. `issue425_comparison_keeps_inherited_scope_for_final_paragraph_marks` fails for `foreign_w=true, inserted=false` with `comparison acceptance does not reproduce the edited stories at body story item[1]`. Its exact input and dispatch are at `crates/rdocx/tests/regression_test.rs:409082` and `crates/rdocx/tests/regression_test.rs:409094`. The preceding duplicate-property and compile-only failures remain separate historical observations. This investigator ran no tests, builds, oracle or wheel checks and changed no source, tests, plan or progress record.

## Defects

### N1, revision resolution loses source namespace scope when merging paragraphs

`crates/rdocx/src/revision.rs:821`
`crates/rdocx/src/revision.rs:1352`
`crates/rdocx/src/revision.rs:1358`
`crates/rdocx/src/revision.rs:1368`

This continues the inherited-scope invariant from investigation pass 1. The comparison repair now closes the unchanged q predecessor and its q properties before marking it at `comparison.rs:4330`. It also gives the generated final w paragraph a local Word binding through `metadata.ids.generated` at `comparison.rs:4470`. Its generated paragraph properties can inherit that binding, as `final_paragraph_properties_xml` returns its property XML at `comparison.rs:6274` rather than independently closing every property subtree. This is valid under that generated paragraph.

Acceptance then removes the predecessor's deleted paragraph mark and merges the paragraphs. `render_merged_paragraphs` keeps the FIRST paragraph's raw opening, takes the LAST paragraph's properties, and renders every paragraph's remaining children under the first opening. Both the selected properties and other transported children use `self.render`, which supplies an empty promoted scope at `revision.rs:571`. `render_ordinary` preserves only element-local declarations plus explicitly supplied promotion at `revision.rs:735`. The parser computed effective source scope at `revision.rs:510`, but `XmlElement` retains local declarations rather than that effective scope at `revision.rs:222` and `revision.rs:1424`.

Consequently, a generated w:pPr that meant Word under the final w:p can mean `urn:producer` after transport under the retained q:p. The destination's inherited w binding changes the expanded names of those properties and descendants. This explains the observed accepted-body property mismatch. Adding a declaration only to this generated pPr would hide this input while leaving transported runs, opaque descendants and other aliases exposed.

The merge also discards its caller's `promoted_namespaces` at `revision.rs:821`. Thus the first merged opening itself has no way to consume declarations from a removed enclosing revision wrapper, even though ordinary children consume that promotion at `revision.rs:827`. This second route is source-traced, not newly executed here, and belongs to the same N1 relocation invariant.

## Shared boundary and consumer inventory

| Consumer | Source and destination obligation | Citation |
|---|---|---|
| Comparison postcondition copies | Accept/reject only newly allocated ids while preserving typed baseline and old markup authority | `crates/rdocx/src/comparison.rs:1133`, `crates/rdocx/src/comparison.rs:1174` |
| Public main-story resolution | All, author, id, id-set and date scopes use the same tree renderer and staged publication | `crates/rdocx/src/revision.rs:293`, `crates/rdocx/src/revision.rs:319` |
| Related-story resolution | The same renderer also resolves header/footer/comment/note story XML | `crates/rdocx/src/revision.rs:389`, `crates/rdocx/src/revision.rs:403` |
| Paragraph merging | First shell, last properties, every retained child, chained paragraphs and crossing removed tables need source-qualified transport | `crates/rdocx/src/revision.rs:775`, `crates/rdocx/src/revision.rs:810`, `crates/rdocx/src/revision.rs:1343` |
| Kept revision unwrapping | Removed wrapper declarations must reach promoted roots without overriding descendant local shadowing | `crates/rdocx/src/revision.rs:657`, `crates/rdocx/src/revision.rs:662` |
| Property replacement | Prior properties and retained current children have different source ancestry and may be combined under a different prior root | `crates/rdocx/src/revision.rs:596`, `crates/rdocx/src/revision.rs:875`, `crates/rdocx/src/revision.rs:907` |
| Legacy mark and numbering retention | Existing promotion handling is part of the same caller audit, not permission to rewrite all revision XML | `crates/rdocx/src/revision.rs:979`, `crates/rdocx/src/revision.rs:1031`, `crates/rdocx/src/revision.rs:1059` |
| Discard-only validation | Rendering discarded content proves selected descendants but publishes no transported bytes | `crates/rdocx/src/revision.rs:1240`, `crates/rdocx/src/revision.rs:1329` |

Rows other than the observed paragraph merge are an enforcement audit inventory. They are not additional executed defects. The same shared fix must not change their original selection, property precedence, paragraph-mark rules, schema order, resolution counts or atomic failure behavior.

## Smallest robust repair recommendation

1. Extend the private resolution renderer's relocation boundary in the existing `revision.rs`. Retain or reconstruct each transported root's actual effective source namespace scope, then close it against its actual destination context. The existing parent links and local declaration records already identify source ancestry. There is no need for a new parser, public API, trait, crate, module or source file. Record this concrete internal extension and its public-resolution consumers in the existing approved plan before implementation. It remains an A2/A3 preservation clarification, not unrestricted revision-bearing-owner editing.
2. Carry the incoming promotion into the first merged paragraph shell, preserving that shell's own declarations and attributes. Qualify the last paragraph's properties and each retained child using THAT child's source ancestry. Do not put the last paragraph's namespace map on the whole merged paragraph, because that would rebind the first paragraph's content and opaque metadata.
3. Preserve local shadowing at every descendant. A source child-local declaration wins over inherited/promoted bindings. A changed destination default namespace may require an explicit empty default binding for a source subtree that originally had no default namespace. Never guess that a prefix, a local name or an undeclared name denotes Word. Never globally rebind the producer's w prefix. Preserve source names and opaque payload bytes, adding only namespace closure needed by actual relocation.
4. Apply one coherent source-context policy to byte-producing promotion and replacement paths in the table above. Ordinary traversal in an unchanged context should keep its existing bytes, rather than adding every inherited declaration to every element. Discard-only validation needs no relocated output. Do not weaken comparison postconditions, silently ignore duplicate properties, drop opaque children or turn the supported foreign-prefix case into a blanket refusal.

## Required discriminating checks after returning to implementation

Retain the exact observed final-mark input with q:pPr/q:rPr, x:note, canonical and foreign producer w, insertion and deletion. Keep the duplicate-property assertion, original plain alias reproduction, separate foreign-body control and all existing 11 successful controls. Require the actual failing foreign-w deletion to become GREEN.

Add one structurally different direct resolution control in the existing regression carrier: merge paragraphs whose moved properties or runs inherit bindings from their source paragraph, with a conflicting binding at the first paragraph, plus opaque child metadata and a child-local shadow. Check expanded element AND attribute names after author/id resolution and save/reopen, not just XML parseability or literal prefix counts. Include accept and reject paths with their appropriate deletion/insertion marks. A three-paragraph chain can combine differing source scopes without enumerating a Cartesian product. A related-story occurrence provides bounded evidence that the shared story renderer receives the same policy. Newly failing controls need actual pre-repair execution before being labelled RED.

Inspect a retained wrapper or property replacement case where local descendants shadow a promoted prefix. This challenges the ordinary recursive propagation at `revision.rs:710` and promotion merge at `revision.rs:1031`. This is an unexecuted challenge here. Preserve the existing effective precedence rather than blindly injecting an ancestor binding on every descendant.

Batch applicable existing tests with the affected comparison controls: `unwrapped_revisions_keep_wrapper_namespace_bindings` (`regression_test.rs:19911`), `rejected_property_changes_keep_owner_namespace_bindings` (`regression_test.rs:19924`), `contextual_paragraph_markers_merge_the_adjacent_paragraphs` (`regression_test.rs:20031`), `contextual_cleanup_preserves_foreign_property_shells` (`regression_test.rs:20000`), `paragraph_mark_removed_before_a_table_removes_an_emptied_paragraph` (`regression_test.rs:47116`) and `rejecting_paragraph_property_change_keeps_mark_and_section_properties` (`regression_test.rs:47284`). Check actual filter membership and preserve individual results. Complete suites and native oracles remain at final sprint union, not this repair checkpoint.

## Not found and remaining gates

No additional defect was established in comparison's repaired predecessor property selection or generated local Word binding during this focused pass. No new public contract, dependency, structural abstraction or material user scope decision is required. Zero independent panic, structure, smell or nitpick findings are reported. This is not an audit of all unrelated feature code.

N1 remains unresolved until implementation, applicable local checks and fresh independent review. N2 from pass 1 is not reintroduced: the plain alias and foreign-body test inputs are now separate. Fresh installed Python, CLI, scoped check/clippy, complete feature microscope and every named final native/hosted/hash/pixel/risk gate remain pending. Return control to root now. No remediation was performed in this pass.
