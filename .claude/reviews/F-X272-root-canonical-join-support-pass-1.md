# F-X272, bounded canonical integration support pass1

**Verdict**: 0 defects, 0 smells, 0 nitpicks in the requested semantic joins.
This is independent integration support, not a new full F272 peer cycle or
whole-provider acceptance. The original caps remain unchanged.

## Identity and scope

Read the immutable packet at
`/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/returned-reader-functional-batch/ROOT_TAKEOVER/root-canonical-join-1`.
All47joined source snapshots match JOINED_SOURCE.json. All72base/ours/canonical
operand pins across24delta paths match PREVIEW.json. RESOLVED.json records nine
explicit decisions. Its preview-state wording is historical preparation, while
JOINED_SOURCE.json explicitly identifies the later applied and checked snapshot.
Canonical delta is7d6a1d28742e94d949e54988376d65398d5787c7 from recovery
Base1d2c5a803d346da4a7b65b1bf8ac7e590ad1492c.

Relevant joined source SHA256 values:

| Path | SHA256 |
|---|---|
| crates/rdocx/src/document.rs | 14c47f7d1c4d39fc56931ecdcdf59cea4e5aa5fed4627cff4eeb8d64fcbf158e |
| crates/rdocx/src/comments.rs | 5b03c3ea2333ff2b4fd771fb82280f28e1d626d6208616609402d0507803f2d0 |
| crates/rdocx-oxml/src/comments.rs | 7a75ee83d09812db3878714266c361cea0aba4c025f1a71972a84aef78ef7ddd |
| crates/rdocx-oxml/src/footnotes.rs | 2c0037f9d1d3ab6f6a404e9a868f24a7139eb2163b548536e499f00d618cd5a8 |
| crates/rdocx-py/tests/test_core.py | 258821cce8bb60a081a04b962c27de0471781850a790f7f463ed74a434a89430 |

The approved F27246/7contract and F277b publication-only numbering, F279c
markup/namespace preservation, F279d crossing endpoint and F280b provisional
identifier/package policies govern this reconciliation. The existing readiness
checklist supplies preparation context, not authority to choose one whole file.

## Numbering and scoped preparation

The joined substantive numbering publisher saves the live numbering model,
identifier sets and paragraph numIds, temporarily canonicalizes them, invokes
the actual scoped document flush inside that lifetime, then restores live state
before returning the publication maps. Its closure result is retained through
restoration. This preserves F277b's package projection and C1's inherited writer
context together:
`/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/returned-reader-functional-batch/ROOT_TAKEOVER/root-canonical-join-1/joined/crates/rdocx/src/document.rs:28390`.

flush_to_package_in forwards its existing template context into that publisher
and retains the returned maps for the final live-to-published mapping fields.
canonicalize_typed_identifiers now performs scoped bookmark canonicalization
without an eager numbering remap. Hyperlink-style preparation follows successful
scoped Field preparation and precedes bundle preflight. The ordinary None route
remains available. These are the intended combined lifetimes and ordering at
joined `crates/rdocx/src/document.rs:24268`, `:27281`, `:27302`, `:28461` and
`:28656`. The four ordinary test-call adaptations pass None to the changed
private preparation method and introduce no product branch.

## Comments and note integration

D's Rust shims still route to optional initials, and staged reply creation calls
the actual projectable crossing-anchor helper before comment authoring. The
helper retains same-story/count/physical end-before-reference and pure-reference
payload admission. C2's empty-part cleanup still requires ownership, absence of
opaque payload and its URI-qualified narrow authored MC attributes:
`/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/returned-reader-functional-batch/ROOT_TAKEOVER/root-canonical-join-1/joined/crates/rdocx/src/comments.rs:1796`,
`:1831`, `:1901`, `:2197` and `:2228`.

The lower comments writer retains shared original destination binding checks,
the actual emitted root scope and paragraph_ids, then selects ordinary versus
inherited-work paragraph emission inside the same substantive writer. It does
not restore the old forced rebinding or replace the scoped writer with ordinary
paragraph byte wrapping:
`/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/returned-reader-functional-batch/ROOT_TAKEOVER/root-canonical-join-1/joined/crates/rdocx-oxml/src/comments.rs:186`,
`:266` and `:486`.

Qualified note id/type selection and canonical signed-normal classification
survive beside C1's actual reader outputs and emitted root scope. The merged
footnote reader uses is_word_attribute for the id/type and keeps omitted or
explicit normal notes at every signed id. The writer retains its scoped
paragraph branch and emitted root frame at joined
`crates/rdocx-oxml/src/footnotes.rs:136`, `:254` and `:270`.

## Canonical policies and Python union

The shared OPC source snapshot is byte-identical to canonical. Native raw
validation remains in the joined document source. Provisional bookmark counts
are observed, cloned and compared against provenance at staged publication,
while the scoped bookmark canonicalization context remains wired. Relevant
joined `crates/rdocx/src/document.rs` locations are5639,5809,6278,24268,27524
and87732. No second validator or imported blanket renumbering was introduced.

Static Python AST source-segment comparison found all four recovery-added or
changed top-level functions and all19canonical-added or changed top-level
functions preserved byte-for-byte in the joined test carrier. The three
recovery-added functions are also exact canonical substrings. The fourth is
the existing carrier's strict Field XML coverage. Thus the selected conflict
does not drop either test population. This checks source custody and membership,
not execution or test sufficiency.

## Findings and retained debt

Defects: none established. Smells: none established. Nitpicks: none established.

CHECK_2.json records the actual all-target rdocx check exit0 in10.413seconds and
the same Native document SHA above. This review performed no Cargo, tests,
network, source, progress, metadata or Git mutation. Only this report was
written. The earlier failed check is retained, not relabeled successful.

Affected final scoped checks, external46/7client compilation, fresh authenticated
Python regression, whole-provider Root final and actual checkpoint integration
remain pending. All original final native/oracle/full/installed/hash/release
riders remain due. This report ends the bounded support pass and returns to
Root orchestration without provider or sprint completion.
