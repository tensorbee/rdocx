# F-X272 first drawing extent final audit

Reviewer: Codex, delegated by Root.

Disposition: retain the bounded increment. Functional defects: 0. Structural smells: 0. Nits: 0. This is not complete drawing qualification, full C1 acceptance or a fourth implementation peer pass.

## Scope and authentication

Reviewed the approved first drawing extent negative action at `.claude/plans/F-X272-design.md:4085` under the existing repository workflow and microscope rules. The packet is `/private/tmp/rdocx-s96-root-c1-implementation/drawing-first-extent/REVIEW_MANIFEST.json`.

Current document SHA-256 is `41568823eaf305032266b3a651cd68019a6672f8611b609cc0d2502dea6ccc34`. Current drawing SHA-256 is `5b4fdad0241f00d603fa4c2daf8d36bb6564a4999c6d6c0db793a7e1e233c8af`. Both current files and both incremental patches match the manifest pins. Applying the supplied patches to their captured before snapshots in memory reconstructs the exact current files. No source edits, builds, tests or network operations were performed during this review.

## Findings and observations

No findings.

The shared decoder at `crates/rdocx-oxml/src/drawing.rs:2061` retains the exact former attribute-loop behavior in both actual consumers at `crates/rdocx-oxml/src/drawing.rs:1032` and `crates/rdocx-oxml/src/drawing.rs:1447`. Every attribute value is decoded before its key is selected. Only literal unprefixed `cx` and `cy` update dimensions, using the existing signed numeric parse. Missing dimensions retain prior values. Attribute order and partial updates before a later error remain unchanged.

Qualification enters Drawing only from an actual nonempty Word drawing in a run at `crates/rdocx-oxml/src/document.rs:1803`. Direct container selection at `crates/rdocx-oxml/src/document.rs:1707` precedes the generic foreign-element rejection and uses the actual reader's lexical container selector. This correctly avoids introducing a stricter namespace interpretation for inline and anchor containers.

The negative certificate at `crates/rdocx-oxml/src/document.rs:1904` permanently marks the first direct element as seen. It decodes only when that first element is an Empty local-name extent. Started extents, later extents and extents nested in consumed subtrees do not acquire a failure certificate. Container children otherwise remain opaque, so nested container early-End cursor behavior is not treated as ordinary full-tree construction. A decoder error propagates through Drawing and Run rather than acquiring a new swallowing boundary. A successful extent does not mark the full drawing known or successful.

The four-case regression at `crates/rdocx-oxml/src/document.rs:9735` covers inline and anchor with canonical and foreign extent names. Actual bad numeric constructors make their enclosing control raw. Their valid numeric counterparts include required docPr metadata, remain constructible and still refuse later genuinely typed over-depth controls. The discriminators at `crates/rdocx-oxml/src/document.rs:9793` preserve unexecuted nested, Started and early-End cases. The actual-reader control at `crates/rdocx-oxml/src/document.rs:9828` retains a prior dimension across a later partial extent and decodes a signed numeric entity.

## Saved evidence

The saved RED document source, executable and log match their receipt hashes. The actual log reports `[false, false, false, false]` instead of the required four true outcomes after the constructor and valid-counterpart assertions. This is a genuine failure of the captured earlier source.

The GREEN receipt pins both current sources before and after execution. Its log matches hash `27f27dffaafde127b329049962fd02a1734e32995e4ecd66e4b3810ce032426c` and reports 285 passed, zero failed, child exit 0. The lower-crate all-target, all-feature Clippy with warnings denied, workspace formatting and prose logs match their receipt hashes and each reports exit 0. Prose reports zero violations.

No quota, public API or error string change was identified in the implemented scope. Broader drawing cursor and scalar qualification, complete original/current provenance and aggregate scheduling, Issue 430 completion and final integrated native, schema, hash, package and CI gates remain open. The bounded negative certificate does not establish those obligations.

Review ended after this report. Writer and build control return to Root.
