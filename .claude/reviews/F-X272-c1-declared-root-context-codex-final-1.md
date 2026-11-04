# F-X272 declared-root context final audit

Reviewer: Codex, delegated by Root.

Disposition: one functional defect requires remediation. Structural smells: 0. Nits: 0. This bounded review does not approve full C1 or add a fourth implementation peer pass.

## Frozen scope and authentication

Reviewed the six source increments and three pinned HLD files in `/private/tmp/rdocx-s96-root-c1-implementation/declared-root-context/REVIEW_MANIFEST.json` under the existing repository workflow and microscope rules. Every current source, patch and HLD hash matches that manifest. Applying each exact source patch to its captured before snapshot in memory reconstructs its current source. No source edits, builds, tests or network operations were performed during review.

## Finding

### P1: Unconstructed spans can conceal paragraphs the part reader actually constructs

At `crates/rdocx-oxml/src/document.rs:1738`, ParagraphCollection treats every non-paragraph child as Unconstructed. Notes and Comments apply the same blanket exclusion to non-entry children at `crates/rdocx-oxml/src/document.rs:1745` and `crates/rdocx-oxml/src/document.rs:1752`. Their complete spans become discarded at `crates/rdocx-oxml/src/document.rs:2482`.

The actual part loops also recognize their wrapper elements at arbitrary encountered depth. Header/footer recognizes another hdr or ftr at `crates/rdocx-oxml/src/header_footer.rs:274` and continues into it. Notes similarly enters footnotes or endnotes at `crates/rdocx-oxml/src/footnotes.rs:152`. Comments similarly enters comments at `crates/rdocx-oxml/src/comments.rs:67`. None of those branches checks that it is handling only the outermost root.

For example, a well-formed `<q:hdr xmlns:q="WORD_URI"><q:hdr><q:p>...</q:p></q:hdr></q:hdr>` is traversed by the existing header reader, which constructs the inner paragraph. The new qualifier instead discards the entire inner hdr span as unconstructed. Typed controls in that paragraph are omitted from the admission facts, yet the new paragraph consumer at `crates/rdocx-oxml/src/header_footer.rs:267` no longer performs the previous independent paragraph admission. Nested notes/comments wrappers have the corresponding mismatch. This can hide actual modeled work and invalidate control occurrence correspondence. It is a static code finding, not a claimed observed crash or test failure.

Keep the historical reader behavior and model its actual wrapper cursor modes, or use a conservative admission fallback when those modes are not proved. Only mark a complete subtree Unconstructed when the concrete reader really skips or captures it. Add a discriminator with nested recognized wrappers and live paragraph controls, including an over-budget typed counterpart. Do not resolve this by silently imposing a new schema rejection policy on the existing permissive readers.

## Other reviewed behavior

No additional functional defects found.

The six declared-root entry points now pass empty inherited scope only at their complete-input seams. Consumed readers and externally supplied context retain their existing defaults. The live-input controls_for consumers receive their actual full input slices, and the ordinary supported paragraph paths reuse one admission rather than granting each child a new allowance. The declared-root regression distinguishes unbound w attributes from bound q owners, including an earlier phantom control that previously consumed a live control receipt. Glossary still executes its strict lexical validation after source admission.

The HLD additions at `docs/hld/03-architecture.md:2663`, `docs/hld/04-opc-and-packaging.md:2039` and `docs/hld/10-bindings-spec.md:2188` describe selected-context admission, current part consumers and checked resource propagation. They do not claim complete original provenance or aggregate generated-work closure. The intended skipped-block description remains appropriate, but the actual skip proof needs the repair above.

## Saved evidence and limits

Both RED packets' saved source snapshots, binaries and logs match their recorded hashes. The first actual RED reports six incorrect resource/error-boundary outcomes and loss of live q-control text. The collection RED reports three false raw-block refusals and observed qualification/physical-event pairs `(2447, 143)`, `(2449, 145)` and `(2449, 145)`.

The GREEN receipt pins all six reviewed current sources before and after execution. Its log matches hash `ebb8def39135fddea00eb19d76dcfbd7b5951f6d8b2fdaeede057a87cb3657b6` and reports 395 passed, zero failed, child exit 0. The lower-crate all-target, all-feature Clippy with warnings denied, workspace formatting and prose logs match their receipt hashes and report exit 0. Prose reports zero violations. These controls do not cover the nested recognized-wrapper mismatch identified above.

Complete original/current provenance and aggregate scheduling, unproved constructor modes, Issue 430 completion and final integrated native, schema, hash, package and CI gates remain open. No full acceptance follows from the supplied passing controls.

Review ended after this report. Writer and build control return to Root for remediation.
