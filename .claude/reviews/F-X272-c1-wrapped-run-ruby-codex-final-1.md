# F-X272 wrapped-run and Ruby prefix final audit

Reviewer: Codex, delegated by Root.

Disposition: retain this bounded increment. Functional defects: 0. Structural smells: 0. Nits: 0. This is neither full C1 acceptance nor another implementation peer pass.

## Scope and source authentication

Reviewed the operative wrapped-run and Ruby prefix plan sections at `.claude/plans/F-X272-design.md:4045` and `.claude/plans/F-X272-design.md:4066` under the existing repository workflow and microscope rules. Only the combined document and run-property increments in `/private/tmp/rdocx-s96-root-c1-implementation/ruby-property-prefix/REVIEW_MANIFEST.json` are assessed here.

Current document SHA-256 is `c812c7f6a58c06bffb0a853ce53de04f3e1f0a88e4d0f2c6a705ddbb9abb949d`. Current run-property SHA-256 is `426d925bd40ce076776fbcdbd052692fa340e1d836b6127e7f34ef3ab883e509`. Both current files, both captured before files and both combined patches match their manifest hashes. Applying each patch in memory reconstructs its exact current source. No source edits, builds, tests or network operations were performed during review.

## Findings and reasoning

No findings.

The simple-field classifier distinguishes Unknown projection, early raw header None and permission to parse the result at `crates/rdocx-oxml/src/document.rs:2234`. Only the result state dispatches direct runs at `crates/rdocx-oxml/src/document.rs:1758`. The raw state keeps the early child boundary, while the Unknown state does not acquire result-constructor proof. A propagated result-child error now reaches its enclosing control. The existing revision swallowing boundary stays explicit at `crates/rdocx-oxml/src/document.rs:2381`. Direct hyperlink runs use the same concrete run actions rather than being opaque to negative facts.

The shared `empty_bold_value` primitive at `crates/rdocx-oxml/src/run_properties.rs:29` uses the existing permissive toggle parser. Its actual Empty-event constructor consumer is at `crates/rdocx-oxml/src/run_properties.rs:838`. Started bold handling is unchanged. The qualifier invokes this same decoder only for Empty property children at `crates/rdocx-oxml/src/document.rs:1935`. Unsupported element children turn the property frame Unknown. Existing negative measurement facts still propagate. Positive property knowledge reaches the parent run only after the property frame closes. The property's ignored non-element event policy agrees with the actual property reader's catch-all handling at `crates/rdocx-oxml/src/run_properties.rs:1231`.

The Ruby prefix control at `crates/rdocx-oxml/src/document.rs:11355` tests omitted, false and unknown legacy bold spellings. It verifies one actual prefix-run construction followed by Ruby None, keeps the fully typed counterpart resource refusal, and checks that an earlier invalid size outranks the later Ruby None. The wrapped-run controls separately retain the missing-instruction early None boundary and the later genuine over-depth refusal. No public signature, quota or source allowance is changed by these increments.

## Actual saved evidence

Both RED document snapshots and saved binaries match their recorded source and binary hashes. Both actual RED logs match their receipt hashes. The wrapped-run RED reports all three child-error cases as `[false, false, false]` instead of `[true, true, true]`, while the missing-header control passes. The Ruby RED reports the same three-false mismatch after its constructor assertions have established the prefix run and late None. These are actual failures of the captured prior sources.

The current GREEN receipt pins both reviewed sources before and after execution. Its actual log matches hash `1b3a692656fc8468e3df8cb69e725d3c489d7cc1e1af1b887f603ce3fcdf1e65` and reports 256 passed, zero failed, child exit 0. The Clippy, formatting and prose receipt logs also match their recorded hashes and report exit 0. Clippy covers the lower crate with all targets, all features and warnings denied. Prose reports zero violations.

This evidence supports the bounded changes above. It does not establish complete constructor facts, original/current provenance or aggregate scheduling, remaining drawing/revision contexts, Issue 430 completion, or final integrated native, schema, hash, package and CI gates. Those disclosed obligations remain open.

Review ended after this report. Writer and build control return to Root.
