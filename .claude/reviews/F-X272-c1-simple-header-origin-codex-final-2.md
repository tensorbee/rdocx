# F-X272 simple-header context remediation review

Reviewer: Codex, delegated by Root.

Disposition: the targeted context regression is resolved. Functional defects: 0. Structural smells: 0. Nits: 0 in this remediation. This is a follow-up to the existing Root increment, not a peer reset or full C1 acceptance.

## Prior review correction

The preceding clean bounded review missed an actual regression introduced by Root's header reuse. Admission could recognize its implicit bare `w` alias while the ordinary preserving factory used the declared `q` scope. Reusing the admitted header could therefore fabricate a field from an instruction attribute the real factory would ignore, or set a dirty flag from an unbound attribute. The actual RED records two fields where one is required. Root found and fixed this regression. It is not attributable to GLM.

## Targeted comparison

Compared current document and text sources against `document-green.rs` and `text-green.rs` in `/private/tmp/rdocx-s96-root-c1-implementation/simple-header-origin`. The exact changes add first-qualified instruction and dirty attribute indices, check their compatibility before cloning a header, pass the consuming start event and prefix context, and add the regression test. The intrinsic scalar calculation and physical input/span checks are unchanged.

The first-match helper at `crates/rdocx-oxml/src/text.rs:7035` follows the same physical attribute order, malformed-attribute skipping and qualified-name predicate as the existing optional attribute reader. It records attribute identity rather than merely whether some matching local name exists. The retained live input and exact opening checks still establish the corresponding bytes.

The compatibility check at `crates/rdocx-oxml/src/text.rs:7082` always compares the first instruction attribute. It compares dirty only after a syntax-bearing header, matching the factory's early None policy. When instruction decoding already produces None, the factory does not read dirty, so requiring dirty equivalence there would be unnecessary. A changed first qualifying instruction or dirty attribute rejects reuse. The check runs before cloning at `crates/rdocx-oxml/src/document.rs:1662`, returning to the actual factory path on mismatch.

The new regression at `crates/rdocx-oxml/src/text.rs:22402` checks that an admission-only instruction remains raw, the actual qualified instruction remains typed and the unbound dirty attribute remains ignored. It also asserts retention of the raw instruction bytes. Existing physical receipt, copied backing, subslice and insertion-mapping controls remain passing in the saved related run.

## Authenticated evidence

Current document SHA-256 is `54605035ff508f0603b6aa6392c42f38c384e826695a390418b3e8e6fd572665`. Current text SHA-256 is `9e163aafdd9cb35f11f807ec36433bd50fc727d483b227b93899550b1f77e420`. Both match the context validation receipts, which report unchanged source during execution.

`text-context-red.rs`, `context-red-binary` and `context-red.log` match `CONTEXT_RED.json`. The log actually fails with two fields versus one. This is the introduced regression's saved failing evidence.

The context-controls log matches its recorded hash `fd431e360c7d967f4bd3a8307f82e9e1a99f4b02f851c8c1ed87add28a7d2d38` and reports 304 passed, zero failed, exit 0. The context Clippy, formatting and prose logs also match their receipt hashes and each reports exit 0. Clippy uses all targets, all features and warnings denied. Prose reports zero violations.

Broader historical declared-root admission alias behavior remains open. This repair restores the targeted factory semantics and does not waive that debt or establish complete constructor provenance, aggregate scheduling or final integrated acceptance.

No source edits, builds, tests or network operations were performed during review. Review ended after this report. Writer and build control return to Root.
