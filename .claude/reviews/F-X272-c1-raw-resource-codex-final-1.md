# F-X272 C1 raw-block resource final review

Reviewer: Codex, delegated by Root.

Disposition: one validation defect requires remediation. Zero additional functional defects, zero structural smells and zero nits found in this bounded increment. This review does not approve full C1 or consume another implementation peer pass.

## Scope and authentication

Reviewed the raw-block source refusal classification section of `.claude/plans/F-X272-design.md:4016`, repository review rules, the operative workflow and microscope instructions. The branch is `sprint/s96`, with the existing dirty sprint carry preserved. No source changes, builds, tests or network operations were performed during this review.

The packet is `/private/tmp/rdocx-s96-root-c1-implementation/raw-block-resource/REVIEW_MANIFEST.json`. All four current source hashes and incremental patch hashes match its pins. The document and placeholder patches reconstruct their supplied frozen GREEN snapshots exactly from their captured before snapshots, checked in memory. The content-control and table changes are solely `cfg(test)` additions to previously unused production convenience entry points. Their before origins are derived, as explicitly disclosed by the manifest.

Current source pins:

| File | SHA-256 |
| --- | --- |
| document.rs | da51b2b2a3bb9cf6a95ad60edaf2939b8fa845e2f473a1cc9f91857e48d648a3 |
| placeholder.rs | f2270c2e22f4922300c7a7c8910978e150472ab7a6e6fbf8c271c8df0b47e5da |
| content_control.rs | 30b1892f70b68861ecb9d7a835a98347a04f3f9b55c870cd1392607d5200346c |
| table.rs | 51f8f2f1ce5d17633d436946d445cfa800e8e0013423b1ead654e2adfaa9903b |

## Finding

### P2: New test fails the required Clippy gate

At `crates/rdocx-oxml/src/placeholder.rs:3946`, the new header publication regression constructs `regex::Regex::new("NEEDLE")` within its loops. The actual saved `clippy.log` reports `clippy::regex_creation_in_loops` at this line. The command uses `--all-targets --all-features -- -D warnings` and exits 101 according to `clippy.json`. Compile the constant regex once outside the loops and reuse it, then rerun affected validation. Root confirmed there is no newer source or successful lint receipt at this freeze. The increment cannot claim a clean lint gate yet.

## Functional and structural review

Additional functional defects: 0. Structural smells: 0. Nits: 0.

The private classification at `crates/rdocx-oxml/src/document.rs:30` retains the original `OxmlError` payload. Existing public entry points convert it back without changing error strings. Mixed producers tag checked arithmetic and limits at their actual sites. XML reader, attribute, decoding and UTF-8 failures retain Input classification. The declaration-pool and completed-capture leaves mapped wholesale to Resource contain only resource failures. Deferred declaration lengths remain selected-use failures. No new budget or string-based error classification was introduced.

The real raw consumer now admits the original recognized paragraph, table or body-control bytes with inherited part context at `crates/rdocx-oxml/src/placeholder.rs:1017`. It preserves the old parse route when `admit_raw_word_source` returns no facts for ordinary Input at `crates/rdocx-oxml/src/document.rs:3215`. Resource failures escape before that candidate is constructed. Successful control facts feed the existing paragraph, table and control constructors. The body owner context is retained for the raw control root. Unsupported root elements continue to count zero.

The control cursor reserves disjoint preorder child ranges at `crates/rdocx-oxml/src/document.rs:1620`. The receiving control constructor retains its projected-alias fallback at `crates/rdocx-oxml/src/content_control.rs:2258`, so the new consumer does not force unproved alias projections through prepared facts.

Literal and regex checked header/footer operations edit cloned candidates at `crates/rdocx-oxml/src/placeholder.rs:1534` and `crates/rdocx-oxml/src/placeholder.rs:2317`. A raw resource error propagates through the checked walker before candidate publication. The new controls exercise an earlier successful paragraph replacement followed by refusal in each raw paragraph, table and control root, asserting unchanged original paragraphs and bytes and zero refused run construction. Ordinary opaque refusal, invalid-property raw owners and actual inherited aliases have separate controls.

## Saved evidence

The RED source snapshots, saved executable and log match the hashes in `RED.json`. The actual log passes ordinary opaque preservation, then fails resource propagation because the old operation incorrectly returns successful count 1. This is an actual RED, not a hypothetical failure.

The GREEN source pins match all four frozen current files. The log matches its receipt hash `80fcbabe741c7aad418266774a832e3f8e7ea579c227733833c1420a8292b43a`. It reports 230 passed, zero failed, with child exit 0, `RUST_MIN_STACK=8388608` and `CARGO_BUILD_JOBS=2`. This authenticates the supplied related-control result, not complete integrated acceptance.

The Clippy log matches receipt hash `0c66af06c8b98072372b81221377e6c000dc0cc69b61e3a2b05bb5715c4dfdba` and establishes the finding above. No fmt or prose completion artifacts were present in this packet at inspection. Their success is not assumed.

Later resource failures inside unproved constructors remain explicitly open. Complete original/derived provenance and aggregate scheduling, Issue 430 readiness and final integrated native, schema, hash, package and CI gates are outside this increment and remain due. Their disclosed incompleteness is not counted as a new defect here.

Review ended after this report. Writer and build control return to Root for remediation.
