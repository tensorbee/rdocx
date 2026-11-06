# F-X280b, final union, Root affected pass 2

**Verdict**: 0 defects, 0 smells, 0 nitpicks. D1 source repair closed. Actual corpus-consumer and final union reruns remain pending. This is affected Root remediation review, not another independent implementation peer.

`crates/oxml-opc/src/package.rs:1073`

The only production change rejects a comment end beyond the physical file instead of requiring equality. Checked arithmetic, disk/count fields, plausible central-directory bounds, maximal candidate count, ZIP64 classification and physical duplicate-name decoder remain unchanged. This matches the actual ZIP provider's documented trailing-data acceptance. No limit, API, decoder or broader policy changes. The exact pinned corpus is preserved.

`crates/oxml-opc/src/package.rs:1490`
`crates/oxml-opc/src/package.rs:1521`

The generated positive proves provider acceptance and actual bounded/unbounded OPC readability with a trailing newline. The negative proves that a physical duplicate collapsed by the provider index still refuses with the correct named part after the same newline. Both preserve source bytes. Actual predecessor execution failed both cases before repair, rather than failing compilation or setup.

Root authenticated final source SHA886e30764dafd53ce1bb821de310ee7c3b44954fd25d7934d5f81b9866dba43d and retained predecessor/fixed executable hashes against SOURCE_BOUNDARY.json in /private/tmp/rdocx-s96-f280b-trailing-compat-9v0m0k5u. Actual fixed observations cover5issue435 cases plus the fake-EOCD, truncated-ZIP64 and2bounded-reader controls,9distinct passes. Scoped check, strict all-target/all-feature OPC Clippy, formatting and whitespace checks pass. Earlier copied-binary execute-mode setup failure is separately retained and not claimed as semantic RED.

The full workspace at originalH remains failed. The six actual pinned-corpus failures need final-source rerun. Three separate missing-soffice failures remain environment/oracle work and are not fixed by this predicate. Genuine pinned Linux corpus/render execution, all other final categories and the actual CI successor remain required. No new ignored test, advisory exception, baseline update or deeper security assessment follows.

## Defects

0. D1 source closure accepted.

## Smells

0.

## Nitpicks

0.

## Not found

No further correctness, compatibility, test-discrimination or scope defect established in the bounded change. Only this review record was written. The audit ends here before commit or further remediation.
