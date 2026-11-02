# F-X282, Codex implementation, pass 2

**Reviewed**: Independent non-author affected review after pass1 remediation, on `work/f-x282-codex`, Base `aff3f2b8744ed86755db8520c9bf6abeed5f2038`. Compared the frozen reviewed source with current source, read both new native fixtures, the new public fixture, changed HLD12 prose, actual receipts, current plan and unchanged publication paths relevant to the findings. Production source is unchanged. The native carrier adds 177 lines inside its existing test module. Interpreted Python tests and HLD12 test prose are the other affected files.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this affected pass. Pass1 D1 and D2 are closed. This is a clean implementation microscope, not final performance acceptance, integrated source acceptance or feature completion.

## Prior findings

### D1, closed with actual strict-error acceptance controls

`crates/rdocx/src/document.rs:49881`

The new retained-source fixture uses the exact four historical lexical forms: duplicate expanded attributes, literal less-than in an attribute, numeric qualified name and unbound prefix. It opens actual packages and changes the style through both the actual checked editor and the new bridge. Crucially, each callback builds the typed `ContentFragment::paragraph` and asserts that its emitted candidate still fails strict validation at `crates/rdocx/src/document.rs:49910`. Source malformedness alone is not used as a substitute for this discriminator.

Both publishers must succeed. Each callback count must be one. The bridge must report checked fallback for each strict-error case. Equal models, exact saved ZIP bytes, retained raw fragments, reopened text and the Normal style are asserted. A blanket strict-error refusal would fail the bridge's success assertion. The ordinary fixture explicitly passes the strict probe and must select direct publication. The earlier qualified foreign-extension control and its explicit namespace-placement difference remain intact.

The actual final native receipt includes this case and succeeds. These current-source assertions discharge the missing local evidence rather than merely adopting the old Root probe.

### D2, closed with the actual checked preflight boundary

`crates/rdocx/src/document.rs:49974`

The new fixture places a nested table inside a cell content control, with distinct outer-before, controlled-table and outer-after values. It proves three reader paragraphs against two direct paragraphs. All three target ordinals enter the actual checked resolver, which returns `cell paragraph modeled/source correspondence is unsupported`. The bridge must preserve that exact preflight refusal with zero callbacks and zero direct or completed checked publications, at `crates/rdocx/src/document.rs:50009`.

Full live model equality, original source bytes, opaque sentinel retention, exact saved bytes and reopened physical paragraph order are checked. A wrong direct eligibility rule that ignored control-owned table descendants would violate the callback, outcome or state assertions. The existing direct nested-table control and successful directly controlled-paragraph controls remain, so unsupported correspondence is not generalized into a refusal of ordinary controls.

The public test at `crates/rdocx-py/tests/test_core.py:16083` tests all three targets and retains held paragraph/run values, fresh values, exact saved bytes, the sentinel and reopened order. The same three public cases pass on the authenticated pre-F282 extension. This establishes existing refusal policy, rather than a new capability reduction. Pass1 asked for the omitted ownership discriminator. It did not authorize expanding the existing checked resolver's supported grammar. Preserving the actual preflight refusal closes D2 without a production change.

## Defects

None found.

## Smells

None found.

## Nitpicks

None recorded.

## Authentication and actual outcomes

The new handoff, final source manifest and evidence index match their supplied SHA-256 identities: `03d08073cf0ceaccb89bdc00f3db2e997c1edc1c54476849437032c3cface627`, `f5896b41227e72f5c5327c3b93dd0ac8e3a8b660f6fc58592014e94ed20b1064` and `379b0b53425f16963cfec0b5b78a4c44ba0389f76efbc761cea3e015c779be46`. All twelve current source hashes match the manifest. All indexed log hashes match their receipts.

Independently reconstructed the prior full document source in memory using current source and the current/frozen patches. Its full hash matches the original reviewed `946f50e7d091186ecffac6a7da709951abdf5e08ee56e863816c6964604216d1`. Compared the complete prefix before the existing native test module and obtained exact equality, SHA-256 `c282062b435a09a8052545f2f5cecc8847d368f7d1c039a778c1cf9ad87fb86f`. The remaining production files match their original reviewed hashes. No production change is hidden by the new test carrier hash.

The unchanged current extension matches `bcd1116683d73c37241ed0d178e2ca55127db73d1901fd601b22797bf9efa74b`. The independently interpreted Python source matches `123eb7f6d36585909d883026c5a0dfb97a0385802580ccbea4600a01ccc5a29d`. The old-policy receipt prints the pre-F282 imported extension identity `9bd8cc2ad5037562441719870e391d7e98a61ee97ea6cb5ecfce5737a30169bf`.

Read actual outcomes, not just aggregate claims: `peer1-native2` reports all fourteen named native cases passing, including the two additions. `peer1-python2` reports three public cases passing. `peer1-python-baseline-policy1` reports the same three passing on the old artifact. Child and wrapper exits are zero and no timeout is recorded. Affected clippy, format, prose and diff receipts also pass. The earlier 42 public passes remain a separate applicable observation. They are not relabelled as a newly executed 45-case suite.

The initial remediation receipts preserve fixture failures. The first fixtures assumed that the control-owned table could be mutated through the existing checked resolver and overasserted whole-model equality for the ordinary direct case. The final assertions correct those assumptions while preserving the four strict-error acceptance controls. These failures are not claimed as production RED/GREEN. The read-only source reconstruction during this review also initially omitted a diff-section terminal newline, then corrected that parsing and verified the full original hash. It made no source or evidence writes.

## Not found and limits

No new correctness, contract, panic, OOXML, test or structure issue was found in the affected repair. The existing at-most-once callback, immutable preflight and checked candidate fallback remain unchanged. The repair does not add a trait, generic, module, file, test binary, feature flag or public API. No F-X278 font-policy change or new F-X272 provider dependency is introduced. Actual checked-policy composition with the future integrated C1 source remains due.

The malformed styles replacement refusal still needs its dedicated labelled behavior commit when Root prepares integration. The strict-qualified foreign extension still records its inherited-scope versus redundant-declaration saved-byte difference. Exact saved-byte/hash acceptance is not inferred from semantic equality. The failed writer proof remains a synthetic guard test.

All exact Linux release Python3.12 A282-1/2/3 ratios, original workloads and cold sampling rules remain pending. Final F-X268 79/252/407, saved bytes/hash49, full workspace/native/oracle, installed runtime/typing, WASM, wheels, archives and hosted CI gates also remain pending. Debug semantic evidence does not discharge them. Root final review and normal workflow completion are separate.

No tests, builds, benchmarks, network operations, Git actions or source edits were performed in this review. This review file is the only write. Stop here and return the sole-writer lease to Root.
