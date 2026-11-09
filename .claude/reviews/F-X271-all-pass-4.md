# F-X271, all aspects, pass 4

**Reviewer**: independent Codex reviewer. This reviewer supplied earlier read-only technical recommendations and did not implement the corpus retry patch.
**Scope**: approved bounded presentation-corpus retry increment against its actual before copies. Three files, 412 added and 17 removed lines in the 481-line incremental patch. Read the current plan, latest pass 3, progress checkpoint, canonical microscope and actual source/tests/HLD. Earlier F-X271 installer/security work and unrelated F-X275 punctuation are outside this product increment.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Scoped retry review is clean. F-X271 remains in progress and this is not final hosted/native/fidelity acceptance.

## Source and evidence identity

Approved plan SHA256 is 1caf3d6bd3a43c669567805d8e4596487f7189c506fee1788aa37901c7630a2a. Incremental patch SHA256 is e2700ade41b44c37583f483b30ac1c7ff431e7f69eaebef6cf774f107096d31d. Handoff SHA256 is 30f14a10613d5aa72c395cfc6dfc2d4248813dbbcd04b5d04aaf42d9b1b2ce49. Packet seal SHA256 supplied by root is 49755148e27cfe498b403267651547a8b19508ea6bd5600fb1ad4d5119704be9.

Independently authenticated all 60 sealed packet members with no mismatch. The actual current three inputs and unchanged manifest agree with source-after.json:

| Input | SHA256 |
|---|---|
| scripts/fetch_pptx_corpus.py | 9263bd30a81f147e8bde0a3693cea0eb8f1dd5be5f926cf4f89f2eae00fa0718 |
| scripts/test_sprint_workflow.py | 5dc51332ee859622a701b2bd12c6951882ca1ec357ad6153632bf68a3df59188 |
| docs/hld/15-build-and-toolchain.md | 7caae1886d3d0970c62efa93b116cba2bcd5fa18c4477a748c80df03d9c20223 |
| scripts/pptx-corpus-manifest.tsv | c51acd82358c04de9b72c3bae2557b945c373d0f67320010b43be579ac80925f |

Root additionally authenticated the 3840 protected inputs before its separately disclosed punctuation correction. No unrelated source change is attributed to this retry increment.

## Examined implementation

_retryable_transport at scripts/fetch_pptx_corpus.py:80 unwraps only exception-valued URLError reasons and detects cycles. HTTPError is classified before the URLError loop can erase its response identity. Only the five approved statuses retry. Certificate/generic SSL, permanent DNS, string reasons and unknown errors remain terminal. Explicit transient exception/errno handling does not inspect message text or arbitrary cause chains. The helper retains the actual HTTPError object for response closure, including a wrapped one.

_fetch_subject at line 104 owns the complete attempt and publication rule. The network catch boundaries at lines 117 and 126 exclude temporary-file open/write/close, digest, replace and cleanup. A TimeoutError from local operations therefore propagates rather than becoming transport recovery. The response context closes ordinary failed/successful responses. The finally block at lines 145 to 151 closes classified HTTP error responses and still attempts owned-temp cleanup if that close fails. Cleanup and response-close exceptions remain terminal and do not proceed to sleep/retry.

The temporary is opened exclusively at line 113, and ownership is set only after successful creation. Foreign/stale collisions are never truncated or unlinked. Partial bytes are discarded between attempts. Exactly three attempts and delays 1 then 2 are enforced at lines 106 and 152 to 154. Diagnostic output carries subject, actual exception and attempt count. Silent EOF goes through digest verification and remains an integrity failure. Atomic replace occurs only after a matching complete digest, leaving an old destination intact through failed attempts.

fetch at line 157 retains cache digest validation, preflight refusal of extra files and final whole-directory verification. The 50 manifest identities, HTTPS paths, checksum pins and existing CLI remain unchanged. HLD15 at line 1116 records the same bounded behavior and correctly calls timeout60 an operation timeout, without inventing a total-subject deadline. No fallback endpoint, dependency, new module, feature flag or public option was introduced.

## Discriminators and actual outcomes

Inspected the new cases in the existing policy carrier beginning at scripts/test_sprint_workflow.py:2512. They assert recovery after initial open timeout and after actual partial reads, exact attempts/delays/timeout/header, discarded prefixes, response closure before retry, original destination custody and final exact bytes. Classification cases include all approved and terminal HTTP statuses, temporary/permanent DNS, direct/wrapped certificate/SSL errors, nested and cyclic reasons, raised incomplete reads and named errnos. HTTP error payloads must not be read.

The local-failure cases inject transient-looking errors into open/write/close/digest/replace/unlink, error-response close and ordinary response close. They require no retry, old destination preservation and correct owned cleanup. The explicit unlink failure leaves its owned file as an actual terminal failure rather than pretending cleanup succeeded. Cache and concurrent exclusive-open collision cases prove no network and preserved foreign bytes. Final directory and manifest checks remain covered. Existing Word corpus controls are included without changing its fetcher policy.

The retained actual old-source recovery test produced three genuine semantic recovery errors, rather than setup/mock errors. The current acquired batch contains 9 tests and 58 passing subcases, with no skips: 48 presentation and 10 Word subcases. These are finite local controls, not an actual 50-deck network retrieval. No tests were rerun by this reviewer.

The worker's full prose result remains NONPASS because a protected preexisting F-X275 plan semicolon was present. Root subsequently changed that one punctuation mark after the lease returned and obtained actual prose and diff exit0. This does not rewrite the old packet's result or qualify new product behavior. The already passing nine-test batch need not be rerun for that unrelated punctuation. Adapter/diff and local HLD prose observations are retained with their actual packet outcomes.

## Defects

Zero independent correctness, contract, resource custody, panic, test, compatibility or structure defects established in this bounded increment. No OOXML or rendering transformation is introduced.

## Smells

Zero. Previous S1 readiness disposition remains closed under the earlier reviews and does not erase the current final debt.

## Nitpicks

Zero.

## Handback

F-X271 remains approved/in progress. Actual source-qualified supported-host retrieval of all 50 decks, genuine fidelity/native CI and the complete integrated policy/product/installed/hash/risk union remain unaccepted. No completion, push, dispatch, tag, release, comment or rendering-baseline change follows this clean scoped review. An unchanged producer suite is not rerun for this report.

Only this review file was written under root's released review lease. No production, tests, plans, HLD, progress, TEMP or other file was changed. No product execution, build, oracle, native operation, network request, commit or push ran. End this clean pass and return control to root.
