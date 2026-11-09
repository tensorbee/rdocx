# F-X272 Issue430 design, Codex final 3

REPAIR REQUIRED. V5 closes the specific cross-crate visibility and lost owner-scope defects in the proposed mechanism, and adds persistent part-presence bits. Those are design corrections, not executed implementation evidence. No fourth design peer, source lease, approval or provider completion is recorded.

V5 proposal SHA-256 is 309996c085c6878cdbf18f1cb54f8bb8731127e47c132fba99900497d039b91e. Root reviewed the affected contract against the actual primary issue JSON, raw classifier and prior final. The original five issue rows remain required.

## Remaining finding

N8. The imported-presence check contradicts the prior-custody lifecycle. /private/tmp/rdocx-deepseek-harness/issue430-contract/DEEPSEEK_PROPOSAL_V5.md:147 requires stored original and edited bits to agree with the current comparison inputs. Consider a valid pending addition made from an absent part, carrying bits 0/1, then saved and reopened. Its current part exists. Comparing that document with a copy containing an unrelated paragraph edit presents current presence 1/1, so the stated check rejects the valid prior 0/1 unit. The user authorized fully valid imported convention markers, and V5:216 promises prior-revision preservation. Historical transaction provenance is not the same as current input presence. Specify how prior units retain their own transaction identity while new units and current part-level decisions are checked, including mixed prior/new pending units, save/reopen, selective resolution and unchanged prior owners. Do not silently disable the prior-revision contract or use current presence as historical proof.

This finding also applies to V5:139, which demands one identical bit pair across every unit in a part. Explain whether independently valid transactions can coexist. If the approved comparison policy already refuses the concrete case for another reason, cite that exact policy and provide a supported prior-marker case instead of inventing support. A prior valid input must not become invalid merely because serialization made the part present.

## Root question correction

Root withdraws its earlier statement that retaining refusal of a newly added unreferenced definition necessarily contracts accepted input. The actual primary issue at /private/tmp/rdocx-deepseek-harness/new-github-intake/issues/issue-430.json requires referenced additions and removals, not new unreferenced definitions. Current comparison rejects the changed owner count at crates/rdocx/src/comparison.rs:2357. Schema validity alone does not mean this API previously supported that input. Q-A430-04 may close as bounded retention of existing atomic refusal, with a discriminator and the unchanged original-unused-note preservation contract. Do not invent note-only custody solely to satisfy Root's earlier scope expansion. If another approved criterion explicitly requires successful new-unused-note comparison, cite it before closing the question. This is a Root instruction confounder.

Q-A430-08 and Q-A430-11 remain explicit draft compatibility choices. This review neither approves the recommendations nor creates additional choices about private helper visibility or XML namespace facts.

Zero new native, schema, hash, bindings or package qualification. No Issue430 comment, source edit or gate execution. Author remediation returns directly to Root within the existing exhausted design cap. STOP.
