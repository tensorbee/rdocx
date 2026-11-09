# S96 opening and standalone value implementation support 7

## Verdict and scope

**D1 resolved. Zero defects, zero smells, zero nitpicks.** This is an evidence-only successor to implementation support6 and Root24. Their unchanged-product source audit is consumed without repeating it. Support6 remains intact, including its original D1 finding. All prior pass numbers remain.

Root remains sole product/Cargo writer. Only this report was written. No source patch, Cargo command, runtime import, test rerun or benchmark was performed. The audit covers the count-correction script, retained original receipts, corrected per-case outcomes, aggregate interpretation and unchanged product boundary. It does not establish Office, performance or final-union acceptance.

## D1 resolution

`correct_native_counts.py:6` preserves each original Native receipt before correction. At11 it parses the actual named libtest outcome, and at12 it checks per-case semantic totals against all63actual harness result lines. Successful child exits are counted separately. The summary replacement at17 explicitly distinguishes actual passes, ignored Office cases and successful commands. The final correction record is written at18. No product or test-execution command exists in this script.

Independent read-only parsing reproduces every corrected row and total:

| Receipt | Actual passed | Actual failed | Ignored | Successful child commands | Commands |
|---|---:|---:|---:|---:|---:|
| Initial Native cohort | 61 | 1 | 1 | 62 | 63 |
| Final retained Native cohort | 62 | 0 | 1 | 63 | 63 |

Both cohorts contain63unique named cases. Every new semantic_outcome agrees with its actual log line, and each original row's name, exit and elapsed value is unchanged. All other original receipt fields remain unchanged except the corrected totals and added interpretation. The initial failed assertion remains failed in its historical receipt.

The corrected initial totals appear at `native-tests-1.json:383`. Final totals appear at `native-final-tests.json:386`, with the Word case explicitly ignored at263. `SCOPED_SUMMARY.json:11` now reports62actual passes, and its Office qualification at30 remains pending. It also retains the initial61pass/1fail/1ignored outcome and separates the corrected followup from final custody consolidation.

The original ignored receipt at `native-final-tests.log:4033` and ignore attribute at `crates/rdocx/src/comments.rs:6663` remain unchanged. Thus D1 is resolved by correct interpretation of the actual execution, not by relabeling an ignored test as executed or treating child0 as a semantic pass. The progress checkpoint explicitly supersedes the earlier63Native-pass claims and keeps Office evidence pending.

## Preservation and authentication

Packet: `.claude/scratch/S96-final-evidence/f268-standalone-value-implementation-20261107-1`. Independently authenticated retained originals match:

- native-tests-1-before-D1-correction.json: SHA-25678af050a95be92e174cbed9a504df62ef38e3133089760ac741a8c63a4c82bcd.
- native-final-tests-before-D1-correction.json: SHA-2567e5596eb3a571fe657de987a71b99bc3960384b18df427c4cf02b68e3e03b250.
- SCOPED_SUMMARY-before-D1-correction.json: SHA-2564cb1cfe1863086f620598a79dcb9016faf13495fe14baff46fcb79c940a00ca5.

D1_NATIVE_COUNTS_CORRECTION.json matches SHA-256383a7aec610edb91b0b86bcc5c1526e2cab85239cbde8f8e78665c127bbe4923. Its counts, ignored name and preserved-original hashes agree with the actual files and independently parsed logs. Non-Native fields of SCOPED_SUMMARY remain equal to the retained original. No other public, mechanism or artifact result was changed by the correction.

SOURCE_BOUNDARY.json remains exactly SHA-256cc4594d258c1f491d9031437b86b7144b9e5f100400534ae4cb7aec1002ee199. All12reviewed product/plan/HLD pins still match. The complete4309input manifest still has exactly those12Root23differences, with4297other inputs unchanged. The retained Native binary remains4b3eb840629ff7a8537652d328d11332afdb6bf4596aa742dd240844de40a1a4. Support6 remains SHA-25634f7cd7105058f1490f7afc5a8ba1e7a05beebfcc30fe5d94deded19fbb4b2c2.

The Native producer's source_unchanged field continues to describe the actual observation window. The later sole Python regression-regex edit remains explicit through ORDINARY_APPLICABILITY. Support6 already authenticated that applicability and the unchanged runtime, diagnostic and ordinary installed payloads. This correction does not claim a new execution at the final metadata state.

## Remaining gates

The four last ordinary elapsed misses remain failed. Fresh ordinary final-source qualification, pinned Office/native evidence, whole installed typing, memory, hosted, archive and full-union gates remain due. No gate is waived, no historical feature status is reset, and no source fix, confirmation rerun, release build, timing, ledger change, commit, push or publication was performed in this pass. Return control for Root25.
