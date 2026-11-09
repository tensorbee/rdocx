# F-X272 Issue430 design, Codex final 4

REPAIR REQUIRED. V6 proposal SHA-256 f375830c9cf7b9e2afe19f648d7b1d352d1e1be7499bf4984cbe3482a18a6a27. Q-A430-04 is closed as retention of existing atomic refusal for new unreferenced definitions. N5 and N6 remain closed at design level. Historical provenance is now distinguished from current input presence, but the added fold does not yet establish N8 lifecycle closure. No source lease, fourth peer, counter reset, implementation or provider completion.

## Findings

N9. The fold contradicts its own example. /private/tmp/rdocx-deepseek-harness/issue430-contract/DEEPSEEK_PROPOSAL_V6.md:194 and :198 make a 1/1 unit set presence to true even when rejected. At :246 the prior 0/1 plus new 1/1 example promises both rejected gives absence. Starting at 0, rejecting the prior unit leaves 0, then the stated 1/1 rule sets 1. No retained notes or references override that contradiction. Define one coherent resolution rule and test its result for both resolution orders and save/reopen. Do not merely change the example to preserve a part that the required original-state reconstruction would remove.

N10. Token numeric order is not transaction chronology. V6:145 equates the token with a revision id. V6:191 and :197 use ascending tokens to choose the earliest transaction's original state. Actual IdAllocator starts next at zero at crates/rdocx/src/comparison.rs:764 and allocates the next unused id at :815, so a later comparison can fill a lower hole beneath an older unit's id. A valid prior 0/1 unit at id20 followed by a new 1/1 unit at id0 sorts the later transaction first and starts from original-present, losing the true original absence when both are rejected. Imported tokens and collision rebinding at V6:159 provide no chronological proof either. Use persisted facts that actually determine the cleanup invariant without pretending arbitrary local ids encode time. Preserve existing allocator scope unless an explicit justified contract amendment is required. A new id convention alone does not qualify already-valid imported markers. Do not add speculative transaction machinery if a smaller invariant suffices.

Zero new C4 or source-budget findings. No test, native, schema, hash or product gate was executed in this review. Tombstone recognition and final-marker cleanup remain part of the affected closure obligation, including what can actually be decoded after revision elements or note definitions disappear.

## Explicit user answers

The user has now answered both remaining compatibility questions. Exact authority is /private/tmp/rdocx-cross-review/ISSUE430_USER_DECISIONS_20261101.json. Q-A430-08: follow the schema, omitted type means normal at any id, with untouched imported XML preserved. Q-A430-11: preserve untouched empty or unknown types and atomically refuse changes to the owner or its references. These are approved policy choices for the planned minor release, not claims of implemented behavior or gate completion. Apply them consistently to the parser rule, classifier/preflight, acceptance rows, examples, compatibility notes, checklist and handoff. Do not ask them again or leave operative BLOCKING-choice wording.

The user answers do not approve the unclosed fold or the whole draft. Root direct review follows author remediation inside the existing exhausted design cap. STOP.
