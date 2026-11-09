# F-X277, Root preparation final 1

REJECTED PREPARATION after the third and last actual design peer. This is a scoped Root check of the remaining blocking contract, not a claim that every other mechanism has received final independent approval. No fourth design peer or counter reset.

## Identity and scope

Grok V3 SHA-256 `d1971df3ea58fab3d8992acea294cc4f4c1f1190b617909df4a4151d060293f2`, native handoff `3383f7c6be0fc5b436a552a8198f6dfdd20816d4f3f2ec537f096da4de6c4282`. Actual DeepSeek design peer 3 report is `/private/tmp/rdocx-cross-review/fx277-deepseek-design-peer3-current.md`, SHA-256 `2a8a102e47405fff4f4fab1c4a16154d0dec4ee63bbc07f2a783064331f4f532`. Root read that complete report and inspected the proposal's numbering/publication contract, registration, current signing path, Python test carrier and equation/validator cross-story contract.

## F1, medium, signing breaks the promised live definition id

Confirmed. `GROK_PROPOSAL_V3.md:167` says the returned numbering id stops being a handle after sign and equates that boundary with process reopen. The same document instance remains alive. `docs/hld/14-development-backlog.md:8798` requires that returned definition id to stay valid for the life of the document. Current `crates/rdocx/src/document.rs:24610` stages and commits a candidate, and the proposed emission remap would commit its published numbering ids. The explicit exception is a contract reduction, not a factual closure.

Preserve the original handle through sign and later mutations while retaining signed package equality on a later unchanged save. Trace every staged/reopened and emission publisher, cache identity and identifier set. A model/package split that lasts only until the next reopen is insufficient. An explicit map or a coherent prepared-versus-emission candidate may work, but its callable paths and failure atomicity must be specified. Do not invent a new residual to waive the registered criterion. Required discriminator: definition id, instance, sign, new instance using the original id, later save and signature verification. No execution of that discriminator is claimed here.

## F2, low, carrier existence is not existing coverage

Confirmed. `GROK_PROPOSAL_V3.md:68` says existing test_core.py covers every publisher, while its line 289 names planned checks. Root's exact search found no insert_column or add_internal_hyperlink call in the current `crates/rdocx-py/tests/test_core.py`. Say these checks will be added to that existing carrier. Do not report an absent test or a future test as existing coverage.

## F3, low, picture publication citation

Confirmed. `GROK_PROPOSAL_V3.md:92` cites document.rs:29869 as the prepare call, but that is insert_picture_candidate's declaration. The prepare call is at `crates/rdocx/src/document.rs:29965`. Keep the function and publisher references distinct.

## Other disposition and limits

The authored radical contract at proposal line 118 adds the required empty degree, while imported absence is preserved and the separate F280 validator can report that imported schema fault without refusing open. No cross-story conflict was found in this bounded comparison. The peer's primary-source closure of Q-S96-08 is retained for later full Root assessment.

All eight criteria remain in scope. Source F272 is not an accepted provider. Actual new native note observations do not qualify this story's equation drawing, list rendering or separator spacing. Fifth TOC hash identity, local RED/GREEN, native/oracle, schema, installed Python, typing/wheels, WASM, archive and hosted CI evidence remain pending.

Return to a separate Grok author repair, then Root affected and complete final contract review. No new design peer. The implementation phase later has its separate cap of three. No source edits, builds, tests, Git mutations, native actions or publication ran in this review.
