# F-X280a, Codex independent final1

**Reviewed**: GLM working diff against f4616029f19e12edc0b808670cff1db5e6baa9f0, four product/test/HLD paths,454insertions and5deletions. Read the approved contract, actual shared builders/publication and Python paths, tests, identified GLM self-review, progress record and actual tool stream. No tests/builds/source edits performed in this pass.
**Verdict**: 2 defects, 0 smells, 0 nitpicks. Production helper appears contract-conforming. Required atomicity coverage and verification provenance need repair before preparation.

## Defects

### D1, the claimed populated-group and companion-preservation cohort is absent

`crates/rpptx/tests/integration.rs:37529` creates an empty group and immediately attempts the refused member addition at:37536. No existing group member is present. Python likewise creates an empty group at `crates/rpptx-py/tests/test_documented_examples.py:12788` and tests its empty collection. Root-held shape/text handles are exercised, but held group-member/text authority is not. The Python package comparison is made only after all four refusals at:12802, rather than after each call.

The approved matrix in `.claude/plans/F-X280a-design.md:65` requires a populated group/candidate path, and its atomic-publication row requires an existing relationship-backed member and opaque package part. The native fixture starts with a blank slide and adds a connector, with no such member/opaque preservation seed. The implementation and self-review therefore describe a stronger cohort than the assertions exercise. Empty-group refusal cannot prove preservation of existing child geometry, companion relationships or held group-member authority.

Repair the actual fixtures, retain the useful empty/local and connector-forced controls, seed the required existing state, and assert unchanged package/count/identity after each refusal. Exercise actual Python group-member/text handles and reuse them after the error. A focused separate populated-group discriminator should reject the old producer or document its actual older error behavior. Do not add a public API or weaken the matrix to match the simpler fixture.

### D2, pipeline statuses are reported instead of checked producer exits

`.claude/scratch/F-X280a-progress.md:98` through:108 records scoped GREEN and rebuilt installed evidence, but the actual runner commands pipe Cargo, uv and pytest output through grep/head/tail without preserving producer exit status. The tool stream at `/private/tmp/rdocx-glm-f280a-story/stdout.jsonl:238` records wrapper exit0 while the Cargo test reports failure. At:258 the same wrapper exit0 accompanies a missing pytest module. At:285 a failing integration batch is followed by another test batch and the wrapper still reports0. Final build/lint/test commands at:313, :340 and:344 use the same status-masking construction.

The semantic observations are not discarded: the old assertion failures and later passed-test summaries are visible, and the current extension hash matches the reported1f8cdc0d512ae1d6e17a6c6b00744b4d38657f8ffcd4dc9febd8838486eb0cdb. They do not establish the separately checked successful child/build outcomes required by verify's Build, artifact and semantic outcome section. Full source/producer/imported-module association must also remain explicit after the fixture repairs.

For the affected final-source checks, capture actual subprocess return codes and complete logs before summarizing output. Stop dependent stages on failed builds, bind exact source/profile/configuration to the artifact, and record the actual imported extension path and hash. Preserve earlier pipeline records as wrapper outcomes and semantic observations, never relabel them as checked child exits. Reuse valid unaffected observations and run only missing or invalidated focused checks, not the workspace/final union. Keep every setup failure separate from real semantic RED.

## Smells

Zero findings.

## Nitpicks

Zero findings. The GLM self-review's optional taste observations do not change the two mandatory findings above.

## Not found

No production defect found in the strictly-negative width/height predicate, raw i64 error formatting, three existing call sites, zero/negative-position handling, table/connector scope, public API compatibility, schema ordering or unmodelled XML behavior. The accepted path still calls positioned_transform unchanged. No new unsafe arithmetic, panic, module, trait, generic or dependency was introduced. The HLD edits stay in the approved ShapeCollection sections. No unrelated Word/OfficeMath source change, commit, publication, baseline rewrite or premature final-acceptance claim was found.

## Handoff

This is the first actual independent assessment of the GLM implementation. Its pass1was explicitly self-review and is not promoted to independent evidence. Preserve original design/implementation caps. Return to a separate GLM remediation phase for D1/D2, then obtain fresh independent review. Do not commit or prepare yet. All original final installed, SDK/native, full, hash, archive, typing/WASM and hosted riders remain due. This pass changed only this review file and ends here.
