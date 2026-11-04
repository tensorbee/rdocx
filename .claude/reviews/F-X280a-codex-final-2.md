# F-X280a, Codex independent affected final2

**Review kind**: actual independent read-only microscope after GLM repair1 and identified GLM self-review pass2. The original implementation cap remains three, with no design-peer reset.
**Reviewed**: working diff against f4616029f19e12edc0b808670cff1db5e6baa9f0, four tracked product/test/HLD paths, 691 insertions and 5 deletions. Exact diff SHA256 d14deb41a4412ed0f09323ff9d06f2fe318de412ca4221d6177f03fab621c2a1.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Prior D1 and D2 are closed for the affected local implementation and evidence. Final integrated qualification remains pending.

Read the approved plan, progress, final1 findings, GLM all-pass2, actual repair stream/result, relevant first-run stream, HLD units/model/binding contracts, source consumers and retained receipts. No build, test, model, imported-extension execution or source/evidence modification was performed in this review. Only this review file was written.

## Defects

Zero findings.

## Smells

Zero findings.

## Nitpicks

Zero findings. The author's optional self-review taste comments do not identify a material defect or future invariant failure.

## D1 closure, actual seeded state and authority

The new native fixture at `crates/rpptx/tests/integration.rs:37700` seeds an ordinary textbox, a relationship-backed picture, and a group containing a real existing textbox before any refusal. It also writes an opaque `/ppt/opaque.bin` through OPC and reopens before establishing its baseline. The negative-width, negative-height and both-negative preset cases each compare complete serialized package bytes and top-level count/IDs. The additional group textbox and root preset refusals compare the same complete bytes. Full package equality covers existing group XML, descendant geometry/IDs, companion relationships/media and opaque bytes, rather than relying only on top-level vectors. After successful subsequent use, the test reopens, validates, checks opaque bytes, seeded identity prefix, exact existing member geometry and group count. The root and empty-group/connector-forced controls at `integration.rs:37382` are retained.

The Python fixture at `crates/rpptx-py/tests/test_documented_examples.py:12775` now holds an ordinary shape/text handle, a picture, a populated group's collection, existing group member and member-text handle. Each of four refused root/group calls is followed by package equality and relevant count checks. The held member paragraph is actually mutated and read after all errors, the member's width and existing root shape are read, and both held collections successfully append. Save/reopen verifies the member text/width and counts. This directly discriminates the owner-authority risk that an empty-group fixture did not test. A new public API or weakened acceptance matrix was not substituted.

The final native populated-group test suffix is byte-identical in current and old-source test carriers. All 29 tracked production files under the old scratch rpptx, rpptx-py and rpptx-oxml source directories equal the claim Base. The old native run's checked child exit101 is visible at repair stream line182 and its retained full log shows the intended negative-size acceptance panic. This is semantic RED on the old producer, distinct from earlier opaque-part seeding and count mistakes. The new native run is checked exit0 with four focused passes at line189.

Old Python maturin build child exit0, a distinct old extension hash/import path, and actual old populated-group acceptance are visible at repair stream lines157-161. After pytest setup, the identical final Python test fails with checked child exit1 and DID NOT RAISE at line171. The new installed test is checked exit0 at line132. The misleading cached-wheel attempt is explicitly discarded, rather than counted as old-source proof. D1 is closed.

## D2 closure, producer outcomes and exact artifacts

The repair run completed with process exit0 after 982.2413391669979 seconds, which is an orchestration result rather than a test result. Independently authenticated RESULT.json's task, stdout and stderr hashes against retained bytes. Repair stdout SHA256 is 566d134879f637b6b73c1e0037de1924a16bb445981e6f168fbe98778c1a89ca. The source diff hash recorded at stream line196 matches the current working diff exactly.

The final commands capture producer status before filtering logs. Actual receipts show fmt final child0 at line175, Clippy child0 at line186, library test child0 and integration child0 at line189, and Python test child0 at line132. Their retained logs contain respectively an empty successful fmt log, completed warnings-denied Clippy, one library pass, four native passes and one Python3.12.13 pass. The intermediate fmt failure and earlier fixture/setup failures are retained in the stream and not promoted to semantic acceptance.

The later --no-run artifact-discovery pipeline at line193 is not independently a checked Cargo success receipt. It is not needed to establish the successful tested producer: the preceding direct checked integration run and its full log identify the executed binary, and line196 plus current bytes authenticate that artifact. No masked discovery command is promoted here to an additional successful build.

Authenticated current retained artifacts:

| Artifact | SHA256 |
|---|---|
| Current native integration executable, `/private/tmp/rdocx-s96-f280a-target/debug/deps/integration-63d3fc61d54bd197` | 03c7b1980887225270e063b74499375585b4e78a2ccdce9ed46420a3cd2ee2a7 |
| Old native integration executable, corresponding oldsrc-red target | 9d28818c9f0c54a65f50dce6cf7fcde8e1c2dcac6f862afdf44889a727c12df8 |
| Current imported `crates/rpptx-py/python/rpptx/_rpptx.abi3.so` | 1f8cdc0d512ae1d6e17a6c6b00744b4d38657f8ffcd4dc9febd8838486eb0cdb |
| Old scratch imported extension | 912e81aef9a25e40701b68a71b1409b033b3c40c103145aaad3dbd0d804dbfd3 |

The actual current import path is printed at line196 and agrees with the retained extension. This remediation changes tests, not the production helper or binding source. Reusing the identified unchanged extension does not claim a newly checked wheel rebuild. The old-source extension is independently distinct. Original pipeline outcomes remain historical wrapper/semantic observations, not checked producer exits, as required by final1. D2 is closed.

Authenticated full-log hashes include native GREEN 7ecec24a8609c412d0f8daa9f532b7824034af644c88e431ce75f74033901b64, Python GREEN 4b4113320fbb6f432735377208b05037245dc0dd36d55e16c27e65cdc7c1fd0a, old native RED 94f278de2ed9d37aa454d35f711420836b82838ed0546bd55c6037e237b91f1f, and old Python RED a07f7e7aaea9d3b02ad2f66b470304aeda4f07d96303fb46f7fb17bf8125f877. These observations are consumed without rerunning them.

## Affected production and consumer audit

`crates/rpptx/src/lib.rs:22871` implements a private strictly-negative width/height predicate with raw signed i64 operands. Nonnegative inputs use positioned_transform unchanged. The three actual builders use it at `lib.rs:22490`, `:22508` and `:22529`, preserving textbox versus preset operation labels and the existing invalid-shape error family. No dimension clamping, absolute value, arithmetic overflow or public surface was introduced. Negative positions remain independent. The unstyled importer route is exercised in the existing unit carrier, including zero/positive dimensions and absence of style.

Existing append_new_member constructs the member before append. Local and candidate paths therefore refuse before publication. Python's root textbox shortcut calls native add before commit, and its ordinary/group path edits a candidate before preparing/publishing owners at `crates/rpptx-py/src/shape.rs:1843`. Held-handle tests complement this source trace. Tables and endpoint-based connectors retain their separate producers and existing behavior, with native controls retained. Boundary controls at `integration.rs:37565` distinguish zero acceptance and signed positions from an erroneous <=0 or position predicate. The guard is not applied globally to imported geometry or unrelated resize operations.

No new parser/serializer, namespace policy, quota, trait, generic, module, dependency or source file was added. No schema-order/raw-preservation change is introduced by the transform guard. HLD edits are limited to the approved ShapeCollection authoring and atomic-publication paragraphs. No Word, math or other story source change is present.

## Remaining qualification and handoff

The narrow local evidence supports implementation review closure, not universal producer parity, native PowerPoint qualification, final installed Python3.9/3.12 union, typing, SDK, WASM, full workspace, hash harness, archive/dry-run or hosted CI. Applicable final riders remain due over the integrated union. No baseline, commit, prepare, completion, tag or publication is authorized or claimed by this review.

This is the independent affected final2 after the actual GLM repair1 and its explicitly identified self-review pass2. Original caps remain unchanged. Return control to Root for adjudication and the normal subsequent workflow. This pass ends when this file is written.
