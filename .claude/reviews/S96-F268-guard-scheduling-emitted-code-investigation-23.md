# S96, actual guard scheduling code, investigation23

## Result and scope

Candidate21 actually restores a compact standard-extend enqueue loop in the emitted optimized guard. The predecessor's per-child saved-Finish bounds, variant, child bounds and Vec push-growth checks are absent from that loop. Reserve/growth handling still exists before extension. Finish installation retains checks once after extension. The empty-child path bypasses both extension and installation.

This is a static generated-code observation on identified artifacts. It does not establish elapsed improvement,19 causality in an earlier timing result, lower memory usage or a satisfied public threshold. It does not replace independent implementation review or the required ordinary artifact qualification. The earlier 75/79 flat result and failed secondary cold insertion remain valid observations. Investigation22's five profiles belong to the Root29 predecessor, not this candidate.

This pass establishes zero product defects, zero smells and zero nitpicks. Root remains the sole source/Cargo writer. No source, plan or existing report was changed.

## Artifact and tool binding

The actual predecessor is the completed Root29 symbol diagnostic at `.claude/scratch/S96-final-evidence/f268-five-remaining-symbol-diagnostic-20261108-1`. Its Native image SHA-256 is c0909f3d0b071c13f7f5b6ce768a2eadfa282e1d36333f73c83c9604aab42875. Its wheel is a8d416dab58b27e610f678c283fac6fce98f75ffdfc7921f98f79cef19f11e2c.

The actual candidate is `.claude/scratch/S96-final-evidence/f268-guard-scheduling-implementation-20261108-1/optimized-symbol`. Its Native image is 1c92db6bfa9e72ffaaf380e8e5e3de1b91225d56ee725942aa99e0051450d7cd. Its wheel is fc5e2528db013a1efe8bfd926e5ffd93f4b7a6fff597594b1c31186842a78c5c. Both BUILD.json records have actual child exit 0 and unchanged producer inputs. Every one of each artifact's ten wheel/payload member hashes was independently checked.

Recorded release commands match after replacing only their output-directory arguments. Both use the pinned Maturin executable, the same CPython interpreter, offline release build, the same manifest/default feature setup and the identical leaf `profile.release.package.rdocx-py.strip="none"` override. Other release optimization/LTO/codegen defaults are retained. These are symbol diagnostics, not ordinary shipping artifacts. This comparison binds the recorded commands and source manifests rather than claiming that arbitrary unrecorded ambient state is proved by an instruction listing.

The two 4309-entry producer manifests differ in exactly four tracked paths: the existing F268 plan, OXML document.rs, HLD03 and HLD04. Native document.rs, comments.rs, binding code and the18/20 implementation inputs are unchanged. The new-summary scheduling branch is the changed compiled Rust policy. Relevant current source is `crates/rdocx-oxml/src/document.rs:1551`. The source delta is narrower than the earlier Root28-versus-Root29 comparison in investigation21.

Static inspection used `/Library/Developer/CommandLineTools/usr/bin/llvm-objdump`, Apple LLVM 21.0.0, selecting the actual word_model_children_maxima symbol from each Mach-O symbol table. No toy function or separately compiled model was substituted. Absolute code addresses below identify these two images only.

## Actual predecessor loop

The Root29 predecessor calls resource_children at 0xeb42d4. Its nonempty new-summary loop contains:

- At0xeb4364/0xeb4368, comparison and branch for the saved pending index.
- At0xeb4378 through 0xeb4380, the retained Finish-variant test.
- At0xeb4384 through 0xeb438c, child length/bounds handling.
- At0xeb439c through 0xeb43b4, pending capacity comparison and a growth path before the next push.

The enqueue body updates the pending length on each iteration. A prior reserve remains in the function, but the emitted loop still defends each individual push. Its completed-overbudget branch separately retains the compact original into_iter/map/extend shape. These checks are correct. Static inspection identifies their existence, not how often each branch executes or mispredicts.

## Actual candidate nonempty loop

The candidate calls resource_children at 0xeb429c, then distinguishes completed state and actual child length. For a nonempty new summary it writes the temporary empty Finish payload at 0xeb42bc through 0xeb42d0. It compares the exact required enqueue count with available pending capacity at 0xeb42d8 through 0xeb42e0. The out-of-line reserve path at 0xeb4468 through 0xeb4484 returns before the bulk loop.

The actual loop at 0xeb4304 through 0xeb4320 is eight instructions:

```text
ldp   x10, x11, [x8, #-0x8]
stp   x13, x10, [x9, #-0x10]
stp   x11, x28, [x9]
str   x26, [x9, #0x10]
add   x8, x8, #0x10
add   x9, x9, #0x38
subs  x25, x25, #0x1
b.ne  0xeb4304
```

The source cursor advances by one borrowed owner, and the destination advances by one queued step. There is no saved-Finish lookup, enum test, child bounds branch, pending push-capacity test or growth call inside this eight-instruction loop. The final pending length is stored once at 0xeb4324. This is the actual desired code shape, not a prediction based solely on a source-level TrustedLen expectation.

After that loop, the code checks the saved Finish index at 0xeb4328/0xeb432c. It locates and tests the existing payload at 0xeb4334 through 0xeb4350, retaining a defensive variant path and generic old-vector drop logic. Correct execution sees the empty placeholder and takes the empty-drop path. The compiler did not erase every descriptor/drop check. It installs the original child vector descriptor at 0xeb4354 through 0xeb4360before returning to the pending-step processing loop.

Those installation checks run once for this nonempty new summary, not once per child. Their presence is consistent with the private invariant assertion in the source. A panic_bounds_check destination remains in the full emitted function. Therefore the accurate claim is removal of repeated checks from the enqueue loop, not removal of every bounds or invariant check from the function.

## Empty and completed-overbudget paths

The actual zero-child test at 0xeb42a8 branches to0xeb43f8. That path directly writes Finish with the original child descriptor and existing key/local fields, updates the pending length and returns to processing at 0xeb4424/0xeb4428. It does not execute the borrowed extend loop or the post-extension placeholder installation. No per-leaf replacement match is introduced on that path.

The completed-but-overbudget path beginning at 0xeb4378 retains the original owned iterator scheduling. Its loop at 0xeb43bc through 0xeb43d4again constructs Enter records in a compact pointer loop. Its reserve and original child-buffer release handling remain outside that loop. This is the required unchanged alternative, not a new summary or permission to skip an offending descendant.

Both algorithms still call actual resource_children before these branches. Neither emitted scheduling comparison establishes that any Source/admission/Field work can be omitted. The design21 correctness proof, complete graph/error controls and actual implementation review remain separate obligations.

## Whole-function and memory tradeoffs

The actual predecessor symbol contains 781 decoded AArch64 instructions and the candidate 779. This tiny whole-function difference does not describe the repeated loop's dynamic cost. Once-per-summary installation code, map/guard logic, reserve paths and error/drop handling contribute to those totals. Static instruction count is not cycles, branch frequency, allocator time or API elapsed time.

The first stack adjustment is `sub sp, sp, #0x150` in the predecessor and `sub sp, sp, #0x160` in the candidate. The emitted fixed native stack reservation therefore grows from 336 to 352 bytes on this build. Queue addressing still advances by 0x38 bytes in both scheduling loops. These are actual emitted storage facts on this target, not a stable portable type ABI or a complete peak-stack/RSS measurement.

Candidate21 retains the original child Vec and its actual spare capacity until Finish, just as19 does. It avoids the old duplicate child buffer but does not restore a length-sized Finish clone or remove the previously disclosed capacity cost. The temporary empty descriptor and later installation stores are visible in the actual code and must not be described as free. The observed extra fixed stack reservation also belongs in final representation and memory calibration applicability.

## Limits and handback

The actual artifact comparison confirms the local enqueue code-generation goal. It does not establish19as the cause of the prior five misses or guarantee21 as their cure. Investigation22 shows that substantial summary hashing, enumeration, parsing, source traversal, allocation/freeing and comparisons remain, and that cold insertion had no sampled named model-guard frame.

No new runtime profile, benchmark or timing was executed in this pass. It read existing build/source/payload receipts and inspected emitted code only. Only this report was written. Root's final composed correctness gates, independent implementation review, fresh ordinary release artifact, new compiler-free unchanged complete qualification and all memory/Office/typing/hosted/archive/full-union riders remain pending at their actual boundaries.
