# F-X272, Root inline and revision functional support, pass 2

**Reviewed**: Only the immutable successor from `ROOT_TAKEOVER/root-returned-private-freeze` to `ROOT_TAKEOVER/root-returned-private-repair`, including D1/D2 remediation, the exact selected-input assertion and the existing formatter Field-activation repair. Seven snapshots are pinned, with four source files changed and three unchanged.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this bounded successor. D1 and D2 are closed locally. Whole C1/provider acceptance remains excluded.

## Defects

0 defects.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## D1 and D2 disposition

D1 is repaired at `crates/rdocx-oxml/src/revision.rs:1095` and `crates/rdocx-oxml/src/content_control.rs:1844`. The actual template capture now guards raw-map construction. An absent capture passes None into the existing mapped emitter, which retains its unmapped branch and returns no map for the transport lease. Ordinary retained Publication therefore avoids the newly introduced generated-map allocation and metadata lease. The active scoped path continues to use the same retained-root piece transcript.

D2 is repaired at `crates/rdocx-oxml/src/content_control.rs:2548` and `crates/rdocx-oxml/src/table.rs:2743`, `3188` and `3721`. All four successful handbacks now forward pending.word_raws alongside their existing companion collections. The actual path prefixing remains in WordBlockSourceMaps::append. Failed owner returns do not publish these pending collections.

The compound regression at `crates/rdocx-oxml/src/document.rs:31160` now checks the exact nine decoded input bytes multiplied by the existing Word input multiplier. This distinguishes loss of the E&F revision input from a successful copy funded by the other selected inputs. The same assertion runs in all six existing contexts. It does not introduce another allowance or change its units.

## Formatter repair

`crates/rdocx-oxml/src/document.rs:7359` registers delivered original Field facts after the borrowed replacement kernel succeeds and before final serialization. It uses the existing writer roots and ordered normalization children. The registration helper at `crates/rdocx-oxml/src/text.rs:3876` selects only actual registry-delivered source/current facts and deduplicates them in the same operation. It does not treat a replacement String or generated Field as fresh ingress.

The placement preserves the replacement kernel's existing dependency and occurrence validation before this added activation. A kernel error returns before the new registration. The final writer can now consume already delivered original Field work even when the successful kernel did not need a Field mutation. No additional parser, public API or numerical budget is introduced. No concrete additional defect was found in this repair.

## Source and evidence

All seven repair hashes authenticate against SOURCE.json.

| Path under crates/rdocx-oxml/src | SHA256 |
|---|---|
| document.rs | 6d6adb92ae5db8dd7a172a5fe783c405956963d0f948bca4e985c13a1faf290d |
| text.rs | 6e3c036545eb4b21bb6aefc9892ba0403b2160080d3a47b63a0b48c2377e4eb0 |
| revision.rs | 1d99de298ed7b31db00057be0d8a15116a89c464cb0638f3698312d8c0d4a282 |
| content_control.rs | 1b07eed7ee7d46d5a5e11d20303f35571b4f99d4b272f0f72f3b0552866c4330 |
| table.rs | 4ef73fe5c252bd4d09995a2e2d59642b2506344cbba39be52131f60a682eb6b5 |
| footnotes.rs | 58ee89d9641a29232f5ac794b8dcad6e70b493ddddef18b5df093eb4486348f8 |
| placeholder.rs | b704f39f0fe4cb499120068c8ebbf9d7ca8e9317ee8209d6e5048d5af9c59374 |

The archived strict-2 log records completed rdocx/rdocx-oxml checking in 22.82 seconds. The requested command is all-targets with warnings denied. No all-features result is inferred. The scoped196-1 log records 196 passes and one failure, the previously archived textbox composition case. It is not a successful whole scoped gate. The formatter regression that failed before this repair now passes in that batch. The strengthened compound regression also passes. No reviewer build, test, endpoint or native application execution was performed.

## Explicit boundaries

The remaining failure, scoped_original_textbox_return_materializes_known_field_roles_before_source_drop, remains real preexisting functional debt. It is not waived or assigned a passing result by closing D1/D2. All exclusions in pass 1 remain, including other row/raw-textbox consumers, changed or rebuilt revision returned contexts, comments/glossary, complete non-emitting expression evaluation, final joins and provider closure.

Final callable inventory, integration, SemVer, full local and integrated checks and hash/native/oracle riders remain pending. This pass does not repeat the broader source review or approve whole F-X272 completion.

## Not found

No additional correctness, panic, namespace/schema-order, preservation, operation-identity or structural-rule finding was found in this exact successor. The zero categories apply only to the reviewed repairs and formatter activation change.
