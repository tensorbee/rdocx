# F-X272 stored ordered input accessor, independent pass1

## Result

Defects:0. Smells:0. Nits:0.

Accept this bounded accessor visibility implementation. No whole dirty carry, Native freezer implementation, cache eligibility, ordinary certificate or C1/provider completion is accepted here.

## Exact scope and identities

Contract is canonical a1ba85a3, Borrow stored ordered Field input without materialization, plan SHA256 a105c31dd3a26f4346bbd495ffbb65b3966334e636bc619b2c663fb5f5ab14f0. Reviewed only the actual worker getter, new focused test and corresponding HLD03/HLD10 paragraphs requested by Root.

Sampled worker text.rs SHA256 `38f7ed8da68ab34244f0ea720876158a4c5c1ce11fe946db55e66ed9ec3284d4`. HLD03 SHA256 `4259a2fe2d44fc0aac6ff24bb4b3aab06e9839523066418e1054b44e13042af8`. HLD10 SHA256 `b42896a624ca47a3e787daf38a76b64dd4fa3e33d4601ad11877f2dbbb49ae63`. These identify a bounded read of the active sole-writer tree, not a new atomic source freeze or review of its other changes. This reviewer changed only this review file and ran no build or test.

## Source check

At `/private/tmp/rdocx-s96-f272-codex-recovery/crates/rdocx-oxml/src/text.rs:6894`, the existing method is now doc-hidden public with exact borrowed signature stored_ordered_result(&self) -> Option<&CT_P>. Its body is byte-identical to the accepted preceding source. It performs only ordered_result.as_ref().map and owner.paragraph.as_ref(). There is no clone, String construction, XML decoding, lazy parser, normalization, admission, cache-display comparison, mutation or input grant.

The return borrows the actual stored paragraph under self's lifetime and preserves None when no stored ordered owner exists. It is intentionally independent of current cache selection. Existing reader/writer branch guards still determine whether that stored state is eligible for effective use. The method does not expose a provenance type, add a model field or change existing public run/projection getters.

## Discriminating test

The new test at `/private/tmp/rdocx-s96-f272-codex-recovery/crates/rdocx-oxml/src/text.rs:47549` constructs an actual parsed ordered Field with a retained control, then changes cached_result to different text. The borrow still points to the exact stored paragraph, retains its control and performs zero COMPLETE_RESULT_QUALIFICATIONS. A newly authored PAGE Field returns None. This distinguishes inert physical input access from cache-display selection and owned normalization.

The counter assertion proves no complete-result qualification was invoked, not a general allocator measurement. Allocation freedom follows the unchanged borrow-only body. The test makes no cache-eligibility or publication claim and does not activate a Word source role.

Root reported the actual targeted command cargo test -p rdocx-oxml --lib stored_ordered_result_borrows_physical_input_without_cache_selection ran one test and passed. The initial unqualified --exact invocation selected zero tests and is not counted as a pass. No historical runtime RED is claimed for a pure visibility addition. This review did not rerun either command or invent an executable/source seal beyond the supplied observation.

## HLD and compatibility

The appended HLD03 paragraph describes stored physical input inspection before cloning and preserves actual reader/writer selection and work admission. HLD10 at `/private/tmp/rdocx-s96-f272-codex-recovery/docs/hld/10-bindings-spec.md:5132` states no allocation, normalization or cache selection and leaves actual writer use authoritative. These match the approved inert accessor, not completed Native source correspondence.

The operation carrier count stays35. This is the separately inventoried fifth borrowed accessor beside the four retained-XML helpers. Exact external client compilation and final Rust0.17.0 additive SemVer inventory remain pending. No Python/WASM expansion, new parser, model field, quota, feature, generic or review allowance is introduced.

## Remaining boundary

Actual membership freezing and same clone/preparation producer correspondence remain separate Root implementation work. No after-the-fact ordinal, equality or persistent candidate-address catalog is authorized. Existing raw/late consumer completeness and every final source/resource/native/installed/history/hash rider remain open.

Review prose was explicitly checked in memory before writing for em dash, en dash and prose semicolon absence. No product or private file was edited for that check. Review pass ended when this file was written and returns control to Root.
